//! Coarse run progress, published for a host UI to poll.
//!
//! This is deliberately a polled global rather than a callback across the FFI
//! boundary. Data tables are evaluated on worker threads, so a callback would
//! fire from whichever worker happened to finish -- and the Excel object model
//! can only be touched from the host's main thread, so every one of those
//! callbacks would have to be marshalled back anyway. Polling inverts it: the
//! host reads whenever it is ready to paint, from whatever thread it likes, and
//! the engine never blocks on the UI.
//!
//! Reads are advisory. The three values are separate atomics and are not
//! sampled atomically together, so a poll landing exactly on a phase change can
//! see a stale `done` against a fresh `total`. That is fine for a progress
//! indicator and not fine for anything else -- nothing should make a decision
//! from these numbers. `begin_with_total` stores `total` before `phase` so a
//! reader that sees a new phase never sees the *previous* phase's total, which
//! is the one skew that would look obviously wrong (a jump to 100% before a
//! phase starts).

use std::sync::atomic::{AtomicU32, AtomicU64, Ordering};

/// No run in flight.
pub const PHASE_IDLE: u32 = 0;
/// Reading the workbook and building the model. The longest phase on a cold
/// run by some margin, which is exactly why it is worth naming separately.
pub const PHASE_LOADING: u32 = 1;
/// Building the dependency graph.
pub const PHASE_ANALYZING: u32 = 2;
/// Evaluating the workbook.
pub const PHASE_EVALUATING: u32 = 3;
/// Evaluating data tables -- the only phase with a meaningful unit count.
pub const PHASE_DATA_TABLES: u32 = 4;

static PHASE: AtomicU32 = AtomicU32::new(PHASE_IDLE);
static DONE: AtomicU64 = AtomicU64::new(0);
static TOTAL: AtomicU64 = AtomicU64::new(0);

/// Enter a phase whose size isn't known up front. Reported as indeterminate:
/// `total` is zero, so a host showing a percentage must fall back to a spinner
/// or a bare phase name rather than dividing.
pub fn begin(phase: u32) {
    begin_with_total(phase, 0);
}

/// Enter a phase of known size.
pub fn begin_with_total(phase: u32, total: u64) {
    DONE.store(0, Ordering::Relaxed);
    TOTAL.store(total, Ordering::Relaxed);
    // Phase last: see the module comment on read skew.
    PHASE.store(phase, Ordering::Relaxed);
}

/// Record completed units within the current phase. Safe to call from worker
/// threads.
pub fn advance(units: u64) {
    DONE.fetch_add(units, Ordering::Relaxed);
}

/// Mark the run finished. A host polling after this sees `PHASE_IDLE` and stops
/// drawing, without having to know the run returned.
pub fn finish() {
    begin_with_total(PHASE_IDLE, 0);
}

/// `(phase, done, total)`. See the module comment: advisory only.
pub fn snapshot() -> (u32, u64, u64) {
    (
        PHASE.load(Ordering::Relaxed),
        DONE.load(Ordering::Relaxed),
        TOTAL.load(Ordering::Relaxed),
    )
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::sync::atomic::AtomicBool;
    use std::sync::{Arc, Mutex, MutexGuard};
    use std::thread;

    // Progress is process-global by design, and Rust runs tests in parallel
    // within one process -- so without this every test's begin() would zero a
    // sibling's counters mid-assert. Serialize them against each other.
    static TEST_LOCK: Mutex<()> = Mutex::new(());

    fn exclusive() -> MutexGuard<'static, ()> {
        // A panicking test poisons the lock; the state is rebuilt by the next
        // begin() anyway, so recover rather than cascading one failure into
        // every other test in the module.
        TEST_LOCK.lock().unwrap_or_else(|err| err.into_inner())
    }

    #[test]
    fn begin_resets_counters_from_a_previous_phase() {
        let _guard = exclusive();
        begin_with_total(PHASE_DATA_TABLES, 10);
        advance(7);
        assert_eq!(snapshot(), (PHASE_DATA_TABLES, 7, 10));

        begin(PHASE_EVALUATING);
        let (phase, done, total) = snapshot();
        assert_eq!(phase, PHASE_EVALUATING);
        assert_eq!(done, 0, "a new phase must not inherit the old count");
        assert_eq!(total, 0, "a sizeless phase reports indeterminate");
    }

    #[test]
    fn finish_returns_to_idle() {
        let _guard = exclusive();
        begin_with_total(PHASE_LOADING, 3);
        advance(3);
        finish();
        assert_eq!(snapshot(), (PHASE_IDLE, 0, 0));
    }

    #[test]
    fn advance_is_safe_from_many_threads() {
        let _guard = exclusive();
        begin_with_total(PHASE_DATA_TABLES, 400);
        let stop = Arc::new(AtomicBool::new(false));

        // A reader running concurrently, standing in for the host's poll loop:
        // it must never panic and must never observe more done than total.
        let reader_stop = Arc::clone(&stop);
        let reader = thread::spawn(move || {
            let mut seen_partial = false;
            while !reader_stop.load(Ordering::Relaxed) {
                let (_, done, total) = snapshot();
                if done > 0 && done < total {
                    seen_partial = true;
                }
            }
            seen_partial
        });

        thread::scope(|scope| {
            for _ in 0..8 {
                scope.spawn(|| {
                    for _ in 0..50 {
                        advance(1);
                    }
                });
            }
        });

        stop.store(true, Ordering::Relaxed);
        reader.join().expect("reader thread panicked");

        let (_, done, total) = snapshot();
        assert_eq!(done, 400, "every increment must be counted exactly once");
        assert_eq!(total, 400);
    }
}
