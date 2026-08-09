use std::ffi::{CStr, CString};
use std::os::raw::c_char;
use std::sync::{Mutex, OnceLock};

use crate::engine::LudicrousSpeedEngine;
use crate::model::{CalcResult, EngineError, WorkbookSnapshot};

static ENGINE: OnceLock<Mutex<LudicrousSpeedEngine>> = OnceLock::new();

#[no_mangle]
pub extern "C" fn ludicrous_run_json(request_json: *const c_char) -> *mut c_char {
    let response = run_json(request_json).unwrap_or_else(error_response);
    string_to_ptr(response)
}

/// Reads the current run's progress. Safe to call from any thread at any time,
/// including while `ludicrous_run_json` is in flight on another thread -- which
/// is the entire point, since that call blocks for the length of the run.
///
/// Null pointers are skipped, so a caller that only wants the phase can pass
/// null for the counts. When no run is in flight the phase is `PHASE_IDLE` (0).
/// A `total` of zero means the current phase is indeterminate: show a phase
/// name, not a percentage.
#[no_mangle]
pub extern "C" fn ludicrous_progress(
    phase_out: *mut u32,
    done_out: *mut u64,
    total_out: *mut u64,
) {
    let (phase, done, total) = crate::progress::snapshot();
    unsafe {
        if !phase_out.is_null() {
            *phase_out = phase;
        }
        if !done_out.is_null() {
            *done_out = done;
        }
        if !total_out.is_null() {
            *total_out = total;
        }
    }
}

#[no_mangle]
pub extern "C" fn ludicrous_free_string(ptr: *mut c_char) {
    if ptr.is_null() {
        return;
    }
    unsafe {
        let _ = CString::from_raw(ptr);
    }
}

fn run_json(request_json: *const c_char) -> Result<String, EngineError> {
    if request_json.is_null() {
        return Err(EngineError::Serialization(
            "request pointer is null".to_string(),
        ));
    }

    let request = unsafe { CStr::from_ptr(request_json) }
        .to_str()
        .map_err(|err| EngineError::Serialization(err.to_string()))?;
    let snapshot: WorkbookSnapshot =
        serde_json::from_str(request).map_err(|err| EngineError::Serialization(err.to_string()))?;
    let engine = ENGINE.get_or_init(|| Mutex::new(LudicrousSpeedEngine::new()));
    let result = engine
        .lock()
        .map_err(|_| EngineError::Evaluation("engine lock is poisoned".to_string()))?
        .run(&snapshot)?;
    serialize_ok(result)
}

fn serialize_ok(result: CalcResult) -> Result<String, EngineError> {
    serde_json::to_string(&FfiResponse::ok(result))
        .map_err(|err| EngineError::Serialization(err.to_string()))
}

fn error_response(err: EngineError) -> String {
    serde_json::to_string(&FfiResponse::<()>::err(err.to_string())).unwrap_or_else(|_| {
        "{\"ok\":false,\"error\":\"failed to serialize error response\"}".to_string()
    })
}

fn string_to_ptr(value: String) -> *mut c_char {
    CString::new(value)
        .unwrap_or_else(|_| {
            CString::new("{\"ok\":false,\"error\":\"response contained nul byte\"}").unwrap()
        })
        .into_raw()
}

#[derive(serde::Serialize)]
struct FfiResponse<T> {
    ok: bool,
    result: Option<T>,
    error: Option<String>,
}

impl<T> FfiResponse<T> {
    fn ok(result: T) -> Self {
        Self {
            ok: true,
            result: Some(result),
            error: None,
        }
    }

    fn err(error: String) -> Self {
        Self {
            ok: false,
            result: None,
            error: Some(error),
        }
    }
}
