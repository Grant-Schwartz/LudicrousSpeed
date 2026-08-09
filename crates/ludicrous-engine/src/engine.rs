use crate::ironcalc_engine::IronCalcEngine;
use crate::model::{CalcPlan, CalcResult, EngineError, WorkbookSnapshot};
use crate::progress;

pub trait CalcEngine {
    fn plan(&self, snapshot: &WorkbookSnapshot) -> Result<CalcPlan, EngineError>;
    fn calculate(
        &self,
        snapshot: &WorkbookSnapshot,
        plan: CalcPlan,
    ) -> Result<CalcResult, EngineError>;
}

#[derive(Debug, Default)]
pub struct LudicrousSpeedEngine {
    inner: IronCalcEngine,
}

impl LudicrousSpeedEngine {
    pub fn new() -> Self {
        Self {
            inner: IronCalcEngine::default(),
        }
    }

    pub fn run(&self, snapshot: &WorkbookSnapshot) -> Result<CalcResult, EngineError> {
        // Published for a polling host UI; see crate::progress. finish() runs on
        // the error paths too, so a failed run leaves the indicator idle rather
        // than frozen at whatever phase it died in.
        progress::begin(progress::PHASE_LOADING);
        let result = self.run_instrumented(snapshot);
        progress::finish();
        result
    }

    fn run_instrumented(&self, snapshot: &WorkbookSnapshot) -> Result<CalcResult, EngineError> {
        snapshot.validate()?;
        let plan = self.plan(snapshot)?;
        // No phase change here on purpose: the cold-load path reports its own
        // load/analyze/evaluate boundaries from inside calculate(), where the
        // work actually happens. A warm run does none of that and simply stays
        // on PHASE_LOADING until data tables start -- acceptable, because a
        // warm run is short enough that the label barely appears.
        self.calculate(snapshot, plan)
    }
}

impl CalcEngine for LudicrousSpeedEngine {
    fn plan(&self, snapshot: &WorkbookSnapshot) -> Result<CalcPlan, EngineError> {
        self.inner.plan(snapshot)
    }

    fn calculate(
        &self,
        snapshot: &WorkbookSnapshot,
        plan: CalcPlan,
    ) -> Result<CalcResult, EngineError> {
        self.inner.calculate(snapshot, plan)
    }
}
