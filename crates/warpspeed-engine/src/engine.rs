use crate::ironcalc_engine::IronCalcEngine;
use crate::model::{CalcPlan, CalcResult, EngineError, WorkbookSnapshot};

pub trait CalcEngine {
    fn plan(&self, snapshot: &WorkbookSnapshot) -> Result<CalcPlan, EngineError>;
    fn calculate(
        &self,
        snapshot: &WorkbookSnapshot,
        plan: CalcPlan,
    ) -> Result<CalcResult, EngineError>;
}

#[derive(Debug, Default)]
pub struct WarpSpeedEngine {
    inner: IronCalcEngine,
}

impl WarpSpeedEngine {
    pub fn new() -> Self {
        Self {
            inner: IronCalcEngine::default(),
        }
    }

    pub fn run(&self, snapshot: &WorkbookSnapshot) -> Result<CalcResult, EngineError> {
        snapshot.validate()?;
        let plan = self.plan(snapshot)?;
        self.calculate(snapshot, plan)
    }
}

impl CalcEngine for WarpSpeedEngine {
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
