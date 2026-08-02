# LBO Test Model Sourcing

Use legally downloadable/public templates only. Do not commit downloaded third-party
workbooks unless their license explicitly permits redistribution. Prefer storing
local-only copies in `fixtures/external`, which is ignored by git.

## Best candidates

1. Macabacus long-form LBO model
   - URL: https://macabacus.com/category/excel/templates
   - Why: likely the best public candidate for a larger, investment-banking-style
     workbook with realistic formula structure.
   - Notes: download terms should be checked before committing or sharing.

2. YouExec LBO model
   - URL: https://www.youexec.com/resources/lbo-model-in-excel-and-google-sheets
   - Why: advertises a 15-sheet LBO model, which is useful for cross-sheet graph
     and workbook-scale testing.
   - Notes: may require account/signup.

3. Wall Street Prep LBO model templates
   - URL: https://www.wallstreetprep.com/knowledge/lbo-model/
   - Why: reputable training models, useful for correctness checks and formula
     coverage.
   - Notes: often require a form submission.

4. Multiple Expansion advanced LBO tutorials
   - URL: https://multipleexpansion.com/
   - Why: advanced LBO tutorials discuss casing, sensitivities, waterfalls, and
     data-table-like workflows.
   - Notes: useful as modeling reference even when the workbook is not large.

5. Damodaran Merger & LBO valuation model on Eloquens
   - URL: https://www.eloquens.com/tool/nRSN/finance/leveraged-buyout-lbo/merger-lbo-valuation
   - Why: reputable author and free model, good for diversified formula coverage.
   - Notes: likely not as large as a full sponsor-style LBO.

## Not ideal as primary regression targets

- Finamodel free templates may be values-only samples; useful for layout examples,
  but not enough for formula engine regression if formulas are removed.
- One-sheet/simple LBO calculators are good smoke tests, but too small to validate
  workbook-scale graph and cache behavior.

## Local handling

1. Download candidate workbooks manually.
2. Save them under `fixtures/external`.
3. Run:

   ```sh
   cargo run -p warpspeed-engine --bin warpspeed-cli -- fixtures/external/model.xlsx analyze
   ```

4. Record formula coverage, fallback reasons, and timings in a local notes file.
5. Promote only sanitized or license-clear workbooks into committed regression
   fixtures.
