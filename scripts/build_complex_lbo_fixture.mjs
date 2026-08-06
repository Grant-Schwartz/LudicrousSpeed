import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = path.resolve("outputs/lbo-fixtures");
const workbookPath = path.join(outputDir, "LudicrousSpeed_Complex_LBO_Test_Model.xlsx");

const wb = Workbook.create();
const sheets = {
  cover: wb.worksheets.add("Cover"),
  assumptions: wb.worksheets.add("Assumptions"),
  operating: wb.worksheets.add("Operating Model"),
  debt: wb.worksheets.add("Debt Schedule"),
  returns: wb.worksheets.add("Returns"),
  sensitivities: wb.worksheets.add("Sensitivity Tables"),
  specs: wb.worksheets.add("Data Table Specs"),
  checks: wb.worksheets.add("Checks"),
  sources: wb.worksheets.add("Sources Audit"),
};

const colors = {
  navy: "#17365D",
  black: "#000000",
  white: "#FFFFFF",
  inputBlue: "#0000FF",
  linkedGreen: "#008000",
  yellow: "#FFF2CC",
  lightBlue: "#D9EAF7",
  lightGreen: "#E2F0D9",
  lightGray: "#F3F6F8",
  red: "#C00000",
  border: "#B7C9D6",
};

const fmt = {
  amount: "$#,##0;[Red]($#,##0);-",
  count: "#,##0;[Red](#,##0);-",
  pct: "0.0%;[Red](0.0%);-",
  pct2: "0.00%;[Red](0.00%);-",
  mult: "0.0x;[Red](0.0x);-",
  irr: "0.0%;[Red](0.0%);-",
};

const years = [2026, 2027, 2028, 2029, 2030, 2031];
const yearHeaders = years.map((y) => `${y}E`);

function range(sheet, a1) {
  return sheet.getRange(a1);
}

function setValues(sheet, a1, values) {
  range(sheet, a1).values = values;
}

function setFormulas(sheet, a1, formulas) {
  range(sheet, a1).formulas = formulas;
}

function section(sheet, a1, title) {
  const r = range(sheet, a1);
  r.values = [[title]];
  r.merge();
  r.format.fill.color = colors.navy;
  r.format.font.color = colors.white;
  r.format.font.bold = true;
}

const revenueCaseAdj =
  "INDEX('Assumptions'!$M$6:$M$8,MATCH('Assumptions'!$M$4,'Assumptions'!$L$6:$L$8,0))";
const marginCaseAdj =
  "INDEX('Assumptions'!$N$6:$N$8,MATCH('Assumptions'!$M$4,'Assumptions'!$L$6:$L$8,0))";

function title(sheet, text) {
  setValues(sheet, "A1:H1", [[text, null, null, null, null, null, null, null]]);
  range(sheet, "A1:H1").merge();
  range(sheet, "A1:H1").format.fill.color = colors.navy;
  range(sheet, "A1:H1").format.font.color = colors.white;
  range(sheet, "A1:H1").format.font.bold = true;
  range(sheet, "A1:H1").format.font.size = 16;
}

function styleSheet(sheet) {
  sheet.showGridLines = false;
  range(sheet, "A:Z").format.font.name = "Arial";
  range(sheet, "A:Z").format.font.size = 10;
}

function styleHeader(sheet, a1) {
  const r = range(sheet, a1);
  r.format.fill.color = colors.lightBlue;
  r.format.font.bold = true;
  r.format.borders = { preset: "outside", style: "thin", color: colors.border };
}

function styleInputs(sheet, a1, numberFormat = null) {
  const r = range(sheet, a1);
  r.format.font.color = colors.inputBlue;
  r.format.fill.color = colors.yellow;
  if (numberFormat) r.setNumberFormat(numberFormat);
}

function styleFormulas(sheet, a1, numberFormat = null) {
  const r = range(sheet, a1);
  r.format.font.color = colors.black;
  if (numberFormat) r.setNumberFormat(numberFormat);
}

function styleLinked(sheet, a1, numberFormat = null) {
  const r = range(sheet, a1);
  r.format.font.color = colors.linkedGreen;
  if (numberFormat) r.setNumberFormat(numberFormat);
}

function autofit(sheet) {
  sheet.getUsedRange().format.autofitColumns();
  sheet.getUsedRange().format.autofitRows();
}

function excelCol(indexOneBased) {
  let col = "";
  let n = indexOneBased;
  while (n > 0) {
    const rem = (n - 1) % 26;
    col = String.fromCharCode(65 + rem) + col;
    n = Math.floor((n - 1) / 26);
  }
  return col;
}

for (const sheet of Object.values(sheets)) styleSheet(sheet);

// Cover
title(sheets.cover, "LudicrousSpeed Complex LBO Recalculation Test Model");
setValues(sheets.cover, "A3:B12", [
  ["Purpose", "Formula-heavy synthetic LBO workbook for testing LudicrousSpeed / IronCalc full-workbook calculation."],
  ["Currency", "$ in millions"],
  ["Version", "v0.1 synthetic fixture"],
  ["Forecast period", "2026E-2031E"],
  ["Primary stress area", "Many sensitivity grids with cross-sheet formulas"],
  ["Formula policy", "Inputs are blue/yellow; formulas are black; cross-sheet links are green"],
  ["Model status", null],
  ["Primary output", "Sponsor IRR / MOIC"],
  ["Note", "The sensitivity sheets are data-table-style formula grids, not Excel native DataTable objects."],
  ["Local use", "Store third-party downloaded models in fixtures/external; this synthetic model is safe to commit/use."],
]);
setFormulas(sheets.cover, "B9:B9", [["='Checks'!F4"]]);
styleHeader(sheets.cover, "A3:A12");
range(sheets.cover, "B3:B12").format.wrapText = true;

// Assumptions
title(sheets.assumptions, "Assumptions");
section(sheets.assumptions, "A3:H3", "Core Transaction Drivers");
setValues(sheets.assumptions, "A4:B24", [
  ["LTM Revenue", 750],
  ["LTM EBITDA Margin", 0.28],
  ["Purchase Multiple", 10.5],
  ["Transaction Fees % EV", 0.025],
  ["Minimum Cash", 25],
  ["Sponsor Fees at Exit", 0.015],
  ["Exit Multiple", 11.0],
  ["Revenue CAGR", 0.06],
  ["EBITDA Margin Expansion", 0.015],
  ["D&A % Revenue", 0.035],
  ["Capex % Revenue", 0.045],
  ["NWC % Revenue", 0.12],
  ["Cash Tax Rate", 0.25],
  ["Management Rollover %", 0.08],
  ["Senior Debt / EBITDA", 4.5],
  ["Sub Debt / EBITDA", 1.5],
  ["Senior Cash Interest", 0.075],
  ["Sub Cash Interest", 0.105],
  ["Senior Amortization %", 0.075],
  ["Cash Sweep %", 0.75],
  ["PIK Interest", 0.04],
]);
styleHeader(sheets.assumptions, "A4:A24");
styleInputs(sheets.assumptions, "B4:B24");
range(sheets.assumptions, "B5:B5").setNumberFormat(fmt.pct);
range(sheets.assumptions, "B6:B6").setNumberFormat(fmt.mult);
range(sheets.assumptions, "B7:B8").setNumberFormat(fmt.pct);
range(sheets.assumptions, "B9:B9").setNumberFormat(fmt.mult);
range(sheets.assumptions, "B10:B18").setNumberFormat(fmt.pct);
range(sheets.assumptions, "B19:B20").setNumberFormat(fmt.mult);
range(sheets.assumptions, "B21:B24").setNumberFormat(fmt.pct);

section(sheets.assumptions, "D3:J3", "Operating Case Ramps");
setValues(sheets.assumptions, "D4:J4", [["Metric", ...yearHeaders]]);
setValues(sheets.assumptions, "D5:J13", [
  ["Revenue Growth", 0.06, 0.065, 0.065, 0.06, 0.055, 0.05],
  ["EBITDA Margin", 0.28, 0.285, 0.29, 0.295, 0.298, 0.30],
  ["D&A % Revenue", 0.035, 0.035, 0.034, 0.034, 0.033, 0.033],
  ["Capex % Revenue", 0.045, 0.045, 0.044, 0.043, 0.042, 0.042],
  ["NWC % Revenue", 0.12, 0.12, 0.118, 0.116, 0.115, 0.115],
  ["Tax Rate", 0.25, 0.25, 0.25, 0.25, 0.25, 0.25],
  ["Senior Rate", 0.075, 0.074, 0.073, 0.072, 0.071, 0.070],
  ["Sub Rate", 0.105, 0.105, 0.103, 0.102, 0.100, 0.100],
  ["Cash Sweep", 0.75, 0.75, 0.78, 0.80, 0.80, 0.80],
]);
styleHeader(sheets.assumptions, "D4:J4");
styleInputs(sheets.assumptions, "E5:J13", fmt.pct);

section(sheets.assumptions, "L3:N3", "Scenario Selector");
setValues(sheets.assumptions, "L4:N8", [
  ["Selected Case", "Base", null],
  ["Case", "Revenue Growth Adj.", "EBITDA Margin Adj."],
  ["Downside", -0.025, -0.025],
  ["Base", 0.000, 0.000],
  ["Upside", 0.025, 0.025],
]);
styleHeader(sheets.assumptions, "L4:N5");
styleInputs(sheets.assumptions, "M6:N8", fmt.pct);
styleInputs(sheets.assumptions, "M4:M4");

// Operating model
title(sheets.operating, "Operating Model");
setValues(sheets.operating, "B3:G3", [yearHeaders]);
styleHeader(sheets.operating, "A3:G3");
section(sheets.operating, "A5:G5", "Income Statement and FCF");
setValues(sheets.operating, "A6:A28", [
  ["Revenue"],
  ["Revenue Growth"],
  ["EBITDA"],
  ["EBITDA Margin"],
  ["D&A"],
  ["EBIT"],
  ["Cash Taxes"],
  ["NOPAT"],
  ["D&A Add-back"],
  ["Capex"],
  ["Change in NWC"],
  ["Unlevered FCF"],
  [""],
  ["Beginning NWC"],
  ["Ending NWC"],
  ["NWC Change Check"],
  [""],
  ["PV Discount Factor"],
  ["PV of FCF"],
  ["Exit EBITDA"],
  ["Exit Enterprise Value"],
  ["Sponsor Exit Fees"],
  ["Enterprise Value Net of Fees"],
]);
setFormulas(sheets.operating, "B6:G28", [
  ["='Assumptions'!B4", `=B6*(1+'Assumptions'!F5+${revenueCaseAdj})`, `=C6*(1+'Assumptions'!G5+${revenueCaseAdj})`, `=D6*(1+'Assumptions'!H5+${revenueCaseAdj})`, `=E6*(1+'Assumptions'!I5+${revenueCaseAdj})`, `=F6*(1+'Assumptions'!J5+${revenueCaseAdj})`],
  ["=B6/'Assumptions'!B4-1", "=C6/B6-1", "=D6/C6-1", "=E6/D6-1", "=F6/E6-1", "=G6/F6-1"],
  [`=B6*('Assumptions'!E6+${marginCaseAdj})`, `=C6*('Assumptions'!F6+${marginCaseAdj})`, `=D6*('Assumptions'!G6+${marginCaseAdj})`, `=E6*('Assumptions'!H6+${marginCaseAdj})`, `=F6*('Assumptions'!I6+${marginCaseAdj})`, `=G6*('Assumptions'!J6+${marginCaseAdj})`],
  ["=B8/B6", "=C8/C6", "=D8/D6", "=E8/E6", "=F8/F6", "=G8/G6"],
  ["=B6*'Assumptions'!E7", "=C6*'Assumptions'!F7", "=D6*'Assumptions'!G7", "=E6*'Assumptions'!H7", "=F6*'Assumptions'!I7", "=G6*'Assumptions'!J7"],
  ["=B8-B10", "=C8-C10", "=D8-D10", "=E8-E10", "=F8-F10", "=G8-G10"],
  ["=MAX(0,B11*'Assumptions'!E10)", "=MAX(0,C11*'Assumptions'!F10)", "=MAX(0,D11*'Assumptions'!G10)", "=MAX(0,E11*'Assumptions'!H10)", "=MAX(0,F11*'Assumptions'!I10)", "=MAX(0,G11*'Assumptions'!J10)"],
  ["=B11-B12", "=C11-C12", "=D11-D12", "=E11-E12", "=F11-F12", "=G11-G12"],
  ["=B10", "=C10", "=D10", "=E10", "=F10", "=G10"],
  ["=-B6*'Assumptions'!E8", "=-C6*'Assumptions'!F8", "=-D6*'Assumptions'!G8", "=-E6*'Assumptions'!H8", "=-F6*'Assumptions'!I8", "=-G6*'Assumptions'!J8"],
  ["=-(B20-B19)", "=-(C20-C19)", "=-(D20-D19)", "=-(E20-E19)", "=-(F20-F19)", "=-(G20-G19)"],
  ["=SUM(B13:B16)", "=SUM(C13:C16)", "=SUM(D13:D16)", "=SUM(E13:E16)", "=SUM(F13:F16)", "=SUM(G13:G16)"],
  [null, null, null, null, null, null],
  ["='Assumptions'!B4*'Assumptions'!E9", "=B20", "=C20", "=D20", "=E20", "=F20"],
  ["=B6*'Assumptions'!E9", "=C6*'Assumptions'!F9", "=D6*'Assumptions'!G9", "=E6*'Assumptions'!H9", "=F6*'Assumptions'!I9", "=G6*'Assumptions'!J9"],
  ["=B20-B19+B16", "=C20-C19+C16", "=D20-D19+D16", "=E20-E19+E16", "=F20-F19+F16", "=G20-G19+G16"],
  [null, null, null, null, null, null],
  ["=1/(1+'Returns'!B23)^1", "=1/(1+'Returns'!B23)^2", "=1/(1+'Returns'!B23)^3", "=1/(1+'Returns'!B23)^4", "=1/(1+'Returns'!B23)^5", "=1/(1+'Returns'!B23)^6"],
  ["=B17*B23", "=C17*C23", "=D17*D23", "=E17*E23", "=F17*F23", "=G17*G23"],
  ["=B8", "=C8", "=D8", "=E8", "=F8", "=G8"],
  ["=B25*'Assumptions'!B10", "=C25*'Assumptions'!B10", "=D25*'Assumptions'!B10", "=E25*'Assumptions'!B10", "=F25*'Assumptions'!B10", "=G25*'Assumptions'!B10"],
  ["=-B26*'Assumptions'!B7", "=-C26*'Assumptions'!B7", "=-D26*'Assumptions'!B7", "=-E26*'Assumptions'!B7", "=-F26*'Assumptions'!B7", "=-G26*'Assumptions'!B7"],
  ["=B26+B27", "=C26+C27", "=D26+D27", "=E26+E27", "=F26+F27", "=G26+G27"],
]);
styleLinked(sheets.operating, "B6:G28");
range(sheets.operating, "B6:G6").setNumberFormat(fmt.amount);
range(sheets.operating, "B8:G8").setNumberFormat(fmt.amount);
range(sheets.operating, "B10:G17").setNumberFormat(fmt.amount);
range(sheets.operating, "B19:G22").setNumberFormat(fmt.amount);
range(sheets.operating, "B24:G28").setNumberFormat(fmt.amount);
range(sheets.operating, "B7:G7").setNumberFormat(fmt.pct);
range(sheets.operating, "B9:G9").setNumberFormat(fmt.pct);
range(sheets.operating, "B23:G23").setNumberFormat("0.000x");

// Debt schedule
title(sheets.debt, "Debt Schedule");
setValues(sheets.debt, "B3:G3", [yearHeaders]);
styleHeader(sheets.debt, "A3:G3");
section(sheets.debt, "A5:G5", "Sources, Uses and Debt Roll-forward");
setValues(sheets.debt, "A6:A34", [
  ["Purchase Enterprise Value"],
  ["Transaction Fees"],
  ["Minimum Cash"],
  ["Total Uses"],
  [""],
  ["LTM EBITDA"],
  ["Senior Debt"],
  ["Sub Debt"],
  ["Management Rollover"],
  ["Sponsor Equity"],
  ["Total Sources"],
  ["Sources / Uses Check"],
  [""],
  ["Beginning Senior Debt"],
  ["Senior Cash Interest"],
  ["Mandatory Amortization"],
  ["Cash Sweep"],
  ["Ending Senior Debt"],
  [""],
  ["Beginning Sub Debt"],
  ["Sub Cash Interest"],
  ["PIK Interest"],
  ["Ending Sub Debt"],
  [""],
  ["Total Debt"],
  ["Cash Available for Debt Paydown"],
  ["Total Cash Interest"],
  ["Interest Coverage"],
  ["Net Debt / EBITDA"],
]);
setFormulas(sheets.debt, "B6:G34", [
  ["='Assumptions'!B4*'Assumptions'!B5*'Assumptions'!B6", "=B6", "=C6", "=D6", "=E6", "=F6"],
  ["=B6*'Assumptions'!B7", "=C6*'Assumptions'!B7", "=D6*'Assumptions'!B7", "=E6*'Assumptions'!B7", "=F6*'Assumptions'!B7", "=G6*'Assumptions'!B7"],
  ["='Assumptions'!B8", "='Assumptions'!B8", "='Assumptions'!B8", "='Assumptions'!B8", "='Assumptions'!B8", "='Assumptions'!B8"],
  ["=SUM(B6:B8)", "=SUM(C6:C8)", "=SUM(D6:D8)", "=SUM(E6:E8)", "=SUM(F6:F8)", "=SUM(G6:G8)"],
  [null, null, null, null, null, null],
  ["='Assumptions'!B4*'Assumptions'!B5", "='Operating Model'!C8", "='Operating Model'!D8", "='Operating Model'!E8", "='Operating Model'!F8", "='Operating Model'!G8"],
  ["=B11*'Assumptions'!B19", "=B23", "=C23", "=D23", "=E23", "=F23"],
  ["=B11*'Assumptions'!B20", "=B28", "=C28", "=D28", "=E28", "=F28"],
  ["=B9*'Assumptions'!B17", "=C9*'Assumptions'!B17", "=D9*'Assumptions'!B17", "=E9*'Assumptions'!B17", "=F9*'Assumptions'!B17", "=G9*'Assumptions'!B17"],
  ["=B9-SUM(B12:B14)", "=C9-SUM(C12:C14)", "=D9-SUM(D12:D14)", "=E9-SUM(E12:E14)", "=F9-SUM(F12:F14)", "=G9-SUM(G12:G14)"],
  ["=SUM(B12:B15)", "=SUM(C12:C15)", "=SUM(D12:D15)", "=SUM(E12:E15)", "=SUM(F12:F15)", "=SUM(G12:G15)"],
  ["=B16-B9", "=C16-C9", "=D16-D9", "=E16-E9", "=F16-F9", "=G16-G9"],
  [null, null, null, null, null, null],
  ["=B12", "=B23", "=C23", "=D23", "=E23", "=F23"],
  ["=B20*'Assumptions'!E11", "=C20*'Assumptions'!F11", "=D20*'Assumptions'!G11", "=E20*'Assumptions'!H11", "=F20*'Assumptions'!I11", "=G20*'Assumptions'!J11"],
  ["=-MIN(B20,'Assumptions'!B23*B20)", "=-MIN(C20,'Assumptions'!B23*C20)", "=-MIN(D20,'Assumptions'!B23*D20)", "=-MIN(E20,'Assumptions'!B23*E20)", "=-MIN(F20,'Assumptions'!B23*F20)", "=-MIN(G20,'Assumptions'!B23*G20)"],
  ["=-MIN(B20+B22,MAX(0,'Operating Model'!B17-B21)*'Assumptions'!E13)", "=-MIN(C20+C22,MAX(0,'Operating Model'!C17-C21)*'Assumptions'!F13)", "=-MIN(D20+D22,MAX(0,'Operating Model'!D17-D21)*'Assumptions'!G13)", "=-MIN(E20+E22,MAX(0,'Operating Model'!E17-E21)*'Assumptions'!H13)", "=-MIN(F20+F22,MAX(0,'Operating Model'!F17-F21)*'Assumptions'!I13)", "=-MIN(G20+G22,MAX(0,'Operating Model'!G17-G21)*'Assumptions'!J13)"],
  ["=SUM(B20:B22)", "=SUM(C20:C22)", "=SUM(D20:D22)", "=SUM(E20:E22)", "=SUM(F20:F22)", "=SUM(G20:G22)"],
  [null, null, null, null, null, null],
  ["=B13", "=B28", "=C28", "=D28", "=E28", "=F28"],
  ["=B25*'Assumptions'!E12", "=C25*'Assumptions'!F12", "=D25*'Assumptions'!G12", "=E25*'Assumptions'!H12", "=F25*'Assumptions'!I12", "=G25*'Assumptions'!J12"],
  ["=B25*'Assumptions'!B24", "=C25*'Assumptions'!B24", "=D25*'Assumptions'!B24", "=E25*'Assumptions'!B24", "=F25*'Assumptions'!B24", "=G25*'Assumptions'!B24"],
  ["=SUM(B25:B27)", "=SUM(C25:C27)", "=SUM(D25:D27)", "=SUM(E25:E27)", "=SUM(F25:F27)", "=SUM(G25:G27)"],
  [null, null, null, null, null, null],
  ["=B23+B28", "=C23+C28", "=D23+D28", "=E23+E28", "=F23+F28", "=G23+G28"],
  ["=MAX(0,'Operating Model'!B17-B21)", "=MAX(0,'Operating Model'!C17-C21)", "=MAX(0,'Operating Model'!D17-D21)", "=MAX(0,'Operating Model'!E17-E21)", "=MAX(0,'Operating Model'!F17-F21)", "=MAX(0,'Operating Model'!G17-G21)"],
  ["=B21+B26", "=C21+C26", "=D21+D26", "=E21+E26", "=F21+F26", "=G21+G26"],
  ["='Operating Model'!B8/B32", "='Operating Model'!C8/C32", "='Operating Model'!D8/D32", "='Operating Model'!E8/E32", "='Operating Model'!F8/F32", "='Operating Model'!G8/G32"],
  ["=B30/'Operating Model'!B8", "=C30/'Operating Model'!C8", "=D30/'Operating Model'!D8", "=E30/'Operating Model'!E8", "=F30/'Operating Model'!F8", "=G30/'Operating Model'!G8"],
]);
styleLinked(sheets.debt, "B6:G34", fmt.amount);
range(sheets.debt, "B33:G34").setNumberFormat(fmt.mult);

// Returns
title(sheets.returns, "Returns");
section(sheets.returns, "A3:E3", "Sponsor Return Summary");
setValues(sheets.returns, "A4:B24", [
  ["Entry Enterprise Value", null],
  ["Entry Fees", null],
  ["Gross Uses", null],
  ["Initial Senior Debt", null],
  ["Initial Sub Debt", null],
  ["Management Rollover", null],
  ["Sponsor Equity", null],
  [""],
  ["Exit Enterprise Value", null],
  ["Less: Sponsor Exit Fees", null],
  ["Less: Ending Senior Debt", null],
  ["Less: Ending Sub Debt", null],
  ["Exit Equity Value", null],
  ["Sponsor Ownership", null],
  ["Sponsor Exit Proceeds", null],
  [""],
  ["MOIC", null],
  ["IRR", null],
  ["WACC / Discount Rate", 0.12],
  ["PV of FCF", null],
  ["Enterprise Value DCF Cross-check", null],
]);
setFormulas(sheets.returns, "B4:B24", [
  ["='Debt Schedule'!B6"],
  ["='Debt Schedule'!B7"],
  ["='Debt Schedule'!B9"],
  ["='Debt Schedule'!B12"],
  ["='Debt Schedule'!B13"],
  ["='Debt Schedule'!B14"],
  ["='Debt Schedule'!B15"],
  [null],
  ["='Operating Model'!G26"],
  ["=-'Operating Model'!G27"],
  ["=-'Debt Schedule'!G23"],
  ["=-'Debt Schedule'!G28"],
  ["=SUM(B12:B15)"],
  ["=1-'Assumptions'!B17"],
  ["=B16*B17"],
  [null],
  ["=B18/B10"],
  ["=(B18/B10)^(1/6)-1"],
  [null],
  ["=SUM('Operating Model'!B24:G24)"],
  ["=B23+'Operating Model'!G28*'Operating Model'!G23"],
]);
styleLinked(sheets.returns, "B4:B22", fmt.amount);
range(sheets.returns, "B20:B21").setNumberFormat(fmt.mult);
range(sheets.returns, "B22:B23").setNumberFormat(fmt.irr);
range(sheets.returns, "B23:B23").format.font.color = colors.inputBlue;
range(sheets.returns, "B23:B23").format.fill.color = colors.yellow;

section(sheets.returns, "D3:K3", "Cash Flow Returns");
setValues(sheets.returns, "D4:L4", [["Metric", "Entry", ...yearHeaders, "Exit"]]);
setValues(sheets.returns, "D5:D10", [
  ["Sponsor Equity"],
  ["FCF to Sponsor"],
  ["Exit Proceeds"],
  ["Net Sponsor Cash Flow"],
  ["Cumulative Cash Flow"],
  ["Return Multiple"],
]);
setFormulas(sheets.returns, "E5:L10", [
  ["=-B10", null, null, null, null, null, null, null],
  [null, "='Operating Model'!B17", "='Operating Model'!C17", "='Operating Model'!D17", "='Operating Model'!E17", "='Operating Model'!F17", "='Operating Model'!G17", null],
  [null, null, null, null, null, null, null, "=B18"],
  ["=SUM(E5:E7)", "=SUM(F5:F7)", "=SUM(G5:G7)", "=SUM(H5:H7)", "=SUM(I5:I7)", "=SUM(J5:J7)", "=SUM(K5:K7)", "=SUM(L5:L7)"],
  ["=E8", "=E9+F8", "=F9+G8", "=G9+H8", "=H9+I8", "=I9+J8", "=J9+K8", "=K9+L8"],
  ["=L9/ABS(E5)", "=L9/ABS(E5)", "=L9/ABS(E5)", "=L9/ABS(E5)", "=L9/ABS(E5)", "=L9/ABS(E5)", "=L9/ABS(E5)", "=L9/ABS(E5)"],
]);
styleHeader(sheets.returns, "D4:L4");
styleLinked(sheets.returns, "E5:L10", fmt.amount);
range(sheets.returns, "E10:L10").setNumberFormat(fmt.mult);

// Sensitivity tables
title(sheets.sensitivities, "Sensitivity Tables");
setValues(sheets.sensitivities, "A3:D3", [["Output Cell", "='Returns'!B21", "Base IRR", "='Returns'!B22"]]);
styleLinked(sheets.sensitivities, "B3:D3");
const tableDefs = [
  ["Exit Multiple x Purchase Multiple", 8.5, 0.5, 7, 8.5, 0.5, 7, "MOIC", "='Returns'!$B$21*($A$1/'Assumptions'!$B$10)*($A$2/'Assumptions'!$B$6)"],
  ["Exit Multiple x Leverage", 8.5, 0.5, 7, 3.5, 0.5, 7, "IRR", "='Returns'!$B$22+($A$1-'Assumptions'!$B$10)*2.0%-($A$2-'Assumptions'!$B$19)*1.2%"],
  ["Revenue CAGR x EBITDA Margin", 0.02, 0.01, 7, 0.24, 0.02, 7, "MOIC", "='Returns'!$B$21*(1+($A$1-'Assumptions'!$B$11)*3+($A$2-'Assumptions'!$B$12)*4)"],
  ["Tax Rate x Capex % Revenue", 0.18, 0.02, 7, 0.025, 0.005, 7, "IRR", "='Returns'!$B$22-($A$1-'Assumptions'!$B$16)*0.4-($A$2-'Assumptions'!$B$14)*0.8"],
  ["Cash Sweep x Senior Rate", 0.50, 0.05, 7, 0.055, 0.005, 7, "MOIC", "='Returns'!$B$21*(1+($A$1-'Assumptions'!$B$23)*0.25-($A$2-'Assumptions'!$B$21)*1.8)"],
  ["NWC % Revenue x EBITDA Margin Expansion", 0.08, 0.01, 7, 0.00, 0.005, 7, "IRR", "='Returns'!$B$22-($A$1-'Assumptions'!$B$15)*0.2+($A$2-'Assumptions'!$B$12)*1.4"],
  ["Exit Multiple x Sub Debt Multiple", 8.5, 0.5, 7, 0.5, 0.25, 7, "MOIC", "='Returns'!$B$21*(1+($A$1-'Assumptions'!$B$10)*0.16-($A$2-'Assumptions'!$B$20)*0.08)"],
  ["Entry Fees x Exit Fees", 0.01, 0.005, 7, 0.005, 0.005, 7, "IRR", "='Returns'!$B$22-($A$1-'Assumptions'!$B$7)*0.5-($A$2-'Assumptions'!$B$7)*0.4"],
  ["Revenue CAGR x Exit Multiple", 0.02, 0.01, 7, 8.5, 0.5, 7, "MOIC", "='Returns'!$B$21*(1+($A$1-'Assumptions'!$B$11)*3.5+($A$2-'Assumptions'!$B$10)*0.12)"],
  ["EBITDA Margin x Senior Debt Multiple", 0.24, 0.02, 7, 3.5, 0.5, 7, "IRR", "='Returns'!$B$22+($A$1-'Assumptions'!$B$12)*1.5-($A$2-'Assumptions'!$B$19)*0.01"],
  ["PIK Rate x Sub Rate", 0.00, 0.01, 7, 0.08, 0.01, 7, "MOIC", "='Returns'!$B$21*(1-($A$1-'Assumptions'!$B$24)*1.1-($A$2-'Assumptions'!$B$22)*0.9)"],
  ["D&A % Revenue x Tax Rate", 0.025, 0.005, 7, 0.18, 0.02, 7, "IRR", "='Returns'!$B$22+($A$1-'Assumptions'!$B$13)*0.3-($A$2-'Assumptions'!$B$16)*0.4"],
  ["Exit Multiple x Cash Sweep", 8.5, 0.5, 7, 0.50, 0.05, 7, "MOIC", "='Returns'!$B$21*(1+($A$1-'Assumptions'!$B$10)*0.15+($A$2-'Assumptions'!$B$23)*0.2)"],
  ["Purchase Multiple x Senior Rate", 8.5, 0.5, 7, 0.055, 0.005, 7, "IRR", "='Returns'!$B$22-($A$1-'Assumptions'!$B$6)*1.0%-($A$2-'Assumptions'!$B$21)*1.0"],
  ["Revenue CAGR x NWC % Revenue", 0.02, 0.01, 7, 0.08, 0.01, 7, "MOIC", "='Returns'!$B$21*(1+($A$1-'Assumptions'!$B$11)*3-($A$2-'Assumptions'!$B$15)*0.35)"],
  ["EBITDA Margin x Exit Fees", 0.24, 0.02, 7, 0.005, 0.005, 7, "IRR", "='Returns'!$B$22+($A$1-'Assumptions'!$B$12)*1.2-($A$2-'Assumptions'!$B$7)*0.5"],
  ["Exit Multiple x Hold Period Proxy", 8.5, 0.5, 7, 4, 1, 7, "MOIC", "='Returns'!$B$21*(1+($A$1-'Assumptions'!$B$10)*0.15)*(1+$A$2*0.01)"],
  ["Management Rollover x Exit Multiple", 0.00, 0.025, 7, 8.5, 0.5, 7, "IRR", "='Returns'!$B$22-($A$1-'Assumptions'!$B$17)*0.2+($A$2-'Assumptions'!$B$10)*2.0%"],
  ["Capex % Revenue x Cash Sweep", 0.025, 0.005, 7, 0.50, 0.05, 7, "MOIC", "='Returns'!$B$21*(1-($A$1-'Assumptions'!$B$14)*1.3+($A$2-'Assumptions'!$B$23)*0.2)"],
  ["Revenue CAGR x Purchase Multiple", 0.02, 0.01, 7, 8.5, 0.5, 7, "IRR", "='Returns'!$B$22+($A$1-'Assumptions'!$B$11)*1.4-($A$2-'Assumptions'!$B$6)*1.5%"],
  ["Senior Debt x Sub Debt", 3.0, 0.5, 7, 0.5, 0.25, 7, "MOIC", "='Returns'!$B$21*(1-($A$1-'Assumptions'!$B$19)*0.08-($A$2-'Assumptions'!$B$20)*0.06)"],
  ["Terminal EBITDA x Exit Multiple", 190, 10, 7, 8.5, 0.5, 7, "MOIC", "=($A$1*$A$2-'Debt Schedule'!$G$30)/'Returns'!$B$10"],
  ["Enterprise Value x Exit Debt", 1700, 100, 7, 0, 50, 7, "IRR", "=MAX(0,($A$1-$A$2)/'Returns'!$B$10)^(1/6)-1"],
  ["FCF Yield x Exit Multiple", 0.05, 0.01, 7, 8.5, 0.5, 7, "MOIC", "='Returns'!$B$21*(1+($A$1-0.08)*2+($A$2-'Assumptions'!$B$10)*0.15)"],
];

let tableIndex = 0;
for (let blockRow = 5; blockRow <= 89; blockRow += 17) {
  for (let blockCol = 1; blockCol <= 13; blockCol += 12) {
    const def = tableDefs[tableIndex++];
    if (!def) continue;
    const [name, rowStart, rowStep, rowCount, colStart, colStep, colCount, metric, formula] = def;
    const row = blockRow;
    const col = blockCol;
    const titleRange = sheets.sensitivities.getRangeByIndexes(row - 1, col - 1, 1, 9);
    titleRange.values = [[`${tableIndex}. ${name}`, null, null, null, null, null, null, null, null]];
    titleRange.merge();
    titleRange.format.fill.color = colors.navy;
    titleRange.format.font.color = colors.white;
    titleRange.format.font.bold = true;

    sheets.sensitivities.getCell(row, col - 1).values = [[metric]];
    for (let j = 0; j < colCount; j++) {
      sheets.sensitivities.getCell(row, col + j).values = [[colStart + colStep * j]];
    }
    for (let i = 0; i < rowCount; i++) {
      sheets.sensitivities.getCell(row + 1 + i, col - 1).values = [[rowStart + rowStep * i]];
      const rowFormulas = [];
      for (let j = 0; j < colCount; j++) {
        const rowDriverAddress = `${excelCol(col)}${row + 2 + i}`;
        const columnDriverAddress = `${excelCol(col + 1 + j)}${row + 1}`;
        rowFormulas.push(
          formula
            .replaceAll("$A$1", rowDriverAddress)
            .replaceAll("$A$2", columnDriverAddress),
        );
      }
      sheets.sensitivities.getRangeByIndexes(row + 1 + i, col, 1, colCount).formulas = [rowFormulas];
    }
    const body = sheets.sensitivities.getRangeByIndexes(row, col - 1, rowCount + 1, colCount + 1);
    body.format.borders = { preset: "all", style: "thin", color: colors.border };
    body.setNumberFormat(metric === "IRR" ? fmt.irr : fmt.mult);
  }
}

// Data table specs
title(sheets.specs, "Data Table Specs");
setValues(sheets.specs, "A3:F3", [["Table", "Sheet", "Range", "Row Driver", "Column Driver", "Output Metric"]]);
styleHeader(sheets.specs, "A3:F3");
const specRows = tableDefs.map((def, idx) => [
  idx + 1,
  "Sensitivity Tables",
  "See generated grid block",
  def[0].split(" x ")[0],
  def[0].split(" x ")[1] ?? "N/A",
  def[7],
]);
setValues(sheets.specs, `A4:F${3 + specRows.length}`, specRows);
range(sheets.specs, `A4:F${3 + specRows.length}`).format.borders = { preset: "all", style: "thin", color: colors.border };

// Checks
title(sheets.checks, "Checks");
setValues(sheets.checks, "A3:G3", [["Check", "Actual", "Expected", "Difference", "Tolerance", "Status", "Notes"]]);
styleHeader(sheets.checks, "A3:G3");
setValues(sheets.checks, "A4:A10", [
  ["Sources equal uses"],
  ["Debt roll-forward sign"],
  ["NWC check"],
  ["MOIC positive"],
  ["IRR not error"],
  ["Sensitivity tables populated"],
  ["Formula-preserving fixture"],
]);
setFormulas(sheets.checks, "B4:F10", [
  ["='Debt Schedule'!B17", 0, "=B4-C4", 0.01, '=IF(ABS(D4)<=E4,"OK","FAIL")'],
  ["=MIN('Debt Schedule'!B30:G30)", 0, "=B5-C5", 0.01, '=IF(B5>=0,"OK","FAIL")'],
  ["=SUM('Operating Model'!B21:G21)", 0, "=B6-C6", 0.01, '=IF(ABS(D6)<=E6,"OK","FAIL")'],
  ["='Returns'!B21", 1, "=B7-C7", 0, '=IF(B7>C7,"OK","FAIL")'],
  ["='Returns'!B22", 0, "=B8-C8", 0, '=IF(ISNUMBER(B8),"OK","FAIL")'],
  ["=COUNTA('Sensitivity Tables'!A1:Z100)", 1, "=B9-C9", 0, '=IF(B9>C9,"OK","FAIL")'],
  [0, 0, "=B10-C10", 0, '=IF(ABS(D10)<=E10,"OK","FAIL")'],
]);
setValues(sheets.checks, "G4:G10", [
  ["Entry capitalization should balance."],
  ["Debt should not become negative."],
  ["NWC change helper row should tie."],
  ["MOIC should be above 1.0x in base case."],
  ["IRR must resolve to a number."],
  ["Sensitivity grid area must contain formulas/values."],
  ["The fixture does not use native Excel DataTable objects."],
]);
setValues(sheets.checks, "A12:F12", [["Overall Status", null, null, null, null, null]]);
setFormulas(sheets.checks, "F12:F12", [['=IF(COUNTIF(F4:F10,"FAIL")=0,"OK","FAIL")']]);
range(sheets.checks, "B4:E10").setNumberFormat(fmt.count);
range(sheets.checks, "F4:F12").format.fill.color = colors.lightGreen;
range(sheets.checks, "F4:F12").format.font.bold = true;

// Sources
title(sheets.sources, "Sources Audit");
setValues(sheets.sources, "A3:E3", [["Item", "Value", "Units", "Source", "Notes"]]);
styleHeader(sheets.sources, "A3:E3");
setValues(sheets.sources, "A4:E12", [
  ["All financial values", "Synthetic", "$mm", "Generated by LudicrousSpeed fixture builder", "Not company-specific; designed for calculation engine regression."],
  ["LBO structure", "Synthetic", "N/A", "Internal test model", "Includes sponsor equity, debt schedule, operating forecast, and returns."],
  ["Sensitivity tables", tableDefs.length, "tables", "Internal test model", "Formula grids stress cross-sheet references and repeated recalculation."],
  ["Data table caveat", "Formula grids", "N/A", "Internal test model", "Artifact builder does not create native Excel DataTable objects in this fixture."],
  ["No external links", "TRUE", "Boolean", "Internal test model", "Useful for isolating IronCalc support before external-link fallback tests."],
  ["No VBA/macros", "TRUE", "Boolean", "Internal test model", "Workbook is .xlsx only."],
  ["Input convention", "Blue/yellow", "Style", "Financial modeling convention", "Editable assumptions are visually tagged."],
  ["Formula convention", "Black/green", "Style", "Financial modeling convention", "Cross-sheet links are green; formulas are preserved."],
  ["Use case", "LudicrousSpeed regression", "N/A", "Internal test model", "Large enough to profile formula coverage and sensitivity recalc paths."],
]);
range(sheets.sources, "A4:E12").format.wrapText = true;

for (const sheet of Object.values(sheets)) {
  try {
    sheet.freezePanes.freezeRows(3);
  } catch {}
  autofit(sheet);
}

await fs.mkdir(outputDir, { recursive: true });
const errors = await wb.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "formula error scan",
});
console.log(errors.ndjson);

const overview = await wb.inspect({
  kind: "workbook,sheet,table",
  maxChars: 6000,
  tableMaxRows: 8,
  tableMaxCols: 8,
});
console.log(overview.ndjson);

for (const sheetName of Object.keys(sheets).map((key) => sheets[key].name)) {
  const blob = await wb.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  const bytes = new Uint8Array(await blob.arrayBuffer());
  await fs.writeFile(path.join(outputDir, `${sheetName.replaceAll(" ", "_")}.png`), bytes);
}

const xlsx = await SpreadsheetFile.exportXlsx(wb);
await xlsx.save(workbookPath);
console.log(workbookPath);
