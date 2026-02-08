/***** CONFIG *****/
const YEAR_SOURCES = [
  { year: "2024", id: "1FykMUoXlTGVIyi4MMplP2fkytCkPE3trnPPLhLjdst0" },
  { year: "2025", id: "1by9lOM8tEUqcrcyMk0fnf1biSbVnhZMMSLInG-VQ0pU" },
  { year: "2026", id: "1csZby0HyYg3qvwJy3z37Wl-zGzvpRmNk8UXw5-4OrmY" },
];

// Pull by column LETTER only: A, B, C, E, H, K, L, O (1-based)
const COLS = [1, 2, 3, 5, 8, 11, 12, 15];

// Skip obvious non-pay-period tabs (adjust if you want; does not affect column mapping)
const SKIP_SHEET_NAME_REGEX = /^(README|INSTRUCTIONS|SETTINGS|CONFIG|SUMMARY|DASHBOARD)$/i;
const SKIP_HIDDEN_SHEETS = true;
const HEADER_ROW = 1;


/***** MENU (use exactly as provided) *****/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Visit Consolidation")
    .addSubMenu(
      SpreadsheetApp.getUi()
        .createMenu("1) Rebuild RAW (direct copy)")
        .addItem("Rebuild 2024 Tab (RAW)", "rebuild2024")
        .addItem("Rebuild 2025 Tab (RAW)", "rebuild2025")
        .addItem("Rebuild 2026 Tab (RAW)", "rebuild2026")
        .addItem("Rebuild ALL Year Tabs (RAW)", "rebuildAllYears")
    )
    .addSubMenu(
      SpreadsheetApp.getUi()
        .createMenu("2) Normalize Canonical Dates")
        .addItem("Normalize Dates: 2024", "normalizeDates2024")
        .addItem("Normalize Dates: 2025", "normalizeDates2025")
        .addItem("Normalize Dates: 2026", "normalizeDates2026")
        .addItem("Normalize Dates: ALL", "normalizeDatesAllYears")
    )
    .addSubMenu(
      SpreadsheetApp.getUi()
        .createMenu("3) Normalize HA Names")
        .addItem("Normalize HA Names: 2024", "normalizeAgencies2024")
        .addItem("Normalize HA Names: 2025", "normalizeAgencies2025")
        .addItem("Normalize HA Names: 2026", "normalizeAgencies2026")
        .addItem("Normalize HA Names: ALL", "normalizeAgenciesAllYears")
    )
    .addToUi();
}


/***** STEP 1: REBUILD RAW (direct copy) *****/
function rebuild2024() { rebuildYear_("2024"); }
function rebuild2025() { rebuildYear_("2025"); }
function rebuild2026() { rebuildYear_("2026"); }

function rebuildAllYears() {
  for (const { year } of YEAR_SOURCES) {
    rebuildYear_(year);
  }
}

function rebuildYear_(year) {
  const master = SpreadsheetApp.getActiveSpreadsheet();
  const srcInfo = YEAR_SOURCES.find(x => x.year === year);
  if (!srcInfo) {
    SpreadsheetApp.getUi().alert("No source configured for year: " + year);
    return;
  }

  const dest = getOrCreateSheet_(master, year);
  dest.clearContents();

  const allRows = buildYearRows_(srcInfo.id);

  if (allRows.length === 0) {
    dest.getRange(1, 1).setValue("No rows found.");
    return;
  }

  dest.getRange(1, 1, allRows.length, allRows[0].length).setValues(allRows);
  dest.setFrozenRows(1);
  try { dest.autoResizeColumns(1, allRows[0].length); } catch (e) {}
}


/***** HELPERS (RAW import builder) *****/
function buildYearRows_(sourceSpreadsheetId) {
  const src = SpreadsheetApp.openById(sourceSpreadsheetId);
  const sheets = src.getSheets();

  let output = [];
  let headerWritten = false;

  for (const sh of sheets) {
    const name = sh.getName();

    if (SKIP_HIDDEN_SHEETS && sh.isSheetHidden()) continue;
    if (SKIP_SHEET_NAME_REGEX.test(name)) continue;

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow < HEADER_ROW) continue;
    if (lastCol < Math.max.apply(null, COLS)) continue; // cannot reach O

    const values = sh.getRange(1, 1, lastRow, lastCol).getValues();
    if (!values || values.length < HEADER_ROW) continue;

    // Header is whatever is in row 1 at A,B,C,E,H,K,L,O (no interpretation)
    const header = pickCols_(values[HEADER_ROW - 1], COLS);

    const data = [];
    for (let r = HEADER_ROW; r < values.length; r++) {
      const row = values[r];
      if (!row || row.every(v => v === "" || v === null)) continue;

      // PURE pull by column letter positions only
      data.push(pickCols_(row, COLS));
    }

    if (!headerWritten) {
      output.push(header);
      headerWritten = true;
    }
    if (data.length) output = output.concat(data);
  }

  return output;
}

function pickCols_(row, cols1Based) {
  return cols1Based.map(c => row[c - 1] ?? "");
}

function getOrCreateSheet_(ss, sheetName) {
  let sh = ss.getSheetByName(sheetName);
  if (!sh) sh = ss.insertSheet(sheetName);
  return sh;
}


/***** PLACEHOLDERS ONLY — DO NOT IMPLEMENT UNTIL YOU TELL ME *****/
function normalizeDates2024() { throw new Error("Not implemented yet (waiting for instruction)."); }
function normalizeDates2025() { throw new Error("Not implemented yet (waiting for instruction)."); }
function normalizeDates2026() { throw new Error("Not implemented yet (waiting for instruction)."); }
function normalizeDatesAllYears() { throw new Error("Not implemented yet (waiting for instruction)."); }

function normalizeAgencies2024() { throw new Error("Not implemented yet (waiting for instruction)."); }
function normalizeAgencies2025() { throw new Error("Not implemented yet (waiting for instruction)."); }
function normalizeAgencies2026() { throw new Error("Not implemented yet (waiting for instruction)."); }
function normalizeAgenciesAllYears() { throw new Error("Not implemented yet (waiting for instruction)."); }
