/***** STEP 2: NORMALIZE CANONICAL DATES (ONLY when called; ONLY on existing rows) *****/

// Output header for the added column
const CANON_DATE_HEADER = "Visit Date Canonical";

// The imported output is ALWAYS 8 columns in this order:
// A: Last, B: First, C: Patient, D: (E from source), E: (H from source) <-- scheduled date value lives here
// F: (K), G: (L), H: (O)
//
// So the date we canonicalize is ALWAYS column 5 on the year tab.
const YEAR_TAB_DATE_COL_1BASED = 5;

function normalizeDates2024() { normalizeDatesOnYearSheet_("2024"); }
function normalizeDates2025() { normalizeDatesOnYearSheet_("2025"); }
function normalizeDates2026() { normalizeDatesOnYearSheet_("2026"); }

function normalizeDatesAllYears() {
  normalizeDatesOnYearSheet_("2024");
  normalizeDatesOnYearSheet_("2025");
  normalizeDatesOnYearSheet_("2026");
}

function normalizeDatesOnYearSheet_(year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(year);

  if (!sh) {
    SpreadsheetApp.getUi().alert("Year tab not found: " + year);
    return;
  }

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) return; // nothing to normalize

  // Ensure the year tab has the imported date column available (col 5)
  if (lastCol < YEAR_TAB_DATE_COL_1BASED) {
    SpreadsheetApp.getUi().alert(
      "Cannot normalize dates for " + year + " because the tab has fewer than " +
      YEAR_TAB_DATE_COL_1BASED + " columns."
    );
    return;
  }

  // Find or create the output column
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v ?? "").trim());
  let canonCol1Based = headers.findIndex(h => h === CANON_DATE_HEADER) + 1;

  if (canonCol1Based < 1) {
    canonCol1Based = lastCol + 1;
    sh.getRange(1, canonCol1Based).setValue(CANON_DATE_HEADER);
  }

  // Read the date values (existing rows only)
  const dateVals = sh.getRange(2, YEAR_TAB_DATE_COL_1BASED, lastRow - 1, 1).getValues();

  // Compute canonical dates
  const out = new Array(dateVals.length);
  for (let i = 0; i < dateVals.length; i++) {
    out[i] = [canonicalizeDateOnly_(dateVals[i][0])];
  }

  // Write
  const outRange = sh.getRange(2, canonCol1Based, out.length, 1);
  outRange.setValues(out);
  outRange.setNumberFormat("m/d/yyyy");
}
