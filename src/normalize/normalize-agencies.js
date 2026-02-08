// FILE: src/normalize/normalize-agencies.js
// PURPOSE: Implements menu functions:
//   normalizeAgencies2024 / normalizeAgencies2025 / normalizeAgencies2026 / normalizeAgenciesAllYears
// RULES:
// - Runs ONLY when called
// - Works ONLY on existing rows already on the year tabs
// - Does NOT change import behavior
// - Uses your CANONICAL_AGENCIES list (already provided elsewhere in your project)

/***** AGENCY NORMALIZATION (canonical resolver) *****/
const _BASE_OVERRIDES = {
  // Trust USA / All About You, All About You, D9 All About You -> CM All About You
  "ALL ABOUT YOU": "CM All About You",
};

let _CANON_BASE_TO_CANON_CACHE = null;

function _getCanonBaseToCanon_() {
  if (_CANON_BASE_TO_CANON_CACHE) return _CANON_BASE_TO_CANON_CACHE;

  if (typeof CANONICAL_AGENCIES === "undefined" || !Array.isArray(CANONICAL_AGENCIES)) {
    SpreadsheetApp.getUi().alert(
      "CANONICAL_AGENCIES is missing.\n\n" +
      "Add your approved CANONICAL_AGENCIES list to the project, then re-run Normalize HA Names."
    );
    throw new Error("CANONICAL_AGENCIES is not defined.");
  }

  const map = {};
  for (const canon of CANONICAL_AGENCIES) {
    const b = _agencyBaseKey_(canon);
    if (!map[b]) map[b] = [];
    map[b].push(canon);
  }

  _CANON_BASE_TO_CANON_CACHE = map;
  return map;
}


function normalizeAgencies2024() { normalizeAgenciesOnYearSheet_("2024"); }
function normalizeAgencies2025() { normalizeAgenciesOnYearSheet_("2025"); }
function normalizeAgencies2026() { normalizeAgenciesOnYearSheet_("2026"); }

function normalizeAgenciesAllYears() {
  normalizeAgenciesOnYearSheet_("2024");
  normalizeAgenciesOnYearSheet_("2025");
  normalizeAgenciesOnYearSheet_("2026");
}

/**
 * HA Name is ALWAYS the 6th output column on the year tabs:
 * A Last, B First, C Patient, D Visit Type, E Visit Scheduled Date, F HA Name, G HA Initial Price, H Price Agreed
 */
const YEAR_TAB_HA_NAME_COL_1BASED = 6;

function normalizeAgenciesOnYearSheet_(year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(year);

  if (!sh) {
    SpreadsheetApp.getUi().alert("Year tab not found: " + year);
    return;
  }

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) return;

  if (lastCol < YEAR_TAB_HA_NAME_COL_1BASED) {
    SpreadsheetApp.getUi().alert(
      "Cannot normalize HA names for " + year + " because the tab has fewer than " +
      YEAR_TAB_HA_NAME_COL_1BASED + " columns."
    );
    return;
  }

  const rng = sh.getRange(2, YEAR_TAB_HA_NAME_COL_1BASED, lastRow - 1, 1);
  const vals = rng.getValues();

  for (let i = 0; i < vals.length; i++) {
    vals[i][0] = resolveCanonicalAgencyName_(vals[i][0]);
  }

  rng.setValues(vals);
}

/***** Resolver *****/
function resolveCanonicalAgencyName_(rawName) {
  const raw = normalizeSpaces_(rawName);
  if (!raw) return "";

  // Exact canonical match wins
  if (CANONICAL_AGENCIES.indexOf(raw) >= 0) return raw;

  const base = _agencyBaseKey_(raw);

  // Overrides for known rollups
  if (_BASE_OVERRIDES[base]) return _BASE_OVERRIDES[base];

  const canonMap = _getCanonBaseToCanon_();
  const candidates = canonMap[base] || [];
  if (candidates.length === 1) return candidates[0];

  if (candidates.length > 1) {
    // Disambiguate if the raw name contains an explicit program token
    const up = raw.toUpperCase();

    if (/\bD10\b/.test(up)) {
      for (const c of candidates) if (/^D10\b/i.test(c)) return c;
    }
    if (/\bD9\b/.test(up)) {
      for (const c of candidates) if (/^D9\b/i.test(c)) return c;
    }
    if (/\bCM\b/.test(up)) {
      for (const c of candidates) if (/^CM\b/i.test(c)) return c;
    }

    // Still ambiguous -> leave as-is (do NOT guess)
    return raw;
  }

  // No match -> leave as-is
  return raw;
}

function _agencyBaseKey_(name) {
  let s = normalizeSpaces_(name).toUpperCase();

  // Remove common “noise” prefixes
  s = s.replace(/^TRUST\s*USA\s*\/\s*/i, "");
  s = s.replace(/^TRUST\s*USA\s*-\s*/i, "");
  s = s.replace(/^TRUST\s*USA\s*/i, "");

  // Strip program prefix if present (we want base name)
  s = s.replace(/^(D9|D10|CM|PBPT)\b\s*[-–—]?\s*/i, "");

  // Normalize punctuation to spaces, keep letters/numbers/spaces
  s = s.replace(/[^A-Z0-9 ]+/g, " ");
  s = s.replace(/\s+/g, " ").trim();
  return s;
}

function normalizeSpaces_(s) {
  return String(s || "")
    .replace(/\u00A0/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}
