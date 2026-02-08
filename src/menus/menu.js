/***** MENU *****/
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
