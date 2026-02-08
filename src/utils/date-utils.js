/***** DATE UTILITIES *****/

/**
 * Returns a Date (midnight) or "".
 * Accepts:
 * - Date objects
 * - Strings like 9/17/2024, 09-17-2024, 2024-09-17
 * - Numeric serial dates (Sheets/Excel style)
 */
function canonicalizeDateOnly_(val) {
  let d = null;

  // Date object
  if (val instanceof Date && !isNaN(val.getTime())) {
    d = val;
  }

  // Number (possible serial)
  else if (typeof val === "number" && isFinite(val)) {
    // Google Sheets serial date origin: 1899-12-30 (same as Excel for modern dates)
    const ms = Math.round((val - 25569) * 86400 * 1000);
    const maybe = new Date(ms);
    if (!isNaN(maybe.getTime())) d = maybe;
  }

  // String
  else if (typeof val === "string") {
    const s = val.replace(/\u00A0/g, " ").trim();
    if (!s) return "";

    // mm/dd/yyyy or mm-dd-yyyy
    let m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
    if (m) {
      let mo = parseInt(m[1], 10);
      let day = parseInt(m[2], 10);
      let yr = parseInt(m[3], 10);
      if (yr < 100) yr += 2000;
      const dd = new Date(yr, mo - 1, day);
      if (!isNaN(dd.getTime())) d = dd;
    } else {
      // yyyy-mm-dd
      m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
      if (m) {
        const yr = parseInt(m[1], 10);
        const mo = parseInt(m[2], 10);
        const day = parseInt(m[3], 10);
        const dd = new Date(yr, mo - 1, day);
        if (!isNaN(dd.getTime())) d = dd;
      } else {
        // Fallback parse
        const parsed = Date.parse(s);
        if (!isNaN(parsed)) d = new Date(parsed);
      }
    }
  }

  if (!d) return "";

  // Normalize to midnight local
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}
