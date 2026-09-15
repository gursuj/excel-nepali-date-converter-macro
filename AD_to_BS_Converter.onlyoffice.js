/**
 * AD (Gregorian) -> BS (Bikram Sambat) date converter
 * ONLYOFFICE Desktop Editors macro (JavaScript) - port of AD_to_BS_Converter.bas
 *
 * Data source: Project_Parva (github.com/dantwoashim/Project_Parva)
 *   backend/app/calendar/constants.py  -> BS_MONTH_LENGTHS
 *   backend/app/calendar/provenance.py -> confidence ranges
 *
 * Coverage: BS 2000-2099 (approx AD 1943-2043)
 * Confidence: BS 2078-2083 = official (structured source backing).
 *             Everything else in range = static_table_unverified.
 * Output is flagged [unverified] outside 2078-2083 - see bsDateText().
 *
 * HOW TO RUN: View tab -> Macros -> paste this whole file -> select AD date
 * cells first -> Run. Result OVERWRITES the selected cell in place (the
 * original AD date is replaced - not reversible except via Ctrl+Z, and
 * that only survives until the next action). Keep a backup column/copy
 * of the AD dates first if you need to keep both.
 *
 * Handles both real Excel-style numeric date serials AND plain text dates
 * (common for CSV-origin data) in "yyyy-mm-dd" or "D-Mon-YY"/"D-Mon-YYYY"
 * format (e.g. "3-Apr-26"). Already-converted BS output is recognized and
 * left alone, so re-running on the same range twice is safe.
 *
 * UNVERIFIED ASSUMPTION (docs didn't confirm this - check on first run):
 *   GetRow() is 0-based (inferred from one doc example, not stated
 *   outright). If reported row numbers are off by one, check this.
 * Test against the known-good dates in README.md before trusting output.
 */
(function () {
  var BS_REF_YEAR = 2082;
  var BS_TABLE_MIN_YEAR = 2000;
  var BS_TABLE_MAX_YEAR = 2099;
  var BS_OFFICIAL_MIN_YEAR = 2078;
  var BS_OFFICIAL_MAX_YEAR = 2083;

  // BS 2082-01-01 = AD 2025-04-14 (Hamro Patro reference)
  var BS_REF_AD_UTC_DAYS = Math.floor(Date.UTC(2025, 3, 14) / 86400000);

  // Same table as AD_to_BS_Converter.bas, transcribed from Project_Parva's constants.py
  var BS_MONTH_LENGTHS = {
    2000: [30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31],
    2001: [31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30],
    2002: [31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30],
    2003: [31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31],
    2004: [30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31],
    2005: [31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30],
    2006: [31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30],
    2007: [31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31],
    2008: [31, 31, 31, 32, 31, 31, 29, 30, 30, 29, 29, 31],
    2009: [31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30],
    2010: [31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30],
    2011: [31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31],
    2012: [31, 31, 31, 32, 31, 31, 29, 30, 30, 29, 30, 30],
    2013: [31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30],
    2014: [31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30],
    2015: [31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31],
    2016: [31, 31, 31, 32, 31, 31, 29, 30, 30, 29, 30, 30],
    2017: [31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30],
    2018: [31, 32, 31, 32, 31, 30, 30, 29, 30, 29, 30, 30],
    2019: [31, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31],
    2020: [31, 31, 31, 32, 31, 31, 30, 29, 30, 29, 30, 30],
    2021: [31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30],
    2022: [31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 30],
    2023: [31, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31],
    2024: [31, 31, 31, 32, 31, 31, 30, 29, 30, 29, 30, 30],
    2025: [31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30],
    2026: [31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31],
    2027: [30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31],
    2028: [31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30],
    2029: [31, 31, 32, 31, 32, 30, 30, 29, 30, 29, 30, 30],
    2030: [31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31],
    2031: [30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31],
    2032: [31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30],
    2033: [31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30],
    2034: [31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31],
    2035: [30, 32, 31, 32, 31, 31, 29, 30, 30, 29, 29, 31],
    2036: [31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30],
    2037: [31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30],
    2038: [31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31],
    2039: [31, 31, 31, 32, 31, 31, 29, 30, 30, 29, 30, 30],
    2040: [31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30],
    2041: [31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30],
    2042: [31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31],
    2043: [31, 31, 31, 32, 31, 31, 29, 30, 30, 29, 30, 30],
    2044: [31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30],
    2045: [31, 32, 31, 32, 31, 30, 30, 29, 30, 29, 30, 30],
    2046: [31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31],
    2047: [31, 31, 31, 32, 31, 31, 30, 29, 30, 29, 30, 30],
    2048: [31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30],
    2049: [31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 30],
    2050: [31, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31],
    2051: [31, 31, 31, 32, 31, 31, 30, 29, 30, 29, 30, 30],
    2052: [31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30],
    2053: [31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 30],
    2054: [31, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31],
    2055: [31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30],
    2056: [31, 31, 32, 31, 32, 30, 30, 29, 30, 29, 30, 30],
    2057: [31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31],
    2058: [30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31],
    2059: [31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30],
    2060: [31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30],
    2061: [31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31],
    2062: [31, 31, 31, 32, 31, 31, 29, 30, 29, 30, 29, 31],
    2063: [31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30],
    2064: [31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30],
    2065: [31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31],
    2066: [31, 31, 31, 32, 31, 31, 29, 30, 30, 29, 29, 31],
    2067: [31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30],
    2068: [31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30],
    2069: [31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31],
    2070: [31, 31, 31, 32, 31, 31, 29, 30, 30, 29, 30, 30],
    2071: [31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30],
    2072: [31, 32, 31, 32, 31, 30, 30, 29, 30, 29, 30, 30],
    2073: [31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31],
    2074: [31, 31, 31, 32, 31, 31, 30, 29, 30, 29, 30, 30],
    2075: [31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30],
    2076: [31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 30],
    2077: [31, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31],
    2078: [31, 31, 31, 32, 31, 31, 30, 29, 30, 29, 30, 30],
    2079: [31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30],
    2080: [31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 30],
    2081: [31, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31],
    2082: [31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30],
    2083: [31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30],
    2084: [31, 31, 32, 31, 31, 30, 30, 30, 29, 30, 30, 30],
    2085: [31, 32, 31, 32, 30, 31, 30, 30, 29, 30, 30, 30],
    2086: [30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 30, 30],
    2087: [31, 31, 32, 31, 31, 31, 30, 30, 30, 30, 30, 30],
    2088: [30, 31, 32, 32, 30, 31, 30, 30, 29, 30, 30, 30],
    2089: [30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 30, 30],
    2090: [30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 30, 30],
    2091: [31, 31, 32, 31, 31, 31, 30, 30, 29, 30, 30, 30],
    2092: [30, 31, 32, 32, 31, 30, 30, 30, 29, 30, 30, 30],
    2093: [30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 30, 30],
    2094: [31, 31, 32, 31, 31, 30, 30, 30, 29, 30, 30, 30],
    2095: [31, 31, 32, 31, 31, 31, 30, 29, 30, 30, 30, 30],
    2096: [30, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30],
    2097: [31, 32, 31, 31, 31, 30, 30, 30, 29, 30, 30, 30],
    2098: [31, 31, 32, 31, 31, 31, 29, 30, 29, 30, 29, 31],
    2099: [31, 31, 32, 31, 31, 31, 30, 29, 29, 30, 30, 31]
  };

  function isOfficialYear(bsYear) {
    return bsYear >= BS_OFFICIAL_MIN_YEAR && bsYear <= BS_OFFICIAL_MAX_YEAR;
  }

  // Excel-style serial (days since 1899-12-30) -> UTC day count, matching BS_REF_AD_UTC_DAYS
  function excelSerialToUtcDays(serial) {
    // 25569 = days between 1899-12-30 and 1970-01-01 (unix epoch)
    return Math.round(serial) - 25569;
  }

  // Core conversion: AD UTC-day-count -> {year, month, day} in BS
  function gregorianDaysToBS(adUtcDays) {
    var remaining = adUtcDays - BS_REF_AD_UTC_DAYS;
    var y = BS_REF_YEAR;
    var m = 1;
    var mLen;

    if (remaining >= 0) {
      while (true) {
        if (y < BS_TABLE_MIN_YEAR || y > BS_TABLE_MAX_YEAR) return null;
        mLen = BS_MONTH_LENGTHS[y][m - 1];
        if (remaining < mLen) return { year: y, month: m, day: remaining + 1 };
        remaining -= mLen;
        m += 1;
        if (m > 12) { m = 1; y += 1; }
      }
    } else {
      while (true) {
        m -= 1;
        if (m < 1) { m = 12; y -= 1; }
        if (y < BS_TABLE_MIN_YEAR || y > BS_TABLE_MAX_YEAR) return null;
        mLen = BS_MONTH_LENGTHS[y][m - 1];
        remaining += mLen;
        if (remaining >= 0) return { year: y, month: m, day: remaining + 1 };
      }
    }
  }

  function pad(n, width) {
    var s = String(n);
    while (s.length < width) s = "0" + s;
    return s;
  }

  // Core: AD UTC-day-count -> formatted "yyyy-mm-dd" BS string, flagged if unverified
  function bsDateTextFromUtcDays(adUtcDays) {
    var bs = gregorianDaysToBS(adUtcDays);
    if (!bs) {
      return "#OUT OF RANGE (BS " + BS_TABLE_MIN_YEAR + "-" + BS_TABLE_MAX_YEAR + " only)";
    }
    var text = pad(bs.year, 4) + "-" + pad(bs.month, 2) + "-" + pad(bs.day, 2);
    if (!isOfficialYear(bs.year)) text += " [unverified]";
    return text;
  }

  // Excel serial number -> formatted BS string
  function bsDateText(excelSerial) {
    return bsDateTextFromUtcDays(excelSerialToUtcDays(excelSerial));
  }

  // Already-converted BS output, e.g. "2082-12-23" or "2082-12-23 [unverified]"
  function looksLikeBsOutput(s) {
    return /^\d{4}-\d{2}-\d{2}/.test(s);
  }

  var MONTH_NAMES = { jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11 };

  // Parses common AD date strings a CSV/text-formatted cell might contain:
  // "yyyy-mm-dd" or "D-Mon-YY"/"D-Mon-YYYY" (e.g. "3-Apr-26", "30 June 2026").
  // Two-digit years are assumed 2000s (matches this tool's actual date range).
  // Returns UTC day count, or null if unrecognized.
  function parseAdStringToUtcDays(raw) {
    var s = raw.trim();
    var m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (m) {
      return Math.floor(Date.UTC(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10)) / 86400000);
    }
    m = s.match(/^(\d{1,2})[\s-]+([A-Za-z]{3,})[\s-]+(\d{2,4})$/);
    if (m) {
      var monKey = m[2].slice(0, 3).toLowerCase();
      if (MONTH_NAMES.hasOwnProperty(monKey)) {
        var yr = parseInt(m[3], 10);
        if (yr < 100) yr += 2000;
        return Math.floor(Date.UTC(yr, MONTH_NAMES[monKey], parseInt(m[1], 10)) / 86400000);
      }
    }
    return null;
  }

  // ---- Entry point ----
  var sheet = Api.GetActiveSheet();
  var selection = sheet.GetSelection();

  var converted = 0;
  var skipped = 0;
  var already = 0;
  var skippedRows = [];
  var sampleSkip = "";
  var maxRow = -1; // lowest selected row seen, 0-based (see assumption below)

  selection.ForEach(function (range) {
    var value = range.GetValue();
    var row0 = range.GetRow(); // Assumption (unverified): 0-based
    if (row0 > maxRow) maxRow = row0;

    var rawText = (typeof value === "string") ? value.trim() : "";

    if (rawText !== "" && looksLikeBsOutput(rawText)) {
      // Already converted by a previous run - re-running is safe, don't
      // re-convert or flag it as a problem. This was the root cause of the
      // "everything skipped" bug: already-BS text was being counted as a
      // failed conversion instead of being recognized and left alone.
      already += 1;
      return;
    }

    var adUtcDays = null;
    if (typeof value === "number" && value > 0) {
      adUtcDays = excelSerialToUtcDays(value);
    } else if (rawText !== "") {
      adUtcDays = parseAdStringToUtcDays(rawText);
    }

    if (adUtcDays !== null) {
      var bsText = bsDateTextFromUtcDays(adUtcDays);
      range.SetNumberFormat("@"); // force text, same as the VBA version
      range.SetValue(bsText); // overwrites the selected cell in place
      converted += 1;
    } else {
      skipped += 1;
      // +1 for the native row number shown in the app (check this against
      // a known skipped row if the numbers look off by one).
      skippedRows.push(row0 + 1);
      if (sampleSkip === "") {
        sampleSkip = "e.g. row " + (row0 + 1) + " raw=[" + rawText + "] type=" + typeof value;
      }
    }
  });

  var msg = "Converted: " + converted + "\nAlready BS (skipped, safe): " + already + "\nSkipped: " + skipped;
  if (skippedRows.length > 0) {
    msg += "\nSkipped rows: " + skippedRows.join(", ");
    msg += "\nSample: " + sampleSkip;
  }

  // No native message box in ONLYOFFICE macros (alert()/window.alert blocked
  // since v7.1). confirm() was tried as a dialog and confirmed NOT to work
  // (silently did nothing - no exception either, so it can't be detected
  // and had to be dropped rather than used as a try/catch fallback trigger).
  // Always writing the cell-based summary below instead.
  if (maxRow >= 0) {
    // Write summary 2 rows below the selection itself (not a fixed guessed
    // cell like A1 - avoids overwriting unrelated data elsewhere on the
    // sheet) and select it. Note: ONLYOFFICE has no reliable "scroll into
    // view" API (known limitation), so Select() marks it active but may not
    // bring it on screen - check the Name Box if you don't see the
    // highlight move.
    var statusRow = maxRow + 2; // 0-based; +2 = one blank row gap below selection
    var status = sheet.GetRangeByNumber(statusRow, 0);
    status.SetValue(msg.replace(/\n/g, " | "));
    status.Select();
  }
})();
