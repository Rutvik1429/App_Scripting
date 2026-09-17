/**
 * Shared helper functions used by every module. Nothing in here talks
 * to a specific sheet's business logic — it's plumbing other files reuse.
 */

function getSheet_(name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error('Sheet "' + name + '" not found. Did you run setupDatabase() in Setup.gs?');
  return sheet;
}

/**
 * Reads the date out of an imported file and returns it as "yyyy-MM-dd"
 * (or "yyyy-MM-dd HH:mm:ss" when the file also carried a time).
 *
 * Every source writes dates differently, so this accepts the common ones:
 *
 *   2026-09-01T14:23:45+0530   Meta Lead Ads
 *   2026-09-01 14:23:45        exported spreadsheets
 *   2026-09-01                 plain ISO
 *   01/09/2026                 day first — the Indian convention
 *   01-09-2026                 day first
 *   1 Sep 2026 / Sep 1, 2026   written out
 *
 * On slash and dash dates the DAY is read first, because that's how dates
 * are written in India. So 05/09/2026 is 5 September, not 9 May. (Meta's own
 * exports use the ISO form above, which has no such ambiguity.)
 *
 * Returns '' when the value isn't a date at all — the caller then falls back
 * to the time of import.
 */
function parseImportDate_(value) {
  if (value === null || value === undefined || value === '') return '';

  var tz = Session.getScriptTimeZone();

  // Already a real date (an Excel cell, say)
  if (value instanceof Date) {
    return isNaN(value.getTime()) ? '' : Utilities.formatDate(value, tz, 'yyyy-MM-dd');
  }

  var text = String(value).trim();
  if (!text) return '';

  var m;

  // 2026-09-01, optionally followed by a time
  m = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T ](\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (m) {
    var iso = pad4_(m[1]) + '-' + pad2_(m[2]) + '-' + pad2_(m[3]);
    if (m[4]) iso += ' ' + pad2_(m[4]) + ':' + pad2_(m[5]) + ':' + pad2_(m[6] || '0');
    return iso;
  }

  // 01/09/2026 or 01-09-2026 — day first
  m = text.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})(?:[T ](\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (m) {
    var day = parseInt(m[1], 10);
    var month = parseInt(m[2], 10);
    var year = parseInt(m[3], 10);
    if (year < 100) year += 2000;

    // If the first number can't be a day, the file is month-first after all.
    if (day > 12 && month <= 12) { /* day-first, as assumed */ }
    else if (month > 12 && day <= 12) { var swap = day; day = month; month = swap; }

    if (month < 1 || month > 12 || day < 1 || day > 31) return '';

    var out = pad4_(String(year)) + '-' + pad2_(String(month)) + '-' + pad2_(String(day));
    if (m[4]) out += ' ' + pad2_(m[4]) + ':' + pad2_(m[5]) + ':' + pad2_(m[6] || '0');
    return out;
  }

  // "1 Sep 2026", "Sep 1, 2026" and similar — let the browser engine try
  var parsed = new Date(text);
  if (!isNaN(parsed.getTime()) && parsed.getFullYear() > 1970 && parsed.getFullYear() < 2200) {
    return Utilities.formatDate(parsed, tz, 'yyyy-MM-dd');
  }

  return '';
}

function pad2_(v) { return ('0' + String(v)).slice(-2); }
function pad4_(v) { return ('0000' + String(v)).slice(-4); }

/**
 * Adds any missing values to a column on the Settings tab.
 *
 * The Clients sheet's dropdown columns (Source, Industry, Status, Priority)
 * read their allowed options from Settings, and REJECT anything not listed —
 * that's why importing a lead with Source "Facebook Ads" failed with "violates
 * the data validation rules". Registering new values here first means the
 * import succeeds AND the dropdown offers them from then on.
 *
 * columnIndex: 1 = Statuses, 2 = Sources, 3 = Industries, 4 = Priorities.
 * Returns how many new values were added.
 */
function ensureSettingsValues_(columnIndex, values) {
  if (!values || !values.length) return 0;

  var settings = getSheet_('Settings');
  var maxRows = Math.max(settings.getMaxRows() - 1, 1);
  var existing = settings.getRange(2, columnIndex, maxRows, 1).getValues();

  var seen = {};
  var lastFilledRow = 1; // the header row
  existing.forEach(function (r, i) {
    var v = String(r[0] || '').trim();
    if (v) {
      seen[v.toLowerCase()] = true;
      lastFilledRow = i + 2;
    }
  });

  var toAdd = [];
  values.forEach(function (v) {
    var name = String(v || '').trim();
    if (!name) return;
    var key = name.toLowerCase();
    if (seen[key]) return;
    seen[key] = true;
    toAdd.push([name]);
  });

  if (!toAdd.length) return 0;

  settings.getRange(lastFilledRow + 1, columnIndex, toAdd.length, 1).setValues(toAdd);
  SpreadsheetApp.flush();
  return toAdd.length;
}

/** Turns a sheet's rows into an array of plain objects keyed by header name.
 *  Always returns an array — even for a totally empty or header-only sheet —
 *  so callers can safely chain .filter()/.map() on the result without a
 *  null/undefined check every time. */
function sheetToObjects_(sheet) {
  var data = sheet.getDataRange().getValues();
  if (!data || data.length === 0) return [];
  var rawHeaders = data.shift();
  if (!rawHeaders || rawHeaders.length === 0) return [];
  var headers = rawHeaders.map(function (h) { return String(h || '').trim(); });
  return data.map(function (row) {
    var obj = {};
    headers.forEach(function (h, i) {
      if (!h) return;
      var val = toSafeValue_(row[i]);
      obj[h] = val;
      // Also provide normalized key for resilient property access
      var normKey = h.toLowerCase().replace(/[\s_-]+/g, '');
      if (!obj.hasOwnProperty(normKey)) {
        obj[normKey] = val;
      }
    });
    return obj;
  });
}

/**
 * IMPORTANT — this is what makes data reach the browser at all.
 *
 * Google Apps Script cannot send a Date object from the server to the
 * browser through google.script.run. If even ONE date is buried anywhere
 * in the response, the entire response silently fails and the browser
 * gets nothing — which is why the Clients list and Follow-ups list came
 * back empty while the (date-free) login and dashboard numbers worked.
 *
 * So every date coming out of the sheet is turned into plain text here,
 * before it's ever sent anywhere.
 */
function toSafeValue_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return value;
}

/** Finds one row by an ID column. Returns { rowIndex, headers, values } or null. */
function findRowById_(sheet, idColumnName, idValue) {
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var idCol = headers.indexOf(idColumnName);
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(idValue)) {
      return { rowIndex: i + 1, headers: headers, values: data[i] };
    }
  }
  return null;
}

/** Generates the next sequential ID (e.g. CID-000042), safe against two people saving at once. */
function generateId_(sheet, prefix, idColumnIndex) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var lastRow = sheet.getLastRow();
    var nextNumber = 1;
    if (lastRow >= 2) {
      var ids = sheet.getRange(2, idColumnIndex, lastRow - 1, 1).getValues();
      var max = 0;
      ids.forEach(function (row) {
        var id = row[0];
        if (id) {
          var num = parseInt(String(id).split('-')[1], 10);
          if (!isNaN(num) && num > max) max = num;
        }
      });
      nextNumber = max + 1;
    }
    var padded = ('000000' + nextNumber).slice(-6);
    return prefix + '-' + padded;
  } finally {
    lock.releaseLock();
  }
}

function formatDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

/** Standard success/failure shape every write function returns. */
function ok_(message, data) {
  return { success: true, message: message, data: data || null };
}
function fail_(message) {
  return { success: false, message: message, data: null };
}
