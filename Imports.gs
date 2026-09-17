/**
 * IMPORT HISTORY
 * ---------------
 * Keeps a record of every bulk import: which file, who ran it, how many rows
 * went in, how many were skipped and exactly why — and lets a whole import be
 * undone in one click if it turns out to be wrong.
 *
 * Two pieces make the undo possible:
 *   1. An "Imports" tab, one row per import run.
 *   2. An "Import_ID" column on the Clients tab, stamped onto every row that
 *      came from an import. Undoing = finding rows with that ID and marking
 *      them Deleted (the same recoverable delete the app uses everywhere else).
 */

var IMPORTS_HEADERS = [
  'Import_ID', 'File_Name', 'Sheet_Name', 'Imported_By', 'Imported_At',
  'Total_Rows', 'Imported_Count', 'Skipped_Duplicates', 'Skipped_Invalid',
  'Default_Source', 'Status', 'Skip_Details'
];

/** Creates the Imports tab if it isn't there yet. Safe to call every time. */
function ensureImportsSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Imports');

  if (!sheet) {
    sheet = ss.insertSheet('Imports');
    sheet.getRange(1, 1, 1, IMPORTS_HEADERS.length)
      .setValues([IMPORTS_HEADERS])
      .setFontWeight('bold')
      .setBackground('#e3efe9');
    sheet.setFrozenRows(1);
    SpreadsheetApp.flush();
  }
  return sheet;
}

/**
 * Adds the Import_ID column to the Clients tab if it's missing. Appended at the
 * far right, so it never disturbs the existing column order — and everything
 * that writes to it looks the column up by name, not by position.
 */
function ensureImportIdColumn_(clientsSheet) {
  var headers = clientsSheet.getRange(1, 1, 1, clientsSheet.getLastColumn()).getValues()[0];
  if (headers.indexOf('Import_ID') !== -1) return;

  var col = clientsSheet.getLastColumn() + 1;
  clientsSheet.getRange(1, col)
    .setValue('Import_ID')
    .setFontWeight('bold')
    .setBackground('#e3efe9');
  SpreadsheetApp.flush();
}

/** The import history, newest first. */
function getImports(callerEmail) {
  requireAuth_(callerEmail);
  ensureImportsSheet_();

  var rows = sheetToObjects_(getSheet_('Imports')) || [];
  return rows
    .filter(function (r) { return r.Import_ID; })
    .map(function (r) {
      return {
        Import_ID: r.Import_ID,
        File_Name: r.File_Name,
        Sheet_Name: r.Sheet_Name,
        Imported_By: r.Imported_By,
        Imported_At: r.Imported_At,
        Total_Rows: Number(r.Total_Rows) || 0,
        Imported_Count: Number(r.Imported_Count) || 0,
        Skipped_Duplicates: Number(r.Skipped_Duplicates) || 0,
        Skipped_Invalid: Number(r.Skipped_Invalid) || 0,
        Default_Source: r.Default_Source,
        Status: r.Status || 'Active'
      };
    })
    .reverse();
}

/** The per-row skip reasons for one import — what the Details view shows. */
function getImportDetails(callerEmail, importId) {
  requireAuth_(callerEmail);
  var found = findRowById_(getSheet_('Imports'), 'Import_ID', importId);
  if (!found) throw new Error('That import could not be found.');

  var obj = {};
  found.headers.forEach(function (h, i) { obj[h] = toSafeValue_(found.values[i]); });

  var details = [];
  try {
    details = JSON.parse(obj.Skip_Details || '[]');
  } catch (e) {
    details = [];
  }

  return {
    Import_ID: obj.Import_ID,
    File_Name: obj.File_Name,
    Imported_Count: Number(obj.Imported_Count) || 0,
    Skipped_Duplicates: Number(obj.Skipped_Duplicates) || 0,
    Skipped_Invalid: Number(obj.Skipped_Invalid) || 0,
    Status: obj.Status || 'Active',
    skipped: details
  };
}

/**
 * Undoes an import: every client stamped with this Import_ID is marked Deleted
 * (recoverable — the rows stay in the sheet), and the history entry is flagged
 * "Rolled back" so it's clear what happened.
 */
function undoImport(callerEmail, importId) {
  requireAuth_(callerEmail);

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var importsSheet = getSheet_('Imports');
    var entry = findRowById_(importsSheet, 'Import_ID', importId);
    if (!entry) throw new Error('That import could not be found.');

    var statusCol = entry.headers.indexOf('Status') + 1;
    if (statusCol > 0 && String(entry.values[statusCol - 1]) === 'Rolled back') {
      return fail_('That import has already been rolled back.');
    }

    var clients = getSheet_('Clients');
    var data = clients.getDataRange().getValues();
    var headers = data[0];

    var importCol = headers.indexOf('Import_ID');
    var clientStatusCol = headers.indexOf('Status');
    if (importCol === -1) throw new Error('The Clients sheet has no Import_ID column, so this import cannot be traced.');
    if (clientStatusCol === -1) throw new Error('The Clients sheet has no Status column.');

    var removed = 0;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][importCol]).trim() !== String(importId).trim()) continue;
      if (String(data[i][clientStatusCol]).trim() === 'Deleted') continue;

      var cell = clients.getRange(i + 1, clientStatusCol + 1);
      cell.setDataValidation(null); // "Deleted" isn't in the Status dropdown
      cell.setValue('Deleted');
      removed++;
    }

    if (statusCol > 0) importsSheet.getRange(entry.rowIndex, statusCol).setValue('Rolled back');
    SpreadsheetApp.flush();

    return ok_(
      removed > 0
        ? removed + ' client(s) from this import were removed.'
        : 'No active clients were left from this import.',
      { removed: removed }
    );
  } catch (e) {
    return fail_(e.message || 'Could not undo this import.');
  } finally {
    lock.releaseLock();
  }
}

/** Writes one row into the Imports tab. Called at the end of a bulk import. */
function logImport_(record) {
  var sheet = ensureImportsSheet_();
  var id = generateId_(sheet, 'IMP', 1);

  // Keep the stored detail sensible in size — the counts above always tell the
  // full story, this list is for showing examples of what was skipped.
  var trimmed = (record.skipped || []).slice(0, 200);

  sheet.appendRow([
    id,
    record.fileName || '(unnamed file)',
    record.sheetName || '',
    record.importedBy || '',
    formatDate_(new Date()),
    record.totalRows || 0,
    record.importedCount || 0,
    record.skippedDuplicates || 0,
    record.skippedInvalid || 0,
    record.defaultSource || '',
    'Active',
    JSON.stringify(trimmed)
  ]);

  SpreadsheetApp.flush();
  return id;
}
