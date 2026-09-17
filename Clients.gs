/**
 * CRUD for the Clients sheet. Every function's first argument is now
 * callerEmail — the browser sends it automatically with every call
 * (see the run() helper in JavaScript.html) — so we know who's asking
 * without relying on Google's own identity detection.
 */

function getClients(callerEmail) {
  requireAuth_(callerEmail);
  var rows = sheetToObjects_(getSheet_('Clients')) || [];
  return rows.filter(function (c) { return c.Status !== 'Deleted'; });
}

function getClientById(callerEmail, clientId) {
  requireAuth_(callerEmail);
  var found = findRowById_(getSheet_('Clients'), 'Client_ID', clientId);
  if (!found) throw new Error('Client not found.');
  var obj = {};
  // toSafeValue_ converts dates to text — see the note in Utils.gs.
  found.headers.forEach(function (h, i) { obj[h] = toSafeValue_(found.values[i]); });
  return obj;
}

function createClient(callerEmail, form) {
  var user = requireAuth_(callerEmail);
  try {
    validateRequired_(form.Client_Name, 'Client name');
    var phoneDigits = validatePhone_(form.Phone);
    validateEmail_(form.Email);

    var sheet = getSheet_('Clients');
    var id = generateId_(sheet, 'CID', 1);
    var now = formatDate_(new Date());

    sheet.appendRow([
      id,
      form.Client_Name,
      form.Company || '',
      phoneDigits,
      form.Email || '',
      form.Address || '',
      form.City || '',
      form.State || '',
      form.Source || '',
      form.Industry || '',
      form.Assigned_To || user.name,
      form.Status || 'New',
      form.Priority || 'Medium',
      now,
      '',
      form.Notes || ''
    ]);

    // Force the write to commit immediately — without this, a follow-up read
    // (like the Dashboard's own count, or the Clients list refreshing right
    // after this save) can run before the new row is actually visible, which
    // is what was causing "Total Clients: 1" but an empty Clients list.
    SpreadsheetApp.flush();

    return ok_('Client created successfully.', { Client_ID: id });
  } catch (e) {
    return fail_(e.message || 'Unable to save client.');
  }
}

function updateClient(callerEmail, clientId, form) {
  requireAuth_(callerEmail);
  try {
    validateRequired_(form.Client_Name, 'Client name');
    var phoneDigits = validatePhone_(form.Phone);
    validateEmail_(form.Email);

    var sheet = getSheet_('Clients');
    var found = findRowById_(sheet, 'Client_ID', clientId);
    if (!found) throw new Error('Client not found.');

    var updates = {
      Client_Name: form.Client_Name,
      Company: form.Company || '',
      Phone: phoneDigits,
      Email: form.Email || '',
      Address: form.Address || '',
      City: form.City || '',
      State: form.State || '',
      Source: form.Source || '',
      Industry: form.Industry || '',
      Status: form.Status || 'New',
      Priority: form.Priority || 'Medium',
      Notes: form.Notes || ''
    };

    if (form.Assigned_To) {
      updates.Assigned_To = form.Assigned_To;
    }

    Object.keys(updates).forEach(function (key) {
      var colIndex = found.headers.indexOf(key) + 1;
      if (colIndex > 0) sheet.getRange(found.rowIndex, colIndex).setValue(updates[key]);
    });

    SpreadsheetApp.flush();
    return ok_('Client updated successfully.');
  } catch (e) {
    return fail_(e.message || 'Unable to update client.');
  }
}

/** Soft delete — sets Status to Deleted instead of removing the row. */
function deleteClient(callerEmail, clientId) {
  requireAuth_(callerEmail);
  try {
    var sheet = getSheet_('Clients');
    var found = findRowById_(sheet, 'Client_ID', clientId);
    if (!found) throw new Error('Client not found.');
    var statusCol = found.headers.indexOf('Status') + 1;
    if (statusCol < 1) throw new Error('The Clients sheet has no "Status" column.');

    // The Status column has a dropdown that only allows New/Contacted/Active/
    // Inactive/Lost — and "Deleted" isn't on that list, so writing it can be
    // rejected. Clearing the dropdown rule on just this one cell lets the
    // soft-delete go through every time.
    var cell = sheet.getRange(found.rowIndex, statusCol);
    cell.setDataValidation(null);
    cell.setValue('Deleted');

    SpreadsheetApp.flush();
    return ok_('Client removed.');
  } catch (e) {
    return fail_(e.message || 'Unable to remove client.');
  }
}

/**
 * Bulk import multiple leads from CSV / file upload.
 * Batch-processes and batch-writes to the Clients sheet for maximum speed.
 */
function importClientsBulk(callerEmail, leadsList, options) {
  var user = requireAuth_(callerEmail);
  if (!leadsList || !Array.isArray(leadsList) || leadsList.length === 0) {
    return fail_('No leads provided for import.');
  }

  options = options || {};
  var skipDuplicates = options.skipDuplicates !== false;
  var defaultSource = options.defaultSource || '';
  var defaultStatus = options.defaultStatus || 'New';
  var defaultPriority = options.defaultPriority || 'Medium';

  var fileName = options.fileName || '';
  var sheetName = options.sheetName || '';

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getSheet_('Clients');
    ensureImportIdColumn_(sheet); // so this import can be undone later

    var data = sheet.getDataRange().getValues();
    var headers = data[0];

    var phoneColIndex = headers.indexOf('Phone');
    var idColIndex = headers.indexOf('Client_ID');

    // Build lookup set of existing phone numbers for fast duplicate checking
    var existingPhones = {};
    var maxIdNum = 0;

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (phoneColIndex !== -1 && row[phoneColIndex]) {
        // Normalized the same way incoming leads are, so "+919824012345" in the
        // sheet and "9824012345" in the file are recognised as the same person.
        var existing = normalizePhone_(row[phoneColIndex]);
        if (existing.ok) existingPhones[existing.value] = true;
      }
      if (idColIndex !== -1 && row[idColIndex]) {
        var idStr = String(row[idColIndex]);
        var num = parseInt(idStr.split('-')[1], 10);
        if (!isNaN(num) && num > maxIdNum) maxIdNum = num;
      }
    }

    var rowsToInsert = [];
    var skippedDuplicates = 0;
    var skippedInvalid = 0;
    var skipDetails = []; // row-by-row reasons, shown in the import history
    var now = formatDate_(new Date());

    leadsList.forEach(function (lead, leadIndex) {
      // +2 because the file's first row is headings and people count from 1
      var fileRow = leadIndex + 2;

      function noteSkip(reason, name, phone) {
        skipDetails.push({
          row: fileRow,
          name: String(name || '').trim() || '(blank)',
          phone: String(phone || '').trim() || '(blank)',
          reason: reason
        });
      }

      if (!lead || !lead.Client_Name || String(lead.Client_Name).trim() === '') {
        skippedInvalid++;
        noteSkip('No name in the Name column', '', lead && lead.Phone);
        return;
      }

      var rawPhone = lead.Phone || '';
      var phoneResult = normalizePhone_(rawPhone);

      if (!phoneResult.ok) {
        skippedInvalid++;
        noteSkip(phoneResult.reason, lead.Client_Name, rawPhone);
        return;
      }
      var cleanPhone = phoneResult.value;

      if (skipDuplicates && existingPhones[cleanPhone]) {
        skippedDuplicates++;
        noteSkip('This phone number is already in the CRM', lead.Client_Name, cleanPhone);
        return;
      }

      // Track phone to prevent duplicates within the same batch file
      existingPhones[cleanPhone] = true;

      maxIdNum++;
      var padded = ('000000' + maxIdNum).slice(-6);
      var newId = 'CID-' + padded;

      rowsToInsert.push([
        newId,
        String(lead.Client_Name).trim(),
        String(lead.Company || '').trim(),
        cleanPhone,
        String(lead.Email || '').trim(),
        String(lead.Address || '').trim(),
        String(lead.City || '').trim(),
        String(lead.State || '').trim(),
        String(lead.Source || defaultSource || 'Imported').trim(),
        String(lead.Industry || 'Other').trim(),
        String(lead.Assigned_To || user.name).trim(),
        String(lead.Status || defaultStatus).trim(),
        String(lead.Priority || defaultPriority).trim(),
        // The lead's OWN date from the file (Meta's created_time, for example)
        // when the file has one, so date filtering reflects when the lead
        // actually came in — not when you happened to import it. Falls back to
        // the time of import when the file carries no usable date.
        parseImportDate_(lead.Lead_Date) || now,
        '', // Next_Followup
        String(lead.Notes || '').trim()
      ]);
    });

    if (rowsToInsert.length === 0) {
      // Still worth logging — this is exactly when you want to see the reasons.
      try {
        logImport_({
          fileName: fileName, sheetName: sheetName, importedBy: user.name,
          totalRows: leadsList.length, importedCount: 0,
          skippedDuplicates: skippedDuplicates, skippedInvalid: skippedInvalid,
          defaultSource: defaultSource, skipped: skipDetails
        });
      } catch (logErr) { /* logging must never block the answer */ }

      var reason = [];
      if (skippedDuplicates > 0) reason.push(skippedDuplicates + ' duplicate(s) skipped');
      if (skippedInvalid > 0) reason.push(skippedInvalid + ' invalid row(s) skipped (missing name or 10-digit phone)');
      return fail_('No new leads were added. (' + (reason.join(', ') || 'All rows skipped') + ') See the Imports screen for the row-by-row reasons.');
    }

    // The Clients sheet's Source and Industry columns have dropdown rules that
    // REJECT any value not listed on the Settings tab. An import carrying a new
    // source like "Facebook Ads" would be blocked outright, so register any new
    // values on the Settings tab first — that keeps the dropdowns working too.
    try {
      ensureSettingsValues_(2, rowsToInsert.map(function (r) { return r[8]; }));  // B: Sources
      ensureSettingsValues_(3, rowsToInsert.map(function (r) { return r[9]; }));  // C: Industries
    } catch (settingsErr) {
      // Not fatal — the retry below still gets the data in.
    }

    var startRow = sheet.getLastRow() + 1;
    var numCols = rowsToInsert[0].length;
    var target = sheet.getRange(startRow, 1, rowsToInsert.length, numCols);

    try {
      target.setValues(rowsToInsert);
    } catch (writeErr) {
      // Some dropdown rule still refused a value (a name not on the Users tab,
      // for instance). Clear the rules on just these new rows and write again,
      // so an import can never fail outright over a dropdown.
      target.setDataValidation(null);
      target.setValues(rowsToInsert);
    }

    SpreadsheetApp.flush();

    // Record the run, then stamp its ID onto every row just written, so the
    // whole batch can be undone later from the Imports screen.
    var importId = '';
    try {
      importId = logImport_({
        fileName: fileName, sheetName: sheetName, importedBy: user.name,
        totalRows: leadsList.length, importedCount: rowsToInsert.length,
        skippedDuplicates: skippedDuplicates, skippedInvalid: skippedInvalid,
        defaultSource: defaultSource, skipped: skipDetails
      });

      var freshHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      var importColIndex = freshHeaders.indexOf('Import_ID');
      if (importColIndex !== -1 && importId) {
        var stamps = rowsToInsert.map(function () { return [importId]; });
        sheet.getRange(startRow, importColIndex + 1, stamps.length, 1).setValues(stamps);
        SpreadsheetApp.flush();
      }
    } catch (logErr) {
      // The leads are safely in — losing the history entry shouldn't fail the import.
    }

    var summaryMsg = rowsToInsert.length + ' leads imported successfully!';
    if (skippedDuplicates > 0) summaryMsg += ' (' + skippedDuplicates + ' duplicate(s) skipped)';
    if (skippedInvalid > 0) summaryMsg += ' (' + skippedInvalid + ' invalid row(s) skipped)';

    return ok_(summaryMsg, {
      imported: rowsToInsert.length,
      duplicatesSkipped: skippedDuplicates,
      invalidSkipped: skippedInvalid,
      total: leadsList.length,
      importId: importId
    });
  } catch (e) {
    return fail_(e.message || 'Bulk import failed.');
  } finally {
    lock.releaseLock();
  }
}
