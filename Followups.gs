/**
 * CRUD for the Followups sheet, plus the today/overdue queries the
 * Dashboard and the reminder trigger both depend on.
 *
 * Functions ending in _ are internal (no login check) — Triggers.gs
 * needs those, since a time-driven trigger has no "logged in browser"
 * to pass a callerEmail from. Everything else takes callerEmail as its
 * first argument, same as Clients.gs.
 */

/**
 * Adds the Outcome_Notes column to the Followups sheet if it isn't there yet.
 *
 * Why it exists: the note you type when you SCHEDULE a call ("wants the 1kg
 * pack, call after 6pm") and the note you type AFTER the call ("not
 * interested, already buys from a local store") are two different things.
 * They used to be squashed into one Notes cell joined with a "|", which made
 * the call outcome impossible to read or report on separately.
 *
 * Safe to run any number of times — it checks first and does nothing if the
 * column already exists. Nothing is overwritten; old rows keep whatever is
 * in their Notes cell and simply have this new column blank.
 */
function ensureOutcomeNotesColumn_(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.indexOf('Outcome_Notes') !== -1) return;

  var col = sheet.getLastColumn() + 1;
  sheet.getRange(1, col)
    .setValue('Outcome_Notes')
    .setFontWeight('bold')
    .setBackground('#e3efe9');
  SpreadsheetApp.flush();
}

function getFollowups_() {
  var sheet = getSheet_('Followups');
  ensureOutcomeNotesColumn_(sheet);
  return sheetToObjects_(sheet) || [];
}
function getFollowups(callerEmail) {
  requireAuth_(callerEmail);
  return getFollowups_();
}

function getTodaysFollowups_() {
  var tz = Session.getScriptTimeZone();
  var today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  return getFollowups_().filter(function (f) {
    var d = f.Followup_Date instanceof Date ? Utilities.formatDate(f.Followup_Date, tz, 'yyyy-MM-dd') : String(f.Followup_Date || '');
    return d === today && f.Status === 'Pending';
  });
}
function getTodaysFollowups(callerEmail) {
  requireAuth_(callerEmail);
  return getTodaysFollowups_();
}

function getOverdueFollowups_() {
  var tz = Session.getScriptTimeZone();
  var today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  return getFollowups_().filter(function (f) {
    var d = f.Followup_Date instanceof Date ? Utilities.formatDate(f.Followup_Date, tz, 'yyyy-MM-dd') : String(f.Followup_Date || '');
    return d && d < today && f.Status === 'Pending';
  });
}
function getOverdueFollowups(callerEmail) {
  requireAuth_(callerEmail);
  return getOverdueFollowups_();
}

function createFollowup(callerEmail, form) {
  var user = requireAuth_(callerEmail);
  try {
    validateRequired_(form.Client_ID, 'Client');
    validateRequired_(form.Followup_Date, 'Follow-up date');

    var client = getClientById(callerEmail, form.Client_ID);
    var followupDate = new Date(form.Followup_Date + 'T00:00:00');

    var sheet = getSheet_('Followups');
    var id = generateId_(sheet, 'FID', 1);
    var now = formatDate_(new Date());

    sheet.appendRow([
      id,
      form.Client_ID,
      client.Client_Name,
      followupDate,
      form.Followup_Time || '',
      form.Type || 'Call',
      form.Priority || 'Medium',
      'Pending',
      form.Assigned_To || user.name,
      form.Notes || '',
      now
    ]);

    // Keep the client's Next_Followup column in sync so it shows on the Clients table too.
    var clientsSheet = getSheet_('Clients');
    var foundClient = findRowById_(clientsSheet, 'Client_ID', form.Client_ID);
    if (foundClient) {
      var col = foundClient.headers.indexOf('Next_Followup') + 1;
      clientsSheet.getRange(foundClient.rowIndex, col).setValue(followupDate);
    }

    SpreadsheetApp.flush();
    return ok_('Follow-up scheduled.', { Followup_ID: id });
  } catch (e) {
    return fail_(e.message || 'Unable to schedule follow-up.');
  }
}

/**
 * Marks a follow-up done and files the call outcome.
 *
 * The outcome goes into its own Outcome_Notes column, NOT appended onto the
 * original scheduling note — so "what I planned to say" and "what actually
 * happened on the call" stay readable as two separate things. If you complete
 * the same follow-up twice, the second note is added underneath the first
 * with a blank line between, rather than replacing it.
 */
function completeFollowup(callerEmail, followupId, outcomeNotes) {
  requireAuth_(callerEmail);
  try {
    var sheet = getSheet_('Followups');
    ensureOutcomeNotesColumn_(sheet);

    var found = findRowById_(sheet, 'Followup_ID', followupId);
    if (!found) throw new Error('Follow-up not found.');

    var statusCol = found.headers.indexOf('Status') + 1;
    sheet.getRange(found.rowIndex, statusCol).setValue('Completed');

    if (outcomeNotes) {
      var outcomeCol = found.headers.indexOf('Outcome_Notes') + 1;
      var existing = found.values[outcomeCol - 1] || '';
      sheet.getRange(found.rowIndex, outcomeCol)
        .setValue(existing ? existing + '\n\n' + outcomeNotes : outcomeNotes);
    }

    SpreadsheetApp.flush();
    return ok_('Follow-up marked complete.');
  } catch (e) {
    return fail_(e.message || 'Unable to update follow-up.');
  }
}

function deleteFollowup(callerEmail, followupId) {
  requireAuth_(callerEmail);
  try {
    var sheet = getSheet_('Followups');
    var found = findRowById_(sheet, 'Followup_ID', followupId);
    if (!found) throw new Error('Follow-up not found.');
    var statusCol = found.headers.indexOf('Status') + 1;
    sheet.getRange(found.rowIndex, statusCol).setValue('Cancelled');
    SpreadsheetApp.flush();
    return ok_('Follow-up cancelled.');
  } catch (e) {
    return fail_(e.message || 'Unable to cancel follow-up.');
  }
}
