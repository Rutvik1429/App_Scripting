/**
 * ALPINO CRM — ONE-TIME DATABASE SETUP SCRIPT
 * ---------------------------------------------
 * What this does: builds the 4 tabs (Clients, Followups, Users, Settings)
 * in this spreadsheet, with the correct headers, frozen header row,
 * dropdown option lists, and data validation wired up.
 *
 * Run this ONCE, right after pasting it in. Safe to re-run later if you
 * haven't entered real data yet — but re-running WIPES any rows you've
 * already added to Clients / Followups / Users, because it clears each
 * sheet before rebuilding it. Duplicate the spreadsheet first if you want
 * a safety copy before re-running.
 */

const CLIENTS_HEADERS = [
  'Client_ID', 'Client_Name', 'Company', 'Phone', 'Email',
  'Address', 'City', 'State', 'Source', 'Industry',
  'Assigned_To', 'Status', 'Priority', 'Created_Date', 'Next_Followup', 'Notes'
];

const FOLLOWUPS_HEADERS = [
  'Followup_ID', 'Client_ID', 'Client_Name', 'Followup_Date', 'Followup_Time',
  'Type', 'Priority', 'Status', 'Assigned_To', 'Notes', 'Created_Date',
  'Outcome_Notes'
];

const USERS_HEADERS = [
  'User_ID', 'Name', 'Email', 'Role', 'Status', 'Created_Date', 'Password'
];

const SETTINGS_HEADERS = [
  'Statuses', 'Sources', 'Industries', 'Priorities',
  'FollowupTypes', 'FollowupStatuses', 'Roles', 'ActiveInactive'
];

const MAX_VALIDATION_ROWS = 1000; // dropdowns apply to rows 2 through 1001

/**
 * The one function you actually run. Everything else below is a helper.
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  createOrResetSheet_(ss, 'Clients', CLIENTS_HEADERS);
  createOrResetSheet_(ss, 'Followups', FOLLOWUPS_HEADERS);
  createOrResetSheet_(ss, 'Users', USERS_HEADERS);
  createOrResetSheet_(ss, 'Settings', SETTINGS_HEADERS);

  populateSettings_(ss);
  populateDefaultAdmin_(ss);
  applyValidation_(ss);
  removeDefaultSheetIfEmpty_(ss);

  Logger.log('Database setup complete — Clients, Followups, Users, and Settings tabs are ready. Default Admin: Admin (admin@alpino.crm / admin123)');
}

/** Creates a sheet if it doesn't exist yet, or clears it if it does, then writes the header row. */
function createOrResetSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(name);
  }
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#e3efe9');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  return sheet;
}

/** Fills the Settings tab with the option lists every dropdown reads from. */
function populateSettings_(ss) {
  const sheet = ss.getSheetByName('Settings');
  const columns = [
    ['New', 'Contacted', 'Active', 'Inactive', 'Lost'],                                   // A: Statuses
    ['Referral', 'Website Inquiry', 'Phone Inquiry', 'Walk-in', 'Social Media', 'Trade Show', 'Other'], // B: Sources
    ['Retailer', 'Distributor', 'Gym / Fitness Studio', 'E-commerce', 'Individual Customer', 'Other'],  // C: Industries
    ['Low', 'Medium', 'High'],                                                             // D: Priorities
    ['Call', 'Email', 'Meeting', 'Visit'],                                                 // E: FollowupTypes
    ['Pending', 'Completed', 'Cancelled'],                                                 // F: FollowupStatuses
    ['Admin', 'Sales Rep'],                                                                // G: Roles
    ['Active', 'Inactive']                                                                 // H: ActiveInactive
  ];
  columns.forEach((values, i) => {
    const colIndex = i + 1;
    sheet.getRange(2, colIndex, values.length, 1).setValues(values.map(v => [v]));
  });
}

/** Wires up every dropdown field to the right source range. */
function applyValidation_(ss) {
  const clients = ss.getSheetByName('Clients');
  const followups = ss.getSheetByName('Followups');
  const users = ss.getSheetByName('Users');
  const settings = ss.getSheetByName('Settings');

  // Clients tab
  setDropdown_(clients, 9, settings.getRange('B2:B50'));   // Source
  setDropdown_(clients, 10, settings.getRange('C2:C50'));  // Industry
  setDropdown_(clients, 11, users.getRange('B2:B100'));    // Assigned_To -> Users' Name column
  setDropdown_(clients, 12, settings.getRange('A2:A50'));  // Status
  setDropdown_(clients, 13, settings.getRange('D2:D50'));  // Priority

  // Followups tab
  setDropdown_(followups, 6, settings.getRange('E2:E50')); // Type
  setDropdown_(followups, 7, settings.getRange('D2:D50')); // Priority
  setDropdown_(followups, 8, settings.getRange('F2:F50')); // Status
  setDropdown_(followups, 9, users.getRange('B2:B100'));   // Assigned_To

  // Users tab
  setDropdown_(users, 4, settings.getRange('G2:G50'));     // Role
  setDropdown_(users, 5, settings.getRange('H2:H50'));     // Status
}

/** Applies a "pick from this range" dropdown rule to one column, rows 2 through MAX_VALIDATION_ROWS. */
function setDropdown_(sheet, columnIndex, sourceRange) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sourceRange, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, columnIndex, MAX_VALIDATION_ROWS, 1).setDataValidation(rule);
}

/** Removes the blank default "Sheet1" Google gives every new spreadsheet, if it's still empty. */
function removeDefaultSheetIfEmpty_(ss) {
  const sheet1 = ss.getSheetByName('Sheet1');
  if (sheet1 && sheet1.getLastRow() === 0 && sheet1.getLastColumn() === 0) {
    ss.deleteSheet(sheet1);
  }
}

/** Seeds the default Admin user with initial credentials. */
function populateDefaultAdmin_(ss) {
  const users = ss.getSheetByName('Users');
  const tz = Session.getScriptTimeZone();
  const now = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
  users.appendRow(['UID-000001', 'Admin', 'admin@alpino.crm', 'Admin', 'Active', now, 'admin123']);
}
