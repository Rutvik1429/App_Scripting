/**
 * Read-only aggregation for the Dashboard panel. Never writes anything.
 */

function getDashboardStats(callerEmail) {
  requireAuth_(callerEmail);
  var clients = getClients(callerEmail);
  var followups = getFollowups(callerEmail);
  return computeDashboardStats_(clients, followups);
}

/**
 * Loads clients, follow-ups, users, and the dashboard stats in ONE call,
 * all computed from the exact same read of the sheets. This is what the
 * app actually uses on every load/refresh now — fetching them as separate
 * calls (the old way) meant the Dashboard's numbers and the Clients/
 * Follow-ups lists could occasionally be read a moment apart and disagree
 * with each other, which is the "Total Clients: 1 but list is empty" bug.
 */
function getAppData(callerEmail) {
  requireAuth_(callerEmail);

  // The Clients sheet is read ONCE here and used for both the client list and
  // the diagnostic counts below. It used to be read twice per page load —
  // once via getClients() and again for the counts — which is slow once you
  // have a thousand-plus rows.
  var allClientRows = sheetToObjects_(getSheet_('Clients')) || [];
  var clients = allClientRows.filter(function (c) { return c.Status !== 'Deleted'; });

  var followups = getFollowups(callerEmail);
  var users = getUsers(callerEmail);
  var stats = computeDashboardStats_(clients, followups);

  // Import history. Wrapped, so a problem here can never take down the app load.
  var imports = [];
  try {
    imports = getImports(callerEmail);
  } catch (e) {
    imports = [];
  }

  // Diagnostic count, so the Clients screen can explain itself when it looks
  // empty — e.g. rows that exist in the sheet but are hidden because their
  // Status column says "Deleted". Derived from the read above, not a new one.
  var deletedRows = allClientRows.length - clients.length;

  // The stats are included as plain top-level fields (not nested inside a
  // "stats" object) — Google Apps Script's browser bridge can be unreliable
  // about sending back an object nested inside another object, so this
  // keeps everything flat and simple to avoid that entirely.
  return {
    clients: clients,
    followups: followups,
    users: users,
    totalClients: stats.totalClients,
    activeClients: stats.activeClients,
    newThisMonth: stats.newThisMonth,
    followupsToday: stats.followupsToday,
    overdue: stats.overdue,
    completedThisWeek: stats.completedThisWeek,
    clientRowsInSheet: allClientRows.length,
    deletedClientRows: deletedRows,
    imports: imports
  };
}

function computeDashboardStats_(clients, followups) {
  var tz = Session.getScriptTimeZone();
  var today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var monthStart = Utilities.formatDate(new Date(new Date().getFullYear(), new Date().getMonth(), 1), tz, 'yyyy-MM-dd');
  var weekAgo = new Date();
  weekAgo.setDate(weekAgo.getDate() - 7);

  function asDateStr(v) {
    return v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : String(v || '');
  }

  var totalClients = clients.length;
  var activeClients = clients.filter(function (c) { return c.Status === 'Active'; }).length;
  var newThisMonth = clients.filter(function (c) { return asDateStr(c.Created_Date) >= monthStart; }).length;

  var followupsToday = followups.filter(function (f) {
    return f.Status === 'Pending' && asDateStr(f.Followup_Date) === today;
  }).length;

  var overdue = followups.filter(function (f) {
    var d = asDateStr(f.Followup_Date);
    return f.Status === 'Pending' && d && d < today;
  }).length;

  var completedThisWeek = followups.filter(function (f) {
    if (f.Status !== 'Completed') return false;
    var d = f.Created_Date instanceof Date ? f.Created_Date : new Date(f.Created_Date);
    return !isNaN(d) && d >= weekAgo;
  }).length;

  return {
    totalClients: totalClients,
    activeClients: activeClients,
    newThisMonth: newThisMonth,
    followupsToday: followupsToday,
    overdue: overdue,
    completedThisWeek: completedThisWeek
  };
}
