/**
 * Time-driven check for due follow-ups. Right now it only logs what it
 * finds — actually emailing people is a deliberate later step, once
 * you've decided who should get pinged and how.
 */

function scanFollowups() {
  var today = getTodaysFollowups_();
  var overdue = getOverdueFollowups_();
  Logger.log('Follow-ups due today: ' + today.length + ' | Overdue: ' + overdue.length);
}

/**
 * Run this ONCE (manually, from the function dropdown) to schedule
 * scanFollowups() to run automatically every morning at 8am.
 * Re-running it is safe — it removes any old trigger first so you
 * never end up with duplicates.
 */
function createDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'scanFollowups') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('scanFollowups').timeBased().everyDays(1).atHour(8).create();
  Logger.log('Daily 8am check for due follow-ups is now scheduled.');
}
