/**
 * ONE-TIME MIGRATION — run this once, then you can ignore or delete
 * this file. It adds a "Password" column to your existing Users tab
 * without touching any of the rows you've already added.
 */
function addPasswordColumn() {
  var sheet = getSheet_('Users');
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  if (headers.indexOf('Password') !== -1) {
    Logger.log('Password column already exists — nothing to do.');
    return;
  }

  var col = sheet.getLastColumn() + 1;
  sheet.getRange(1, col).setValue('Password').setFontWeight('bold').setBackground('#e3efe9');
  Logger.log('Password column added. Now open the Users tab and type an initial password for each active user in that new column.');
}
