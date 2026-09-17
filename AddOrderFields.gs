/**
 * ONE-TIME MIGRATION — run this once, then ignore or delete this file.
 * Adds Order_Value and Order_Count columns to the Clients tab, so
 * converted deals have somewhere to live (matching what your old
 * "sales" tab tracked).
 */
function addOrderFields() {
  var sheet = getSheet_('Clients');
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var toAdd = ['Order_Value', 'Order_Count'].filter(function (h) { return headers.indexOf(h) === -1; });

  if (toAdd.length === 0) {
    Logger.log('Order_Value and Order_Count already exist — nothing to do.');
    return;
  }

  var startCol = sheet.getLastColumn() + 1;
  toAdd.forEach(function (header, i) {
    sheet.getRange(1, startCol + i).setValue(header).setFontWeight('bold').setBackground('#e3efe9');
  });
  Logger.log('Added: ' + toAdd.join(', '));
}
