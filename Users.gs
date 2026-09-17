/**
 * Just enough Users functionality to populate the login picker and
 * "Assigned To" dropdowns, and let an Admin add teammates from inside
 * the app later. Adding your very FIRST user still has to be done by
 * hand on the Users sheet.
 */

function getUsers(callerEmail) {
  requireAuth_(callerEmail);
  return sheetToObjects_(getSheet_('Users')).filter(function (u) { return u.Status === 'Active'; });
}

function createUser(callerEmail, form) {
  requireAdmin_(callerEmail);
  try {
    validateRequired_(form.Name, 'Name');
    validateRequired_(form.Email, 'Email');
    validateEmail_(form.Email);

    var sheet = getSheet_('Users');
    ensurePasswordColumn_(sheet);

    var id = generateId_(sheet, 'UID', 1);
    var now = formatDate_(new Date());
    var initialPw = form.Password ? String(form.Password).trim() : '123456';
    var pwHash = hashPassword_(initialPw);

    sheet.appendRow([id, form.Name.trim(), form.Email.trim(), form.Role || 'Sales Rep', 'Active', now, pwHash]);
    SpreadsheetApp.flush();

    return ok_('User added successfully.', { User_ID: id });
  } catch (e) {
    return fail_(e.message || 'Unable to add user.');
  }
}
