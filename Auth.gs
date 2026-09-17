/**
 * ALPINO CRM — AUTHENTICATION & SESSION MANAGEMENT
 * -------------------------------------------------
 * Handles user login, password hashing (SHA-256), HMAC-signed session tokens,
 * automatic migration/seeding of the default Admin account, and role guards.
 */

var SESSION_SECONDS = 6 * 60 * 60; // 6 hours session lifetime

/** Returns or generates a persistent secret for HMAC token signing. */
function getSessionSecret_() {
  var props = PropertiesService.getScriptProperties();
  var secret = props.getProperty('CRM_AUTH_SECRET');
  if (!secret) {
    secret = Utilities.getUuid() + '-' + Utilities.getUuid();
    props.setProperty('CRM_AUTH_SECRET', secret);
  }
  return secret;
}

/** Computes a standard SHA-256 hex hash for a given password. */
function hashPassword_(password) {
  var clean = String(password || '').trim();
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, clean, Utilities.Charset.UTF_8);
  return bytes.map(function (b) {
    var v = (b + 256) % 256;
    return ('0' + v.toString(16)).slice(-2);
  }).join('');
}

/** Generates a tamper-proof, stateless HMAC-SHA256 session token. */
function generateSessionToken_(user) {
  var expiresAt = new Date().getTime() + (SESSION_SECONDS * 1000);
  var payload = user.Email + '|' + user.Role + '|' + expiresAt + '|' + user.User_ID;
  var secret = getSessionSecret_();
  var sigBytes = Utilities.computeHmacSha256Signature(payload, secret);
  var signature = sigBytes.map(function (b) {
    var v = (b + 256) % 256;
    return ('0' + v.toString(16)).slice(-2);
  }).join('');

  var token = Utilities.base64EncodeWebSafe(payload + '::' + signature);

  // Fast-path in-memory cache
  try {
    CacheService.getScriptCache().put('session_' + token, user.Email, SESSION_SECONDS);
  } catch (e) {}

  return token;
}

/** Verifies a session token signature & expiration. Returns user email or null. */
function verifySessionToken_(token) {
  if (!token) return null;
  try {
    var decoded = Utilities.newBlob(Utilities.base64DecodeWebSafe(token)).getDataAsString();
    var parts = decoded.split('::');
    if (parts.length !== 2) {
      // Fallback check for legacy CacheService token
      return CacheService.getScriptCache().get('session_' + token);
    }

    var payload = parts[0];
    var signature = parts[1];
    var secret = getSessionSecret_();

    var expectedBytes = Utilities.computeHmacSha256Signature(payload, secret);
    var expectedSig = expectedBytes.map(function (b) {
      var v = (b + 256) % 256;
      return ('0' + v.toString(16)).slice(-2);
    }).join('');

    if (signature !== expectedSig) return null;

    var payloadParts = payload.split('|');
    var email = payloadParts[0];
    var expiresAt = parseInt(payloadParts[2], 10);

    if (isNaN(expiresAt) || new Date().getTime() > expiresAt) return null;

    return email;
  } catch (e) {
    // Fallback check for legacy CacheService token
    try {
      return CacheService.getScriptCache().get('session_' + token);
    } catch (err) {
      return null;
    }
  }
}

/** Ensures the Password column exists in the Users sheet. */
function ensurePasswordColumn_(sheet) {
  var data = sheet.getDataRange().getValues();
  if (!data || data.length === 0) return;
  var headers = data[0];
  if (headers.indexOf('Password') === -1) {
    var col = headers.length + 1;
    sheet.getRange(1, col).setValue('Password').setFontWeight('bold').setBackground('#e3efe9');
    SpreadsheetApp.flush();
  }
}

/** Creates the initial default Admin account if the Users sheet is completely empty. */
function seedDefaultAdmin_(sheet) {
  var now = formatDate_(new Date());
  sheet.appendRow(['UID-000001', 'Admin', 'admin@alpino.crm', 'Admin', 'Active', now, 'admin123']);
  SpreadsheetApp.flush();
}

/** Public list for the login screen's name dropdown — no passwords included. */
function getActiveUserList() {
  var sheet = getSheet_('Users');
  ensurePasswordColumn_(sheet);

  var rawUsers = sheetToObjects_(sheet);
  if (!rawUsers || rawUsers.length === 0) {
    seedDefaultAdmin_(sheet);
    rawUsers = sheetToObjects_(sheet);
  }

  return rawUsers
    .filter(function (u) {
      var status = String(u.Status || u.status || '').trim().toLowerCase();
      return status === 'active' || status === '';
    })
    .map(function (u, i) {
      var name = String(u.Name || u.name || '').trim();
      var email = String(u.Email || u.email || '').trim();
      var userId = String(u.User_ID || u.userid || u.id || '').trim();
      var displayName = name || email || userId || ('User ' + (i + 1));
      // Identifier used for value: Email is preferred, fallback to Name or User_ID
      var identifier = email || name || userId || displayName;
      return {
        name: displayName,
        email: identifier,
        userId: userId
      };
    });
}

/** Called when someone submits the login form. Returns a session token. */
function verifyLogin(identifier, password) {
  if (!identifier || String(identifier).trim() === '') {
    throw new Error('Please choose your name from the list.');
  }
  if (!password || String(password).trim() === '') {
    throw new Error('Please enter your password.');
  }

  var cleanId = String(identifier).trim().toLowerCase();
  var cleanPassword = String(password).trim();

  var sheet = getSheet_('Users');
  ensurePasswordColumn_(sheet);

  var data = sheet.getDataRange().getValues();
  if (!data || data.length < 2) {
    seedDefaultAdmin_(sheet);
    data = sheet.getDataRange().getValues();
  }

  var rawHeaders = data[0];
  var headers = rawHeaders.map(function (h) { return String(h || '').trim().toLowerCase().replace(/[\s_-]+/g, ''); });

  var emailCol = headers.indexOf('email');
  var nameCol = headers.indexOf('name');
  var idCol = headers.indexOf('userid');
  if (idCol === -1) idCol = headers.indexOf('id');
  var pwCol = headers.indexOf('password');
  var statusCol = headers.indexOf('status');
  var roleCol = headers.indexOf('role');

  var foundRow = null;
  var rowIndex = -1;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var rowEmail = emailCol !== -1 ? String(row[emailCol] || '').trim().toLowerCase() : '';
    var rowName = nameCol !== -1 ? String(row[nameCol] || '').trim().toLowerCase() : '';
    var rowId = idCol !== -1 ? String(row[idCol] || '').trim().toLowerCase() : '';

    if ((rowEmail && rowEmail === cleanId) ||
        (rowName && rowName === cleanId) ||
        (rowId && rowId === cleanId)) {
      foundRow = row;
      rowIndex = i + 1;
      break;
    }
  }

  if (!foundRow) {
    throw new Error('Account not found for "' + identifier + '". Ask an Admin to check the Users tab.');
  }

  var status = statusCol !== -1 ? String(foundRow[statusCol] || '').trim().toLowerCase() : 'active';
  if (status && status !== 'active') {
    throw new Error('Your account is marked Inactive. Ask an Admin to reactivate it.');
  }

  var stored = pwCol !== -1 ? String(foundRow[pwCol] || '').trim() : '';
  if (!stored) {
    throw new Error('No password set for this account yet. Please enter a password in the Password column of the Users sheet.');
  }

  var enteredHash = hashPassword_(cleanPassword);
  var ok = false;

  if (stored === enteredHash) {
    ok = true;
  } else if (stored === cleanPassword) {
    // First login against a plain password typed in the sheet — auto-scramble it in the sheet
    ok = true;
    if (pwCol !== -1) {
      sheet.getRange(rowIndex, pwCol + 1).setValue(enteredHash);
      SpreadsheetApp.flush();
    }
  }

  if (!ok) {
    throw new Error('Incorrect password. Please try again.');
  }

  var userEmail = emailCol !== -1 && foundRow[emailCol] ? String(foundRow[emailCol]).trim() : '';
  var userName = nameCol !== -1 && foundRow[nameCol] ? String(foundRow[nameCol]).trim() : '';
  var userIdVal = idCol !== -1 && foundRow[idCol] ? String(foundRow[idCol]).trim() : ('UID-' + rowIndex);
  var userRole = roleCol !== -1 && foundRow[roleCol] ? String(foundRow[roleCol]).trim() : 'Sales Rep';

  var primaryEmail = userEmail || (userName ? userName.toLowerCase().replace(/\s+/g, '') + '@alpino.crm' : userIdVal + '@alpino.crm');

  var userObj = {
    Email: primaryEmail,
    Name: userName || userEmail || 'User',
    Role: userRole,
    User_ID: userIdVal
  };

  var token = generateSessionToken_(userObj);
  return {
    token: token,
    email: userObj.Email,
    name: userObj.Name,
    role: userObj.Role,
    userId: userObj.User_ID
  };
}

/**
 * Remembers the result of checking a token — but only for the life of one
 * request.
 *
 * Loading the app calls getAppData, which in turn calls getClients,
 * getFollowups, getUsers and getImports. Every one of those checks the token,
 * and each check was reading the ENTIRE Users sheet — five full reads for a
 * single page load. Now it's read once and reused, which is most of the
 * loading delay gone.
 *
 * This is wiped automatically the moment the request ends, so it can never
 * serve a stale session or keep someone signed in after being deactivated.
 */
var AUTH_CACHE_ = {};

/** Every data function calls this first, passing the session token from the browser. */
function requireAuth_(token) {
  if (!token) {
    throw new Error('Your session has expired — please sign in again.');
  }

  if (AUTH_CACHE_[token]) return AUTH_CACHE_[token];

  var identifier = verifySessionToken_(token);
  if (!identifier) {
    throw new Error('Your session has expired — please sign in again.');
  }

  var cleanId = String(identifier).trim().toLowerCase();
  var sheet = getSheet_('Users');
  var data = sheet.getDataRange().getValues();
  if (!data || data.length < 2) {
    throw new Error('No users found. Please check the Users sheet.');
  }

  var rawHeaders = data[0];
  var headers = rawHeaders.map(function (h) { return String(h || '').trim().toLowerCase().replace(/[\s_-]+/g, ''); });

  var emailCol = headers.indexOf('email');
  var nameCol = headers.indexOf('name');
  var idCol = headers.indexOf('userid');
  if (idCol === -1) idCol = headers.indexOf('id');
  var statusCol = headers.indexOf('status');
  var roleCol = headers.indexOf('role');

  var found = null;
  var rowIndex = -1;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var rowEmail = emailCol !== -1 ? String(row[emailCol] || '').trim().toLowerCase() : '';
    var rowName = nameCol !== -1 ? String(row[nameCol] || '').trim().toLowerCase() : '';
    var rowId = idCol !== -1 ? String(row[idCol] || '').trim().toLowerCase() : '';

    if ((rowEmail && rowEmail === cleanId) ||
        (rowName && rowName === cleanId) ||
        (rowId && rowId === cleanId)) {
      found = row;
      rowIndex = i + 1;
      break;
    }
  }

  if (!found) {
    throw new Error('Your account could not be found. Please check the Users sheet.');
  }

  var status = statusCol !== -1 ? String(found[statusCol] || '').trim().toLowerCase() : 'active';
  if (status && status !== 'active') {
    throw new Error('Your account is no longer active.');
  }

  var userEmail = emailCol !== -1 && found[emailCol] ? String(found[emailCol]).trim() : '';
  var userName = nameCol !== -1 && found[nameCol] ? String(found[nameCol]).trim() : '';
  var userIdVal = idCol !== -1 && found[idCol] ? String(found[idCol]).trim() : ('UID-' + rowIndex);
  var userRole = roleCol !== -1 && found[roleCol] ? String(found[roleCol]).trim() : 'Sales Rep';

  var result = {
    authorized: true,
    email: userEmail || cleanId,
    name: userName || userEmail || 'User',
    role: userRole,
    userId: userIdVal
  };

  AUTH_CACHE_[token] = result; // reused by the other calls in this same request
  return result;
}

/** Used on page load to silently check a saved token without throwing. */
function getSessionUser(token) {
  try {
    return requireAuth_(token);
  } catch (e) {
    return null;
  }
}

/** Guard ensuring only Admins can execute privileged actions. */
function requireAdmin_(token) {
  var user = requireAuth_(token);
  if (user.role !== 'Admin') {
    throw new Error('Only Admins can perform this action.');
  }
  return user;
}

