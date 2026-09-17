/**
 * Server-side validation. The browser does light checks too (required
 * fields, input types), but this is the version that actually counts —
 * never trust the browser alone.
 */

function validateRequired_(value, fieldName) {
  if (value === undefined || value === null || String(value).trim() === '') {
    throw new Error(fieldName + ' is required.');
  }
}

/**
 * Turns a phone number written any which way into one consistent stored form.
 *
 * Indian numbers are stored as bare 10 digits, the way the rest of your list
 * looks. Foreign numbers keep their country code, because without it nobody
 * can actually dial them:
 *
 *   "+91 82384 38286"  -> 8238438286      (India)
 *   "082384 38286"     -> 8238438286      (India, trunk 0 removed)
 *   "8238438286"       -> 8238438286      (India)
 *   "p:+19527152132"   -> +19527152132    (USA)
 *   "+971 54 498 8755" -> +971544988755   (UAE)
 *
 * Note there's no chopping a fixed number of digits off the front — country
 * codes are 1, 2 or 3 digits long (1 = USA, 91 = India, 971 = UAE), so
 * guessing would quietly corrupt the number.
 *
 * Returns { ok: true, value: '<stored form>' } or { ok: false, reason: '...' }.
 */
function normalizePhone_(raw) {
  var text = String(raw == null ? '' : raw).trim();
  var hadPlus = text.indexOf('+') !== -1;
  var digits = text.replace(/\D/g, '');

  if (!digits) {
    return { ok: false, reason: 'No phone number' };
  }

  // --- Indian numbers, stored bare ---
  if (digits.length === 12 && digits.indexOf('91') === 0) {
    return { ok: true, value: digits.slice(2) };
  }
  if (digits.length === 11 && digits.charAt(0) === '0') {
    return { ok: true, value: digits.slice(1) };
  }
  if (digits.length === 10 && !hadPlus) {
    return { ok: true, value: digits };
  }

  // --- Anything else: keep it international, country code and all ---
  // 8 to 15 digits is the worldwide range (the E.164 standard caps it at 15).
  if (digits.length >= 8 && digits.length <= 15) {
    return { ok: true, value: '+' + digits };
  }

  return {
    ok: false,
    reason: digits.length < 8
      ? 'Phone is only ' + digits.length + ' digits — too short to be a real number'
      : 'Phone is ' + digits.length + ' digits — too long to be a real number'
  };
}

/**
 * Used by the Add / Edit Client form. Throws a message people can act on.
 */
function validatePhone_(phone) {
  validateRequired_(phone, 'Phone number');
  var result = normalizePhone_(phone);

  if (!result.ok) {
    var digits = String(phone).replace(/\D/g, '');
    throw new Error(
      'That phone number doesn\'t look right (' + digits.length + ' digits). ' +
      'Enter 10 digits for an Indian number, or include the country code for a ' +
      'foreign one — for example +971544988755.'
    );
  }
  return result.value;
}

/**
 * Email is optional, but if one is typed it has to be a real address.
 * Any domain is allowed — gmail.com, yahoo.in, or a client's own business
 * domain like sales@nutritionhub.com.
 */
function validateEmail_(email) {
  if (!email || String(email).trim() === '') return;
  var value = String(email).trim();
  var pattern = /^[^\s@]+@[^\s@]+\.[A-Za-z]{2,}$/;
  if (!pattern.test(value)) {
    throw new Error('Please enter a valid email address, like name@company.com.');
  }
}
