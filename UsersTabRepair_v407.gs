/**
 * UsersTabRepair_v408
 *
 * Use this when users cannot login after sheet cleanup.
 *
 * This script only touches the Users tab.
 * It does not delete any tabs.
 * It does not clean other sheets.
 *
 * What it does:
 * 1. Fixes Users tab headers.
 * 2. Keeps all existing user rows.
 * 3. Adds missing password_hash and temporary_password columns.
 * 4. Lets you reset one user password safely.
 * 5. Lets you create or repair one Super Admin user.
 * 6. Protects Users tab so only owner or listed Super Admin editor can edit.
 *
 * First run:
 * repairUsersTabStructureV407()
 *
 * Then run one of these:
 * resetUserPasswordByUsernameV407('your_username_here', 'NewPassword123')
 * ensureSuperAdminUserV407('admin', 'Your Name', 'your@email.com', 'NewPassword123')
 */


// FIX #8: Hardcoded personal email removed from production code.
// Set this to the Google account email of the Super Admin who should be
// able to edit the Users sheet directly in the Spreadsheet UI.
// Leave as empty string '' to skip adding an extra editor beyond the script owner.
const USERS_TAB_SUPER_ADMIN_EDITOR_EMAIL_V407 = PropertiesService
  .getScriptProperties()
  .getProperty('USERS_TAB_EDITOR_EMAIL') || '';


/**
 * Main repair function.
 * Run this first.
 */
function repairUsersTabStructureV407() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Users');

  if (!sheet) {
    sheet = ss.insertSheet('Users');
  }

  const requiredHeaders = getUsersRequiredHeadersV407_();
  const oldValues = sheet.getDataRange().getValues();
  const oldHeaders = oldValues.length ? oldValues[0].map(v => String(v || '').trim()) : [];
  const oldRows = oldValues.length > 1 ? oldValues.slice(1) : [];

  const newRows = oldRows
    .filter(row => row.some(cell => String(cell || '').trim() !== ''))
    .map(row => {
      return requiredHeaders.map(header => {
        const oldIndex = findUsersHeaderIndexV407_(oldHeaders, header);
        return oldIndex >= 0 ? row[oldIndex] : '';
      });
    });

  removeUsersFilterSafelyV407_(sheet);
  sheet.clear();

  sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);

  if (newRows.length) {
    sheet.getRange(2, 1, newRows.length, requiredHeaders.length).setValues(newRows);
  }

  formatUsersTabV407_(sheet, requiredHeaders);
  protectUsersTabV407_();

  Logger.log('Users tab repaired. Rows kept: ' + newRows.length);
}


/**
 * Reset password for one existing user by username.
 * This writes password_hash and temporary_password.
 */
function resetUserPasswordByUsernameV407(username, newPassword) {
  username = String(username || '').trim();
  newPassword = String(newPassword || '').trim();

  if (!username) throw new Error('Username is required.');
  if (!newPassword || newPassword.length < 8) {
    throw new Error('Password must be at least 8 characters.');
  }

  repairUsersTabStructureV407();

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h || '').trim());

  const usernameCol = getUsersColV407_(headers, 'username');
  const passwordHashCol = getUsersColV407_(headers, 'password_hash');
  const tempPasswordCol = getUsersColV407_(headers, 'temporary_password');
  const statusCol = getUsersColV407_(headers, 'status');
  const updatedAtCol = getUsersColV407_(headers, 'updated_at');

  let foundRow = 0;

  for (let r = 2; r <= data.length; r++) {
    const rowUsername = String(sheet.getRange(r, usernameCol).getValue() || '').trim();
    if (rowUsername.toLowerCase() === username.toLowerCase()) {
      foundRow = r;
      break;
    }
  }

  if (!foundRow) {
    throw new Error('User not found: ' + username);
  }

  const hash = makePortalPasswordHashV407_(newPassword);

  sheet.getRange(foundRow, passwordHashCol).setValue(hash);
  sheet.getRange(foundRow, tempPasswordCol).setValue(newPassword);
  sheet.getRange(foundRow, statusCol).setValue('ACTIVE');
  sheet.getRange(foundRow, updatedAtCol).setValue(new Date());

  Logger.log('Password reset for username: ' + username);
  return {
    success: true,
    username: username,
    temporaryPassword: newPassword
  };
}


/**
 * Create or repair Super Admin user.
 * Use this if you are locked out.
 */
function ensureSuperAdminUserV407(username, fullName, email, newPassword) {
  username = String(username || '').trim();
  fullName = String(fullName || '').trim();
  email = String(email || '').trim();
  newPassword = String(newPassword || '').trim();

  if (!username) throw new Error('Username is required.');
  if (!fullName) throw new Error('Full name is required.');
  if (!email) throw new Error('Email is required.');
  if (!newPassword || newPassword.length < 8) {
    throw new Error('Password must be at least 8 characters.');
  }

  repairUsersTabStructureV407();

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h || '').trim());

  const cols = {};
  getUsersRequiredHeadersV407_().forEach(h => cols[h] = getUsersColV407_(headers, h));

  let foundRow = 0;

  for (let r = 2; r <= data.length; r++) {
    const rowUsername = String(sheet.getRange(r, cols.username).getValue() || '').trim();
    if (rowUsername.toLowerCase() === username.toLowerCase()) {
      foundRow = r;
      break;
    }
  }

  if (!foundRow) {
    foundRow = Math.max(sheet.getLastRow() + 1, 2);
    sheet.getRange(foundRow, cols.user_id).setValue('USR-' + new Date().getTime());
    sheet.getRange(foundRow, cols.created_at).setValue(new Date());
  }

  sheet.getRange(foundRow, cols.username).setValue(username);
  sheet.getRange(foundRow, cols.full_name).setValue(fullName);
  sheet.getRange(foundRow, cols.email).setValue(email);
  sheet.getRange(foundRow, cols.role_code).setValue('SUPER_ADMIN');
  sheet.getRange(foundRow, cols.status).setValue('ACTIVE');
  sheet.getRange(foundRow, cols.password_hash).setValue(makePortalPasswordHashV407_(newPassword));
  sheet.getRange(foundRow, cols.temporary_password).setValue(newPassword);
  sheet.getRange(foundRow, cols.updated_at).setValue(new Date());

  protectUsersTabV407_();

  Logger.log('Super Admin repaired or created.');
  Logger.log('Username: ' + username);
  Logger.log('Temporary password: ' + newPassword);

  return {
    success: true,
    username: username,
    temporaryPassword: newPassword
  };
}


/**
 * Protect Users tab.
 * Owner can always edit. Added email can edit if permissions allow it.
 * FIX #8: Editor email is read from Script Properties, not hardcoded.
 * To set: go to Apps Script > Project Settings > Script Properties,
 * add key USERS_TAB_EDITOR_EMAIL with the Super Admin Google email as value.
 */
function protectUsersTabV407_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Users');
  if (!sheet) return;

  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  protections.forEach(p => {
    try {
      p.remove();
    } catch (error) {}
  });

  const protection = sheet.protect();
  protection.setDescription('Protected Users tab. Edit only by Super Admin.');

  try {
    protection.setWarningOnly(false);
  } catch (error) {}

  try {
    const me = Session.getEffectiveUser().getEmail();
    if (me) protection.addEditor(me);
  } catch (error) {}

  try {
    const editorEmail = USERS_TAB_SUPER_ADMIN_EDITOR_EMAIL_V407;
    if (editorEmail) {
      protection.addEditor(editorEmail);
    }
  } catch (error) {
    Logger.log('Could not add Super Admin editor: ' + error.message);
  }

  try {
    const me = Session.getEffectiveUser().getEmail();
    const editorEmail = USERS_TAB_SUPER_ADMIN_EDITOR_EMAIL_V407;
    const editors = protection.getEditors();
    editors.forEach(editor => {
      const email = editor.getEmail();
      if (
        email &&
        email !== me &&
        email !== editorEmail
      ) {
        protection.removeEditor(editor);
      }
    });
  } catch (error) {}
}


/**
 * Do not change header names unless backend is also updated.
 */
function getUsersRequiredHeadersV407_() {
  return [
    'user_id',
    'username',
    'full_name',
    'email',
    'phone',
    'role_code',
    'employee_id',
    'restaurant_id',
    'assigned_restaurant_ids',
    'region',
    'status',
    'password_hash',
    'temporary_password',
    'created_at',
    'updated_at',
    'last_login_at',
    'notes'
  ];
}


function formatUsersTabV407_(sheet, headers) {
  sheet.setFrozenRows(1);

  const lastRow = Math.max(sheet.getLastRow(), 2);
  const lastCol = headers.length;

  sheet.getRange(1, 1, 1, lastCol)
    .setFontWeight('bold')
    .setBackground('#163a63')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  sheet.getRange(1, 1, lastRow, lastCol)
    .setVerticalAlignment('middle')
    .setWrap(true);

  headers.forEach((header, index) => {
    sheet.setColumnWidth(index + 1, getUsersColumnWidthV407_(header));
  });

  removeUsersFilterSafelyV407_(sheet);
  sheet.getRange(1, 1, lastRow, lastCol).createFilter();
}


function getUsersColumnWidthV407_(header) {
  const h = String(header || '').toLowerCase();

  if (h === 'password_hash') return 320;
  if (h === 'temporary_password') return 190;
  if (h.includes('email')) return 240;
  if (h.includes('name')) return 220;
  if (h.includes('restaurant')) return 220;
  if (h.includes('role')) return 180;
  if (h.includes('phone')) return 150;
  if (h.includes('date') || h.includes('_at')) return 180;
  if (h.includes('notes')) return 260;
  if (h.includes('id')) return 170;

  return 150;
}


function removeUsersFilterSafelyV407_(sheet) {
  try {
    const filter = sheet.getFilter();
    if (filter) filter.remove();
  } catch (error) {
    Logger.log('Could not remove filter: ' + error.message);
  }
}


function findUsersHeaderIndexV407_(oldHeaders, newHeader) {
  const target = normalizeUsersHeaderV407_(newHeader);

  for (let i = 0; i < oldHeaders.length; i++) {
    if (normalizeUsersHeaderV407_(oldHeaders[i]) === target) return i;
  }

  const aliases = {
    user_id: ['userid', 'user_id', 'id'],
    username: ['username', 'user_name', 'login', 'login_name'],
    full_name: ['fullname', 'full_name', 'name'],
    email: ['email', 'email_address', 'emailaddress'],
    phone: ['phone', 'mobile', 'mobile_number'],
    role_code: ['rolecode', 'role_code', 'role', 'role_selection'],
    employee_id: ['employeeid', 'employee_id'],
    restaurant_id: ['restaurantid', 'restaurant_id', 'primary_restaurant'],
    assigned_restaurant_ids: ['assignedrestaurantids', 'assigned_restaurant_ids', 'assigned_restaurants'],
    password_hash: ['passwordhash', 'password_hash', 'hash'],
    temporary_password: ['temporarypassword', 'temporary_password', 'temp_password', 'password'],
    status: ['status', 'user_status'],
    created_at: ['createdat', 'created_at', 'timestamp'],
    updated_at: ['updatedat', 'updated_at']
  };

  const possible = aliases[target] || [];

  for (let i = 0; i < oldHeaders.length; i++) {
    if (possible.includes(normalizeUsersHeaderV407_(oldHeaders[i]))) return i;
  }

  return -1;
}


function normalizeUsersHeaderV407_(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, '_')
    .replace(/[^a-z0-9_]/g, '');
}


function getUsersColV407_(headers, headerName) {
  const index = headers.findIndex(h => normalizeUsersHeaderV407_(h) === normalizeUsersHeaderV407_(headerName));
  if (index < 0) throw new Error('Missing Users column: ' + headerName);
  return index + 1;
}


/**
 * Password hash helper.
 * Uses your existing project hash function if it exists.
 * If not found, uses SHA-256 fallback.
 */
function makePortalPasswordHashV407_(password) {
  password = String(password || '');

  if (typeof hashPassword_ === 'function') {
    return hashPassword_(password);
  }

  if (typeof hashPassword === 'function') {
    return hashPassword(password);
  }

  if (typeof makePasswordHash_ === 'function') {
    return makePasswordHash_(password);
  }

  if (typeof makePasswordHash === 'function') {
    return makePasswordHash(password);
  }

  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    password,
    Utilities.Charset.UTF_8
  );

  return digest.map(byte => {
    const v = (byte < 0 ? byte + 256 : byte).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}


/**
 * Check Users tab health.
 */
function checkUsersTabLoginHealthV407() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  if (!sheet) throw new Error('Users tab not found.');

  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h || '').trim());

  const usernameCol = getUsersColV407_(headers, 'username');
  const roleCol = getUsersColV407_(headers, 'role_code');
  const statusCol = getUsersColV407_(headers, 'status');
  const hashCol = getUsersColV407_(headers, 'password_hash');
  const tempCol = getUsersColV407_(headers, 'temporary_password');

  let total = 0;
  let active = 0;
  let missingHash = 0;
  let superAdmins = 0;

  for (let r = 2; r <= data.length; r++) {
    const username = String(sheet.getRange(r, usernameCol).getValue() || '').trim();
    if (!username) continue;

    total++;

    const status = String(sheet.getRange(r, statusCol).getValue() || '').trim().toUpperCase();
    const role = String(sheet.getRange(r, roleCol).getValue() || '').trim().toUpperCase();
    const hash = String(sheet.getRange(r, hashCol).getValue() || '').trim();
    const temp = String(sheet.getRange(r, tempCol).getValue() || '').trim();

    if (status === 'ACTIVE') active++;
    if (role === 'SUPER_ADMIN') superAdmins++;
    if (!hash && !temp) missingHash++;
  }

  Logger.log('Total users: ' + total);
  Logger.log('Active users: ' + active);
  Logger.log('Super Admin users: ' + superAdmins);
  Logger.log('Users missing password_hash and temporary_password: ' + missingHash);

  return {
    totalUsers: total,
    activeUsers: active,
    superAdmins: superAdmins,
    usersMissingPassword: missingHash
  };
}
