function resetUserSystem() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  if (!sheet) return;
  
  // 1. Delete all rows whose ID starts with "USER"
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx = headers.indexOf('ID');
  if (idIdx === -1) return;
  
  // Delete from bottom to top to avoid shifting issues
  for (let i = data.length - 1; i > 0; i--) {
    const id = String(data[i][idIdx] || '');
    if (id.startsWith('USER')) {
      sheet.deleteRow(i + 1);
    }
  }
  
  // 2. Reset the counter to 1
  const props = PropertiesService.getScriptProperties();
  props.setProperty('USER_COUNTER', '1');
}
// ==================================================
// CARZOO CRM - User Authentication System (MD5)
// ==================================================

const USER_SHEET_NAME = "Users";
const COUNTER_KEY = "USER_COUNTER";

// ----- Helper: MD5 hash -----
function md5Hash(str) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, str);
  return digest.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

// ----- Get or initialize the user counter -----
function getNextUserId() {
  const props = PropertiesService.getScriptProperties();
  let counter = props.getProperty(COUNTER_KEY);
  if (!counter) {
    // Counter not set – initialize based on existing users
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USER_SHEET_NAME);
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      let maxNum = 0;
      // Find the ID column index
      const headers = data[0].map(h => h.toString().trim());
      const idIdx = headers.findIndex(h => h.toLowerCase() === 'id');
      if (idIdx !== -1) {
        for (let i = 1; i < data.length; i++) {
          const id = String(data[i][idIdx] || '');
          const match = id.match(/^USER(\d+)$/);
          if (match) {
            const num = parseInt(match[1], 10);
            if (num > maxNum) maxNum = num;
          }
        }
      }
      counter = String(maxNum + 1);
    } else {
      counter = "1";
    }
    props.setProperty(COUNTER_KEY, counter);
  }
  // Increment and return new ID number
  const nextNum = parseInt(counter, 10);
  const newCounter = String(nextNum + 1);
  props.setProperty(COUNTER_KEY, newCounter);
  return nextNum;
}

// ----- doPost (for future use, but we use GET now) -----
function doPost(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  try {
    if (!e || !e.postData) {
      output.setContent(JSON.stringify({ status: "error", message: "Invalid request" }));
      return output;
    }
    const sheet = getOrCreateUserSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().trim());
    let request = JSON.parse(e.postData.contents);
    const action = request.action;
    if (action === 'register') return handleRegister(sheet, data, headers, request, output);
    if (action === 'login') return handleLogin(sheet, data, headers, request, output);
    output.setContent(JSON.stringify({ status: "error", message: "Unknown action" }));
    return output;
  } catch (error) {
    output.setContent(JSON.stringify({ status: "error", message: error.toString() }));
    return output;
  }
}

// ----- doGet (main entry) -----
function doGet(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  try {
    if (!e || !e.parameter) {
      output.setContent(JSON.stringify({ status: "error", message: "Invalid request" }));
      return output;
    }
    const sheet = getOrCreateUserSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().trim());
    const action = e.parameter.action;

    if (action === 'login') return handleLogin(sheet, data, headers, e.parameter, output);
    if (action === 'register') return handleRegister(sheet, data, headers, e.parameter, output);
    if (action === 'getUsers') return handleGetUsers(sheet, data, headers, e.parameter, output);
    if (action === 'deleteUser') return handleDeleteUser(sheet, data, headers, e.parameter, output);
    if (action === 'updateUser') return handleUpdateUser(sheet, data, headers, e.parameter, output);

    output.setContent(JSON.stringify({ status: "error", message: "Unknown action: " + action }));
    return output;
  } catch (error) {
    output.setContent(JSON.stringify({ status: "error", message: error.toString() }));
    return output;
  }
}

// ----- getOrCreateUserSheet (with hashed default password) -----
function getOrCreateUserSheet() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USER_SHEET_NAME);
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(USER_SHEET_NAME);
    sheet.appendRow(['ID', 'Name', 'Username', 'Password', 'isAdmin', 'createdAt']);
    // Default admin: password "admin123" hashed with MD5
    const defaultHashed = md5Hash('admin123');
    sheet.appendRow(['ADMIN001', 'Admin', 'admin', defaultHashed, 'TRUE', new Date().toISOString()]);
    // Initialize counter to 1 (for next user after ADMIN001)
    const props = PropertiesService.getScriptProperties();
    props.setProperty(COUNTER_KEY, '1');
  }
  return sheet;
}

// ----- handleRegister (stores hashed password, uses counter for ID) -----
function handleRegister(sheet, data, headers, request, output) {
  const name = request.name || '';
  const username = request.username || '';
  const password = request.password || ''; // already hashed by frontend
  const isAdmin = (request.isAdmin === 'true' || request.isAdmin === true);

  if (!username || !password || !name) {
    output.setContent(JSON.stringify({ status: "error", message: "All fields required" }));
    return output;
  }

  const usernameIdx = headers.findIndex(h => h.toLowerCase() === 'username');
  if (usernameIdx === -1) {
    output.setContent(JSON.stringify({ status: "error", message: "Username column missing" }));
    return output;
  }

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][usernameIdx]).toLowerCase() === username.toLowerCase()) {
      output.setContent(JSON.stringify({ status: "error", message: "Username already exists" }));
      return output;
    }
  }

  // ----- Generate unique ID using counter -----
  const nextNum = getNextUserId();
  const id = 'USER' + String(nextNum).padStart(4, '0');
  const adminStatus = isAdmin ? 'TRUE' : 'FALSE';
  // Store the hashed password (already hashed from frontend)
  sheet.appendRow([id, name, username, password, adminStatus, new Date().toISOString()]);

  output.setContent(JSON.stringify({
    status: "success",
    message: "User created",
    user: { id, name, username, isAdmin }
  }));
  return output;
}

// ----- handleLogin (compares hashed, migrates plain text) -----
function handleLogin(sheet, data, headers, request, output) {
  const username = request.username || '';
  const password = request.password || ''; // already MD5 hashed by frontend

  if (!username || !password) {
    output.setContent(JSON.stringify({ status: "error", message: "Username and password required" }));
    return output;
  }

  const usernameIdx = headers.findIndex(h => h.toLowerCase() === 'username');
  const passwordIdx = headers.findIndex(h => h.toLowerCase() === 'password');
  const nameIdx = headers.findIndex(h => h.toLowerCase() === 'name');
  const idIdx = headers.findIndex(h => h.toLowerCase() === 'id');
  const adminIdx = headers.findIndex(h => h.toLowerCase() === 'isadmin');

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[usernameIdx]).toLowerCase() === username.toLowerCase()) {
      const storedPassword = String(row[passwordIdx] || '');
      let isMatch = false;

      // Check if stored is MD5 hash (32 hex chars) or plain text
      if (storedPassword.length === 32 && /^[a-f0-9]{32}$/.test(storedPassword)) {
        if (storedPassword === password) isMatch = true;
      } else {
        // Plain text – migrate to hash
        if (storedPassword === password) {
          isMatch = true;
          const hashed = md5Hash(password);
          sheet.getRange(i + 1, passwordIdx + 1).setValue(hashed);
        }
      }

      if (isMatch) {
        const isAdmin = (adminIdx !== -1) ? (String(row[adminIdx] || '').toUpperCase() === 'TRUE') : false;
        output.setContent(JSON.stringify({
          status: "success",
          user: {
            id: String(row[idIdx] || ''),
            name: String(row[nameIdx] || ''),
            username: String(row[usernameIdx] || ''),
            isAdmin: isAdmin
          }
        }));
        return output;
      }
    }
  }

  output.setContent(JSON.stringify({ status: "error", message: "Invalid credentials" }));
  return output;
}

// ----- isUserAdmin (unchanged) -----
function isUserAdmin(sheet, data, headers, userId) {
  if (!userId) return false;
  const idIdx = headers.findIndex(h => h.toLowerCase() === 'id');
  const adminIdx = headers.findIndex(h => h.toLowerCase() === 'isadmin');
  if (idIdx === -1 || adminIdx === -1) return false;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idIdx]) === userId) {
      const val = String(data[i][adminIdx] || '').toUpperCase();
      return val === 'TRUE' || val === 'YES' || val === '1';
    }
  }
  return false;
}

// ----- handleGetUsers (unchanged) -----
function handleGetUsers(sheet, data, headers, request, output) {
  const adminId = request.adminId || '';
  if (!isUserAdmin(sheet, data, headers, adminId)) {
    output.setContent(JSON.stringify({ status: "error", message: "Unauthorized" }));
    return output;
  }
  const idIdx = headers.findIndex(h => h.toLowerCase() === 'id');
  const nameIdx = headers.findIndex(h => h.toLowerCase() === 'name');
  const usernameIdx = headers.findIndex(h => h.toLowerCase() === 'username');
  const adminIdx = headers.findIndex(h => h.toLowerCase() === 'isadmin');
  const createdAtIdx = headers.findIndex(h => h.toLowerCase() === 'createdat');
  const users = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const isAdmin = (adminIdx !== -1) ? (String(row[adminIdx] || '').toUpperCase() === 'TRUE') : false;
    users.push({
      id: String(row[idIdx] || ''),
      name: String(row[nameIdx] || ''),
      username: String(row[usernameIdx] || ''),
      isAdmin: isAdmin,
      createdAt: String(row[createdAtIdx] || '')
    });
  }
  output.setContent(JSON.stringify({ status: "success", users }));
  return output;
}

// ----- handleDeleteUser (unchanged – counter stays) -----
function handleDeleteUser(sheet, data, headers, request, output) {
  const adminId = request.adminId || '';
  const userId = request.userId || '';
  if (!isUserAdmin(sheet, data, headers, adminId)) {
    output.setContent(JSON.stringify({ status: "error", message: "Unauthorized" }));
    return output;
  }
  const idIdx = headers.findIndex(h => h.toLowerCase() === 'id');
  let rowToDelete = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idIdx]) === userId) {
      rowToDelete = i + 1;
      break;
    }
  }
  if (rowToDelete === -1) {
    output.setContent(JSON.stringify({ status: "error", message: "User not found" }));
    return output;
  }
  sheet.deleteRow(rowToDelete);
  output.setContent(JSON.stringify({ status: "success", message: "User deleted" }));
  return output;
}

// ----- handleUpdateUser (password is already hashed by frontend) -----
function handleUpdateUser(sheet, data, headers, request, output) {
  const adminId = request.adminId || '';
  const userId = request.userId || '';
  const name = request.name || '';
  const username = request.username || '';
  const password = request.password || ''; // already hashed by frontend
  const isAdmin = request.isAdmin;

  if (!isUserAdmin(sheet, data, headers, adminId)) {
    output.setContent(JSON.stringify({ status: "error", message: "Unauthorized" }));
    return output;
  }
  const idIdx = headers.findIndex(h => h.toLowerCase() === 'id');
  const nameIdx = headers.findIndex(h => h.toLowerCase() === 'name');
  const usernameIdx = headers.findIndex(h => h.toLowerCase() === 'username');
  const passwordIdx = headers.findIndex(h => h.toLowerCase() === 'password');
  const adminIdx = headers.findIndex(h => h.toLowerCase() === 'isadmin');

  let rowToUpdate = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idIdx]) === userId) {
      rowToUpdate = i + 1;
      break;
    }
  }
  if (rowToUpdate === -1) {
    output.setContent(JSON.stringify({ status: "error", message: "User not found" }));
    return output;
  }
  if (name && nameIdx !== -1) sheet.getRange(rowToUpdate, nameIdx + 1).setValue(name);
  if (username && usernameIdx !== -1) sheet.getRange(rowToUpdate, usernameIdx + 1).setValue(username);
  if (password && passwordIdx !== -1) sheet.getRange(rowToUpdate, passwordIdx + 1).setValue(password);
  if (isAdmin !== undefined && isAdmin !== null && adminIdx !== -1) {
    const adminValue = (isAdmin === 'true' || isAdmin === true) ? 'TRUE' : 'FALSE';
    sheet.getRange(rowToUpdate, adminIdx + 1).setValue(adminValue);
  }
  output.setContent(JSON.stringify({ status: "success", message: "User updated" }));
  return output;
}