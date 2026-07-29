// Google Apps Script Backend for House Tracker
// Deploy as Web App with Execute as: Me, Who has access: Anyone

const SHEET_NAME_ENTRIES = 'Entries';
const SHEET_NAME_STUDENTS = 'Students';
const SHEET_NAME_CONFIG = 'Config';
const SHEET_NAME_USERS = 'Users';

// Default admin credentials (change these!)
const DEFAULT_ADMIN = {
  username: 'admin',
  password: 'admin123', // Change this!
  fullName: 'Administrator',
  role: 'admin'
};

function doGet(e) {
  const action = e.parameter.action;
  
  try {
    switch(action) {
      case 'getEntries':
        return jsonResponse(getEntries());
      case 'getStudents':
        return jsonResponse(getStudents());
      case 'getConfig':
        return jsonResponse(getConfig());
      case 'getHouseEntries':
        return jsonResponse(getHouseEntries(e.parameter.house));
      case 'getStudentData':
        return jsonResponse(getStudentData(e.parameter.house));
      case 'getUsers':
        return jsonResponse(getUsers());
      default:
        return jsonResponse({error: 'Invalid action'}, 400);
    }
  } catch(err) {
    return jsonResponse({error: err.message}, 500);
  }
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const action = data.action;
  
  try {
    switch(action) {
      // Authentication
      case 'login':
        return jsonResponse(login(data.username, data.password));
      case 'register':
        return jsonResponse(register(data.user));
      case 'updateUser':
        return jsonResponse(updateUser(data.user));
      case 'deleteUser':
        return jsonResponse(deleteUser(data.username));
      case 'changePassword':
        return jsonResponse(changePassword(data.username, data.oldPassword, data.newPassword));
      case 'resetPassword':
        return jsonResponse(resetPassword(data.username, data.newPassword));
      
      // Data operations (require admin for write operations)
      case 'addEntry':
        return jsonResponse(addEntry(data.entry));
      case 'addEntries':
        return jsonResponse(addEntries(data.entries));
      case 'addStudents':
        return jsonResponse(addStudents(data.students));
      case 'deleteEntry':
        return jsonResponse(deleteEntry(data.id));
      case 'deleteEntries':
        return jsonResponse(deleteEntries(data.ids));
      case 'deleteStudentData':
        return jsonResponse(deleteStudentData(data.house, data.cls));
      case 'resetAll':
        return jsonResponse(resetAll());
      case 'updateConfig':
        return jsonResponse(updateConfig(data.config));
      case 'importExcelData':
        return jsonResponse(importExcelData(data));
      default:
        return jsonResponse({error: 'Invalid action'}, 400);
    }
  } catch(err) {
    return jsonResponse({error: err.message}, 500);
  }
}

function jsonResponse(data, statusCode) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// Handle CORS - return response with proper headers
function doOptions(e) {
  return ContentService.createTextOutput('')
    .setMimeType(ContentService.MimeType.TEXT);
}

// ===== SHEET HELPERS =====

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if(!sheet) {
    sheet = ss.insertSheet(name);
    initializeSheet(sheet, name);
  }
  return sheet;
}

function initializeSheet(sheet, name) {
  if(name === SHEET_NAME_ENTRIES) {
    sheet.appendRow(['ID', 'House', 'Category', 'SubCategory', 'Position', 'Points', 'Event', 'Note', 'Date', 'Exam']);
    sheet.setFrozenRows(1);
    sheet.getRange('A1:J1').setFontWeight('bold').setBackground('#334155').setFontColor('#f1f5f9');
  } else if(name === SHEET_NAME_STUDENTS) {
    sheet.appendRow(['Name', 'House', 'Class', 'Section', 'Subjects', 'Total', 'Exam', 'Date']);
    sheet.setFrozenRows(1);
    sheet.getRange('A1:H1').setFontWeight('bold').setBackground('#334155').setFontColor('#f1f5f9');
  } else if(name === SHEET_NAME_CONFIG) {
    sheet.appendRow(['Key', 'Value']);
    sheet.appendRow(['columnMapping', JSON.stringify({name:-1,house:-1,marks:-1,class:-1,section:-1,sNo:-1,admNo:-1,subjects:[],total:-1})]);
    sheet.appendRow(['formState', '{}']);
    sheet.setFrozenRows(1);
    sheet.getRange('A1:B1').setFontWeight('bold').setBackground('#334155').setFontColor('#f1f5f9');
  } else if(name === SHEET_NAME_USERS) {
    sheet.appendRow(['Username', 'Password', 'FullName', 'Role', 'CreatedAt']);
    sheet.setFrozenRows(1);
    sheet.getRange('A1:E1').setFontWeight('bold').setBackground('#334155').setFontColor('#f1f5f9');
    
    // Add default admin user
    const hashedPassword = hashPassword(DEFAULT_ADMIN.password);
    sheet.appendRow([
      DEFAULT_ADMIN.username,
      hashedPassword,
      DEFAULT_ADMIN.fullName,
      DEFAULT_ADMIN.role,
      new Date().toISOString()
    ]);
    
    Logger.log('Created default admin user: ' + DEFAULT_ADMIN.username);
  }
}

// ===== ENTRIES =====

function getEntries() {
  const sheet = getSheet(SHEET_NAME_ENTRIES);
  const data = sheet.getDataRange().getValues();
  if(data.length <= 1) return [];
  
  const headers = data[0];
  return data.slice(1).map(row => ({
    id: row[0],
    house: row[1],
    category: row[2],
    subCategory: row[3],
    position: row[4],
    points: row[5],
    event: row[6],
    note: row[7],
    date: row[8],
    exam: row[9]
  }));
}

function getHouseEntries(house) {
  return getEntries().filter(e => e.house === house);
}

function addEntry(entry) {
  const sheet = getSheet(SHEET_NAME_ENTRIES);
  const id = entry.id || Date.now().toString(36) + Math.random().toString(36).substr(2,5);
  const date = entry.date || new Date().toISOString();
  
  sheet.appendRow([
    id,
    entry.house,
    entry.category,
    entry.subCategory || '',
    entry.position || 0,
    entry.points,
    entry.event || '',
    entry.note || '',
    date,
    entry.exam || ''
  ]);
  
  return {success: true, id: id};
}

function addEntries(entries) {
  const sheet = getSheet(SHEET_NAME_ENTRIES);
  const ids = [];
  
  entries.forEach(entry => {
    const id = entry.id || Date.now().toString(36) + Math.random().toString(36).substr(2,5);
    const date = entry.date || new Date().toISOString();
    
    sheet.appendRow([
      id,
      entry.house,
      entry.category,
      entry.subCategory || '',
      entry.position || 0,
      entry.points,
      entry.event || '',
      entry.note || '',
      date,
      entry.exam || ''
    ]);
    
    ids.push(id);
  });
  
  return {success: true, ids: ids, count: ids.length};
}

function deleteEntry(id) {
  const sheet = getSheet(SHEET_NAME_ENTRIES);
  const data = sheet.getDataRange().getValues();
  
  for(let i = 1; i < data.length; i++) {
    if(data[i][0] === id) {
      sheet.deleteRow(i + 1);
      return {success: true};
    }
  }
  
  return {success: false, error: 'Entry not found'};
}

function deleteEntries(ids) {
  const sheet = getSheet(SHEET_NAME_ENTRIES);
  const data = sheet.getDataRange().getValues();
  let deleted = 0;
  
  // Sort indices in reverse to avoid shifting issues
  const rowsToDelete = [];
  for(let i = 1; i < data.length; i++) {
    if(ids.includes(data[i][0])) {
      rowsToDelete.push(i + 1);
    }
  }
  
  // Delete from bottom to top
  rowsToDelete.reverse().forEach(row => {
    sheet.deleteRow(row);
    deleted++;
  });
  
  return {success: true, deleted: deleted};
}

// ===== STUDENTS =====

function getStudents() {
  const sheet = getSheet(SHEET_NAME_STUDENTS);
  const data = sheet.getDataRange().getValues();
  if(data.length <= 1) return [];
  
  return data.slice(1).map(row => ({
    name: row[0],
    house: row[1],
    cls: row[2],
    sec: row[3],
    subjects: JSON.parse(row[4] || '{}'),
    total: row[5],
    exam: row[6],
    date: row[7]
  }));
}

function getStudentData(house) {
  return getStudents().filter(s => s.house === house);
}

function addStudents(students) {
  const sheet = getSheet(SHEET_NAME_STUDENTS);
  let count = 0;
  
  students.forEach(student => {
    sheet.appendRow([
      student.name,
      student.house,
      student.cls || '',
      student.sec || '',
      JSON.stringify(student.subjects || {}),
      student.total || 0,
      student.exam || '',
      student.date || new Date().toISOString()
    ]);
    count++;
  });
  
  return {success: true, count: count};
}

function deleteStudentData(house, cls) {
  const sheet = getSheet(SHEET_NAME_STUDENTS);
  const data = sheet.getDataRange().getValues();
  let deleted = 0;
  
  const rowsToDelete = [];
  for(let i = 1; i < data.length; i++) {
    if(data[i][1] === house && (cls === undefined || data[i][2] === cls)) {
      rowsToDelete.push(i + 1);
    }
  }
  
  rowsToDelete.reverse().forEach(row => {
    sheet.deleteRow(row);
    deleted++;
  });
  
  return {success: true, deleted: deleted};
}

// ===== CONFIG =====

function getConfig() {
  const sheet = getSheet(SHEET_NAME_CONFIG);
  const data = sheet.getDataRange().getValues();
  const config = {};
  
  for(let i = 1; i < data.length; i++) {
    config[data[i][0]] = data[i][1];
  }
  
  return config;
}

function updateConfig(config) {
  const sheet = getSheet(SHEET_NAME_CONFIG);
  const data = sheet.getDataRange().getValues();
  
  Object.entries(config).forEach(([key, value]) => {
    let found = false;
    for(let i = 1; i < data.length; i++) {
      if(data[i][0] === key) {
        sheet.getRange(i + 1, 2).setValue(typeof value === 'object' ? JSON.stringify(value) : value);
        found = true;
        break;
      }
    }
    if(!found) {
      sheet.appendRow([key, typeof value === 'object' ? JSON.stringify(value) : value]);
    }
  });
  
  return {success: true};
}

// ===== RESET =====

function resetAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Delete and recreate sheets (keep Users)
  [SHEET_NAME_ENTRIES, SHEET_NAME_STUDENTS, SHEET_NAME_CONFIG].forEach(name => {
    const sheet = ss.getSheetByName(name);
    if(sheet) ss.deleteSheet(sheet);
  });
  
  // Recreate
  getSheet(SHEET_NAME_ENTRIES);
  getSheet(SHEET_NAME_STUDENTS);
  getSheet(SHEET_NAME_CONFIG);
  
  return {success: true, message: 'All data has been reset'};
}

// Reset everything including users
function resetAllIncludingUsers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Delete all sheets
  [SHEET_NAME_ENTRIES, SHEET_NAME_STUDENTS, SHEET_NAME_CONFIG, SHEET_NAME_USERS].forEach(name => {
    const sheet = ss.getSheetByName(name);
    if(sheet) ss.deleteSheet(sheet);
  });
  
  // Recreate all
  getSheet(SHEET_NAME_ENTRIES);
  getSheet(SHEET_NAME_STUDENTS);
  getSheet(SHEET_NAME_CONFIG);
  getSheet(SHEET_NAME_USERS);
  
  return {success: true, message: 'All data including users has been reset'};
}

// ===== EXCEL IMPORT =====

function importExcelData(data) {
  const {entries, students, examName, category} = data;
  
  // Add entries
  const entryResult = addEntries(entries);
  
  // Add students
  const studentResult = addStudents(students);
  
  return {
    success: true,
    entriesAdded: entryResult.count,
    studentsAdded: studentResult.count
  };
}

// ===== AUTHENTICATION =====

// Simple password hashing (for production, use a more secure method)
function hashPassword(password) {
  let hash = 0;
  for (let i = 0; i < password.length; i++) {
    const char = password.charCodeAt(i);
    hash = ((hash << 5) - hash) + char;
    hash = hash & hash; // Convert to 32bit integer
  }
  return 'h_' + Math.abs(hash).toString(36);
}

// Verify password
function verifyPassword(password, hashedPassword) {
  return hashPassword(password) === hashedPassword;
}

// Login function
function login(username, password) {
  const sheet = getSheet(SHEET_NAME_USERS);
  const data = sheet.getDataRange().getValues();
  
  for(let i = 1; i < data.length; i++) {
    if(data[i][0].toLowerCase() === username.toLowerCase()) {
      if(verifyPassword(password, data[i][1])) {
        return {
          success: true,
          user: {
            username: data[i][0],
            fullName: data[i][2],
            role: data[i][3]
          }
        };
      } else {
        return {success: false, error: 'Invalid password'};
      }
    }
  }
  
  return {success: false, error: 'User not found'};
}

// Register new user (admin only)
function register(userData) {
  const {username, password, fullName, role} = userData;
  
  if(!username || !password || !fullName) {
    return {success: false, error: 'Username, password, and full name are required'};
  }
  
  if(password.length < 4) {
    return {success: false, error: 'Password must be at least 4 characters'};
  }
  
  const sheet = getSheet(SHEET_NAME_USERS);
  const data = sheet.getDataRange().getValues();
  
  // Check if username already exists
  for(let i = 1; i < data.length; i++) {
    if(data[i][0].toLowerCase() === username.toLowerCase()) {
      return {success: false, error: 'Username already exists'};
    }
  }
  
  // Add new user
  const hashedPassword = hashPassword(password);
  sheet.appendRow([
    username,
    hashedPassword,
    fullName,
    role || 'user', // Default to 'user' role
    new Date().toISOString()
  ]);
  
  return {
    success: true,
    message: 'User registered successfully',
    user: {username, fullName, role: role || 'user'}
  };
}

// Get all users (admin only)
function getUsers() {
  const sheet = getSheet(SHEET_NAME_USERS);
  const data = sheet.getDataRange().getValues();
  
  return data.slice(1).map(row => ({
    username: row[0],
    fullName: row[2],
    role: row[3],
    createdAt: row[4]
  }));
}

// Update user (admin only)
function updateUser(userData) {
  const {username, fullName, role} = userData;
  
  const sheet = getSheet(SHEET_NAME_USERS);
  const data = sheet.getDataRange().getValues();
  
  for(let i = 1; i < data.length; i++) {
    if(data[i][0].toLowerCase() === username.toLowerCase()) {
      if(fullName) sheet.getRange(i + 1, 3).setValue(fullName);
      if(role) sheet.getRange(i + 1, 4).setValue(role);
      return {success: true, message: 'User updated successfully'};
    }
  }
  
  return {success: false, error: 'User not found'};
}

// Delete user (admin only)
function deleteUser(username) {
  if(username.toLowerCase() === 'admin') {
    return {success: false, error: 'Cannot delete admin user'};
  }
  
  const sheet = getSheet(SHEET_NAME_USERS);
  const data = sheet.getDataRange().getValues();
  
  for(let i = 1; i < data.length; i++) {
    if(data[i][0].toLowerCase() === username.toLowerCase()) {
      sheet.deleteRow(i + 1);
      return {success: true, message: 'User deleted successfully'};
    }
  }
  
  return {success: false, error: 'User not found'};
}

// Change password
function changePassword(username, oldPassword, newPassword) {
  if(!oldPassword || !newPassword) {
    return {success: false, error: 'Old and new passwords are required'};
  }
  
  if(newPassword.length < 4) {
    return {success: false, error: 'New password must be at least 4 characters'};
  }
  
  const sheet = getSheet(SHEET_NAME_USERS);
  const data = sheet.getDataRange().getValues();
  
  for(let i = 1; i < data.length; i++) {
    if(data[i][0].toLowerCase() === username.toLowerCase()) {
      if(verifyPassword(oldPassword, data[i][1])) {
        const hashedPassword = hashPassword(newPassword);
        sheet.getRange(i + 1, 2).setValue(hashedPassword);
        return {success: true, message: 'Password changed successfully'};
      } else {
        return {success: false, error: 'Invalid old password'};
      }
    }
  }
  
  return {success: false, error: 'User not found'};
}

// Reset password (admin only - no old password needed)
function resetPassword(username, newPassword) {
  if(!newPassword) {
    return {success: false, error: 'New password is required'};
  }
  
  if(newPassword.length < 4) {
    return {success: false, error: 'Password must be at least 4 characters'};
  }
  
  const sheet = getSheet(SHEET_NAME_USERS);
  const data = sheet.getDataRange().getValues();
  
  for(let i = 1; i < data.length; i++) {
    if(data[i][0].toLowerCase() === username.toLowerCase()) {
      const hashedPassword = hashPassword(newPassword);
      sheet.getRange(i + 1, 2).setValue(hashedPassword);
      return {success: true, message: 'Password reset successfully'};
    }
  }
  
  return {success: false, error: 'User not found'};
}

// ===== TEST FUNCTION =====

function testBackend() {
  // Test adding an entry
  const testEntry = {
    house: 'Compassion',
    category: 'Sports',
    subCategory: 'Athletics',
    position: 1,
    points: 30,
    event: 'Test Event',
    note: 'Testing backend',
    date: new Date().toISOString()
  };
  
  const result = addEntry(testEntry);
  Logger.log('Add entry result: ' + JSON.stringify(result));
  
  // Test getting entries
  const entries = getEntries();
  Logger.log('Total entries: ' + entries.length);
  
  return 'Test completed successfully';
}
