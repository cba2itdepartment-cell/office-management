const APP_CONFIG = {
  dbSpreadsheetName: 'GitHub Office Management',
  uploadsFolderName: 'ManagementSystem_Uploads',
  defaultSheets: {
    Users: ['UserID', 'Username', 'PasswordHash', 'Role', 'Permissions', 'CreatedAt'],
    MainWorks: ['MainWorkID', 'Name', 'Category', 'Nature', 'DayMonth', 'WorkDetails', 'SOP', 'ChecklistJSON', 'AttachmentURLs', 'CreatedBy', 'CreatedAt'],
    Tasks: ['TaskID', 'MainWorkID', 'TaskType', 'AssignedTo', 'Details', 'AssignDate', 'DueDate', 'Status', 'MonitoredBy', 'CreatedAt'],
    Students: ['StudentID', 'StudentName', 'ContactInfo', 'RegistrationDate', 'Status'],
    StudentDocuments: ['DocID', 'StudentID', 'DocumentName', 'DriveUrl', 'Status', 'DueSubmissionDate', 'ApplicationPrefix', 'FollowupDate', 'Notes'],
    Followups: ['FollowupID', 'ReferenceType', 'ReferenceId', 'Description', 'TargetDate', 'Status', 'AssignedTo', 'CreatedAt']
  }
};

function doGet(e) {
  try {
    const template = HtmlService.createTemplateFromFile('Index');
    return template.evaluate()
      .setTitle('Office Management System')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, shrink-to-fit=no');
  } catch (err) {
    return HtmlService.createHtmlOutput('<html><body><h3>Application Error</h3><p>' + String(err) + '</p></body></html>').setTitle('Error');
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getDatabaseSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty('GITHUB_OFFICE_MANAGEMENT_DB_ID');

  if (savedId) {
    try {
      return SpreadsheetApp.openById(savedId);
    } catch (error) {
      props.deleteProperty('GITHUB_OFFICE_MANAGEMENT_DB_ID');
    }
  }

  try {
    const fileIterator = DriveApp.getFilesByName(APP_CONFIG.dbSpreadsheetName);
    if (fileIterator.hasNext()) {
      const file = fileIterator.next();
      const ss = SpreadsheetApp.openById(file.getId());
      props.setProperty('GITHUB_OFFICE_MANAGEMENT_DB_ID', ss.getId());
      return ss;
    }
  } catch (error) {
    Logger.log('Error searching: ' + error);
  }

  try {
    const newSpreadsheet = SpreadsheetApp.create(APP_CONFIG.dbSpreadsheetName);
    Utilities.sleep(500);
    props.setProperty('GITHUB_OFFICE_MANAGEMENT_DB_ID', newSpreadsheet.getId());
    return newSpreadsheet;
  } catch (error) {
    throw error;
  }
}

function getOrCreateSheet(sheetName, defaultHeaders) {
  const ss = getDatabaseSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    Utilities.sleep(500);
  }

  if (defaultHeaders && defaultHeaders.length > 0) {
    const lastRow = sheet.getLastRow();
    if (lastRow === 0) {
      sheet.getRange(1, 1, 1, defaultHeaders.length).setValues([defaultHeaders]);
    }
  }

  return sheet;
}

function ensureDatabase() {
  const ss = getDatabaseSpreadsheet();
  const sheetNames = Object.keys(APP_CONFIG.defaultSheets);

  sheetNames.forEach(sheetName => {
    const headers = APP_CONFIG.defaultSheets[sheetName];
    getOrCreateSheet(sheetName, headers);
    Utilities.sleep(200);
  });

  return { success: true, sheets: getAllSheets() };
}

function getAllSheets() {
  const ss = getDatabaseSpreadsheet();
  const sheets = ss.getSheets();
  return sheets.map(s => s.getName());
}

function sheetToObjects(sheetName) {
  try {
    const sheet = getDatabaseSpreadsheet().getSheetByName(sheetName);
    if (!sheet) return [];

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if (lastRow < 2 || lastCol === 0) return [];

    const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const headers = data[0].map(h => String(h || '').trim());

    return data.slice(1).filter(row => row.some(cell => String(cell || '').trim() !== '')).map(row => {
      const obj = {};
      headers.forEach((header, index) => {
        if (header) obj[header] = row[index] !== undefined ? row[index] : '';
      });
      return obj;
    });
  } catch (error) {
    return [];
  }
}

function getSheetDataWithHeaders(sheetName) {
  try {
    ensureDatabase();
    const sheet = getOrCreateSheet(sheetName, APP_CONFIG.defaultSheets[sheetName] || []);
    const lastCol = sheet.getLastColumn();
    const headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim()).filter(h => h !== '') : [];
    const rows = sheetToObjects(sheetName);
    return { success: true, headers: headers, data: rows };
  } catch (err) {
    return { success: false, error: String(err) };
  }
}

function addNewRow(sheetName, rowObject) {
  try {
    ensureDatabase();
    const sheet = getOrCreateSheet(sheetName, APP_CONFIG.defaultSheets[sheetName] || []);
    const lastCol = sheet.getLastColumn();
    const headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim()).filter(h => h !== '') : [];

    if (headers.length === 0) {
      const newHeaders = Object.keys(rowObject || {});
      sheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
      sheet.appendRow(newHeaders.map(key => rowObject[key] || ''));
      return { success: true };
    }

    const values = headers.map(header => rowObject && rowObject[header] ? rowObject[header] : '');
    sheet.appendRow(values);
    return { success: true };
  } catch (err) {
    return { success: false, message: String(err) };
  }
}

function updateRow(sheetName, idField, idValue, updateObject) {
  try {
    const sheet = getOrCreateSheet(sheetName, APP_CONFIG.defaultSheets[sheetName] || []);
    const data = sheet.getDataRange().getValues();
    const headers = data[0] ? data[0].map(h => String(h || '').trim()) : [];

    if (!headers.length) return { success: false, message: 'No headers' };

    const idIndex = headers.indexOf(String(idField || '').trim());
    if (idIndex === -1) return { success: false, message: 'ID field not found' };

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idIndex] || '').trim() === String(idValue || '').trim()) {
        headers.forEach((header, colIndex) => {
          if (updateObject[header] !== undefined) {
            sheet.getRange(i + 1, colIndex + 1).setValue(updateObject[header]);
          }
        });
        return { success: true };
      }
    }

    return { success: false, message: 'Row not found' };
  } catch (err) {
    return { success: false, message: String(err) };
  }
}

function deleteRow(sheetName, idField, idValue) {
  try {
    const sheet = getOrCreateSheet(sheetName, APP_CONFIG.defaultSheets[sheetName] || []);
    const data = sheet.getDataRange().getValues();
    const headers = data[0] ? data[0].map(h => String(h || '').trim()) : [];

    if (!headers.length) return { success: false };

    const idIndex = headers.indexOf(String(idField || '').trim());
    if (idIndex === -1) return { success: false };

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idIndex] || '').trim() === String(idValue || '').trim()) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }

    return { success: false };
  } catch (err) {
    return { success: false };
  }
}

function getRowById(sheetName, idField, idValue) {
  const rows = sheetToObjects(sheetName);
  return rows.find(row => String(row[idField] || '').trim() === String(idValue || '').trim()) || null;
}

function hashPassword(password) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(password || ''), Utilities.Charset.UTF_8);
  return digest.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

function generateId(prefix) {
  const now = new Date();
  const stamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
  const randomPart = Math.random().toString(36).slice(2, 8).toUpperCase();
  return prefix + '-' + stamp + '-' + randomPart;
}

function safeJsonParse(value, fallback) {
  if (value === null || value === undefined || value === '') return fallback === undefined ? [] : fallback;
  if (typeof value === 'object') return value;
  try {
    const parsed = JSON.parse(value);
    return parsed !== undefined ? parsed : (fallback === undefined ? [] : fallback);
  } catch (err) {
    return fallback === undefined ? [] : fallback;
  }
}

function getUserByUsername(username) {
  const users = sheetToObjects('Users');
  return users.find(u => String(u.Username || '').trim().toLowerCase() === String(username || '').trim().toLowerCase()) || null;
}

function loginUser(username, password) {
  try {
    ensureDatabase();
    const user = getUserByUsername(username);
    if (!user) return { success: false, message: 'User not found' };
    const hashedInput = hashPassword(password);
    if (String(user.PasswordHash || '').trim() !== hashedInput) return { success: false, message: 'Incorrect password' };
    return { success: true, user: { UserID: user.UserID || '', Username: user.Username, Role: user.Role || 'User', Permissions: safeJsonParse(user.Permissions, []) } };
  } catch (err) {
    return { success: false, message: 'Login failed' };
  }
}

function getUsers() {
  return sheetToObjects('Users');
}

function createUser(payload) {
  try {
    ensureDatabase();
    const raw = payload || {};
    const username = String(raw.Username || '').trim();
    const password = String(raw.Password || '');
    const role = String(raw.Role || 'User').trim();

    if (!username || !password) return { success: false, message: 'Username and password required' };

    const existingUser = getUserByUsername(username);
    if (existingUser) return { success: false, message: 'Username exists' };

    const sheet = getOrCreateSheet('Users', APP_CONFIG.defaultSheets.Users);
    const userId = generateId('USR');

    sheet.appendRow([userId, username, hashPassword(password), role, JSON.stringify([]), new Date().toISOString()]);
    return { success: true, userId: userId };
  } catch (err) {
    return { success: false, message: String(err) };
  }
}

function getMainWorks() {
  const mainWorks = sheetToObjects('MainWorks');
  const tasks = sheetToObjects('Tasks');

  return mainWorks.map(mainWork => {
    const relatedTasks = tasks.filter(task => String(task.MainWorkID || '').trim() === String(mainWork.MainWorkID || '').trim());
    return {
      ...mainWork,
      ChecklistJSON: safeJsonParse(mainWork.ChecklistJSON, []),
      AttachmentURLs: safeJsonParse(mainWork.AttachmentURLs, []),
      Tasks: relatedTasks
    };
  });
}

function saveMainWork(payload) {
  try {
    ensureDatabase();
    const data = payload || {};
    const sheet = getOrCreateSheet('MainWorks', APP_CONFIG.defaultSheets.MainWorks);

    const mainWorkId = String(data.MainWorkID || '').trim() || generateId('MW');

    const rowData = {
      MainWorkID: mainWorkId,
      Name: data.Name || '',
      Category: data.Category || 'Government',
      Nature: data.Nature || 'Daily',
      DayMonth: data.DayMonth || '',
      WorkDetails: data.WorkDetails || '',
      SOP: data.SOP || '',
      ChecklistJSON: JSON.stringify(Array.isArray(data.ChecklistJSON) ? data.ChecklistJSON : []),
      AttachmentURLs: JSON.stringify(Array.isArray(data.AttachmentURLs) ? data.AttachmentURLs : []),
      CreatedBy: data.CreatedBy || 'System',
      CreatedAt: new Date().toISOString()
    };

    const existing = getRowById('MainWorks', 'MainWorkID', mainWorkId);
    if (existing) {
      return updateRow('MainWorks', 'MainWorkID', mainWorkId, rowData);
    }

    sheet.appendRow([rowData.MainWorkID, rowData.Name, rowData.Category, rowData.Nature, rowData.DayMonth, rowData.WorkDetails, rowData.SOP, rowData.ChecklistJSON, rowData.AttachmentURLs, rowData.CreatedBy, rowData.CreatedAt]);
    return { success: true, MainWorkID: mainWorkId };
  } catch (err) {
    return { success: false, message: String(err) };
  }
}

function getTasks(filterStatus) {
  const tasks = sheetToObjects('Tasks');
  const mainWorks = sheetToObjects('MainWorks');

  return tasks.map(task => {
    const mainWork = mainWorks.find(mw => String(mw.MainWorkID || '').trim() === String(task.MainWorkID || '').trim());
    return { ...task, MainWorkName: mainWork ? mainWork.Name : 'N/A', Status: task.Status || 'Pending' };
  }).filter(task => {
    if (!filterStatus || filterStatus === 'All') return true;
    return String(task.Status || '').trim().toLowerCase() === filterStatus.toLowerCase();
  });
}

function createTask(payload) {
  try {
    ensureDatabase();
    const data = payload || {};
    const sheet = getOrCreateSheet('Tasks', APP_CONFIG.defaultSheets.Tasks);
    const taskId = String(data.TaskID || '').trim() || generateId('TASK');

    const row = [taskId, data.MainWorkID || '', data.TaskType || 'Manual', data.AssignedTo || '', data.Details || '', data.AssignDate || '', data.DueDate || '', data.Status || 'Pending', data.MonitoredBy || '', new Date().toISOString()];
    sheet.appendRow(row);
    return { success: true, TaskID: taskId };
  } catch (err) {
    return { success: false, message: String(err) };
  }
}

function updateTaskStatus(taskId, status) {
  try {
    const row = getRowById('Tasks', 'TaskID', taskId);
    if (!row) return { success: false, message: 'Task not found' };
    return updateRow('Tasks', 'TaskID', taskId, { Status: status || 'Completed' });
  } catch (err) {
    return { success: false, message: String(err) };
  }
}

function getFollowups() {
  return sheetToObjects('Followups').map(row => ({ ...row, Status: row.Status || 'Open' }));
}

function createFollowup(payload) {
  try {
    ensureDatabase();
    const data = payload || {};
    const sheet = getOrCreateSheet('Followups', APP_CONFIG.defaultSheets.Followups);
    const followupId = generateId('FUP');

    sheet.appendRow([followupId, data.ReferenceType || 'Custom', data.ReferenceId || '', data.Description || '', data.TargetDate || '', data.Status || 'Open', data.AssignedTo || '', new Date().toISOString()]);
    return { success: true, FollowupID: followupId };
  } catch (err) {
    return { success: false, message: String(err) };
  }
}

function updateFollowupStatus(followupId, status) {
  try {
    const row = getRowById('Followups', 'FollowupID', followupId);
    if (!row) return { success: false };
    return updateRow('Followups', 'FollowupID', followupId, { Status: status || 'Closed' });
  } catch (err) {
    return { success: false };
  }
}

function getStudents() {
  const students = sheetToObjects('Students');
  const documents = sheetToObjects('StudentDocuments');

  return students.map(student => ({
    ...student,
    Documents: documents.filter(doc => String(doc.StudentID || '').trim() === String(student.StudentID || '').trim())
  }));
}

function registerStudent(studentData, documents) {
  try {
    ensureDatabase();

    const studentSheet = getOrCreateSheet('Students', APP_CONFIG.defaultSheets.Students);
    const docSheet = getOrCreateSheet('StudentDocuments', APP_CONFIG.defaultSheets.StudentDocuments);
    const followupsSheet = getOrCreateSheet('Followups', APP_CONFIG.defaultSheets.Followups);

    const studentId = generateId('STU');

    studentSheet.appendRow([studentId, studentData.StudentName || '', studentData.ContactInfo || '', new Date().toISOString(), studentData.Status || 'Active']);

    const documentList = Array.isArray(documents) ? documents : [];

    documentList.forEach(doc => {
      const docId = generateId('DOC');
      const appPrefix = doc.ApplicationPrefix || ('APP-' + studentId.slice(-5).toUpperCase());
      const dueDate = doc.DueSubmissionDate || '';
      const followupDate = doc.FollowupDate || dueDate;
      const status = doc.Status || 'Pending';

      docSheet.appendRow([docId, studentId, doc.DocumentName || 'Document', doc.DriveUrl || '', status, dueDate, appPrefix, followupDate, doc.Notes || '']);

      if (String(status || '').trim().toLowerCase() !== 'uploaded') {
        followupsSheet.appendRow([generateId('FUP'), 'StudentDoc', docId, 'Follow up for ' + (doc.DocumentName || 'document'), dueDate || followupDate || '', 'Open', doc.AssignedTo || '', new Date().toISOString()]);
      }
    });

    return { success: true, StudentID: studentId };
  } catch (err) {
    return { success: false, message: String(err) };
  }
}

function getDashboardData() {
  try {
    ensureDatabase();
    const allSheets = getAllSheets();
    const summary = {};

    allSheets.forEach(sheetName => {
      summary[sheetName] = sheetToObjects(sheetName).length;
    });

    const tasks = sheetToObjects('Tasks');
    const followups = sheetToObjects('Followups');
    const students = sheetToObjects('Students');
    const mainWorks = sheetToObjects('MainWorks');

    summary.totalMainWorks = mainWorks.length;
    summary.pendingTasks = tasks.filter(t => {
      const status = String(t.Status || '').trim().toLowerCase();
      return status === 'pending' || status === 'in progress' || status === 'in-progress';
    }).length;
    summary.openFollowups = followups.filter(f => String(f.Status || '').trim().toLowerCase() === 'open').length;
    summary.activeStudents = students.filter(s => {
      const status = String(s.Status || '').trim().toLowerCase();
      return status === 'active' || status === 'new' || status === 'registered' || status === '';
    }).length;

    return { success: true, summary: summary, sheets: allSheets };
  } catch (err) {
    return { success: false, message: String(err) };
  }
}

function setupDatabase() {
  return ensureDatabase();
}
