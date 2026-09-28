// ==================== APP CONFIGURATION ====================
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

// ==================== MAIN ENTRY POINTS ====================

function doGet(e) {
  try {
    const template = HtmlService.createTemplateFromFile('Index');
    return template.evaluate()
      .setTitle('Office Management System')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, shrink-to-fit=no');
  } catch (err) {
    return HtmlService.createHtmlOutput(
      '<html><body><h3>Application Error</h3><p>' + String(err) + '</p></body></html>'
    ).setTitle('Office Management Error');
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ==================== DATABASE CONNECTION ====================

function getScriptProperties() {
  return PropertiesService.getScriptProperties();
}

function getDatabaseSpreadsheet() {
  const props = getScriptProperties();
  const savedId = props.getProperty('GITHUB_OFFICE_MANAGEMENT_DB_ID');

  // Try to open by stored ID
  if (savedId) {
    try {
      const ss = SpreadsheetApp.openById(savedId);
      return ss;
    } catch (error) {
      props.deleteProperty('GITHUB_OFFICE_MANAGEMENT_DB_ID');
    }
  }

  // Try to find by name
  const fileIterator = DriveApp.getFilesByName(APP_CONFIG.dbSpreadsheetName);
  if (fileIterator.hasNext()) {
    const file = fileIterator.next();
    const ss = SpreadsheetApp.openById(file.getId());
    props.setProperty('GITHUB_OFFICE_MANAGEMENT_DB_ID', ss.getId());
    return ss;
  }

  // Try active spreadsheet
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active && active.getName() === APP_CONFIG.dbSpreadsheetName) {
      props.setProperty('GITHUB_OFFICE_MANAGEMENT_DB_ID', active.getId());
      return active;
    }
  } catch (error) {
    // continue
  }

  // Create new spreadsheet
  const newSpreadsheet = SpreadsheetApp.create(APP_CONFIG.dbSpreadsheetName);
  props.setProperty('GITHUB_OFFICE_MANAGEMENT_DB_ID', newSpreadsheet.getId());

  return newSpreadsheet;
}

function getUploadsFolder() {
  const folderIterator = DriveApp.getFoldersByName(APP_CONFIG.uploadsFolderName);
  if (folderIterator.hasNext()) {
    return folderIterator.next();
  }
  return DriveApp.createFolder(APP_CONFIG.uploadsFolderName);
}

// ==================== SHEET MANAGEMENT ====================

function getOrCreateSheet(sheetName, defaultHeaders) {
  const ss = getDatabaseSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  if (defaultHeaders && defaultHeaders.length > 0) {
    ensureSheetHeaders(sheet, defaultHeaders);
  }

  return sheet;
}

function ensureSheetHeaders(sheet, defaultHeaders) {
  if (!sheet) return;

  const currentHeaders = getSheetHeaders(sheet);
  const missing = defaultHeaders.filter((h) => !currentHeaders.includes(h));

  if (missing.length > 0) {
    const nextCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, nextCol, 1, missing.length).setValues([missing]);
  }
}

function getSheetHeaders(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return [];

  const row = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  return row.map((cell) => String(cell || '').trim()).filter((h) => h !== '');
}

function getAllSheets() {
  const ss = getDatabaseSpreadsheet();
  return ss.getSheets().map(function(sheet) {
    return sheet.getName();
  });
}

function ensureDatabase() {
  const ss = getDatabaseSpreadsheet();

  Object.keys(APP_CONFIG.defaultSheets).forEach(function(sheetName) {
    const headers = APP_CONFIG.defaultSheets[sheetName];
    getOrCreateSheet(sheetName, headers);
  });

  return {
    success: true,
    message: 'Database ensured',
    sheets: getAllSheets()
  };
}

// ==================== DATA OPERATIONS ====================

function sheetToObjects(sheetName) {
  const sheet = getDatabaseSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 2 || lastCol === 0) return [];

  const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = data[0].map((header) => String(header || '').trim());

  return data.slice(1)
    .filter((row) => row.some((cell) => String(cell || '').trim() !== ''))
    .map((row) => {
      const obj = {};
      headers.forEach((header, index) => {
        if (header) {
          obj[header] = row[index] !== undefined ? row[index] : '';
        }
      });
      return obj;
    });
}

function getSheetDataWithHeaders(sheetName) {
  ensureDatabase();
  const sheet = getOrCreateSheet(sheetName, APP_CONFIG.defaultSheets[sheetName] || []);
  const headers = getSheetHeaders(sheet);
  const rows = sheetToObjects(sheetName);

  return {
    success: true,
    headers: headers,
    data: rows
  };
}

function addNewRow(sheetName, rowObject) {
  try {
    ensureDatabase();
    const sheet = getOrCreateSheet(sheetName, APP_CONFIG.defaultSheets[sheetName] || []);
    const headers = getSheetHeaders(sheet);

    if (headers.length === 0) {
      const newHeaders = Object.keys(rowObject || {});
      sheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
      sheet.appendRow(newHeaders.map((key) => rowObject[key] || ''));
      return { success: true, message: 'Row added with new headers' };
    }

    const values = headers.map((header) => 
      rowObject && Object.prototype.hasOwnProperty.call(rowObject, header) ? rowObject[header] : ''
    );
    sheet.appendRow(values);

    return { success: true, message: 'Row added' };
  } catch (err) {
    return { success: false, message: String(err) };
  }
}

function updateRow(sheetName, idField, idValue, updateObject) {
  try {
    const sheet = getOrCreateSheet(sheetName, APP_CONFIG.defaultSheets[sheetName] || []);
    const data = sheet.getDataRange().getValues();
    const headers = data[0] ? data[0].map((h) => String(h || '').trim()) : [];

    if (!headers.length) {
      return { success: false, message: 'No headers found' };
    }

    const idIndex = headers.indexOf(String(idField || '').trim());
    if (idIndex === -1) {
      return { success: false, message: 'ID field not found: ' + idField };
    }

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idIndex] || '').trim() === String(idValue || '').trim()) {
        headers.forEach((header, colIndex) => {
          if (Object.prototype.hasOwnProperty.call(updateObject, header)) {
            sheet.getRange(i + 1, colIndex + 1).setValue(updateObject[header]);
          }
        });
        return { success: true, message: 'Row updated' };
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

    if (!data.length) {
      return { success: false, message: 'No data found' };
    }

    const headers = data[0].map((h) => String(h || '').trim());
    const idIndex = headers.indexOf(String(idField || '').trim());

    if (idIndex === -1) {
      return { success: false, message: 'ID field not found' };
    }

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idIndex] || '').trim() === String(idValue || '').trim()) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Row deleted' };
      }
    }

    return { success: false, message: 'Row not found' };
  } catch (err) {
    return { success: false, message: String(err) };
  }
}

function getRowById(sheetName, idField, idValue) {
  const rows = sheetToObjects(sheetName);
  return rows.find((row) => String(row[idField] || '').trim() === String(idValue || '').trim()) || null;
}

// ==================== AUTHENTICATION ====================

function hashPassword(password) {
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(password || ''),
    Utilities.Charset.UTF_8
  );
  return digest.map((b) => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

function generateId(prefix) {
  const now = new Date();
  const stamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
  const randomPart = Math.random().toString(36).slice(2, 8).toUpperCase();
  return prefix + '-' + stamp + '-' + randomPart;
}

function safeJsonParse(value, fallback) {
  if (value === null || value === undefined || value === '') {
    return fallback === undefined ? [] : fallback;
  }
  if (typeof value === 'object') return value;

  try {
    const parsed = JSON.parse(value);
    return parsed !== undefined ? parsed : (fallback === undefined ? [] : fallback);
  } catch (err) {
    return fallback === undefined ? [] : fallback;
  }
}

function normalizePermissions(value) {
  const parsed = safeJsonParse(value, []);
  if (Array.isArray(parsed)) return parsed;
  if (typeof parsed === 'string') {
    try {
      const arr = JSON.parse(parsed);
      return Array.isArray(arr) ? arr : [];
    } catch (err) {
      return [];
    }
  }
  return [];
}

function getUserByUsername(username) {
  const users = sheetToObjects('Users');
  return users.find((u) => 
    String(u.Username || '').trim().toLowerCase() === String(username || '').trim().toLowerCase()
  ) || null;
}

function loginUser(username, password) {
  try {
    ensureDatabase();
    const user = getUserByUsername(username);

    if (!user) {
      return { success: false, message: 'User not found' };
    }

    const hashedInput = hashPassword(password);
    if (String(user.PasswordHash || '').trim() !== hashedInput) {
      return { success: false, message: 'Incorrect password' };
    }

    return {
      success: true,
      user: {
        UserID: user.UserID || user.Id || '',
        Username: user.Username,
        Role: user.Role || 'User',
        Permissions: normalizePermissions(user.Permissions)
      }
    };
  } catch (err) {
    return { success: false, message: 'Login failed: ' + String(err) };
  }
}

// ==================== USER MANAGEMENT ====================

function getUsers() {
  return sheetToObjects('Users').map((u) => ({
    ...u,
    Permissions: normalizePermissions(u.Permissions)
  }));
}

function createUser(payload) {
  try {
    ensureDatabase();
    const raw = payload || {};
    const username = String(raw.Username || '').trim();
    const password = String(raw.Password || '');
    const role = String(raw.Role || 'User').trim();
    const permissions = Array.isArray(raw.Permissions) ? raw.Permissions : [];

    if (!username) {
      return { success: false, message: 'Username is required' };
    }

    const existingUser = getUserByUsername(username);
    if (existingUser) {
      return { success: false, message: 'Username already exists' };
    }

    const sheet = getOrCreateSheet('Users', APP_CONFIG.defaultSheets.Users);
    const userId = generateId('USR');

    sheet.appendRow([
      userId,
      username,
      hashPassword(password),
      role,
      JSON.stringify(permissions),
      new Date().toISOString()
    ]);

    return { success: true, userId: userId };
  } catch (err) {
    return { success: false, message: String(err) };
  }
}

// ==================== MAIN WORKS ====================

function getMainWorks() {
  const mainWorks = sheetToObjects('MainWorks');
  const tasks = sheetToObjects('Tasks');

  return mainWorks.map((mainWork) => {
    const relatedTasks = tasks.filter((task) => 
      String(task.MainWorkID || '').trim() === String(mainWork.MainWorkID || '').trim()
    );

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
      ChecklistJSON: JSON.stringify(safeJsonParse(data.ChecklistJSON, [])),
      AttachmentURLs: JSON.stringify(safeJsonParse(data.AttachmentURLs, [])),
      CreatedBy: data.CreatedBy || 'System',
      CreatedAt: new Date().toISOString()
    };

    const existing = getRowById('MainWorks', 'MainWorkID', mainWorkId);
    if (existing) {
      return updateRow('MainWorks', 'MainWorkID', mainWorkId, rowData);
    }

    sheet.appendRow([
      rowData.MainWorkID,
      rowData.Name,
      rowData.Category,
      rowData.Nature,
      rowData.DayMonth,
      rowData.WorkDetails,
      rowData.SOP,
      rowData.ChecklistJSON,
      rowData.AttachmentURLs,
      rowData.CreatedBy,
      rowData.CreatedAt
    ]);

    return { success: true, MainWorkID: mainWorkId };
  } catch (err) {
    return { success: false, message: String(err) };
  }
}

// ==================== TASKS ====================

function getTasks(filterStatus) {
  const tasks = sheetToObjects('Tasks');
  const mainWorks = sheetToObjects('MainWorks');

  const normalizedFilter = String(filterStatus || 'All').trim();

  return tasks
    .map((task) => {
      const mainWork = mainWorks.find((mw) => 
        String(mw.MainWorkID || '').trim() === String(task.MainWorkID || '').trim()
      );
      return {
        ...task,
        MainWorkName: mainWork ? mainWork.Name : 'N/A',
        Status: task.Status || 'Pending'
      };
    })
    .filter((task) => {
      if (!normalizedFilter || normalizedFilter === 'All') return true;
      return String(task.Status || '').trim().toLowerCase() === normalizedFilter.toLowerCase();
    })
    .sort((a, b) => {
      const aTime = new Date(a.DueDate || 0).getTime();
      const bTime = new Date(b.DueDate || 0).getTime();
      return aTime - bTime;
    });
}

function createTask(payload) {
  try {
    ensureDatabase();
    const data = payload || {};
    const sheet = getOrCreateSheet('Tasks', APP_CONFIG.defaultSheets.Tasks);
    const taskId = String(data.TaskID || '').trim() || generateId('TASK');

    const row = [
      taskId,
      data.MainWorkID || '',
      data.TaskType || 'Manual',
      data.AssignedTo || '',
      data.Details || '',
      data.AssignDate || '',
      data.DueDate || '',
      data.Status || 'Pending',
      data.MonitoredBy || '',
      new Date().toISOString()
    ];

    const existing = getRowById('Tasks', 'TaskID', taskId);
    if (existing) {
      return updateRow('Tasks', 'TaskID', taskId, {
        TaskID: taskId,
        MainWorkID: data.MainWorkID || '',
        TaskType: data.TaskType || 'Manual',
        AssignedTo: data.AssignedTo || '',
        Details: data.Details || '',
        AssignDate: data.AssignDate || '',
        DueDate: data.DueDate || '',
        Status: data.Status || 'Pending',
        MonitoredBy: data.MonitoredBy || '',
        CreatedAt: new Date().toISOString()
      });
    }

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

// ==================== FOLLOW-UPS ====================

function getFollowups() {
  return sheetToObjects('Followups').map((row) => ({
    ...row,
    Status: row.Status || 'Open'
  }));
}

function createFollowup(payload) {
  try {
    ensureDatabase();
    const data = payload || {};
    const sheet = getOrCreateSheet('Followups', APP_CONFIG.defaultSheets.Followups);
    const followupId = generateId('FUP');

    sheet.appendRow([
      followupId,
      data.ReferenceType || 'Custom',
      data.ReferenceId || '',
      data.Description || '',
      data.TargetDate || '',
      data.Status || 'Open',
      data.AssignedTo || '',
      new Date().toISOString()
    ]);

    return { success: true, FollowupID: followupId };
  } catch (err) {
    return { success: false, message: String(err) };
  }
}

function updateFollowupStatus(followupId, status) {
  try {
    const row = getRowById('Followups', 'FollowupID', followupId);
    if (!row) return { success: false, message: 'Follow-up not found' };

    return updateRow('Followups', 'FollowupID', followupId, { Status: status || 'Closed' });
  } catch (err) {
    return { success: false, message: String(err) };
  }
}

// ==================== STUDENTS ====================

function getStudents() {
  const students = sheetToObjects('Students');
  const documents = sheetToObjects('StudentDocuments');

  return students.map((student) => ({
    ...student,
    Documents: documents.filter((doc) => 
      String(doc.StudentID || '').trim() === String(student.StudentID || '').trim()
    )
  }));
}

function registerStudent(studentData, documents) {
  try {
    ensureDatabase();

    const studentSheet = getOrCreateSheet('Students', APP_CONFIG.defaultSheets.Students);
    const docSheet = getOrCreateSheet('StudentDocuments', APP_CONFIG.defaultSheets.StudentDocuments);
    const followupsSheet = getOrCreateSheet('Followups', APP_CONFIG.defaultSheets.Followups);

    const studentId = generateId('STU');

    studentSheet.appendRow([
      studentId,
      studentData.StudentName || '',
      studentData.ContactInfo || '',
      new Date().toISOString(),
      studentData.Status || 'Active'
    ]);

    const documentList = Array.isArray(documents) ? documents : [];

    documentList.forEach((doc) => {
      const docId = generateId('DOC');
      const appPrefix = doc.ApplicationPrefix || ('APP-' + studentId.slice(-5).toUpperCase());
      const dueDate = doc.DueSubmissionDate || '';
      const followupDate = doc.FollowupDate || dueDate;
      const status = doc.Status || 'Pending';

      docSheet.appendRow([
        docId,
        studentId,
        doc.DocumentName || 'Document',
        doc.DriveUrl || '',
        status,
        dueDate,
        appPrefix,
        followupDate,
        doc.Notes || ''
      ]);

      if (String(status || '').trim().toLowerCase() !== 'uploaded') {
        followupsSheet.appendRow([
          generateId('FUP'),
          'StudentDoc',
          docId,
          'Follow up for ' + (doc.DocumentName || 'document'),
          dueDate || followupDate || '',
          'Open',
          doc.AssignedTo || '',
          new Date().toISOString()
        ]);
      }
    });

    return { success: true, StudentID: studentId };
  } catch (err) {
    return { success: false, message: String(err) };
  }
}

// ==================== DASHBOARD ====================

function getDashboardData() {
  try {
    ensureDatabase();
    const allSheets = getAllSheets();
    const summary = {};

    allSheets.forEach((sheetName) => {
      summary[sheetName] = sheetToObjects(sheetName).length;
    });

    const tasks = sheetToObjects('Tasks');
    const followups = sheetToObjects('Followups');
    const students = sheetToObjects('Students');
    const mainWorks = sheetToObjects('MainWorks');

    summary.totalMainWorks = mainWorks.length;
    summary.pendingTasks = tasks.filter((t) => {
      const status = String(t.Status || '').trim().toLowerCase();
      return status === 'pending' || status === 'in progress' || status === 'in-progress';
    }).length;
    summary.openFollowups = followups.filter((f) => 
      String(f.Status || '').trim().toLowerCase() === 'open'
    ).length;
    summary.activeStudents = students.filter((s) => {
      const status = String(s.Status || '').trim().toLowerCase();
      return status === 'active' || status === 'new' || status === 'registered' || status === '';
    }).length;

    return {
      success: true,
      summary: summary,
      sheets: allSheets
    };
  } catch (err) {
    return { success: false, message: String(err) };
  }
}

// ==================== FILE UPLOAD ====================

function getMimeTypeFromFilename(fileName) {
  const name = String(fileName || '').toLowerCase();

  if (name.endsWith('.png')) return 'image/png';
  if (name.endsWith('.jpg') || name.endsWith('.jpeg')) return 'image/jpeg';
  if (name.endsWith('.pdf')) return 'application/pdf';
  if (name.endsWith('.doc')) return 'application/msword';
  if (name.endsWith('.docx')) return 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
  if (name.endsWith('.xlsx')) return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  if (name.endsWith('.xls')) return 'application/vnd.ms-excel';
  if (name.endsWith('.csv')) return 'text/csv';
  return 'application/octet-stream';
}

function uploadBase64File(base64Data, fileName, folderName) {
  try {
    if (!base64Data) {
      return { success: false, message: 'No file data supplied' };
    }

    const folder = folderName ? DriveApp.getFoldersByName(folderName) : null;
    const destinationFolder = folder && folder.hasNext() ? folder.next() : getUploadsFolder();

    const cleanData = String(base64Data || '').includes(',') ? String(base64Data).split(',')[1] : String(base64Data || '');
    const bytes = Utilities.base64Decode(cleanData);
    const blob = Utilities.newBlob(bytes, getMimeTypeFromFilename(fileName), fileName);

    const uploadedFile = destinationFolder.createFile(blob);
    uploadedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return {
      success: true,
      url: 'https://drive.google.com/uc?export=view&id=' + uploadedFile.getId(),
      fileId: uploadedFile.getId()
    };
  } catch (err) {
    return { success: false, message: String(err) };
  }
}

// ==================== INITIALIZATION ====================

function setupDatabase() {
  return ensureDatabase();
}
