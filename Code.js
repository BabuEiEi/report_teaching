// --- CONFIGURATION ---
// กรุณาแทนที่ด้วย ID ของ Google Sheet และ Google Drive Folder ของคุณ
const SPREADSHEET_ID = 'กรุณาใส่ ID Google Sheet'; 
const FOLDER_ID = 'กรุณาใส่ ID Folder ใน Google Drive'; 

// --- MAIN HANDLERS ---

/**
 * ฟังก์ชันหลักสำหรับรับค่าจาก Frontend (Method POST)
 * ใช้สำหรับ API Call ต่างๆ (Login, Save, Get Data)
 */
function doPost(e) {
  try {
    // ตรวจสอบว่ามีข้อมูลส่งมาหรือไม่
    if (!e || !e.parameter) {
      throw new Error("No parameters found");
    }

    const params = e.parameter;
    const action = params.action;
    
    // ตรวจสอบว่ามีการส่งข้อมูล data มาหรือไม่
    let payload = {};
    if (params.data) {
      payload = JSON.parse(params.data);
    }
    
    let result = {};
    
    // เลือกฟังก์ชันที่จะทำงานตาม action ที่ส่งมา
    switch(action) {
      case 'login':
        result = handleLogin(payload);
        break;
      case 'getDashboardData':
        result = getDashboardData();
        break;
      case 'saveReport':
        result = saveReport(payload);
        break;
      case 'getReports':
        result = getReports(payload);
        break;
      case 'deleteReport':
        result = deleteReport(payload);
        break;
      case 'approveReport':
        result = approveReport(payload);
        break;
      case 'getSettings':
        result = getSettings();
        break;
      case 'saveSettings':
        result = saveSettings(payload);
        break;
      case 'setupDatabase':
        result = setupDatabase();
        break;
      case 'getUsers':
        result = getUsers();
        break;
      case 'saveUser':
        result = saveUser(payload);
        break;
      case 'deleteUser':
        result = deleteUser(payload);
        break;
      default:
        result = { success: false, message: 'Invalid Action: ' + action };
    }
    
    // ส่งค่ากลับเป็น JSON
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (err) {
    // กรณีเกิด Error ให้ส่งกลับไปบอก Frontend
    return ContentService.createTextOutput(JSON.stringify({ 
      success: false, 
      message: err.toString(), 
      stack: err.stack 
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * ฟังก์ชันหลักที่ทำงานเมื่อผู้ใช้เปิด Web App (Method GET)
 * ทำหน้าที่แสดงผลไฟล์ index.html
 */
function doGet(e) {
  // สร้าง Output จากไฟล์ index.html
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('ระบบรายงานการสอนออนไลน์') // ตั้งชื่อ Tab Browser
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL) // อนุญาตให้แสดงผลใน iframe หรือเว็บอื่นได้
      .addMetaTag('viewport', 'width=device-width, initial-scale=1'); // รองรับการแสดงผลบนมือถือ
}

// --- HELPER FUNCTIONS ---

function getDb() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// 0. Setup Database
function setupDatabase() {
  const ss = getDb();
  
  // 1. Users Sheet
  let sheetUsers = ss.getSheetByName('Users');
  if (!sheetUsers) {
    sheetUsers = ss.insertSheet('Users');
    sheetUsers.appendRow(['Username', 'Password', 'Name', 'Role']);
    sheetUsers.appendRow(['admin', 'admin1234', 'Administrator', 'admin']); // Default Admin
  }
  
  // 2. Reports Sheet
  let sheetReports = ss.getSheetByName('Reports');
  if (!sheetReports) {
    sheetReports = ss.insertSheet('Reports');
    // Header based on existing code indices
    sheetReports.appendRow([
      'ID', 'Timestamp', 'Term', 'TeacherName', 'Subject', 'Code', 
      'ClassLevel', 'Date', 'Method', 'PlanLink', 'MediaLink', 'Activity', 
      'Assessment', 'Problems', 'Suggestions', 'EvidenceURL', 'AdminComment', 
      'Username', 'Department', 'Status', 'TimePeriod'
    ]);
  }
  
  // 3. Settings Sheet
  let sheetSettings = ss.getSheetByName('Settings');
  if (!sheetSettings) {
    sheetSettings = ss.insertSheet('Settings');
    sheetSettings.appendRow(['Key', 'Value']);
    sheetSettings.appendRow(['schoolName', 'โรงเรียนตัวอย่าง']);
    sheetSettings.appendRow(['directorName', 'ผู้อำนวยการ']);
    sheetSettings.appendRow(['departments', 'ภาษาไทย,คณิตศาสตร์,วิทยาศาสตร์,สังคมศึกษา,ภาษาอังกฤษ']);
  }
  
  return { success: true, message: 'Database setup complete. Default admin created (user: admin, pass: admin1234)' };
}

// 1. Authentication
function handleLogin(data) {
  const sheet = getDb().getSheetByName('Users');
  if (!sheet) return { success: false, message: 'Sheet "Users" not found' };

  const rows = sheet.getDataRange().getValues();
  // Header: Username, Password, Name, Role
  
  for (let i = 1; i < rows.length; i++) {
    // เปรียบเทียบ Username และ Password (ควรระวังเรื่อง Case Sensitive ถ้าต้องการ)
    if (String(rows[i][0]) === String(data.username) && String(rows[i][1]) === String(data.password)) {
      return { 
        success: true, 
        user: { 
          username: rows[i][0], 
          name: rows[i][2], 
          role: rows[i][3] 
        } 
      };
    }
  }
  return { success: false, message: 'ชื่อผู้ใช้งานหรือรหัสผ่านไม่ถูกต้อง' };
}

// 2. Dashboard Stats
function getDashboardData() {
  const sheet = getDb().getSheetByName('Reports');
  if (!sheet) return { success: false, message: 'Sheet "Reports" not found' };

  const rows = sheet.getDataRange().getValues();
  const data = rows.slice(1); // ตัด Header ออก
  
  let stats = {
    total: data.length,
    online: 0,
    demand: 0,
    methods: {},
    teacherCounts: {}
  };
  
  data.forEach(r => {
    // r[8] = Method, r[3] = TeacherName
    const method = r[8];
    const teacher = r[3];
    
    if(method === 'Online') stats.online++;
    if(method === 'On-Demand') stats.demand++;
    
    // นับจำนวนรูปแบบการสอน
    stats.methods[method] = (stats.methods[method] || 0) + 1;
    
    // นับจำนวนรายงานของครูแต่ละคน
    if (teacher) {
      stats.teacherCounts[teacher] = (stats.teacherCounts[teacher] || 0) + 1;
    }
  });
  
  // จัดอันดับครู Top 10
  const topTeachers = Object.keys(stats.teacherCounts).map(name => {
    return { name: name, count: stats.teacherCounts[name] };
  }).sort((a,b) => b.count - a.count).slice(0, 10);
  
  return { success: true, data: { ...stats, topTeachers } };
}

// 3. Save Report (Insert or Update) & Image Upload
function saveReport(data) {
  const ss = getDb();
  const sheet = ss.getSheetByName('Reports');
  if (!sheet) return { success: false, message: 'Sheet "Reports" not found' };

  let evidenceUrl = data.existingEvidence || '';
  
  // จัดการอัปโหลดรูปภาพ (ถ้ามี multiple files)
  if (data.evidenceFiles && data.evidenceFiles.length > 0) {
    try {
      const mainFolder = DriveApp.getFolderById(FOLDER_ID);
      
      // 1. หาหรือสร้างโฟลเดอร์ของครู (ตามชื่อครู)
      const folderName = data.teacherName || data.username;
      const folders = mainFolder.getFoldersByName(folderName);
      let userFolder;
      
      if (folders.hasNext()) {
        userFolder = folders.next();
      } else {
        userFolder = mainFolder.createFolder(folderName);
      }
      
      // 2. อัปโหลดรูปภาพทั้งหมดลงในโฟลเดอร์ของครู
      let uploadedUrls = [];
      data.evidenceFiles.forEach(fileData => {
        const blob = Utilities.newBlob(Utilities.base64Decode(fileData.base64), fileData.mime, fileData.name);
        const file = userFolder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        uploadedUrls.push("https://lh3.googleusercontent.com/d/" + file.getId());
      });
      
      // 3. รวมลิงก์เป็น string เดียว (คั่นด้วย comma)
      evidenceUrl = uploadedUrls.join(',');
      
    } catch(e) {
      return { success: false, message: 'Upload Failed: ' + e.message };
    }
  }

  const timestamp = new Date();
  
  // ตรวจสอบว่าเป็นโหมดแก้ไขหรือเพิ่มใหม่
  if (data.id) {
    // --- โหมดแก้ไข (Edit) ---
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.id)) { // คอลัมน์ 0 คือ ID
        const rowNum = i + 1;
        
        // อัปเดตข้อมูลทีละเซลล์ให้ตรงกับคอลัมน์
        sheet.getRange(rowNum, 3).setValue(data.term);        // Term
        sheet.getRange(rowNum, 5).setValue(data.subject);     // Subject
        sheet.getRange(rowNum, 6).setValue(data.code);        // Code
        sheet.getRange(rowNum, 7).setValue(data.classLevel);  // Class
        sheet.getRange(rowNum, 8).setValue(data.date);        // Date
        sheet.getRange(rowNum, 9).setValue(data.method);      // Method
        sheet.getRange(rowNum, 10).setValue(data.planLink);   // Plan Link
        sheet.getRange(rowNum, 12).setValue(data.activity);   // Activity
        sheet.getRange(rowNum, 13).setValue(data.assessment); // Assessment
        sheet.getRange(rowNum, 14).setValue(data.problems);   // Problems
        sheet.getRange(rowNum, 15).setValue(data.suggestions);// Suggestions
        sheet.getRange(rowNum, 16).setValue(evidenceUrl);     // Evidence URL (Updated if new files uploaded)
        sheet.getRange(rowNum, 19).setValue(data.department); // Department
        sheet.getRange(rowNum, 21).setValue(data.timePeriod); // TimePeriod
        
        return { success: true, message: 'Updated successfully' };
      }
    }
    return { success: false, message: 'Report ID not found for update' };
    
  } else {
    // --- โหมดเพิ่มใหม่ (Insert) ---
    const newId = Utilities.getUuid();
    // เรียงลำดับข้อมูลให้ตรงกับ Header ของ Sheet Reports
    // 0:ID, 1:Timestamp, 2:Term, 3:TeacherName, 4:Subject, 5:Code, 
    // 6:Class, 7:Date, 8:Method, 9:PlanLink, 10:MediaLink(Blank), 11:Activity, 
    // 12:Assessment, 13:Problems, 14:Suggestions, 15:EvidenceURL, 
    // 16:AdminComment, 17:Username, 18:Department, 19:Status, 20:TimePeriod
    sheet.appendRow([
      newId, 
      timestamp, 
      data.term, 
      data.teacherName, 
      data.subject, 
      data.code,
      data.classLevel, 
      data.date, 
      data.method, 
      data.planLink, 
      '', // MediaLink (ไม่ได้ใช้ในฟอร์มปัจจุบัน)
      data.activity, 
      data.assessment, 
      data.problems, 
      data.suggestions, 
      evidenceUrl, 
      '', // AdminComment เริ่มต้นว่าง
      data.username, 
      data.department, 
      'Pending', // Status เริ่มต้น
      data.timePeriod // TimePeriod
    ]);
    
    return { success: true, message: 'Saved successfully' };
  }
}

// 4. Get Reports List
function getReports(data) {
  const sheet = getDb().getSheetByName('Reports');
  if (!sheet) return { success: false, message: 'Sheet "Reports" not found' };

  const rows = sheet.getDataRange().getValues();
  const reports = [];
  
  // เริ่มที่ i=1 เพื่อข้าม Header
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    
    // แปลงข้อมูลแถวเป็น Object
    const obj = {
      id: r[0],
      term: r[2],
      teacherName: r[3],
      subject: r[4],
      code: r[5],
      classLevel: r[6],
      date: formatDate(r[7]), // จัดรูปแบบวันที่ให้สวยงามถ้าจำเป็น
      method: r[8],
      planLink: r[9],
      activity: r[11],
      assessment: r[12],
      problems: r[13],
      suggestions: r[14],
      evidenceUrl: r[15],
      adminComment: r[16],
      username: r[17],
      department: r[18],
      status: r[19] || 'Pending',
      timePeriod: r[20] || ''
    };
    
    // Logic การกรองข้อมูล
    if (data.role === 'admin') {
      let isMatch = true;
      
      // กรองตามกลุ่มสาระ (ถ้ามีส่งมา)
      if (data.filterDepartment && data.filterDepartment !== "" && obj.department !== data.filterDepartment) {
        isMatch = false;
      }
      
      // กรองตามชื่อครู (ถ้ามีส่งมา)
      if (data.filterTeacher && data.filterTeacher !== "" && obj.teacherName !== data.filterTeacher) {
        isMatch = false;
      }
      
      if (isMatch) reports.push(obj);
      
    } else if (data.username === obj.username) {
      // ครูเห็นเฉพาะของตัวเอง
      reports.push(obj);
    }
  }
  return { success: true, data: reports };
}

// Helper: Format Date Object to String (ISO for HTML Input)
function formatDate(dateVal) {
  if (dateVal instanceof Date) {
    // คืนค่าเป็น ISO String (YYYY-MM-DD) เพื่อให้ input type="date" อ่านง่าย หรือส่งไปแสดงผล
    return Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return dateVal;
}

// 5. Delete Report
function deleteReport(data) {
  const sheet = getDb().getSheetByName('Reports');
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, message: 'ID not found' };
}

// 6. Approve Report (Admin)
function approveReport(data) {
  const sheet = getDb().getSheetByName('Reports');
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      sheet.getRange(i+1, 17).setValue(data.comment); // Admin Comment (Col 16 -> Index 17 in human count)
      sheet.getRange(i+1, 20).setValue('Approved'); // Status (Col 19 -> Index 20 in human count)
      return { success: true };
    }
  }
  return { success: false, message: 'ID not found' };
}

// 7. System Settings
function getSettings() {
  const sheet = getDb().getSheetByName('Settings');
  if (!sheet) return { success: false, data: {} }; // Return empty if not setup yet

  const rows = sheet.getDataRange().getValues();
  let settings = {};
  rows.forEach(r => {
    if(r[0]) settings[r[0]] = r[1];
  });
  return { success: true, data: settings };
}

function saveSettings(data) {
  const sheet = getDb().getSheetByName('Settings');
  sheet.clear(); // ล้างค่าเก่าทั้งหมด
  
  // บันทึกค่าใหม่
  if (data.schoolName) sheet.appendRow(['schoolName', data.schoolName]);
  if (data.directorName) sheet.appendRow(['directorName', data.directorName]);
  if (data.departments) sheet.appendRow(['departments', data.departments]);
  
  return { success: true };
}

// 8. User Management
function getUsers() {
  const sheet = getDb().getSheetByName('Users');
  if (!sheet) return { success: false, message: 'Sheet "Users" not found' };

  const rows = sheet.getDataRange().getValues();
  // ข้าม Header และ map ข้อมูล
  const users = rows.slice(1).map(r => ({ 
    username: r[0], 
    password: r[1], // ใน Production จริงไม่ควรส่ง Password กลับมา แต่เพื่อความง่ายในโปรเจกต์นี้
    name: r[2], 
    role: r[3] 
  }));
  return { success: true, data: users };
}

function saveUser(data) {
  const sheet = getDb().getSheetByName('Users');
  
  // ตรวจสอบ Username ซ้ำ
  const rows = sheet.getDataRange().getValues();
  for(let i=1; i<rows.length; i++) {
    if(String(rows[i][0]) === String(data.username)) {
      return { success: false, message: 'Username นี้มีอยู่ในระบบแล้ว' };
    }
  }
  
  sheet.appendRow([data.username, data.password, data.name, data.role]);
  return { success: true };
}

function deleteUser(data) {
  const sheet = getDb().getSheetByName('Users');
  const rows = sheet.getDataRange().getValues();
  for(let i=1; i<rows.length; i++) {
    if(String(rows[i][0]) === String(data.username)) {
      sheet.deleteRow(i+1);
      return { success: true };
    }
  }
  return { success: false, message: 'Username not found' };
}
