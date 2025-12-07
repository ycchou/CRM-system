/* =========================================================================
   CRM 系統後端邏輯 (Code.gs) - 含 Admin 管理功能 & 試算表選單跳轉
   ========================================================================= */

const CONFIG = {
  SHEET_NAMES: { 
    USERS: "Users", 
    MEMBERS: "會員資料", 
    DASHBOARD: "儀表板" 
  }
};

/* --- 1. WEB APP 基礎設定 & 試算表選單 --- */

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('會員管理系統')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSheetUrl() {
  return SpreadsheetApp.getActiveSpreadsheet().getUrl();
}

// [NEW] 當試算表開啟時，建立自訂選單
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('工會 CRM 系統')
    .addItem('🚀 開啟 CRM 系統', 'openWebApp')
    .addToUi();
}

// [NEW] 執行跳轉到 Web App 的動作 (已修正為固定網址)
function openWebApp() {
  // 將網址固定為您提供的正確版本
  var url = "https://script.google.com/macros/s/AKfycby1nzC1VZtyWYoph-vqY-UDCa2oPDrRXVUJ_0AHVQPUZm2QgynmeHkUr9sH0dxHwLBs/exec";
  
  // 透過 HTML Service 執行 client-side script 來開新分頁
  var html = HtmlService.createHtmlOutput(
    '<html><script>' +
    'window.open("' + url + '", "_blank");' +
    'google.script.host.close();' +
    '</script></html>'
  ).setWidth(250).setHeight(50);
  
  SpreadsheetApp.getUi().showModalDialog(html, '正在開啟系統...');
}

/* --- 2. 帳號與權限管理 (Auth System) --- */

function verifyToken(token) {
  if (!token) return { valid: false, message: "無 Token" };
  try {
    const sheet = getSheet(CONFIG.SHEET_NAMES.USERS);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][6] === token) {
        if (data[i][7] && new Date() > new Date(data[i][7])) {
          return { valid: false, message: "登入逾時，請重新登入" };
        }
        return { valid: true, username: data[i][1], name: data[i][2], role: data[i][5], uid: data[i][0] };
      }
    }
    return { valid: false, message: "無效的 Token，請重新登入" };
  } catch (e) {
    return { valid: false, message: "驗證錯誤: " + e.message };
  }
}

function loginUser(email, pass) {
  try {
    const sheet = getSheet(CONFIG.SHEET_NAMES.USERS);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === email) {
        if (generateHash(pass, data[i][4]) === data[i][3]) {
          if (data[i][5] === 'Pending') return { success: false, message: "帳號審核中，請聯繫管理員" };
          const token = Utilities.getUuid();
          sheet.getRange(i + 1, 7).setValue(token);
          sheet.getRange(i + 1, 8).setValue(new Date(Date.now() + 3600000));
          return { success: true, token: token, role: data[i][5], username: email, name: data[i][2] };
        }
      }
    }
    return { success: false, message: "帳號或密碼錯誤" };
  } catch (e) {
    return { success: false, message: "系統錯誤: " + e.message };
  }
}

function handleRegister(email, pass, name) {
  try {
    const sheet = getSheet(CONFIG.SHEET_NAMES.USERS);
    const data = sheet.getDataRange().getValues();
    if (data.length > 1 && data.slice(1).some(r => r[1] === email)) {
      return { success: false, message: "此 Email 已存在" };
    }
    const salt = generateSalt(10);
    sheet.appendRow([Utilities.getUuid(), email, name, generateHash(pass, salt), salt, 'Pending', '', '', new Date()]);
    return { success: true, message: "註冊申請已送出，請等待管理員開通權限" };
  } catch (e) {
    return { success: false, message: "註冊錯誤: " + e.message };
  }
}

// 取得所有使用者列表 (僅 Admin 可用)
function getAllUsers(token) {
  const user = verifyToken(token);
  if (!user.valid || user.role !== 'Admin') throw new Error("權限不足");
  
  const sheet = getSheet(CONFIG.SHEET_NAMES.USERS);
  const data = sheet.getDataRange().getValues();
  // 回傳 id, username(email), name, role
  return data.slice(1).map(r => ({ id: r[0], username: r[1], name: r[2], role: r[5] }));
}

// 管理員更新使用者 (修改權限或刪除)
function adminUpdateUser(token, targetUid, action, newRole) {
  const user = verifyToken(token);
  if (!user.valid || user.role !== 'Admin') throw new Error("權限不足");
  if (targetUid === user.uid && action === 'delete') throw new Error("不能刪除自己");

  const sheet = getSheet(CONFIG.SHEET_NAMES.USERS);
  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === targetUid) { rowIndex = i + 1; break; }
  }
  if (rowIndex === -1) return { success: false, message: "找不到使用者" };

  if (action === 'delete') {
    sheet.deleteRow(rowIndex);
    return { success: true, message: "已刪除使用者" };
  } else {
    // Column F is index 6
    sheet.getRange(rowIndex, 6).setValue(newRole);
    return { success: true, message: "權限已更新" };
  }
}

/* --- 3. CRM 業務邏輯 --- */

function getDashboardAnalytics(token) {
  const user = verifyToken(token);
  if (!user.valid) throw new Error(user.message);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DASHBOARD);
  if (!dashSheet) throw new Error(`找不到工作表: ${CONFIG.SHEET_NAMES.DASHBOARD}`);
  
  const dashData = dashSheet.getDataRange().getValues();
  let totalActiveMembers = 0, newThisMonth = 0, leavesThisMonth = 0;
  let genderArr = [], laborArr = [], campusArr = [], joinLeaveByYear = [], joinLeaveByMonth = [];
  const now = new Date();
  const currentYear = now.getFullYear();
  const currentMonth = now.getMonth() + 1;

  let currentSection = ''; 
  for (let i = 1; i < dashData.length; i++) { 
    const label = String(dashData[i][0]).trim();
    const value = dashData[i][1];
    if (label === '會籍狀態') { currentSection = 'STATUS'; continue; }
    if (label === '性別') { currentSection = 'GENDER'; continue; }
    if (label === '勞工類別') { currentSection = 'LABOR'; continue; }
    if (label === '院區') { currentSection = 'CAMPUS'; continue; }
    if (label && (value !== '' && value !== undefined)) {
      if (currentSection === 'STATUS') { if (label === '在籍') totalActiveMembers = value; } 
      else if (currentSection === 'GENDER') genderArr.push({ key: label, value: value });
      else if (currentSection === 'LABOR') laborArr.push({ key: label, value: value });
      else if (currentSection === 'CAMPUS') campusArr.push({ key: label, value: value });
    }
  }

  for (let i = 2; i < dashData.length; i++) {
    const yYear = dashData[i][3];
    if (typeof yYear === 'number' && yYear > 2000) joinLeaveByYear.push({ key: String(yYear), joins: dashData[i][4] || 0, leaves: dashData[i][5] || 0 });
    const mYear = dashData[i][7];
    const mMonth = dashData[i][8];
    if (typeof mYear === 'number' && typeof mMonth === 'number') {
      const paddedMonth = mMonth < 10 ? '0' + mMonth : mMonth;
      joinLeaveByMonth.push({ key: `${mYear}-${paddedMonth}`, joins: dashData[i][9] || 0, leaves: dashData[i][10] || 0 });
      if (mYear === currentYear && mMonth === currentMonth) { newThisMonth = dashData[i][9] || 0; leavesThisMonth = dashData[i][10] || 0; }
    }
  }

  const memberSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MEMBERS);
  const memberData = memberSheet.getDataRange().getValues();
  const headers = memberData[0];
  const colMap = {};
  headers.forEach((h, i) => colMap[String(h).trim()] = i);
  const ageStats = { '20-29': 0, '30-39': 0, '40-49': 0, '50-59': 0, '60+': 0 };
  const memberTypeStats = {};
  let totalAge = 0, ageCount = 0;
  for (let i = 1; i < memberData.length; i++) {
    const row = memberData[i];
    const statusRaw = String(row[colMap['會籍狀態']] || '').trim();
    if (!statusRaw.includes('退會')) {
      const birthRaw = row[colMap['生日']];
      const mType = row[colMap['會員類別']];
      if (mType) memberTypeStats[mType] = (memberTypeStats[mType] || 0) + 1;
      if (birthRaw instanceof Date) {
        const age = currentYear - birthRaw.getFullYear();
        if (age > 0 && age < 120) {
          totalAge += age; ageCount++;
          if (age < 30) ageStats['20-29']++; else if (age < 40) ageStats['30-39']++;
          else if (age < 50) ageStats['40-49']++; else if (age < 60) ageStats['50-59']++; else ageStats['60+']++;
        }
      }
    }
  }
  const avgAge = ageCount > 0 ? (totalAge / ageCount).toFixed(1) : 0;
  const sortByValueDesc = (a, b) => b.value - a.value;

  return {
    totalActiveMembers, newThisMonth, leavesThisMonth, avgAge,
    campusDistribution: campusArr.sort(sortByValueDesc),
    genderDistribution: genderArr.sort(sortByValueDesc),
    laborTypeDistribution: laborArr.sort(sortByValueDesc),
    memberTypeDistribution: Object.keys(memberTypeStats).map(k => ({ key: k, value: memberTypeStats[k] })).sort(sortByValueDesc),
    ageDistribution: Object.keys(ageStats).map(k => ({ key: k, value: ageStats[k] })),
    joinLeaveByMonth: joinLeaveByMonth.slice(-24), joinLeaveByYear
  };
}

function getMembers(token) {
  const user = verifyToken(token);
  if (!user.valid) throw new Error(user.message);
  const sheet = getSheet(CONFIG.SHEET_NAMES.MEMBERS);
  const data = sheet.getDataRange().getValues();
  const rawHeaders = data[0];
  const members = [];
  const timeZone = Session.getScriptTimeZone();
  const colMap = {};
  rawHeaders.forEach((h, i) => colMap[String(h).trim()] = i);
  const statusIdx = colMap['會籍狀態'];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0] && !row[1]) continue;
    const statusRaw = String(row[statusIdx] || '').trim();
    if (statusRaw.includes('退會')) continue;
    let memberObj = { row: i + 1 };
    rawHeaders.forEach((h, index) => {
      let val = row[index];
      if (val instanceof Date) val = Utilities.formatDate(val, timeZone, 'yyyy-MM-dd');
      memberObj[h] = val;
    });
    members.push(memberObj);
  }
  return members.reverse();
}

function addMember(token, formObject) {
  const user = verifyToken(token);
  if (!user.valid) throw new Error(user.message);
  if (user.role === 'Viewer') throw new Error("權限不足：您僅能檢視資料");
  const sheet = getSheet(CONFIG.SHEET_NAMES.MEMBERS);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const newRow = [];
  if (!formObject['會籍狀態']) formObject['會籍狀態'] = '在籍';
  if (!formObject['入會日期']) formObject['入會日期'] = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd');
  headers.forEach(header => newRow.push(formObject[header] || ''));
  sheet.appendRow(newRow);
  return `成功新增會員：${formObject['姓名']}`;
}

function updateMember(token, formObject) {
  const user = verifyToken(token);
  if (!user.valid) throw new Error(user.message);
  if (user.role === 'Viewer') throw new Error("權限不足");
  const sheet = getSheet(CONFIG.SHEET_NAMES.MEMBERS);
  const rowIndex = parseInt(formObject.row);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const currentRow = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
  const updatedRow = [];
  headers.forEach((header, index) => updatedRow.push(formObject.hasOwnProperty(header) ? formObject[header] : currentRow[index]));
  sheet.getRange(rowIndex, 1, 1, updatedRow.length).setValues([updatedRow]);
  return `成功更新會員資料：${formObject['姓名']}`;
}

function deleteMember(token, rowIndex) {
  const user = verifyToken(token);
  if (!user.valid) throw new Error(user.message);
  if (user.role === 'Viewer') throw new Error("權限不足");
  const sheet = getSheet(CONFIG.SHEET_NAMES.MEMBERS);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colMap = {};
  headers.forEach((h, i) => colMap[String(h).trim()] = i);
  const statusIdx = colMap['會籍狀態'];
  const leaveDateIdx = colMap['退會日期'];
  if (statusIdx !== undefined) {
     sheet.getRange(rowIndex, statusIdx + 1).setValue('退會');
     if (leaveDateIdx !== undefined) sheet.getRange(rowIndex, leaveDateIdx + 1).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd'));
     return "已將該會員標記為「退會」";
  } else {
     sheet.deleteRow(parseInt(rowIndex));
     return "已刪除該筆資料";
  }
}

function reconcileDues(token, reportData) {
  const user = verifyToken(token);
  if (!user.valid) return { success: false, error: user.message };
  try {
    const sheet = getSheet(CONFIG.SHEET_NAMES.MEMBERS);
    const dbData = sheet.getDataRange().getValues();
    const headers = dbData[0];
    const colMap = {};
    headers.forEach((h, i) => colMap[String(h).trim()] = i);
    const dbMembers = {};
    const dbKeyField = '員工編號'; 
    const reportKeyField = '員編';   
    if (colMap[dbKeyField] === undefined) return { success: false, error: `資料庫中找不到「${dbKeyField}」欄位` };
    
    // 建立現有會員索引
    for (let i = 1; i < dbData.length; i++) {
      const row = dbData[i];
      const key = String(row[colMap[dbKeyField]] || '').trim();
      if (key) dbMembers[key] = { id: key, name: row[colMap['姓名']], campus: row[colMap['院區']], unit: row[colMap['單位']], isLeft: String(row[colMap['會籍狀態']] || '').includes('退會'), matched: false };
    }
    
    const missing = [], extra = [], newMembers = [], specialCases = [];  
    
    reportData.forEach(row => {
      const reportId = String(row[reportKeyField] || '').trim();
      if(!reportId) return;
      const reportName = row['姓名'] || row['名字'];
      
      // === 核心修改：強力清洗資料 (移除單引號與非數字)，並取絕對值 ===
      let rawAmount = row['金額'] || row['扣款金額'] || 0;
      let cleanString = String(rawAmount).replace(/[^0-9-]/g, ''); 
      const amount = Math.abs(parseInt(cleanString) || 0);
      // ========================================================
      
      const reportCampus = row['院區'] || '';
      
      // === 邏輯修改：只要金額符合，優先歸類為新入會/特殊個案 (無論是否在資料庫) ===
      if (amount === 1450) {
          newMembers.push({ id: reportId, name: reportName, campus: reportCampus, amount: amount });
          // 若在資料庫中也標記為已核對，避免出現在 Missing
          if (dbMembers[reportId]) dbMembers[reportId].matched = true;
      } else if (amount === 1700) {
          specialCases.push({ id: reportId, name: reportName, campus: reportCampus, amount: amount });
          if (dbMembers[reportId]) dbMembers[reportId].matched = true;
      } else {
          // 其他金額 (如 250) 走標準流程
          if (dbMembers[reportId]) {
            dbMembers[reportId].matched = true;
            if (dbMembers[reportId].isLeft) extra.push({ id: reportId, name: reportName, campus: reportCampus, amount: amount });
          } else {
            // 不在資料庫且金額不是1450/1700 -> 歸類為新成員待確認
            newMembers.push({ id: reportId, name: reportName, campus: reportCampus, amount: amount });
          }
      }
    });
    
    for (const key in dbMembers) if (!dbMembers[key].isLeft && !dbMembers[key].matched) missing.push({ id: dbMembers[key].id, name: dbMembers[key].name, campus: dbMembers[key].campus, unit: dbMembers[key].unit });
    
    return { success: true, results: { missing, extra, newMembers, specialCases } };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function getSheet(name) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); }
function generateHash(input, salt) { return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, input + salt).map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join(''); }
function generateSalt(len) { let s = "", c = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"; for (let i = 0; i < len; i++) s += c.charAt(Math.floor(Math.random() * c.length)); return s; }
