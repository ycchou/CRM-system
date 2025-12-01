/* =========================================================================
   CRM 系統後端邏輯 (Code.gs) - 儀表板直讀版
   功能：
   1. 核心數據 (KPI、趨勢、分佈) 直接讀取 '儀表板' 工作表，確保與 Excel 顯示一致。
   2. 自動將分佈數據 (院區、性別等) 由大到小排序。
   3. 僅「年齡分佈」與「會員類別」因儀表板未列出，維持從原始資料計算。
   ========================================================================= */

const SHEET_NAME_MEMBERS = '會員資料';
const SHEET_NAME_DASHBOARD = '儀表板';

/* --- WEB APP 基礎設定 --- */

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('工會會員管理系統')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSheetUrl() {
  return SpreadsheetApp.getActiveSpreadsheet().getUrl();
}

/* --- 儀表板數據運算 (DASHBOARD ANALYTICS) --- */

function getDashboardAnalytics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashSheet = ss.getSheetByName(SHEET_NAME_DASHBOARD);
  if (!dashSheet) throw new Error(`找不到工作表: ${SHEET_NAME_DASHBOARD}`);
  
  // 取得儀表板所有資料
  const dashData = dashSheet.getDataRange().getValues();
  
  // 初始化回傳變數
  let totalActiveMembers = 0;
  let newThisMonth = 0;
  let leavesThisMonth = 0;
  
  // 用陣列儲存統計數據 (準備進行排序)
  let genderArr = [];
  let laborArr = [];
  let campusArr = [];
  let joinLeaveByYear = [];
  let joinLeaveByMonth = [];

  // 設定當前時間 (用於抓取本月KPI)
  const now = new Date();
  const currentYear = now.getFullYear();
  const currentMonth = now.getMonth() + 1; // JS 0-11, Sheet 1-12

  // --- 步驟 1: 解析 A:B 欄位 (分佈與KPI) ---
  // 使用狀態機模式 (State Machine) 讀取區塊
  let currentSection = ''; 

  for (let i = 1; i < dashData.length; i++) { // 跳過第一列標題
    const label = String(dashData[i][0]).trim(); // A欄
    const value = dashData[i][1];                // B欄

    // 偵測區塊標題
    if (label === '會籍狀態') { currentSection = 'STATUS'; continue; }
    if (label === '性別') { currentSection = 'GENDER'; continue; }
    if (label === '勞工類別') { currentSection = 'LABOR'; continue; }
    if (label === '院區') { currentSection = 'CAMPUS'; continue; }

    // 根據當前區塊讀取數據 (若 Value 有值)
    if (label && (value !== '' && value !== undefined)) {
      if (currentSection === 'STATUS') {
        if (label === '在籍') totalActiveMembers = value;
      } else if (currentSection === 'GENDER') {
        genderArr.push({ key: label, value: value });
      } else if (currentSection === 'LABOR') {
        laborArr.push({ key: label, value: value });
      } else if (currentSection === 'CAMPUS') {
        campusArr.push({ key: label, value: value });
      }
    }
  }

  // --- 步驟 2: 解析 D:F 欄位 (年份趨勢) ---
  // D=年份, E=入會, F=退會 (Index: 3, 4, 5)
  for (let i = 2; i < dashData.length; i++) {
    const yYear = dashData[i][3];
    const yJoin = dashData[i][4];
    const yLeave = dashData[i][5];

    if (typeof yYear === 'number' && yYear > 2000) {
      joinLeaveByYear.push({
        key: String(yYear),
        joins: yJoin || 0,
        leaves: yLeave || 0
      });
    }
  }

  // --- 步驟 3: 解析 H:K 欄位 (月份趨勢) ---
  // H=年份, I=月份, J=入會, K=退會 (Index: 7, 8, 9, 10)
  for (let i = 2; i < dashData.length; i++) {
    const mYear = dashData[i][7];
    const mMonth = dashData[i][8];
    const mJoin = dashData[i][9];
    const mLeave = dashData[i][10];

    if (typeof mYear === 'number' && typeof mMonth === 'number') {
      const paddedMonth = mMonth < 10 ? '0' + mMonth : mMonth;
      const monthKey = `${mYear}-${paddedMonth}`;
      
      joinLeaveByMonth.push({
        key: monthKey,
        joins: mJoin || 0,
        leaves: mLeave || 0
      });

      // 判斷本月 KPI
      if (mYear === currentYear && mMonth === currentMonth) {
        newThisMonth = mJoin || 0;
        leavesThisMonth = mLeave || 0;
      }
    }
  }

  // --- 步驟 4: 補足儀表板缺少的數據 (從原始資料計算) ---
  // 儀表板 CSV 內沒有「年齡」與「會員類別(一般/贊助)」，需從原始表計算
  const memberSheet = ss.getSheetByName(SHEET_NAME_MEMBERS);
  const memberData = memberSheet.getDataRange().getValues();
  const headers = memberData[0];
  const colMap = {};
  headers.forEach((h, i) => colMap[String(h).trim()] = i);
  
  const ageStats = { '20-29': 0, '30-39': 0, '40-49': 0, '50-59': 0, '60+': 0 };
  const memberTypeStats = {}; // 用於補充會員類別
  let totalAge = 0;
  let ageCount = 0;

  for (let i = 1; i < memberData.length; i++) {
    const row = memberData[i];
    const statusRaw = String(row[colMap['會籍狀態']] || '').trim();
    
    // 只計算在籍
    if (!statusRaw.includes('退會')) {
      const birthRaw = row[colMap['生日']];
      const mType = row[colMap['會員類別']]; // 一般會員/贊助會員

      // 統計會員類別
      if (mType) {
        memberTypeStats[mType] = (memberTypeStats[mType] || 0) + 1;
      }

      // 統計年齡
      if (birthRaw instanceof Date) {
        const age = currentYear - birthRaw.getFullYear();
        if (age > 0 && age < 120) {
          totalAge += age;
          ageCount++;
          if (age < 30) ageStats['20-29']++;
          else if (age < 40) ageStats['30-39']++;
          else if (age < 50) ageStats['40-49']++;
          else if (age < 60) ageStats['50-59']++;
          else ageStats['60+']++;
        }
      }
    }
  }
  const avgAge = ageCount > 0 ? (totalAge / ageCount).toFixed(1) : 0;


  // --- 步驟 5: 排序與回傳 (Sort High to Low) ---
  // 輔助排序函式
  const sortByValueDesc = (a, b) => b.value - a.value;

  return {
    totalActiveMembers: totalActiveMembers,
    newThisMonth: newThisMonth,
    leavesThisMonth: leavesThisMonth,
    avgAge: avgAge,
    
    // 已排序的分佈數據
    campusDistribution: campusArr.sort(sortByValueDesc),
    genderDistribution: genderArr.sort(sortByValueDesc),
    laborTypeDistribution: laborArr.sort(sortByValueDesc),
    
    // 從原始資料補算的
    memberTypeDistribution: objToArray(memberTypeStats).sort(sortByValueDesc),
    ageDistribution: objToArray(ageStats), // 年齡通常按區間排序，不按數量排序

    // 趨勢圖
    joinLeaveByMonth: joinLeaveByMonth.slice(-24), // 取近24個月
    joinLeaveByYear: joinLeaveByYear
  };
}


/* --- 輔助函式 --- */

function objToArray(obj) {
  return Object.keys(obj).map(k => ({ key: k, value: obj[k] }));
}

/* --- 會員管理 CRUD (維持原有功能) --- */

function getMembers() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_MEMBERS);
  if (!sheet) return [];
  
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
      if (val instanceof Date) {
        val = Utilities.formatDate(val, timeZone, 'yyyy-MM-dd');
      }
      memberObj[h] = val;
    });
    members.push(memberObj);
  }
  return members.reverse();
}

function addMember(formObject) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_MEMBERS);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const newRow = [];
  
  if (!formObject['會籍狀態']) formObject['會籍狀態'] = '在籍';
  if (!formObject['入會日期']) {
    formObject['入會日期'] = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd');
  }

  headers.forEach(header => {
    newRow.push(formObject[header] || '');
  });

  sheet.appendRow(newRow);
  return `成功新增會員：${formObject['姓名']}`;
}

function updateMember(formObject) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_MEMBERS);
  const rowIndex = parseInt(formObject.row);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const currentRow = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
  const updatedRow = [];

  headers.forEach((header, index) => {
    if (formObject.hasOwnProperty(header)) {
      updatedRow.push(formObject[header]);
    } else {
      updatedRow.push(currentRow[index]);
    }
  });

  sheet.getRange(rowIndex, 1, 1, updatedRow.length).setValues([updatedRow]);
  return `成功更新會員資料：${formObject['姓名']}`;
}

function deleteMember(rowIndex) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_MEMBERS);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colMap = {};
  headers.forEach((h, i) => colMap[String(h).trim()] = i);

  const statusIdx = colMap['會籍狀態'];
  const leaveDateIdx = colMap['退會日期'];
  
  if (statusIdx !== undefined) {
     sheet.getRange(rowIndex, statusIdx + 1).setValue('退會');
     if (leaveDateIdx !== undefined) {
         const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd');
         sheet.getRange(rowIndex, leaveDateIdx + 1).setValue(today);
     }
     return "已將該會員標記為「退會」";
  } else {
     sheet.deleteRow(parseInt(rowIndex));
     return "已刪除該筆資料";
  }
}

/* --- 帳務核對功能 (維持原有功能) --- */

function reconcileDues(reportData) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_MEMBERS);
    const dbData = sheet.getDataRange().getValues();
    const headers = dbData[0];
    const colMap = {};
    headers.forEach((h, i) => colMap[String(h).trim()] = i);

    const dbMembers = {};
    const dbKeyField = '員工編號'; 
    const reportKeyField = '員編';   
    
    if (colMap[dbKeyField] === undefined) {
        return { success: false, error: `系統試算表中找不到「${dbKeyField}」欄位。` };
    }

    for (let i = 1; i < dbData.length; i++) {
      const row = dbData[i];
      const key = String(row[colMap[dbKeyField]] || '').trim();
      const name = row[colMap['姓名']];
      const campus = row[colMap['院區']];
      const unit = row[colMap['單位']];
      const statusRaw = String(row[colMap['會籍狀態']] || '').trim();
      const isLeft = statusRaw.includes('退會');

      if (key) {
        dbMembers[key] = {
          id: key,
          name: name,
          campus: campus,
          unit: unit,
          isLeft: isLeft, 
          matched: false 
        };
      }
    }

    const missing = [];       
    const extra = [];         
    const newMembers = [];    
    const specialCases = [];  

    reportData.forEach(row => {
      const reportId = String(row[reportKeyField] || '').trim();
      if(!reportId) return;

      const reportName = row['姓名'] || row['名字'];
      const amount = parseInt(row['金額'] || row['扣款金額'] || 0);
      const reportCampus = row['院區'] || '';

      if (dbMembers[reportId]) {
        dbMembers[reportId].matched = true;
        if (dbMembers[reportId].isLeft) {
          extra.push({ id: reportId, name: reportName, campus: reportCampus, amount: amount });
        }
      } else {
        if (amount === 1450) {
          newMembers.push({ id: reportId, name: reportName, campus: reportCampus, amount: amount });
        } else if (amount === 1700) {
          specialCases.push({ id: reportId, name: reportName, campus: reportCampus, amount: amount });
        } else {
           newMembers.push({ id: reportId, name: reportName, campus: reportCampus, amount: amount });
        }
      }
    });

    for (const key in dbMembers) {
      const member = dbMembers[key];
      if (!member.isLeft && !member.matched) {
        missing.push({
          id: member.id,
          name: member.name,
          campus: member.campus,
          unit: member.unit
        });
      }
    }

    return {
      success: true,
      results: { missing, extra, newMembers, specialCases }
    };

  } catch (e) {
    return { success: false, error: e.message };
  }
}
