/**
 * ====================================================================
 * PARTNER RELATION SUPPORT SYSTEM (PRS)
 * Backend Logic (Google Apps Script)
 * Version: Full Cleaned (No Duplicates) + Account Verify + Fast Report
 * ====================================================================
 */

const SPREADSHEET_ID = ""; // *** กรุณาใส่ ID ของ Google Sheet ที่นี่ ***

// Sheet Names Configuration
const SHEET_NAME = "Database_Master";
const ANNUAL_SHEET_NAME = "Database_Annual";
const ONBOARD_SHEET_NAME = "Database_Onboard"; 
const FIRSTBK_SHEET_NAME = "Database_Firstbk"; 
const CONFIG_SHEET_NAME = "Config";
const CONFIG_CHECKLIST_SHEET_NAME = "Config_Checklist"; 

// ==========================================
// --- 1. CORE & ROUTING ---
// ==========================================

function doGet(e) {
  let page = e.parameter.page || 'home'; 
  let html;
  const user = Session.getActiveUser().getEmail();

  switch(page) {
    case 'home': html = HtmlService.createTemplateFromFile('Home'); break;
    case 'onboard': html = HtmlService.createTemplateFromFile('Onboard'); break;
    case 'firstbk': html = HtmlService.createTemplateFromFile('Firstbk'); break;
    
    // 🌟 เพิ่มเงื่อนไขสำหรับดึงหน้าจอใหม่ที่รวม Master & Annual เข้าด้วยกัน 🌟
    case 'verify_workspace': html = HtmlService.createTemplateFromFile('VerifyWorkspace'); break;
    
    // (เก็บ index กับ year_verify เดิมไว้ก่อนได้ เผื่อพิมพ์ URL เข้าตรงๆ)
    case 'index': html = HtmlService.createTemplateFromFile('index'); break;
    case 'year_verify': html = HtmlService.createTemplateFromFile('YearVerify'); break;
    
    case 'account_verify': html = HtmlService.createTemplateFromFile('AccountVerify'); break;
    case 'fast_report': html = HtmlService.createTemplateFromFile('FastReport'); break;
    case 'config': 
      if (!isUserAdmin(user)) return HtmlService.createHtmlOutput("<h3>Access Denied / คุณไม่มีสิทธิ์เข้าถึงหน้านี้</h3><p>กรุณาติดต่อผู้ดูแลระบบ</p>");
      html = HtmlService.createTemplateFromFile('Config'); 
      break;
    default: html = HtmlService.createTemplateFromFile('Home');
  }

  return html.evaluate()
    .setTitle('Partner Relation (PRS)')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
}

function getSheet(name) {
  let ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  
  if (name === CONFIG_CHECKLIST_SHEET_NAME) {
      if (!sheet) {
        sheet = ss.insertSheet(CONFIG_CHECKLIST_SHEET_NAME);
        sheet.appendRow(["ID", "Label", "Section", "Type", "Options", "Active", "Condition"]);
        const defaults = [
            ['c1_1', '1.1 การกดรับงาน', 'pre', 'select', JSON.stringify(['มีงานขึ้นปกติ', 'มีงานขึ้นแต่น้อย', 'งานไม่ขึ้น']), true, ""],
            ['c1_2', '1.2 การโทรยืนยัน', 'pre', 'select', JSON.stringify(['โทรยืนยันแล้ว', 'ไม่ได้โทรยืนยัน']), true, ""],
            ['c1_3', '1.3 ระยะทาง', 'pre', 'select', JSON.stringify(['สามารถเดินทางได้', 'ค่อนข้างไกล', 'ไกล']), true, ""],
            ['c1_4', '1.4 ปัญหาแอพ', 'pre', 'select', JSON.stringify(['ใช้งานได้ปกติ', 'ใช้งานค่อนข้างยาก', 'ใช้งานไม่ได้เลย']), true, ""],
            ['c1_5', '1.5 อุปกรณ์', 'pre', 'select', JSON.stringify(['อุปกรณ์ครบ', 'ขาดบางอย่าง', 'ไม่มีอุปกรณ์']), true, ""],
            ['c1_6', '1.6 ถ่ายรูป Check-in', 'post', 'select', JSON.stringify(['ผ่านปกติ', 'ไม่สามารถทำได้']), true, ""],
            ['c1_7', 'เข้าใจ AI Feature', 'post', 'select', JSON.stringify(['เข้าใจ', 'ไม่เข้าใจ']), true, ""],
            ['c2_1', '2.1 ประเมินหน้างาน', 'post', 'select', JSON.stringify(['งานตรงปก', 'งานไม่ตรงปก']), true, ""],
            ['c2_2', '2.2 งานรีโนเวท', 'post', 'select', JSON.stringify(['ไม่พบ', 'พบเจอ']), true, ""],
            ['c2_3', '2.3 งานเกินขอบเขต', 'post', 'select', JSON.stringify(['ไม่พบ', 'พบเจอ']), true, ""],
            ['c2_4', 'การจัดการงานเกิน', 'post', 'select', JSON.stringify(['คุยกับลูกค้าเรียบร้อย', 'ไม่พบ/ลูกค้าไม่อยู่']), true, ""],
            ['c2_5', '2.5 สัตว์เลี้ยง', 'post', 'select', JSON.stringify(['ไม่พบ', 'พบเจอ']), true, ""],
            ['c2_6', '2.6 อุปกรณ์หน้างาน', 'post', 'select', JSON.stringify(['ครบ', 'ไม่ครบ', 'ชำรุด']), true, ""],
            ['c2_7', '2.7 ปัญหาที่พบ', 'post', 'select', JSON.stringify(['ไม่พบ', 'พบเจอ']), true, ""],
            ['c2_8', 'ข้อเสนอแนะ', 'post', 'text', '[]', true, ""]
        ];
        sheet.getRange(2, 1, defaults.length, 7).setValues(defaults);
      } else {
        if (sheet.getLastColumn() < 7) sheet.getRange(1, 7).setValue("Condition");
      }
  }

  if (name === FIRSTBK_SHEET_NAME) {
    if (!sheet) {
      sheet = ss.insertSheet(FIRSTBK_SHEET_NAME);
      sheet.appendRow([
        "Onboard ID", "Maid Code", "Name", "Phone", "Center", 
        "Booking Code", "Job ID", "Clean Date", "Accept Date", "Status",
        "Check_1_1", "Check_1_2", "Check_1_3", "Check_1_4", "Check_1_5", 
        "Advice", "Officer", "Timestamp", "ReviewScore", "CustomerComment", "ProblemID", "History",
        "Check_1_6", "Check_1_7", 
        "Check_2_1", "Check_2_2", "Check_2_3", "Check_2_4", "Check_2_5", "Check_2_6", "Check_2_7", "Check_2_8", 
        "WorkHours", "Clean Time", "ExtraData_JSON" 
      ]);
    } else {
        const currentCols = sheet.getMaxColumns();
        if (currentCols < 35) {
            sheet.insertColumnsAfter(currentCols, 35 - currentCols);
        }
    }
  }

  // 🌟 อัปเดตตาราง Config: รองรับคอลัมน์ที่ 9 "Onboard Tags" 🌟
  if (name === CONFIG_SHEET_NAME) {
      if (!sheet) {
        sheet = ss.insertSheet(CONFIG_SHEET_NAME);
        sheet.appendRow(["Admin Emails", "Centers", "Results", "FT Statuses", "Onboard Groups", "Onboard Types", "Master Types", "Onboard Centers", "Onboard Tags"]); 
      } else {
        if (sheet.getLastColumn() < 9) {
           if(sheet.getLastColumn() < 5) sheet.getRange(1, 5).setValue("Onboard Groups");
           if(sheet.getLastColumn() < 6) sheet.getRange(1, 6).setValue("Onboard Types");
           if(sheet.getLastColumn() < 7) sheet.getRange(1, 7).setValue("Master Types");
           if(sheet.getLastColumn() < 8) sheet.getRange(1, 8).setValue("Onboard Centers"); 
           if(sheet.getLastColumn() < 9) sheet.getRange(1, 9).setValue("Onboard Tags"); 
        }
      }
  }
  
  if (name === ONBOARD_SHEET_NAME) {
    if (!sheet) {
      sheet = ss.insertSheet(ONBOARD_SHEET_NAME);
      sheet.appendRow([
        "ID", "Training Date", "Maid Code", "Name", "Group", 
        "Phone", "ID Card", "Type", "Latest Followup", "Date 2 (Unused)", "Date 3 (Unused)", 
        "Open Date", "Call Status", "First Job", "Job ID", 
        "Trainer", "History Data (JSON)", "FastTrack Status", "Center", "Skip FastTrack", "Master Type", "Tags"
      ]);
    } else {
        const currentCols = sheet.getMaxColumns();
        if (currentCols < 22) {
           sheet.insertColumnsAfter(currentCols, 22 - currentCols);
        }
    }
  }

  if (name === ANNUAL_SHEET_NAME) {
    if (!sheet) {
      sheet = ss.insertSheet(ANNUAL_SHEET_NAME);
      sheet.appendRow([
        "ID", "Ref Code", "Name", "Group", "ID Card", 
        "Birth Date", "Phone", "Consent Status", "Amount", "Outstanding", "Deduction Status",
        "Channel", "Status Process", "Result Date", "Result", "Last Followup",
        "Officer Email", "Submit Date", "Note", "Export Status"
      ]);
    } else {
        const currentCols = sheet.getMaxColumns();
        if (currentCols < 20) {
           sheet.insertColumnsAfter(currentCols, 20 - currentCols);
        }
    }
  }

  if (name === SHEET_NAME) {
      if(!sheet) {
        sheet = ss.insertSheet(SHEET_NAME);
        sheet.appendRow(["ID", "Ref Code", "Name", "ID Card", "Phone", "Training Date", "Submit Date", "Officer", "Center", "Result Date", "Result", "FT Status", "Note", "Export Status", "Type"]);
      }
  }
  
  return sheet;
}

// ==========================================
// --- CHECKLIST CONFIG LOGIC ---
// ==========================================

function getChecklistConfig() {
  const sheet = getSheet(CONFIG_CHECKLIST_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  
  return data.map(row => ({
    id: String(row[0]),
    label: String(row[1]),
    section: String(row[2]),
    type: String(row[3]),
    options: row[4] ? JSON.parse(row[4]) : [],
    active: row[5] === true || row[5] === "TRUE",
    condition: row[6] ? JSON.parse(row[6]) : null 
  }));
}

function saveChecklistConfig(config) {
  const sheet = getSheet(CONFIG_CHECKLIST_SHEET_NAME);
  const lock = LockService.getScriptLock();
  try {
     lock.waitLock(5000);
     if (sheet.getLastRow() > 1) {
         sheet.getRange(2, 1, sheet.getLastRow()-1, 7).clearContent(); 
     }
     
     if (config && config.length > 0) {
         const rows = config.map(q => [
             q.id, q.label, q.section, q.type, 
             JSON.stringify(q.options), 
             q.active,
             q.condition ? JSON.stringify(q.condition) : ""
         ]);
         sheet.getRange(2, 1, rows.length, 7).setValues(rows);
     }
     return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
  finally { lock.releaseLock(); }
}

// ==========================================
// --- GENERAL CONFIG LOGIC ---
// ==========================================

function getConfigs() {
  const sheet = getSheet(CONFIG_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  const maxCols = Math.max(sheet.getLastColumn(), 9);
  
  if (lastRow <= 1) return { admins: [], centers: [], results: [], ftStatuses: [], onboardGroups: [], onboardTypes: [], masterTypes: [], onboardCenters: [], onboardTags: [] };
  
  const data = sheet.getRange(2, 1, lastRow - 1, maxCols).getValues();
  const getCol = (r, idx) => String(r[idx] || "").trim();

  return {
    admins: data.map(r => getCol(r, 0).toLowerCase()).filter(s => s !== ""),
    centers: data.map(r => getCol(r, 1)).filter(s => s !== ""),
    results: data.map(r => getCol(r, 2)).filter(s => s !== ""),
    ftStatuses: data.map(r => getCol(r, 3)).filter(s => s !== ""),
    onboardGroups: data.map(r => getCol(r, 4)).filter(s => s !== ""),
    onboardTypes: data.map(r => getCol(r, 5)).filter(s => s !== ""),
    masterTypes: data.map(r => getCol(r, 6)).filter(s => s !== ""),
    onboardCenters: data.map(r => getCol(r, 7)).filter(s => s !== ""),
    onboardTags: data.map(r => getCol(r, 8)).filter(s => s !== "")
  };
}

function isUserAdmin(email) {
  const configs = getConfigs();
  return configs.admins.includes(String(email).trim().toLowerCase());
}

function getClientConfig() {
  try {
    const configs = typeof getConfigsCached === "function" ? getConfigsCached() : getConfigs(); 
    const currentUser = Session.getActiveUser().getEmail() || "Unknown User";
    const adminList = configs.admins || [];
    const isAdminUser = adminList.includes(String(currentUser).trim().toLowerCase());

    return {
      isAdmin: isAdminUser,
      userEmail: currentUser,
      admins: adminList,
      centers: configs.centers || [],
      results: configs.results || [],
      ftStatuses: configs.ftStatuses || [],
      onboardGroups: configs.onboardGroups || [],
      onboardTypes: configs.onboardTypes || [],
      masterTypes: configs.masterTypes || [],
      onboardCenters: configs.onboardCenters || [],
      onboardTags: configs.onboardTags || [] 
    };
    
  } catch (error) {
    return {
      isAdmin: false,
      userEmail: "System Error",
      error: error.toString()
    };
  }
}

function addConfigItem(type, value) {
  if (!value) return { success: false, message: "ค่าว่าง" };
  const sheet = getSheet(CONFIG_SHEET_NAME);
  let colIndex;
  switch(type) {
    case 'admin': colIndex = 1; break;
    case 'center': colIndex = 2; break; 
    case 'result': colIndex = 3; break;
    case 'ftStatus': colIndex = 4; break;
    case 'onboardGroup': colIndex = 5; break;
    case 'onboardType': colIndex = 6; break;
    case 'masterType': colIndex = 7; break;
    case 'onboardCenter': colIndex = 8; break; 
    case 'onboardTag': colIndex = 9; break; // เพิ่มบรรทัดนี้
    default: return { success: false, message: "Invalid type" };
  }

  let targetRow = 2;
  while (targetRow <= sheet.getLastRow() && sheet.getRange(targetRow, colIndex).getValue() !== "") targetRow++;
  sheet.getRange(targetRow, colIndex).setValue(value);
  if (typeof clearConfigCache === "function") clearConfigCache();
  return { success: true };
}

function removeConfigItem(type, value) {
  const sheet = getSheet(CONFIG_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: false };
  
  let colIndex;
  switch(type) {
    case 'admin': colIndex = 1; break;
    case 'center': colIndex = 2; break;
    case 'result': colIndex = 3; break;
    case 'ftStatus': colIndex = 4; break;
    case 'onboardGroup': colIndex = 5; break;
    case 'onboardType': colIndex = 6; break;
    case 'masterType': colIndex = 7; break;
    case 'onboardCenter': colIndex = 8; break;
    case 'onboardTag': colIndex = 9; break; // เพิ่มบรรทัดนี้
    default: return { success: false };
  }

  const range = sheet.getRange(2, colIndex, lastRow - 1, 1);
  const values = range.getValues().flat().map(v => String(v).trim().toLowerCase());
  const index = values.indexOf(String(value).trim().toLowerCase());
  if (index !== -1) {
    sheet.getRange(index + 2, colIndex).clearContent();
    const newRange = sheet.getRange(2, colIndex, lastRow - 1, 1);
    const newValues = newRange.getValues().filter(r => r[0] !== "");
    newRange.clearContent();
    if(newValues.length > 0) sheet.getRange(2, colIndex, newValues.length, 1).setValues(newValues);
  }
  if (typeof clearConfigCache === "function") clearConfigCache();
  return { success: true };
}

function saveConfigOrder(type, newList) {
  const sheet = getSheet(CONFIG_SHEET_NAME);
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    let colIndex;
    switch(type) {
      case 'admin': colIndex = 1; break;
      case 'center': colIndex = 2; break;
      case 'result': colIndex = 3; break;
      case 'ftStatus': colIndex = 4; break;
      case 'onboardGroup': colIndex = 5; break;
      case 'onboardType': colIndex = 6; break;
      case 'masterType': colIndex = 7; break;
      case 'onboardCenter': colIndex = 8; break;
      case 'onboardTag': colIndex = 9; break; // เพิ่มบรรทัดนี้
      default: return { success: false, message: "Invalid type" };
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
       sheet.getRange(2, colIndex, lastRow - 1, 1).clearContent();
    }
    
    if (newList && newList.length > 0) {
       const dataToWrite = newList.map(item => [item]);
       sheet.getRange(2, colIndex, dataToWrite.length, 1).setValues(dataToWrite);
    }
    if (typeof clearConfigCache === "function") clearConfigCache();
    return { success: true };
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// --- 2. MASTER DATA LOGIC (New Verify) ---
// ==========================================

function getInitialData(filterVal) {
  const sheet = getSheet(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  const currentUser = Session.getActiveUser().getEmail();
  let data = [];
  
  if (lastRow > 1) {
    const values = sheet.getRange(2, 1, lastRow - 1, 15).getDisplayValues(); 
    data = values.reduce((acc, row, index) => {
      // Basic validation
      if (!row[1] && !row[2]) return acc; // No Code and No Name
      
      // Date Filter
      if (filterVal && !isDateInMonth(row[6], filterVal)) return acc;

      acc.push({
        rowIndex: index + 2, id: row[0], code: row[1], name: row[2], idCard: row[3],
        phone: row[4], trainingDate: row[5], submitDate: row[6], officer: row[7],
        center: row[8], resultDate: row[9], result: row[10], ftStatus: row[11],
        note: row[12], exportStatus: row[13], type: row[14]
      });
      return acc;
    }, []);
    // Sort by Submit Date Descending
    data.sort((a, b) => parseDateForSort(b.submitDate) - parseDateForSort(a.submitDate));
  }
  const configs = getClientConfig();
  return { 
    currentUser: currentUser, 
    isAdmin: configs.isAdmin, 
    centers: configs.centers, 
    results: configs.results, 
    ftStatuses: configs.ftStatuses, 
    masterTypes: configs.masterTypes,
    data: data 
  };
}

function saveData(formData) {
  const sheet = getSheet(SHEET_NAME);
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); 
    let rowNumber;
    let newId = formData.id;
    let currentExportStatus = ""; 
    const lastRow = sheet.getLastRow();

    if (formData.id) {
      if (lastRow < 2) throw new Error("Database Empty");
      const ids = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues().flat().map(id => String(id).trim());
      const index = ids.indexOf(String(formData.id).trim());
      if (index === -1) throw new Error("ID not found");
      rowNumber = index + 2;
      const currentValues = sheet.getRange(rowNumber, 12, 1, 3).getValues()[0]; 
      if (currentValues[0] !== formData.ftStatus) currentExportStatus = ""; else currentExportStatus = currentValues[2];
    } else {
      let maxId = 0;
      if (lastRow >= 2) {
         const existingIds = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
         existingIds.forEach(id => { let num = Number(id); if (!isNaN(num) && num > maxId) maxId = num; });
      }
      newId = (maxId + 1).toString();
      rowNumber = lastRow + 1;
    }

    const rowData = [
      newId, formData.code, formData.name, "'"+formData.idCard, "'"+formData.phone,
      formatDateForSheet(formData.trainingDate), formatDateForSheet(formData.submitDate), 
      formData.officer, formData.center, formatDateForSheet(formData.resultDate), 
      formData.result, formData.ftStatus, formData.note, currentExportStatus, formData.type
    ];
    sheet.getRange(rowNumber, 1, 1, 15).setValues([rowData]);
    return { success: true, message: "บันทึกข้อมูลเรียบร้อย", item: { ...formData, id: newId, exportStatus: currentExportStatus } };
  } catch (e) { return { success: false, message: e.toString() }; } 
  finally { lock.releaseLock(); }
}

function deleteData(id) {
  const currentUser = Session.getActiveUser().getEmail();
  if (!isUserAdmin(currentUser)) return { success: false, message: "คุณไม่มีสิทธิ์ลบข้อมูล" };
  const sheet = getSheet(SHEET_NAME);
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues().flat();
    const index = ids.indexOf(id);
    if (index === -1) return { success: false, message: "ไม่พบข้อมูล" };
    sheet.deleteRow(index + 2);
    return { success: true, message: "ลบข้อมูลเรียบร้อยแล้ว" };
  } catch (e) { return { success: false, message: e.toString() }; } 
  finally { lock.releaseLock(); }
}

function exportMasterCSV(ftStatus, filterVal, isPreview) {
  const sheet = getSheet(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return isPreview ? [] : { content: "", count: 0 };

  const range = sheet.getRange(2, 1, lastRow - 1, 14);
  const displayValues = range.getDisplayValues();
  
  let csvContent = ""; 
  let previewData = []; 
  let count = 0;
  const timestamp = "Exported " + getDateStr();
  
  let exportStatusValues = sheet.getRange(2, 14, lastRow - 1, 1).getValues();

  for (let i = 0; i < displayValues.length; i++) {
    const row = displayValues[i];
    
    if (row[13] !== "") continue; 
    if (filterVal && !isDateMatchFilter(row[6], filterVal)) continue; 
    if (row[11] !== ftStatus) continue;

    let idCard = row[3].toString().replace(/'/g, "").replace(/[\r\n]+/g, "").trim();
    
    let codeMap = 0;
    if(ftStatus === 'Verified') codeMap = 1;
    else if(ftStatus === 'Pending Result') codeMap = 2; 
    else if(ftStatus === 'Not Verified') codeMap = 3;
    else if(ftStatus === 'In Progress') codeMap = 4;

    if (isPreview) {
        previewData.push({ code: row[1], name: row[2], idCard: idCard, ftStatus: row[11], mappedCode: codeMap });
    } else {
        csvContent += `"${idCard}",${codeMap}\n`;
        count++;
        exportStatusValues[i][0] = timestamp;
    }
  }

  if (isPreview) return previewData;
  if (count > 0) {
      sheet.getRange(2, 14, lastRow - 1, 1).setValues(exportStatusValues);
  }
  
  return { content: "\uFEFF" + csvContent.trim(), count: count, filename: `Master_Export_${ftStatus}_${getDateStr()}.csv` };
}

function getAllProviderOptions() {
  const sheet = getSheet(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, 4).getDisplayValues();
  return data.map(row => ({
    code: row[1], name: row[2], idCard: row[3], searchText: `${row[1]} | ${row[2]} | ${row[3]}` 
  })).filter(item => item.code && item.name);
}

// ==========================================
// --- 4. ANNUAL VERIFICATION LOGIC ---
// ==========================================

function getAnnualData(filterVal) { 
  const sheet = getSheet(ANNUAL_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  const currentUser = Session.getActiveUser().getEmail();
  let data = [];

  if (lastRow > 1) {
    const maxCols = sheet.getLastColumn();
    const colsToRead = maxCols < 20 ? maxCols : 20; 
    const values = sheet.getRange(2, 1, lastRow - 1, colsToRead).getDisplayValues(); 
    
    data = values.reduce((acc, row, index) => {
        if (filterVal) {
            if (row.length > 17) {
                if (!isDateMatchFilter(row[17], filterVal)) return acc;
            }
        }
        acc.push({
            rowIndex: index + 2,
            id: row[0] || "",
            refCode: row[1] || "",
            name: row[2] || "",
            group: row[3] || "",
            idCard: row[4] || "",
            birthDate: row[5] || "",
            phone: row[6] || "",
            consentStatus: row[7] || "",
            amount: row[8] || "",
            outstanding: row[9] || "",
            deductionStatus: row[10] || "",
            channel: row[11] || "",
            statusProcess: row[12] || "",
            resultDate: row[13] || "",
            result: row[14] || "",
            lastFollowup: row[15] || "",
            officerEmail: row[16] || "",
            submitDate: row[17] || "",
            note: row[18] || "",
            exportStatus: row[19] || ""
        });
        return acc;
    }, []);
    data.sort((a, b) => parseDateForSort(b.submitDate) - parseDateForSort(a.submitDate));
  }
  const configs = getClientConfig();
  return { currentUser: currentUser, isAdmin: configs.isAdmin, results: configs.results, ftStatuses: configs.ftStatuses, data: data };
}

function importAnnualData(records, selectedDateStr) {
  const sheet = getSheet(ANNUAL_SHEET_NAME);
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const lastRow = sheet.getLastRow();
    
    const timeNow = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
    const uploader = Session.getActiveUser().getEmail(); 
    const logMessage = `${timeNow} โดย ${uploader}`;
    PropertiesService.getScriptProperties().setProperty('LAST_IMPORT_LOG', logMessage);
    
    let existingMap = new Map(); 
    if (lastRow > 1) {
        const allData = sheet.getRange(2, 1, lastRow - 1, 20).getDisplayValues();
        allData.forEach((row, idx) => {
            const clean = String(row[4]).replace(/'/g, "").trim(); 
            if(clean) {
                existingMap.set(clean, { rowIndex: idx + 2, data: row });
            }
        });
    }

    const year = new Date().getFullYear().toString().substr(-2);
    const prefix = `AN-${year}`;
    let maxSeq = 0;
    if (lastRow > 1) {
       const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
       ids.forEach(id => {
           if(String(id).startsWith(prefix)) {
               const parts = String(id).split('-');
               if(parts.length >= 3) maxSeq = Math.max(maxSeq, parseInt(parts[2]) || 0);
           }
       });
    }

    const newRows = [];
    const updates = []; 
    const addedIds = [];
    const updatedIds = [];
    const submitDateVal = selectedDateStr ? formatDateForSheet(selectedDateStr) : formatDateForSheet(new Date());

    function isDifferent(val1, val2) {
        const s1 = String(val1 || "").replace(/^'/, "").trim();
        const s2 = String(val2 || "").replace(/^'/, "").trim();
        if (s1 === s2) return false;
        const n1 = parseFloat(s1.replace(/,/g, ''));
        const n2 = parseFloat(s2.replace(/,/g, ''));
        if (!isNaN(n1) && !isNaN(n2) && n1 === n2) return false;
        return true;
    }

    records.forEach(rec => {
        const cleanCard = String(rec.idCard).replace(/'/g, "").trim();
        const newUpdateValues = [
            String(rec.refCode || ""), String(rec.name || ""), String(rec.group || ""),
            "'" + cleanCard, formatDateForSheet(rec.birthDate), "'" + String(rec.phone || ""),
            String(rec.consentStatus || ""), String(rec.amount || ""), String(rec.outstanding || ""),
            String(rec.deductionStatus || "")
        ];

        if(existingMap.has(cleanCard)) {
            const existing = existingMap.get(cleanCard);
            const currentValues = existing.data.slice(1, 11);
            let changed = false;
            for(let i=0; i<10; i++) {
                if (isDifferent(currentValues[i], newUpdateValues[i])) {
                    changed = true;
                    break;
                }
            }
            if (changed) {
                updates.push({ row: existing.rowIndex, col: 2, data: [newUpdateValues] });
                updatedIds.push(cleanCard); 
            }
        } else {
            // Insert New
            maxSeq++;
            const newId = `${prefix}-${String(maxSeq).padStart(4, '0')}`;
            const rowDataFull = [
                newId, ...newUpdateValues, 
                rec.channel, 
                "", 
                formatDateForSheet(rec.resultDate), rec.result, formatDateForSheet(rec.lastFollowup),
                rec.officerEmail, submitDateVal, rec.note || "", ""
            ];
            newRows.push(rowDataFull);
            addedIds.push(cleanCard);
        }
    });

    updates.forEach(u => {
        sheet.getRange(u.row, u.col, 1, 10).setValues(u.data);
    });

    if(newRows.length > 0) {
        sheet.getRange(lastRow + 1, 1, newRows.length, 20).setValues(newRows);
    }

    return { success: true, updated: updates.length, added: newRows.length, addedIds: addedIds, updatedIds: updatedIds };

  } catch(e) { return { success: false, message: e.toString() }; }
  finally { lock.releaseLock(); }
}

function getLastImportLog() {
  return PropertiesService.getScriptProperties().getProperty('LAST_IMPORT_LOG') || "";
}

function saveAnnualData(form) {
  const sheet = getSheet(ANNUAL_SHEET_NAME);
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    let rowNumber;
    let newId = form.id;
    const lastRow = sheet.getLastRow();
    let currentExportStatus = "";

    if (form.id) {
       const ids = sheet.getRange(2, 1, lastRow > 1 ? lastRow - 1 : 1, 1).getDisplayValues().flat();
       const index = ids.indexOf(String(form.id));
       if (index === -1) throw new Error("ID not found");
       rowNumber = index + 2;
    } else {
       const year = new Date().getFullYear().toString().substr(-2);
       const prefix = `AN-${year}`;
       let maxSeq = 0;
       if (lastRow > 1) {
           const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
           ids.forEach(id => {
               if(String(id).startsWith(prefix)) {
                   const parts = String(id).split('-');
                   if(parts.length >= 3) maxSeq = Math.max(maxSeq, parseInt(parts[2]) || 0);
               }
           });
       }
       newId = `${prefix}-${String(maxSeq + 1).padStart(4, '0')}`;
       rowNumber = lastRow + 1;
    }

    const rowData = [
        newId, form.refCode, form.name, form.group, "'" + form.idCard,
        formatDateForSheet(form.birthDate), "'" + form.phone,
        form.consentStatus, form.amount, form.outstanding, form.deductionStatus,
        form.channel, form.statusProcess, formatDateForSheet(form.resultDate),
        form.result, formatDateForSheet(form.lastFollowup),
        form.officerEmail, formatDateForSheet(form.submitDate),
        form.note, currentExportStatus
    ];

    sheet.getRange(rowNumber, 1, 1, 20).setValues([rowData]);
    return { success: true, message: "บันทึกข้อมูลเรียบร้อย", item: { ...form, id: newId, exportStatus: currentExportStatus } };
  } catch(e) { return { success: false, message: e.toString() }; }
  finally { lock.releaseLock(); }
}

function updateAnnualNote(id, newNote) {
  const sheet = getSheet(ANNUAL_SHEET_NAME);
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const lastRow = sheet.getLastRow();
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues().flat();
    const index = ids.indexOf(String(id));
    if (index === -1) return { success: false, message: "ไม่พบข้อมูล" };
    sheet.getRange(index + 2, 19).setValue(newNote); 
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
  finally { lock.releaseLock(); }
}

function deleteAnnualData(id) {
    const currentUser = Session.getActiveUser().getEmail();
    if (!isUserAdmin(currentUser)) return { success: false, message: "No Permission" };
    const sheet = getSheet(ANNUAL_SHEET_NAME);
    const lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
        const ids = sheet.getRange(2, 1, sheet.getLastRow(), 1).getDisplayValues().flat();
        const index = ids.indexOf(id);
        if(index !== -1) { sheet.deleteRow(index+2); return {success:true}; }
        return {success:false, message: "Not found"};
    } catch(e) { return { success: false, message: e.toString() }; }
    finally { lock.releaseLock(); }
}

function exportAnnualCSV(groupType, filterVal, isPreview) {
  const sheet = getSheet(ANNUAL_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return isPreview ? [] : { content: "", count: 0 };

  const range = sheet.getRange(2, 1, lastRow - 1, 20);
  const displayValues = range.getDisplayValues();
  
  let csvContent = ""; 
  let previewData = []; 
  let count = 0;
  const timestamp = "Exported " + getDateStr();
  let newStatuses = displayValues.map(row => [row[19]]); 

  for (let i = 0; i < displayValues.length; i++) {
    const row = displayValues[i];
    if (row[19] !== "") continue;
    if (filterVal && !isDateMatchFilter(row[17], filterVal)) continue;

    const statusProcess = String(row[12]).trim();
    if (statusProcess !== "ผลตรวจออกแล้ว") continue;

    let idCard = row[4].toString().replace(/'/g, "").replace(/[\r\n]+/g, "").trim();
    let status = row[12].toString().replace(/[\r\n]+/g, "").trim();
    
    let shouldExport = true; 

    if (shouldExport) {
        if (isPreview) {
            previewData.push({ refCode: row[1], name: row[2], idCard: idCard, result: row[14] || status });
        } else {
            csvContent += `"${idCard}",1\n`;
            count++;
            newStatuses[i][0] = timestamp;
        }
    }
  }

  if (isPreview) return previewData;
  if (count > 0) sheet.getRange(2, 20, lastRow - 1, 1).setValues(newStatuses);
  
  return { content: "\uFEFF" + csvContent.trim(), count: count, filename: `Annual_Export_${filterVal || 'ALL'}_${getDateStr()}.csv` };
}

function exportAnnualReport(filterVal) {
  const sheet = getSheet(ANNUAL_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { content: "", count: 0 };

  const displayValues = sheet.getRange(2, 1, lastRow - 1, 20).getDisplayValues();
  let csvContent = "ลำดับ,ID,รหัส,ชื่อ-นามสกุล,กลุ่ม,เลขบัตร,วันเกิด,เบอร์โทร,ยินยอม,ยอดเงิน,ยอดค้าง,สถานะหัก,ช่องทาง,สถานะ,วันรับผล,ผล,ติดตามล่าสุด,จนท.,วันที่ส่ง,หมายเหตุ\n";
  let count = 0;

  displayValues.sort((a, b) => parseDateForSort(a[17]) - parseDateForSort(b[17]));

  for (let i = 0; i < displayValues.length; i++) {
    const row = displayValues[i];
    if (filterVal && !isDateMatchFilter(row[17], filterVal)) continue;

    count++;
    const rowString = [
        count, row[0], row[1], row[2], row[3], `'${row[4]}`, 
        row[5], `'${row[6]}`, row[7], row[8], row[9], row[10],
        row[11], row[12], row[13], row[14], row[15],
        row[16], row[17], row[18]
    ].map(f => {
        let cleanVal = String(f || "").replace(/"/g, '""').replace(/[\r\n]+/g, " ");
        return `"${cleanVal}"`;
    }).join(",");
    
    csvContent += rowString + "\n";
  }
  return { content: "\uFEFF" + csvContent.trim(), count: count, filename: `Annual_Report_${filterVal || 'ALL'}_${getDateStr()}.csv` };
}

// ==========================================
// --- 5. ONBOARD DATA LOGIC ---
// ==========================================

function getOnboardData(filterVal) {
  const sheet = getSheet(ONBOARD_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  const currentUser = Session.getActiveUser().getEmail();
  let data = [];

  const masterSheet = getSheet(SHEET_NAME);
  const masterLastRow = masterSheet.getLastRow();
  let masterStatusMap = new Map();
  
  if (masterLastRow > 1) {
      const masterData = masterSheet.getRange(2, 1, masterLastRow - 1, 12).getValues();
      masterData.forEach(r => {
          let cleanId = String(r[3]).replace(/'/g, "").trim(); 
          let status = r[11];
          let submitDate = r[6];
          if (cleanId) {
              if (!masterStatusMap.has(cleanId)) {
                  masterStatusMap.set(cleanId, { status: status, date: submitDate });
              } else {
                  let current = masterStatusMap.get(cleanId);
                  let currDate = current.date instanceof Date ? current.date : new Date(0);
                  let newDate = submitDate instanceof Date ? submitDate : new Date(0);
                  if (newDate > currDate) {
                      masterStatusMap.set(cleanId, { status: status, date: submitDate });
                  }
              }
          }
      });
  }

  if (lastRow > 1) {
    const maxCols = sheet.getLastColumn();
    // Read up to Col 22 (Tags)
    const colsToRead = maxCols < 22 ? maxCols : 22;
    const values = sheet.getRange(2, 1, lastRow - 1, colsToRead).getDisplayValues();

    data = values.reduce((acc, row, index) => {
      // row[1] = Training Date
      if (filterVal && !isDateMatchFilter(row[1], filterVal)) return acc;
      
      let history = [];
      try { if (row[16] && row[16].startsWith('[')) history = JSON.parse(row[16]); } catch (e) {}

      let tags = [];
      try { if (row.length > 21 && row[21]) tags = JSON.parse(row[21]); } catch(e) { if(row[21]) tags = [row[21]]; }

      let cleanObId = String(row[6]).replace(/'/g, "").trim(); 
      let lookup = masterStatusMap.get(cleanObId); 
      let realStatus = "";
      let isExpired = false;

      if (lookup) {
          if (lookup.date instanceof Date) {
             const sixMonthsAgo = new Date();
             sixMonthsAgo.setMonth(sixMonthsAgo.getMonth() - 6);
             if (lookup.date < sixMonthsAgo) {
                 isExpired = true;
             }
          }
          if (!isExpired) {
             realStatus = lookup.status;
          }
      }

      if (!lookup && !realStatus && row.length > 17 && row[17]) {
          realStatus = row[17]; 
      }
      
      const centerVal = row.length > 18 ? row[18] : "";
      const skipVal = row.length > 19 ? row[19] : "";
      const masterTypeVal = row.length > 20 ? row[20] : "";

      acc.push({
        rowIndex: index + 2,
        id: row[0],
        trainingDate: convertToStandardDate(row[1]),
        maidCode: row[2],
        name: row[3],
        group: row[4],
        phone: row[5],
        idCard: row[6], 
        type: row[7],
        latestFollowup: convertToStandardDate(row[8]), 
        openDate: convertToStandardDate(row[11]),
        callStatus: row[12],
        firstJob: row[13],
        jobId: row[14],
        trainer: row[15],
        history: history, 
        fastTrackStatus: realStatus, 
        center: centerVal, 
        skipFastTrack: skipVal, 
        masterType: masterTypeVal,
        tags: tags
      });
      return acc;
    }, []);
    data.sort((a, b) => parseDateForSort(b.trainingDate) - parseDateForSort(a.trainingDate));
  }
  
 const configs = getClientConfig();
  return { 
    currentUser: currentUser, 
    isAdmin: configs.isAdmin, 
    groups: configs.onboardGroups, 
    onboardTypes: configs.onboardTypes,
    masterTypes: configs.masterTypes, 
    onboardCenters: configs.onboardCenters, 
    onboardTags: configs.onboardTags, // 🌟 เพิ่มบรรทัดนี้ลงไป เพื่อส่ง Tag ไปให้หน้า Onboard 🌟
    data: data 
  };
}

function saveOnboardData(form) {
  const sheet = getSheet(ONBOARD_SHEET_NAME);
  if (sheet.getLastColumn() < 22) {
      sheet.getRange(1, 22).setValue("Tags");
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    let rowNumber;
    let newId = form.id;
    const lastRow = sheet.getLastRow();

    if (form.id) {
       const ids = sheet.getRange(2, 1, lastRow > 1 ? lastRow - 1 : 1, 1).getDisplayValues().flat();
       const index = ids.indexOf(String(form.id));
       if (index === -1) throw new Error("ID not found");
       rowNumber = index + 2;
    } else {
       const year = new Date().getFullYear().toString().substr(-2);
       const prefix = `OB-${year}`;
       let maxSeq = 0;
       if (lastRow > 1) {
           const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
           ids.forEach(id => {
               if(String(id).startsWith(prefix)) {
                   const parts = String(id).split('-');
                   if(parts.length >= 3) maxSeq = Math.max(maxSeq, parseInt(parts[2]) || 0);
               }
           });
       }
       newId = `${prefix}-${String(maxSeq + 1).padStart(4, '0')}`;
       rowNumber = lastRow + 1;
    }
    
    let currentFastTrackStatus = "";
    if (form.id) {
        currentFastTrackStatus = sheet.getRange(rowNumber, 18).getValue();
    }
    
    // ** NEW LOGIC: Update FT Status from Manual Input **
    if (form.skipFastTrack === true && form.manualFtStatus) {
        currentFastTrackStatus = form.manualFtStatus;
    }

    let historyJson = "";
    let latestFollowupDate = "";
    if (form.history && form.history.length > 0) {
        form.history.sort((a, b) => parseDateForSort(b.date) - parseDateForSort(a.date));
        historyJson = JSON.stringify(form.history);
        latestFollowupDate = formatDateForSheet(form.history[0].date); 
    }
    
    let tagsJson = "";
    if (form.tags && form.tags.length > 0) tagsJson = JSON.stringify(form.tags);

    const rowData = [
        newId, 
        formatDateForSheet(form.trainingDate), 
        form.maidCode, 
        form.name, 
        form.group,
        "'" + form.phone, 
        "'" + form.idCard, 
        form.type,
        latestFollowupDate, 
        "", 
        "", 
        formatDateForSheet(form.openDate),
        form.callStatus ? "✓" : "", 
        form.firstJob ? "✓" : "", 
        form.jobId,
        form.trainer,
        historyJson, 
        currentFastTrackStatus,
        form.center, 
        form.skipFastTrack ? "TRUE" : "", 
        form.masterType,
        tagsJson
    ];

    sheet.getRange(rowNumber, 1, 1, rowData.length).setValues([rowData]); 
    return { success: true, message: "บันทึกข้อมูลเรียบร้อย" };
  } catch(e) { return { success: false, message: e.toString() }; }
  finally { lock.releaseLock(); }
}

function deleteOnboardData(id) {
    const currentUser = Session.getActiveUser().getEmail();
    if (!isUserAdmin(currentUser)) return { success: false, message: "No Permission (Admin only)" };
    
    const sheet = getSheet(ONBOARD_SHEET_NAME);
    const lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
        const ids = sheet.getRange(2, 1, sheet.getLastRow(), 1).getDisplayValues().flat();
        const index = ids.indexOf(id);
        if(index !== -1) { sheet.deleteRow(index+2); return {success:true}; }
        return {success:false, message: "Not found"};
    } catch(e) { return { success: false, message: e.toString() }; }
    finally { lock.releaseLock(); }
}

function sendToFastTrack(onboardId) {
    const obSheet = getSheet(ONBOARD_SHEET_NAME);
    const masterSheet = getSheet(SHEET_NAME);
    const lock = LockService.getScriptLock();
    
    try {
        lock.waitLock(10000);
        const obIds = obSheet.getRange(2, 1, obSheet.getLastRow()-1, 1).getDisplayValues().flat();
        const obIndex = obIds.indexOf(onboardId);
        if (obIndex === -1) throw new Error("Onboard ID not found");
        
        const obRowRange = obSheet.getRange(obIndex + 2, 1, 1, 21); 
        const obData = obRowRange.getValues()[0];
        
        const maidCode = obData[2];
        const name = obData[3];
        const phone = obData[5];
        const idCard = obData[6];
        const trainingDate = obData[1];
        const officer = obData[15]; 
        const center = obData[18];  
        const masterType = obData[20]; 
        
        if (!name || !idCard) throw new Error("กรุณาระบุ ชื่อ และ เลขบัตรประชาชน ก่อนส่งตรวจ");

        const masterLastRow = masterSheet.getLastRow();
        let masterId = "";
        
        if (masterLastRow > 1) {
            const masterCheckData = masterSheet.getRange(2, 1, masterLastRow - 1, 7).getValues();
            const cleanIdCard = String(idCard).replace(/'/g, "").trim();
            
            let lastSubmitDate = null;
            let foundDuplicate = false;

            for (let i = 0; i < masterCheckData.length; i++) {
                let rowId = String(masterCheckData[i][3]).replace(/'/g, "").trim();
                if (rowId === cleanIdCard) {
                    foundDuplicate = true;
                    let rowDate = masterCheckData[i][6];
                    if (rowDate instanceof Date) {
                        if (!lastSubmitDate || rowDate > lastSubmitDate) {
                            lastSubmitDate = rowDate;
                        }
                    }
                }
            }

            if (foundDuplicate && lastSubmitDate) {
                const sixMonthsAgo = new Date();
                sixMonthsAgo.setMonth(sixMonthsAgo.getMonth() - 6);
                if (lastSubmitDate > sixMonthsAgo) {
                      return { success: false, message: `ไม่สามารถส่งตรวจได้: เลขบัตรนี้มีการส่งตรวจแล้วเมื่อ ${formatDateForSheet(lastSubmitDate)} (ยังไม่ครบ 6 เดือน)` };
                }
            }
            
            let maxId = 0;
            const existingIds = masterSheet.getRange(2, 1, masterLastRow - 1, 1).getValues().flat();
            existingIds.forEach(id => { let num = Number(id); if (!isNaN(num) && num > maxId) maxId = num; });
            masterId = (maxId + 1).toString();
        } else {
            masterId = "1";
        }

        const submitDate = new Date(); 
        const masterRow = [
            masterId, maidCode, name, "'" + idCard, "'" + phone, 
            trainingDate, formatDateForSheet(submitDate), officer, center, 
            "", "", "รอตรวจสอบ", "", "", masterType 
        ];

        masterSheet.appendRow(masterRow);
        obSheet.getRange(obIndex + 2, 18).setValue("Sent"); 

        return { success: true, message: "ส่งข้อมูลไปยัง Fast Track เรียบร้อยแล้ว" };

    } catch(e) {
        return { success: false, message: e.toString() };
    } finally {
        lock.releaseLock();
    }
}

function exportOnboardReport(filterVal) {
  const sheet = getSheet(ONBOARD_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { content: "", count: 0 };

  const values = sheet.getRange(2, 1, lastRow - 1, 21).getDisplayValues();
  let csvContent = "วันที่อบรม,ID,ชื่อ-นามสกุล,กลุ่ม,ศูนย์,เบอร์โทร,เลขบัตร ปชช.,สถานะ(Onboard),ประเภท(Master),วันเปิดระบบ,ผู้ดูแล,สถานะ FT,ติดตามล่าสุด\n";
  let count = 0;
  
  values.sort((a, b) => parseDateForSort(b[1]) - parseDateForSort(a[1]));

  for (let i = 0; i < values.length; i++) {
      const row = values[i];
      if (filterVal && !isDateMatchFilter(row[1], filterVal)) continue;
      
      count++;
      
      const rowString = [
          row[1], row[2], row[3], row[4], row[18], `"${row[5]}"`, `"${row[6]}"`, 
          row[7], row[20], row[11], row[15], row[17], row[8]
      ].map(f => `"${String(f || "").replace(/"/g, '""')}"`).join(",");
      
      csvContent += rowString + "\n";
  }
  
  const timestamp = getDateStr();
  return { content: "\uFEFF" + csvContent, count: count, filename: `Onboard_Report_${filterVal || 'ALL'}_${timestamp}.csv` };
}

// ==========================================
// --- 6. FIRST JOB TRACKING LOGIC ---
// ==========================================

// *** UPDATED: Get First Job Tracking Data ***
function getFirstBkData() {
  const onboardSheet = getSheet(ONBOARD_SHEET_NAME);
  const firstBkSheet = getSheet(FIRSTBK_SHEET_NAME);
  const currentUser = Session.getActiveUser().getEmail();
  let data = [];
  
  const lastRow = onboardSheet.getLastRow();
  if (lastRow > 1) {
    const obValues = onboardSheet.getRange(2, 1, lastRow - 1, 21).getDisplayValues();
    
    let fbMap = new Map();
    const fbLastRow = firstBkSheet.getLastRow();
    
    if (fbLastRow > 1) {
        // 🌟 ปรับให้อ่านถึงคอลัมน์ที่ 37 (เพื่อรวม PreAdvice และ PostAdvice)
        const fbData = firstBkSheet.getRange(2, 1, fbLastRow - 1, 37).getDisplayValues(); 
        fbData.forEach(r => {
             let history = [];
             try { if (r[21] && r[21].startsWith('[')) history = JSON.parse(r[21]); } catch(e) {}
             
             let extraData = {};
             try { if (r[34] && r[34].startsWith('{')) extraData = JSON.parse(r[34]); } catch(e) {}

             const baseChecklist = {
                 c1_1: r[10], c1_2: r[11], c1_3: r[12], c1_4: r[13], c1_5: r[14],
                 c1_6: r[22], c1_7: r[23],
                 c2_1: r[24], c2_2: r[25], c2_3: r[26], c2_4: r[27], 
                 c2_5: r[28], c2_6: r[29], c2_7: r[30], c2_8: r[31],
                 advice: r[15], officer: r[16], timestamp: r[17],
                 reviewScore: r[18], customerComment: r[19], problemId: r[20]
             };
             
             const mergedChecklist = { ...baseChecklist, ...extraData };

             fbMap.set(String(r[0]), { 
                 bookingCode: r[5], jobId: r[6], cleanDate: r[7], acceptDate: r[8], status: r[9],
                 checklist: mergedChecklist,
                 history: history,
                 workHours: r[32], cleanTime: r[33],
                 // 🌟 อ่านค่าจากคอลัมน์ที่ 36(AJ) และ 37(AK) (ถ้ายกเลิก/ไม่มีข้อมูล ให้ดึงจาก advice เก่าแทน)
                 preAdvice: r[35] || r[15] || "", 
                 postAdvice: r[36] || ""
             });
        });
    }

    data = obValues.reduce((acc, row) => {
        const statusOnboard = String(row[7]).trim(); 
        const id = String(row[0]);
        
        const fbRecord = fbMap.get(id);
        const bookingCode = fbRecord ? fbRecord.bookingCode : "";
        const jobId = fbRecord ? fbRecord.jobId : "";
        const cleanDateStr = fbRecord ? fbRecord.cleanDate : "";
        const status = fbRecord ? fbRecord.status : ""; 

        let cleanTime = 0;
        if (cleanDateStr) cleanTime = parseDateForSort(cleanDateStr);

        let processStatus = "รอรับงานแรก";
        if (bookingCode) {
            if (status === 'PreCallDone') processStatus = "โทรเยี่ยมเรียบร้อย";
            else if (status === 'Done') processStatus = "จบงานสมบูรณ์";
            else processStatus = "รอโทรเยี่ยม";
        } else {
            processStatus = "รอรับงานแรก";
        }
        
        const isPreCallDone = status === 'PreCallDone' || status === 'Done';
        const isPostCallDone = status === 'Done';

        if (statusOnboard === 'เปิดระบบ' || bookingCode) {
           acc.push({
             id: id, maidCode: row[2], name: row[3], phone: row[5], center: row[18],
             bookingCode: bookingCode, jobId: jobId, cleanDate: cleanDateStr, cleanTimestamp: cleanTime,
             acceptDate: fbRecord ? fbRecord.acceptDate : "", 
             processStatus: processStatus,
             isPreCallDone: isPreCallDone, isPostCallDone: isPostCallDone,
             checklist: fbRecord ? fbRecord.checklist : null,
             history: fbRecord ? fbRecord.history : [],
             workHours: fbRecord ? fbRecord.workHours : "",
             // 🌟 ส่งค่าที่อ่านได้กลับไปหน้าเว็บ
             preAdvice: fbRecord ? fbRecord.preAdvice : "",
             postAdvice: fbRecord ? fbRecord.postAdvice : ""
          });
        }
        
        return acc;
    }, []);
  }

  const configs = getClientConfig();
  return { currentUser: currentUser, isAdmin: configs.isAdmin, onboardCenters: configs.onboardCenters, data: data };
}

function saveFirstJobDetails(form) {
  const firstBkSheet = getSheet(FIRSTBK_SHEET_NAME);
  const onboardSheet = getSheet(ONBOARD_SHEET_NAME);
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    
    const lastRow = firstBkSheet.getLastRow();
    let rowIndex = -1;
    
    if (lastRow > 1) {
        const ids = firstBkSheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues().flat();
        const found = ids.indexOf(String(form.id));
        if (found !== -1) rowIndex = found + 2;
    }
    
    // Parse Date & Time
    let formattedCleanDate = form.cleanDate;
    let dbTime = "";
    if (!formattedCleanDate && form.cleanDate) formattedCleanDate = form.cleanDate; 

    const status = "Assigned";
    const timestamp = formatDateForSheet(new Date());
    const logEntry = { date: timestamp, note: `[Assign Job] จ่ายงานแรก: ${form.bookingCode} (Clean: ${form.cleanDate})`, by: form.officer };

    if (rowIndex !== -1) {
        // อัปเดตข้อมูลเดิมที่มีอยู่แล้ว
        firstBkSheet.getRange(rowIndex, 6).setValue(form.bookingCode);
        firstBkSheet.getRange(rowIndex, 7).setValue(form.jobId);
        firstBkSheet.getRange(rowIndex, 8).setValue(formattedCleanDate);     
        firstBkSheet.getRange(rowIndex, 9).setValue(form.acceptDate); // 🌟 เซฟวันที่กดรับงานลงคอลัมน์ 9 (I)
        firstBkSheet.getRange(rowIndex, 33).setValue(form.workHours); 
        firstBkSheet.getRange(rowIndex, 34).setValue(dbTime);     
        
        const currStatus = firstBkSheet.getRange(rowIndex, 10).getValue();
        if(!currStatus) firstBkSheet.getRange(rowIndex, 10).setValue(status);
        firstBkSheet.getRange(rowIndex, 18).setValue(timestamp); 
        
        const historyCell = firstBkSheet.getRange(rowIndex, 22);
        let history = [];
        try { const val = historyCell.getValue(); if (val && String(val).startsWith('[')) history = JSON.parse(val); } catch(e) {}
        history.unshift(logEntry);
        historyCell.setValue(JSON.stringify(history));
    } else {
        // กรณีเพิ่มข้อมูลใหม่ (ดึงมาจาก Onboard)
        const obLastRow = onboardSheet.getLastRow();
        const obIds = onboardSheet.getRange(2, 1, obLastRow - 1, 1).getDisplayValues().flat();
        const obIndex = obIds.indexOf(String(form.id));
        if (obIndex === -1) throw new Error("ไม่พบข้อมูลใน Onboard");
        
        const obData = onboardSheet.getRange(obIndex + 2, 1, 1, 21).getValues()[0];
        const historyJson = JSON.stringify([logEntry]);

        // 🌟 ปรับวิธีสร้าง Row ใหม่ให้อ่านง่าย และป้องกันคอลัมน์เคลื่อน
        const newRow = Array(35).fill(""); // สร้าง Array ว่างๆ 35 ช่อง (เผื่อไว้จนถึง ExtraData)
        newRow[0] = String(form.id);
        newRow[1] = obData[2];             // รหัสแม่บ้าน
        newRow[2] = obData[3];             // ชื่อ-นามสกุล
        newRow[3] = "'" + obData[5];       // เบอร์โทร
        newRow[4] = obData[18];            // ศูนย์ (Center)
        newRow[5] = form.bookingCode;      // รหัสการจอง
        newRow[6] = form.jobId;            // Job ID
        newRow[7] = formattedCleanDate;    // วันที่ทำงาน
        newRow[8] = form.acceptDate;       // 🌟 วันที่กดรับงาน
        newRow[9] = status;                // สถานะ Assigned
        newRow[16] = form.officer;         // เจ้าหน้าที่
        newRow[17] = timestamp;            // เวลาที่อัปเดต
        newRow[21] = historyJson;          // ประวัติ (Log)
        newRow[32] = form.workHours;       // ชั่วโมงทำงาน
        newRow[33] = dbTime;               // เวลา
        newRow[34] = "{}";                 // Extra JSON

        firstBkSheet.appendRow(newRow);
    }
    
    return { success: true, message: "บันทึกข้อมูลงานเรียบร้อย" };
  } catch(e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function saveFirstJobChecklist(form) {
  const firstBkSheet = getSheet(FIRSTBK_SHEET_NAME);
  const onboardSheet = getSheet(ONBOARD_SHEET_NAME); 
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const lastRow = firstBkSheet.getLastRow();
    const ids = firstBkSheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues().flat();
    const index = ids.indexOf(String(form.id));
    
    if (index === -1) throw new Error("ไม่พบข้อมูลงาน (ต้องบันทึกงานก่อน)");
    const rowNumber = index + 2;

    const historyCell = firstBkSheet.getRange(rowNumber, 22); 
    let history = [];
    try { const val = historyCell.getValue(); if (val && String(val).startsWith('[')) history = JSON.parse(val); } catch(e) {}

    let logTitle = "";
    let isComplete = false;
    let noteContent = "";
    
    // Prepare Extra Data (JSON for all answers)
    const extraData = {...form}; 
    delete extraData.id; delete extraData.officer; delete extraData.type; delete extraData.callResult;

    if (form.type === 'precall') {
        if (form.callResult === 'contacted') {
            logTitle = "[Pre-Call] โทรเยี่ยม (ติดต่อได้)";
            isComplete = true; 
            firstBkSheet.getRange(rowNumber, 10).setValue("PreCallDone");
            
            // Standard Cols 11-15
            const preData = [form.c1_1||"", form.c1_2||"", form.c1_3||"", form.c1_4||"", form.c1_5||""];
            firstBkSheet.getRange(rowNumber, 11, 1, 5).setValues([preData]);
            
            // 🌟 บันทึกคำแนะนำ "ก่อนเริ่มงาน" ลงคอลัมน์ที่ 36 (AJ)
            firstBkSheet.getRange(rowNumber, 36).setValue(form.preAdvice || "");
            
            noteContent = `${logTitle}\nบันทึกข้อมูลเรียบร้อย`;
        } else {
            logTitle = `[Pre-Call] โทรไม่ติด (${form.callResult})`;
            noteContent = logTitle;
            isComplete = false; 
        }
    } else {
        // Post-Call
        logTitle = "[Post-Call] ติดตามหลังจบงาน";
        isComplete = true;
        firstBkSheet.getRange(rowNumber, 10).setValue("Done");
        
        const obLastRow = onboardSheet.getLastRow();
        const obIds = onboardSheet.getRange(2, 1, obLastRow - 1, 1).getDisplayValues().flat();
        const obIndex = obIds.indexOf(String(form.id));
        if (obIndex !== -1) onboardSheet.getRange(obIndex + 2, 14).setValue("✓"); 
        
        noteContent = `${logTitle}\nReview: ${form.reviewScore || '-'}`;
              
        // Standard Cols 23-32
        const postData = [
            form.c1_6||"", form.c1_7||"",
            form.c2_1||"", form.c2_2||"", form.c2_3||"", form.c2_4||"", form.c2_5||"", form.c2_6||"", form.c2_7||"", form.c2_8||""
        ];
        firstBkSheet.getRange(rowNumber, 23, 1, 10).setValues([postData]);
        
        firstBkSheet.getRange(rowNumber, 19).setValue(form.reviewScore || "");
        firstBkSheet.getRange(rowNumber, 20).setValue(form.customerComment || "");
        firstBkSheet.getRange(rowNumber, 21).setValue(form.problemId || "");
        
        // 🌟 บันทึกคำแนะนำ "หลังจบงาน" ลงคอลัมน์ที่ 37 (AK)
        firstBkSheet.getRange(rowNumber, 37).setValue(form.postAdvice || "");
    }
    
    // Save Advice (คอลัมน์ 16 แบบเดิม เผื่อไว้ให้เป็น Fallback) & Officer
    if(form.advice) firstBkSheet.getRange(rowNumber, 16).setValue(form.advice);
    firstBkSheet.getRange(rowNumber, 17).setValue(form.officer);
    firstBkSheet.getRange(rowNumber, 18).setValue(formatDateForSheet(new Date()));
    
    // Save Extra Data (JSON) to Column 35 (AI)
    firstBkSheet.getRange(rowNumber, 35).setValue(JSON.stringify(extraData));

    const logEntry = { date: formatDateForSheet(new Date()), note: noteContent, by: form.officer };
    history.unshift(logEntry);
    historyCell.setValue(JSON.stringify(history));

    return { success: true, message: isComplete ? "บันทึกเรียบร้อย" : "บันทึกสถานะเรียบร้อย" };
  } catch(e) { return { success: false, message: e.toString() }; } 
  finally { lock.releaseLock(); }
}

function returnFirstJob(id, reason, officer, problemId) {
  const sheet = getSheet(FIRSTBK_SHEET_NAME);
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues().flat();
    const index = ids.indexOf(String(id));
    if (index === -1) return { success: false, message: "ไม่พบข้อมูล" };
    
    const row = index + 2;
    const historyCell = sheet.getRange(row, 22);
    let history = [];
    try { const val = historyCell.getValue(); if (val && String(val).startsWith('[')) history = JSON.parse(val); } catch(e) {}
    
    const timestamp = formatDateForSheet(new Date());
    
    let noteText = `[Returned] คืนงานเนื่องจาก: ${reason}`;
    if (problemId) {
        noteText += `<br>Problem ID: <a href="https://admin-test.beneat.co/report-problems/${problemId}/edit" target="_blank">${problemId}</a>`;
    }
    
    history.unshift({ date: timestamp, note: noteText, by: officer });
    
    // Clear Assignment
    sheet.getRange(row, 6, 1, 5).clearContent(); 
    sheet.getRange(row, 33).clearContent();
    sheet.getRange(row, 34).clearContent();
    // Clear Checklist
    sheet.getRange(row, 11, 1, 5).clearContent(); 
    sheet.getRange(row, 23, 1, 10).clearContent();
    sheet.getRange(row, 10).clearContent(); // Clear status too
    
    historyCell.setValue(JSON.stringify(history));
    return { success: true };
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function exportFirstBkReport() {
  const sheet = getSheet(FIRSTBK_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { content: "", count: 0 };

  const values = sheet.getRange(2, 1, lastRow - 1, 34).getDisplayValues();
  let csvContent = "รหัสแม่บ้าน,ชื่อ-นามสกุล,ศูนย์,เบอร์โทร,รหัสการจอง,JobID,วันทำความสะอาด,เวลา,จำนวนชั่วโมง,สถานะ," + 
                   "1.1รับงาน,1.2โทรยืนยัน,1.3ระยะทาง,1.4แอพ,1.5อุปกรณ์,1.6ถ่ายรูป,1.7AI," + 
                   "2.1ประเมิน,2.2รีโนเวท,2.3เกินขอบเขต,2.4จัดการเกิน,2.5สัตว์เลี้ยง,2.6อุปกรณ์หน้างาน,2.7ปัญหา,2.8อื่นๆ," + 
                   "คำแนะนำ,ผู้ติดตาม,รีวิว,คอมเม้น,รหัสปัญหา\n";
  let count = 0;

  for (let i = 0; i < values.length; i++) {
      const row = values[i];
      count++;
      
      const rowString = [
          row[1], row[2], row[4], `"${row[3]}"`, row[5], row[6], row[7], row[33], row[32], 
          row[9],
          row[10], row[11], row[12], row[13], row[14], row[22], row[23], 
          row[24], row[25], row[26], row[27], row[28], row[29], row[30], row[31], 
          row[15], row[16], row[18], row[19], row[20]
      ].map(f => `"${String(f || "").replace(/"/g, '""')}"`).join(",");
      
      csvContent += rowString + "\n";
  }
  
  return { content: "\uFEFF" + csvContent.trim(), count: count, filename: `FirstJob_Report_${getDateStr()}.csv` };
}

// ==========================================
// --- NEW: FAST REPORT LOGIC ---
// ==========================================

function getFastReportData(dateRange, centerFilter) {
    const obSheet = getSheet(ONBOARD_SHEET_NAME);
    const fbSheet = getSheet(FIRSTBK_SHEET_NAME);
    const mSheet = getSheet(SHEET_NAME); 
    const aSheet = getSheet(ANNUAL_SHEET_NAME);
    const config = getClientConfig();

    let result = {
        centers: config.onboardCenters,
        onboard: { total: 0, openSystem: 0, waiting: 0, stopped: 0 },
        // 🌟 เพิ่มตัวแปร jobList สำหรับเก็บ Array รายชื่องานที่จบแล้วส่งไปให้ Popup หน้าบ้าน
        firstJob: { total: 0, waitingJob: 0, waitingCall: 0, callDone: 0, completed: 0, avgRating: 0, jobList: [] },
        newVerify: { total: 0, verified: 0, pending: 0, notVerified: 0 },
        annualVerify: { total: 0, verified: 0, paid: 0, unpaid: 0 }
    };

    // 1. Onboard Stats
    const obData = getOnboardData(dateRange).data; 
    let validMaidIds = new Set();
    
    // Trend Data Preparation
    let trendMap = {}; // Key: YYYY-MM-DD
    
    // Helper to ensure key exists
    const initDay = (dateStr) => {
        if(!trendMap[dateStr]) trendMap[dateStr] = { onboard: 0, firstJob: 0 };
    };

    obData.forEach(row => {
        if (centerFilter !== 'ALL' && row.center !== centerFilter) return;
        
        validMaidIds.add(row.id);
        result.onboard.total++;
        if (row.type === 'เปิดระบบ') result.onboard.openSystem++;
        else if (row.type === 'อยู่ระหว่างเตรียมความพร้อมก่อนเปิดระบบ') result.onboard.waiting++;
        else if (row.type === 'ยุติการอบรม') result.onboard.stopped++;
        
        // Trend
        if (row.trainingDate) {
             initDay(row.trainingDate);
             trendMap[row.trainingDate].onboard++;
        }
    });

    // 2. First Job Stats
    const fbData = getFirstBkData().data; 
    let totalScore = 0;
    let countScore = 0;

    fbData.forEach(row => {
        if (!validMaidIds.has(row.id)) return;

        result.firstJob.total++;
        if (row.processStatus === 'จบงานสมบูรณ์') {
            result.firstJob.completed++;
            
            // 🌟 เพิ่มข้อมูลงานลงใน Array ส่งให้ UI แสดง Popup
            result.firstJob.jobList.push({
                id: row.id,
                name: row.name,
                maidCode: row.maidCode,
                bookingCode: row.bookingCode, // <-- 🌟 เพิ่ม Booking Code
                jobId: row.jobId,             // <-- 🌟 เพิ่ม Job ID สำหรับทำลิ้งก์
                cleanDate: row.cleanDate,
                rating: (row.checklist && row.checklist.reviewScore) ? row.checklist.reviewScore : "NO_REVIEW",
                comment: (row.checklist && row.checklist.customerComment) ? row.checklist.customerComment : "",
                problemId: (row.checklist && row.checklist.problemId) ? row.checklist.problemId : ""
            });
            
        }
        else if (row.processStatus === 'โทรเยี่ยมเรียบร้อย') result.firstJob.callDone++;
        else if (row.processStatus === 'รอโทรเยี่ยม') result.firstJob.waitingCall++;
        else result.firstJob.waitingJob++;
        
        if (row.checklist && row.checklist.reviewScore && row.checklist.reviewScore !== "NO_REVIEW") {
            let score = parseFloat(row.checklist.reviewScore);
            if (!isNaN(score)) {
                totalScore += score;
                countScore++;
            }
        }
        
        // Trend (Use Clean Date)
        let cDate = "";
        if (row.cleanTimestamp) {
            let d = new Date(row.cleanTimestamp);
            let y = d.getFullYear();
            let m = (d.getMonth()+1).toString().padStart(2,'0');
            let dd = d.getDate().toString().padStart(2,'0');
            cDate = `${y}-${m}-${dd}`;
        }
        
        if (cDate) {
            initDay(cDate);
            if (row.processStatus === 'จบงานสมบูรณ์') {
                trendMap[cDate].firstJob++;
            }
        }
    });
    
    if (countScore > 0) result.firstJob.avgRating = (totalScore / countScore).toFixed(1);
    
    // Sort and Format Trend Data
    const sortedDates = Object.keys(trendMap).sort();
    result.trend = {
        labels: sortedDates.map(d => {
             const parts = d.split('-');
             return `${parts[2]}/${parts[1]}`; // DD/MM
        }),
        onboard: sortedDates.map(d => trendMap[d].onboard),
        firstJob: sortedDates.map(d => trendMap[d].firstJob)
    };

    // 3. New Verification Stats
    const mRows = mSheet.getLastRow();
    if (mRows > 1) {
        const mValues = mSheet.getRange(2, 1, mRows - 1, 12).getValues();
        mValues.forEach(row => {
            // FIX: Check dateRange existence
            if (dateRange && !isDateMatchFilter(row[5], dateRange)) return;
            if (centerFilter !== 'ALL' && row[8] !== centerFilter) return;

            result.newVerify.total++;
            const status = row[11];
            if (status === 'Verified') result.newVerify.verified++;
            else if (status === 'Pending Result' || status === 'In Progress') result.newVerify.pending++;
            else if (status === 'Not Verified') result.newVerify.notVerified++;
        });
    }

    // 4. Annual Verification Stats
    const aRows = aSheet.getLastRow();
    if (aRows > 1) {
        const aValues = aSheet.getRange(2, 1, aRows - 1, 20).getValues();
        aValues.forEach(row => {
            // FIX: Check dateRange existence
            if (dateRange && !isDateMatchFilter(row[17], dateRange)) return;
            
            result.annualVerify.total++;
            const process = row[12];
            if (process === 'ผลตรวจออกแล้ว') result.annualVerify.verified++;
            
            const pay = row[10];
            if (pay === 'ชำระครบ' || pay === 'หักสำเร็จ') result.annualVerify.paid++;
            else result.annualVerify.unpaid++;
        });
    }

    return result;
}

// ==========================================
// --- NEW: ACCOUNT VERIFICATION LOGIC ---
// ==========================================

function getAccountVerifyReferenceData() {
    const masterSheet = getSheet(SHEET_NAME);
    const annualSheet = getSheet(ANNUAL_SHEET_NAME);
    const onboardSheet = getSheet(ONBOARD_SHEET_NAME); // NEW: Import Onboard sheet
    const currentUser = Session.getActiveUser().getEmail();
    
    let refData = {};
    
    // 0. Get from Onboard (Start here for latest onboarding status)
    const obRows = onboardSheet.getLastRow();
    if (obRows > 1) {
        // Col ID(0), Name(3), ID Card(6), Type(7)
        // Use getDisplayValues to ensure string format
        const obData = onboardSheet.getRange(2, 1, obRows - 1, 21).getDisplayValues();
        obData.forEach(r => {
             const cleanId = String(r[6]).replace(/'/g, "").replace(/\D/g, "").trim();
             if (cleanId) {
                 refData[cleanId] = {
                     type: 'Onboard (เปิดระบบ)',
                     name: r[3],
                     result: '-', // No verification result yet usually
                     status: r[7] // Onboard Status (e.g. "เปิดระบบ", "รอเอกสาร")
                 };
             }
        });
    }

    // 1. Get from Master (New Verification) - Overwrite if found
    const mRows = masterSheet.getLastRow();
    if (mRows > 1) {
        const mData = masterSheet.getRange(2, 1, mRows - 1, 12).getDisplayValues();
        mData.forEach(r => {
             const cleanId = String(r[3]).replace(/'/g, "").replace(/\D/g, "").trim();
             if (cleanId) {
                 refData[cleanId] = {
                     type: 'ผู้ให้บริการใหม่',
                     name: r[2],
                     result: r[10], // Result
                     status: r[11]  // FT Status
                 };
             }
        });
    }

    // 2. Get from Annual (Yearly Verification) - Overwrite if found (Latest status)
    const aRows = annualSheet.getLastRow();
    if (aRows > 1) {
        const aData = annualSheet.getRange(2, 1, aRows - 1, 15).getDisplayValues();
        aData.forEach(r => {
             const cleanId = String(r[4]).replace(/'/g, "").replace(/\D/g, "").trim();
             if (cleanId) {
                 refData[cleanId] = {
                     type: 'ตรวจประวัติรายปี',
                     name: r[2],
                     result: r[14], // Result
                     status: r[14]  // Status
                 };
             }
        });
    }

    return { referenceData: refData, currentUser: currentUser };
}

// ==========================================
// --- UTILITIES ---
// ==========================================

function parseDateForSort(dateStr) {
  if (!dateStr) return 0;
  dateStr = String(dateStr).trim();
  
  // Support DD/MM/YYYY HH:mm
  if (dateStr.match(/^\d{2}\/\d{2}\/\d{4} \d{2}:\d{2}/)) {
      const [dPart, tPart] = dateStr.split(' ');
      const [d, m, y] = dPart.split('/').map(Number);
      const [hr, min] = tPart.split(':').map(Number);
      let year = y > 2400 ? y - 543 : y;
      return new Date(year, m - 1, d, hr, min).getTime();
  }
  
  if (dateStr.includes(' ')) dateStr = dateStr.split(' ')[0];
  
  // YYYY-MM-DD
  if (dateStr.match(/^\d{4}-\d{2}-\d{2}/)) {
     const parts = dateStr.split('-');
     return new Date(parts[0], parts[1]-1, parts[2]).getTime();
  }
  
  let parts = dateStr.split(/[-/]/);
  if (parts.length === 3) {
      if (parts[0].length === 4) {
           let y = parseInt(parts[0]), m = parseInt(parts[1]) - 1, d = parseInt(parts[2]);
           if (y > 2400) y -= 543;
           return new Date(y, m, d).getTime();
      }
      let d = parseInt(parts[0]), m = parseInt(parts[1]) - 1, y = parseInt(parts[2]);
      if (y > 2400) y -= 543;
      return new Date(y, m, d).getTime();
  }
  return 0;
}

function formatDateForSheet(dateStr) {
  if (!dateStr) return "";
  if (Object.prototype.toString.call(dateStr) === '[object Date]') {
       let d = dateStr.getDate().toString().padStart(2, '0');
       let m = (dateStr.getMonth()+1).toString().padStart(2, '0');
       let y = dateStr.getFullYear();
       return `${d}/${m}/${y > 2400 ? y-543 : y}`;
  }
  if (dateStr.match(/^\d{4}-\d{2}-\d{2}$/)) { const [year, month, day] = dateStr.split('-'); let y = parseInt(year); if (y > 2400) y -= 543; return `${day}/${month}/${y}`; } 
  return dateStr; 
}

function parseYearFromDate(dateStr) {
    if(!dateStr) return "";
    let parts = String(dateStr).split(/[-/]/);
    let y;
    if (parts.length === 3) {
        if(parts[0].length === 4) y = parseInt(parts[0]); else if(parts[2].length === 4) y = parseInt(parts[2]);
    }
    if(y) return y > 2400 ? y : y + 543; 
    return "";
}

function getDateStr() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmm"); }

function isDateMatchFilter(dateStr, filterVal) {
  if (!dateStr || !filterVal) return false;
  dateStr = String(dateStr);
  
  if (dateStr.includes(' ')) {
      dateStr = dateStr.split(' ')[0];
  }

  let separator = null;
  if (filterVal.includes(" to ")) separator = " to ";
  else if (filterVal.includes(" ถึง ")) separator = " ถึง ";
  else if (filterVal.includes(" - ")) separator = " - ";

  if (separator) {
      const [startStr, endStr] = filterVal.split(separator);
      const rowTime = parseDateForSort(dateStr);
      const startTime = parseDateForSort(startStr);
      const endTime = parseDateForSort(endStr) + (24 * 60 * 60 * 1000) - 1; // Include end of day
      
      if (rowTime === 0 || startTime === 0 || endTime === 0) return false;
      return rowTime >= startTime && rowTime <= endTime;
  }
  
  if (filterVal.match(/^\d{4}-\d{2}-\d{2}$/)) {
      const rowTime = parseDateForSort(dateStr);
      const filterTime = parseDateForSort(filterVal);
      if (rowTime === 0) return false; 
      return rowTime === filterTime;
  }
  
  return isDateInMonth(dateStr, filterVal);
}

function isDateInMonth(dateStr, filter) {
  if (!dateStr || !filter) return false;
  dateStr = String(dateStr);
  
  if (dateStr.includes(' ')) {
      dateStr = dateStr.split(' ')[0];
  }

  let filterYear, filterMonth;
  if (filter.indexOf('-') > -1) { [filterYear, filterMonth] = filter.split('-'); } else { filterYear = filter; }
  
  if (dateStr.startsWith(filter)) return true;
  
  let parts = dateStr.split(/[-/]/); 
  if (parts.length === 3) {
      let y, m;
      if (parts[0].length === 4) { y = parts[0]; m = parts[1]; } 
      else if (parts[2].length === 4) { y = parts[2]; m = parts[1]; } 
      
      if (y && m) {
        if (parseInt(y) > 2400) y = (parseInt(y) - 543).toString();
        m = m.toString().padStart(2, '0');
        
        if (filterMonth) return y === filterYear && m === filterMonth; 
        else return y === filterYear;
      }
  }
  return false;
}

function convertToStandardDate(dateStr) {
    if (!dateStr) return "";
    let timestamp = parseDateForSort(dateStr);
    if (timestamp === 0) return dateStr; 
    let d = new Date(timestamp);
    let y = d.getFullYear();
    let m = (d.getMonth() + 1).toString().padStart(2, '0');
    let day = d.getDate().toString().padStart(2, '0');
    return `${y}-${m}-${day}`;
}

// ==========================================
// 🌟 BULK UPDATE & HISTORY LOGS LOGIC 🌟
// ==========================================

const UPLOAD_LOG_SHEET = "UploadLogs";

function processBulkUpdate(records, moduleName, fileName, userEmail) {
  const ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  
  let logSheet = ss.getSheetByName("UploadLogs");
  if (!logSheet) {
    logSheet = ss.insertSheet("UploadLogs");
    logSheet.appendRow(["Timestamp", "User Email", "Module", "File Name", "Success Count", "Failed Count", "Failed Details", "Success Details"]);
  }

  let targetSheet;
  let idCardColIdx; 

  if (moduleName === 'MASTER') {
    targetSheet = ss.getSheetByName(SHEET_NAME); 
    idCardColIdx = 4; // คอลัมน์ D
  } else if (moduleName === 'ANNUAL') {
    targetSheet = ss.getSheetByName(ANNUAL_SHEET_NAME); 
    idCardColIdx = 5; // คอลัมน์ E
  } else {
    return { success: false, message: "Invalid module name" };
  }

  if (!targetSheet) return { success: false, message: "ไม่พบฐานข้อมูล" };

  const lastRow = targetSheet.getLastRow();
  if (lastRow <= 1) return { success: false, message: "ไม่มีข้อมูลในฐานข้อมูลให้ค้นหา" };

  const dataRange = targetSheet.getRange(2, 1, lastRow - 1, targetSheet.getLastColumn());
  const sheetData = dataRange.getDisplayValues();

  // 🌟 ดึงชื่อหัวตารางเพื่อหาคอลัมน์แบบอัตโนมัติ 🌟
  const headers = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
  
  // -- สำหรับหน้า ANNUAL (รายปี) --
  let colAnnualResult = headers.findIndex(h => String(h).includes('ผลตรวจ') || String(h).includes('result')) + 1;
  if (colAnnualResult <= 0 && moduleName === 'ANNUAL') colAnnualResult = 15; // default O
  
  let colAnnualStatusProcess = headers.findIndex(h => String(h).trim().toLowerCase() === 'status process' || String(h).includes('สถานะส่ง')) + 1;
  if (colAnnualStatusProcess <= 0 && moduleName === 'ANNUAL') colAnnualStatusProcess = 13; // default M
  
  let colAnnualResultDate = headers.findIndex(h => String(h).trim().toLowerCase() === 'result date' || String(h).includes('วันรับผล')) + 1;
  if (colAnnualResultDate <= 0 && moduleName === 'ANNUAL') colAnnualResultDate = 14; // default N

  // -- สำหรับหน้า MASTER (ประวัติใหม่) --
  let colMasterFtStatus = headers.findIndex(h => String(h).trim().toLowerCase().includes('สถานะ ft')) + 1;
  if (colMasterFtStatus <= 0 && moduleName === 'MASTER') colMasterFtStatus = 12; // default L
  
  let colMasterResult = headers.findIndex(h => String(h).includes('ผลตรวจสอบ') || String(h) === 'ผลตรวจ') + 1;
  if (colMasterResult <= 0 && moduleName === 'MASTER') colMasterResult = 11; // default K
  
  let colMasterResultDate = headers.findIndex(h => String(h).includes('วันที่รับผล') || String(h).includes('วันรับผล')) + 1;
  if (colMasterResultDate <= 0 && moduleName === 'MASTER') colMasterResultDate = 10; // default J

  let successCount = 0;
  let failedItems = [];
  let successItems = [];
  let updatedIds = []; 
  let updatesToMake = []; 

  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  records.forEach((record, index) => {
    if(!record.idCard && !record.status) return;
    
    const cleanCsvId = String(record.idCard).replace(/\D/g, '');
    const rawStatus = String(record.status).trim();
    const lowerStatus = rawStatus.toLowerCase();
    
    if (cleanCsvId === '' || rawStatus === '' || cleanCsvId.length < 10) {
       if(index > 0) failedItems.push({ idCard: record.idCard, status: rawStatus, reason: "ข้อมูลไม่ครบ/รูปแบบผิด" });
       return;
    }

    let rowIndexFound = -1;
    for (let i = 0; i < sheetData.length; i++) {
       const cleanSheetId = String(sheetData[i][idCardColIdx - 1]).replace(/\D/g, '');
       if (cleanCsvId === cleanSheetId) {
           rowIndexFound = i;
           break;
       }
    }

    if (rowIndexFound !== -1) {
       const targetRow = rowIndexFound + 2;

       if (moduleName === 'MASTER') {
           // 🌟 Logic สำหรับตรวจประวัติใหม่ (MASTER) 🌟
           if (lowerStatus.includes('kyc ไม่ผ่าน') || lowerStatus.includes('kycไม่ผ่าน') || lowerStatus.includes('ไม่ผ่าน')) {
               if (colMasterResult > 0) updatesToMake.push({ row: targetRow, col: colMasterResult, val: 'พบประวัติ' });
               if (colMasterFtStatus > 0) updatesToMake.push({ row: targetRow, col: colMasterFtStatus, val: 'Not Verified' });
               if (colMasterResultDate > 0) updatesToMake.push({ row: targetRow, col: colMasterResultDate, val: todayStr });
           } 
           else if (lowerStatus.includes('ได้รับผลแล้ว') || lowerStatus.includes('อนุมัติโดยพนักงาน')) {
               if (colMasterResult > 0) updatesToMake.push({ row: targetRow, col: colMasterResult, val: 'ไม่พบประวัติ' });
               if (colMasterFtStatus > 0) updatesToMake.push({ row: targetRow, col: colMasterFtStatus, val: 'Verified' });
               if (colMasterResultDate > 0) updatesToMake.push({ row: targetRow, col: colMasterResultDate, val: todayStr });
           } 
           else {
               // กรณีสถานะอื่นๆ ที่นอกเหนือเงื่อนไข ให้อัปเดตแค่ สถานะ FT
               if (colMasterFtStatus > 0) updatesToMake.push({ row: targetRow, col: colMasterFtStatus, val: rawStatus });
           }
       } 
       else if (moduleName === 'ANNUAL') {
           // 🌟 Logic สำหรับตรวจรายปี (ANNUAL) 🌟
           if (colAnnualResult > 0) updatesToMake.push({ row: targetRow, col: colAnnualResult, val: rawStatus });
           
           if (lowerStatus.includes('ได้รับผลแล้ว') || lowerStatus.includes('อนุมัติโดยพนักงาน')) {
               if (colAnnualStatusProcess > 0) updatesToMake.push({ row: targetRow, col: colAnnualStatusProcess, val: 'ผลตรวจออกแล้ว' });
               if (colAnnualResultDate > 0) updatesToMake.push({ row: targetRow, col: colAnnualResultDate, val: todayStr });
           }
       }

       updatedIds.push(String(sheetData[rowIndexFound][0])); 
       successItems.push({ idCard: record.idCard, status: rawStatus });
       successCount++;
    } else {
       failedItems.push({ idCard: record.idCard, status: rawStatus, reason: "ไม่พบเลขบัตรในระบบ" });
    }
  });

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    
    updatesToMake.forEach(update => {
       targetSheet.getRange(update.row, update.col).setValue(update.val);
    });

    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    logSheet.insertRowAfter(1); 
    logSheet.getRange(2, 1, 1, 8).setValues([[
        timestamp, userEmail, moduleName, fileName, successCount, failedItems.length, JSON.stringify(failedItems), JSON.stringify(successItems)
    ]]);
    
    if (typeof clearConfigCache === "function") clearConfigCache();

    return { 
      success: true, 
      successCount: successCount, 
      failedCount: failedItems.length, 
      updatedIds: updatedIds 
    };

  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function getUploadLogs(moduleName) {
  try {
      const ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
      const logSheet = ss.getSheetByName(UPLOAD_LOG_SHEET);
      if (!logSheet) return [];

      const lastRow = logSheet.getLastRow();
      if (lastRow <= 1) return [];

      const logs = logSheet.getRange(2, 1, lastRow - 1, 8).getDisplayValues(); 
      return logs
        .filter(row => row[2] === moduleName)
        .map(row => {
          let failed = [];
          let success = [];
          try { failed = row[6] ? JSON.parse(row[6]) : []; } catch(e) {}
          try { success = row[7] ? JSON.parse(row[7]) : []; } catch(e) {}
          
          return {
            timestamp: row[0],
            user: row[1],
            fileName: row[3],
            successCount: row[4],
            failedCount: row[5],
            failedItems: failed,
            successItems: success
          };
        });
  } catch (error) {
      return [];
  }
}

// ==========================================
// 🌟 UNIFIED BULK UPDATE (Master + Annual) 🌟
// ==========================================
function processUnifiedBulkUpdate(records, fileName, userEmail) {
  const ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  
  let logSheet = ss.getSheetByName("UploadLogs");
  if (!logSheet) {
    logSheet = ss.insertSheet("UploadLogs");
    logSheet.appendRow(["Timestamp", "User Email", "Module", "File Name", "Success Count", "Failed Count", "Failed Details", "Success Details"]);
  }

  // --- 1. เตรียมข้อมูล Master (ประวัติใหม่) ---
  const masterSheet = ss.getSheetByName(SHEET_NAME);
  const masterData = masterSheet ? masterSheet.getDataRange().getDisplayValues() : [];
  const m_headers = masterData.length > 0 ? masterData[0] : [];
  let m_idCol = 4; // D
  let m_ftStatusCol = m_headers.findIndex(h => String(h).trim().toLowerCase().includes('สถานะ ft')) + 1 || 12;
  let m_resultCol = m_headers.findIndex(h => String(h).includes('ผลตรวจสอบ') || String(h) === 'ผลตรวจ') + 1 || 11;
  let m_resultDateCol = m_headers.findIndex(h => String(h).includes('วันที่รับผล') || String(h).includes('วันรับผล')) + 1 || 10;

  // --- 2. เตรียมข้อมูล Annual (รายปี) ---
  const annualSheet = ss.getSheetByName(ANNUAL_SHEET_NAME);
  const annualData = annualSheet ? annualSheet.getDataRange().getDisplayValues() : [];
  const a_headers = annualData.length > 0 ? annualData[0] : [];
  let a_idCol = 5; // E
  let a_resultCol = a_headers.findIndex(h => String(h).includes('ผลตรวจ') || String(h).includes('result')) + 1 || 15;
  let a_statusProcessCol = a_headers.findIndex(h => String(h).trim().toLowerCase() === 'status process' || String(h).includes('สถานะส่ง')) + 1 || 13;
  let a_resultDateCol = a_headers.findIndex(h => String(h).trim().toLowerCase() === 'result date' || String(h).includes('วันรับผล')) + 1 || 14;

  let masterUpdates = [];
  let annualUpdates = [];
  let successItems = [];
  let failedItems = [];
  
  let masterUpdateCount = 0;
  let annualUpdateCount = 0;
  let updatedMasterIds = [];
  let updatedAnnualIds = [];

  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  // --- 3. วนลูปจับคู่ข้อมูล ---
  records.forEach((record, index) => {
    if(!record.idCard && !record.status) return;
    
    const cleanCsvId = String(record.idCard).replace(/\D/g, '');
    const rawStatus = String(record.status).trim();
    const lowerStatus = rawStatus.toLowerCase();
    
    if (cleanCsvId === '' || rawStatus === '' || cleanCsvId.length < 10) {
       if(index > 0) failedItems.push({ idCard: record.idCard, status: rawStatus, reason: "ข้อมูลไม่ครบ" });
       return;
    }

    let foundInAnnual = false;
    let annualRowIdx = -1;

    // 1. ค้นหาใน Annual ก่อนเป็นอันดับแรก
    if (annualData.length > 1) {
        for (let i = 1; i < annualData.length; i++) {
           const sheetId = String(annualData[i][a_idCol - 1]).replace(/\D/g, '');
           if (cleanCsvId === sheetId) { annualRowIdx = i; foundInAnnual = true; break; }
        }
    }

    let foundInMaster = false;
    let masterRowIdx = -1;

    // 2. ค้นหาใน Master เฉพาะกรณีที่ "ไม่เจอใน Annual" เท่านั้น
    if (!foundInAnnual && masterData.length > 1) {
        for (let i = 1; i < masterData.length; i++) {
           const sheetId = String(masterData[i][m_idCol - 1]).replace(/\D/g, '');
           if (cleanCsvId === sheetId) { masterRowIdx = i; foundInMaster = true; break; }
        }
    }

    if (!foundInMaster && !foundInAnnual) {
        failedItems.push({ idCard: record.idCard, status: rawStatus, reason: "ไม่พบในระบบใดเลย" });
        return;
    }

    // 🌟 Logic ฝั่งรายปี (Annual) 🌟
    if (foundInAnnual) {
        const targetRow = annualRowIdx + 1;
        if(a_resultCol > 0) annualUpdates.push({ row: targetRow, col: a_resultCol, val: rawStatus });
        if (lowerStatus.includes('ได้รับผลแล้ว') || lowerStatus.includes('อนุมัติโดยพนักงาน')) {
           if(a_statusProcessCol > 0) annualUpdates.push({ row: targetRow, col: a_statusProcessCol, val: 'ผลตรวจออกแล้ว' });
           if(a_resultDateCol > 0) annualUpdates.push({ row: targetRow, col: a_resultDateCol, val: todayStr });
        }
        updatedAnnualIds.push(String(annualData[annualRowIdx][0]));
        annualUpdateCount++;
    }

    // 🌟 Logic ฝั่งประวัติใหม่ (Master) 🌟
    if (foundInMaster) {
        const targetRow = masterRowIdx + 1;
        if (lowerStatus.includes('kyc ไม่ผ่าน') || lowerStatus.includes('kycไม่ผ่าน') || lowerStatus.includes('ไม่ผ่าน')) {
           if(m_resultCol > 0) masterUpdates.push({ row: targetRow, col: m_resultCol, val: 'พบประวัติ' });
           if(m_ftStatusCol > 0) masterUpdates.push({ row: targetRow, col: m_ftStatusCol, val: 'Not Verified' });
           if(m_resultDateCol > 0) masterUpdates.push({ row: targetRow, col: m_resultDateCol, val: todayStr });
        } else if (lowerStatus.includes('ได้รับผลแล้ว') || lowerStatus.includes('อนุมัติโดยพนักงาน')) {
           if(m_resultCol > 0) masterUpdates.push({ row: targetRow, col: m_resultCol, val: 'ไม่พบประวัติ' });
           if(m_ftStatusCol > 0) masterUpdates.push({ row: targetRow, col: m_ftStatusCol, val: 'Verified' });
           if(m_resultDateCol > 0) masterUpdates.push({ row: targetRow, col: m_resultDateCol, val: todayStr });
        } else {
           if(m_ftStatusCol > 0) masterUpdates.push({ row: targetRow, col: m_ftStatusCol, val: rawStatus });
        }
        updatedMasterIds.push(String(masterData[masterRowIdx][0]));
        masterUpdateCount++;
    }

    successItems.push({ idCard: record.idCard, status: rawStatus, masterMatch: foundInMaster, annualMatch: foundInAnnual });
  });

  // --- 4. บันทึกลง Google Sheets ---
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    
    if (masterSheet && masterUpdates.length > 0) {
        masterUpdates.forEach(u => masterSheet.getRange(u.row, u.col).setValue(u.val));
    }
    if (annualSheet && annualUpdates.length > 0) {
        annualUpdates.forEach(u => annualSheet.getRange(u.row, u.col).setValue(u.val));
    }

    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    logSheet.insertRowAfter(1); 
    logSheet.getRange(2, 1, 1, 8).setValues([[
        timestamp, userEmail, "UNIFIED_VERIFY", fileName, successItems.length, failedItems.length, JSON.stringify(failedItems), JSON.stringify(successItems)
    ]]);
    
    if (typeof clearConfigCache === "function") clearConfigCache();

    return { 
      success: true, 
      masterCount: masterUpdateCount,
      annualCount: annualUpdateCount,
      failedCount: failedItems.length, 
      updatedMasterIds: updatedMasterIds,
      updatedAnnualIds: updatedAnnualIds,
      successItems: successItems,
      failedItems: failedItems
    };

  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 🌟 GLOBAL SMART SEARCH (ค้นหาอัจฉริยะทะลุทุกชีต) 🌟
// ==========================================
function globalSmartSearch(keyword) {
  if (!keyword || keyword.length < 3) return { success: true, data: [] };
  
  keyword = keyword.toLowerCase().trim();
  const ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  
  let results = [];

  // ฟังก์ชันช่วยหาคอลัมน์จากชื่อหัวตาราง
  function getColIndex(headers, possibleNames) {
      return headers.findIndex(h => {
          const lowerH = String(h).toLowerCase().trim();
          return possibleNames.some(pn => lowerH.includes(pn));
      });
  }

  // ฟังก์ชันช่วยค้นหาในแต่ละชีต
  function searchInSheet(sheetName, moduleName) {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return;
      
      const data = sheet.getDataRange().getDisplayValues();
      if (data.length <= 1) return;

      const headers = data[0];
      const nameIdx = getColIndex(headers, ['ชื่อ', 'name']);
      const idIdx = getColIndex(headers, ['บัตร', 'id card', 'เลขบัตร']);
      const codeIdx = getColIndex(headers, ['รหัส', 'code', 'maid', 'ref']);

      for (let i = 1; i < data.length; i++) {
          const row = data[i];
          const name = nameIdx >= 0 ? String(row[nameIdx]) : "";
          const idCard = idIdx >= 0 ? String(row[idIdx]) : "";
          const code = codeIdx >= 0 ? String(row[codeIdx]) : "";
          
          const searchStr = (name + " " + idCard + " " + code).toLowerCase();
          
          // ถ้ามีคำค้นหาอยู่ในแถวนี้
          if (searchStr.includes(keyword)) {
              const cleanId = idCard.replace(/\D/g, '');
              const uniqueKey = cleanId || name; // ใช้เลขบัตรเป็น Key หลักในการรวมข้อมูล
              
              let existing = results.find(r => r.key === uniqueKey);
              if (existing) {
                  // ถ้าเจอคนนี้แล้วในระบบอื่น ให้เพิ่มป้ายสถานะระบบใหม่เข้าไป
                  if (!existing.foundIn.includes(moduleName)) {
                      existing.foundIn.push(moduleName);
                  }
                  if (!existing.code && code) existing.code = code;
              } else {
                  // ถ้าเพิ่งเจอครั้งแรก
                  results.push({
                      key: uniqueKey,
                      name: name || "-",
                      idCard: idCard || "-",
                      code: code || "",
                      foundIn: [moduleName]
                  });
              }
          }
      }
  }

  try {
      // วิ่งค้นหาใน 3 ชีตหลัก (ใช้ตัวแปรชื่อชีตที่คุณตั้งไว้ด้านบนของ Code.gs)
      if (typeof ONBOARD_SHEET_NAME !== 'undefined') searchInSheet(ONBOARD_SHEET_NAME, "Onboard");
      if (typeof SHEET_NAME !== 'undefined') searchInSheet(SHEET_NAME, "Master");
      if (typeof ANNUAL_SHEET_NAME !== 'undefined') searchInSheet(ANNUAL_SHEET_NAME, "Annual");

      // ส่งผลลัพธ์กลับไปแสดงแค่ 15 รายการแรก เพื่อป้องกันหน้าจอกระตุก
      return { success: true, data: results.slice(0, 15) };
  } catch (e) {
      return { success: false, message: e.toString() };
  }
}
