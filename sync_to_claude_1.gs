// ================================================================
// 張老師專業心算班 — AppSheet 打卡資料自動同步
// 設定說明：
//   1. 開啟您的 Google Sheets
//   2. 點選「擴充功能」→「Apps Script」
//   3. 把這份程式碼全部貼上（取代原本內容）
//   4. 修改下方「請填入您的設定」區塊
//   5. 點選「部署」→「新增部署」→ 類型選「網頁應用程式」
//      執行身分：「我」/ 存取權：「任何人」→ 部署
//   6. 複製產生的網址，填入 Claude 管理系統的設定欄位
//   7. 回到 Apps Script，執行 setupTrigger() 一次（設定自動觸發）
// ================================================================

// ======== 請填入您的設定 ========
const SHEET_NAME = '簽到紀錄';        // Google Sheets 工作表名稱
const CLAUDE_WEBHOOK_URL = '';         // 留空即可（Claude 系統會主動來拉資料）
const SYNC_BATCH_SIZE = 50;            // 每次同步筆數
// =================================

// ================================================================
// 主要功能：提供 API 讓 Claude 系統來抓最新打卡資料
// ================================================================
function doGet(e) {
  try {
    const params = e.parameter;
    const action = params.action || 'latest';
    
    if (action === 'latest') {
      return getLatestCheckins(params);
    } else if (action === 'stats') {
      return getStats();
    } else if (action === 'test') {
      return testConnection();
    }
    
    return jsonResponse({ error: '未知的 action' });
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

// 取得最新打卡記錄
function getLatestCheckins(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    return jsonResponse({ error: `找不到工作表：${SHEET_NAME}` });
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return jsonResponse({ checkins: [], total: 0 });
  }
  
  // 取得最後 N 筆（預設 50 筆）
  const limit = parseInt(params.limit) || SYNC_BATCH_SIZE;
  const startRow = Math.max(2, lastRow - limit + 1);
  const numRows = lastRow - startRow + 1;
  
  // 讀取資料（欄位：日期, 簽到時間, 拍照, 姓名, ID, 簽名, 今日名次）
  const data = sheet.getRange(startRow, 1, numRows, 7).getValues();
  
  const checkins = data
    .filter(row => row[3]) // 姓名不為空
    .map(row => {
      const dateSerial = row[0];
      const timeSerial = row[1];
      
      // 轉換 Excel 日期序號為日期字串
      let dateStr = '';
      let timeStr = '';
      try {
        if (dateSerial instanceof Date) {
          dateStr = Utilities.formatDate(dateSerial, 'Asia/Taipei', 'yyyy-MM-dd');
        } else if (typeof dateSerial === 'number') {
          const d = new Date((dateSerial - 25569) * 86400 * 1000);
          dateStr = Utilities.formatDate(d, 'Asia/Taipei', 'yyyy-MM-dd');
        }
        
        if (timeSerial instanceof Date) {
          timeStr = Utilities.formatDate(timeSerial, 'Asia/Taipei', 'HH:mm');
        } else if (typeof timeSerial === 'number' && timeSerial > 0) {
          const totalMin = Math.round(timeSerial * 24 * 60);
          const h = Math.floor(totalMin / 60) % 24;
          const m = totalMin % 60;
          timeStr = `${String(h).padStart(2,'0')}:${String(m).padStart(2,'0')}`;
        }
      } catch(e) {}
      
      return {
        date: dateStr,
        time: timeStr,
        name: String(row[3]).trim(),
        hasPhoto: !!row[2],
        hasSignature: !!row[5],
        rank: row[6] || ''
      };
    });
  
  return jsonResponse({
    checkins: checkins,
    total: checkins.length,
    sheetLastRow: lastRow,
    syncTime: Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss')
  });
}

// 取得統計資料
function getStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    return jsonResponse({ error: `找不到工作表：${SHEET_NAME}` });
  }
  
  const lastRow = sheet.getLastRow();
  const totalCheckins = Math.max(0, lastRow - 1);
  
  // 讀取所有姓名欄統計
  let uniqueDates = new Set();
  let studentCount = {};
  
  if (totalCheckins > 0) {
    const allData = sheet.getRange(2, 1, totalCheckins, 4).getValues();
    allData.forEach(row => {
      if (row[0]) uniqueDates.add(String(row[0]));
      if (row[3]) {
        const name = String(row[3]).trim();
        studentCount[name] = (studentCount[name] || 0) + 1;
      }
    });
  }
  
  // 最常出席前10名
  const topStudents = Object.entries(studentCount)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 10)
    .map(([name, count]) => ({ name, count }));
  
  return jsonResponse({
    totalCheckins,
    totalDays: uniqueDates.size,
    totalStudents: Object.keys(studentCount).length,
    topStudents,
    syncTime: Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss')
  });
}

// 測試連線
function testConnection() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  return jsonResponse({
    status: 'ok',
    spreadsheetName: ss.getName(),
    sheetFound: !!sheet,
    sheetRows: sheet ? sheet.getLastRow() : 0,
    message: '連線成功！張老師心算班資料同步服務正常運作。'
  });
}

// ================================================================
// 自動觸發：每次有新資料寫入時觸發通知（選用）
// ================================================================
function onSheetEdit(e) {
  // 只在簽到紀錄工作表有異動時處理
  if (e.source.getActiveSheet().getName() !== SHEET_NAME) return;
  
  // 記錄最後同步時間到設定工作表
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName('_同步設定');
  if (!configSheet) {
    configSheet = ss.insertSheet('_同步設定');
    configSheet.getRange('A1').setValue('最後更新時間');
    configSheet.getRange('B1').setValue('今日簽到數');
  }
  configSheet.getRange('A2').setValue(new Date());
  
  // 計算今日簽到數
  const sheet = ss.getSheetByName(SHEET_NAME);
  const today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const dates = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    const todayCount = dates.filter(r => {
      try {
        const d = r[0] instanceof Date ? r[0] : new Date((r[0] - 25569) * 86400 * 1000);
        return Utilities.formatDate(d, 'Asia/Taipei', 'yyyy-MM-dd') === today;
      } catch(e) { return false; }
    }).length;
    configSheet.getRange('B2').setValue(todayCount);
  }
}

// ================================================================
// 設定觸發器（執行一次即可）
// ================================================================
function setupTrigger() {
  // 刪除舊的觸發器
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  
  // 設定工作表編輯觸發器
  ScriptApp.newTrigger('onSheetEdit')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
  
  Logger.log('✅ 觸發器設定完成！每次簽到後會自動更新同步時間。');
  
  // 測試一下
  const result = JSON.parse(testConnection().getContent());
  Logger.log('連線測試：' + JSON.stringify(result));
}

// ================================================================
// 工具函式
// ================================================================
function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data, null, 2))
    .setMimeType(ContentService.MimeType.JSON);
}
