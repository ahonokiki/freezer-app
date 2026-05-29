/* =========================================================
   冷凍在庫システム  Google Apps Script（最新版）
   - 在庫カウントの保存（doPost type=count）
   - 設定（品目マスタ＋基準値）の保存（doPost type=config）
   - 最新在庫の取得（doGet action=latest）
   - 設定の取得（doGet action=config）
   ========================================================= */
const SHEET_NAME = '冷凍庫在庫履歴';
const CONFIG_SHEET = '設定';
const MAX_ITEMS = 50;

function ss(){ return SpreadsheetApp.getActiveSpreadsheet(); }
function mainSheet(){ return ss().getSheetByName(SHEET_NAME) || ss().getSheets()[0]; }
function configSheet(){
  let s = ss().getSheetByName(CONFIG_SHEET);
  if (!s) s = ss().insertSheet(CONFIG_SHEET);
  return s;
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // --- 設定の保存 ---
    if (data.type === 'config') {
      const blob = JSON.stringify({ items: data.items || [], itemSettings: data.itemSettings || {}, ts: data.ts || Date.now() });
      configSheet().getRange(1, 1).setValue(blob);
      return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
    }

    // --- 在庫カウントの保存 ---
    const sheet = mainSheet();
    const row = [
      new Date(),
      data.date || '',
      data.reporter || '',
      data.memo || '',
      data.stock_level || '',
      data.frost_level || '',
    ];
    for (let i = 1; i <= MAX_ITEMS; i++) {
      row.push(data.counts && data.counts[i] !== undefined ? data.counts[i] : '');
    }
    sheet.appendRow(row);
    return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ok:false, error:err.message})).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  // --- 設定の取得 ---
  if (e.parameter.action === 'config') {
    const v = configSheet().getRange(1, 1).getValue();
    if (!v) return ContentService.createTextOutput(JSON.stringify({})).setMimeType(ContentService.MimeType.JSON);
    try {
      return ContentService.createTextOutput(String(v)).setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService.createTextOutput(JSON.stringify({})).setMimeType(ContentService.MimeType.JSON);
    }
  }

  // --- 最新在庫の取得 ---
  if (e.parameter.action === 'latest') {
    const sheet = mainSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow < 1) return ContentService.createTextOutput(JSON.stringify({})).setMimeType(ContentService.MimeType.JSON);
    const lastCol = Math.max(MAX_ITEMS + 6, sheet.getLastColumn());
    const values = sheet.getRange(lastRow, 1, 1, lastCol).getValues()[0];
    const counts = {};
    for (let i = 1; i <= MAX_ITEMS; i++) counts[i] = values[5 + i] ?? '';
    return ContentService.createTextOutput(JSON.stringify({
      ts: values[0],
      date: values[1] instanceof Date ? Utilities.formatDate(values[1], 'JST', 'yyyy-MM-dd') : String(values[1] || ''),
      reporter: values[2] || '',
      memo: values[3] || '',
      stock_level: values[4] || '',
      frost_level: values[5] || '',
      counts: counts,
    })).setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput('OK');
}
