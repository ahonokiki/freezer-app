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

// 値を日付に変換（Dateオブジェクトでも日付文字列でも対応。非日付は null）
function toDate_(v){
  if (v instanceof Date) return v;
  if (typeof v === 'number' && v > 0) { const d = new Date(v); if (!isNaN(d.getTime())) return d; }
  if (typeof v === 'string' && v.trim() !== '') { const d = new Date(v); if (!isNaN(d.getTime())) return d; }
  return null;
}
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
      const blob = JSON.stringify({ items: data.items || [], itemSettings: data.itemSettings || {}, vendors: data.vendors || [], ts: data.ts || Date.now() });
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

  // --- 診断（原因切り分け用） ---
  if (e.parameter.action === 'debug') {
    const sheet = mainSheet();
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    let col0 = '', col0type = '', sample = [];
    if (lastRow >= 1) {
      col0 = sheet.getRange(lastRow, 1).getValue();
      col0type = (col0 instanceof Date) ? 'Date' : (typeof col0);
      sample = sheet.getRange(lastRow, 1, 1, Math.min(8, Math.max(lastCol,1))).getValues()[0].map(String);
    }
    return ContentService.createTextOutput(JSON.stringify({
      sheetName: sheet.getName(),
      allSheets: ss().getSheets().map(s => s.getName()),
      lastRow: lastRow, lastCol: lastCol,
      col0type: col0type, lastRowSample: sample,
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // --- 最新在庫の取得 ---
  // 品目ごとに「最後に値が入った行」を採用する。
  // これにより、今回カウントしなかった品は“前回の確認値”と“その確認日”が残る。
  if (e.parameter.action === 'latest') {
    const sheet = mainSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow < 1) return ContentService.createTextOutput(JSON.stringify({})).setMimeType(ContentService.MimeType.JSON);
    const lastCol = Math.max(MAX_ITEMS + 6, sheet.getLastColumn());
    const all = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const counts = {};
    const dates = {};   // 品目ごとの在庫確認日（yyyy-MM-dd）
    for (let i = 1; i <= MAX_ITEMS; i++) { counts[i] = ''; dates[i] = ''; }
    let newest = null, newestDate = '';  // 最新の在庫送信行（全体ヘッダ表示用）
    for (let r = 0; r < all.length; r++) {
      const rv = all[r];
      // 日付として解釈できる行だけをデータ行とみなす（文字列日付でもOK／ヘッダはスキップ）
      const d = toDate_(rv[1]) || toDate_(rv[0]);
      if (!d) continue;
      const rowDate = Utilities.formatDate(d, 'JST', 'yyyy-MM-dd');
      newest = rv; newestDate = rowDate;
      for (let i = 1; i <= MAX_ITEMS; i++) {
        const cell = rv[5 + i];
        if (cell !== '' && cell !== null && cell !== undefined) {
          counts[i] = cell;     // 後の行ほど新しい → 最後に入った値で上書き
          dates[i]  = rowDate;  // その値が入った日＝最終確認日
        }
      }
    }
    if (!newest) return ContentService.createTextOutput(JSON.stringify({})).setMimeType(ContentService.MimeType.JSON);
    return ContentService.createTextOutput(JSON.stringify({
      ts: newest[0],
      date: newestDate,
      reporter: newest[2] || '',
      memo: newest[3] || '',
      stock_level: newest[4] || '',
      frost_level: newest[5] || '',
      counts: counts,
      dates: dates,
    })).setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput('OK');
}
