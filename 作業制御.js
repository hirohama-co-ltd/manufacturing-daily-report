// ========================================
// 作業マスタ・作業QR
// ========================================

var WORK_MASTER_SHEET_NAME = '作業マスタ';

/**
 * 作業マスタ読み込み（日報用スプレッドシート内）
 * 列: QRコード | 作業区分名 | 表示色 | 表示順
 * 表示色: 製造/修理/調整/未選択、日本語色名（赤・青紫等）、または #hex
 */
function loadWorkMasterFromSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var defaultList = [
    { code: 'WORK_REPAIR', name: '修理', colorKey: '修理', sortOrder: 1 },
    { code: 'WORK_ADJUST', name: '調整', colorKey: '調整', sortOrder: 2 }
  ];
  if (!ss) return defaultList;

  var sheet = ss.getSheetByName(WORK_MASTER_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return defaultList;

  var lastRow = sheet.getLastRow();
  var lastCol = Math.max(sheet.getLastColumn(), 4);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(function(h) { return String(h || '').trim(); });
  var codeCol = findMasterColumnIndex_(headers, ['QRコード', 'QR', 'コード']);
  var nameCol = findMasterColumnIndex_(headers, ['作業区分名', '作業区分', '名称', '名前']);
  var colorCol = findMasterColumnIndex_(headers, ['表示色', '色', 'カラー']);
  var sortCol = findMasterColumnIndex_(headers, ['表示順', '順序', '順']);
  if (codeCol === -1) codeCol = 0;
  if (nameCol === -1) nameCol = 1;
  if (colorCol === -1) colorCol = 2;
  if (sortCol === -1) sortCol = 3;

  var numRows = lastRow - 1;
  if (numRows < 1) return defaultList;
  var rows = sheet.getRange(2, 1, numRows, lastCol).getValues();
  var list = [];
  for (var i = 0; i < rows.length; i++) {
    var code = normalizeServerCode(String(rows[i][codeCol] || ''));
    var name = String(rows[i][nameCol] || '').trim();
    if (!name) continue;
    var colorKey = normalizeWorkColorKey_(rows[i][colorCol], name);
    var sortOrder = parseInt(rows[i][sortCol], 10) || (i + 1);
    list.push({ code: code || ('WORK_' + i), name: name, colorKey: colorKey, sortOrder: sortOrder });
  }
  list.sort(function(a, b) { return a.sortOrder - b.sortOrder; });
  return list.length > 0 ? list : defaultList;
}

/**
 * 作業QR照合
 */
function processWorkQRScan(rawCode) {
  try {
    var code = normalizeServerCode(rawCode);
    if (!code) return { success: false };

    var list = loadWorkMasterFromSheet();
    for (var i = 0; i < list.length; i++) {
      var masterCode = normalizeServerCode(list[i].code);
      if (masterCode === code || code === list[i].name) {
        return {
          success: true,
          workCode: list[i].code,
          workName: list[i].name,
          colorKey: list[i].colorKey
        };
      }
    }
    return { success: false };
  } catch (e) {
    return { success: false, msg: e.toString() };
  }
}
