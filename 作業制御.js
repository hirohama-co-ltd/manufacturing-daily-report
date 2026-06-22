// ========================================
// 作業マスタ・作業QR
// ========================================

var WORK_MASTER_SHEET_NAME = '作業マスタ';

/**
 * 作業マスタ読み込み（日報用スプレッドシート内）
 * 列: QRコード | 作業区分名 | 表示色 | 表示順
 * 表示色: 修理 / 調整 / 製造 / 未選択 または #hex（チャート用クラスへ変換）
 */
function loadWorkMasterFromSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cacheKey = 'workMaster_' + ss.getId();
  var cached = getCachedJson_(cacheKey);
  if (cached && cached.length) return cached;

  var sheet = ss.getSheetByName(WORK_MASTER_SHEET_NAME);
  var defaultList = [
    { code: 'WORK_REPAIR', name: '修理', colorKey: '修理', sortOrder: 1 },
    { code: 'WORK_ADJUST', name: '調整', colorKey: '調整', sortOrder: 2 }
  ];
  if (!sheet || sheet.getLastRow() < 2) return defaultList;

  var rows = sheet.getRange(2, 1, sheet.getLastRow(), 4).getValues();
  var list = [];
  for (var i = 0; i < rows.length; i++) {
    var code = normalizeServerCode(String(rows[i][0] || ''));
    var name = String(rows[i][1] || '').trim();
    if (!name) continue;
    var colorKey = String(rows[i][2] || name).trim() || name;
    var sortOrder = parseInt(rows[i][3], 10) || (i + 1);
    list.push({ code: code || ('WORK_' + i), name: name, colorKey: colorKey, sortOrder: sortOrder });
  }
  list.sort(function(a, b) { return a.sortOrder - b.sortOrder; });
  var result = list.length > 0 ? list : defaultList;
  putCachedJson_(cacheKey, result);
  return result;
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
