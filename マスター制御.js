/**
 * 👥 バーコード総合受付窓口（作業者QRか製品ラベルかを自動判別）
 */
function processGeneralBarcodeScan(rawCode) {
  var normalized = normalizeServerCode(rawCode);
  if (!normalized) {
    return { success: false, msg: "空の値です" };
  }

  // 1. 作業者マスタ
  var workerCheck = processWorkerQRScan(normalized);
  if (workerCheck.success) {
    return { success: true, dataType: "worker", workerName: workerCheck.workerName };
  }

  // 2. 作業マスタ（作業QR）
  var workCheck = processWorkQRScan(normalized);
  if (workCheck.success) {
    workCheck.dataType = "work";
    return workCheck;
  }

  // 3. 製品ラベル（ラベル全体は元文字列から品目CD・ケースNoを切り出す）
  var productCheck = processBarcodeScan(String(rawCode || '').trim() || normalized);
  if (productCheck.success) {
    productCheck.dataType = "product";
    return productCheck;
  }

  return { success: false, msg: "登録がないコードです。" };
}

/**
 * 👥 作業者マスタ（社員マスタ）専用の照合関数
 */
function processWorkerQRScan(rawCode) {
  try {
    var code = normalizeEmployeeCode_(rawCode);
    if (!code || !MASTER_SS_ID) return { success: false };

    var masterSs = SpreadsheetApp.openById(MASTER_SS_ID);
    var workerSheet = masterSs.getSheetByName(EMPLOYEE_MASTER_SHEET_NAME)
      || masterSs.getSheetByName(WORKER_MASTER_LEGACY_SHEET)
      || masterSs.getSheetByName('作業者');
    if (!workerSheet) return { success: false };

    var lastRow = workerSheet.getLastRow();
    if (lastRow < 2) return { success: false };

    var headers = workerSheet.getRange(1, 1, 1, workerSheet.getLastColumn()).getValues()[0]
      .map(function(h) { return String(h || '').trim(); });
    var idCol = findMasterColumnIndex_(headers, ['社員ID', 'QRコード', 'ID']);
    var nameCol = findMasterColumnIndex_(headers, ['氏名', '作業者名', '名前']);
    var activeCol = findMasterColumnIndex_(headers, ['有効']);
    if (idCol === -1 || nameCol === -1) return { success: false };

    var wData = workerSheet.getRange(2, 1, lastRow, workerSheet.getLastColumn()).getValues();
    for (var i = 0; i < wData.length; i++) {
      if (activeCol >= 0) {
        var active = String(wData[i][activeCol] || '').trim();
        if (active && active !== '有効' && active !== 'TRUE' && active !== '1') continue;
      }
      var masterCodeStr = normalizeEmployeeCode_(String(wData[i][idCol]));
      if (masterCodeStr === code) {
        return { success: true, workerName: String(wData[i][nameCol]).trim() };
      }
    }
    return { success: false };
  } catch (e) {
    return { success: false };
  }
}

function findMasterColumnIndex_(headers, aliases) {
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i] || '').replace(/\s/g, '');
    for (var j = 0; j < aliases.length; j++) {
      if (h === aliases[j] || h.indexOf(aliases[j]) !== -1) return i;
    }
  }
  return -1;
}

/** ライン名の表記ゆれを統一（全角ハイフン・余分な空白など） */
function normalizeLineKey_(raw) {
  return String(raw || '').trim()
    .replace(/[\uFF0D\u2212\u2010\u2011\u2013\u2014\uFE58\uFE63]/g, '-')
    .replace(/\u3000/g, ' ')
    .replace(/\s+/g, ' ');
}

var MASTER_LINE_KEY_ALL = '全';
var MASTER_LINE_KEY_CAP = 'キャップ';
var MASTER_LINE_KEY_KUCHIGANE = '口金';

function isPseudoMasterLineKey_(key) {
  var k = normalizeLineKey_(key);
  return k === MASTER_LINE_KEY_ALL || k === MASTER_LINE_KEY_CAP || k === MASTER_LINE_KEY_KUCHIGANE;
}

/** ラインマスタ B列「グループ」（キャップ / 口金） */
function normalizeLineGroup_(raw) {
  var g = String(raw || '').trim();
  if (!g) return '';
  if (g === 'キャップ' || g === 'ｷｬｯﾌﾟ' || g.toLowerCase() === 'cap') return MASTER_LINE_KEY_CAP;
  if (g === '口金' || g === 'くちがね') return MASTER_LINE_KEY_KUCHIGANE;
  return normalizeLineKey_(g);
}

/**
 * 点検マスタのライン列 → 展開先の実ライン名配列
 * 「全」=全ライン、「キャップ」「口金」=ラインマスタのグループ列
 */
function resolveMasterTargetLines_(lineRaw, lineIndex) {
  var key = normalizeLineKey_(lineRaw);
  if (!key) return [];
  lineIndex = lineIndex || getLineMasterIndex_();
  if (key === MASTER_LINE_KEY_ALL) return (lineIndex.lineList || []).slice();
  if (key === MASTER_LINE_KEY_CAP) return (lineIndex.capLines || []).slice();
  if (key === MASTER_LINE_KEY_KUCHIGANE) return (lineIndex.kuchiganeLines || []).slice();
  return [key];
}

function normalizeEmployeeCode_(raw) {
  var code = normalizeServerCode(raw);
  if (!code) return '';
  if (code.indexOf('EMP') !== 0) code = 'EMP' + code;
  return code;
}

function shouldSkipMasterSheetForProductScan_(sheetName) {
  if (!sheetName) return true;
  if (sheetName.indexOf('不良') !== -1 || sheetName.indexOf('ライン') !== -1) return true;
  if (sheetName === EMPLOYEE_MASTER_SHEET_NAME || sheetName === WORKER_MASTER_LEGACY_SHEET) return true;
  if (sheetName.indexOf('社員') !== -1 || sheetName.indexOf('作業者') !== -1) return true;
  if (sheetName.indexOf('ポータル') !== -1 || sheetName.indexOf('ワークフロー') !== -1) return true;
  return false;
}

var PRODUCT_COLUMN_ALIASES_ = {
  code: ['製品コード', '品目CD', '品目コード'],
  name: ['製品名', '品名'],
  category: ['分類'],
  qty: ['入数', 'ケース入数', '数量']
};

function mapProductMasterColumns_(headers) {
  return {
    code: findMasterColumnIndex_(headers, PRODUCT_COLUMN_ALIASES_.code),
    name: findMasterColumnIndex_(headers, PRODUCT_COLUMN_ALIASES_.name),
    category: findMasterColumnIndex_(headers, PRODUCT_COLUMN_ALIASES_.category),
    qty: findMasterColumnIndex_(headers, PRODUCT_COLUMN_ALIASES_.qty)
  };
}

function lookupProductInMasterSheet_(sheet, targetCodeStr) {
  if (!sheet) return null;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  var lastCol = Math.max(sheet.getLastColumn(), PRODUCT_MASTER_HEADERS.length);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(function(h) { return String(h || '').trim(); });
  var col = mapProductMasterColumns_(headers);
  if (col.code < 0 || col.name < 0) return null;

  var targetCodeNum = parseInt(targetCodeStr, 10);
  var values = sheet.getRange(2, 1, lastRow, lastCol).getValues();
  for (var i = 0; i < values.length; i++) {
    var rawMasterCode = values[i][col.code];
    if (!rawMasterCode) continue;
    var masterCodeStr = normalizeServerCode(String(rawMasterCode));
    var masterCodeNum = parseInt(masterCodeStr, 10);
    var masterCodePadded = masterCodeStr.padStart(targetCodeStr.length, '0');
    var targetCodePadded = targetCodeStr.padStart(masterCodeStr.length, '0');
    if (masterCodeStr === targetCodeStr || masterCodeStr === targetCodeStr.replace(/^0+/, '') ||
        masterCodePadded === targetCodeStr || targetCodePadded === masterCodeStr ||
        (!isNaN(targetCodeNum) && !isNaN(masterCodeNum) && masterCodeNum === targetCodeNum)) {
      return {
        productName: String(values[i][col.name] || '品名空欄').trim(),
        qtyPerCase: col.qty >= 0
          ? (parseInt(String(values[i][col.qty]).replace(/[^0-9]/g, ''), 10) || 0)
          : 0
      };
    }
  }
  return null;
}

/**
 * 📦 商品マスタとの照合
 */
function processBarcodeScan(barcodeStr) {
  try {
    if (!barcodeStr || barcodeStr.trim() === "") {
      return { success: false, msg: "空の値です" };
    }
    var targetCodeStr = parseCleanProductCode(barcodeStr);
    var targetCodeNum = parseInt(targetCodeStr, 10);

    var masterSs = SpreadsheetApp.openById(MASTER_SS_ID);
    var productName = '未登録品種';
    var qtyPerCase = 0;
    var found = false;

    var productSheet = masterSs.getSheetByName(PRODUCT_MASTER_SHEET_NAME);
    var hit = lookupProductInMasterSheet_(productSheet, targetCodeStr);
    if (hit) {
      productName = hit.productName;
      qtyPerCase = hit.qtyPerCase;
      found = true;
    }

    if (!found) {
      var sheets = masterSs.getSheets();
      for (var s = 0; s < sheets.length; s++) {
        var sheet = sheets[s];
        var sheetName = sheet.getName();
        if (sheetName === PRODUCT_MASTER_SHEET_NAME) continue;
        if (shouldSkipMasterSheetForProductScan_(sheetName)) continue;
        hit = lookupProductInMasterSheet_(sheet, targetCodeStr);
        if (hit) {
          productName = hit.productName;
          qtyPerCase = hit.qtyPerCase;
          found = true;
          break;
        }
      }
    }

    if (!found) {
      return { success: false, msg: '製品コード「' + targetCodeStr + '」は製品マスタに登録されていません。' };
    }

    var startCaseNo = extractCaseNoFromLabel(barcodeStr);

    return {
      success: true,
      productCode: targetCodeStr,
      productName: productName,
      qtyPerCase: qtyPerCase,
      startCaseNo: startCaseNo
    };

  } catch (e) {
    return { success: false, msg: "マスタ通信エラー: " + e.toString() };
  }
}

/**
 * ⚠ 不良品QRスキャン・マスタ照合
 */
function processDefectQRScan(rawCode) {
  try {
    var cleanCode = normalizeServerCode(rawCode);
    if (!cleanCode) return { success: false, msg: '空の値です' };
    var masterSs = SpreadsheetApp.openById(MASTER_SS_ID);
    var sheet = masterSs.getSheetByName("不良マスタ") || masterSs.getSheetByName("不良品マスタ");
    if(!sheet) return { success: true, defectName: cleanCode };
    var mData = sheet.getDataRange().getValues();
    for(var i=1; i<mData.length; i++) {
      var mCode = normalizeServerCode(String(mData[i][0]));
      var mName = String(mData[i][1]).trim();
      if(mCode === cleanCode) return { success: true, defectName: mName };
    }
    return { success: true, defectName: cleanCode };
  } catch(e) {
    return { success: false, msg: e.toString() };
  }
}

/**
 * 製品ラベルから品目CDを抽出（設定.gs の位置指定を使用）
 */
function parseCleanProductCode(raw) {
  return extractProductCodeFromLabel(raw);
}

/**
 * スキャン文字列の正規化（制御文字・URL・GS1区切り・全角英数など）
 */
function normalizeServerCode(raw) {
  var s = String(raw || '');
  s = s.replace(/^\uFEFF/, '');
  s = s.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '');
  s = s.replace(/\u001D/g, '');
  s = s.replace(/\s+/g, '');
  s = s.trim();
  if (!s) return '';

  if (/^https?:\/\//i.test(s)) {
    var textIdx = s.indexOf('text=');
    if (textIdx >= 0) {
      var tail = s.substring(textIdx + 5);
      var amp = tail.indexOf('&');
      if (amp >= 0) tail = tail.substring(0, amp);
      try {
        s = decodeURIComponent(tail);
      } catch (e) {
        s = tail;
      }
    }
  }

  s = s.replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(ch) {
    return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0);
  });
  s = s.replace(/\s+/g, '').trim();

  if (s.indexOf('.') !== -1) {
    var p = s.split('.');
    s = p[0].trim();
  }
  return s.toUpperCase();
}