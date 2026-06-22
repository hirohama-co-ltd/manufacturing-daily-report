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
    var code = normalizeServerCode(rawCode);
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
    if (idCol === -1 || nameCol === -1) return { success: false };

    var wData = workerSheet.getRange(2, 1, lastRow, workerSheet.getLastColumn()).getValues();
    for (var i = 0; i < wData.length; i++) {
      var masterCodeStr = normalizeServerCode(String(wData[i][idCol]));
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
    var sheets = masterSs.getSheets();
    
    var productName = "未登録品種";
    var qtyPerCase = 0;
    var found = false;

    for (var s = 0; s < sheets.length; s++) {
      var sheet = sheets[s];
      var sheetName = sheet.getName();
      if (sheetName.indexOf("不良") !== -1 || sheetName.indexOf("ライン") !== -1) continue;

      var mData = sheet.getDataRange().getValues();
      if (mData.length < 2) continue;

      var idxCode = -1; var idxName = -1; var idxQty = -1;
      var startRow = 1;

      for (var r = 0; r < Math.min(5, mData.length); r++) {
        var isHeader = false;
        for (var c = 0; c < mData[r].length; c++) {
          var valStr = String(mData[r][c]);
          if (valStr.indexOf("コード") !== -1 || valStr.indexOf("CD") !== -1 || valStr.indexOf("名") !== -1 || valStr.indexOf("入数") !== -1) {
            isHeader = true; break;
          }
        }
        if (isHeader) {
          startRow = r + 1;
          for (var c = 0; c < mData[r].length; c++) {
            var h = String(mData[r][c]).replace(/\s/g, "");
            if (h.indexOf("コード") !== -1 || h.indexOf("CD") !== -1) idxCode = c;
            else if (h.indexOf("名") !== -1) idxName = c;
            else if (h.indexOf("入数") !== -1 || h.indexOf("数量") !== -1) idxQty = c;
          }
          break;
        }
      }
      if (idxCode === -1 || idxName === -1) continue;

      for (var i = startRow; i < mData.length; i++) {
        var rawMasterCode = mData[i][idxCode]; if (!rawMasterCode) continue;
        var masterCodeStr = normalizeServerCode(String(rawMasterCode));
        var masterCodeNum = parseInt(masterCodeStr, 10);
        var masterCodePadded = masterCodeStr.padStart(targetCodeStr.length, "0");
        var targetCodePadded = targetCodeStr.padStart(masterCodeStr.length, "0");

        if (masterCodeStr === targetCodeStr || masterCodeStr === targetCodeStr.replace(/^0+/, "") || masterCodePadded === targetCodeStr || targetCodePadded === masterCodeStr || (!isNaN(targetCodeNum) && !isNaN(masterCodeNum) && masterCodeNum === targetCodeNum)) {
          productName = String(mData[i][idxName] || "品名空欄").trim();
          qtyPerCase = (idxQty !== -1) ? (parseInt(String(mData[i][idxQty]).replace(/[^0-9]/g, ""), 10) || 0) : 0;
          found = true; break;
        }
      }
      if (found) break;
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
  return s;
}