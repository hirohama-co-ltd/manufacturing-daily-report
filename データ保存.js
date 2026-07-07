var PROD_HEADERS = ['作業日', '識別No', '品目CD', '品名', 'ケース入数', '始ケース', '終ケース', '良品数量', '状態', '不良内訳', '作業者'];
var TL_HEADERS = ['作業日', 'ID', 'ステータス', '開始時間', '終了時間'];
var CHECK_TYPE_LABELS = { daily: '日常点検', sensor: 'センサチェック', regular: '定時検査', switchover: '切替点検', '設定': '設定' };

/**
 * スプレッドシートへの実績・タイムライン・点検の確定保存
 */
function updateTimelineAndProducts(timelineData, productsData, currentJob, activeLineNo, lineCheckData, workDate) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dateKey = normalizeWorkDate(workDate);

    saveTimelineForDate_(ss, dateKey, timelineData);
    saveProductsForDate_(ss, dateKey, productsData, currentJob);
    saveCheckResultsForDate_(ss, dateKey, lineCheckData);

    return '✅ [ ' + dateKey + ' / ' + activeLineNo + ' ] の製造実績・タイムライン・点検結果を保存しました！';
  } finally {
    lock.releaseLock();
  }
}

function saveTimelineForDate_(ss, workDate, timelineData) {
  var sheet = ss.getSheetByName('タイムライン') || ss.insertSheet('タイムライン');
  var kept = readTimelineRows_(sheet).filter(function(r) { return r.workDate !== workDate; });
  var newRows = (timelineData || []).map(function(b) {
    return {
      workDate: workDate,
      id: b.id,
      status: b.status,
      startTime: b.startTime,
      endTime: b.endTime
    };
  });
  writeTimelineSheet_(sheet, kept.concat(newRows));
}

function saveProductsForDate_(ss, workDate, productsData, currentJob) {
  var sheet = ss.getSheetByName('製造実績') || ss.insertSheet('製造実績');
  var kept = readProductRows_(sheet).filter(function(r) { return r.workDate !== workDate; });
  var newRows = [];

  (productsData || []).forEach(function(p) {
    newRows.push(buildProductRow_(workDate, p, '完了'));
  });
  if (currentJob) {
    newRows.push(buildProductRow_(workDate, currentJob, '製造中'));
  }

  writeProductSheet_(sheet, kept.concat(newRows));
}

function buildProductRow_(workDate, p, status) {
  var defParts = [];
  if (p.defects) {
    Object.keys(p.defects).forEach(function(k) {
      if (p.defects[k] > 0) defParts.push(k + ':' + p.defects[k]);
    });
  }
  var defStr = defParts.length > 0 ? defParts.join(',') : '-';
  return {
    workDate: workDate,
    idName: p.idName,
    productCode: p.productCode,
    productName: p.productName,
    qtyPerCase: p.qtyPerCase || 0,
    startCaseNo: p.startCaseNo,
    endCaseNo: status === '製造中' ? '' : (p.endCaseNo || ''),
    totalQty: status === '製造中' ? 0 : (parseInt(p.totalQty, 10) || 0),
    status: status,
    defects: defStr,
    workers: p.workers || '-'
  };
}

/**
 * 日報集計シートを生成
 */
function generateDailyReportSheet(workDate, activeLineNo) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dateKey = normalizeWorkDate(workDate);
  var sheet = ss.getSheetByName(REPORT_SHEET_NAME) || ss.insertSheet(REPORT_SHEET_NAME);
  sheet.clear();
  sheet.setColumnWidths(1, 6, 120);

  var prodRows = readProductRows_(ss.getSheetByName('製造実績')).filter(function(r) { return r.workDate === dateKey; });
  var tlRows = readTimelineRows_(ss.getSheetByName('タイムライン')).filter(function(r) { return r.workDate === dateKey; });
  var checkRows = getMergedCheckResultRowsForReport_(dateKey, activeLineNo);

  var rows = [];
  rows.push(['製造日報 集計レポート']);
  rows.push(['作業日', dateKey]);
  rows.push(['出力ライン', activeLineNo || '（全ライン）']);
  rows.push(['出力日時', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')]);
  rows.push([]);

  rows.push(['■ 製造実績']);
  rows.push(['識別No', '品目CD', '品名', '始ケース', '終ケース', '良品数量', '状態', '不良内訳', '作業者']);
  if (prodRows.length === 0) {
    rows.push(['（データなし）']);
  } else {
    prodRows.forEach(function(p) {
      rows.push([p.idName, p.productCode, p.productName, p.startCaseNo, p.endCaseNo, p.totalQty, p.status, p.defects, p.workers]);
    });
  }
  rows.push([]);

  rows.push(['■ タイムライン']);
  rows.push(['ステータス', '開始', '終了']);
  if (tlRows.length === 0) {
    rows.push(['（データなし）']);
  } else {
    tlRows.forEach(function(t) {
      rows.push([t.status, t.startTime, t.endTime]);
    });
  }
  rows.push([]);

  rows.push(['■ 点検結果']);
  rows.push(['ライン', '種別', 'スロット', '項目', '判定種別', '結果', '記録時刻']);
  if (checkRows.length === 0) {
    rows.push(['（データなし）']);
  } else {
    checkRows.forEach(function(c) {
      rows.push([c.line, CHECK_TYPE_LABELS[c.type] || c.type, c.slot || '', c.text, c.judgeType, c.value, c.recordedAt || '']);
    });
  }

  if (rows.length > 0) {
    sheet.getRange(1, 1, rows.length, Math.max.apply(null, rows.map(function(r) { return r.length; }))).setValues(
      padRows_(rows)
    );
  }
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setFontSize(12);

  return { success: true, sheetName: REPORT_SHEET_NAME, message: dateKey + ' の集計シート「' + REPORT_SHEET_NAME + '」を作成しました。' };
}

/**
 * 集計シートをPDF化（Base64でクライアントへ返却）
 */
function exportDailyReportPdf(workDate, activeLineNo) {
  var gen = generateDailyReportSheet(workDate, activeLineNo);
  if (!gen.success) return { success: false, msg: '集計シートの作成に失敗しました。' };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(REPORT_SHEET_NAME);
  if (!sheet) return { success: false, msg: '集計シートが見つかりません。' };

  var baseUrl = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?';
  var params = {
    format: 'pdf',
    size: 'A4',
    portrait: 'true',
    fitw: 'true',
    sheetnames: 'false',
    printtitle: 'false',
    pagenumbers: 'false',
    gridlines: 'false',
    fzr: 'false',
    gid: sheet.getSheetId()
  };
  var query = Object.keys(params).map(function(k) { return k + '=' + params[k]; }).join('&');
  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(baseUrl + query, {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    return { success: false, msg: 'PDF生成に失敗しました（HTTP ' + response.getResponseCode() + '）' };
  }

  var dateKey = normalizeWorkDate(workDate);
  var fileName = '製造日報_' + dateKey + (activeLineNo ? '_' + activeLineNo : '') + '.pdf';
  return {
    success: true,
    fileName: fileName,
    base64: Utilities.base64Encode(response.getBlob().getBytes()),
    message: gen.message
  };
}

// --- シート読み書き（内部） ---

function readProductRows_(sheet, filterDateKey) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  var lastCol = Math.max(sheet.getLastColumn(), PROD_HEADERS.length);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h).trim(); });
  var hasDate = headers[0] === '作業日';
  var data = sheet.getRange(2, 1, sheet.getLastRow(), lastCol).getValues();
  var todayKey = normalizeWorkDate(new Date());
  var rows = [];

  for (var i = 0; i < data.length; i++) {
    if (hasDate) {
      if (!data[i][0]) continue;
      var rowDate = normalizeWorkDate(data[i][0]);
      if (filterDateKey && rowDate !== filterDateKey) continue;
      rows.push({
        workDate: rowDate,
        idName: String(data[i][1]),
        productCode: String(data[i][2]),
        productName: String(data[i][3]),
        qtyPerCase: parseInt(data[i][4], 10) || 0,
        startCaseNo: String(data[i][5]),
        endCaseNo: String(data[i][6]),
        totalQty: parseInt(data[i][7], 10) || 0,
        status: String(data[i][8]),
        defects: String(data[i][9] || '-'),
        workers: String(data[i][10] || '-')
      });
    } else {
      if (!data[i][0]) continue;
      rows.push({
        workDate: todayKey,
        idName: String(data[i][0]),
        productCode: String(data[i][1]),
        productName: String(data[i][2]),
        qtyPerCase: parseInt(data[i][3], 10) || 0,
        startCaseNo: String(data[i][4]),
        endCaseNo: String(data[i][5]),
        totalQty: parseInt(data[i][6], 10) || 0,
        status: String(data[i][7]),
        defects: String(data[i][8] || '-'),
        workers: String(data[i][9] || '-')
      });
    }
  }
  return rows;
}

function writeProductSheet_(sheet, rows) {
  sheet.clearContents();
  sheet.getRange(1, 1, 1, PROD_HEADERS.length).setValues([PROD_HEADERS]);
  if (!rows || rows.length === 0) return;
  var values = rows.map(function(r) {
    return [r.workDate, r.idName, r.productCode, r.productName, r.qtyPerCase, r.startCaseNo, r.endCaseNo, r.totalQty, r.status, r.defects, r.workers];
  });
  sheet.getRange(2, 1, values.length, PROD_HEADERS.length).setValues(values);
}

function readTimelineRows_(sheet, filterDateKey) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  var lastCol = Math.max(sheet.getLastColumn(), TL_HEADERS.length);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h).trim(); });
  var hasDate = headers[0] === '作業日';
  var data = sheet.getRange(2, 1, sheet.getLastRow(), lastCol).getValues();
  var todayKey = normalizeWorkDate(new Date());
  var rows = [];

  for (var i = 0; i < data.length; i++) {
    if (hasDate) {
      if (!data[i][0]) continue;
      var rowDate = normalizeWorkDate(data[i][0]);
      if (filterDateKey && rowDate !== filterDateKey) continue;
      rows.push({
        workDate: rowDate,
        id: data[i][1],
        status: String(data[i][2]),
        startTime: formatTime(data[i][3]),
        endTime: formatTime(data[i][4])
      });
    } else {
      rows.push({
        workDate: todayKey,
        id: data[i][0],
        status: String(data[i][1]),
        startTime: formatTime(data[i][2]),
        endTime: formatTime(data[i][3])
      });
    }
  }
  return rows;
}

function writeTimelineSheet_(sheet, rows) {
  sheet.clearContents();
  sheet.getRange(1, 1, 1, TL_HEADERS.length).setValues([TL_HEADERS]);
  if (!rows || rows.length === 0) return;
  var values = rows.map(function(r) {
    return [r.workDate, r.id, r.status, r.startTime, r.endTime];
  });
  sheet.getRange(2, 1, values.length, TL_HEADERS.length).setValues(values);
}

function padRows_(rows) {
  var maxCol = 1;
  rows.forEach(function(r) { if (r.length > maxCol) maxCol = r.length; });
  return rows.map(function(r) {
    var copy = r.slice();
    while (copy.length < maxCol) copy.push('');
    return copy;
  });
}

function parseDefectsFromString_(defectStr) {
  var defectsObj = {};
  var s = String(defectStr || '');
  if (s && s !== '-') {
    s.split(',').forEach(function(p) {
      var kv = p.split(':');
      if (kv.length === 2) defectsObj[kv[0]] = parseInt(kv[1], 10) || 0;
    });
  }
  return defectsObj;
}

/**
 * 指定作業日の製造・タイムライン読込（画面制御から利用）
 */
function loadProductionDataForDate(workDate) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dateKey = normalizeWorkDate(workDate);
  var products = [];
  var currentJob = null;

  readProductRows_(ss.getSheetByName('製造実績'), dateKey).forEach(function(r) {
    var defectsObj = parseDefectsFromString_(r.defects);
    if (r.status === '製造中') {
      currentJob = {
        idName: r.idName,
        productCode: r.productCode,
        productName: r.productName,
        qtyPerCase: r.qtyPerCase,
        startCaseNo: r.startCaseNo,
        endCaseNo: '',
        totalQty: 0,
        defects: defectsObj,
        workers: r.workers
      };
      return;
    }
    products.push({
      idName: r.idName,
      productCode: r.productCode,
      productName: r.productName,
      qtyPerCase: r.qtyPerCase,
      startCaseNo: r.startCaseNo,
      endCaseNo: r.endCaseNo,
      totalQty: r.totalQty,
      defects: defectsObj,
      workers: r.workers
    });
  });

  var timeline = readTimelineRows_(ss.getSheetByName('タイムライン'), dateKey)
    .map(function(r) {
      return { id: r.id || Date.now(), status: r.status, startTime: r.startTime, endTime: r.endTime };
    });

  if (timeline.length === 0) {
    timeline = [{ id: Date.now(), status: '未選択', startTime: '08:30', endTime: '08:45' }];
  } else {
    timeline = normalizeTimelineBlocks(timeline);
  }

  return { products: products, currentJob: currentJob, timeline: timeline };
}

/**
 * 指定作業日より前の、直近作業日の最終製造品（朝の切替点検用）
 */
function loadLastPriorProductForDate(workDate) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('製造実績');
  if (!sheet) return null;
  var dateKey = normalizeWorkDate(workDate);
  var all = readProductRows_(sheet);
  var priorRows = all.filter(function(r) { return r.workDate && r.workDate < dateKey; });
  if (!priorRows.length) return null;

  var maxDate = priorRows[0].workDate;
  for (var i = 1; i < priorRows.length; i++) {
    if (priorRows[i].workDate > maxDate) maxDate = priorRows[i].workDate;
  }
  var onLastDay = priorRows.filter(function(r) { return r.workDate === maxDate; });
  var completed = onLastDay.filter(function(r) { return r.status === '完了'; });
  var pick = completed.length ? completed[completed.length - 1] : onLastDay[onLastDay.length - 1];
  if (!pick) return null;

  return {
    workDate: maxDate,
    productCode: pick.productCode || '',
    productName: pick.productName || '',
    idName: pick.idName || ''
  };
}

/**
 * 共通マスタからライン一覧・グループ情報を読み込む
 * @return {{lineList: string[], capLines: string[], kuchiganeLines: string[]}}
 */
function getLineMasterIndex_() {
  var cacheKey = 'lineMasterIndex_' + (MASTER_SS_ID || 'default');
  var cached = getCachedJson_(cacheKey);
  if (cached && cached.lineList) return cached;

  var defaultLines = ['1号ライン', '2号ライン', '3号ライン'];
  var index = { lineList: defaultLines.slice(), capLines: [], kuchiganeLines: [] };
  if (!MASTER_SS_ID) {
    putCachedJson_(cacheKey, index);
    return index;
  }
  try {
    var masterSs = SpreadsheetApp.openById(MASTER_SS_ID);
    var sheet = masterSs.getSheetByName(LINE_MASTER_SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) {
      putCachedJson_(cacheKey, index);
      return index;
    }
    var lastCol = Math.max(sheet.getLastColumn(), 1);
    var values = sheet.getRange(2, 1, sheet.getLastRow(), lastCol).getValues();
    var lineList = [];
    var capLines = [];
    var kuchiganeLines = [];
    for (var i = 0; i < values.length; i++) {
      var lineName = normalizeLineKey_(values[i][0]);
      if (!lineName || isPseudoMasterLineKey_(lineName)) continue;
      var group = lastCol >= 2 ? normalizeLineGroup_(values[i][1]) : '';
      lineList.push(lineName);
      if (group === MASTER_LINE_KEY_CAP) capLines.push(lineName);
      else if (group === MASTER_LINE_KEY_KUCHIGANE) kuchiganeLines.push(lineName);
    }
    if (lineList.length > 0) {
      index = { lineList: lineList, capLines: capLines, kuchiganeLines: kuchiganeLines };
    }
    putCachedJson_(cacheKey, index);
    return index;
  } catch (e) {
    putCachedJson_(cacheKey, index);
    return index;
  }
}

/**
 * 共通マスタからライン一覧を読み込む
 */
function loadLineMasterFromSheet() {
  return getLineMasterIndex_().lineList;
}

