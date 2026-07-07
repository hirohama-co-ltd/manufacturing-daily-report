// ========================================
// 点検・検査（案A: スロット方式）
// ========================================

var CHECK_RESULT_HEADERS = ['作業日', 'ライン', '種別', 'スロット', '項目', '判定種別', '結果', '記録時刻', '品種選択'];
var CHECK_MASTER_DEFAULT_REVISION = '2000-01-01';
var CHECK_MASTER_REV_HEADERS = ['改定日', 'アクション', '旧項目名'];
var SENSOR_SLOTS_BASE = ['午前', '午後'];
var REGULAR_SLOTS_ALL = ['始業', '10時', '13時', '15時', '17時', '終業'];
var HIN_SWITCH_SLOT_PREFIX = '品切';
var SWITCHOVER_SLOT_PREFIX = '切替';

/**
 * クライアント向け定数
 */
function getCheckSlotConstants() {
  return {
    sensorSlotsBase: SENSOR_SLOTS_BASE,
    regularSlots: REGULAR_SLOTS_ALL,
    hinSwitchPrefix: HIN_SWITCH_SLOT_PREFIX,
    switchoverPrefix: SWITCHOVER_SLOT_PREFIX
  };
}

function parseApplySlots_(raw) {
  var s = String(raw || '').trim();
  if (!s) return [];
  return s.split(/[,、\s]+/).map(function(p) { return normalizeSlotLabel_(p.trim()); }).filter(Boolean);
}

function normalizeSlotLabel_(label) {
  var s = String(label || '').trim();
  if (!s) return '';
  if (s === '10' || s === '10:00') return '10時';
  if (s === '13' || s === '13:00') return '13時';
  if (s === '15' || s === '15:00') return '15時';
  if (s === '17' || s === '17:00') return '17時';
  if (s.indexOf('品切') === 0) return s;
  if (s === '品切' || s === '品種切替' || s === '品切替') return HIN_SWITCH_SLOT_PREFIX;
  return s;
}

function hinSwitchSlotName_(n) {
  return HIN_SWITCH_SLOT_PREFIX + String(n);
}

function switchoverSlotName_(n) {
  return SWITCHOVER_SLOT_PREFIX + String(n);
}

function isHinSwitchSlot_(slot) {
  return String(slot || '').indexOf(HIN_SWITCH_SLOT_PREFIX) === 0;
}

function isSwitchoverSlot_(slot) {
  return String(slot || '').indexOf(SWITCHOVER_SLOT_PREFIX) === 0;
}

function masterRowAppliesToSlot_(applySlots, slot, hinSwitchCount) {
  if (!applySlots || applySlots.length === 0) return true;
  for (var i = 0; i < applySlots.length; i++) {
    var a = applySlots[i];
    if (a === slot) return true;
    if (a === HIN_SWITCH_SLOT_PREFIX && isHinSwitchSlot_(slot)) return true;
  }
  return false;
}

function isKnownSensorApplySlot_(slot) {
  if (SENSOR_SLOTS_BASE.indexOf(slot) >= 0) return true;
  if (slot === HIN_SWITCH_SLOT_PREFIX) return true;
  return isHinSwitchSlot_(slot);
}

/** センサチェックマスタの実施区分を解釈（列ずれ・日付誤読時は全スロットにフォールバック） */
function resolveSensorApplySlots_(headers, row) {
  var defaultSlots = SENSOR_SLOTS_BASE.concat([HIN_SWITCH_SLOT_PREFIX]);
  var applyCol = findMasterColumnIndex_(headers, ['実施区分']);
  if (applyCol === -1 && headers.length > 2) {
    var h2 = String(headers[2] || '').replace(/\s/g, '');
    if (h2.indexOf('改定') < 0 && h2.indexOf('アクション') < 0 && h2.indexOf('旧') < 0) {
      applyCol = 2;
    }
  }
  var applySlots = applyCol >= 0 ? parseApplySlots_(row[applyCol]) : [];
  if (!applySlots.length) return defaultSlots.slice();
  var filtered = [];
  applySlots.forEach(function(s) {
    if (isKnownSensorApplySlot_(s) && filtered.indexOf(s) < 0) filtered.push(s);
  });
  return filtered.length ? filtered : defaultSlots.slice();
}

function parseMasterRevisionDate_(cell) {
  if (cell instanceof Date && !isNaN(cell.getTime())) {
    return normalizeWorkDate(cell);
  }
  var s = String(cell != null ? cell : '').trim();
  if (!s) return CHECK_MASTER_DEFAULT_REVISION;
  return normalizeWorkDate(s);
}

function normalizeMasterAction_(raw) {
  var a = String(raw || '').trim();
  if (!a || a === '追加') return '追加';
  if (a === '変更') return '変更';
  if (a === '廃止' || a === '削除') return '廃止';
  return '追加';
}

/** 作業日 D にこのマスタ行を表示するか */
function isMasterRowActiveOnDate_(action, revisionDate, workDate) {
  var rev = parseMasterRevisionDate_(revisionDate);
  var d = normalizeWorkDate(workDate);
  if (action === '廃止') return rev > d;
  return rev <= d;
}

function parseMasterRevisionFields_(headers, row) {
  var revCol = findMasterColumnIndex_(headers, ['改定日']);
  if (revCol === -1) {
    return {
      revisionDate: CHECK_MASTER_DEFAULT_REVISION,
      action: '追加',
      prevName: ''
    };
  }
  var actCol = findMasterColumnIndex_(headers, ['アクション']);
  var prevCol = findMasterColumnIndex_(headers, ['旧項目名', '旧名称']);
  return {
    revisionDate: parseMasterRevisionDate_(row[revCol]),
    action: normalizeMasterAction_(actCol >= 0 ? row[actCol] : ''),
    prevName: prevCol >= 0 ? String(row[prevCol] || '').trim() : ''
  };
}

function shouldIncludeMasterRow_(headers, row, workDate) {
  var revCol = findMasterColumnIndex_(headers, ['改定日']);
  if (revCol === -1) return true;
  var fields = parseMasterRevisionFields_(headers, row);
  return isMasterRowActiveOnDate_(fields.action, fields.revisionDate, workDate);
}

function getMasterRevisionFields_(headers, row) {
  return parseMasterRevisionFields_(headers, row);
}

function createEmptyLineCheckBundle_() {
  var bundle = {
    daily: [],
    sensorBySlot: {},
    regularBySlot: {},
    switchoverBySlot: {},
    hinSwitchCount: 1,
    switchoverCount: 0,
    varietySelection: '未選択'
  };
  SENSOR_SLOTS_BASE.forEach(function(s) { bundle.sensorBySlot[s] = []; });
  bundle.sensorBySlot[hinSwitchSlotName_(1)] = [];
  REGULAR_SLOTS_ALL.forEach(function(s) { bundle.regularBySlot[s] = []; });
  return bundle;
}

function ensureLineBundle_(data, lineName) {
  if (!data[lineName]) data[lineName] = createEmptyLineCheckBundle_();
  return data[lineName];
}

function ensureHinSwitchSlots_(bundle) {
  var n = Math.max(1, parseInt(bundle.hinSwitchCount, 10) || 1);
  bundle.hinSwitchCount = n;
  for (var i = 1; i <= n; i++) {
    var key = hinSwitchSlotName_(i);
    if (!bundle.sensorBySlot[key]) bundle.sensorBySlot[key] = [];
  }
}

function ensureSwitchoverSlots_(bundle) {
  var n = Math.max(0, parseInt(bundle.switchoverCount, 10) || 0);
  bundle.switchoverCount = n;
  if (!bundle.switchoverBySlot) bundle.switchoverBySlot = {};
  for (var i = 1; i <= n; i++) {
    var key = switchoverSlotName_(i);
    if (!bundle.switchoverBySlot[key]) bundle.switchoverBySlot[key] = [];
  }
}

function pushUniqueItem_(list, item) {
  for (var i = 0; i < list.length; i++) {
    if (list[i].text === item.text) return;
  }
  list.push(item);
}

function newJudgeItem_(text, judgeType, prevName) {
  return {
    text: text,
    judgeType: judgeType || '合否',
    value: (judgeType === '合否' ? '-' : ''),
    recordedAt: '',
    prevName: prevName || '',
    legacy: false
  };
}

function newSensorItem_(text, prevName) {
  return { text: text, value: '-', recordedAt: '', prevName: prevName || '', legacy: false };
}

function newDailyItem_(text, prevName) {
  return { text: text, value: false, prevName: prevName || '', legacy: false };
}

/**
 * マスタ・実績シートのデータ行（2行目〜最終行）を読む
 * ※旧 getRange(…, getLastRow()-1, …) は最終データ行を読み落としていた
 */
function readSheetDataRows_(sheet, numCols) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  var lastRow = sheet.getLastRow();
  var cols = numCols || Math.max(sheet.getLastColumn(), 1);
  return sheet.getRange(2, 1, lastRow, cols).getValues();
}

/**
 * 各種点検マスタの読み込み（作業日時点で有効な行のみ、スロット展開）
 */
function loadCheckMasterFromSheet(workDate) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return {};
  var dateKey = normalizeWorkDate(workDate || new Date());

  var result = {};
  var lineIndex = getLineMasterIndex_();
  loadDailyMaster_(ss, result, dateKey, lineIndex);
  loadSensorMaster_(ss, result, dateKey, lineIndex);
  loadRegularMaster_(ss, result, dateKey, lineIndex);
  loadSwitchoverMaster_(ss, result, dateKey);

  Object.keys(result).forEach(function(line) {
    ensureHinSwitchSlots_(result[line]);
    ensureSwitchoverSlots_(result[line]);
  });
  return result;
}

function loadDailyMaster_(ss, result, workDate, lineIndex) {
  var sheet = ss.getSheetByName('日常点検マスタ');
  if (!sheet || sheet.getLastRow() < 2) return;
  var lastCol = Math.max(sheet.getLastColumn(), 2);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(function(h) { return String(h || '').trim(); });
  var rows = readSheetDataRows_(sheet, lastCol);
  for (var i = 0; i < rows.length; i++) {
    if (!shouldIncludeMasterRow_(headers, rows[i], workDate)) continue;
    var itemText = String(rows[i][1]).trim();
    if (!itemText) continue;
    var rev = getMasterRevisionFields_(headers, rows[i]);
    var targets = resolveMasterTargetLines_(rows[i][0], lineIndex);
    targets.forEach(function(lineNo) {
      var bundle = ensureLineBundle_(result, lineNo);
      pushUniqueItem_(bundle.daily, newDailyItem_(itemText, rev.prevName));
    });
  }
}

function loadSensorMaster_(ss, result, workDate, lineIndex) {
  var sheet = ss.getSheetByName('センサチェックマスタ');
  if (!sheet || sheet.getLastRow() < 2) return;
  var lastCol = Math.max(sheet.getLastColumn(), 3);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h).trim(); });
  var rows = readSheetDataRows_(sheet, lastCol);

  for (var i = 0; i < rows.length; i++) {
    if (!shouldIncludeMasterRow_(headers, rows[i], workDate)) continue;
    var itemText = String(rows[i][1]).trim();
    if (!itemText) continue;
    var applySlots = resolveSensorApplySlots_(headers, rows[i]);
    var rev = getMasterRevisionFields_(headers, rows[i]);
    var targets = resolveMasterTargetLines_(rows[i][0], lineIndex);
    targets.forEach(function(lineNo) {
      var bundle = ensureLineBundle_(result, lineNo);
      ensureHinSwitchSlots_(bundle);

      SENSOR_SLOTS_BASE.forEach(function(slot) {
        if (masterRowAppliesToSlot_(applySlots, slot, bundle.hinSwitchCount)) {
          pushUniqueItem_(bundle.sensorBySlot[slot], newSensorItem_(itemText, rev.prevName));
        }
      });
      for (var h = 1; h <= bundle.hinSwitchCount; h++) {
        var hSlot = hinSwitchSlotName_(h);
        if (masterRowAppliesToSlot_(applySlots, hSlot, bundle.hinSwitchCount)) {
          pushUniqueItem_(bundle.sensorBySlot[hSlot], newSensorItem_(itemText, rev.prevName));
        }
      }
    });
  }
}

function loadSwitchoverMaster_(ss, result, workDate) {
  var sheet = ss.getSheetByName('切替点検マスタ');
  if (!sheet) {
    Logger.log('切替点検マスタ: シートがありません。スプレッドシートメニュー「切替点検マスタシートを作成」を実行してください。');
    return;
  }
  if (sheet.getLastRow() < 2) {
    Logger.log('切替点検マスタ: データ行がありません（2行目以降にライン・項目を入力）');
    return;
  }
  var lastCol = Math.max(sheet.getLastColumn(), 2);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(function(h) { return String(h || '').trim(); });
  var rows = readSheetDataRows_(sheet, lastCol);
  for (var i = 0; i < rows.length; i++) {
    if (!shouldIncludeMasterRow_(headers, rows[i], workDate)) continue;
    var lineNo = normalizeLineKey_(rows[i][0]);
    var itemText = String(rows[i][1]).trim();
    if (!lineNo || !itemText) continue;
    var bundle = ensureLineBundle_(result, lineNo);
    ensureSwitchoverSlots_(bundle);
    if (!bundle.switchoverMasterItems) bundle.switchoverMasterItems = [];
    bundle.switchoverMasterItems.push(itemText);
  }
}

function loadRegularMaster_(ss, result, workDate, lineIndex) {
  var sheet = ss.getSheetByName('定時検査マスタ');
  if (!sheet || sheet.getLastRow() < 2) return;
  var lastCol = Math.max(sheet.getLastColumn(), 4);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(function(h) { return String(h || '').trim(); });
  var judgeCol = findMasterColumnIndex_(headers, ['判定種別']);
  var applyCol = findMasterColumnIndex_(headers, ['実施区分']);
  if (judgeCol === -1) judgeCol = 2;
  if (applyCol === -1) applyCol = 3;
  var rows = readSheetDataRows_(sheet, lastCol);

  for (var i = 0; i < rows.length; i++) {
    if (!shouldIncludeMasterRow_(headers, rows[i], workDate)) continue;
    var itemText = String(rows[i][1]).trim();
    if (!itemText) continue;
    var judgeType = judgeCol >= 0 && rows[i][judgeCol] ? String(rows[i][judgeCol]).trim() : '合否';
    var applySlots = applyCol >= 0 ? parseApplySlots_(rows[i][applyCol]) : REGULAR_SLOTS_ALL.slice();
    if (applySlots.length === 0) applySlots = REGULAR_SLOTS_ALL.slice();
    var rev = getMasterRevisionFields_(headers, rows[i]);
    var targets = resolveMasterTargetLines_(rows[i][0], lineIndex);
    targets.forEach(function(lineNo) {
      var bundle = ensureLineBundle_(result, lineNo);
      applySlots.forEach(function(slot) {
        if (REGULAR_SLOTS_ALL.indexOf(slot) === -1) return;
        if (!bundle.regularBySlot[slot]) bundle.regularBySlot[slot] = [];
        pushUniqueItem_(bundle.regularBySlot[slot], newJudgeItem_(itemText, judgeType, rev.prevName));
      });
    });
  }
}

/**
 * 点検マスタのディープコピー（JSON より軽量）
 */
function cloneCheckMasterBundle_(bundle) {
  var cloned = createEmptyLineCheckBundle_();
  cloned.hinSwitchCount = bundle.hinSwitchCount || 1;
  cloned.switchoverCount = bundle.switchoverCount || 0;
  cloned.varietySelection = bundle.varietySelection || '未選択';

  cloned.daily = (bundle.daily || []).map(function(item) {
    return {
      text: item.text,
      value: !!item.value,
      prevName: item.prevName || '',
      legacy: !!item.legacy
    };
  });

  Object.keys(bundle.sensorBySlot || {}).forEach(function(slot) {
    cloned.sensorBySlot[slot] = (bundle.sensorBySlot[slot] || []).map(function(item) {
      return {
        text: item.text,
        value: item.value || '-',
        recordedAt: item.recordedAt || '',
        prevName: item.prevName || '',
        legacy: !!item.legacy
      };
    });
  });
  ensureHinSwitchSlots_(cloned);

  Object.keys(bundle.regularBySlot || {}).forEach(function(slot) {
    cloned.regularBySlot[slot] = (bundle.regularBySlot[slot] || []).map(function(item) {
      return {
        text: item.text,
        judgeType: item.judgeType || '合否',
        value: item.value !== undefined && item.value !== null && item.value !== ''
          ? item.value
          : ((item.judgeType === '合否') ? '-' : ''),
        recordedAt: item.recordedAt || '',
        prevName: item.prevName || '',
        legacy: !!item.legacy
      };
    });
  });

  if (bundle.switchoverMasterItems) {
    cloned.switchoverMasterItems = bundle.switchoverMasterItems.slice();
  }

  Object.keys(bundle.switchoverBySlot || {}).forEach(function(slot) {
    cloned.switchoverBySlot[slot] = (bundle.switchoverBySlot[slot] || []).map(function(item) {
      return {
        text: item.text,
        value: item.value || '-',
        recordedAt: item.recordedAt || '',
        prevName: item.prevName || '',
        legacy: !!item.legacy
      };
    });
  });
  ensureSwitchoverSlots_(cloned);
  return cloned;
}

function cloneCheckMasterData_(masterData) {
  var result = {};
  Object.keys(masterData || {}).forEach(function(line) {
    result[line] = cloneCheckMasterBundle_(masterData[line]);
  });
  return result;
}

/**
 * 点検マスタ＋保存済み実績をマージ
 */
function mergeCheckDataWithResults(masterData, workDate) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('点検実績');
  var saved = sheet ? readCheckResultRows_(sheet, workDate) : [];
  var result = cloneCheckMasterData_(masterData);

  saved.forEach(function(row) {
    if (row.type === '設定' && row.text === '品種選択') {
      var b0 = ensureLineBundle_(result, row.line);
      b0.varietySelection = row.value || '未選択';
      return;
    }
    if (row.type === '設定' && row.text === '品切回数') {
      var b1 = ensureLineBundle_(result, row.line);
      b1.hinSwitchCount = Math.max(1, parseInt(row.value, 10) || 1);
      ensureHinSwitchSlots_(b1);
      return;
    }
    if (row.type === '設定' && row.text === '切替回数') {
      var b2 = ensureLineBundle_(result, row.line);
      b2.switchoverCount = Math.max(0, parseInt(row.value, 10) || 0);
      ensureSwitchoverSlots_(b2);
      return;
    }

    var bundle = ensureLineBundle_(result, row.line);
    ensureHinSwitchSlots_(bundle);
    ensureSwitchoverSlots_(bundle);
    var slot = row.slot || '';

    if (row.type === 'daily') {
      applyDailyValue_(bundle.daily, row.text, row.value === '済');
    } else if (row.type === 'sensor') {
      var sensorSlot = slot || SENSOR_SLOTS_BASE[0];
      if (!bundle.sensorBySlot[sensorSlot]) bundle.sensorBySlot[sensorSlot] = [];
      applySensorValue_(bundle.sensorBySlot[sensorSlot], row.text, row.value, row.recordedAt);
    } else if (row.type === 'regular') {
      var regSlot = slot || REGULAR_SLOTS_ALL[0];
      if (!bundle.regularBySlot[regSlot]) bundle.regularBySlot[regSlot] = [];
      applyRegularValue_(bundle.regularBySlot[regSlot], row.text, row.judgeType, row.value, row.recordedAt);
    } else if (row.type === 'switchover') {
      var swSlot = slot || switchoverSlotName_(1);
      if (!bundle.switchoverBySlot[swSlot]) bundle.switchoverBySlot[swSlot] = [];
      applySensorValue_(bundle.switchoverBySlot[swSlot], row.text, row.value, row.recordedAt);
    }
  });

  Object.keys(masterData || {}).forEach(function(line) {
    if (!masterData[line] || !masterData[line].switchoverMasterItems || !masterData[line].switchoverMasterItems.length) return;
    var bundle = ensureLineBundle_(result, line);
    bundle.switchoverMasterItems = masterData[line].switchoverMasterItems.slice();
  });
  Object.keys(result).forEach(function(line) {
    ensureHinSwitchSlots_(result[line]);
    ensureSwitchoverSlots_(result[line]);
  });
  return result;
}

function applyDailyValue_(list, text, value) {
  for (var i = 0; i < list.length; i++) {
    if (list[i].text === text) {
      list[i].value = value;
      return true;
    }
    if (list[i].prevName && list[i].prevName === text) {
      list[i].value = value;
      return true;
    }
  }
  list.push({ text: text, value: value, prevName: '', legacy: true });
  return true;
}

function applySensorValue_(list, text, value, recordedAt) {
  for (var i = 0; i < list.length; i++) {
    if (list[i].text === text || (list[i].prevName && list[i].prevName === text)) {
      list[i].value = value;
      if (recordedAt) list[i].recordedAt = recordedAt;
      return true;
    }
  }
  list.push({
    text: text,
    value: value,
    recordedAt: recordedAt || '',
    prevName: '',
    legacy: true
  });
  return true;
}

function applyRegularValue_(list, text, judgeType, value, recordedAt) {
  for (var i = 0; i < list.length; i++) {
    if (list[i].text === text || (list[i].prevName && list[i].prevName === text)) {
      list[i].value = value;
      if (recordedAt) list[i].recordedAt = recordedAt;
      return true;
    }
  }
  list.push({
    text: text,
    judgeType: judgeType || '合否',
    value: value,
    recordedAt: recordedAt || '',
    prevName: '',
    legacy: true
  });
  return true;
}

/**
 * 点検結果の保存
 */
function saveCheckResultsForDate_(ss, workDate, lineCheckData) {
  if (!lineCheckData) return;
  var sheet = ss.getSheetByName('点検実績') || ss.insertSheet('点検実績');
  var kept = readCheckResultRows_(sheet).filter(function(r) { return r.workDate !== workDate; });
  var newRows = [];
  var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm');

  Object.keys(lineCheckData).forEach(function(lineName) {
    var bundle = lineCheckData[lineName];
    if (!bundle) return;
    ensureHinSwitchSlots_(bundle);

    newRows.push({
      workDate: workDate, line: lineName, type: '設定', slot: '', text: '品種選択',
      judgeType: '-', value: bundle.varietySelection || '未選択', recordedAt: '', variety: ''
    });
    newRows.push({
      workDate: workDate, line: lineName, type: '設定', slot: '', text: '品切回数',
      judgeType: '-', value: String(bundle.hinSwitchCount || 1), recordedAt: '', variety: ''
    });
    newRows.push({
      workDate: workDate, line: lineName, type: '設定', slot: '', text: '切替回数',
      judgeType: '-', value: String(bundle.switchoverCount || 0), recordedAt: '', variety: ''
    });

    (bundle.daily || []).forEach(function(item) {
      newRows.push({
        workDate: workDate, line: lineName, type: 'daily', slot: '当日',
        text: item.text, judgeType: '-',
        value: item.value ? '済' : '未', recordedAt: item.recordedAt || '', variety: ''
      });
    });

    Object.keys(bundle.sensorBySlot || {}).forEach(function(slot) {
      (bundle.sensorBySlot[slot] || []).forEach(function(item) {
        newRows.push({
          workDate: workDate, line: lineName, type: 'sensor', slot: slot,
          text: item.text, judgeType: '-',
          value: String(item.value || '-'), recordedAt: item.recordedAt || '', variety: ''
        });
      });
    });

    Object.keys(bundle.regularBySlot || {}).forEach(function(slot) {
      (bundle.regularBySlot[slot] || []).forEach(function(item) {
        newRows.push({
          workDate: workDate, line: lineName, type: 'regular', slot: slot,
          text: item.text, judgeType: item.judgeType || '合否',
          value: String(item.value !== undefined && item.value !== null ? item.value : '-'),
          recordedAt: item.recordedAt || '', variety: ''
        });
      });
    });

    Object.keys(bundle.switchoverBySlot || {}).forEach(function(slot) {
      (bundle.switchoverBySlot[slot] || []).forEach(function(item) {
        newRows.push({
          workDate: workDate, line: lineName, type: 'switchover', slot: slot,
          text: item.text, judgeType: '-',
          value: String(item.value || '-'), recordedAt: item.recordedAt || '', variety: ''
        });
      });
    });
  });

  writeCheckResultSheet_(sheet, kept.concat(newRows));
}

function readCheckResultRows_(sheet, filterDateKey) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  var lastCol = Math.max(sheet.getLastColumn(), CHECK_RESULT_HEADERS.length);
  var headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h).trim(); });
  var hasSlot = headerRow.indexOf('スロット') >= 0;
  var colMap = buildCheckResultColMap_(headerRow, hasSlot);

  var data = readSheetDataRows_(sheet, lastCol);
  var rows = [];
  for (var i = 0; i < data.length; i++) {
    if (!data[i][colMap.workDate]) continue;
    var rowDate = normalizeWorkDate(data[i][colMap.workDate]);
    if (filterDateKey && rowDate !== filterDateKey) continue;
    rows.push({
      workDate: rowDate,
      line: normalizeLineKey_(data[i][colMap.line]),
      type: String(data[i][colMap.type]),
      slot: hasSlot ? String(data[i][colMap.slot] || '') : '',
      text: String(data[i][colMap.text]),
      judgeType: String(data[i][colMap.judgeType] || '-'),
      value: String(data[i][colMap.value] !== undefined && data[i][colMap.value] !== null ? data[i][colMap.value] : ''),
      recordedAt: colMap.recordedAt >= 0 ? String(data[i][colMap.recordedAt] || '') : '',
      variety: colMap.variety >= 0 ? String(data[i][colMap.variety] || '') : ''
    });
  }
  return rows;
}

function buildCheckResultColMap_(headerRow, hasSlot) {
  if (hasSlot) {
    return {
      workDate: 0, line: 1, type: 2, slot: 3, text: 4, judgeType: 5, value: 6,
      recordedAt: headerRow.indexOf('記録時刻'), variety: headerRow.indexOf('品種選択')
    };
  }
  return {
    workDate: 0, line: 1, type: 2, slot: -1, text: 3, judgeType: 4, value: 5,
    recordedAt: -1, variety: -1
  };
}

function writeCheckResultSheet_(sheet, rows) {
  sheet.clearContents();
  sheet.getRange(1, 1, 1, CHECK_RESULT_HEADERS.length).setValues([CHECK_RESULT_HEADERS]);
  if (!rows || rows.length === 0) return;
  var values = rows.map(function(r) {
    return [
      r.workDate, r.line, r.type, r.slot || '', r.text, r.judgeType || '-',
      r.value, r.recordedAt || '', r.variety || ''
    ];
  });
  sheet.getRange(2, 1, values.length, CHECK_RESULT_HEADERS.length).setValues(values);
}

/**
 * マトリクス表示用データ
 */
function buildCheckMatrixData(lineCheckData, lineNo) {
  var bundle = lineCheckData && lineNo ? lineCheckData[lineNo] : null;
  if (!bundle) return { daily: [], sensor: { slots: [], rows: [] }, regular: { slots: [], rows: [] } };

  ensureHinSwitchSlots_(bundle);
  ensureSwitchoverSlots_(bundle);
  var sensorSlots = SENSOR_SLOTS_BASE.slice();
  for (var h = 1; h <= bundle.hinSwitchCount; h++) sensorSlots.push(hinSwitchSlotName_(h));
  var switchoverSlots = [];
  for (var s = 1; s <= bundle.switchoverCount; s++) switchoverSlots.push(switchoverSlotName_(s));

  return {
    varietySelection: bundle.varietySelection || '未選択',
    daily: (bundle.daily || []).map(function(it) {
      return { text: it.text, value: it.value ? '済' : '未' };
    }),
    sensor: {
      slots: sensorSlots,
      rows: buildMatrixRowsFromSlots_(bundle.sensorBySlot, sensorSlots, false)
    },
    switchover: {
      slots: switchoverSlots,
      rows: buildMatrixRowsFromSlots_(bundle.switchoverBySlot, switchoverSlots, false)
    },
    regular: {
      slots: REGULAR_SLOTS_ALL.slice(),
      rows: buildMatrixRowsFromSlots_(bundle.regularBySlot, REGULAR_SLOTS_ALL, true)
    }
  };
}

function buildMatrixRowsFromSlots_(bySlot, slots, includeJudgeType) {
  var itemMap = {};
  slots.forEach(function(slot) {
    (bySlot[slot] || []).forEach(function(it) {
      if (!itemMap[it.text]) {
        itemMap[it.text] = { text: it.text, judgeType: it.judgeType || '', cells: {} };
      }
      itemMap[it.text].cells[slot] = it.value !== undefined && it.value !== null && it.value !== '' ? String(it.value) : '-';
      if (it.recordedAt) itemMap[it.text].cells[slot + '_at'] = it.recordedAt;
    });
  });
  return Object.keys(itemMap).map(function(k) { return itemMap[k]; });
}

/**
 * レポート出力用：作業日時点のマスタ＋実績マージ結果を平坦化
 */
function getMergedCheckResultRowsForReport_(workDate, activeLineNo) {
  var dateKey = normalizeWorkDate(workDate);
  var checkMaster = loadCheckMasterFromSheet(dateKey);
  var merged = mergeCheckDataWithResults(checkMaster, dateKey);
  var rows = [];

  Object.keys(merged).forEach(function(lineName) {
    if (activeLineNo && lineName !== activeLineNo) return;
    var bundle = merged[lineName];
    if (!bundle) return;

    (bundle.daily || []).forEach(function(item) {
      rows.push({
        line: lineName,
        type: 'daily',
        slot: '当日',
        text: item.text,
        judgeType: '-',
        value: item.value ? '済' : '未',
        recordedAt: ''
      });
    });

    Object.keys(bundle.sensorBySlot || {}).forEach(function(slot) {
      (bundle.sensorBySlot[slot] || []).forEach(function(item) {
        rows.push({
          line: lineName,
          type: 'sensor',
          slot: slot,
          text: item.text,
          judgeType: '-',
          value: String(item.value || '-'),
          recordedAt: item.recordedAt || ''
        });
      });
    });

    Object.keys(bundle.regularBySlot || {}).forEach(function(slot) {
      (bundle.regularBySlot[slot] || []).forEach(function(item) {
        rows.push({
          line: lineName,
          type: 'regular',
          slot: slot,
          text: item.text,
          judgeType: item.judgeType || '合否',
          value: String(item.value !== undefined && item.value !== null ? item.value : '-'),
          recordedAt: item.recordedAt || ''
        });
      });
    });

    Object.keys(bundle.switchoverBySlot || {}).forEach(function(slot) {
      (bundle.switchoverBySlot[slot] || []).forEach(function(item) {
        rows.push({
          line: lineName,
          type: 'switchover',
          slot: slot,
          text: item.text,
          judgeType: '-',
          value: String(item.value || '-'),
          recordedAt: item.recordedAt || ''
        });
      });
    });
  });

  return rows;
}

/**
 * 点検マスタ改定日・アクションの入力チェック（メニュー用）
 */
function validateCheckMasterRevisions_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetNames = ['日常点検マスタ', 'センサチェックマスタ', '定時検査マスタ', '切替点検マスタ'];
  var warnings = [];

  sheetNames.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() < 2) return;
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
      .map(function(h) { return String(h || '').trim(); });
    if (findMasterColumnIndex_(headers, ['改定日']) === -1) {
      warnings.push(name + ': 「改定日」列がありません（既存行は全期間有効として扱われます）');
      return;
    }
    var actCol = findMasterColumnIndex_(headers, ['アクション']);
    var prevCol = findMasterColumnIndex_(headers, ['旧項目名', '旧名称']);
    var rows = readSheetDataRows_(sheet, lastCol);
    for (var i = 0; i < rows.length; i++) {
      var rowNum = i + 2;
      var action = normalizeMasterAction_(actCol >= 0 ? rows[i][actCol] : '');
      var prevName = prevCol >= 0 ? String(rows[i][prevCol] || '').trim() : '';
      if (action === '変更' && !prevName) {
        warnings.push(name + ' 行' + rowNum + ': アクション「変更」ですが旧項目名が空です');
      }
    }
  });

  return warnings;
}
