// ========================================
// 切替点検マスタ（C-8 等）定義・取込
// ========================================

var SWITCHOVER_MASTER_HEADERS = [
  'ライン', 'No', '機械区分', '項目', '判定種別', '点検者', '改定日', 'アクション', '旧項目名'
];

/** 切替1回分のヘッダ・フッター項目キー */
var SWITCHOVER_META_FIELD_KEYS = [
  '切替品目_前', '切替品目_後', 'ロット番号_前', 'ロット番号_後',
  '切替開始時刻', '製造開始時刻', '第1点検者', '第2点検者',
  '第1点検者結果', '第2点検者結果', '発見箇所'
];

/** 点検者列の正規化（両方 / 1 / 2） */
function normalizeSwitchoverInspectorMode_(mode) {
  var m = String(mode || '').trim();
  if (m === '1' || m === '第1' || m === '第1のみ') return '1';
  if (m === '2' || m === '第2' || m === '第2のみ') return '2';
  return '両方';
}

function switchoverItemNeedsInspector1_(itemOrDef) {
  return normalizeSwitchoverInspectorMode_(itemOrDef && itemOrDef.inspectorMode) !== '2';
}

function switchoverItemNeedsInspector2_(itemOrDef) {
  return normalizeSwitchoverInspectorMode_(itemOrDef && itemOrDef.inspectorMode) !== '1';
}

/** 切替点検表（C-8専用）2024-4-15改定 — 34項目 */
var SWITCHOVER_C8_MASTER_ROWS = [
  [1, '絞り', 'マガジン'],
  [2, '絞り', '金型'],
  [3, '絞り', '出口シュート'],
  [4, '絞り', '立上コンベア付近'],
  [5, '爪曲げ', '入り口シュート'],
  [6, '爪曲げ', '裏向き排出のカゴ'],
  [7, '爪曲げ', '金型'],
  [8, '爪曲げ', '出口シュート'],
  [9, '爪曲げ', '立上コンベア付近'],
  [10, 'PEオート', '入り口シュート'],
  [11, 'PEオート', 'PEホッパー'],
  [12, 'PEオート', 'PEパーツフィーダ'],
  [13, 'PEオート', 'マガジン'],
  [14, 'PEオート', '機械内'],
  [15, 'PEオート', '不良排出シュート・カゴ'],
  [16, 'PEオート', '出口シュート'],
  [17, 'セット機', '立ち上げコンベア'],
  [18, 'セット機', 'キャップ入り口コンベア'],
  [19, 'セット機', '機械内'],
  [20, 'セット機', 'セット不良排出カゴ'],
  [21, 'セット機', '封印パーツフィーダー（ブリキ・ポリプロ）'],
  [22, 'セット機', '封印ホッパー（ブリキ・ポリプロ）'],
  [23, 'セット機', 'ホッパー下の排出縁カゴ内'],
  [24, 'セット機', '封印振り分け分岐部'],
  [25, 'セット機', 'セット機封印入り口コンベア上'],
  [26, 'セット機', '出口シュート'],
  [27, 'セット機', 'バイパスシュート'],
  [28, '画像処理装置', '表面・裏面コンベア上'],
  [29, '画像処理装置', 'リング浮き検出排出カゴ'],
  [30, '画像処理装置', '表面画像排出カゴ'],
  [31, '画像処理装置', '裏面画像排出カゴ'],
  [32, '画像処理装置', '機械上'],
  [33, '画像処理装置', 'ケース強制排出'],
  [34, '画像処理装置', 'ケース内確認']
];

function buildSwitchoverC8MasterSheetRows_() {
  return buildSwitchoverMasterSheetRowsForLine_('C-8', SWITCHOVER_C8_MASTER_ROWS, '2024-04-15');
}

/** 切替点検表（C-9専用）2026-4-15改定 — 本表32項目 + 別表12項目 */
var SWITCHOVER_C9_MASTER_ROWS = [
  [1, '絞り', 'マガジン'],
  [2, '絞り', '金型内'],
  [3, '絞り', 'スタッカーチェーンカバー'],
  [4, '絞り', '出口シュート'],
  [5, '絞り', 'メインモーター上'],
  [6, '絞り', '立上コンベア側面及び上部カバー'],
  [7, '爪曲げ', '入口シュート'],
  [8, '爪曲げ', '裏向き排出のカゴ'],
  [9, '爪曲げ', '金型'],
  [10, '爪曲げ', 'メインモーター付近'],
  [11, '爪曲げ', '立上コンベア側面及び上部カバー'],
  [12, '爪曲げ', '油受け'],
  [13, '爪曲げ', 'ゲージ置場・排出カゴ'],
  [14, '爪曲げ', '出口シュート'],
  [15, 'オートパッカー', '入口シュート'],
  [16, 'オートパッカー', '機械内及びターレット内'],
  [17, 'オートパッカー', '出口シュートから立上コンベアー'],
  [18, 'オートパッカー', 'バイパスシュート'],
  [19, 'オートパッカー', '排出シュート'],
  [20, 'オート', '立ち上げコンベアー接続部分'],
  [21, 'オート', '立ち上げコンベアー頂上付近'],
  [23, '小爪曲げ', '入口シュート'],
  [24, '小爪曲げ', '振り分け供給部シュート'],
  [25, '小爪曲げ', '機械内及びターレット内'],
  [26, '小爪曲げ', '出口シュート'],
  [27, '小爪曲げ', '排出シュート'],
  [28, 'セット機', 'バイパスシュート'],
  [29, 'セット機', 'キャップ入口シュート'],
  [30, 'セット機', 'ドライポンプ付近'],
  [31, 'セット機', '封印入口シュート(ポリプロ・ブリキ)'],
  [32, 'セット機', '封印立ち上げコンベア(ポリプロ・ブリキ)'],
  [33, 'セット機', 'パーツフィーダ、ホッパー(ポリプロ・ブリキ)'],
  [34, 'セット機', '排出シュート及び箱'],
  [35, 'セット機', '出口コンベア'],
  [36, 'セット機', '画像処理排出シュート(表面)'],
  [37, 'セット機', '移送コンベア裏及び下'],
  [38, 'セット機', 'セット機内'],
  [39, 'セット機', 'スイッチボックス上'],
  [40, '検査コンベアー', 'シュート出口からコンベア上'],
  [41, '検査コンベアー', '配線付近'],
  [42, '検査コンベアー', '画像処理排出シュート(裏面)'],
  [43, '検査コンベアー', 'ラウンジ内、不良仮置きカゴ'],
  [44, '検査コンベアー', '高さ検出排出シュート'],
  [45, '検査コンベアー', 'ケース内残品']
];

function buildSwitchoverC9MasterSheetRows_() {
  return buildSwitchoverMasterSheetRowsForLine_('C-9', SWITCHOVER_C9_MASTER_ROWS, '2026-04-15');
}

function buildSwitchoverMasterSheetRowsForLine_(lineKey, masterRows, revisionDate) {
  return masterRows.map(function(row) {
    return [lineKey, row[0], row[1], row[2], 'XO', '両方', revisionDate, '追加', ''];
  });
}

/**
 * 切替点検マスタの指定ライン行を差し替えて書き込む
 */
function importSwitchoverMasterForLine_(ss, lineKey, newRows, options) {
  options = options || {};
  var replaceLine = options.replaceLine !== false;
  lineKey = normalizeLineKey_(lineKey);
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var spec = {
    name: '切替点検マスタ',
    headers: SWITCHOVER_MASTER_HEADERS,
    tabColor: '#fbcfe8',
    samples: null
  };
  ensureSheetWithHeaders_(ss, spec, true);
  var sheet = ss.getSheetByName(spec.name);
  if (!sheet) throw new Error('切替点検マスタシートを作成できませんでした。');

  if (replaceLine && sheet.getLastRow() >= 2) {
    var lastCol = Math.max(sheet.getLastColumn(), SWITCHOVER_MASTER_HEADERS.length);
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
      .map(function(h) { return String(h || '').trim(); });
    var lineCol = findMasterColumnIndex_(headers, ['ライン']);
    if (lineCol === -1) lineCol = 0;
    var data = readSheetDataRows_(sheet, lastCol);
    var kept = [];
    for (var i = 0; i < data.length; i++) {
      var rowLine = normalizeLineKey_(data[i][lineCol]);
      if (rowLine !== lineKey) kept.push(data[i]);
    }
    var dataRowCount = Math.max(sheet.getLastRow() - 1, 0);
    if (dataRowCount > 0) {
      sheet.getRange(2, 1, dataRowCount, lastCol).clearContent();
    }
    var writeRow = 2;
    kept.forEach(function(row) {
      var cols = Math.max(row.length, SWITCHOVER_MASTER_HEADERS.length);
      var padded = row.slice();
      while (padded.length < cols) padded.push('');
      sheet.getRange(writeRow, 1, 1, cols).setValues([padded]);
      writeRow++;
    });
    if (newRows.length) {
      sheet.getRange(writeRow, 1, newRows.length, SWITCHOVER_MASTER_HEADERS.length).setValues(newRows);
    }
  } else if (newRows.length) {
    var startRow = sheet.getLastRow() < 2 ? 2 : sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, newRows.length, SWITCHOVER_MASTER_HEADERS.length).setValues(newRows);
  }
  try { sheet.setTabColor(spec.tabColor); } catch (e) { /* ignore */ }
  clearMasterCaches_();
  return newRows.length;
}

/**
 * 切替点検マスタ（C-8 34項目）をシートへ書き込む
 */
function importSwitchoverMasterC8_(ss, options) {
  return importSwitchoverMasterForLine_(ss, 'C-8', buildSwitchoverC8MasterSheetRows_(), options);
}

/**
 * 切替点検マスタ（C-9 44項目）をシートへ書き込む
 */
function importSwitchoverMasterC9_(ss, options) {
  return importSwitchoverMasterForLine_(ss, 'C-9', buildSwitchoverC9MasterSheetRows_(), options);
}

function createEmptySwitchoverMeta_() {
  var meta = {};
  SWITCHOVER_META_FIELD_KEYS.forEach(function(k) {
    if (k === '第1点検者結果' || k === '第2点検者結果') {
      meta[k] = '-';
    } else {
      meta[k] = '';
    }
  });
  return meta;
}

function newSwitchoverItemFromDef_(def, prevName) {
  var judgeType = def.judgeType || '合否';
  return {
    no: def.no || 0,
    section: def.section || '',
    text: def.text,
    judgeType: judgeType,
    inspectorMode: normalizeSwitchoverInspectorMode_(def.inspectorMode),
    value: '-',
    value1: '-',
    value2: '-',
    recordedAt: '',
    recordedAt1: '',
    recordedAt2: '',
    prevName: prevName || '',
    legacy: false
  };
}

function normalizeSwitchoverItem_(item, def) {
  var judgeType = (def && def.judgeType) || item.judgeType || '合否';
  var normalized = {
    no: (def && def.no) || item.no || 0,
    section: (def && def.section) || item.section || '',
    text: (def && def.text) || item.text,
    judgeType: judgeType,
    inspectorMode: normalizeSwitchoverInspectorMode_((def && def.inspectorMode) || item.inspectorMode),
    value: item.value || '-',
    value1: item.value1 !== undefined ? item.value1 : '-',
    value2: item.value2 !== undefined ? item.value2 : '-',
    recordedAt: item.recordedAt || '',
    recordedAt1: item.recordedAt1 || '',
    recordedAt2: item.recordedAt2 || '',
    prevName: item.prevName || (def && def.prevName) || '',
    legacy: !!item.legacy
  };
  if (judgeType !== 'XO' && normalized.value === '-' && normalized.value1 !== '-') {
    normalized.value = normalized.value1;
  }
  return normalized;
}

function isSwitchoverXoForm_(bundle) {
  if (!bundle) return false;
  if (bundle.switchoverMasterDefs && bundle.switchoverMasterDefs.length) {
    return bundle.switchoverMasterDefs.some(function(d) { return d.judgeType === 'XO'; });
  }
  var slots = bundle.switchoverBySlot || {};
  return Object.keys(slots).some(function(slot) {
    return (slots[slot] || []).some(function(it) { return it.judgeType === 'XO'; });
  });
}

function switchoverItemMergeKey_(itemOrDef) {
  var no = parseInt(itemOrDef && itemOrDef.no, 10) || 0;
  if (no > 0) return 'no:' + no;
  return 'text:' + String((itemOrDef && itemOrDef.text) || '').trim();
}

function buildSwitchoverPrevLookup_(existing, defs) {
  var lookup = {};
  var duplicateTexts = {};
  (defs || []).forEach(function(d) {
    var t = String(d.text || '').trim();
    duplicateTexts[t] = (duplicateTexts[t] || 0) + 1;
  });
  (existing || []).forEach(function(it, idx) {
    lookup[switchoverItemMergeKey_(it)] = it;
    lookup['idx:' + idx] = it;
    var t = String(it.text || '').trim();
    if ((duplicateTexts[t] || 0) <= 1) lookup['text:' + t] = it;
  });
  return lookup;
}

function findPrevSwitchoverItem_(def, index, lookup) {
  if (!lookup) return null;
  var no = parseInt(def && def.no, 10) || 0;
  if (no > 0 && lookup['no:' + no]) return lookup['no:' + no];
  var t = String((def && def.text) || '').trim();
  if (lookup['text:' + t]) return lookup['text:' + t];
  if (lookup['idx:' + index] !== undefined) return lookup['idx:' + index];
  return null;
}

/** 点検実績シート保存用（同名項目は No 付き） */
function formatSwitchoverItemStoredText_(item) {
  var no = parseInt(item && item.no, 10) || 0;
  var text = String((item && item.text) || '').trim();
  if (no > 0) return 'No' + no + ':' + text;
  return text;
}

/** 点検実績シート読込用 */
function parseSwitchoverItemStoredText_(stored) {
  var s = String(stored || '').trim();
  var m = /^No(\d+):([\s\S]*)$/.exec(s);
  if (m) return { no: parseInt(m[1], 10) || 0, text: m[2] };
  return { no: 0, text: s };
}

function findSwitchoverListIndex_(list, itemNo, itemText) {
  var i;
  if (itemNo > 0) {
    for (i = 0; i < list.length; i++) {
      if ((parseInt(list[i].no, 10) || 0) === itemNo) return i;
    }
  }
  for (i = 0; i < list.length; i++) {
    if (list[i].text === itemText || (list[i].prevName && list[i].prevName === itemText)) return i;
  }
  return -1;
}

function formatSwitchoverCellValue_(item) {
  if (!item) return '-';
  if (item.judgeType === 'XO') {
    return String(item.value1 || '-') + '/' + String(item.value2 || '-');
  }
  if (item.value !== undefined && item.value !== null && item.value !== '') {
    return String(item.value);
  }
  return '-';
}

function ensureSwitchoverSlotItemsFromMaster_(bundle) {
  if (!bundle || !bundle.switchoverMasterDefs || !bundle.switchoverMasterDefs.length) return;
  ensureSwitchoverSlots_(bundle);
  for (var i = 1; i <= bundle.switchoverCount; i++) {
    var slot = switchoverSlotName_(i);
    var existing = bundle.switchoverBySlot[slot] || [];
    var lookup = buildSwitchoverPrevLookup_(existing, bundle.switchoverMasterDefs);
    bundle.switchoverBySlot[slot] = bundle.switchoverMasterDefs.map(function(def, index) {
      var prev = findPrevSwitchoverItem_(def, index, lookup);
      if (prev) return normalizeSwitchoverItem_(prev, def);
      return newSwitchoverItemFromDef_(def, def.prevName || '');
    });
  }
}

function applySwitchoverInspectorValue_(list, itemNo, itemText, inspectorNo, value, recordedAt) {
  var key = inspectorNo === 2 ? 'value2' : 'value1';
  var atKey = inspectorNo === 2 ? 'recordedAt2' : 'recordedAt1';
  var idx = findSwitchoverListIndex_(list, itemNo, itemText);
  if (idx >= 0) {
    list[idx][key] = value;
    if (recordedAt) list[idx][atKey] = recordedAt;
    return true;
  }
  var legacy = {
    no: itemNo || 0,
    text: itemText,
    judgeType: 'XO',
    value: '-',
    value1: inspectorNo === 1 ? value : '-',
    value2: inspectorNo === 2 ? value : '-',
    recordedAt: '',
    recordedAt1: inspectorNo === 1 && recordedAt ? recordedAt : '',
    recordedAt2: inspectorNo === 2 && recordedAt ? recordedAt : '',
    prevName: '',
    legacy: true
  };
  list.push(legacy);
  return true;
}

function applySwitchoverSimpleValue_(list, itemNo, itemText, value, recordedAt) {
  var idx = findSwitchoverListIndex_(list, itemNo, itemText);
  if (idx >= 0) {
    list[idx].value = value;
    if (recordedAt) list[idx].recordedAt = recordedAt;
    return true;
  }
  list.push({
    no: itemNo || 0,
    text: itemText,
    value: value,
    recordedAt: recordedAt || '',
    prevName: '',
    legacy: true
  });
  return true;
}

function applySwitchoverMetaValue_(bundle, slot, key, value) {
  if (!bundle.switchoverMetaBySlot) bundle.switchoverMetaBySlot = {};
  if (!bundle.switchoverMetaBySlot[slot]) {
    bundle.switchoverMetaBySlot[slot] = createEmptySwitchoverMeta_();
  }
  bundle.switchoverMetaBySlot[slot][key] = value;
}

function menuImportSwitchoverMasterC8() {
  var ui = SpreadsheetApp.getUi();
  var confirm = ui.alert(
    '切替点検マスタ（C-8）',
    'C-8 の切替点検34項目（2024-4-15改定）を「切替点検マスタ」に書き込みます。\n'
      + '既存の C-8 行は置き換え、他ラインの行は保持します。\n\n続行しますか？',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;
  var count = importSwitchoverMasterC8_(SpreadsheetApp.getActiveSpreadsheet(), { replaceLine: true });
  ui.alert(
    '切替点検マスタ（C-8）',
    'C-8 の項目を ' + count + ' 行書き込みました。\n\n'
      + 'clasp push → Webアプリ再読み込み後、C-8 ラインで切替点検を確認してください。',
    ui.ButtonSet.OK
  );
}

function menuImportSwitchoverMasterC9() {
  var ui = SpreadsheetApp.getUi();
  var confirm = ui.alert(
    '切替点検マスタ（C-9）',
    'C-9 の切替点検44項目（2026-4-15改定・本表32+別表12）を「切替点検マスタ」に書き込みます。\n'
      + '既存の C-9 行は置き換え、他ラインの行は保持します。\n\n続行しますか？',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;
  var count = importSwitchoverMasterC9_(SpreadsheetApp.getActiveSpreadsheet(), { replaceLine: true });
  ui.alert(
    '切替点検マスタ（C-9）',
    'C-9 の項目を ' + count + ' 行書き込みました。\n\n'
      + 'clasp push → Webアプリ再読み込み後、C-9 ラインで切替点検を確認してください。',
    ui.ButtonSet.OK
  );
}
