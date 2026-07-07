/**
 * Webアプリケーションへのアクセス制御
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('製造実績 ＆ ケース管理システム')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 画面起動時のデータ読み込み・初期化処理の統合窓口
 * @param {string} workDate - yyyy-MM-dd（省略時は本日）
 */
function getCombinedInitialData(workDate) {
  var dateKey = normalizeWorkDate(workDate);
  var prodData = loadProductionDataForDate(dateKey);
  var lineList = loadLineMasterFromSheet();
  var workMasterList = loadWorkMasterFromSheet();
  var checkMaster = loadCheckMasterFromSheet(dateKey);
  var lineCheckData = mergeCheckDataWithResults(checkMaster, dateKey);

  return {
    workDate: dateKey,
    products: prodData.products,
    timeline: prodData.timeline,
    currentJob: prodData.currentJob,
    lastPriorProduct: loadLastPriorProductForDate(dateKey),
    lineList: lineList,
    lineCheckData: lineCheckData,
    barcodeConfig: getBarcodeConfigForClient(),
    checkSlotConstants: getCheckSlotConstants(),
    workMasterList: workMasterList,
    isToday: dateKey === normalizeWorkDate(new Date()),
    checkMasterGuideUrl: getCheckMasterGuideUrl_()
  };
}

/**
 * 共通関数：時刻フォーマット整形（HH:mm）
 * スプレッドシートの時刻シリアル値・Date オブジェクトにも対応
 * @return {string|null} 解釈できない場合は null
 */
function formatTime(val) {
  if (val === null || val === undefined || val === '') return null;
  if (val instanceof Date && !isNaN(val.getTime())) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'HH:mm');
  }
  if (typeof val === 'number' && !isNaN(val)) {
    var mins;
    if (val >= 0 && val < 1) {
      mins = Math.round(val * 24 * 60);
    } else if (val >= 1 && val < 24 * 60) {
      mins = Math.round(val);
    } else {
      return null;
    }
    mins = ((mins % (24 * 60)) + (24 * 60)) % (24 * 60);
    return minutesToTimeString_(mins);
  }
  var s = String(val).trim();
  var m = s.match(/^(\d{1,2}):(\d{2})/);
  if (m) {
    return String(parseInt(m[1], 10)).padStart(2, '0') + ':' + m[2];
  }
  return null;
}

function minutesToTimeString_(mins) {
  var h = Math.floor(mins / 60);
  var m = mins % 60;
  return String(h).padStart(2, '0') + ':' + String(m).padStart(2, '0');
}

/** タイムライン用：分を5分単位に丸める */
function roundMinutesTo5_(mins) {
  var step = 5;
  return Math.round(mins / step) * step;
}

/**
 * タイムライン1行分の時刻を補正（空・逆転・軸外を修正）
 */
function normalizeTimelineBlock_(block, index) {
  var axisStart = 7 * 60 + 30;
  var axisEnd = 21 * 60;
  var slot = (index || 0) * 15;

  var startMin = timeStringToMinutes_(formatTime(block.startTime));
  var endMin = timeStringToMinutes_(formatTime(block.endTime));

  if (startMin === null) startMin = axisStart + slot;
  if (endMin === null) endMin = startMin + 15;
  if (endMin <= startMin) endMin = startMin + 15;

  startMin = Math.max(axisStart, Math.min(startMin, axisEnd - 15));
  endMin = Math.max(startMin + 5, Math.min(endMin, axisEnd));
  startMin = roundMinutesTo5_(startMin);
  endMin = roundMinutesTo5_(endMin);
  if (endMin <= startMin) endMin = startMin + 5;

  return {
    id: block.id || Date.now() + index,
    status: block.status || '未選択',
    startTime: minutesToTimeString_(startMin),
    endTime: minutesToTimeString_(endMin)
  };
}

function timeStringToMinutes_(hhmm) {
  if (!hhmm) return null;
  var parts = String(hhmm).split(':');
  if (parts.length < 2) return null;
  var h = parseInt(parts[0], 10);
  var m = parseInt(parts[1], 10);
  if (isNaN(h) || isNaN(m)) return null;
  return h * 60 + m;
}

/**
 * タイムライン配列をまとめて正規化
 */
function normalizeTimelineBlocks(blocks) {
  if (!blocks || blocks.length === 0) {
    return [normalizeTimelineBlock_({ id: Date.now(), status: '未選択', startTime: '08:30', endTime: '08:45' }, 0)];
  }
  var result = [];
  for (var i = 0; i < blocks.length; i++) {
    result.push(normalizeTimelineBlock_(blocks[i], i));
  }
  for (var j = 1; j < result.length; j++) {
    if (timeStringToMinutes_(result[j].startTime) < timeStringToMinutes_(result[j - 1].endTime)) {
      result[j].startTime = result[j - 1].endTime;
      var s = timeStringToMinutes_(result[j].startTime);
      var e = timeStringToMinutes_(result[j].endTime);
      if (e <= s) result[j].endTime = minutesToTimeString_(s + 15);
    }
  }
  return result;
}
