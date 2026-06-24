// ========================================
// ⚙️ システム全体共通の設定項目
// ========================================

// 📂 共通マスタスプレッドシートのID
var MASTER_SS_ID = '1FrxPVUeKecY8SXwc5daMxjGT0MzQKZ_toa77PfO4iQo';

// 📋 共通マスタ内のライン一覧シート名
var LINE_MASTER_SHEET_NAME = 'ラインマスタ';

// 👥 社員マスタ（製造日報の作業者QR・出張旅費精算で共用）
var EMPLOYEE_MASTER_SHEET_NAME = '社員マスタ';
var WORKER_MASTER_LEGACY_SHEET = '作業者マスタ';

// 🏷️ 製品ラベル切り出し（いずれも「何文字目から」= 1始まり）
// 品目CD: 10文字目から6桁（21桁ラベル／旧実装 substring(9,15) と同等）
var PRODUCT_CODE_START  = 10;
var PRODUCT_CODE_LENGTH = 6;

// 📦 始ケースNo（例: 7文字目から3桁）
var CASE_NO_START       = 7;
var CASE_NO_LENGTH      = 3;

// 📄 レポート出力シート名
var REPORT_SHEET_NAME = '日報集計';

// マスタデータのスクリプトキャッシュ有効秒数（2回目以降の起動短縮）
var MASTER_CACHE_TTL_SEC = 600;

/** 作業マスタ「表示色」：日本語色名 → 16進カラー */
var WORK_COLOR_NAME_TO_HEX = {
  '赤': '#ef4444',
  '朱色': '#e45c2a',
  'オレンジ': '#f59e0b',
  '黄色': '#eab308',
  '黄緑': '#84cc16',
  '青紫': '#6366f1',
  '黒': '#334155',
  '青緑': '#14b8a6',
  '空色': '#38bdf8',
  '青': '#3b82f6',
  '紺色': '#1e3a8a',
  '紫': '#8b5cf6',
  '赤紫': '#c026d3',
  'ピンク': '#ec4899',
  '薄茶色': '#d4a574',
  '茶色': '#78350f',
  '灰色': '#94a3b8'
};

function getWorkColorNameMap_() {
  return WORK_COLOR_NAME_TO_HEX;
}

/**
 * 作業マスタ「表示色」を正規化（#hex / 日本語色名 / プリセット名）
 */
function normalizeWorkColorKey_(raw, fallbackName) {
  var ck = String(raw != null && raw !== '' ? raw : fallbackName || '').trim();
  ck = ck.replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(ch) {
    return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0);
  });
  ck = ck.replace(/\u3000/g, '').replace(/\s+/g, '');
  if (/^#[0-9A-Fa-f]{3,8}$/.test(ck)) return ck.toLowerCase();
  if (WORK_COLOR_NAME_TO_HEX[ck]) return WORK_COLOR_NAME_TO_HEX[ck];
  var aliases = { '生産': '製造', 'メンテ': '修理', 'メンテナンス': '修理' };
  if (aliases[ck]) ck = aliases[ck];
  return ck;
}

/**
 * 1始まりの位置指定で部分文字列を切り出す
 */
function sliceLabelByPosition(raw, startOneBased, length) {
  var s = String(raw || '').trim();
  if (!s || !startOneBased || !length) return '';
  var start = startOneBased - 1;
  if (s.length < start + length) return '';
  return s.substring(start, start + length);
}

/**
 * 製品ラベルから品目CDを抽出
 */
function extractProductCodeFromLabel(raw) {
  var part = sliceLabelByPosition(raw, PRODUCT_CODE_START, PRODUCT_CODE_LENGTH);
  if (part) return part;
  return String(raw || '').trim();
}

/**
 * 製品ラベルからケースNoを抽出（数値化、空なら "1"）
 */
function extractCaseNoFromLabel(raw) {
  var part = sliceLabelByPosition(raw, CASE_NO_START, CASE_NO_LENGTH);
  if (!part) return '1';
  var n = parseInt(part, 10);
  return String(isNaN(n) ? 1 : (n || 1));
}

/**
 * 作業日を yyyy-MM-dd に正規化
 */
function normalizeWorkDate(dateInput) {
  var tz = Session.getScriptTimeZone();
  if (!dateInput) {
    return Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  }
  if (dateInput instanceof Date && !isNaN(dateInput.getTime())) {
    return Utilities.formatDate(dateInput, tz, 'yyyy-MM-dd');
  }
  var s = String(dateInput).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  if (/^\d{4}\/\d{2}\/\d{2}$/.test(s)) return s.replace(/\//g, '-');
  var d = new Date(s);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  }
  return Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
}

/**
 * クライアントへ渡すバーコード設定
 */
function getBarcodeConfigForClient() {
  return {
    productCodeStart: PRODUCT_CODE_START,
    productCodeLength: PRODUCT_CODE_LENGTH,
    caseNoStart: CASE_NO_START,
    caseNoLength: CASE_NO_LENGTH,
    workColorNameMap: getWorkColorNameMap_()
  };
}

/**
 * スクリプトキャッシュから JSON を取得
 */
function getCachedJson_(key) {
  try {
    var raw = CacheService.getScriptCache().get(key);
    if (raw) return JSON.parse(raw);
  } catch (e) { /* ignore */ }
  return null;
}

/**
 * スクリプトキャッシュへ JSON を保存（100KB超は自動スキップ）
 */
function putCachedJson_(key, value, expirationInSeconds) {
  try {
    CacheService.getScriptCache().put(
      key,
      JSON.stringify(value),
      expirationInSeconds || MASTER_CACHE_TTL_SEC
    );
  } catch (e) { /* payload too large */ }
}

/**
 * マスタ読込キャッシュをクリア（マスタ編集後に反映させる）
 */
function clearMasterCaches_() {
  try {
    var cache = CacheService.getScriptCache();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) {
      cache.remove('checkMaster_' + ss.getId());
      cache.remove('workMaster_' + ss.getId());
    }
    if (MASTER_SS_ID) cache.remove('lineMaster_' + MASTER_SS_ID);
  } catch (e) { /* ignore */ }
}
