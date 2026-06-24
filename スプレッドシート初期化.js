/**
 * 製造日報アプリ用スプレッドシートの一括初期化
 *
 * 【実行方法】
 * 1. このプロジェクトをスプレッドシートに紐づけた状態で Apps Script エディタを開く
 * 2. 関数 initializeSpreadsheet を選択 → 「実行」
 *    またはメニュー「製造日報」→「全シート＋ヘッダーを一括作成」
 *
 * 【作業マスタとQR】
 * ・QRにエンコードするのは A列「QRコード」のみ（例: WORK_REPAIR）
 * ・B列「作業区分名」は画面表示用（修理・調整）。QRにはしない
 * ・製造①②は作業マスタに登録しない（製品スキャンで動的にタイムラインへ追加）
 */

var WORK_MASTER_HEADERS = ['QRコード', '作業区分名', '表示色', '表示順', 'QR表示'];
var WORK_MASTER_SAMPLES = [
  ['WORK_REPAIR', '修理', '修理', 1],
  ['WORK_ADJUST', '調整', '調整', 2]
];

/** 点検マスタ共通：改定日・アクション・旧項目名のデフォルト */
var CHECK_MASTER_REV_DEFAULTS = ['2000-01-01', '追加', ''];

/** 共通マスタ「社員マスタ」（旧作業者マスタと統合） */
var EMPLOYEE_MASTER_HEADERS = [
  '社員ID', '氏名', 'Email', '事業所', '部署', 'ロール', '日当(円)', '有効', 'QR表示'
];
var EMPLOYEE_MASTER_SAMPLES = [
  ['EMP001', '山田太郎', '', '本社工場', '製造課', '一般', '', '', '有効'],
  ['EMP002', '鈴木次郎', '', '本社工場', '製造課,品質管理課', '課長,品質管理担当', '', '', '有効']
];

/**
 * @param {Object} options
 * @param {boolean} options.addSamples サンプル行を入れる（既定: true）
 * @param {boolean} options.forceHeaders 既存シートの1行目もヘッダーで上書き（既定: false）
 * @return {string}
 */
function initializeSpreadsheet(options) {
  options = options || {};
  var addSamples = options.addSamples !== false;
  var forceHeaders = options.forceHeaders === true;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logs = [];

  var specs = [
    { name: '製造実績', headers: PROD_HEADERS, tabColor: '#fed7aa' },
    { name: 'タイムライン', headers: TL_HEADERS, tabColor: '#bfdbfe' },
    { name: '点検実績', headers: CHECK_RESULT_HEADERS, tabColor: '#bbf7d0' },
    {
      name: '作業マスタ',
      headers: WORK_MASTER_HEADERS,
      tabColor: '#fecaca',
      samples: addSamples ? WORK_MASTER_SAMPLES : null,
      qrSourceCol: 1,
      qrDisplayCol: 5
    },
    {
      name: '日常点検マスタ',
      headers: ['ライン', '項目', '改定日', 'アクション', '旧項目名'],
      tabColor: '#e9d5ff',
      samples: addSamples ? [
        ['1号ライン', '清掃・ゴミ箱の確認', '2000-01-01', '追加', ''],
        ['1号ライン', '潤滑油・消耗品の確認', '2000-01-01', '追加', ''],
        ['2号ライン', '清掃・ゴミ箱の確認', '2000-01-01', '追加', '']
      ] : null
    },
    {
      name: 'センサチェックマスタ',
      headers: ['ライン', '項目', '実施区分', '改定日', 'アクション', '旧項目名'],
      tabColor: '#fde68a',
      samples: addSamples ? [
        ['1号ライン', '検知センサ A（通過確認）', '午前,午後,品切', '2000-01-01', '追加', ''],
        ['1号ライン', '停止センサ B', '午前,午後', '2000-01-01', '追加', ''],
        ['2号ライン', '検知センサ A（通過確認）', '午前,午後,品切', '2000-01-01', '追加', '']
      ] : null
    },
    {
      name: '定時検査マスタ',
      headers: ['ライン', '項目', '判定種別', '実施区分', '改定日', 'アクション', '旧項目名'],
      tabColor: '#a5f3fc',
      samples: addSamples ? [
        ['1号ライン', '外観・異物の有無', '合否', '始業,10時,13時,15時,17時,終業', '2000-01-01', '追加', ''],
        ['1号ライン', '重量', 'g', '13時,終業', '2000-01-01', '追加', ''],
        ['2号ライン', '外観・異物の有無', '合否', '始業,13時,終業', '2000-01-01', '追加', '']
      ] : null
    },
    {
      name: '切替点検マスタ',
      headers: ['ライン', '項目', '改定日', 'アクション', '旧項目名'],
      tabColor: '#fbcfe8',
      samples: addSamples ? [
        ['1号ライン', '金型・治具の取り外し確認', '2000-01-01', '追加', ''],
        ['1号ライン', '前品種残材・ラベルの除去確認', '2000-01-01', '追加', ''],
        ['1号ライン', '切替後の試運転・サンプル確認', '2000-01-01', '追加', ''],
        ['2号ライン', '金型・治具の取り外し確認', '2000-01-01', '追加', ''],
        ['2号ライン', '前品種残材・ラベルの除去確認', '2000-01-01', '追加', '']
      ] : null
    },
    { name: REPORT_SHEET_NAME, headers: ['項目', '内容'], tabColor: '#f1f5f9' }
  ];

  specs.forEach(function(spec, index) {
    var result = ensureSheetWithHeaders_(ss, spec, forceHeaders);
    logs.push(result);
    var sheet = ss.getSheetByName(spec.name);
    if (sheet) {
      try {
        sheet.setTabColor(spec.tabColor || null);
        sheet.setPosition(index + 1);
      } catch (e) { /* タブ色非対応環境は無視 */ }
      if (spec.qrSourceCol && spec.qrDisplayCol) {
        applyQrDisplayFormulas_(sheet, spec.qrSourceCol, spec.qrDisplayCol);
      }
    }
  });

  writeSetupGuideSheet_(ss, forceHeaders);
  writeCheckMasterGuideSheet_(ss, forceHeaders);

  var created = logs.filter(function(l) { return l.created; }).map(function(l) { return l.name; });
  var headerUpdated = logs.filter(function(l) { return l.headerSet; }).map(function(l) { return l.name; });

  var msg = 'スプレッドシート初期化が完了しました。\n\n';
  if (created.length) msg += '【新規作成】\n・' + created.join('\n・') + '\n\n';
  else msg += '【新規作成】なし（既存シートを利用）\n\n';
  if (headerUpdated.length) msg += '【ヘッダー設定】\n・' + headerUpdated.join('\n・') + '\n';
  msg += '\n【作業マスタ】QRは A列「QRコード」を encode してください（作業区分名は不可）。\n';
  msg += '【共通マスタ】作業者・商品・不良・ラインは別ブック。メニュー「共通マスタを初期化」を参照。\n';

  Logger.log(msg);
  return msg;
}

/**
 * ヘッダー行のみ再設定（2行目以降のデータは削除しない）
 */
function resetAllSheetHeaders() {
  return initializeSpreadsheet({ addSamples: false, forceHeaders: true });
}

/**
 * 共通マスタ（設定.gs の MASTER_SS_ID）を初期化
 */
function initializeMasterSpreadsheet(options) {
  options = options || {};
  var addSamples = options.addSamples !== false;
  var forceHeaders = options.forceHeaders === true;

  if (!MASTER_SS_ID) {
    throw new Error('設定.gs の MASTER_SS_ID が未設定です。');
  }

  var ss = SpreadsheetApp.openById(MASTER_SS_ID);
  var specs = [
    {
      name: LINE_MASTER_SHEET_NAME,
      headers: ['ライン名'],
      samples: addSamples ? [['1号ライン'], ['2号ライン'], ['3号ライン']] : null
    },
    {
      name: EMPLOYEE_MASTER_SHEET_NAME,
      headers: EMPLOYEE_MASTER_HEADERS,
      samples: addSamples ? EMPLOYEE_MASTER_SAMPLES : null,
      qrSourceCol: 1,
      qrDisplayCol: 9
    },
    {
      name: '不良マスタ',
      headers: ['QRコード', '不良名', 'QR表示'],
      samples: addSamples ? [['D001', 'キズ'], ['D002', '汚れ']] : null,
      qrSourceCol: 1,
      qrDisplayCol: 3
    }
  ];

  specs.forEach(function(spec) {
    var sheet = ensureSheetWithHeaders_(ss, spec, forceHeaders);
    if (spec.qrSourceCol && spec.qrDisplayCol) {
      var sh = ss.getSheetByName(spec.name);
      if (sh) applyQrDisplayFormulas_(sh, spec.qrSourceCol, spec.qrDisplayCol);
    }
  });

  return '共通マスタ（ID: ' + MASTER_SS_ID + '）を初期化しました。\n\n'
    + '・' + LINE_MASTER_SHEET_NAME + '\n'
    + '・社員マスタ（社員ID・氏名・Email・権限・QR表示 等）\n'
    + '・不良マスタ\n\n'
    + '【QR表示】A列「社員ID」を QR 化します（= 旧作業者QRコード）。\n'
    + '【統合】既存の作業者マスタがある場合はメニュー「作業者マスタを社員マスタへ統合」を実行。\n'
    + '※ 品目マスタはシート名任意。各シートに「品目CD」「品名」「入数」列を用意してください。';
}

/**
 * QR表示列に quickchart 数式を設定（2行目以降）
 * @param {Sheet} sheet
 * @param {number} sourceCol 1始まり（QRコード列）
 * @param {number} displayCol 1始まり（QR表示列）
 */
function applyQrDisplayFormulas_(sheet, sourceCol, displayCol) {
  if (!sheet || sheet.getLastRow() < 2) return;
  var lastRow = sheet.getLastRow();
  var sourceLetter = columnToLetter_(sourceCol);
  for (var r = 2; r <= lastRow; r++) {
    var code = String(sheet.getRange(r, sourceCol).getValue() || '').trim();
    if (!code) continue;
    var formula = '=IMAGE("https://quickchart.io/qr?size=150&text="&ENCODEURL(' + sourceLetter + r + '))';
    sheet.getRange(r, displayCol).setFormula(formula);
  }
  sheet.setColumnWidth(displayCol, 160);
}

function columnToLetter_(col) {
  var letter = '';
  while (col > 0) {
    var mod = (col - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

/**
 * @param {Spreadsheet} ss
 * @param {Object} spec
 * @param {boolean} forceHeaders
 */
function ensureSheetWithHeaders_(ss, spec, forceHeaders) {
  var sheet = ss.getSheetByName(spec.name);
  var created = false;

  if (!sheet) {
    sheet = ss.insertSheet(spec.name);
    created = true;
  }

  var headerNeedsUpdate = created || forceHeaders || sheet.getLastRow() === 0 || !headersMatch_(
    sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), spec.headers.length)).getValues()[0],
    spec.headers
  );

  if (headerNeedsUpdate) {
    sheet.getRange(1, 1, 1, spec.headers.length).setValues([spec.headers]);
    sheet.getRange(1, 1, 1, spec.headers.length)
      .setFontWeight('bold')
      .setBackground('#e2e8f0')
      .setWrap(true);
  }

  if (spec.samples && (created || forceHeaders) && spec.samples.length > 0) {
    var sampleColCount = spec.headers.length;
    var padded = spec.samples.map(function(row) {
      var copy = row.slice();
      while (copy.length < sampleColCount) copy.push('');
      return copy.slice(0, sampleColCount);
    });
    var existingRows = Math.max(0, sheet.getLastRow() - 1);
    if (created || existingRows === 0) {
      if (sheet.getLastRow() > 1) {
        sheet.deleteRows(2, sheet.getLastRow() - 1);
      }
      sheet.getRange(2, 1, padded.length, sampleColCount).setValues(padded);
    }
  }

  sheet.setFrozenRows(1);

  if (spec.name === '作業マスタ') {
    sheet.setColumnWidth(1, 120);
    sheet.setColumnWidth(2, 100);
    sheet.setColumnWidth(3, 72);
    sheet.setColumnWidth(4, 56);
  }

  if (spec.name === REPORT_SHEET_NAME && created) {
    sheet.getRange(2, 1, 2, 2).setValues([['説明', 'アプリのレポート出力で自動生成されます']]);
  }

  return { name: spec.name, created: created, headerSet: headerNeedsUpdate };
}

function headersMatch_(row, expected) {
  if (!row || !expected) return false;
  for (var i = 0; i < expected.length; i++) {
    if (String(row[i] || '').trim() !== expected[i]) return false;
  }
  return true;
}

/** 2列ガイド用：各行を必ず2列に揃える（setValues エラー防止） */
function normalizeGuideRowsTo2Cols_(rows) {
  return rows.map(function(row) {
    row = row || [];
    return [
      row[0] != null ? String(row[0]) : '',
      row[1] != null ? String(row[1]) : ''
    ];
  });
}

/**
 * セットアップ手順シート
 */
function writeSetupGuideSheet_(ss, forceRewrite) {
  var name = 'セットアップ手順';
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  } else if (!forceRewrite && sheet.getLastRow() > 3) {
    return;
  }

  sheet.clear();
  var guide = [
    ['製造日報アプリ｜セットアップ手順'],
    [],
    ['手順', '内容'],
    ['1', 'メニュー「製造日報」→「全シート＋ヘッダーを一括作成」を実行'],
    ['2', '作業マスタ: A列QRコードをQR化（E列QR表示に数式自動設定）。製造①②は登録しない'],
    ['3', '点検マスタ（日常/センサ/定時/切替）にライン・項目・実施区分・改定日・アクションを登録（詳細は「点検マスタ運用」シート）'],
    ['4', '設定.gs の MASTER_SS_ID を共通マスタのIDに設定'],
    ['5', 'メニュー「共通マスタを初期化」→ 社員マスタ（EMP001形式+QR表示）等'],
    ['6', 'Apps Script を Webアプリとしてデプロイ'],
    [],
    ['シート（日報用ブック）', '用途'],
    ['製造実績', '日別の製造・ケース・不良・作業者'],
    ['タイムライン', '日別の作業区分と時間帯（修理/調整/製造①…）'],
    ['点検実績', '日常点検・センサ・定時の入力結果（スロット別）'],
    ['作業マスタ', '固定作業のQR・表示名・色。QRはA列のみ'],
    ['日常点検マスタ', 'ライン別の日常点検項目'],
    ['センサチェックマスタ', '午前/午後/品切 の実施区分'],
    ['定時検査マスタ', '始業/10時/13時等の実施区分'],
    ['切替点検マスタ', '品種切替時の点検項目（ライン別）'],
    ['日報集計', 'レポート出力用（自動生成）'],
    [],
    ['点検マスタ 改定日ルール', ''],
    ['改定日', 'この行の変更が有効になる日（yyyy-MM-dd）'],
    ['アクション「追加」', '改定日以降の作業日に表示（改定日より前は非表示）'],
    ['アクション「廃止」', '改定日より前の作業日にのみ表示（改定日以降は非表示）'],
    ['アクション「変更」', '新項目行。旧項目名に改名前の名称を入力（旧行は別途「廃止」行を追加）'],
    ['旧項目名', '変更時のみ。点検実績の旧名称と紐づけ'],
    ['既存行の移行', '改定日=2000-01-01、アクション=追加 で全過去日有効'],
    [],
    ['作業マスタ QRルール', ''],
    ['QRにする列', 'A列「QRコード」（例: WORK_REPAIR）'],
    ['QRにしない列', 'B列「作業区分名」（修理・調整など表示用）'],
    ['製造①②', '品種スキャンで自動追加。マスタに書かない'],
    ['作業QRスキャン', '前作業終了＝読取時刻、新作業開始＝読取時刻（モーダルで修正可）'],
    [],
    ['共通マスタ（別ブック）', ''],
    ['作業者マスタ', '（旧）QRコード / 作業者名 → 社員マスタへ統合済み'],
    ['社員マスタ', '社員ID(QR) / 氏名 / Email / 権限 / QR表示（製造日報・出張旅費精算共用）'],
    ['不良マスタ', 'QRコード / 不良名 / QR表示'],
    ['ラインマスタ', 'ライン名一覧'],
    ['品目', '任意シート名で品目CD・品名・入数']
  ];

  sheet.getRange(1, 1, guide.length, 2).setValues(normalizeGuideRowsTo2Cols_(guide));
  sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setFontSize(12);
  sheet.setColumnWidths(1, 1, 200);
  sheet.setColumnWidths(2, 1, 420);
  sheet.setFrozenRows(1);
}

var CHECK_MASTER_GUIDE_SHEET_NAME = '点検マスタ運用';

/**
 * 点検マスタ運用説明シート
 */
function writeCheckMasterGuideSheet_(ss, forceRewrite) {
  var sheet = ss.getSheetByName(CHECK_MASTER_GUIDE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CHECK_MASTER_GUIDE_SHEET_NAME);
  } else if (!forceRewrite && sheet.getLastRow() > 3) {
    return;
  }

  sheet.clear();
  var guide = [
    ['製造日報アプリ｜点検マスタ運用説明'],
    [],
    ['概要', '日常点検・センサチェック・定時検査・切替点検の項目は、日報用スプレッドシートの各点検マスタで管理します。Webアプリは作業日ごとに有効な行だけを読み込み、点検実績とマージして表示します。'],
    [],
    ['対象シート', '列構成'],
    ['日常点検マスタ', 'ライン / 項目 / 改定日 / アクション / 旧項目名'],
    ['切替点検マスタ', 'ライン / 項目 / 改定日 / アクション / 旧項目名'],
    ['センサチェックマスタ', 'ライン / 項目 / 実施区分 / 改定日 / アクション / 旧項目名'],
    ['定時検査マスタ', 'ライン / 項目 / 判定種別 / 実施区分 / 改定日 / アクション / 旧項目名'],
    ['実施区分', 'センサ: 午前・午後・品切 等（カンマ区切り可）。定時: 始業・10時・13時 等。空欄=全スロット'],
    ['ライン', 'Webアプリのライン選択と同じ表記（例: 1号ライン）'],
    [],
    ['改定日・アクション・旧項目名', ''],
    ['改定日', 'この行の変更が有効になる日（yyyy-MM-dd）。空欄は 2000-01-01 扱い'],
    ['アクション「追加」', '改定日以降の作業日に表示。改定日より前の作業日では非表示'],
    ['アクション「廃止」', '改定日より前の作業日にのみ表示。改定日以降の作業日では非表示'],
    ['アクション「変更」', '新しい項目名の行。旧項目名に改名前の名称を入力する'],
    ['旧項目名', '変更時のみ入力。保存済み点検実績の旧名称と紐づける'],
    ['変更時の注意', '名称変更は「変更」行＋旧名称の「廃止」行の2行セットで登録する'],
    [],
    ['運用例：新規追加', ''],
    ['例（日常点検）', '1号ライン | 潤滑油レベル確認 | 2026-04-01 | 追加 | （空）'],
    ['効果', '2026-04-01 以降の作業日に「潤滑油レベル確認」が表示される'],
    [],
    ['運用例：廃止', ''],
    ['例（日常点検）', '1号ライン | 旧チェック項目 | 2026-04-01 | 廃止 | （空）'],
    ['効果', '2026-03-31 以前の作業日では表示、2026-04-01 以降は非表示'],
    [],
    ['運用例：名称変更', ''],
    ['変更行', '1号ライン | 新チェック項目 | 2026-04-01 | 変更 | 旧チェック項目'],
    ['廃止行', '1号ライン | 旧チェック項目 | 2026-04-01 | 廃止 | （空）'],
    ['効果', '2026-04-01 以降は新名称。それ以前は旧名称。過去の点検実績は旧名称で紐づく'],
    [],
    ['既存行の移行', ''],
    ['推奨設定', '改定日=2000-01-01、アクション=追加、旧項目名=空'],
    ['意味', '初回導入時の既存項目を、すべての過去日・未来日で有効にする'],
    [],
    ['過去データの扱い', ''],
    ['旧バッジ', 'マスタに無いが点検実績に残る項目は、Webアプリで「旧」と表示される'],
    ['過去日の閲覧', '本日以外の作業日を開くと、点検入力は読み取り専用（変更不可）'],
    ['過去日の保存', '保存時に確認ダイアログ。現在のマスタ構成で点検実績が上書きされる'],
    [],
    ['運用ツール', ''],
    ['入力チェック', 'メニュー「製造日報」→「点検マスタ改定日の入力チェック」'],
    ['説明の更新', 'メニュー「製造日報」→「点検マスタ運用説明を更新」'],
    [],
    ['よくあるミス', ''],
    ['変更だけ追加', '名称変更時に旧名称の「廃止」行を忘れると、新旧両方が表示される場合がある'],
    ['ライン名不一致', 'マスタのライン名がWebアプリの選択肢と1文字でも違うと項目が出ない'],
    ['改定日の形式', 'yyyy-MM-dd 形式（例: 2026-04-01）。日付として解釈できない値は避ける'],
    ['変更で旧項目名が空', '入力チェックメニューで警告される。必ず旧名称を入力する']
  ];

  sheet.getRange(1, 1, guide.length, 2).setValues(normalizeGuideRowsTo2Cols_(guide));
  sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setFontSize(12);
  sheet.setColumnWidths(1, 1, 220);
  sheet.setColumnWidths(2, 1, 520);
  sheet.setFrozenRows(1);
  try { sheet.setTabColor('#dbeafe'); } catch (e) { /* ignore */ }
}

/**
 * Webアプリから点検マスタ運用シートへのURL
 */
function getCheckMasterGuideUrl_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CHECK_MASTER_GUIDE_SHEET_NAME);
  if (sheet) return ss.getUrl() + '#gid=' + sheet.getSheetId();
  return ss.getUrl();
}

function menuRefreshCheckMasterGuide() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  writeCheckMasterGuideSheet_(ss, true);
  ui.alert(
    '点検マスタ運用',
    '「' + CHECK_MASTER_GUIDE_SHEET_NAME + '」シートを更新しました。\n\n'
      + 'シートタブから内容を確認してください。\n'
      + 'Webアプリの点検タブにも運用案内リンクが表示されます（再デプロイ後）。',
    ui.ButtonSet.OK
  );
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('製造日報')
    .addItem('全シート＋ヘッダーを一括作成', 'menuInitializeSpreadsheet')
    .addItem('ヘッダーのみ再設定（データ保持）', 'menuResetHeaders')
    .addSeparator()
    .addItem('切替点検マスタシートを作成', 'menuEnsureSwitchoverMasterSheet')
    .addItem('切替点検マスタの読込確認', 'menuCheckSwitchoverMaster')
    .addItem('点検マスタ改定日の入力チェック', 'menuValidateCheckMasterRevisions')
    .addItem('点検マスタ運用説明を更新', 'menuRefreshCheckMasterGuide')
    .addSeparator()
    .addItem('作業マスタのQR表示を再生成', 'menuRefreshWorkMasterQr')
    .addSeparator()
    .addItem('共通マスタを初期化（別ブック）', 'menuInitializeMaster')
    .addItem('作業者マスタを社員マスタへ統合', 'menuMergeWorkerMaster')
    .addToUi();
}

/**
 * 「切替点検マスタ」シートだけを確実に用意する（既存データは保持）
 */
function menuEnsureSwitchoverMasterSheet() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var spec = {
    name: '切替点検マスタ',
    headers: ['ライン', '項目', '改定日', 'アクション', '旧項目名'],
    tabColor: '#fbcfe8',
    samples: null
  };
  var result = ensureSheetWithHeaders_(ss, spec, false);
  var sheet = ss.getSheetByName(spec.name);
  if (sheet) {
    try { sheet.setTabColor(spec.tabColor); } catch (e) { /* ignore */ }
  }
  var msg = result.created
    ? '「切替点検マスタ」を新規作成しました。'
    : '「切替点検マスタ」は既にあります。';
  msg += '\n\nA1=ライン / B1=項目 / C1=改定日 / D1=アクション / E1=旧項目名\n2行目以降に、画面上のライン名と同じ表記で登録してください。\n';
  msg += '（例: 1号ライン）\n\n登録後は clasp push → clasp deploy 済みの Webアプリを再読み込みしてください。';
  ui.alert('切替点検マスタ', msg, ui.ButtonSet.OK);
}

/**
 * 切替点検マスタが読めるか確認（トラブルシュート用）
 */
function menuCheckSwitchoverMaster() {
  var ui = SpreadsheetApp.getUi();
  var master = loadCheckMasterFromSheet(new Date());
  var lines = Object.keys(master).filter(function(line) {
    return master[line].switchoverMasterItems && master[line].switchoverMasterItems.length > 0;
  });
  if (lines.length === 0) {
    ui.alert(
      '切替点検マスタ',
      '読み込める項目がありません。\n\n'
        + '・シート名が「切替点検マスタ」か\n'
        + '・1行目が「ライン」「項目」か\n'
        + '・2行目以降にデータがあるか\n'
        + '・A列のライン名が Webアプリのライン選択と一致しているか\n\n'
        + 'を確認してください。',
      ui.ButtonSet.OK
    );
    return;
  }
  var detail = lines.map(function(line) {
    return '・' + line + ' … ' + master[line].switchoverMasterItems.length + ' 項目';
  }).join('\n');
  ui.alert('切替点検マスタ', '読込 OK:\n' + detail, ui.ButtonSet.OK);
}

function menuValidateCheckMasterRevisions() {
  var ui = SpreadsheetApp.getUi();
  var warnings = validateCheckMasterRevisions_();
  if (!warnings.length) {
    ui.alert('点検マスタ改定日', '入力チェック: 問題は見つかりませんでした。', ui.ButtonSet.OK);
    return;
  }
  ui.alert('点検マスタ改定日', '確認事項:\n\n・' + warnings.join('\n・'), ui.ButtonSet.OK);
}

function menuRefreshWorkMasterQr() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('作業マスタ');
  if (!sheet) {
    SpreadsheetApp.getUi().alert('作業マスタシートがありません。先に一括作成を実行してください。');
    return;
  }
  applyQrDisplayFormulas_(sheet, 1, 5);
  clearMasterCaches_();
  SpreadsheetApp.getUi().alert('作業マスタの E列（QR表示）を再生成しました。');
}

function menuInitializeSpreadsheet() {
  var ui = SpreadsheetApp.getUi();
  var confirm = ui.alert(
    'スプレッドシート初期化',
    '必要なシートとヘッダーを作成します。\n作業マスタには QR表示列の数式も設定します。\n既存データ行は削除しません（新規シート時はサンプル行あり）。\n\n実行しますか？',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;
  var msg = initializeSpreadsheet({ addSamples: true, forceHeaders: false });
  clearMasterCaches_();
  ui.alert('完了', msg, ui.ButtonSet.OK);
}

function menuResetHeaders() {
  var ui = SpreadsheetApp.getUi();
  var confirm = ui.alert(
    'ヘッダー再設定',
    '各シートの1行目を標準ヘッダーで上書きします。\n2行目以降のデータは残ります。\n\n実行しますか？',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;
  var msg = resetAllSheetHeaders();
  ui.alert('完了', msg, ui.ButtonSet.OK);
}

function menuInitializeMaster() {
  var ui = SpreadsheetApp.getUi();
  var confirm = ui.alert(
    '共通マスタ初期化',
    '別ブックにラインマスタ・社員マスタ（EMP形式）・不良マスタを作成し、QR表示数式を設定します。\n\n実行しますか？',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;
  try {
    var msg = initializeMasterSpreadsheet({ addSamples: true, forceHeaders: false });
    clearMasterCaches_();
    ui.alert('完了', msg, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('エラー', e.message, ui.ButtonSet.OK);
  }
}

/**
 * 旧「作業者マスタ」の内容を「社員マスタ」へ統合（出張旅費精算アプリと同じロジック）
 */
function mergeWorkerMasterIntoEmployeeMaster() {
  if (!MASTER_SS_ID) {
    throw new Error('設定.js の MASTER_SS_ID が未設定です。');
  }

  var ss = SpreadsheetApp.openById(MASTER_SS_ID);
  ensureEmployeeMasterSheetInCommon_(ss, false);

  var empSheet = ss.getSheetByName(EMPLOYEE_MASTER_SHEET_NAME);
  var workerSheet = ss.getSheetByName(WORKER_MASTER_LEGACY_SHEET);
  if (!workerSheet || workerSheet.getLastRow() < 2) {
    applyQrDisplayFormulas_(empSheet, 1, 9);
    clearMasterCaches_();
    return '作業者マスタにデータがありません。\n社員マスタのみ QR 表示を更新しました。';
  }

  var wHeaders = workerSheet.getRange(1, 1, 1, workerSheet.getLastColumn()).getValues()[0]
    .map(function(h) { return String(h || '').trim(); });
  var wIdCol = findEmployeeColumnIndex_(wHeaders, ['社員ID', 'QRコード', 'ID']);
  var wNameCol = findEmployeeColumnIndex_(wHeaders, ['氏名', '作業者名', '名前']);
  var wRoleCol = findEmployeeColumnIndex_(wHeaders, ['権限']);
  if (wIdCol === -1 || wNameCol === -1) {
    throw new Error('作業者マスタの列（QRコード/社員ID、作業者名/氏名）が読み取れません。');
  }

  var empIndex = buildEmployeeIdIndex_(empSheet);
  var wData = workerSheet.getRange(2, 1, workerSheet.getLastRow(), workerSheet.getLastColumn()).getValues();
  var added = 0;
  var updated = 0;
  var skipped = 0;

  for (var i = 0; i < wData.length; i++) {
    var workerId = String(wData[i][wIdCol] || '').trim();
    var workerName = String(wData[i][wNameCol] || '').trim();
    var workerRole = wRoleCol !== -1 ? String(wData[i][wRoleCol] || '').trim() : '';
    if (!workerId || !workerName) {
      skipped++;
      continue;
    }

    if (empIndex[workerId]) {
      var rowNum = empIndex[workerId];
      empSheet.getRange(rowNum, 2).setValue(workerName);
      if (workerRole) empSheet.getRange(rowNum, 6).setValue(workerRole);
      if (!String(empSheet.getRange(rowNum, 8).getValue() || '').trim()) {
        empSheet.getRange(rowNum, 8).setValue('有効');
      }
      updated++;
    } else {
      empSheet.appendRow([workerId, workerName, '', '', '', workerRole, '', '有効', '']);
      empIndex[workerId] = empSheet.getLastRow();
      added++;
    }
  }

  applyQrDisplayFormulas_(empSheet, 1, 9);
  clearMasterCaches_();

  return '作業者マスタ → 社員マスタ の統合が完了しました。\n\n'
    + '【追加】' + added + ' 件\n'
    + '【更新】' + updated + ' 件\n'
    + '【スキップ】' + skipped + ' 件\n\n'
    + '※ Email・日当は社員マスタで個別に登録してください。\n'
    + '※ 旧「作業者マスタ」シートは手動で削除できます（任意）。';
}

function ensureEmployeeMasterSheetInCommon_(ss, forceHeaders) {
  var sheet = ss.getSheetByName(EMPLOYEE_MASTER_SHEET_NAME);
  var created = false;
  if (!sheet) {
    sheet = ss.insertSheet(EMPLOYEE_MASTER_SHEET_NAME);
    created = true;
  }

  var headerNeedsUpdate = created || forceHeaders || sheet.getLastRow() === 0 || !headersMatch_(
    sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), EMPLOYEE_MASTER_HEADERS.length)).getValues()[0],
    EMPLOYEE_MASTER_HEADERS
  );

  if (headerNeedsUpdate) {
    sheet.getRange(1, 1, 1, EMPLOYEE_MASTER_HEADERS.length).setValues([EMPLOYEE_MASTER_HEADERS]);
    sheet.getRange(1, 1, 1, EMPLOYEE_MASTER_HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#e2e8f0')
      .setWrap(true);
    sheet.setFrozenRows(1);
  }
  return { sheet: sheet, created: created };
}

function buildEmployeeIdIndex_(empSheet) {
  var index = {};
  if (!empSheet || empSheet.getLastRow() < 2) return index;
  var ids = empSheet.getRange(2, 1, empSheet.getLastRow(), 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    var id = String(ids[i][0] || '').trim();
    if (id) index[id] = i + 2;
  }
  return index;
}

function findEmployeeColumnIndex_(headers, aliases) {
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i] || '').replace(/\s/g, '');
    for (var j = 0; j < aliases.length; j++) {
      if (h === aliases[j] || h.indexOf(aliases[j]) !== -1) return i;
    }
  }
  return -1;
}

function menuMergeWorkerMaster() {
  var ui = SpreadsheetApp.getUi();
  var confirm = ui.alert(
    '作業者マスタ統合',
    '共通マスタの「作業者マスタ」の内容を「社員マスタ」へ取り込みます。\n'
      + '同じ社員IDは氏名・権限のみ更新し、Email 等は上書きしません。\n\n実行しますか？',
    ui.ButtonSet.YES_NO
  );
  if ( confirm !== ui.Button.YES) return;
  try {
    var msg = mergeWorkerMasterIntoEmployeeMaster();
    ui.alert('完了', msg, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('エラー', e.message, ui.ButtonSet.OK);
  }
}
