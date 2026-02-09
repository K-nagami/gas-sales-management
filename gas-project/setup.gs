/**
 * 売上管理統合システム - 初期セットアップ
 *
 * setup() を実行すると以下を自動構築:
 *   1. スプレッドシート「売上管理_統合システム」（4シート）
 *   2. Googleドキュメントテンプレート3種（templates.gs の関数を呼出）
 *   3. 設定シート初期値
 *   4. サンプルデータ投入
 *   5. ダッシュボード（関数・グラフ）
 *
 * 依存: templates.gs （createAllDocumentTemplates 関数）
 */

// ============================================================================
// メインセットアップ
// ============================================================================

/**
 * 売上管理統合システムの初期セットアップを実行
 */
function setup() {
  try {
    Logger.log('=== 売上管理統合システム セットアップ開始 ===');

    // --- STEP 1: スプレッドシート作成 ---
    Logger.log('STEP 1: スプレッドシート作成...');
    var ss = SpreadsheetApp.create('売上管理_統合システム');
    var ssId = ss.getId();

    // デフォルトシート (Sheet1) を削除して4シートを作成
    var defaultSheet = ss.getSheets()[0];
    ss.insertSheet('設定', 0);
    ss.insertSheet('案件マスタ', 1);
    ss.insertSheet('明細', 2);
    ss.insertSheet('ダッシュボード', 3);
    ss.deleteSheet(defaultSheet);
    Logger.log('  ✓ スプレッドシート: ' + ss.getUrl());

    // --- STEP 2: ドキュメントテンプレート作成 ---
    Logger.log('STEP 2: ドキュメントテンプレート作成...');
    var tplIds = createAllDocumentTemplates(); // templates.gs の関数
    Logger.log('  ✓ 見積書: ' + tplIds.estimateId);
    Logger.log('  ✓ 発注書: ' + tplIds.orderFormId);
    Logger.log('  ✓ 請求書: ' + tplIds.invoiceId);

    // --- STEP 3: 設定シート初期化 ---
    Logger.log('STEP 3: 設定シート初期化...');
    initSettingsSheet_(ss, tplIds);
    Logger.log('  ✓ 設定シート完了');

    // --- STEP 4: 案件マスタ初期化 ---
    Logger.log('STEP 4: 案件マスタシート...');
    initProjectMasterSheet_(ss);
    Logger.log('  ✓ 案件マスタ完了');

    // --- STEP 5: 明細シート初期化 ---
    Logger.log('STEP 5: 明細シート...');
    initLineItemsSheet_(ss);
    Logger.log('  ✓ 明細シート完了');

    // --- STEP 6: ダッシュボード構築 ---
    Logger.log('STEP 6: ダッシュボード...');
    initDashboardSheet_(ss);
    Logger.log('  ✓ ダッシュボード完了');

    // --- 完了報告 ---
    Logger.log('');
    Logger.log('=== セットアップ完了 ===');
    Logger.log('スプレッドシートURL: ' + ss.getUrl());
    Logger.log('見積書テンプレート: https://docs.google.com/document/d/' + tplIds.estimateId);
    Logger.log('発注書テンプレート: https://docs.google.com/document/d/' + tplIds.orderFormId);
    Logger.log('請求書テンプレート: https://docs.google.com/document/d/' + tplIds.invoiceId);

    // UIアラート（スプレッドシートにバインドされている場合）
    try {
      SpreadsheetApp.getUi().alert(
        'セットアップ完了',
        'スプレッドシートとテンプレートを作成しました。\n' +
        '「設定」シートで自社情報・帳票保存先フォルダIDを入力してください。',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } catch (uiError) {
      // スタンドアロンスクリプトの場合UIは使えない → ログのみ
    }

  } catch (error) {
    Logger.log('★ セットアップエラー: ' + error.message);
    Logger.log('スタックトレース: ' + error.stack);
    throw error;
  }
}

// ============================================================================
// 設定シート初期化
// ============================================================================

/**
 * @param {Spreadsheet} ss
 * @param {Object} tplIds { estimateId, orderFormId, invoiceId }
 * @private
 */
function initSettingsSheet_(ss, tplIds) {
  var sheet = ss.getSheetByName('設定');

  var data = [
    ['社名',                       '（後で入力）'],
    ['代表者名',                   '（後で入力）'],
    ['郵便番号',                   '〒000-0000'],
    ['住所',                       '（後で入力）'],
    ['TEL',                        '000-0000-0000'],
    ['メール',                     'example@example.com'],
    ['適格請求書発行事業者番号',    'T0000000000000'],
    ['振込先銀行',                 '（後で入力）'],
    ['振込先支店',                 '（後で入力）'],
    ['口座種別',                   '普通'],
    ['口座番号',                   '0000000'],
    ['口座名義',                   '（後で入力）'],
    ['消費税率',                   0.10],
    ['発注書注意書き',             '※本発注書にご署名・ご捺印の上、PDF又は原本をご返送ください。'],
    ['支払条件',                   '月末締め翌月末払い'],
    ['帳票保存先フォルダID',        '（GoogleドライブのフォルダIDを後で入力）'],
    ['見積書テンプレートID',        tplIds.estimateId],
    ['発注書テンプレートID',        tplIds.orderFormId],
    ['請求書テンプレートID',        tplIds.invoiceId],
    ['自社担当者名',                '（後で入力）'],
    ['自社URL',                     ''],
    ['社印画像ファイルID',          '']
  ];

  sheet.getRange(1, 1, data.length, 2).setValues(data);

  // 書式設定
  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 450);
  sheet.getRange('A1:A' + data.length).setFontWeight('bold');
  sheet.getRange('B13').setNumberFormat('0.00');
}

// ============================================================================
// 案件マスタシート初期化
// ============================================================================

/**
 * @param {Spreadsheet} ss
 * @private
 */
function initProjectMasterSheet_(ss) {
  var sheet = ss.getSheetByName('案件マスタ');

  // ヘッダー
  var headers = [
    '案件ID', '顧客名', '案件名', '見積日', '受注日',
    '請求日', '入金予定日', '入金日', 'ステータス',
    '見積金額（税抜）', '消費税', '合計金額',
    '見積書URL', '発注書URL', '請求書URL', '備考', 'ヨミランク',
    '登録日', '提案日', '失注日',
    '仕入先名', '仕入金額（税抜）', '仕入消費税', '仕入合計',
    '仕入先請求番号', '支払日', '仕入先請求書URL', '粗利', '粗利率'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4472C4').setFontColor('#FFFFFF');

  // サンプルデータ（数式は文字列で入力 → GASが自動的に数式として認識）
  var samples = [
    ['EST-202602-001', '株式会社サンプル商事', 'Google Workspace導入支援',
     new Date(2026, 1, 1), new Date(2026, 1, 5), '', '', '',
     '受注', '', '', '', '', '', '', '初期設定+研修込み', 'A',
     new Date(2026, 0, 25), new Date(2026, 0, 28), ''],
    ['EST-202602-002', '有限会社テスト工業', 'GA4分析レポート作成',
     new Date(2026, 1, 3), '', '', '', '',
     '見積もり提示', '', '', '', '', '', '', '月次レポート3ヶ月契約', 'B',
     new Date(2026, 1, 1), new Date(2026, 1, 2), ''],
    ['EST-202601-003', '合同会社デモ販売', 'SEOコンテンツ制作（10記事）',
     new Date(2026, 0, 15), new Date(2026, 0, 20), new Date(2026, 0, 31),
     new Date(2026, 1, 28), '',
     '請求済', '', '', '', '', '', '', 'キーワード選定済み', 'A',
     new Date(2026, 0, 10), new Date(2026, 0, 12), '']
  ];
  sheet.getRange(2, 1, samples.length, samples[0].length).setValues(samples);

  // J/K/L列の数式をセット（行2〜4）
  for (var r = 2; r <= 4; r++) {
    sheet.getRange(r, 10).setFormula('=SUMIF(明細!A:A,A' + r + ',明細!E:E)');           // J: 見積金額（税抜）
    sheet.getRange(r, 11).setFormula('=FLOOR(J' + r + '*設定!B13)');                     // K: 消費税（切り捨て）
    sheet.getRange(r, 12).setFormula('=J' + r + '+K' + r);                              // L: 合計金額
    // 仕入関連数式
    sheet.getRange(r, 23).setFormula('=FLOOR(V' + r + '*設定!B13)');                     // W: 仕入消費税
    sheet.getRange(r, 24).setFormula('=V' + r + '+W' + r);                              // X: 仕入合計
    sheet.getRange(r, 28).setFormula('=L' + r + '-X' + r);                              // AB: 粗利
    sheet.getRange(r, 29).setFormula('=IFERROR(AB' + r + '/L' + r + ',0)');             // AC: 粗利率
  }

  // Q列幅
  sheet.setColumnWidth(17, 90);

  // ステータス列のプルダウン（v2: 7段階）
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['商談見込み', '提案中', '見積もり提示', '受注', '請求済', '入金済', '失注'])
    .setAllowInvalid(false)
    .build();
  sheet.getRange('I2:I1000').setDataValidation(statusRule);

  // ヨミランク列（Q列）のプルダウン
  var yomiRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['A', 'B', 'C', 'D'])
    .setAllowInvalid(true)
    .build();
  sheet.getRange('Q2:Q1000').setDataValidation(yomiRule);

  // 書式
  sheet.getRange('D:H').setNumberFormat('yyyy/mm/dd');
  sheet.getRange('R:T').setNumberFormat('yyyy/mm/dd');
  sheet.getRange('Z:Z').setNumberFormat('yyyy/mm/dd');
  sheet.getRange('J:L').setNumberFormat('#,##0');
  sheet.getRange('V:X').setNumberFormat('#,##0');
  sheet.getRange('AB:AB').setNumberFormat('#,##0');
  sheet.getRange('AC:AC').setNumberFormat('0.0%');

  // 列幅
  var widths = [130, 160, 200, 100, 100, 100, 100, 100, 90, 120, 100, 120, 200, 200, 200, 180, 90, 100, 100, 100,
    140, 120, 100, 120, 140, 100, 200, 100, 80];
  for (var c = 0; c < widths.length; c++) {
    sheet.setColumnWidth(c + 1, widths[c]);
  }
}

// ============================================================================
// 明細シート初期化
// ============================================================================

/**
 * @param {Spreadsheet} ss
 * @private
 */
function initLineItemsSheet_(ss) {
  var sheet = ss.getSheetByName('明細');

  // ヘッダー
  var headers = ['案件ID', '品目名', '単価', '数量', '小計'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4472C4').setFontColor('#FFFFFF');

  // サンプルデータ
  var items = [
    ['EST-202602-001', 'Google Workspace Business Standard 初期設定', 80000, 1],
    ['EST-202602-001', '管理者向け操作研修（半日）',                   40000, 1],
    ['EST-202602-001', 'ユーザー向け操作マニュアル作成',               30000, 1],
    ['EST-202602-002', 'GA4現状分析',                                  50000, 1],
    ['EST-202602-002', '月次分析レポート作成',                          30000, 3],
    ['EST-202601-003', 'SEO記事制作（3000字〜）',                      25000, 10],
    ['EST-202601-003', 'キーワード選定・構成案作成',                    15000, 1]
  ];
  sheet.getRange(2, 1, items.length, 4).setValues(items);

  // E列（小計）の数式
  for (var r = 2; r <= items.length + 1; r++) {
    sheet.getRange(r, 5).setFormula('=C' + r + '*D' + r);
  }

  // 書式
  sheet.getRange('C:E').setNumberFormat('#,##0');
  sheet.setColumnWidth(1, 130);
  sheet.setColumnWidth(2, 320);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 80);
  sheet.setColumnWidth(5, 110);
}

// ============================================================================
// ダッシュボードシート初期化
// ============================================================================

/**
 * @param {Spreadsheet} ss
 * @private
 */
function initDashboardSheet_(ss) {
  var sheet = ss.getSheetByName('ダッシュボード');

  // ========== タイトル ==========
  sheet.getRange('A1').setValue('売上管理ダッシュボード');
  sheet.getRange('A1').setFontWeight('bold').setFontSize(16).setBackground('#4472C4').setFontColor('#FFFFFF');
  sheet.getRange('A1:G1').merge();

  // ========== セクション1: 月別ヨミサマリー (B2起点) ==========
  sheet.getRange('B2').setValue('■ 月別ヨミサマリー').setFontWeight('bold').setFontSize(13);

  var summaryHeaders = ['対象月', '商談〜見積', '受注', '請求済', '入金済', '合計'];
  sheet.getRange('B3:G3').setValues([summaryHeaders]);
  sheet.getRange('B3:G3').setFontWeight('bold').setBackground('#D9E2F3');

  // 当月〜6ヶ月先（指示書: 当月〜6ヶ月先）
  var now = new Date();
  for (var i = 0; i < 7; i++) {
    var r = 4 + i;
    var targetDate = new Date(now.getFullYear(), now.getMonth() + i, 1);
    var y = targetDate.getFullYear();
    var m = targetDate.getMonth() + 1; // 1-based
    var mStr = String(m).padStart(2, '0');
    var nextM = m + 1;
    var nextY = y;
    if (nextM > 12) { nextM = 1; nextY = y + 1; }

    sheet.getRange('B' + r).setValue(y + '-' + mStr);

    // 商談〜見積: 見積日(D列)が対象月内 かつ ステータス=商談見込み/提案中/見積もり提示
    sheet.getRange('C' + r).setFormula(
      '=SUMPRODUCT(((案件マスタ!I$2:I$1000="商談見込み")+(案件マスタ!I$2:I$1000="提案中")+(案件マスタ!I$2:I$1000="見積もり提示"))*(YEAR(案件マスタ!D$2:D$1000)=' + y + ')*(MONTH(案件マスタ!D$2:D$1000)=' + m + ')*(案件マスタ!J$2:J$1000))'
    );
    // 受注: 受注日(E列)基準
    sheet.getRange('D' + r).setFormula(
      '=SUMPRODUCT((案件マスタ!I$2:I$1000="受注")*(YEAR(案件マスタ!E$2:E$1000)=' + y + ')*(MONTH(案件マスタ!E$2:E$1000)=' + m + ')*(案件マスタ!J$2:J$1000))'
    );
    // 請求済: 請求日(F列)基準
    sheet.getRange('E' + r).setFormula(
      '=SUMPRODUCT((案件マスタ!I$2:I$1000="請求済")*(YEAR(案件マスタ!F$2:F$1000)=' + y + ')*(MONTH(案件マスタ!F$2:F$1000)=' + m + ')*(案件マスタ!J$2:J$1000))'
    );
    // 入金済: 入金日(H列)基準
    sheet.getRange('F' + r).setFormula(
      '=SUMPRODUCT((案件マスタ!I$2:I$1000="入金済")*(YEAR(案件マスタ!H$2:H$1000)=' + y + ')*(MONTH(案件マスタ!H$2:H$1000)=' + m + ')*(案件マスタ!J$2:J$1000))'
    );
    // 合計
    sheet.getRange('G' + r).setFormula('=SUM(C' + r + ':F' + r + ')');
  }
  sheet.getRange('C4:G10').setNumberFormat('#,##0');

  // ========== セクション2: KPI (B12起点) ==========
  sheet.getRange('B12').setValue('■ KPI').setFontWeight('bold').setFontSize(13);

  sheet.getRange('B13').setValue('当月受注率');
  sheet.getRange('C13').setFormula(
    '=IFERROR(COUNTIFS(案件マスタ!I:I,"受注",案件マスタ!E:E,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),案件マスタ!E:E,"<"&DATE(YEAR(TODAY()),MONTH(TODAY())+1,1))/' +
    'COUNTIFS(案件マスタ!D:D,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),案件マスタ!D:D,"<"&DATE(YEAR(TODAY()),MONTH(TODAY())+1,1)),0)'
  );
  sheet.getRange('C13').setNumberFormat('0.0%');

  sheet.getRange('B14').setValue('当月売上（請求済合計）');
  sheet.getRange('C14').setFormula(
    '=SUMIFS(案件マスタ!L:L,案件マスタ!I:I,"請求済",案件マスタ!F:F,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),案件マスタ!F:F,"<"&DATE(YEAR(TODAY()),MONTH(TODAY())+1,1))'
  );
  sheet.getRange('C14').setNumberFormat('#,##0');

  sheet.getRange('B15').setValue('当月入金済合計');
  sheet.getRange('C15').setFormula(
    '=SUMIFS(案件マスタ!L:L,案件マスタ!I:I,"入金済",案件マスタ!H:H,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),案件マスタ!H:H,"<"&DATE(YEAR(TODAY()),MONTH(TODAY())+1,1))'
  );
  sheet.getRange('C15').setNumberFormat('#,##0');

  sheet.getRange('B16').setValue('未入金残高');
  sheet.getRange('C16').setFormula(
    '=SUMIF(案件マスタ!I:I,"請求済",案件マスタ!L:L)'
  );
  sheet.getRange('C16').setNumberFormat('#,##0');

  sheet.getRange('B13:B16').setFontWeight('bold');

  // ========== セクション3: 入金遅延アラート (B18起点) ==========
  sheet.getRange('B18').setValue('■ 入金遅延アラート').setFontWeight('bold').setFontSize(13);

  var alertHeaders = ['案件ID', '顧客名', '案件名', '入金予定日', '合計金額'];
  sheet.getRange('B19:F19').setValues([alertHeaders]);
  sheet.getRange('B19:F19').setFontWeight('bold').setBackground('#F8D7DA');

  // FILTER関数で入金予定日 < TODAY() かつ 入金日が空の案件を表示
  sheet.getRange('B20').setFormula(
    '=IFERROR(FILTER({案件マスタ!A2:A,案件マスタ!B2:B,案件マスタ!C2:C,案件マスタ!G2:G,案件マスタ!L2:L},' +
    '(案件マスタ!G2:G<TODAY())*(案件マスタ!G2:G<>"")*(案件マスタ!H2:H="")*(案件マスタ!I2:I<>"入金済")),"該当なし")'
  );

  // ========== セクション4: 顧客別売上 (B25起点) ==========
  sheet.getRange('B25').setValue('■ 顧客別売上').setFontWeight('bold').setFontSize(13);

  var custHeaders = ['顧客名', '合計金額（税込）', '案件数'];
  sheet.getRange('B26:D26').setValues([custHeaders]);
  sheet.getRange('B26:D26').setFontWeight('bold').setBackground('#D9E2F3');

  // UNIQUE + SUMIFS で顧客名別集計
  sheet.getRange('B27').setFormula(
    '=IFERROR(SORT(UNIQUE(FILTER(案件マスタ!B2:B,案件マスタ!B2:B<>""))),"データなし")'
  );
  // C27以降は配列数式で一括計算
  sheet.getRange('C27').setFormula(
    '=IFERROR(ARRAYFORMULA(IF(B27:B40<>"",SUMIF(案件マスタ!B:B,B27:B40,案件マスタ!L:L),"")),"")'
  );
  sheet.getRange('D27').setFormula(
    '=IFERROR(ARRAYFORMULA(IF(B27:B40<>"",COUNTIF(案件マスタ!B:B,B27:B40),"")),"")'
  );
  sheet.getRange('C27:C40').setNumberFormat('#,##0');

  // ========== セクション5: 月別推移チャート用データ (B32起点) → B42起点に移動（顧客別と重ならないよう） ==========
  var chartStartRow = 42;
  sheet.getRange('B' + chartStartRow).setValue('■ 月別推移データ').setFontWeight('bold').setFontSize(13);

  var chartHeaders = ['対象月', '見積件数', '受注件数', '売上金額', '入金金額'];
  sheet.getRange('B' + (chartStartRow + 1) + ':F' + (chartStartRow + 1)).setValues([chartHeaders]);
  sheet.getRange('B' + (chartStartRow + 1) + ':F' + (chartStartRow + 1)).setFontWeight('bold').setBackground('#D9E2F3');

  var chartDataStart = chartStartRow + 2; // 44
  for (var j = 11; j >= 0; j--) {
    var cr = chartDataStart + (11 - j);
    var cd = new Date(now.getFullYear(), now.getMonth() - j, 1);
    var cy = cd.getFullYear();
    var cm = cd.getMonth() + 1;
    var cmStr = String(cm).padStart(2, '0');

    sheet.getRange('B' + cr).setValue(cy + '-' + cmStr);

    // 見積件数
    sheet.getRange('C' + cr).setFormula(
      '=COUNTIFS(案件マスタ!D:D,">="&DATE(' + cy + ',' + cm + ',1),案件マスタ!D:D,"<"&DATE(' + cy + ',' + (cm + 1) + ',1))'
    );
    // 受注件数
    sheet.getRange('D' + cr).setFormula(
      '=COUNTIFS(案件マスタ!E:E,">="&DATE(' + cy + ',' + cm + ',1),案件マスタ!E:E,"<"&DATE(' + cy + ',' + (cm + 1) + ',1))'
    );
    // 売上金額
    sheet.getRange('E' + cr).setFormula(
      '=SUMIFS(案件マスタ!L:L,案件マスタ!F:F,">="&DATE(' + cy + ',' + cm + ',1),案件マスタ!F:F,"<"&DATE(' + cy + ',' + (cm + 1) + ',1))'
    );
    // 入金金額
    sheet.getRange('F' + cr).setFormula(
      '=SUMIFS(案件マスタ!L:L,案件マスタ!H:H,">="&DATE(' + cy + ',' + cm + ',1),案件マスタ!H:H,"<"&DATE(' + cy + ',' + (cm + 1) + ',1))'
    );
  }

  var chartDataEnd = chartDataStart + 11;
  sheet.getRange('E' + chartDataStart + ':F' + chartDataEnd).setNumberFormat('#,##0');

  // ========== グラフ作成 ==========
  var chartRange = sheet.getRange('B' + (chartStartRow + 1) + ':F' + chartDataEnd);
  var chart = sheet.newChart()
    .setChartType(Charts.ChartType.COMBO)
    .addRange(chartRange)
    .setOption('title', '月別推移')
    .setOption('legend', { position: 'bottom' })
    .setOption('hAxis', { title: '月' })
    .setOption('series', {
      0: { type: 'bars', color: '#4472C4' },    // 見積件数
      1: { type: 'bars', color: '#70AD47' },     // 受注件数
      2: { type: 'line', color: '#FF0000', targetAxisIndex: 1 }, // 売上金額
      3: { type: 'line', color: '#FFC000', targetAxisIndex: 1 }  // 入金金額
    })
    .setOption('vAxes', {
      0: { title: '件数' },
      1: { title: '金額' }
    })
    .setPosition(chartDataEnd + 2, 2, 0, 0)
    .build();
  sheet.insertChart(chart);

  // ========== 列幅 ==========
  sheet.setColumnWidth(1, 20);
  sheet.setColumnWidth(2, 180);
  for (var w = 3; w <= 7; w++) {
    sheet.setColumnWidth(w, 130);
  }
}

// ============================================================================
// v2マイグレーション（既存ユーザー向け）
// ============================================================================

/**
 * 既存のスプレッドシートをv2に移行する
 *
 * 実行内容:
 *   1. 案件マスタ Q1 にヘッダー「ヨミランク」追加
 *   2. Q列にA/B/C/Dのプルダウン設定
 *   3. I列のプルダウンを新ステータス7種に変更
 *   4. E1ヘッダーを「受注日」に変更
 *   5. 旧ステータスを新ステータスに一括変換
 */
function migrateV2() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('案件マスタ');
    if (!sheet) throw new Error('「案件マスタ」シートが見つかりません');

    Logger.log('=== v2 マイグレーション開始 ===');

    // --- 1. Q1ヘッダー ---
    sheet.getRange('Q1').setValue('ヨミランク');
    sheet.getRange('Q1').setFontWeight('bold').setBackground('#4472C4').setFontColor('#FFFFFF');
    sheet.setColumnWidth(17, 90);
    Logger.log('  ✓ Q1 ヘッダー追加');

    // --- 2. Q列プルダウン ---
    var yomiRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['A', 'B', 'C', 'D'])
      .setAllowInvalid(true)
      .build();
    sheet.getRange('Q2:Q1000').setDataValidation(yomiRule);
    Logger.log('  ✓ Q列 プルダウン設定');

    // --- 3. I列プルダウン更新 ---
    var statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['商談見込み', '提案中', '見積もり提示', '受注', '請求済', '入金済', '失注'])
      .setAllowInvalid(false)
      .build();
    sheet.getRange('I2:I1000').setDataValidation(statusRule);
    Logger.log('  ✓ I列 プルダウン更新（7段階）');

    // --- 4. E1ヘッダー変更 ---
    sheet.getRange('E1').setValue('受注日');
    Logger.log('  ✓ E1 ヘッダー「受注日」に変更');

    // --- 5. ステータス一括変換 ---
    var data = sheet.getDataRange().getValues();
    var statusMap = {
      '見積中': '見積もり提示',
      '発注済': '受注'
      // 請求済, 入金済, 失注 はそのまま
    };
    var converted = 0;

    for (var i = 1; i < data.length; i++) {
      var oldStatus = data[i][8]; // I列
      if (oldStatus && statusMap[oldStatus]) {
        sheet.getRange(i + 1, 9).setValue(statusMap[oldStatus]);
        converted++;
      }
    }
    Logger.log('  ✓ ステータス変換: ' + converted + '件');

    Logger.log('');
    Logger.log('=== v2 マイグレーション完了 ===');
    Logger.log('変換ルール: 見積中→見積もり提示 / 発注済→受注');

    try {
      SpreadsheetApp.getUi().alert(
        'v2マイグレーション完了',
        'ステータス ' + converted + '件を変換しました。\n' +
        '見積中→見積もり提示 / 発注済→受注\n\n' +
        'ヨミランク(Q列)は各案件に手動でA/B/C/Dを設定してください。',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } catch (uiErr) { /* スタンドアロンの場合 */ }

  } catch (error) {
    Logger.log('★ マイグレーションエラー: ' + error.message);
    throw error;
  }
}

// ============================================================================
// v3 マイグレーション（R/S/T列追加: 登録日・提案日・失注日）
// ============================================================================

/**
 * 既存スプレッドシートにR列（登録日）/S列（提案日）/T列（失注日）を追加
 * 既存データの登録日にはD列（見積日）の値をコピー
 * 提案中ステータスの案件にはD列の値を提案日にもコピー
 * 失注ステータスの案件にはD列の値を失注日にもコピー
 *
 * ★ 実行場所: スプレッドシート → 拡張機能 → Apps Script → migrateV3 を実行
 */
function migrateV3() {
  try {
    Logger.log('=== v3 マイグレーション開始 ===');
    Logger.log('追加列: R(登録日) / S(提案日) / T(失注日)');

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('案件マスタ');
    if (!sheet) throw new Error('「案件マスタ」シートが見つかりません');

    // --- 1. R1/S1/T1 ヘッダー追加 ---
    var newHeaders = [['登録日', '提案日', '失注日']];
    sheet.getRange('R1:T1').setValues(newHeaders);
    sheet.getRange('R1:T1').setFontWeight('bold').setBackground('#4472C4').setFontColor('#FFFFFF');
    Logger.log('  ✓ R/S/T列 ヘッダー追加');

    // --- 2. R:T列の書式を日付に設定 ---
    sheet.getRange('R:T').setNumberFormat('yyyy/mm/dd');
    Logger.log('  ✓ R:T列 日付書式設定');

    // --- 3. 列幅設定 ---
    sheet.setColumnWidth(18, 100); // R列
    sheet.setColumnWidth(19, 100); // S列
    sheet.setColumnWidth(20, 100); // T列
    Logger.log('  ✓ R/S/T列 列幅設定');

    // --- 4. 既存データのマイグレーション ---
    var data = sheet.getDataRange().getValues();
    var migrated = 0;
    var proposalSet = 0;
    var lostSet = 0;

    for (var i = 1; i < data.length; i++) {
      var row = i + 1;
      var quoteDate = data[i][3];   // D列: 見積日
      var status = data[i][8];      // I列: ステータス
      var existingR = data[i][17];  // R列: 登録日（既存）

      if (!data[i][0]) continue; // 案件IDがなければスキップ

      // R列（登録日）: 空の場合、D列の値をコピー
      if (!existingR && quoteDate) {
        sheet.getRange(row, 18).setValue(new Date(quoteDate));
        migrated++;
      }

      // S列（提案日）: 提案中以降のステータスでS列が空の場合
      var proposalAndAfter = ['提案中', '見積もり提示', '受注', '請求済', '入金済'];
      if (!data[i][18] && proposalAndAfter.indexOf(status) !== -1 && quoteDate) {
        sheet.getRange(row, 19).setValue(new Date(quoteDate));
        proposalSet++;
      }

      // T列（失注日）: 失注ステータスでT列が空の場合
      if (!data[i][19] && status === '失注' && quoteDate) {
        sheet.getRange(row, 20).setValue(new Date(quoteDate));
        lostSet++;
      }
    }

    Logger.log('  ✓ 登録日マイグレーション: ' + migrated + '件');
    Logger.log('  ✓ 提案日マイグレーション: ' + proposalSet + '件');
    Logger.log('  ✓ 失注日マイグレーション: ' + lostSet + '件');

    Logger.log('');
    Logger.log('=== v3 マイグレーション完了 ===');
    Logger.log('※ マイグレーション値は見積日(D列)からの推定です');
    Logger.log('※ 今後はステータス変更時に自動入力されます');

    try {
      SpreadsheetApp.getUi().alert(
        'v3マイグレーション完了',
        'R/S/T列（登録日・提案日・失注日）を追加しました。\n\n' +
        '登録日: ' + migrated + '件（見積日から推定）\n' +
        '提案日: ' + proposalSet + '件（見積日から推定）\n' +
        '失注日: ' + lostSet + '件（見積日から推定）\n\n' +
        '※ 今後はステータス変更時に自動入力されます。\n' +
        '※ Web Appを再デプロイしてください。',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } catch (uiErr) { /* スタンドアロンの場合 */ }

  } catch (error) {
    Logger.log('★ v3マイグレーションエラー: ' + error.message);
    throw error;
  }
}

// ============================================================================
// v4 マイグレーション（顧客マスタ新設・設定拡張・テンプレートリデザイン）
// ============================================================================

/**
 * 既存スプレッドシートをv4に移行する
 *
 * 実行内容:
 *   1. 顧客マスタシート作成
 *   2. 設定シートにB20-B22を追加
 *   3. テンプレート3種を再生成 → 設定シートB17-B19を更新
 *
 * ★ 実行場所: スプレッドシート → 拡張機能 → Apps Script → migrateV4 を実行
 */
function migrateV4() {
  try {
    Logger.log('=== v4 マイグレーション開始 ===');
    Logger.log('変更: 顧客マスタ新設 / 設定拡張(B20-B22) / テンプレートリデザイン');

    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // --- 1. 顧客マスタシート作成 ---
    createCustomerMasterSheet_(ss);
    Logger.log('  ✓ 顧客マスタシート作成完了');

    // --- 2. 設定シート拡張 ---
    extendSettingsForV4_(ss);
    Logger.log('  ✓ 設定シート拡張完了 (B20-B22)');

    // --- 3. テンプレート再生成 ---
    Logger.log('  テンプレート再生成中...');
    var tplIds = createAllDocumentTemplates();
    var settingsSheet = ss.getSheetByName('設定');
    settingsSheet.getRange('B17').setValue(tplIds.estimateId);
    settingsSheet.getRange('B18').setValue(tplIds.orderFormId);
    settingsSheet.getRange('B19').setValue(tplIds.invoiceId);
    Logger.log('  ✓ テンプレート再生成完了');
    Logger.log('    見積書: ' + tplIds.estimateId);
    Logger.log('    発注書: ' + tplIds.orderFormId);
    Logger.log('    請求書: ' + tplIds.invoiceId);

    Logger.log('');
    Logger.log('=== v4 マイグレーション完了 ===');

    try {
      SpreadsheetApp.getUi().alert(
        'v4マイグレーション完了',
        '以下を実行しました:\n\n' +
        '1. 顧客マスタシートを作成\n' +
        '2. 設定シートにB20-B22を追加\n' +
        '3. テンプレート3種を再生成\n\n' +
        '★ 次のステップ:\n' +
        '  - 顧客マスタに顧客情報（住所・担当者等）を入力\n' +
        '  - 設定シートB20(自社担当者名)、B21(自社URL)、B22(社印画像ID)を入力\n' +
        '  - Web Appを再デプロイ',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } catch (uiErr) { /* スタンドアロンの場合 */ }

  } catch (error) {
    Logger.log('★ v4マイグレーションエラー: ' + error.message);
    Logger.log('スタックトレース: ' + error.stack);
    throw error;
  }
}

/**
 * 顧客マスタシートを作成
 * 既存の案件マスタB列から顧客名を抽出してプレースホルダー行を作成
 * @param {Spreadsheet} ss
 * @private
 */
function createCustomerMasterSheet_(ss) {
  // 既存シートがあればスキップ
  if (ss.getSheetByName('顧客マスタ')) {
    Logger.log('  顧客マスタシートは既に存在します（スキップ）');
    return;
  }

  var sheet = ss.insertSheet('顧客マスタ');

  // ヘッダー（v4b: 呼称+名称分離）
  var headers = ['呼称', '名称', '郵便番号', '住所1', '住所2(ビル名)', '役職', '担当者名', 'メールアドレス', '備考'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4472C4').setFontColor('#FFFFFF');

  // 列幅設定
  var widths = [180, 200, 120, 250, 200, 100, 120, 200, 200];
  for (var c = 0; c < widths.length; c++) {
    sheet.setColumnWidth(c + 1, widths[c]);
  }

  // 案件マスタB列から既存顧客名を抽出
  var masterSheet = ss.getSheetByName('案件マスタ');
  if (masterSheet) {
    var masterData = masterSheet.getDataRange().getValues();
    var customerNames = {};
    for (var i = 1; i < masterData.length; i++) {
      var name = masterData[i][1];
      if (name && name !== '' && !customerNames[name]) {
        customerNames[name] = true;
      }
    }

    // プレースホルダー行を作成（呼称=名称で初期セット）
    var names = Object.keys(customerNames);
    if (names.length > 0) {
      var rows = [];
      for (var j = 0; j < names.length; j++) {
        rows.push([names[j], names[j], '', '', '', '', '', '', '']);
      }
      sheet.getRange(2, 1, rows.length, 9).setValues(rows);
      Logger.log('  顧客名 ' + names.length + '件をプレースホルダーとして追加');
    }
  }
}

/**
 * 設定シートにB20-B22を追加（未存在時のみ）
 * @param {Spreadsheet} ss
 * @private
 */
function extendSettingsForV4_(ss) {
  var sheet = ss.getSheetByName('設定');
  if (!sheet) throw new Error('「設定」シートが見つかりません');

  // B20が空かチェック
  var existingA20 = sheet.getRange('A20').getValue();
  if (existingA20 === '自社担当者名') {
    Logger.log('  設定B20-B22は既に存在します（スキップ）');
    return;
  }

  var newData = [
    ['自社担当者名',       '（後で入力）'],
    ['自社URL',            ''],
    ['社印画像ファイルID', '']
  ];
  sheet.getRange(20, 1, 3, 2).setValues(newData);
  sheet.getRange('A20:A22').setFontWeight('bold');
}

// ============================================================================
// v4b マイグレーション（顧客マスタ: 呼称/名称分離）
// ============================================================================

/**
 * 顧客マスタに「名称」列（B列）を追加し、既存データを1列右にシフト
 *
 * 変更内容:
 *   1. 顧客マスタ B列に「名称」を挿入（既存列を右にシフト）
 *   2. A1=「呼称」、B1=「名称」にヘッダー更新
 *   3. 既存の呼称（A列）をB列（名称）にもコピー（初期値として同一にする）
 *
 * ★ 実行場所: スプレッドシート → 拡張機能 → Apps Script → migrateV4b を実行
 */
function migrateV4b() {
  try {
    Logger.log('=== v4b マイグレーション開始 ===');
    Logger.log('変更: 顧客マスタに名称列を追加（呼称/名称分離）');

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('顧客マスタ');

    if (!sheet) {
      // 顧客マスタがなければ新スキーマで作成
      Logger.log('  顧客マスタが存在しません。新スキーマで作成します。');
      createCustomerMasterSheet_(ss);
      Logger.log('=== v4b マイグレーション完了（新規作成）===');
      try {
        SpreadsheetApp.getUi().alert(
          'v4bマイグレーション完了',
          '顧客マスタシートを新スキーマで作成しました。\n\n' +
          '新スキーマ: 呼称 | 名称 | 郵便番号 | 住所1 | 住所2 | 役職 | 担当者名 | メール | 備考',
          SpreadsheetApp.getUi().ButtonSet.OK
        );
      } catch (uiErr) { /* スタンドアロンの場合 */ }
      return;
    }

    // 既に名称列があるかチェック（A1=呼称 かつ B1=名称）
    var headerA = sheet.getRange('A1').getValue();
    var headerB = sheet.getRange('B1').getValue();
    if (headerA === '呼称' && headerB === '名称') {
      Logger.log('  顧客マスタは既にv4bスキーマです（スキップ）');
      try {
        SpreadsheetApp.getUi().alert('v4bマイグレーション', '既にv4bスキーマです。処理をスキップしました。', SpreadsheetApp.getUi().ButtonSet.OK);
      } catch (uiErr) { /* スタンドアロンの場合 */ }
      return;
    }

    // --- 1. B列に列を挿入 ---
    sheet.insertColumnBefore(2);
    Logger.log('  ✓ B列に列を挿入');

    // --- 2. ヘッダー更新 ---
    sheet.getRange('A1').setValue('呼称');
    sheet.getRange('B1').setValue('名称');
    sheet.getRange('B1').setFontWeight('bold').setBackground('#4472C4').setFontColor('#FFFFFF');
    Logger.log('  ✓ ヘッダー更新: A1=呼称, B1=名称');

    // --- 3. 既存データのA列（呼称）をB列（名称）にコピー ---
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var nicknames = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      sheet.getRange(2, 2, lastRow - 1, 1).setValues(nicknames);
      Logger.log('  ✓ 既存 ' + (lastRow - 1) + '件の呼称を名称にコピー');
    }

    // --- 4. B列の列幅設定 ---
    sheet.setColumnWidth(2, 200);
    Logger.log('  ✓ B列の列幅設定');

    Logger.log('');
    Logger.log('=== v4b マイグレーション完了 ===');
    Logger.log('新スキーマ: 呼称 | 名称 | 郵便番号 | 住所1 | 住所2 | 役職 | 担当者名 | メール | 備考');

    try {
      SpreadsheetApp.getUi().alert(
        'v4bマイグレーション完了',
        '顧客マスタに「名称」列を追加しました。\n\n' +
        '新スキーマ: 呼称 | 名称 | 郵便番号 | 住所1 | ...\n\n' +
        '既存の呼称を名称にもコピーしました。\n' +
        '必要に応じて名称（帳票記載用）を修正してください。\n\n' +
        '★ Web Appを再デプロイしてください。',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } catch (uiErr) { /* スタンドアロンの場合 */ }

  } catch (error) {
    Logger.log('★ v4bマイグレーションエラー: ' + error.message);
    Logger.log('スタックトレース: ' + error.stack);
    throw error;
  }
}

// ============================================================================
// v5 マイグレーション（仕入・粗利列追加: U〜AC列）
// ============================================================================

/**
 * 案件マスタに仕入・粗利列（U〜AC）を追加
 *
 * 追加列:
 *   U(21): 仕入先名、V(22): 仕入金額（税抜）、W(23): 仕入消費税（数式）
 *   X(24): 仕入合計（数式）、Y(25): 仕入先請求番号、Z(26): 支払日
 *   AA(27): 仕入先請求書URL、AB(28): 粗利（数式）、AC(29): 粗利率（数式）
 *
 * ★ 実行場所: スプレッドシート → 拡張機能 → Apps Script → migrateV5 を実行
 */
function migrateV5() {
  try {
    Logger.log('=== v5 マイグレーション開始 ===');
    Logger.log('変更: 案件マスタに仕入・粗利列（U〜AC）を追加');

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('案件マスタ');
    if (!sheet) throw new Error('「案件マスタ」シートが見つかりません');

    // 既にU1にヘッダーがあるかチェック
    var existingU1 = sheet.getRange('U1').getValue();
    if (existingU1 === '仕入先名') {
      Logger.log('  仕入列は既に存在します（スキップ）');
      try {
        SpreadsheetApp.getUi().alert('v5マイグレーション', '仕入列は既に存在します。処理をスキップしました。', SpreadsheetApp.getUi().ButtonSet.OK);
      } catch (uiErr) { /* スタンドアロンの場合 */ }
      return;
    }

    // --- 1. ヘッダー追加 ---
    var newHeaders = [['仕入先名', '仕入金額（税抜）', '仕入消費税', '仕入合計', '仕入先請求番号', '支払日', '仕入先請求書URL', '粗利', '粗利率']];
    sheet.getRange('U1:AC1').setValues(newHeaders);
    sheet.getRange('U1:AC1').setFontWeight('bold').setBackground('#4472C4').setFontColor('#FFFFFF');
    Logger.log('  ✓ U〜AC列 ヘッダー追加');

    // --- 2. 書式設定 ---
    sheet.getRange('Z:Z').setNumberFormat('yyyy/mm/dd');
    sheet.getRange('V:X').setNumberFormat('#,##0');
    sheet.getRange('AB:AB').setNumberFormat('#,##0');
    sheet.getRange('AC:AC').setNumberFormat('0.0%');
    Logger.log('  ✓ 書式設定');

    // --- 3. 列幅設定 ---
    var colWidths = { 21: 140, 22: 120, 23: 100, 24: 120, 25: 140, 26: 100, 27: 200, 28: 100, 29: 80 };
    for (var col in colWidths) {
      sheet.setColumnWidth(parseInt(col), colWidths[col]);
    }
    Logger.log('  ✓ 列幅設定');

    // --- 4. 既存データ行に数式をセット ---
    var lastRow = sheet.getLastRow();
    var formulaCount = 0;
    for (var r = 2; r <= lastRow; r++) {
      var pid = sheet.getRange(r, 1).getValue();
      if (!pid) continue;
      sheet.getRange(r, 23).setFormula('=FLOOR(V' + r + '*設定!B13)');        // W: 仕入消費税
      sheet.getRange(r, 24).setFormula('=V' + r + '+W' + r);                 // X: 仕入合計
      sheet.getRange(r, 28).setFormula('=L' + r + '-X' + r);                 // AB: 粗利
      sheet.getRange(r, 29).setFormula('=IFERROR(AB' + r + '/L' + r + ',0)'); // AC: 粗利率
      formulaCount++;
    }
    Logger.log('  ✓ 数式セット: ' + formulaCount + '行');

    Logger.log('');
    Logger.log('=== v5 マイグレーション完了 ===');

    try {
      SpreadsheetApp.getUi().alert(
        'v5マイグレーション完了',
        '案件マスタに仕入・粗利列を追加しました。\n\n' +
        '追加列: 仕入先名 | 仕入金額 | 仕入消費税 | 仕入合計 | 請求番号 | 支払日 | 請求書URL | 粗利 | 粗利率\n\n' +
        '数式セット: ' + formulaCount + '行\n\n' +
        '★ Web Appを再デプロイしてください。',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } catch (uiErr) { /* スタンドアロンの場合 */ }

  } catch (error) {
    Logger.log('★ v5マイグレーションエラー: ' + error.message);
    Logger.log('スタックトレース: ' + error.stack);
    throw error;
  }
}
