/**
 * 売上管理システム - Googleドキュメントテンプレート生成 v4
 *
 * 初期セットアップ時に3つの帳票テンプレートを作成する。
 * プレースホルダーは main.gs の buildReplacements() と完全一致させること。
 *
 * v4 プレースホルダー一覧:
 *   {{自社名}}, {{自社住所}}, {{自社TEL}}, {{自社メール}}, {{適格番号}}
 *   {{自社担当者名}}, {{自社URL}}, {{社印画像}}
 *   {{顧客名}}, {{顧客郵便番号}}, {{顧客住所1}}, {{顧客住所2}}, {{顧客役職}}, {{顧客担当者名}}
 *   {{案件名}}, {{案件ID}}
 *   {{見積日}}, {{発注日}}, {{請求日}}, {{入金予定日}}
 *   {{明細テーブル}}
 *   {{小計}}, {{消費税}}, {{合計金額}}
 *   {{10%対象}}, {{10%消費税}}
 *   {{振込先情報}}, {{支払条件}}, {{発注書注意書き}}
 */

// ============================================================================
// メイン関数
// ============================================================================

/**
 * 3つのGoogleドキュメントテンプレート（見積書・発注書・請求書）を作成
 * @returns {Object} { estimateId, orderFormId, invoiceId }
 */
function createAllDocumentTemplates() {
  try {
    Logger.log('=== ドキュメントテンプレート作成開始 ===');

    var estimateId = createTemplateDoc_('【テンプレート】見積書', buildEstimateContent_);
    Logger.log('✓ 見積書テンプレート: ' + estimateId);

    var orderFormId = createTemplateDoc_('【テンプレート】発注書', buildOrderFormContent_);
    Logger.log('✓ 発注書テンプレート: ' + orderFormId);

    var invoiceId = createTemplateDoc_('【テンプレート】請求書', buildInvoiceContent_);
    Logger.log('✓ 請求書テンプレート: ' + invoiceId);

    Logger.log('=== テンプレート作成完了 ===');

    return {
      estimateId: estimateId,
      orderFormId: orderFormId,
      invoiceId: invoiceId
    };
  } catch (error) {
    Logger.log('テンプレート作成エラー: ' + error.message);
    throw new Error('テンプレート作成に失敗: ' + error.message);
  }
}

// ============================================================================
// ヘルパー関数
// ============================================================================

/**
 * Googleドキュメントを作成し、コンテンツビルダーで内容を構築する
 * @param {string} title ドキュメントタイトル
 * @param {Function} contentBuilder body を受け取るコールバック関数
 * @returns {string} ドキュメントID
 * @private
 */
function createTemplateDoc_(title, contentBuilder) {
  try {
    var doc = DocumentApp.create(title);
    var body = doc.getBody();
    body.clear();

    contentBuilder(body);

    doc.saveAndClose();
    return doc.getId();
  } catch (error) {
    Logger.log('ドキュメント作成エラー (' + title + '): ' + error.message);
    throw error;
  }
}

/**
 * 2列ボーダレスヘッダーテーブルを構築する共通ヘルパー
 * @param {DocumentApp.Body} body
 * @param {Array} leftLines 左セルの行テキスト配列
 * @param {Array} rightLines 右セルの行テキスト配列
 * @returns {DocumentApp.Table}
 * @private
 */
function buildHeaderTable_(body, leftLines, rightLines) {
  var table = body.appendTable([[leftLines.join('\n'), rightLines.join('\n')]]);

  // ボーダーなし
  table.setBorderWidth(0);

  // 列幅の調整（左50%、右50%の概算）
  var leftCell = table.getRow(0).getCell(0);
  var rightCell = table.getRow(0).getCell(1);
  leftCell.setWidth(250);
  rightCell.setWidth(250);

  // セルのパディング
  leftCell.setPaddingTop(2);
  leftCell.setPaddingBottom(2);
  rightCell.setPaddingTop(2);
  rightCell.setPaddingBottom(2);

  return table;
}

/**
 * サマリーテーブル（3列: 小計・消費税・合計金額）を追加
 * @param {DocumentApp.Body} body
 * @param {string} totalLabel 合計金額ラベル（見積金額/発注金額/請求金額）
 * @private
 */
function appendSummaryTable_(body, totalLabel) {
  var table = body.appendTable([
    ['小計', '消費税(10%)', totalLabel],
    ['{{小計}}', '{{消費税}}', '{{合計金額}}']
  ]);

  // ヘッダー行スタイル
  var headerRow = table.getRow(0);
  for (var c = 0; c < 3; c++) {
    headerRow.getCell(c).editAsText().setBold(true).setFontSize(9);
    headerRow.getCell(c).setBackgroundColor('#F2F2F2');
  }

  // データ行スタイル
  var dataRow = table.getRow(1);
  for (var d = 0; d < 3; d++) {
    dataRow.getCell(d).editAsText().setFontSize(10);
  }
  // 合計金額を太字14pt
  dataRow.getCell(2).editAsText().setBold(true).setFontSize(14);

  return table;
}

/**
 * 内訳セクションを追加（10%対象税抜・10%消費税）
 * @param {DocumentApp.Body} body
 * @private
 */
function appendBreakdownSection_(body) {
  var p = body.appendParagraph('10%対象（税抜） {{10%対象}}  /  10%消費税 {{10%消費税}}');
  p.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  p.editAsText().setFontSize(9);
}

/**
 * 備考欄テーブルを追加（枠付き）
 * @param {DocumentApp.Body} body
 * @param {string} defaultText 備考欄のデフォルトテキスト
 * @private
 */
function appendRemarksTable_(body, defaultText) {
  var table = body.appendTable([
    ['備考'],
    [defaultText || '']
  ]);

  var headerRow = table.getRow(0);
  headerRow.getCell(0).editAsText().setBold(true).setFontSize(9);
  headerRow.getCell(0).setBackgroundColor('#F2F2F2');

  var dataRow = table.getRow(1);
  dataRow.getCell(0).editAsText().setFontSize(9);
  dataRow.setMinimumHeight(50);

  return table;
}

// ============================================================================
// 見積書テンプレート v4
// ============================================================================

/**
 * 見積書テンプレートの内容を構築
 * @param {DocumentApp.Body} body
 * @private
 */
function buildEstimateContent_(body) {
  // === タイトル ===
  var title = body.appendParagraph('見積書');
  title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  title.editAsText().setBold(true).setFontSize(20);

  body.appendParagraph('');

  // === 2カラムヘッダー（ボーダレステーブル） ===
  var leftLines = [
    '{{顧客名}} 御中',
    '〒{{顧客郵便番号}}',
    '{{顧客住所1}}',
    '{{顧客住所2}}',
    '{{顧客役職}} {{顧客担当者名}} 様'
  ];
  var rightLines = [
    '見積日: {{見積日}}',
    '見積書番号: {{案件ID}}',
    '登録番号: {{適格番号}}',
    '',
    '{{自社名}}',
    '{{自社担当者名}}',
    '{{自社住所}}',
    'TEL: {{自社TEL}}',
    '{{自社URL}}',
    '{{社印画像}}'
  ];

  var headerTable = body.appendTable([[leftLines.join('\n'), rightLines.join('\n')]]);
  headerTable.setBorderWidth(0);
  var leftCell = headerTable.getRow(0).getCell(0);
  var rightCell = headerTable.getRow(0).getCell(1);
  leftCell.setWidth(250);
  rightCell.setWidth(250);
  leftCell.setPaddingTop(2);
  leftCell.setPaddingBottom(2);
  rightCell.setPaddingTop(2);
  rightCell.setPaddingBottom(2);

  // 左セル: 顧客名を太字14ptに
  leftCell.editAsText().setFontSize(10);
  leftCell.editAsText().setBold(false);
  // 顧客名行のみ太字14pt
  var leftText = leftCell.editAsText();
  var customerEnd = leftLines[0].length;
  leftText.setBold(0, customerEnd - 1, true);
  leftText.setFontSize(0, customerEnd - 1, 14);

  // 右セル: 自社名を太字11pt
  rightCell.editAsText().setFontSize(10);
  rightCell.editAsText().setBold(false);
  var rightText = rightCell.editAsText();
  // 自社名の位置を計算（4行目: 空行の後）
  var rightOffset = 0;
  for (var ri = 0; ri < 4; ri++) {
    rightOffset += rightLines[ri].length + 1; // +1 for \n
  }
  var companyNameLen = rightLines[4].length;
  rightText.setBold(rightOffset, rightOffset + companyNameLen - 1, true);
  rightText.setFontSize(rightOffset, rightOffset + companyNameLen - 1, 11);

  body.appendParagraph('');

  // === 件名 ===
  body.appendParagraph('件名　{{案件名}}').editAsText().setFontSize(10);
  body.appendParagraph('');

  // === サマリーテーブル ===
  appendSummaryTable_(body, '見積金額');

  body.appendParagraph('');

  // === 明細テーブル ===
  body.appendParagraph('{{明細テーブル}}').editAsText().setFontSize(10);
  body.appendParagraph('');

  // === 内訳 ===
  appendBreakdownSection_(body);
  body.appendParagraph('');

  // === 備考欄 ===
  appendRemarksTable_(body, '支払条件: {{支払条件}}\n有効期限: 発行日より30日間');
}

// ============================================================================
// 発注書テンプレート v4
// ============================================================================

/**
 * 発注書テンプレートの内容を構築
 * @param {DocumentApp.Body} body
 * @private
 */
function buildOrderFormContent_(body) {
  // === タイトル ===
  var title = body.appendParagraph('発注書');
  title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  title.editAsText().setBold(true).setFontSize(20);

  body.appendParagraph('');

  // === 2カラムヘッダー ===
  // 左セル: 自社名（発注を受ける側）
  var leftLines = [
    '{{自社名}} 御中',
    '{{自社住所}}',
    'TEL: {{自社TEL}}'
  ];
  // 右セル: 発注日・番号 + 顧客情報（発注者）
  var rightLines = [
    '発注日: {{発注日}}',
    '発注書番号: {{案件ID}}',
    '',
    '{{顧客名}}',
    '〒{{顧客郵便番号}}',
    '{{顧客住所1}}',
    '{{顧客住所2}}',
    '{{顧客役職}} {{顧客担当者名}}'
  ];

  var headerTable = body.appendTable([[leftLines.join('\n'), rightLines.join('\n')]]);
  headerTable.setBorderWidth(0);
  var leftCell = headerTable.getRow(0).getCell(0);
  var rightCell = headerTable.getRow(0).getCell(1);
  leftCell.setWidth(250);
  rightCell.setWidth(250);
  leftCell.setPaddingTop(2);
  leftCell.setPaddingBottom(2);
  rightCell.setPaddingTop(2);
  rightCell.setPaddingBottom(2);

  // 左セル: 自社名御中を太字14pt
  leftCell.editAsText().setFontSize(10);
  leftCell.editAsText().setBold(false);
  var leftText = leftCell.editAsText();
  var recipientEnd = leftLines[0].length;
  leftText.setBold(0, recipientEnd - 1, true);
  leftText.setFontSize(0, recipientEnd - 1, 14);

  // 右セル全体
  rightCell.editAsText().setFontSize(10);
  rightCell.editAsText().setBold(false);

  body.appendParagraph('');
  body.appendParagraph('下記の通り発注いたします。').editAsText().setFontSize(10);
  body.appendParagraph('');

  // === 件名 ===
  body.appendParagraph('件名　{{案件名}}').editAsText().setFontSize(10);
  body.appendParagraph('');

  // === サマリーテーブル ===
  appendSummaryTable_(body, '発注金額');

  body.appendParagraph('');

  // === 明細テーブル ===
  body.appendParagraph('{{明細テーブル}}').editAsText().setFontSize(10);
  body.appendParagraph('');

  // === 内訳 ===
  appendBreakdownSection_(body);
  body.appendParagraph('');

  // === 備考欄 ===
  appendRemarksTable_(body, '支払条件: {{支払条件}}\n{{発注書注意書き}}');

  body.appendParagraph('');

  // === 署名欄テーブル ===
  var signTable = body.appendTable([
    ['発注日', '発注者名', 'ご担当者名・印'],
    ['     年     月     日', '{{顧客名}}', '                        印']
  ]);

  var signHeader = signTable.getRow(0);
  for (var s = 0; s < 3; s++) {
    signHeader.getCell(s).editAsText().setBold(true).setFontSize(9);
    signHeader.getCell(s).setBackgroundColor('#F2F2F2');
  }
  var signData = signTable.getRow(1);
  for (var sd = 0; sd < 3; sd++) {
    signData.getCell(sd).editAsText().setFontSize(10);
  }
  signData.setMinimumHeight(40);
}

// ============================================================================
// 請求書テンプレート v4
// ============================================================================

/**
 * 請求書テンプレートの内容を構築
 * @param {DocumentApp.Body} body
 * @private
 */
function buildInvoiceContent_(body) {
  // === タイトル ===
  var title = body.appendParagraph('請求書');
  title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  title.editAsText().setBold(true).setFontSize(20);

  body.appendParagraph('');

  // === 2カラムヘッダー（ボーダレステーブル） ===
  var leftLines = [
    '{{顧客名}} 御中',
    '〒{{顧客郵便番号}}',
    '{{顧客住所1}}',
    '{{顧客住所2}}',
    '{{顧客役職}} {{顧客担当者名}} 様'
  ];
  var rightLines = [
    '請求日: {{請求日}}',
    '請求書番号: {{案件ID}}',
    '登録番号: {{適格番号}}',
    '',
    '{{自社名}}',
    '{{自社担当者名}}',
    '{{自社住所}}',
    'TEL: {{自社TEL}}',
    '{{自社URL}}',
    '{{社印画像}}'
  ];

  var headerTable = body.appendTable([[leftLines.join('\n'), rightLines.join('\n')]]);
  headerTable.setBorderWidth(0);
  var leftCell = headerTable.getRow(0).getCell(0);
  var rightCell = headerTable.getRow(0).getCell(1);
  leftCell.setWidth(250);
  rightCell.setWidth(250);
  leftCell.setPaddingTop(2);
  leftCell.setPaddingBottom(2);
  rightCell.setPaddingTop(2);
  rightCell.setPaddingBottom(2);

  // 左セル: 顧客名を太字14pt
  leftCell.editAsText().setFontSize(10);
  leftCell.editAsText().setBold(false);
  var leftText = leftCell.editAsText();
  var customerEnd = leftLines[0].length;
  leftText.setBold(0, customerEnd - 1, true);
  leftText.setFontSize(0, customerEnd - 1, 14);

  // 右セル: 自社名を太字11pt
  rightCell.editAsText().setFontSize(10);
  rightCell.editAsText().setBold(false);
  var rightText = rightCell.editAsText();
  var rightOffset = 0;
  for (var ri = 0; ri < 4; ri++) {
    rightOffset += rightLines[ri].length + 1;
  }
  var companyNameLen = rightLines[4].length;
  rightText.setBold(rightOffset, rightOffset + companyNameLen - 1, true);
  rightText.setFontSize(rightOffset, rightOffset + companyNameLen - 1, 11);

  body.appendParagraph('');

  // === 件名 ===
  body.appendParagraph('件名　{{案件名}}').editAsText().setFontSize(10);
  body.appendParagraph('');

  // === サマリーテーブル ===
  appendSummaryTable_(body, '請求金額');

  body.appendParagraph('');

  // === 入金・振込テーブル（請求書のみ） ===
  var payTable = body.appendTable([
    ['入金期日', '振込先'],
    ['{{入金予定日}}', '{{振込先情報}}']
  ]);

  var payHeader = payTable.getRow(0);
  for (var ph = 0; ph < 2; ph++) {
    payHeader.getCell(ph).editAsText().setBold(true).setFontSize(9);
    payHeader.getCell(ph).setBackgroundColor('#F2F2F2');
  }
  var payData = payTable.getRow(1);
  payData.getCell(0).editAsText().setFontSize(10);
  payData.getCell(0).setWidth(120);
  payData.getCell(1).editAsText().setFontSize(10);

  body.appendParagraph('');

  // === 明細テーブル ===
  body.appendParagraph('{{明細テーブル}}').editAsText().setFontSize(10);
  body.appendParagraph('');

  // === 内訳 ===
  appendBreakdownSection_(body);
  body.appendParagraph('');

  // === 備考欄 ===
  appendRemarksTable_(body, '{{支払条件}}');
}
