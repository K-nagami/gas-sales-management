/**
 * 売上管理システム - Googleドキュメントテンプレート生成
 *
 * 初期セットアップ時に3つの帳票テンプレートを作成する。
 * プレースホルダーは main.gs の buildReplacements() と完全一致させること。
 *
 * プレースホルダー一覧（INSTRUCTIONS.md準拠）:
 *   {{自社名}}, {{自社住所}}, {{自社TEL}}, {{自社メール}}, {{適格番号}}
 *   {{顧客名}}, {{案件名}}, {{案件ID}}
 *   {{見積日}}, {{発注日}}, {{請求日}}
 *   {{明細テーブル}}
 *   {{小計}}, {{消費税}}, {{合計金額}}
 *   {{振込先情報}}, {{支払条件}}, {{発注書注意書き}}
 *   {{入金予定日}}
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

// ============================================================================
// 見積書テンプレート
// ============================================================================

/**
 * 見積書テンプレートの内容を構築
 * @param {DocumentApp.Body} body
 * @private
 */
function buildEstimateContent_(body) {
  // タイトル
  var title = body.appendParagraph('【見積書】');
  title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  title.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  title.editAsText().setBold(true).setFontSize(18);

  body.appendHorizontalRule();

  // 見積書番号・発行日
  body.appendParagraph('見積書番号: {{案件ID}}').editAsText().setFontSize(10);
  body.appendParagraph('発行日: {{見積日}}').editAsText().setFontSize(10);
  body.appendParagraph('');

  // 顧客名
  var customer = body.appendParagraph('{{顧客名}} 御中');
  customer.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  customer.editAsText().setFontSize(14).setBold(true);
  body.appendParagraph('');

  // 案件名
  body.appendParagraph('件名: {{案件名}}').editAsText().setFontSize(10);
  body.appendParagraph('');

  body.appendParagraph('下記の通りお見積り申し上げます。').editAsText().setFontSize(10);
  body.appendParagraph('');

  // 自社情報（右寄せ）
  var lines = ['{{自社名}}', '{{自社住所}}', 'TEL: {{自社TEL}}', 'Email: {{自社メール}}', '適格請求書発行事業者番号: {{適格番号}}'];
  for (var i = 0; i < lines.length; i++) {
    var p = body.appendParagraph(lines[i]);
    p.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    p.editAsText().setFontSize(10);
  }

  body.appendParagraph('');
  body.appendHorizontalRule();

  // 明細セクション
  var detailHeader = body.appendParagraph('【明細】');
  detailHeader.editAsText().setBold(true).setFontSize(12);
  body.appendParagraph('{{明細テーブル}}').editAsText().setFontSize(10);
  body.appendParagraph('');
  body.appendHorizontalRule();

  // 金額
  body.appendParagraph('小計:     {{小計}}').editAsText().setFontSize(10);
  body.appendParagraph('消費税:   {{消費税}}').editAsText().setFontSize(10);
  var totalP = body.appendParagraph('合計金額: {{合計金額}}');
  totalP.editAsText().setBold(true).setFontSize(12);
  body.appendParagraph('');

  // 有効期限・備考
  body.appendParagraph('有効期限: 発行日より30日間').editAsText().setFontSize(10);
  body.appendParagraph('備考: {{支払条件}}').editAsText().setFontSize(10);
}

// ============================================================================
// 発注書テンプレート
// ============================================================================

/**
 * 発注書テンプレートの内容を構築
 * @param {DocumentApp.Body} body
 * @private
 */
function buildOrderFormContent_(body) {
  // タイトル
  var title = body.appendParagraph('【発注書】');
  title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  title.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  title.editAsText().setBold(true).setFontSize(18);

  body.appendHorizontalRule();

  // 関連見積番号
  body.appendParagraph('関連見積番号: {{案件ID}}').editAsText().setFontSize(10);
  body.appendParagraph('');

  // 宛先（自社 = 発注を受ける側）
  var recipient = body.appendParagraph('{{自社名}} 御中');
  recipient.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  recipient.editAsText().setFontSize(14).setBold(true);
  body.appendParagraph('');

  body.appendParagraph('下記の通り発注いたします。').editAsText().setFontSize(10);
  body.appendParagraph('');
  body.appendHorizontalRule();

  // 明細セクション
  var detailHeader = body.appendParagraph('【明細】');
  detailHeader.editAsText().setBold(true).setFontSize(12);
  body.appendParagraph('{{明細テーブル}}').editAsText().setFontSize(10);
  body.appendParagraph('');
  body.appendHorizontalRule();

  // 金額
  body.appendParagraph('小計:     {{小計}}').editAsText().setFontSize(10);
  body.appendParagraph('消費税:   {{消費税}}').editAsText().setFontSize(10);
  var totalP = body.appendParagraph('合計金額: {{合計金額}}');
  totalP.editAsText().setBold(true).setFontSize(12);
  body.appendParagraph('');

  // 支払条件
  body.appendParagraph('支払条件: {{支払条件}}').editAsText().setFontSize(10);
  body.appendParagraph('');

  // 注意書き
  body.appendParagraph('{{発注書注意書き}}').editAsText().setFontSize(9);
  body.appendParagraph('');
  body.appendHorizontalRule();

  // 署名欄
  body.appendParagraph('発注日:          年      月      日').editAsText().setFontSize(10);
  body.appendParagraph('貴社名: {{顧客名}}').editAsText().setFontSize(10);
  body.appendParagraph('ご担当者名:                        印').editAsText().setFontSize(10);
}

// ============================================================================
// 請求書テンプレート
// ============================================================================

/**
 * 請求書テンプレートの内容を構築
 * @param {DocumentApp.Body} body
 * @private
 */
function buildInvoiceContent_(body) {
  // タイトル
  var title = body.appendParagraph('【請求書】');
  title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  title.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  title.editAsText().setBold(true).setFontSize(18);

  body.appendHorizontalRule();

  // 請求書番号・発行日
  body.appendParagraph('請求書番号: {{案件ID}}').editAsText().setFontSize(10);
  body.appendParagraph('発行日: {{請求日}}').editAsText().setFontSize(10);
  body.appendParagraph('');

  // 顧客名
  var customer = body.appendParagraph('{{顧客名}} 御中');
  customer.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  customer.editAsText().setFontSize(14).setBold(true);
  body.appendParagraph('');

  body.appendParagraph('下記の通りご請求申し上げます。').editAsText().setFontSize(10);
  body.appendParagraph('');

  // 自社情報（右寄せ）
  var lines = ['{{自社名}}', '{{自社住所}}', 'TEL: {{自社TEL}}', 'Email: {{自社メール}}', '適格請求書発行事業者番号: {{適格番号}}'];
  for (var i = 0; i < lines.length; i++) {
    var p = body.appendParagraph(lines[i]);
    p.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    p.editAsText().setFontSize(10);
  }

  body.appendParagraph('');
  body.appendHorizontalRule();

  // 明細セクション
  var detailHeader = body.appendParagraph('【明細】');
  detailHeader.editAsText().setBold(true).setFontSize(12);
  body.appendParagraph('{{明細テーブル}}').editAsText().setFontSize(10);
  body.appendParagraph('');
  body.appendHorizontalRule();

  // 金額
  body.appendParagraph('小計:     {{小計}}').editAsText().setFontSize(10);
  body.appendParagraph('消費税(10%): {{消費税}}').editAsText().setFontSize(10);
  var totalP = body.appendParagraph('合計金額: {{合計金額}}');
  totalP.editAsText().setBold(true).setFontSize(12);
  body.appendParagraph('');
  body.appendHorizontalRule();

  // お振込先
  var bankHeader = body.appendParagraph('【お振込先】');
  bankHeader.editAsText().setBold(true).setFontSize(12);
  body.appendParagraph('{{振込先情報}}').editAsText().setFontSize(10);
  body.appendParagraph('');

  // 支払期限・条件
  body.appendParagraph('お支払期限: {{入金予定日}}').editAsText().setFontSize(10);
  body.appendParagraph('{{支払条件}}').editAsText().setFontSize(10);
}
