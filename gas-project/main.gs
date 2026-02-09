/**
 * 売上管理統合システム - メインアプリケーション
 *
 * 機能:
 *   - カスタムメニュー「帳票」
 *   - PDF生成（見積書・発注書・請求書）
 *   - 税理士CSV出力
 *   - 案件ID自動採番
 *   - ステータス変更トリガー（onEdit）
 *
 * プレースホルダー対応表（templates.gs と一致させること）:
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
// メニュー
// ============================================================================

/**
 * スプレッドシートを開いた時にカスタムメニューを追加
 */
function onOpen() {
  try {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('帳票')
      .addItem('見積書PDF生成', 'generateEstimatePDF')
      .addItem('発注書PDF生成', 'generateOrderFormPDF')
      .addItem('請求書PDF生成', 'generateInvoicePDF')
      .addSeparator()
      .addItem('税理士CSV出力', 'exportCSVForAccountant')
      .addItem('案件ID自動採番', 'assignProjectId')
      .addToUi();
  } catch (error) {
    Logger.log('メニュー設定エラー: ' + error.message);
  }
}

// ============================================================================
// PDF生成 - 見積書
// ============================================================================

/**
 * アクティブ行の案件に対して見積書PDFを生成
 */
function generateEstimatePDF() {
  try {
    var rowData = getActiveRowData_();
    if (!rowData) return;

    var settings = getSettingsData();
    var items = getItemsForProject(rowData.projectId);
    if (items.length === 0) {
      SpreadsheetApp.getUi().alert('明細データが見つかりません: ' + rowData.projectId);
      return;
    }

    var replacements = buildReplacements_(rowData, settings, items);
    var fileName = '見積書_' + rowData.projectId + '_' + rowData.customerName + '.pdf';
    var transactionDate = replacements['{{見積日}}'];

    var pdfUrl = fillTemplateAndConvertToPDF(
      settings.estimateTemplateId,
      replacements,
      items,
      fileName,
      settings.folderId,
      transactionDate,
      settings.sealImageId
    );

    // M列（13列目）に見積書URLを記入
    rowData.sheet.getRange(rowData.row, 13).setValue(pdfUrl);
    SpreadsheetApp.getUi().alert('見積書を生成しました。\n' + pdfUrl);

  } catch (error) {
    SpreadsheetApp.getUi().alert('見積書生成エラー:\n' + error.message);
  }
}

// ============================================================================
// PDF生成 - 発注書
// ============================================================================

/**
 * アクティブ行の案件に対して発注書PDFを生成
 */
function generateOrderFormPDF() {
  try {
    var rowData = getActiveRowData_();
    if (!rowData) return;

    var settings = getSettingsData();
    var items = getItemsForProject(rowData.projectId);
    if (items.length === 0) {
      SpreadsheetApp.getUi().alert('明細データが見つかりません: ' + rowData.projectId);
      return;
    }

    var replacements = buildReplacements_(rowData, settings, items);
    // 発注書固有のプレースホルダー
    replacements['{{発注書注意書き}}'] = settings.orderNote;

    var fileName = '発注書_' + rowData.projectId + '_' + rowData.customerName + '.pdf';
    var transactionDate = replacements['{{発注日}}'];

    var pdfUrl = fillTemplateAndConvertToPDF(
      settings.orderTemplateId,
      replacements,
      items,
      fileName,
      settings.folderId,
      transactionDate,
      settings.sealImageId
    );

    // N列（14列目）に発注書URLを記入
    rowData.sheet.getRange(rowData.row, 14).setValue(pdfUrl);
    SpreadsheetApp.getUi().alert('発注書を生成しました。\n' + pdfUrl);

  } catch (error) {
    SpreadsheetApp.getUi().alert('発注書生成エラー:\n' + error.message);
  }
}

// ============================================================================
// PDF生成 - 請求書
// ============================================================================

/**
 * アクティブ行の案件に対して請求書PDFを生成
 */
function generateInvoicePDF() {
  try {
    var rowData = getActiveRowData_();
    if (!rowData) return;

    var settings = getSettingsData();
    var items = getItemsForProject(rowData.projectId);
    if (items.length === 0) {
      SpreadsheetApp.getUi().alert('明細データが見つかりません: ' + rowData.projectId);
      return;
    }

    var replacements = buildReplacements_(rowData, settings, items);
    // 請求書固有: 振込先情報を組み立て
    replacements['{{振込先情報}}'] = settings.bankName + ' ' + settings.branchName + '\n' +
      settings.accountType + ' ' + settings.accountNumber + '\n' +
      '口座名義: ' + settings.accountHolder;
    // 入金予定日
    replacements['{{入金予定日}}'] = rowData.values[6] ? formatDateValue_(rowData.values[6]) : '（未設定）';

    var fileName = '請求書_' + rowData.projectId + '_' + rowData.customerName + '.pdf';
    var transactionDate = replacements['{{請求日}}'];

    var pdfUrl = fillTemplateAndConvertToPDF(
      settings.invoiceTemplateId,
      replacements,
      items,
      fileName,
      settings.folderId,
      transactionDate,
      settings.sealImageId
    );

    // O列（15列目）に請求書URLを記入
    rowData.sheet.getRange(rowData.row, 15).setValue(pdfUrl);
    SpreadsheetApp.getUi().alert('請求書を生成しました。\n' + pdfUrl);

  } catch (error) {
    SpreadsheetApp.getUi().alert('請求書生成エラー:\n' + error.message);
  }
}

// ============================================================================
// PDF生成 共通基盤
// ============================================================================

/**
 * テンプレートを複製 → プレースホルダー置換 → 明細テーブル挿入 → PDF変換
 *
 * @param {string} templateId テンプレートドキュメントID
 * @param {Object} replacements { placeholder: value } の辞書
 * @param {Array} items 明細行データ [{name, unitPrice, quantity, subtotal}]
 * @param {string} fileName 生成するPDFのファイル名
 * @param {string} folderId 保存先GoogleドライブフォルダID
 * @param {string} transactionDate 取引日（YYYY/MM/DD形式）
 * @param {string} sealImageId 社印画像のファイルID（省略可）
 * @returns {string} 生成されたPDFのURL
 */
function fillTemplateAndConvertToPDF(templateId, replacements, items, fileName, folderId, transactionDate, sealImageId) {
  var tempDocId = null;

  try {
    // テンプレートを複製
    var templateFile = DriveApp.getFileById(templateId);
    var tempFile = templateFile.makeCopy('_temp_' + Date.now());
    tempDocId = tempFile.getId();

    // ドキュメントを開く
    var doc = DocumentApp.openById(tempDocId);
    var body = doc.getBody();

    // プレースホルダー置換（明細テーブル・社印画像以外）
    for (var placeholder in replacements) {
      if (placeholder === '{{明細テーブル}}') continue;
      if (placeholder === '{{社印画像}}') continue;
      body.replaceText(escapeRegExp_(placeholder), replacements[placeholder] || '');
    }

    // 社印画像挿入（プレースホルダー置換後、テーブル挿入前）
    insertSealImage_(body, sealImageId || '');

    // 明細テーブルの挿入（5列: 取引日・摘要・数量・単価・明細金額）
    insertItemsTable_(body, items, transactionDate || '');

    doc.saveAndClose();

    // PDF変換
    var pdfBlob = DriveApp.getFileById(tempDocId).getAs('application/pdf');
    pdfBlob.setName(fileName);

    // 帳票種別を判定してサブフォルダに振り分け
    var subFolderName = '';
    if (fileName.indexOf('見積書') === 0) {
      subFolderName = '見積書';
    } else if (fileName.indexOf('発注書') === 0) {
      subFolderName = '発注書';
    } else if (fileName.indexOf('請求書') === 0) {
      subFolderName = '請求書';
    }

    var parentFolder = DriveApp.getFolderById(folderId);
    var targetFolder = getOrCreateSubFolder_(parentFolder, subFolderName);
    var pdfFile = targetFolder.createFile(pdfBlob);
    var pdfUrl = pdfFile.getUrl();

    // テンポラリファイル削除
    tempFile.setTrashed(true);

    return pdfUrl;

  } catch (error) {
    // クリーンアップ
    if (tempDocId) {
      try { DriveApp.getFileById(tempDocId).setTrashed(true); } catch (e) { /* 無視 */ }
    }
    throw error;
  }
}

/**
 * 親フォルダ内にサブフォルダを取得（なければ自動作成）
 * @param {Folder} parentFolder 親フォルダ
 * @param {string} subFolderName サブフォルダ名
 * @returns {Folder} サブフォルダ
 * @private
 */
function getOrCreateSubFolder_(parentFolder, subFolderName) {
  if (!subFolderName) return parentFolder;

  var folders = parentFolder.getFoldersByName(subFolderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parentFolder.createFolder(subFolderName);
}

/**
 * {{明細テーブル}} プレースホルダーを検索し、その位置に明細テーブルを挿入する
 * @param {DocumentApp.Body} body
 * @param {Array} items
 * @param {string} transactionDate 取引日（YYYY/MM/DD形式）
 * @private
 */
function insertItemsTable_(body, items, transactionDate) {
  // プレースホルダーの段落を検索
  var searchResult = body.findText('\\{\\{明細テーブル\\}\\}');
  if (!searchResult) {
    // プレースホルダーが見つからない場合は末尾に追加
    appendItemsTableToBody_(body, items, transactionDate);
    return;
  }

  // プレースホルダーの段落を取得
  var element = searchResult.getElement();
  var paragraph = element.getParent();
  var parentIndex = body.getChildIndex(paragraph);

  // プレースホルダー段落を削除
  body.removeChild(paragraph);

  // テーブルデータを構築（5列: 取引日・摘要・数量・単価・明細金額）
  var tableData = buildTableDataV4_(items, transactionDate);

  // 指定位置にテーブルを挿入
  var table = body.insertTable(parentIndex, tableData);

  // ヘッダー行のスタイル
  var headerRow = table.getRow(0);
  for (var c = 0; c < headerRow.getNumCells(); c++) {
    headerRow.getCell(c).editAsText().setBold(true).setFontSize(9);
    headerRow.getCell(c).setBackgroundColor('#4472C4');
    headerRow.getCell(c).editAsText().setForegroundColor('#FFFFFF');
  }

  // データ行のスタイル
  for (var r = 1; r < table.getNumRows(); r++) {
    var row = table.getRow(r);
    for (var cc = 0; cc < row.getNumCells(); cc++) {
      row.getCell(cc).editAsText().setFontSize(9);
    }
  }
}

/**
 * v4テーブルデータ配列を構築（5列: 取引日・摘要・数量・単価・明細金額）
 * @param {Array} items
 * @param {string} transactionDate 取引日
 * @returns {Array} 2D配列
 * @private
 */
function buildTableDataV4_(items, transactionDate) {
  var data = [['取引日', '摘要', '数量', '単価', '明細金額']];
  var txDate = transactionDate || '';

  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    data.push([
      txDate,
      item.name,
      String(item.quantity),
      '¥' + formatNumber_(item.unitPrice),
      '¥' + formatNumber_(item.subtotal)
    ]);
  }

  return data;
}

/**
 * Body末尾にテーブルを追加（フォールバック用）
 * @private
 */
function appendItemsTableToBody_(body, items, transactionDate) {
  var tableData = buildTableDataV4_(items, transactionDate);
  body.appendTable(tableData);
}

// ============================================================================
// データ取得関数
// ============================================================================

/**
 * アクティブ行のデータを取得して返す
 * @returns {Object|null} { sheet, row, projectId, customerName, values }
 * @private
 */
function getActiveRowData_() {
  var sheet = SpreadsheetApp.getActiveSheet();

  if (sheet.getName() !== '案件マスタ') {
    SpreadsheetApp.getUi().alert('「案件マスタ」シートで実行してください');
    return null;
  }

  var row = sheet.getActiveRange().getRow();
  if (row <= 1) {
    SpreadsheetApp.getUi().alert('データ行を選択してください（ヘッダー行は不可）');
    return null;
  }

  var values = sheet.getRange(row, 1, 1, 17).getValues()[0];
  var projectId = values[0];
  var customerName = values[1];

  if (!projectId || !customerName) {
    SpreadsheetApp.getUi().alert('案件IDまたは顧客名が入力されていません');
    return null;
  }

  return {
    sheet: sheet,
    row: row,
    projectId: projectId,
    customerName: customerName,
    values: values
  };
}

/**
 * 設定シートから全設定を取得
 * @returns {Object}
 */
function getSettingsData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('設定');
    if (!sheet) throw new Error('「設定」シートが見つかりません');

    var vals = sheet.getRange('B1:B22').getValues();

    return {
      companyName:        vals[0][0],   // B1: 社名
      representative:     vals[1][0],   // B2: 代表者名
      postalCode:         vals[2][0],   // B3: 郵便番号
      address:            vals[3][0],   // B4: 住所
      tel:                vals[4][0],   // B5: TEL
      email:              vals[5][0],   // B6: メール
      invoiceNumber:      vals[6][0],   // B7: 適格請求書発行事業者番号
      bankName:           vals[7][0],   // B8: 振込先銀行
      branchName:         vals[8][0],   // B9: 振込先支店
      accountType:        vals[9][0],   // B10: 口座種別
      accountNumber:      vals[10][0],  // B11: 口座番号
      accountHolder:      vals[11][0],  // B12: 口座名義
      taxRate:            parseFloat(vals[12][0]) || 0.10, // B13: 消費税率
      orderNote:          vals[13][0],  // B14: 発注書注意書き
      paymentTerms:       vals[14][0],  // B15: 支払条件
      folderId:           vals[15][0],  // B16: 帳票保存先フォルダID
      estimateTemplateId: vals[16][0],  // B17: 見積書テンプレートID
      orderTemplateId:    vals[17][0],  // B18: 発注書テンプレートID
      invoiceTemplateId:  vals[18][0],  // B19: 請求書テンプレートID
      companyStaff:       vals[19][0],  // B20: 自社担当者名
      companyUrl:         vals[20][0],  // B21: 自社URL
      sealImageId:        vals[21][0]   // B22: 社印画像ファイルID
    };
  } catch (error) {
    throw new Error('設定データ取得エラー: ' + error.message);
  }
}

/**
 * 指定された案件IDの明細行を取得
 * @param {string} projectId
 * @returns {Array} [{name, unitPrice, quantity, subtotal}]
 */
function getItemsForProject(projectId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('明細');
    if (!sheet) throw new Error('「明細」シートが見つかりません');

    var data = sheet.getDataRange().getValues();
    var items = [];

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === projectId) {
        var unitPrice = parseFloat(data[i][2]) || 0;
        var quantity = parseFloat(data[i][3]) || 0;
        items.push({
          name: data[i][1] || '',
          unitPrice: unitPrice,
          quantity: quantity,
          subtotal: unitPrice * quantity
        });
      }
    }
    return items;
  } catch (error) {
    throw new Error('明細取得エラー: ' + error.message);
  }
}

/**
 * 顧客マスタから顧客名で検索し、顧客情報を返す
 * シート未存在時や該当なしの場合は空文字を返す（グレースフルデグレード）
 * @param {string} customerName 顧客名
 * @returns {Object} { postalCode, address1, address2, title, contactName }
 * @private
 */
function getCustomerData_(customerName) {
  var empty = { displayName: '', postalCode: '', address1: '', address2: '', title: '', contactName: '' };
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('顧客マスタ');
    if (!sheet) return empty;

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === customerName) {
        return {
          displayName: data[i][1] || '',  // B列: 名称（帳票記載用）
          postalCode:  data[i][2] || '',   // C列: 郵便番号
          address1:    data[i][3] || '',   // D列: 住所1
          address2:    data[i][4] || '',   // E列: 住所2
          title:       data[i][5] || '',   // F列: 役職
          contactName: data[i][6] || ''    // G列: 担当者名
        };
      }
    }
    return empty;
  } catch (e) {
    return empty;
  }
}

// ============================================================================
// 置換データ構築
// ============================================================================

/**
 * テンプレート用の置換辞書を構築
 * プレースホルダー名は templates.gs と完全一致
 *
 * @param {Object} rowData getActiveRowData_() の戻り値
 * @param {Object} settings getSettingsData() の戻り値
 * @param {Array} items getItemsForProject() の戻り値
 * @returns {Object}
 * @private
 */
function buildReplacements_(rowData, settings, items) {
  var vals = rowData.values;
  var taxRate = settings.taxRate || 0.10;

  // 小計計算
  var subtotal = 0;
  for (var i = 0; i < items.length; i++) {
    subtotal += items[i].subtotal;
  }
  var tax = Math.floor(subtotal * taxRate);  // 切り捨て
  var total = subtotal + tax;

  // 顧客マスタからデータ取得
  var customer = getCustomerData_(rowData.customerName);

  var r = {};

  // 自社情報
  r['{{自社名}}']   = settings.companyName;
  r['{{自社住所}}'] = settings.postalCode + ' ' + settings.address;
  r['{{自社TEL}}']  = settings.tel;
  r['{{自社メール}}'] = settings.email;
  r['{{適格番号}}'] = settings.invoiceNumber;
  r['{{自社担当者名}}'] = settings.companyStaff || '';
  r['{{自社URL}}']  = settings.companyUrl || '';

  // 案件情報
  r['{{案件ID}}']   = rowData.projectId;
  r['{{顧客名}}']   = customer.displayName || rowData.customerName;
  r['{{案件名}}']   = vals[2] || '';

  // 顧客マスタ情報
  r['{{顧客郵便番号}}'] = customer.postalCode;
  r['{{顧客住所1}}']   = customer.address1;
  r['{{顧客住所2}}']   = customer.address2;
  r['{{顧客役職}}']     = customer.title;
  r['{{顧客担当者名}}'] = customer.contactName;

  // 日付（Date型で入っている場合を考慮）
  r['{{見積日}}']   = vals[3] ? formatDateValue_(vals[3]) : '';
  r['{{発注日}}']   = vals[4] ? formatDateValue_(vals[4]) : '';
  r['{{請求日}}']   = vals[5] ? formatDateValue_(vals[5]) : '';
  r['{{入金予定日}}'] = vals[6] ? formatDateValue_(vals[6]) : '';

  // 金額
  r['{{小計}}']     = '¥' + formatNumber_(subtotal);
  r['{{消費税}}']   = '¥' + formatNumber_(tax);
  r['{{合計金額}}'] = '¥' + formatNumber_(total);

  // 内訳（10%対象・消費税）
  r['{{10%対象}}']  = '¥' + formatNumber_(subtotal);
  r['{{10%消費税}}'] = '¥' + formatNumber_(tax);

  // 条件・注意書き
  r['{{支払条件}}']     = settings.paymentTerms || '';
  r['{{発注書注意書き}}'] = settings.orderNote || '';
  r['{{振込先情報}}']   = ''; // 請求書生成時にのみ上書きする

  return r;
}

/**
 * {{社印画像}} プレースホルダーを検索し、社印画像を挿入する
 * ID未設定時はプレースホルダーを削除するだけ（エラーにしない）
 * @param {DocumentApp.Body} body
 * @param {string} sealImageId Google DriveのファイルID
 * @private
 */
function insertSealImage_(body, sealImageId) {
  var searchResult = body.findText('\\{\\{社印画像\\}\\}');
  if (!searchResult) return;

  var element = searchResult.getElement();
  var paragraph = element.getParent();

  if (!sealImageId || sealImageId === '') {
    // ID未設定: プレースホルダーテキストを削除
    paragraph.asText().replaceText('\\{\\{社印画像\\}\\}', '');
    return;
  }

  try {
    var file = DriveApp.getFileById(sealImageId);
    var blob = file.getBlob();

    // プレースホルダーテキストを削除
    paragraph.asText().replaceText('\\{\\{社印画像\\}\\}', '');

    // 画像を挿入（80×80px）
    var img = paragraph.asParagraph().appendInlineImage(blob);
    img.setWidth(80);
    img.setHeight(80);
  } catch (e) {
    // 画像取得に失敗した場合はプレースホルダーを削除のみ
    paragraph.asText().replaceText('\\{\\{社印画像\\}\\}', '');
    Logger.log('社印画像挿入スキップ: ' + e.message);
  }
}

// ============================================================================
// 案件ID自動採番
// ============================================================================

/**
 * 案件マスタで案件IDが空欄かつ顧客名が入力済みの行に自動採番
 * 形式: EST-YYYYMM-NNN（月内の連番）
 */
function assignProjectId() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('案件マスタ');
    if (!sheet) throw new Error('「案件マスタ」シートが見つかりません');

    var data = sheet.getDataRange().getValues();
    var count = 0;

    for (var i = 1; i < data.length; i++) {
      var existingId = data[i][0];   // A列: 案件ID
      var customer   = data[i][1];   // B列: 顧客名
      var quoteDate  = data[i][3];   // D列: 見積日

      // 案件IDが空 & 顧客名あり & 見積日あり → 採番対象
      if ((!existingId || existingId === '') && customer && customer !== '' && quoteDate) {
        var newId = generateProjectId_(data, quoteDate);
        sheet.getRange(i + 1, 1).setValue(newId);
        data[i][0] = newId; // 後続の採番で重複しないよう更新
        count++;
      }
    }

    if (count > 0) {
      SpreadsheetApp.getUi().alert(count + '件の案件IDを採番しました');
    } else {
      SpreadsheetApp.getUi().alert('採番対象の案件がありません\n（案件ID空欄 かつ 顧客名・見積日が入力済みの行が必要です）');
    }

  } catch (error) {
    SpreadsheetApp.getUi().alert('案件ID採番エラー: ' + error.message);
  }
}

/**
 * 新しい案件IDを生成する
 * @param {Array} allData 全行データ
 * @param {Date} quoteDate 見積日
 * @returns {string} EST-YYYYMM-NNN
 * @private
 */
function generateProjectId_(allData, quoteDate) {
  var d = new Date(quoteDate);
  var yyyy = d.getFullYear();
  var mm = String(d.getMonth() + 1).padStart(2, '0');
  var prefix = 'EST-' + yyyy + mm + '-';

  // 同月の最大番号を検索
  var maxNum = 0;
  for (var i = 1; i < allData.length; i++) {
    var id = allData[i][0];
    if (id && typeof id === 'string' && id.indexOf(prefix) === 0) {
      var numPart = parseInt(id.substring(prefix.length), 10);
      if (!isNaN(numPart) && numPart > maxNum) {
        maxNum = numPart;
      }
    }
  }

  return prefix + String(maxNum + 1).padStart(3, '0');
}

// ============================================================================
// 税理士CSV出力
// ============================================================================

/**
 * 指定月の請求済/入金済案件をCSVで出力しGoogleドライブに保存
 */
function exportCSVForAccountant() {
  try {
    var ui = SpreadsheetApp.getUi();

    // 対象月を入力
    var response = ui.prompt(
      '税理士CSV出力',
      '対象月を入力してください（例: 2026-01）',
      ui.ButtonSet.OK_CANCEL
    );
    if (response.getSelectedButton() === ui.Button.CANCEL) return;

    var targetMonth = response.getResponseText().trim();
    if (!/^\d{4}-\d{2}$/.test(targetMonth)) {
      ui.alert('形式エラー: YYYY-MM で入力してください');
      return;
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheet = ss.getSheetByName('案件マスタ');
    var itemsSheet = ss.getSheetByName('明細');
    if (!masterSheet || !itemsSheet) throw new Error('シートが見つかりません');

    var masterData = masterSheet.getDataRange().getValues();
    var itemsData  = itemsSheet.getDataRange().getValues();
    var settings   = getSettingsData();
    var taxRate    = settings.taxRate || 0.10;

    // CSVヘッダー
    var csvLines = ['請求日,入金日,顧客名,案件名,品目名,単価,数量,小計,消費税,合計'];

    for (var i = 1; i < masterData.length; i++) {
      var projectId   = masterData[i][0];
      var customerName = masterData[i][1];
      var projectName = masterData[i][2];
      var invoiceDate = masterData[i][5]; // F列
      var paymentDate = masterData[i][7]; // H列（入金日）
      var status      = masterData[i][8]; // I列

      if (!projectId) continue;
      if (status !== '請求済' && status !== '入金済') continue;
      if (!invoiceDate) continue;

      // 対象月フィルタ
      var invDate = new Date(invoiceDate);
      var invYM = invDate.getFullYear() + '-' + String(invDate.getMonth() + 1).padStart(2, '0');
      if (invYM !== targetMonth) continue;

      // 該当案件の明細を取得
      var projectSubtotal = 0;
      var projectItems = [];
      for (var j = 1; j < itemsData.length; j++) {
        if (itemsData[j][0] === projectId) {
          var uPrice = parseFloat(itemsData[j][2]) || 0;
          var qty    = parseFloat(itemsData[j][3]) || 0;
          var sub    = uPrice * qty;
          projectItems.push({name: itemsData[j][1], unitPrice: uPrice, quantity: qty, subtotal: sub});
          projectSubtotal += sub;
        }
      }

      var projectTax   = Math.floor(projectSubtotal * taxRate);
      var projectTotal = projectSubtotal + projectTax;

      // CSV行を追加
      for (var k = 0; k < projectItems.length; k++) {
        var item = projectItems[k];
        var line = [];
        if (k === 0) {
          line.push(formatDateValue_(invoiceDate));
          line.push(paymentDate ? formatDateValue_(paymentDate) : '');
          line.push(escapeCsvField_(customerName));
          line.push(escapeCsvField_(projectName));
        } else {
          line.push('', '', '', '');
        }
        line.push(escapeCsvField_(item.name));
        line.push(item.unitPrice);
        line.push(item.quantity);
        line.push(item.subtotal);
        line.push(k === 0 ? projectTax : '');
        line.push(k === 0 ? projectTotal : '');
        csvLines.push(line.join(','));
      }
    }

    if (csvLines.length <= 1) {
      ui.alert('対象月（' + targetMonth + '）に該当する請求済/入金済の案件がありません');
      return;
    }

    // BOM付きUTF-8（Excel互換）でCSVファイルを生成
    var BOM = '\uFEFF';
    var csvContent = BOM + csvLines.join('\r\n');
    var blob = Utilities.newBlob(csvContent, 'text/csv', '税理士_' + targetMonth + '.csv');

    var folder = DriveApp.getFolderById(settings.folderId);
    var csvFile = folder.createFile(blob);

    ui.alert('CSV出力完了\n\n' + csvFile.getUrl());

  } catch (error) {
    SpreadsheetApp.getUi().alert('CSV出力エラー: ' + error.message);
  }
}

// ============================================================================
// ステータス変更トリガー（onEdit）
// ============================================================================

/**
 * 案件マスタのステータス列（I列 = 9列目）が変更された時に日付を自動入力
 *
 * - 「見積もり提示」→ D列（見積日）にTODAY
 * - 「受注」→ E列（受注日）にTODAY
 * - 「請求済」→ F列（請求日）にTODAY
 * - 「入金済」→ H列（入金日）にTODAY
 *
 * ※ 既に日付が入っている場合は上書きしない
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;

    var sheet = e.range.getSheet();
    if (sheet.getName() !== '案件マスタ') return;

    var col = e.range.getColumn();
    var row = e.range.getRow();
    if (col !== 9 || row <= 1) return; // I列（9列目）のみ対象

    var newStatus = e.value;
    var today = new Date(); // Date型で入力（文字列ではない）

    if (newStatus === '提案中') {
      var currentS = sheet.getRange(row, 19).getValue(); // S列: 提案日
      if (!currentS || currentS === '') {
        sheet.getRange(row, 19).setValue(today);
      }
    } else if (newStatus === '見積もり提示') {
      var currentD = sheet.getRange(row, 4).getValue(); // D列: 見積日
      if (!currentD || currentD === '') {
        sheet.getRange(row, 4).setValue(today);
      }
    } else if (newStatus === '受注') {
      var current = sheet.getRange(row, 5).getValue(); // E列: 受注日
      if (!current || current === '') {
        sheet.getRange(row, 5).setValue(today);
      }
    } else if (newStatus === '請求済') {
      var current2 = sheet.getRange(row, 6).getValue(); // F列: 請求日
      if (!current2 || current2 === '') {
        sheet.getRange(row, 6).setValue(today);
      }
    } else if (newStatus === '入金済') {
      var current3 = sheet.getRange(row, 8).getValue(); // H列: 入金日
      if (!current3 || current3 === '') {
        sheet.getRange(row, 8).setValue(today);
      }
    } else if (newStatus === '失注') {
      var currentT = sheet.getRange(row, 20).getValue(); // T列: 失注日
      if (!currentT || currentT === '') {
        sheet.getRange(row, 20).setValue(today);
      }
    }

  } catch (error) {
    // onEdit では UI.alert() を使わない（simple trigger の制約）
    Logger.log('onEditエラー: ' + error.message);
  }
}

// ============================================================================
// ユーティリティ関数
// ============================================================================

/**
 * 日付値（Date型 or 文字列）を YYYY/MM/DD 形式に変換
 * @param {Date|string} value
 * @returns {string}
 * @private
 */
function formatDateValue_(value) {
  if (!value || value === '') return '';
  var d = new Date(value);
  if (isNaN(d.getTime())) return String(value);
  var yyyy = d.getFullYear();
  var mm = String(d.getMonth() + 1).padStart(2, '0');
  var dd = String(d.getDate()).padStart(2, '0');
  return yyyy + '/' + mm + '/' + dd;
}

/**
 * 数値をカンマ区切り文字列に変換（切り捨て済み）
 * @param {number} num
 * @returns {string}
 * @private
 */
function formatNumber_(num) {
  if (num === null || num === undefined || isNaN(num)) return '0';
  return Math.floor(num).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

/**
 * 正規表現のメタ文字をエスケープ
 * @param {string} str
 * @returns {string}
 * @private
 */
function escapeRegExp_(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * CSVフィールドをエスケープ（カンマ・改行・ダブルクォートを含む場合）
 * @param {string} field
 * @returns {string}
 * @private
 */
function escapeCsvField_(field) {
  if (field === null || field === undefined) return '';
  var str = String(field);
  if (str.indexOf(',') >= 0 || str.indexOf('"') >= 0 || str.indexOf('\n') >= 0) {
    return '"' + str.replace(/"/g, '""') + '"';
  }
  return str;
}
