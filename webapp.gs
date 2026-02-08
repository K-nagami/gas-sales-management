/**
 * 売上管理統合システム - Web App サーバーサイド v2
 *
 * v2変更点:
 *   - 7段階ステータス対応（商談見込み→提案中→見積もり提示→受注→請求済→入金済/失注）
 *   - ヨミランク（A/B/C/D）による加重売上予測
 *   - ファネル分析データ
 *   - 入金予定vs実績ギャップ分析
 */

// ============================================================================
// Web App エントリーポイント
// ============================================================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('webapp')
    .setTitle('売上管理ダッシュボード');
}

// ============================================================================
// 定数
// ============================================================================

var STATUSES = ['商談見込み', '提案中', '見積もり提示', '受注', '請求済', '入金済', '失注'];
var ACTIVE_PIPELINE = ['商談見込み', '提案中', '見積もり提示', '受注', '請求済', '入金済'];
var ACTIVE_STATUSES = ['商談見込み', '提案中', '見積もり提示', '受注'];
var YOMI_RATES = { 'A': 0.90, 'B': 0.70, 'C': 0.50, 'D': 0.30 };

/**
 * 決算月1月（期首2月）ベースの期間範囲を返す
 * @param {string} period - 'month','1q','2q','3q','4q','year'
 * @returns {{start: Date, end: Date}}
 */
function getFiscalDateRange_(period) {
  var now = new Date();
  var y = now.getFullYear();
  var m = now.getMonth(); // 0-based (0=Jan)

  // 今期の期首年: 2月始まりなので、1月なら前年が期首年
  var fyStart = m >= 1 ? y : y - 1; // m>=1 means Feb(1)~Dec(11) → current year, Jan(0) → prev year

  var ranges = {
    'month': { start: new Date(y, m, 1), end: new Date(y, m + 1, 0, 23, 59, 59) },
    '1q':    { start: new Date(fyStart, 1, 1), end: new Date(fyStart, 4, 0, 23, 59, 59) },   // 2-4月
    '2q':    { start: new Date(fyStart, 4, 1), end: new Date(fyStart, 7, 0, 23, 59, 59) },   // 5-7月
    '3q':    { start: new Date(fyStart, 7, 1), end: new Date(fyStart, 10, 0, 23, 59, 59) },  // 8-10月
    '4q':    { start: new Date(fyStart, 10, 1), end: new Date(fyStart + 1, 1, 0, 23, 59, 59) }, // 11-1月
    'year':  { start: new Date(fyStart, 1, 1), end: new Date(fyStart + 1, 1, 0, 23, 59, 59) }  // 2月-1月
  };

  return ranges[period] || ranges['year'];
}

// ============================================================================
// データ取得API
// ============================================================================

/**
 * ステータス別の基準日を返す
 * 各ステータスが「いつ」そのステータスになったかを判定する日付カラムのマッピング
 *
 * | ステータス     | 基準日列       | row index |
 * |---------------|---------------|-----------|
 * | 商談見込み     | R列(登録日)    | row[17]   |
 * | 提案中        | S列(提案日)    | row[18]   |
 * | 見積もり提示   | D列(見積日)    | row[3]    |
 * | 受注          | E列(受注日)    | row[4]    |
 * | 請求済        | F列(請求日)    | row[5]    |
 * | 入金済        | H列(入金日)    | row[7]    |
 * | 失注          | T列(失注日)    | row[19]   |
 */
function getStatusBaseDate_(row, status) {
  var map = {
    '商談見込み':   row[17],  // R列: 登録日
    '提案中':       row[18],  // S列: 提案日
    '見積もり提示': row[3],   // D列: 見積日
    '受注':         row[4],   // E列: 受注日
    '請求済':       row[5],   // F列: 請求日
    '入金済':       row[7],   // H列: 入金日
    '失注':         row[19]   // T列: 失注日
  };
  var d = map[status];
  return d ? new Date(d) : null;
}

/**
 * 日付が期間範囲内かチェック
 */
function isInRange_(dateVal, dateRange) {
  if (!dateVal) return false;
  var d = new Date(dateVal);
  if (isNaN(d.getTime())) return false;
  return d >= dateRange.start && d <= dateRange.end;
}

/**
 * ダッシュボード用データをまとめて返す（v3: 現ステータス基準日方式）
 *
 * KPI/ファネル共通ルール:
 *   各案件の「現在のステータスに対応する基準日」が期間内か判定
 *   例: ステータス=受注 → E列(受注日)が期間内なら集計対象
 */
function getDashboardData(period) {
  try {
    period = period || 'year';
    var dateRange = getFiscalDateRange_(period);

    var ss = getSpreadsheet_();
    var masterSheet = ss.getSheetByName('案件マスタ');
    var settingsSheet = ss.getSheetByName('設定');

    var masterData = masterSheet.getDataRange().getValues();
    var taxRate = parseFloat(settingsSheet.getRange('B13').getValue()) || 0.10;

    var now = new Date();
    var thisYear = now.getFullYear();
    var thisMonth = now.getMonth(); // 0-based

    // 初期化
    var kpi = {
      totalProjects: 0,
      activeProjects: 0,
      confirmedRevenue: 0,
      weightedPipeline: 0,
      totalForecast: 0,
      unpaidAmount: 0,
      lostCount: 0,
      orderRate: 0
    };

    var funnel = [];
    for (var f = 0; f < ACTIVE_PIPELINE.length; f++) {
      funnel.push({ stage: ACTIVE_PIPELINE[f], count: 0, amount: 0 });
    }

    var forecast = {
      confirmed: 0,
      weightedPipeline: 0,
      totalForecast: 0,
      byRank: { A: 0, B: 0, C: 0, D: 0, unset: 0 }
    };

    var monthlyData = {};
    var paymentGapData = {};
    var customerData = {};
    var overdueAlerts = [];
    var recentProjects = [];

    // 受注率: D列(見積日)が期間内の件数 / E列(受注日)が期間内の件数
    var estimatesInPeriod = 0;
    var ordersInPeriod = 0;

    for (var i = 1; i < masterData.length; i++) {
      var row = masterData[i];
      if (!row[0]) continue;

      var projectId = row[0];
      var customer = row[1];
      var projectName = row[2];
      var quoteDate = row[3];
      var orderDate = row[4];
      var invoiceDate = row[5];
      var dueDate = row[6];
      var paymentDate = row[7];
      var status = row[8];
      var subtotal = parseFloat(row[9]) || 0;
      var tax = Math.floor(subtotal * taxRate);
      var total = subtotal + tax;
      var yomiRank = row[16] || '';

      // ★ 現ステータスの基準日を取得
      var baseDate = getStatusBaseDate_(row, status);
      var inPeriod = isInRange_(baseDate, dateRange);

      // === KPI集計（期間内のみ） ===
      if (inPeriod) {
        // 総案件数 = 失注以外
        if (status !== '失注') {
          kpi.totalProjects++;
        }
        // アクティブ = 商談見込み〜受注
        if (ACTIVE_STATUSES.indexOf(status) !== -1) {
          kpi.activeProjects++;
        }
        // 失注カウント
        if (status === '失注') {
          kpi.lostCount++;
        }
        // 未入金（請求済）: F列が期間内
        if (status === '請求済') {
          kpi.unpaidAmount += total;
        }
        // ヨミ管理（売上予測）
        if (status === '受注' || status === '請求済' || status === '入金済') {
          forecast.confirmed += total;
          kpi.confirmedRevenue += total;
        } else if (status !== '失注') {
          var rate = YOMI_RATES[yomiRank] || 0;
          var weighted = Math.floor(total * rate);
          forecast.weightedPipeline += weighted;
          if (yomiRank && forecast.byRank[yomiRank] !== undefined) {
            forecast.byRank[yomiRank] += total;
          } else {
            forecast.byRank.unset += total;
          }
        }
      }

      // === ファネル集計（期間フィルタあり・ステータス別基準日で判定） ===
      for (var fi = 0; fi < ACTIVE_PIPELINE.length; fi++) {
        if (ACTIVE_PIPELINE[fi] === status) {
          if (inPeriod) {
            funnel[fi].count++;
            funnel[fi].amount += total;
          }
          break;
        }
      }

      // === 受注率: D列(見積日)が期間内 → 分母, E列(受注日)が期間内 → 分子 ===
      if (isInRange_(quoteDate, dateRange)) {
        estimatesInPeriod++;
      }
      if (isInRange_(orderDate, dateRange) && (status === '受注' || status === '請求済' || status === '入金済')) {
        ordersInPeriod++;
      }

      // === 以下はグラフ・アラート等（期間フィルタなし・全件対象） ===

      // 入金遅延アラート
      if (dueDate && !paymentDate && status !== '入金済' && status !== '失注') {
        var due = new Date(dueDate);
        if (due < now) {
          var daysOverdue = Math.floor((now - due) / (1000 * 60 * 60 * 24));
          overdueAlerts.push({
            projectId: projectId,
            customer: customer,
            projectName: projectName,
            dueDate: formatDate_(dueDate),
            total: total,
            daysOverdue: daysOverdue
          });
        }
      }

      // 顧客別売上
      if (!customerData[customer]) {
        customerData[customer] = { total: 0, count: 0 };
      }
      customerData[customer].total += total;
      customerData[customer].count++;

      // 月別データ（過去12ヶ月）
      if (quoteDate) {
        var qd2 = new Date(quoteDate);
        var monthKey = qd2.getFullYear() + '-' + String(qd2.getMonth() + 1).padStart(2, '0');
        if (!monthlyData[monthKey]) {
          monthlyData[monthKey] = { estimates: 0, orders: 0, revenue: 0, paid: 0 };
        }
        monthlyData[monthKey].estimates++;
      }
      if (orderDate && (status === '受注' || status === '請求済' || status === '入金済')) {
        var od2 = new Date(orderDate);
        var mKey2 = od2.getFullYear() + '-' + String(od2.getMonth() + 1).padStart(2, '0');
        if (!monthlyData[mKey2]) monthlyData[mKey2] = { estimates: 0, orders: 0, revenue: 0, paid: 0 };
        monthlyData[mKey2].orders++;
      }
      if (invoiceDate && (status === '請求済' || status === '入金済')) {
        var id2 = new Date(invoiceDate);
        var mKey3 = id2.getFullYear() + '-' + String(id2.getMonth() + 1).padStart(2, '0');
        if (!monthlyData[mKey3]) monthlyData[mKey3] = { estimates: 0, orders: 0, revenue: 0, paid: 0 };
        monthlyData[mKey3].revenue += total;
      }
      if (paymentDate && status === '入金済') {
        var pd2 = new Date(paymentDate);
        var mKey4 = pd2.getFullYear() + '-' + String(pd2.getMonth() + 1).padStart(2, '0');
        if (!monthlyData[mKey4]) monthlyData[mKey4] = { estimates: 0, orders: 0, revenue: 0, paid: 0 };
        monthlyData[mKey4].paid += total;
      }

      // 入金ギャップ（入金予定日ベース = planned、入金日ベース = actual）
      if (dueDate && (status === '請求済' || status === '入金済')) {
        var dd = new Date(dueDate);
        var pgKey = dd.getFullYear() + '-' + String(dd.getMonth() + 1).padStart(2, '0');
        if (!paymentGapData[pgKey]) paymentGapData[pgKey] = { planned: 0, actual: 0 };
        paymentGapData[pgKey].planned += total;
      }
      if (paymentDate && status === '入金済') {
        var pd3 = new Date(paymentDate);
        var pgKey2 = pd3.getFullYear() + '-' + String(pd3.getMonth() + 1).padStart(2, '0');
        if (!paymentGapData[pgKey2]) paymentGapData[pgKey2] = { planned: 0, actual: 0 };
        paymentGapData[pgKey2].actual += total;
      }

      // 案件一覧
      recentProjects.push({
        projectId: projectId,
        customer: customer,
        projectName: projectName,
        quoteDate: quoteDate ? formatDate_(quoteDate) : '',
        status: status,
        total: total,
        yomiRank: yomiRank
      });
    }

    // 受注率
    kpi.orderRate = estimatesInPeriod > 0 ? Math.round((ordersInPeriod / estimatesInPeriod) * 100) : 0;

    // ヨミ合計
    forecast.totalForecast = forecast.confirmed + forecast.weightedPipeline;
    kpi.weightedPipeline = forecast.weightedPipeline;
    kpi.totalForecast = forecast.totalForecast;

    // ファネル転換率を算出
    for (var fc = 1; fc < funnel.length; fc++) {
      funnel[fc].conversionRate = funnel[fc - 1].count > 0
        ? Math.round((funnel[fc].count / (funnel[fc - 1].count + funnel[fc].count)) * 100)
        : 0;
    }
    if (funnel.length > 0) funnel[0].conversionRate = 100;

    // 月別データを配列に変換（過去12ヶ月分、ソート済み）
    var monthlyArray = [];
    for (var m = 11; m >= 0; m--) {
      var td = new Date(thisYear, thisMonth - m, 1);
      var mk = td.getFullYear() + '-' + String(td.getMonth() + 1).padStart(2, '0');
      var md = monthlyData[mk] || { estimates: 0, orders: 0, revenue: 0, paid: 0 };
      monthlyArray.push({ month: mk, estimates: md.estimates, orders: md.orders, revenue: md.revenue, paid: md.paid });
    }

    // 入金ギャップを配列に変換（過去6ヶ月 + 今後3ヶ月）
    var paymentGapArray = [];
    for (var pg = 5; pg >= -3; pg--) {
      var pgd = new Date(thisYear, thisMonth - pg, 1);
      var pgk = pgd.getFullYear() + '-' + String(pgd.getMonth() + 1).padStart(2, '0');
      var pgv = paymentGapData[pgk] || { planned: 0, actual: 0 };
      paymentGapArray.push({
        month: pgk,
        planned: pgv.planned,
        actual: pgv.actual,
        gap: pgv.planned - pgv.actual
      });
    }

    // 顧客別を配列に変換
    var customerArray = [];
    for (var c in customerData) {
      customerArray.push({ name: c, total: customerData[c].total, count: customerData[c].count });
    }
    customerArray.sort(function(a, b) { return b.total - a.total; });

    return {
      kpi: kpi,
      forecast: forecast,
      period: period,
      funnel: funnel,
      monthly: monthlyArray,
      paymentGap: paymentGapArray,
      customers: customerArray,
      overdueAlerts: overdueAlerts,
      recentProjects: recentProjects
    };

  } catch (error) {
    throw new Error('ダッシュボードデータ取得エラー: ' + error.message);
  }
}

/**
 * 案件一覧を取得
 */
function getProjectsList() {
  try {
    var ss = getSpreadsheet_();
    var sheet = ss.getSheetByName('案件マスタ');
    var data = sheet.getDataRange().getValues();
    var settingsSheet = ss.getSheetByName('設定');
    var taxRate = parseFloat(settingsSheet.getRange('B13').getValue()) || 0.10;

    var projects = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      var subtotal = parseFloat(data[i][9]) || 0;
      projects.push({
        row: i + 1,
        projectId: data[i][0],
        customer: data[i][1],
        projectName: data[i][2],
        quoteDate: data[i][3] ? formatDate_(data[i][3]) : '',
        orderDate: data[i][4] ? formatDate_(data[i][4]) : '',
        invoiceDate: data[i][5] ? formatDate_(data[i][5]) : '',
        dueDate: data[i][6] ? formatDate_(data[i][6]) : '',
        paymentDate: data[i][7] ? formatDate_(data[i][7]) : '',
        status: data[i][8],
        subtotal: subtotal,
        tax: Math.floor(subtotal * taxRate),
        total: subtotal + Math.floor(subtotal * taxRate),
        estimateUrl: data[i][12],
        orderUrl: data[i][13],
        invoiceUrl: data[i][14],
        memo: data[i][15],
        yomiRank: data[i][16] || ''
      });
    }
    return projects;
  } catch (error) {
    throw new Error('案件一覧取得エラー: ' + error.message);
  }
}

/**
 * 指定案件の明細を取得
 */
function getItemsForProjectWeb(projectId) {
  try {
    var ss = getSpreadsheet_();
    var sheet = ss.getSheetByName('明細');
    var data = sheet.getDataRange().getValues();
    var items = [];

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === projectId) {
        var unitPrice = parseFloat(data[i][2]) || 0;
        var quantity = parseFloat(data[i][3]) || 0;
        items.push({
          row: i + 1,
          name: data[i][1],
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

// ============================================================================
// データ書き込みAPI
// ============================================================================

/**
 * 新規案件を登録
 */
function addNewProject(formData) {
  try {
    var ss = getSpreadsheet_();
    var sheet = ss.getSheetByName('案件マスタ');
    var data = sheet.getDataRange().getValues();

    // 案件ID生成
    var quoteDate = new Date(formData.quoteDate);
    var projectId = generateProjectId_(data, quoteDate);

    // 新行を追加
    var newRow = sheet.getLastRow() + 1;
    var today = new Date();
    sheet.getRange(newRow, 1).setValue(projectId);
    sheet.getRange(newRow, 2).setValue(formData.customer);
    sheet.getRange(newRow, 3).setValue(formData.projectName);
    sheet.getRange(newRow, 4).setValue(quoteDate);
    if (formData.dueDate) sheet.getRange(newRow, 7).setValue(new Date(formData.dueDate));
    var initStatus = formData.status || '商談見込み';
    sheet.getRange(newRow, 9).setValue(initStatus);
    sheet.getRange(newRow, 16).setValue(formData.memo || '');
    if (formData.yomiRank) sheet.getRange(newRow, 17).setValue(formData.yomiRank);

    // R列（登録日）: 案件登録時に常にセット
    sheet.getRange(newRow, 18).setValue(today);

    // S列（提案日）: 初期ステータスが提案中以降ならセット
    if (initStatus === '提案中') {
      sheet.getRange(newRow, 19).setValue(today);
    }
    // T列（失注日）: 初期ステータスが失注ならセット
    if (initStatus === '失注') {
      sheet.getRange(newRow, 20).setValue(today);
    }

    // J/K/L列の数式
    sheet.getRange(newRow, 10).setFormula('=SUMIF(明細!A:A,A' + newRow + ',明細!E:E)');
    sheet.getRange(newRow, 11).setFormula('=FLOOR(J' + newRow + '*設定!B13)');
    sheet.getRange(newRow, 12).setFormula('=J' + newRow + '+K' + newRow);

    return { success: true, projectId: projectId };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * 明細行を追加
 */
function addNewItem(formData) {
  try {
    var ss = getSpreadsheet_();
    var sheet = ss.getSheetByName('明細');
    var newRow = sheet.getLastRow() + 1;

    sheet.getRange(newRow, 1).setValue(formData.projectId);
    sheet.getRange(newRow, 2).setValue(formData.itemName);
    sheet.getRange(newRow, 3).setValue(parseFloat(formData.unitPrice) || 0);
    sheet.getRange(newRow, 4).setValue(parseFloat(formData.quantity) || 0);
    sheet.getRange(newRow, 5).setFormula('=C' + newRow + '*D' + newRow);

    return { success: true };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * ステータスを更新
 */
function updateProjectStatus(projectId, newStatus) {
  try {
    var ss = getSpreadsheet_();
    var sheet = ss.getSheetByName('案件マスタ');
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === projectId) {
        var row = i + 1;
        sheet.getRange(row, 9).setValue(newStatus);

        // 日付自動入力
        var today = new Date();
        if (newStatus === '提案中' && !data[i][18]) {
          sheet.getRange(row, 19).setValue(today);   // S列: 提案日
        } else if (newStatus === '見積もり提示' && !data[i][3]) {
          sheet.getRange(row, 4).setValue(today);    // D列: 見積日
        } else if (newStatus === '受注' && !data[i][4]) {
          sheet.getRange(row, 5).setValue(today);    // E列: 受注日
        } else if (newStatus === '請求済' && !data[i][5]) {
          sheet.getRange(row, 6).setValue(today);    // F列: 請求日
        } else if (newStatus === '入金済' && !data[i][7]) {
          sheet.getRange(row, 8).setValue(today);    // H列: 入金日
        } else if (newStatus === '失注' && !data[i][19]) {
          sheet.getRange(row, 20).setValue(today);   // T列: 失注日
        }

        return { success: true };
      }
    }
    return { success: false, message: '案件が見つかりません' };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * ヨミランクを更新
 */
function updateYomiRank(projectId, rank) {
  try {
    var ss = getSpreadsheet_();
    var sheet = ss.getSheetByName('案件マスタ');
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === projectId) {
        sheet.getRange(i + 1, 17).setValue(rank); // Q列
        return { success: true };
      }
    }
    return { success: false, message: '案件が見つかりません' };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * 備考（メモ）を更新
 */
function updateMemo(projectId, memo) {
  try {
    var ss = getSpreadsheet_();
    var sheet = ss.getSheetByName('案件マスタ');
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === projectId) {
        sheet.getRange(i + 1, 16).setValue(memo); // P列（備考）
        return { success: true };
      }
    }
    return { success: false, message: '案件が見つかりません' };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * ステータス・ヨミランク・備考を一括更新
 */
function updateProjectBulk(projectId, updates) {
  try {
    var ss = getSpreadsheet_();
    var sheet = ss.getSheetByName('案件マスタ');
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === projectId) {
        var row = i + 1;
        var changed = [];

        // ステータス更新
        if (updates.status && updates.status !== data[i][8]) {
          sheet.getRange(row, 9).setValue(updates.status);
          var today = new Date();
          if (updates.status === '提案中' && !data[i][18]) {
            sheet.getRange(row, 19).setValue(today); // S列: 提案日
          } else if (updates.status === '見積もり提示' && !data[i][3]) {
            sheet.getRange(row, 4).setValue(today);  // D列: 見積日
          } else if (updates.status === '受注' && !data[i][4]) {
            sheet.getRange(row, 5).setValue(today);  // E列: 受注日
          } else if (updates.status === '請求済' && !data[i][5]) {
            sheet.getRange(row, 6).setValue(today);  // F列: 請求日
          } else if (updates.status === '入金済' && !data[i][7]) {
            sheet.getRange(row, 8).setValue(today);  // H列: 入金日
          } else if (updates.status === '失注' && !data[i][19]) {
            sheet.getRange(row, 20).setValue(today);  // T列: 失注日
          }
          changed.push('ステータス');
        }

        // ヨミランク更新
        if (updates.yomiRank !== undefined && updates.yomiRank !== (data[i][16] || '')) {
          sheet.getRange(row, 17).setValue(updates.yomiRank);
          changed.push('ヨミランク');
        }

        // 備考更新
        if (updates.memo !== undefined && updates.memo !== (data[i][15] || '')) {
          sheet.getRange(row, 16).setValue(updates.memo);
          changed.push('備考');
        }

        if (changed.length === 0) {
          return { success: true, message: '変更なし' };
        }
        return { success: true, message: changed.join('・') + 'を更新しました' };
      }
    }
    return { success: false, message: '案件が見つかりません' };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * PDF生成をトリガー（Web App経由）
 */
function generatePDFFromWeb(projectId, docType) {
  try {
    var ss = getSpreadsheet_();
    var masterSheet = ss.getSheetByName('案件マスタ');
    var data = masterSheet.getDataRange().getValues();

    var rowIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === projectId) {
        rowIndex = i;
        break;
      }
    }
    if (rowIndex === -1) return { success: false, message: '案件が見つかりません' };

    var values = data[rowIndex];
    var settings = getSettingsData();
    var items = getItemsForProject(projectId);
    if (items.length === 0) return { success: false, message: '明細がありません' };

    var rowData = {
      sheet: masterSheet,
      row: rowIndex + 1,
      projectId: projectId,
      customerName: values[1],
      values: values
    };

    var replacements = buildReplacements_(rowData, settings, items);
    var templateId, fileName, urlCol;

    if (docType === 'estimate') {
      templateId = settings.estimateTemplateId;
      fileName = '見積書_' + projectId + '_' + values[1] + '.pdf';
      urlCol = 13;
    } else if (docType === 'order') {
      templateId = settings.orderTemplateId;
      fileName = '発注書_' + projectId + '_' + values[1] + '.pdf';
      replacements['{{発注書注意書き}}'] = settings.orderNote;
      urlCol = 14;
    } else if (docType === 'invoice') {
      templateId = settings.invoiceTemplateId;
      fileName = '請求書_' + projectId + '_' + values[1] + '.pdf';
      replacements['{{振込先情報}}'] = settings.bankName + ' ' + settings.branchName + '\n' +
        settings.accountType + ' ' + settings.accountNumber + '\n口座名義: ' + settings.accountHolder;
      replacements['{{入金予定日}}'] = values[6] ? formatDate_(values[6]) : '（未設定）';
      urlCol = 15;
    } else {
      return { success: false, message: '不明な帳票種別' };
    }

    var pdfUrl = fillTemplateAndConvertToPDF(templateId, replacements, items, fileName, settings.folderId);
    masterSheet.getRange(rowIndex + 1, urlCol).setValue(pdfUrl);

    return { success: true, url: pdfUrl };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// ============================================================================
// 案件削除
// ============================================================================

/**
 * 案件と関連明細を削除（ハードデリート）
 * 明細シート → 案件マスタ の順で削除（行ずれ防止のため下から上へ）
 */
function deleteProject(projectId) {
  try {
    var ss = getSpreadsheet_();
    var masterSheet = ss.getSheetByName('案件マスタ');
    var itemSheet = ss.getSheetByName('明細');

    // 案件マスタから対象行を特定
    var masterData = masterSheet.getDataRange().getValues();
    var masterRowIndex = -1;
    for (var i = 1; i < masterData.length; i++) {
      if (masterData[i][0] === projectId) {
        masterRowIndex = i + 1; // シート行番号（1-based）
        break;
      }
    }
    if (masterRowIndex === -1) {
      return { success: false, message: '案件が見つかりません: ' + projectId };
    }

    // 明細シートから該当行を検索（下から上へ削除するため行番号を収集）
    var itemData = itemSheet.getDataRange().getValues();
    var itemRowsToDelete = [];
    for (var j = 1; j < itemData.length; j++) {
      if (itemData[j][0] === projectId) {
        itemRowsToDelete.push(j + 1); // シート行番号
      }
    }

    // 明細を下から上へ削除（インデックスずれ防止）
    for (var k = itemRowsToDelete.length - 1; k >= 0; k--) {
      itemSheet.deleteRow(itemRowsToDelete[k]);
    }

    // 案件マスタの行を削除
    masterSheet.deleteRow(masterRowIndex);

    return {
      success: true,
      message: projectId + ' を削除しました（明細 ' + itemRowsToDelete.length + '行含む）',
      deletedItemCount: itemRowsToDelete.length
    };
  } catch (error) {
    return { success: false, message: '削除エラー: ' + error.message };
  }
}

// ============================================================================
// 一括請求書PDF生成
// ============================================================================

/**
 * 複数案件の請求書を一括生成
 * @param {string[]} projectIds - 案件IDの配列
 * @returns {Object} 生成結果サマリー
 */
function generateBulkInvoices(projectIds) {
  try {
    var ss = getSpreadsheet_();
    var masterSheet = ss.getSheetByName('案件マスタ');
    var masterData = masterSheet.getDataRange().getValues();
    var settings = getSettingsData();

    var results = [];
    var errors = [];

    for (var p = 0; p < projectIds.length; p++) {
      var pid = projectIds[p];
      try {
        // 案件マスタから対象行を検索
        var rowIndex = -1;
        for (var i = 1; i < masterData.length; i++) {
          if (masterData[i][0] === pid) {
            rowIndex = i;
            break;
          }
        }
        if (rowIndex === -1) {
          errors.push({ projectId: pid, reason: '案件が見つかりません' });
          continue;
        }

        var values = masterData[rowIndex];
        var items = getItemsForProject(pid);
        if (items.length === 0) {
          errors.push({ projectId: pid, reason: '明細がありません' });
          continue;
        }

        var rowData = {
          sheet: masterSheet,
          row: rowIndex + 1,
          projectId: pid,
          customerName: values[1],
          values: values
        };

        var replacements = buildReplacements_(rowData, settings, items);
        replacements['{{振込先情報}}'] = settings.bankName + ' ' + settings.branchName + '\n' +
          settings.accountType + ' ' + settings.accountNumber + '\n口座名義: ' + settings.accountHolder;
        replacements['{{入金予定日}}'] = values[6] ? formatDate_(values[6]) : '（未設定）';

        var fileName = '請求書_' + pid + '_' + values[1] + '.pdf';
        var pdfUrl = fillTemplateAndConvertToPDF(settings.invoiceTemplateId, replacements, items, fileName, settings.folderId);

        // O列（15列目）に請求書URLを保存
        masterSheet.getRange(rowIndex + 1, 15).setValue(pdfUrl);

        results.push({ projectId: pid, url: pdfUrl });
      } catch (innerError) {
        errors.push({ projectId: pid, reason: innerError.message });
      }

      // API制限回避: 500ms待機
      if (p < projectIds.length - 1) {
        Utilities.sleep(500);
      }
    }

    return {
      success: true,
      total: projectIds.length,
      completed: results.length,
      failed: errors.length,
      results: results,
      errors: errors
    };
  } catch (error) {
    return { success: false, message: '一括生成エラー: ' + error.message, total: 0, completed: 0, failed: 0, results: [], errors: [] };
  }
}

// ============================================================================
// ユーティリティ
// ============================================================================

/**
 * スプレッドシートを取得（バインドスクリプト or IDで取得）
 * @private
 */
function getSpreadsheet_() {
  try {
    return SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {
    var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if (id) return SpreadsheetApp.openById(id);
    throw new Error('スプレッドシートが見つかりません。SPREADSHEET_IDをスクリプトプロパティに設定してください。');
  }
}

/**
 * 日付フォーマット
 * @private
 */
function formatDate_(value) {
  if (!value) return '';
  var d = new Date(value);
  if (isNaN(d.getTime())) return String(value);
  return d.getFullYear() + '/' + String(d.getMonth() + 1).padStart(2, '0') + '/' + String(d.getDate()).padStart(2, '0');
}
