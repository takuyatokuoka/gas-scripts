/**
 * summary_dashboard.gs
 *
 * 売上データを月ごとに集計し、「月次サマリー」シートに書き込んで
 * 棒グラフで可視化するスクリプト。
 * 毎朝9時の自動実行トリガーは setupDailyTrigger() を一度だけ手動実行して設定する。
 */

// ===== 定数 =====
var SHEET_DATA    = '売上データ';
var SHEET_SUMMARY = '月次サマリー';
var CHART_TITLE   = '月次売上推移';

// ===== メイン処理 =====

/**
 * 売上データを集計して月次サマリーを作成し、棒グラフを更新する。
 * トリガーからも手動実行からも呼び出せる。
 */
function createSummaryDashboard() {
  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet   = ss.getSheetByName(SHEET_DATA);
  var summarySheet = getOrCreateSheet(ss, SHEET_SUMMARY);

  if (!dataSheet) {
    throw new Error('「' + SHEET_DATA + '」シートが見つかりません。');
  }

  // 売上データを月ごとに集計する
  var monthlyData = aggregateByMonth(dataSheet);

  // 月次サマリーシートを書き直す
  writeSummary(summarySheet, monthlyData);

  // 棒グラフを作成／更新する
  updateBarChart(summarySheet);

  Logger.log('月次サマリーの作成が完了しました。');
}

// ===== 集計処理 =====

/**
 * 「売上データ」シートの全行を読み込み、月ごとに合計売上と件数を集計する。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dataSheet
 * @returns {Array<{month: string, total: number, count: number}>} 月順にソートされた集計結果
 */
function aggregateByMonth(dataSheet) {
  var lastRow = dataSheet.getLastRow();

  // ヘッダー行を除いたデータがない場合は空を返す
  if (lastRow < 2) {
    return [];
  }

  // A〜D列のデータをまとめて取得（1行目はヘッダーなので2行目から）
  var values = dataSheet.getRange(2, 1, lastRow - 1, 4).getValues();

  // 月ごとの集計を保持するオブジェクト（キー: "YYYY-MM"）
  var map = {};

  values.forEach(function(row) {
    var dateVal = row[0];
    var amount  = row[3];

    // 日付・金額が空の行はスキップ
    if (!dateVal || amount === '' || amount === null) return;

    // Date オブジェクトに変換（文字列・Date 両対応）
    var date = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
    if (isNaN(date.getTime())) return;

    var year  = date.getFullYear();
    var month = date.getMonth() + 1; // 0始まりを補正
    var key   = year + '-' + zeroPad(month); // ソート用キー（例: "2026-01"）
    var label = year + '年' + month + '月';  // 表示用ラベル

    if (!map[key]) {
      map[key] = { month: label, total: 0, count: 0 };
    }
    map[key].total += Number(amount);
    map[key].count += 1;
  });

  // キーで昇順ソートして配列に変換
  return Object.keys(map)
    .sort()
    .map(function(key) { return map[key]; });
}

// ===== サマリー書き込み =====

/**
 * 月次サマリーシートをクリアしてヘッダーとデータを書き込む。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} summarySheet
 * @param {Array} monthlyData aggregateByMonth() の戻り値
 */
function writeSummary(summarySheet, monthlyData) {
  // シートを全クリア
  summarySheet.clearContents();

  // ヘッダー行を書き込む
  summarySheet.getRange(1, 1, 1, 3).setValues([['月', '合計売上', '件数']]);

  if (monthlyData.length === 0) {
    Logger.log('集計対象のデータがありませんでした。');
    return;
  }

  // データ行を一括書き込み
  var rows = monthlyData.map(function(item) {
    return [item.month, item.total, item.count];
  });
  summarySheet.getRange(2, 1, rows.length, 3).setValues(rows);

  // 合計売上列を通貨書式で整形
  summarySheet.getRange(2, 2, rows.length, 1)
    .setNumberFormat('¥#,##0');
}

// ===== グラフ作成／更新 =====

/**
 * 月次サマリーシートの合計売上を棒グラフで可視化する。
 * 既存の同タイトルのグラフがあれば削除してから再作成する。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} summarySheet
 */
function updateBarChart(summarySheet) {
  // 既存グラフを削除する（タイトルが一致するものだけ）
  summarySheet.getCharts().forEach(function(chart) {
    if (chart.getOptions().get('title') === CHART_TITLE) {
      summarySheet.removeChart(chart);
    }
  });

  var lastRow = summarySheet.getLastRow();

  // データ行がなければグラフは作成しない
  if (lastRow < 2) return;

  // A列（月）と B列（合計売上）を使用
  var dataRange = summarySheet.getRange(1, 1, lastRow, 2);

  var chart = summarySheet.newChart()
    .setChartType(Charts.ChartType.COLUMN) // 縦棒グラフ
    .addRange(dataRange)
    .setOption('title', CHART_TITLE)
    .setOption('hAxis.title', '月')
    .setOption('vAxis.title', '合計売上（円）')
    .setOption('legend.position', 'none')
    .setOption('width', 600)
    .setOption('height', 400)
    .setPosition(2, 5, 0, 0) // E2 セルの位置に配置
    .build();

  summarySheet.insertChart(chart);
}

// ===== トリガー設定 =====

/**
 * 毎朝9時に createSummaryDashboard() を自動実行するトリガーを登録する。
 *
 * 使い方：
 *   GAS エディタでこの関数を一度だけ手動実行してください。
 *   以降は毎朝9時に集計処理が自動で走るようになります。
 *   重複登録を防ぐため、既存の同関数向けトリガーは事前に削除します。
 */
function setupDailyTrigger() {
  var targetFunctionName = 'createSummaryDashboard';

  // 同名のトリガーが既にあれば削除して重複を防ぐ
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === targetFunctionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 毎日9時のトリガーを新規登録
  ScriptApp.newTrigger(targetFunctionName)
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();

  Logger.log('毎朝9時のトリガーを登録しました。');
}

// ===== テスト用サンプルデータ =====

/**
 * 「売上データ」シートを作成してサンプルデータを挿入する。
 *
 * 使い方：
 *   GAS エディタでこの関数を一度だけ手動実行してください。
 *   既存の「売上データ」シートがある場合は内容を上書きします。
 *   動作確認が済んだら実データに差し替えてください。
 */
function insertSampleData() {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = getOrCreateSheet(ss, SHEET_DATA);

  // シートを全クリアしてからヘッダーとサンプルデータを書き込む
  dataSheet.clearContents();

  var header = [['日付', '担当者名', '商品名', '金額']];
  var rows = [
    [new Date('2026/01/05'), '田中', '商品A',  50000],
    [new Date('2026/01/15'), '鈴木', '商品B',  30000],
    [new Date('2026/01/28'), '田中', '商品C',  70000],
    [new Date('2026/02/03'), '佐藤', '商品A',  45000],
    [new Date('2026/02/14'), '鈴木', '商品D',  90000],
    [new Date('2026/02/20'), '田中', '商品B',  60000],
    [new Date('2026/03/07'), '佐藤', '商品C',  38000],
    [new Date('2026/03/19'), '田中', '商品A',  82000],
    [new Date('2026/03/25'), '鈴木', '商品D', 110000],
    [new Date('2026/04/02'), '佐藤', '商品B',  55000],
    [new Date('2026/04/18'), '田中', '商品C',  67000],
    [new Date('2026/05/09'), '鈴木', '商品A',  48000],
    [new Date('2026/05/22'), '佐藤', '商品D',  95000],
  ];

  dataSheet.getRange(1, 1, 1, 4).setValues(header);
  dataSheet.getRange(2, 1, rows.length, 4).setValues(rows);

  // 日付列を見やすい書式に設定
  dataSheet.getRange(2, 1, rows.length, 1).setNumberFormat('yyyy/mm/dd');
  // 金額列を通貨書式に設定
  dataSheet.getRange(2, 4, rows.length, 1).setNumberFormat('¥#,##0');

  Logger.log('サンプルデータを挿入しました（' + rows.length + '件）。');
}

// ===== ユーティリティ =====

/**
 * 指定した名前のシートを取得する。存在しない場合は新規作成する。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} sheetName
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateSheet(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

/**
 * 数値を2桁のゼロ埋め文字列に変換する（月のソートキー用）。
 * @param {number} n
 * @returns {string}
 */
function zeroPad(n) {
  return n < 10 ? '0' + n : String(n);
}
