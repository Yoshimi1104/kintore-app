/**
 * Config.gs
 * マルチユーザー対応：ユーザーごとに独立したスプレッドシートを使用。
 * UserPropertiesに保存済みのSpreadsheetはそのまま利用し、未設定のユーザーには
 * 新しい専用Spreadsheetを作成する。
 */

function getSpreadsheet() {
  var props = PropertiesService.getUserProperties();
  var ssId = props.getProperty('SPREADSHEET_ID');

  if (ssId) {
    return SpreadsheetApp.openById(ssId);
  }

  // 新規ユーザー：既存の共有Spreadsheetは採用せず、専用Spreadsheetを作成してセットアップ
  var ss = SpreadsheetApp.create('筋トレアプリ データ');
  ssId = ss.getId();
  props.setProperty('SPREADSHEET_ID', ssId);
  setupAllSheetsOnSS(ss);
  return ss;
}
