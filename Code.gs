/**
 * Code.gs
 * Webアプリのメインエントリーポイント
 * セットアップ用関数もここに集約しています
 */

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('筋トレ記録アプリ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ===== セットアップ関数（GASエディタから手動実行） =====

function migrateExistingData() {
  var props = PropertiesService.getUserProperties();
  if (props.getProperty('SPREADSHEET_ID')) {
    Logger.log('すでに設定済みです: ' + props.getProperty('SPREADSHEET_ID'));
    return;
  }
  Logger.log('移行は実行しません。未設定のユーザーは初回利用時に専用スプレッドシートが作成されます。');
}

function addPhase7TrainingHeaders() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('トレーニング記録');
  if (!sheet) {
    Logger.log('トレーニング記録シートが見つかりません');
    return;
  }
  var lastCol = sheet.getLastColumn();
  var existing = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  for (var i = 0; i < existing.length; i++) {
    if (existing[i] === '睡眠時間(時間)') {
      Logger.log('Phase7ヘッダーはすでに追加済みです');
      return;
    }
  }
  var newHeaders = [
    '睡眠時間(時間)', '睡眠の質(1-5)', '疲労度(1-5)', '筋肉痛(1-5)',
    'トレ前体調', 'トレ前メモ', 'サプリメント', 'クールダウン'
  ];
  var startCol = lastCol + 1;
  var range = sheet.getRange(1, startCol, 1, newHeaders.length);
  range.setValues([newHeaders]);
  range.setBackground('#4A90D9').setFontColor('#FFFFFF')
       .setFontWeight('bold').setHorizontalAlignment('center');
  sheet.autoResizeColumns(startCol, newHeaders.length);
  Logger.log('完了: Phase7ヘッダーを追加しました');
}

// ===== 診断用関数（データ確認のため一時追加） =====

function debugGetTrainingRawData() {
  try {
    var ss = getSpreadsheet();
    var result = { trainingRows: [], setRows: [] };

    var trSheet = ss.getSheetByName('トレーニング記録');
    if (trSheet && trSheet.getLastRow() > 1) {
      var numRows = Math.min(trSheet.getLastRow() - 1, 20);
      var rows = trSheet.getRange(2, 1, numRows, 11).getValues();
      result.trainingRows = rows
        .filter(function(r) { return r[0] !== ''; })
        .map(function(r) {
          return {
            id:      String(r[0]),
            date:    String(r[1]),
            name:    String(r[4]),
            sets:    r[6],
            weight:  r[7]
          };
        });
    }

    var setsSheet = ss.getSheetByName('セット詳細');
    if (setsSheet && setsSheet.getLastRow() > 1) {
      var numSets = Math.min(setsSheet.getLastRow() - 1, 60);
      var sRows = setsSheet.getRange(2, 1, numSets, 5).getValues();
      result.setRows = sRows
        .filter(function(r) { return r[0] !== ''; })
        .map(function(r) {
          return {
            id:         String(r[0]),
            trainingId: String(r[1]),
            setNum:     r[2],
            weight:     r[3],
            reps:       r[4]
          };
        });
    }

    result.trainingTotal = result.trainingRows.length;
    result.setTotal = result.setRows.length;
    return { success: true, data: result };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function authorizeDriveFile() {
  // drive.file スコープの再認証トリガー用。データ変更なし・読み取りのみ。
  var result = Drive.Files.list({ pageSize: 1, fields: 'files(id)' });
  Logger.log('drive.file 認証OK: ' + JSON.stringify(result));
}

function setupBodyDataSheet() {
  var ss = getSpreadsheet();
  var existing = ss.getSheetByName('身体データ記録');
  if (existing) {
    Logger.log('身体データ記録シートはすでに存在します');
    return;
  }
  var sheet = ss.insertSheet('身体データ記録');
  var headers = [
    '記録ID', '日付', '体重(kg)', '胸囲(cm)',
    'ウエスト(cm)', 'ヒップ(cm)', '太もも(cm)',
    '自由メモ', '登録日時', '写真DriveID'
  ];
  var range = sheet.getRange(1, 1, 1, headers.length);
  range.setValues([headers]);
  range.setBackground('#8E44AD').setFontColor('#FFFFFF')
       .setFontWeight('bold').setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  Logger.log('完了: 身体データ記録シートを作成しました');
}
