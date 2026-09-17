/**
 * SetupSheets.gs
 * スプレッドシートの全シートを自動作成するセットアップスクリプト
 *
 * 使い方：
 * Google Apps Script エディタで「setupAllSheets」を選択して実行してください。
 */

// 手動実行用ラッパー（GASエディタから呼ぶ）
function setupAllSheets() {
  setupAllSheetsOnSS(getSpreadsheet());
}

// Config.gs の getSpreadsheet() から新規ユーザー初回時に呼ばれる
function setupAllSheetsOnSS(ss) {
  const sheetConfigs = [
    {
      name: 'トレーニング記録',
      headers: [
        '記録ID', '日付', '開始時刻', '終了時刻',
        '種目名', 'トレーニング部位', 'セット数', '最大重量(kg)',
        'RPE/きつさ(1-10)', 'フォーム・感覚メモ', '登録日時',
        '睡眠時間(時間)', '睡眠の質(1-5)', '疲労度(1-5)', '筋肉痛(1-5)',
        'トレ前体調', 'トレ前メモ', 'サプリメント', 'クールダウン'
      ],
      color: '#4A90D9'
    },
    {
      name: 'セット詳細',
      headers: [
        'セット詳細ID', '記録ID', 'セット番号',
        '重量(kg)', '回数', 'RPE', 'メモ'
      ],
      color: '#5BA85A'
    },
    {
      name: '有酸素運動記録',
      headers: [
        '記録ID', '日付', '種類', '時間(分)',
        '距離(km)', '速度(km/h)', '傾斜(%)',
        '消費カロリー(kcal)', 'メモ', '登録日時'
      ],
      color: '#E67E22'
    },
    {
      name: '体調記録',
      headers: [
        '記録ID', '日付', '体重(kg)', '睡眠時間(時間)',
        '睡眠の質(1-5)', '疲労度(1-10)', '筋肉痛の部位',
        '筋肉痛の程度(1-5)', '体調', '気分',
        '痛みがある場所', '痛みの程度(1-5)', '自由メモ', '登録日時'
      ],
      color: '#9B59B6'
    },
    {
      name: '食事・栄養記録',
      headers: [
        '記録ID', '日付', 'タンパク質(g)', '総カロリー(kcal)',
        '水分摂取(ml)', 'トレ前の食事メモ', 'トレ後の食事メモ',
        'その他メモ', '登録日時'
      ],
      color: '#E74C3C'
    },
    {
      name: 'ケア記録',
      headers: [
        '記録ID', '日付', 'ストレッチ実施', 'ストレッチ時間(分)',
        '筋膜リリース実施', 'クールダウン有酸素', 'クールダウン時間(分)',
        'その他ケア内容', 'メモ', '登録日時'
      ],
      color: '#1ABC9C'
    },
    {
      name: '種目マスタ',
      headers: ['種目名', '部位', '器具'],
      color: '#95A5A6'
    },
    {
      name: '身体データ記録',
      headers: [
        '記録ID', '日付', '体重(kg)', '胸囲(cm)',
        'ウエスト(cm)', 'ヒップ(cm)', '太もも(cm)',
        '自由メモ', '登録日時', '写真DriveID'
      ],
      color: '#8E44AD'
    },
    {
      name: 'AI分析履歴',
      headers: ['分析ID', '分析日時', '分析対象期間', '分析結果'],
      color: '#F39C12'
    }
  ];

  sheetConfigs.forEach(config => {
    createOrUpdateSheet(ss, config.name, config.headers, config.color);
  });

  addSampleExercises(ss);

  Logger.log('✅ セットアップ完了！作成したシート: ' + sheetConfigs.map(c => c.name).join(', '));
}

/**
 * シートを作成または更新する
 */
function createOrUpdateSheet(ss, sheetName, headers, headerColor) {
  let sheet = ss.getSheetByName(sheetName);


  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    // すでにヘッダー行があれば上書きしない（データ保護のため）
    const existingHeader = sheet.getRange(1, 1).getValue();
    if (existingHeader !== '') {
      Logger.log(sheetName + ' はすでにデータがあるためスキップしました。');
      return;
    }
  }

  // ヘッダー行を書き込む
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);

  // ヘッダーのスタイルを設定する
  headerRange
    .setBackground(headerColor)
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // 1行目を固定する（スクロールしてもヘッダーが見えるように）
  sheet.setFrozenRows(1);

  // 列幅を自動調整する
  sheet.autoResizeColumns(1, headers.length);

  Logger.log(sheetName + ' を作成しました。');
}

/**
 * 種目マスタにサンプル種目を追加する
 */
function addSampleExercises(ss) {
  const sheet = ss.getSheetByName('種目マスタ');
  if (!sheet) return;

  // データがすでにあれば追加しない
  if (sheet.getLastRow() > 1) return;

  const exercises = [
    // 胸
    ['ベンチプレス', '胸', 'バーベル'],
    ['ダンベルフライ', '胸', 'ダンベル'],
    ['ダンベルベンチプレス', '胸', 'ダンベル'],
    ['チェストプレス', '胸', 'マシン'],
    ['ペックデック', '胸', 'マシン'],
    // 背中
    ['デッドリフト', '背中', 'バーベル'],
    ['ラットプルダウン', '背中', 'マシン'],
    ['シーテッドロウ', '背中', 'マシン'],
    ['ダンベルロウ', '背中', 'ダンベル'],
    ['チンニング(懸垂)', '背中', '自重'],
    // 肩
    ['ショルダープレス', '肩', 'バーベル'],
    ['ダンベルショルダープレス', '肩', 'ダンベル'],
    ['サイドレイズ', '肩', 'ダンベル'],
    ['フロントレイズ', '肩', 'ダンベル'],
    // 腕
    ['バーベルカール', '腕(二頭筋)', 'バーベル'],
    ['ダンベルカール', '腕(二頭筋)', 'ダンベル'],
    ['トライセプスプレスダウン', '腕(三頭筋)', 'マシン'],
    ['スカルクラッシャー', '腕(三頭筋)', 'バーベル'],
    // 脚
    ['スクワット', '脚', 'バーベル'],
    ['レッグプレス', '脚', 'マシン'],
    ['レッグエクステンション', '脚(大腿四頭筋)', 'マシン'],
    ['レッグカール', '脚(ハムストリングス)', 'マシン'],
    ['カーフレイズ', '脚(ふくらはぎ)', 'マシン'],
    ['ランジ', '脚', 'ダンベル'],
    // 体幹
    ['プランク', '体幹', '自重'],
    ['クランチ', '腹筋', '自重'],
    ['レッグレイズ', '腹筋', '自重'],
  ];

  const dataRange = sheet.getRange(2, 1, exercises.length, 3);
  dataRange.setValues(exercises);

  // 交互に色を付けて見やすくする
  exercises.forEach((_, i) => {
    const row = sheet.getRange(i + 2, 1, 1, 3);
    if (i % 2 === 0) {
      row.setBackground('#F8F9FA');
    }
  });
}
