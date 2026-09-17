/**
 * DataService.gs
 * スプレッドシートへのデータ読み書き処理
 */

// ===== 筋トレ記録 =====

function saveTrainingData(data) {
  try {
    const ss = getSpreadsheet();
    const trainingSheet = ss.getSheetByName('トレーニング記録');
    const setsSheet     = ss.getSheetByName('セット詳細');

    if (!trainingSheet) throw new Error('トレーニング記録シートが見つかりません');
    if (!setsSheet)     throw new Error('セット詳細シートが見つかりません');

    const now     = new Date();
    const dateStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMMdd');
    const tsStr   = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');

    // Phase 7: セッション単位の付加情報（最初の種目行のみ保存）
    const pre  = data.preCondition || {};
    const supp = data.supplements  || '';
    const cd   = data.cooldown     || '';

    data.exercises.forEach((ex, idx) => {
      const trRowNum   = String(trainingSheet.getLastRow()).padStart(3, '0');
      const trainingId = 'TR-' + dateStr + '-' + trRowNum;

      let maxWeight = 0;
      let maxRpe    = '';
      ex.sets.forEach(s => {
        const w = parseFloat(s.weight) || 0;
        if (w > maxWeight) maxWeight = w;
        const r = parseInt(s.rpe) || 0;
        if (r > 0 && (maxRpe === '' || r > parseInt(maxRpe))) maxRpe = String(r);
      });

      // 基本11列（既存）
      const row = [
        trainingId,
        data.date        || '',
        data.startTime   || '',
        '',
        ex.name          || '',
        ex.bodyPart      || '',
        ex.sets.length,
        maxWeight        || '',
        maxRpe,
        ex.memo          || '',
        tsStr
      ];

      // 最初の種目行にのみセッション情報を追加（列12〜19）
      if (idx === 0) {
        row.push(
          pre.sleepHours   || '',  // 列12: 睡眠時間
          pre.sleepQuality || '',  // 列13: 睡眠の質(1-5)
          pre.fatigue      || '',  // 列14: 疲労度(1-5)
          pre.soreness     || '',  // 列15: 筋肉痛(1-5)
          pre.condition    || '',  // 列16: トレ前体調
          pre.memo         || '',  // 列17: トレ前メモ
          supp,                    // 列18: サプリメント
          cd                       // 列19: クールダウン
        );
      }

      trainingSheet.appendRow(row);

      ex.sets.forEach((s, j) => {
        const sdRowNum = String(setsSheet.getLastRow()).padStart(3, '0');
        const setId    = 'SD-' + dateStr + '-' + sdRowNum;

        setsSheet.appendRow([
          setId,
          trainingId,
          j + 1,
          s.weight || '',
          s.reps   || '',
          s.rpe    || '',
          ''
        ]);
      });
    });

    return { success: true };
  } catch (e) {
    return { success: false, error: e.message, stack: String(e.stack || '') };
  }
}

// ===== トレーニング履歴 =====

function getTrainingHistory() {
  try {
    var ss = getSpreadsheet();
    var trainingSheet = ss.getSheetByName('トレーニング記録');

    if (!trainingSheet || trainingSheet.getLastRow() <= 1) {
      return { success: true, sessions: [] };
    }

    // セット詳細を先に Map 化（trainingId → sets[]）
    var setsMap = {};
    var setsSheet = ss.getSheetByName('セット詳細');
    if (setsSheet && setsSheet.getLastRow() > 1) {
      var setsRows = setsSheet.getRange(2, 1, setsSheet.getLastRow() - 1, 7).getValues();
      setsRows.forEach(function(r) {
        var tid = String(r[1] || '');
        if (!tid) return;
        if (!setsMap[tid]) setsMap[tid] = [];
        setsMap[tid].push({
          setNum: Number(r[2]) || 0,
          weight: r[3] !== '' ? String(r[3]) : '',
          reps:   r[4] !== '' ? String(r[4]) : '',
          rpe:    r[5] !== '' ? String(r[5]) : ''
        });
      });
    }

    // トレーニング記録を日付でグループ化（Phase7列 col12〜19 も読む）
    var lastCol = trainingSheet.getLastColumn();
    var readCols = Math.max(11, Math.min(lastCol, 19));
    var trRows = trainingSheet.getRange(2, 1, trainingSheet.getLastRow() - 1, readCols).getValues();
    var dateMap = {};
    var dateOrder = [];

    trRows.forEach(function(r) {
      if (!r[1]) return;
      var date = bodyDateStr(r[1]);
      if (!date) return;

      if (!dateMap[date]) {
        dateMap[date] = { date: date, exercises: [], preCondition: null };
        dateOrder.push(date);
      }

      // トレ前コンディション（col12〜19 = r[11]〜r[18]）は先頭種目行のみ保存されている
      if (readCols >= 12 && !dateMap[date].preCondition) {
        var hasAnyPre = false;
        for (var ci = 11; ci < readCols; ci++) {
          if (r[ci] !== '' && r[ci] !== null) { hasAnyPre = true; break; }
        }
        if (hasAnyPre) {
          dateMap[date].preCondition = {
            sleepHours:   String(r[11] || ''),
            sleepQuality: readCols > 12 ? String(r[12] || '') : '',
            fatigue:      readCols > 13 ? String(r[13] || '') : '',
            soreness:     readCols > 14 ? String(r[14] || '') : '',
            condition:    readCols > 15 ? String(r[15] || '') : '',
            memo:         readCols > 16 ? String(r[16] || '') : '',
            supplements:  readCols > 17 ? String(r[17] || '') : '',
            cooldown:     readCols > 18 ? String(r[18] || '') : ''
          };
        }
      }

      var tid = String(r[0] || '');
      var tidSets = setsMap[tid] || [];
      dateMap[date].exercises.push({
        id:        tid,
        name:      String(r[4] || ''),
        bodyPart:  String(r[5] || ''),
        totalSets: tidSets.length > 0 ? tidSets.length : (r[6] !== '' ? Number(r[6]) : 0),
        maxWeight: r[7] !== '' ? String(r[7]) : '',
        maxRpe:    r[8] !== '' ? String(r[8]) : '',
        memo:      String(r[9] || ''),
        sets:      tidSets
      });
    });

    // 重複排除して日付降順ソート
    var seen = {};
    var sessions = dateOrder
      .filter(function(d) { if (seen[d]) return false; seen[d] = true; return true; })
      .sort(function(a, b) { return b.localeCompare(a); })
      .map(function(d) { return dateMap[d]; });

    return { success: true, sessions: sessions };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ===== ホーム画面 =====

function getHomeData(dateStr) {
  try {
    const ss = getSpreadsheet();
    const now = new Date();
    const todayStr = dateStr || Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd');

    function toDateStr(val) {
      if (!val) return '';
      if (val instanceof Date) return Utilities.formatDate(val, 'Asia/Tokyo', 'yyyy-MM-dd');
      return String(val).trim();
    }

    // 筋トレサマリー
    let training = null;
    const trainingSheet = ss.getSheetByName('トレーニング記録');
    if (trainingSheet && trainingSheet.getLastRow() > 1) {
      const rows = trainingSheet.getRange(2, 1, trainingSheet.getLastRow() - 1, 11).getValues();
      const todayRows = rows.filter(r => toDateStr(r[1]) === todayStr);
      if (todayRows.length > 0) {
        const allNames = todayRows.map(r => String(r[4])).filter(Boolean);
        const uniqueNames = allNames.filter((v, i, a) => a.indexOf(v) === i);

        // セット詳細シートから正確なセット数を計算する
        const todayIds = todayRows.map(r => String(r[0])).filter(Boolean);
        let totalSets = 0;
        const setsSheetForHome = ss.getSheetByName('セット詳細');
        if (setsSheetForHome && setsSheetForHome.getLastRow() > 1) {
          const setsData = setsSheetForHome.getRange(2, 1, setsSheetForHome.getLastRow() - 1, 2).getValues();
          totalSets = setsData.filter(r => todayIds.indexOf(String(r[1])) >= 0).length;
        }
        // セット詳細シートにデータがなければトレーニング記録シートの値をフォールバックで使う
        if (totalSets === 0) {
          totalSets = todayRows.reduce((sum, r) => sum + (parseInt(r[6]) || 0), 0);
        }

        training = { exerciseCount: uniqueNames.length, exercises: uniqueNames, totalSets: totalSets };
      }
    }

    // 体調サマリー
    let condition = null;
    const conditionSheet = ss.getSheetByName('体調記録');
    if (conditionSheet && conditionSheet.getLastRow() > 1) {
      const rows = conditionSheet.getRange(2, 1, conditionSheet.getLastRow() - 1, 14).getValues();
      const todayRows = rows.filter(r => toDateStr(r[1]) === todayStr);
      if (todayRows.length > 0) {
        const row = todayRows[todayRows.length - 1];
        condition = {
          weight:    row[2],
          fatigue:   row[5],
          condition: row[8],
          mood:      row[9]
        };
      }
    }

    // 有酸素サマリー（今日の全記録）
    let cardio = null;
    const cardioSheet = ss.getSheetByName('有酸素運動記録');
    if (cardioSheet && cardioSheet.getLastRow() > 1) {
      const rows = cardioSheet.getRange(2, 1, cardioSheet.getLastRow() - 1, 9).getValues();
      const todayRows = rows.filter(r => toDateStr(r[1]) === todayStr);
      if (todayRows.length > 0) {
        cardio = todayRows.map(r => ({
          type:     String(r[2] || ''),
          duration: r[3] !== '' && r[3] !== null ? String(r[3]) : '',
          distance: r[4] !== '' && r[4] !== null ? String(r[4]) : '',
          calories: r[7] !== '' && r[7] !== null ? String(r[7]) : ''
        }));
      }
    }

    // 食事サマリー（今日の最新記録）
    let meal = null;
    const mealSheet = ss.getSheetByName('食事・栄養記録');
    if (mealSheet && mealSheet.getLastRow() > 1) {
      const rows = mealSheet.getRange(2, 1, mealSheet.getLastRow() - 1, 8).getValues();
      const todayRows = rows.filter(r => toDateStr(r[1]) === todayStr);
      if (todayRows.length > 0) {
        const row = todayRows[todayRows.length - 1];
        meal = {
          protein:  row[2] !== '' && row[2] !== null ? String(row[2]) : '',
          calories: row[3] !== '' && row[3] !== null ? String(row[3]) : '',
          water:    row[4] !== '' && row[4] !== null ? String(row[4]) : '',
          preMeal:  String(row[5] || ''),
          postMeal: String(row[6] || ''),
          memo:     String(row[7] || '')
        };
      }
    }

    // ケアサマリー（今日の最新記録）
    let care = null;
    const careSheet = ss.getSheetByName('ケア記録');
    if (careSheet && careSheet.getLastRow() > 1) {
      const rows = careSheet.getRange(2, 1, careSheet.getLastRow() - 1, 9).getValues();
      const todayRows = rows.filter(r => toDateStr(r[1]) === todayStr);
      if (todayRows.length > 0) {
        const row = todayRows[todayRows.length - 1];
        care = {
          stretch:    String(row[2] || ''),
          foamRoller: String(row[4] || ''),
          cardio:     String(row[5] || ''),
          otherCare:  String(row[7] || ''),
          memo:       String(row[8] || '')
        };
      }
    }

    return { success: true, today: todayStr, training: training, condition: condition, cardio: cardio, meal: meal, care: care };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ===== 身体データ記録 =====

function uploadBodyPhoto(base64Data, mimeType, filename) {
  try {
    // DriveApp.createFolder/getFoldersByName は drive スコープが必要なため使用しない。
    // Drive Advanced Service (v3) は drive.file スコープで動作する。
    const props = PropertiesService.getUserProperties();
    let folderId = props.getProperty('BODY_PHOTO_FOLDER_ID');

    if (folderId) {
      try {
        var meta = Drive.Files.get(folderId, { fields: 'id,trashed' });
        if (meta.trashed) folderId = null;
      } catch (e) {
        folderId = null; // フォルダが消えていたら再作成
      }
    }

    if (!folderId) {
      var newFolder = Drive.Files.create({
        name: '筋トレアプリ_身体写真',
        mimeType: 'application/vnd.google-apps.folder'
      });
      folderId = newFolder.id;
      props.setProperty('BODY_PHOTO_FOLDER_ID', folderId);
    }

    const decoded = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(decoded, mimeType, filename);
    var uploadedFile = Drive.Files.create(
      { name: filename, parents: [folderId] },
      blob,
      { fields: 'id' }
    );

    return { success: true, fileId: uploadedFile.id };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function saveBodyData(data) {
  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName('身体データ記録');

    // シートが存在しない場合は初回保存時に自動作成
    if (!sheet) {
      sheet = ss.insertSheet('身体データ記録');
      const headers = [
        '記録ID', '日付', '体重(kg)', '胸囲(cm)',
        'ウエスト(cm)', 'ヒップ(cm)', '太もも(cm)',
        '自由メモ', '登録日時', '写真DriveID'
      ];
      const hr = sheet.getRange(1, 1, 1, headers.length);
      hr.setValues([headers]);
      hr.setBackground('#8E44AD')
        .setFontColor('#FFFFFF')
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
      sheet.setFrozenRows(1);
      sheet.autoResizeColumns(1, headers.length);
    }

    const now     = new Date();
    const dateStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMMdd');
    const rowNum  = String(sheet.getLastRow()).padStart(3, '0');
    const id      = 'BD-' + dateStr + '-' + rowNum;

    // 列1-9は既存データと同じ並び。列10に写真DriveIDを追加（過去データとの互換性を維持）
    sheet.appendRow([
      id,
      data.date        || '',
      data.weight      || '',
      data.chest       || '',
      data.waist       || '',
      data.hip         || '',
      data.thigh       || '',
      data.memo        || '',
      Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'),
      data.photoFileId || ''
    ]);

    return { success: true, id: id };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ===== 写真レコード挿入（アップロード済みfileIdをシートに記録するだけ） =====
function insertBodyPhotoRecord(data) {
  try {
    var ss = getSpreadsheet();
    var sheets = _ensureBodyPhotoSheets(ss);
    var photoId = Utilities.getUuid();
    sheets.photoSheet.appendRow([
      photoId,
      String(data.recordId || ''),
      String(data.date     || ''),
      String(data.partId   || ''),
      String(data.fileId   || ''),
      true,
      new Date().toISOString()
    ]);
    return { success: true, photoId: photoId };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function getBodyHistory() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('身体データ記録');
    if (!sheet || sheet.getLastRow() <= 1) {
      return { success: true, records: [] };
    }

    // 既存データは9列、新規データは10列（写真DriveID）
    const lastCol = sheet.getLastColumn();
    const readCols = Math.min(Math.max(lastCol, 9), 10);
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, readCols).getValues();
    const records = rows
      .filter(r => r[1] !== '' && r[1] !== null)
      .map(r => ({
        id:          String(r[0] || ''),
        date:        bodyDateStr(r[1]),
        weight:      r[2] !== '' ? String(r[2]) : '',
        chest:       r[3] !== '' ? String(r[3]) : '',
        waist:       r[4] !== '' ? String(r[4]) : '',
        hip:         r[5] !== '' ? String(r[5]) : '',
        thigh:       r[6] !== '' ? String(r[6]) : '',
        memo:        String(r[7] || ''),
        photoFileId: r.length > 9 ? String(r[9] || '') : ''
      }))
      .sort((a, b) => b.date.localeCompare(a.date));

    return { success: true, records: records };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ===== 身体写真 変更・削除 =====

function updateBodyPhoto(recordId, newBase64, newMime, newFilename) {
  var newFileId = null;
  try {
    // 1. 新しい写真を先にアップロード
    var uploadResult = uploadBodyPhoto(newBase64, newMime, newFilename);
    if (!uploadResult.success) return uploadResult;
    newFileId = uploadResult.fileId;

    // 2. 対象行を検索
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('身体データ記録');
    if (!sheet) throw new Error('身体データ記録シートが見つかりません');
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) throw new Error('記録が見つかりません');

    var rows = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
    var targetRow = -1;
    var oldFileId = '';
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) === String(recordId)) {
        targetRow = i + 2;
        oldFileId = String(rows[i][9] || '');
        break;
      }
    }

    if (targetRow === -1) {
      // 行が見つからない場合は新しいファイルをロールバック削除
      try { Drive.Files.remove(newFileId); } catch(e2) {}
      throw new Error('記録ID ' + recordId + ' が見つかりません');
    }

    // 3. スプレッドシートを更新（先に成功させる）
    sheet.getRange(targetRow, 10).setValue(newFileId);

    // 4. 古いファイルを安全に削除（失敗しても続行）
    if (oldFileId) {
      try { Drive.Files.remove(oldFileId); } catch(e3) {}
    }

    return { success: true, newFileId: newFileId };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function deleteBodyPhoto(recordId) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('身体データ記録');
    if (!sheet) throw new Error('身体データ記録シートが見つかりません');
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) throw new Error('記録が見つかりません');

    var rows = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
    var targetRow = -1;
    var fileId = '';
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) === String(recordId)) {
        targetRow = i + 2;
        fileId = String(rows[i][9] || '');
        break;
      }
    }

    if (targetRow === -1) throw new Error('記録が見つかりません');

    // スプレッドシートから先にクリア
    sheet.getRange(targetRow, 10).setValue('');

    // DriveファイルをIDで削除（失敗しても続行）
    if (fileId) {
      try { Drive.Files.remove(fileId); } catch(e2) {}
    }

    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function bodyDateStr(val) {
  if (!val) return '';
  if (val instanceof Date) return Utilities.formatDate(val, 'Asia/Tokyo', 'yyyy-MM-dd');
  return String(val).trim();
}

// ===== ケア記録 =====

function saveCareData(data) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('ケア記録');
    if (!sheet) throw new Error('ケア記録シートが見つかりません');

    const now = new Date();
    const dateStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMMdd');
    const rowNum  = String(sheet.getLastRow()).padStart(3, '0');
    const id      = 'KA-' + dateStr + '-' + rowNum;

    sheet.appendRow([
      id,
      data.date         || '',
      data.stretch      ? 'はい' : '',  // ストレッチ実施
      '',                               // ストレッチ時間(分)
      data.foamRoller   ? 'はい' : '',  // 筋膜リリース実施
      data.cardio       || '',          // クールダウン有酸素（種類＋時間を文字列で格納）
      '',                               // クールダウン時間(分)
      data.otherCare    || '',          // その他ケア内容
      data.memo         || '',          // メモ
      Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss')
    ]);

    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ===== 体調記録 履歴 =====

function getConditionHistory() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('体調記録');
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, records: [] };
    var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getValues();
    var records = rows
      .filter(function(r) { return r[1] !== '' && r[1] !== null; })
      .map(function(r) {
        return {
          id:            String(r[0] || ''),
          date:          bodyDateStr(r[1]),
          weight:        r[2] !== '' ? String(r[2]) : '',
          sleepHours:    r[3] !== '' ? String(r[3]) : '',
          sleepQuality:  r[4] !== '' ? String(r[4]) : '',
          fatigue:       r[5] !== '' ? String(r[5]) : '',
          soreness:      String(r[6] || ''),
          sorenessLevel: r[7] !== '' ? String(r[7]) : '',
          condition:     String(r[8] || ''),
          mood:          String(r[9] || ''),
          painLocation:  String(r[10] || ''),
          painLevel:     r[11] !== '' ? String(r[11]) : '',
          memo:          String(r[12] || '')
        };
      })
      .sort(function(a, b) { return b.date.localeCompare(a.date); });
    return { success: true, records: records };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function getCareHistory() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('ケア記録');
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, records: [] };
    var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
    var records = rows
      .filter(function(r) { return r[1] !== '' && r[1] !== null; })
      .map(function(r) {
        return {
          date:       bodyDateStr(r[1]),
          stretch:    String(r[2] || ''),
          foamRoller: String(r[4] || ''),
          cardio:     String(r[5] || ''),
          otherCare:  String(r[7] || ''),
          memo:       String(r[8] || '')
        };
      })
      .sort(function(a, b) { return b.date.localeCompare(a.date); });
    return { success: true, records: records };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function getCardioHistory() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('有酸素運動記録');
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, records: [] };
    var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
    var records = rows
      .filter(function(r) { return r[1] !== '' && r[1] !== null; })
      .map(function(r) {
        return {
          id:       String(r[0] || ''),
          date:     bodyDateStr(r[1]),
          type:     String(r[2] || ''),
          duration: r[3] !== '' ? String(r[3]) : '',
          distance: r[4] !== '' ? String(r[4]) : '',
          speed:    r[5] !== '' ? String(r[5]) : '',
          slope:    r[6] !== '' ? String(r[6]) : '',
          calories: r[7] !== '' ? String(r[7]) : '',
          memo:     String(r[8] || '')
        };
      })
      .sort(function(a, b) { return b.date.localeCompare(a.date); });
    return { success: true, records: records };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function getMealHistory() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('食事・栄養記録');
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, records: [] };
    // 旧データは9列、STEP2以降は22列。getLastColumn()でシート最大幅を取得しキャップ
    var readCols = Math.min(Math.max(sheet.getLastColumn(), 9), 22);
    var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, readCols).getValues();
    var records = rows
      .filter(function(r) { return r[1] !== '' && r[1] !== null; })
      .map(function(r) {
        return {
          // 既存フィールド（後方互換）
          id:               String(r[0] || ''),
          date:             bodyDateStr(r[1]),
          protein:          r[2] !== '' ? String(r[2]) : '',
          calories:         r[3] !== '' ? String(r[3]) : '',
          water:            r[4] !== '' ? String(r[4]) : '',
          preMeal:          String(r[5]  || ''),
          postMeal:         String(r[6]  || ''),
          memo:             String(r[7]  || ''),
          // STEP2追加フィールド（旧データは空文字になる）
          inputMode:        String(r[9]  || ''),
          mealAmount:       String(r[10] || ''),
          proteinLevel:     String(r[11] || ''),
          carbLevel:        String(r[12] || ''),
          fatLevel:         String(r[13] || ''),
          preWorkoutMeal:   String(r[14] || ''),
          preWorkoutTiming: String(r[15] || ''),
          preWorkoutCarbs:  String(r[16] || ''),
          fat:              (r[17] !== undefined && r[17] !== '') ? String(r[17]) : '',
          carbs:            (r[18] !== undefined && r[18] !== '') ? String(r[18]) : '',
          fiber:            (r[19] !== undefined && r[19] !== '') ? String(r[19]) : '',
          salt:             (r[20] !== undefined && r[20] !== '') ? String(r[20]) : '',
          caffeine:         (r[21] !== undefined && r[21] !== '') ? String(r[21]) : ''
        };
      })
      .sort(function(a, b) { return b.date.localeCompare(a.date); });
    return { success: true, records: records };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ===== 体調記録 =====

function saveConditionData(data) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('体調記録');
    if (!sheet) throw new Error('体調記録シートが見つかりません');

    const now = new Date();
    const dateStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMMdd');
    const rowNum = String(sheet.getLastRow()).padStart(3, '0');
    const id = 'CD-' + dateStr + '-' + rowNum;

    sheet.appendRow([
      id,
      data.date        || '',
      data.weight      || '',
      data.sleepHours  || '',
      data.sleepQuality || '',
      data.fatigue     || '',
      data.soreness    || '',
      data.sorenessLevel || '',
      data.condition   || '',
      data.mood        || '',
      data.painLocation || '',
      data.painLevel   || '',
      data.memo        || '',
      Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss')
    ]);

    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ===== 有酸素運動記録 =====

function saveCardioData(data) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('有酸素運動記録');
    if (!sheet) throw new Error('有酸素運動記録シートが見つかりません');

    const now = new Date();
    const dateStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMMdd');
    const rowNum  = String(sheet.getLastRow()).padStart(3, '0');
    const id      = 'CA-' + dateStr + '-' + rowNum;

    sheet.appendRow([
      id,
      data.date     || '',
      data.type     || '',
      data.duration || '',
      data.distance || '',
      data.speed    || '',
      data.slope    || '',
      data.calories || '',
      data.memo     || '',
      Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss')
    ]);

    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ===== 食事・栄養記録 =====

function saveMealData(data) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('食事・栄養記録');
    if (!sheet) throw new Error('食事・栄養記録シートが見つかりません');

    const now = new Date();
    const dateStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMMdd');
    const rowNum  = String(sheet.getLastRow()).padStart(3, '0');
    const id      = 'ML-' + dateStr + '-' + rowNum;

    // 数値バリデーション：空欄は空欄のまま保存（0扱いしない）、負数はNG
    function safeNum(val) {
      if (val === '' || val === null || val === undefined) return '';
      var n = parseFloat(val);
      return (!isNaN(n) && n >= 0) ? n : '';
    }

    sheet.appendRow([
      // 既存9列（後方互換維持）
      id,
      data.date     || '',
      safeNum(data.protein),   // タンパク質(g) — 詳細モード
      safeNum(data.calories),  // 総カロリー(kcal) — 詳細モード
      safeNum(data.water),     // 水分(ml) — 両モード共用
      data.preMeal  || '',     // トレ前メモ(テキスト) — 詳細/レガシー
      data.postMeal || '',     // トレ後メモ(テキスト) — 詳細/レガシー
      data.memo     || '',     // 一般メモ
      Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'),
      // STEP2追加列（列10〜22）
      data.inputMode          || '',  // 入力方式（かんたん/詳細）
      data.mealAmount         || '',  // 食事量(簡易)
      data.proteinLevel       || '',  // タンパク質(簡易)
      data.carbLevel          || '',  // 炭水化物(簡易)
      data.fatLevel           || '',  // 脂質(簡易)
      data.preWorkoutMeal     || '',  // トレ前食事(簡易)
      data.preWorkoutTiming   || '',  // トレ前タイミング
      data.preWorkoutCarbs    || '',  // トレ前炭水化物(簡易)
      safeNum(data.fat),              // 脂質(g) — 詳細
      safeNum(data.carbs),            // 炭水化物(g) — 詳細
      safeNum(data.fiber),            // 食物繊維(g) — 任意
      safeNum(data.salt),             // 塩分(g) — 任意
      safeNum(data.caffeine)          // カフェイン(mg) — 任意
    ]);

    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ===== 種目マスタ =====

function getExerciseMaster() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('種目マスタ');
    if (!sheet || sheet.getLastRow() <= 1) {
      return { success: true, exercises: [] };
    }
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
    const exercises = rows
      .filter(r => r[0] !== '' && r[0] !== null)
      .map(r => ({
        name:      String(r[0] || ''),
        part:      String(r[1] || ''),
        equipment: String(r[2] || '')
      }));
    return { success: true, exercises: exercises };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ===== 種目マスタ 追加・更新 =====
function addOrUpdateExerciseMaster(name, part) {
  try {
    const ss    = getSpreadsheet();
    const sheet = ss.getSheetByName('種目マスタ');
    if (!sheet) throw new Error('種目マスタシートが見つかりません');
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const vals = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
      for (let i = 0; i < vals.length; i++) {
        if (String(vals[i][0]) === String(name)) {
          sheet.getRange(i + 2, 2).setValue(part);
          return { success: true, action: 'updated' };
        }
      }
    }
    sheet.appendRow([name, part, '']);
    return { success: true, action: 'added' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ===== STEP3分析: 体調・食事・種目マスタを一括取得 =====
function getStep3Bundle() {
  try {
    var c = getConditionHistory();
    var m = getMealHistory();
    var e = getExerciseMaster();
    return {
      success:    true,
      conditions: (c && c.success) ? c.records   : [],
      meals:      (m && m.success) ? m.records   : [],
      exercises:  (e && e.success) ? e.exercises : []
    };
  } catch(e) {
    return { success: false, error: e.message, conditions: [], meals: [], exercises: [] };
  }
}

// ===== 履歴編集：筋トレ =====
function updateTrainingSession(data) {
  try {
    const ss        = getSpreadsheet();
    const trSheet   = ss.getSheetByName('トレーニング記録');
    const setsSheet = ss.getSheetByName('セット詳細');
    if (!trSheet)   throw new Error('トレーニング記録シートが見つかりません');
    if (!setsSheet) throw new Error('セット詳細シートが見つかりません');

    const now    = new Date();
    const nowStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMMdd');

    // 各種目行を ID で特定して更新
    const trLastRow = trSheet.getLastRow();
    if (trLastRow > 1) {
      const readCols = Math.min(trSheet.getLastColumn(), 19);
      const trRows   = trSheet.getRange(2, 1, trLastRow - 1, readCols).getValues();
      data.exercises.forEach(function(ex) {
        if (!ex.id) return;
        const sets      = ex.sets || [];
        let   maxWeight = 0, maxRpe = '';
        sets.forEach(function(s) {
          const w = parseFloat(s.weight) || 0;
          if (w > maxWeight) maxWeight = w;
          const r = parseInt(s.rpe) || 0;
          if (r > 0 && (maxRpe === '' || r > parseInt(maxRpe))) maxRpe = String(r);
        });
        for (let i = 0; i < trRows.length; i++) {
          if (String(trRows[i][0]) === String(ex.id)) {
            const rowNum = i + 2;
            trSheet.getRange(rowNum, 2).setValue(data.date || '');
            trSheet.getRange(rowNum, 5).setValue(ex.name     || '');
            trSheet.getRange(rowNum, 6).setValue(ex.bodyPart || '');
            trSheet.getRange(rowNum, 7).setValue(sets.length);
            trSheet.getRange(rowNum, 8).setValue(maxWeight   || '');
            trSheet.getRange(rowNum, 9).setValue(maxRpe);
            trSheet.getRange(rowNum, 10).setValue(ex.memo   || '');
            break;
          }
        }
      });
    }

    // セット詳細の更新：対象 TR-ID のセットを削除 → 再挿入
    const exerciseIds = data.exercises.map(function(ex) { return String(ex.id || ''); })
                                      .filter(function(id) { return id; });
    if (exerciseIds.length > 0) {
      const sLastRow = setsSheet.getLastRow();
      if (sLastRow > 1) {
        const sRows     = setsSheet.getRange(2, 1, sLastRow - 1, 2).getValues();
        const toDelete  = [];
        sRows.forEach(function(r, i) {
          if (exerciseIds.indexOf(String(r[1])) >= 0) toDelete.push(i + 2);
        });
        for (let i = toDelete.length - 1; i >= 0; i--) {
          setsSheet.deleteRow(toDelete[i]);
        }
      }
      data.exercises.forEach(function(ex) {
        if (!ex.id) return;
        (ex.sets || []).forEach(function(s, j) {
          const sdRow = String(setsSheet.getLastRow()).padStart(3, '0');
          setsSheet.appendRow([
            'SD-' + nowStr + '-' + sdRow,
            ex.id,
            j + 1,
            s.weight || '',
            s.reps   || '',
            s.rpe    || '',
            ''
          ]);
        });
      });
    }

    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ===== 履歴編集：有酸素 =====
function updateCardioRecord(data) {
  try {
    const ss    = getSpreadsheet();
    const sheet = ss.getSheetByName('有酸素運動記録');
    if (!sheet) throw new Error('有酸素運動記録シートが見つかりません');
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) throw new Error('記録が見つかりません');
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    let targetRow = -1;
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(data.id)) { targetRow = i + 2; break; }
    }
    if (targetRow === -1) throw new Error('記録ID ' + data.id + ' が見つかりません');
    function safeVal(v) { return (v !== undefined && v !== null) ? v : ''; }
    sheet.getRange(targetRow, 2).setValue(safeVal(data.date));
    sheet.getRange(targetRow, 3).setValue(safeVal(data.type));
    sheet.getRange(targetRow, 4).setValue(safeVal(data.duration));
    sheet.getRange(targetRow, 5).setValue(safeVal(data.distance));
    sheet.getRange(targetRow, 6).setValue(safeVal(data.speed));
    sheet.getRange(targetRow, 7).setValue(safeVal(data.slope));
    sheet.getRange(targetRow, 8).setValue(safeVal(data.calories));
    sheet.getRange(targetRow, 9).setValue(safeVal(data.memo));
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ===== 履歴編集：体調 =====
function updateConditionRecord(data) {
  try {
    const ss    = getSpreadsheet();
    const sheet = ss.getSheetByName('体調記録');
    if (!sheet) throw new Error('体調記録シートが見つかりません');
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) throw new Error('記録が見つかりません');
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    let targetRow = -1;
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(data.id)) { targetRow = i + 2; break; }
    }
    if (targetRow === -1) throw new Error('記録ID ' + data.id + ' が見つかりません');
    function safeVal(v) { return (v !== undefined && v !== null) ? v : ''; }
    sheet.getRange(targetRow, 2).setValue(safeVal(data.date));
    sheet.getRange(targetRow, 3).setValue(safeVal(data.weight));
    sheet.getRange(targetRow, 4).setValue(safeVal(data.sleepHours));
    sheet.getRange(targetRow, 5).setValue(safeVal(data.sleepQuality));
    sheet.getRange(targetRow, 6).setValue(safeVal(data.fatigue));
    sheet.getRange(targetRow, 7).setValue(safeVal(data.soreness));
    sheet.getRange(targetRow, 8).setValue(safeVal(data.sorenessLevel));
    sheet.getRange(targetRow, 9).setValue(safeVal(data.condition));
    sheet.getRange(targetRow, 10).setValue(safeVal(data.mood));
    sheet.getRange(targetRow, 11).setValue(safeVal(data.painLocation));
    sheet.getRange(targetRow, 12).setValue(safeVal(data.painLevel));
    sheet.getRange(targetRow, 13).setValue(safeVal(data.memo));
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ===== 履歴編集：食事 =====
function updateMealRecord(data) {
  try {
    const ss    = getSpreadsheet();
    const sheet = ss.getSheetByName('食事・栄養記録');
    if (!sheet) throw new Error('食事・栄養記録シートが見つかりません');
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) throw new Error('記録が見つかりません');
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    let targetRow = -1;
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(data.id)) { targetRow = i + 2; break; }
    }
    if (targetRow === -1) throw new Error('記録ID ' + data.id + ' が見つかりません');
    function safeNum(val) {
      if (val === '' || val === null || val === undefined) return '';
      const n = parseFloat(val);
      return (!isNaN(n) && n >= 0) ? n : '';
    }
    function safeStr(v) { return (v !== undefined && v !== null) ? String(v) : ''; }
    // 既存9列
    sheet.getRange(targetRow, 2).setValue(safeStr(data.date));
    sheet.getRange(targetRow, 3).setValue(safeNum(data.protein));
    sheet.getRange(targetRow, 4).setValue(safeNum(data.calories));
    sheet.getRange(targetRow, 5).setValue(safeNum(data.water));
    sheet.getRange(targetRow, 6).setValue(safeStr(data.preMeal));
    sheet.getRange(targetRow, 7).setValue(safeStr(data.postMeal));
    sheet.getRange(targetRow, 8).setValue(safeStr(data.memo));
    // STEP2列 (10〜22)
    const readCols = sheet.getLastColumn();
    if (readCols >= 10) {
      sheet.getRange(targetRow, 10).setValue(safeStr(data.inputMode));
      sheet.getRange(targetRow, 11).setValue(safeStr(data.mealAmount));
      sheet.getRange(targetRow, 12).setValue(safeStr(data.proteinLevel));
      sheet.getRange(targetRow, 13).setValue(safeStr(data.carbLevel));
      sheet.getRange(targetRow, 14).setValue(safeStr(data.fatLevel));
      sheet.getRange(targetRow, 15).setValue(safeStr(data.preWorkoutMeal));
      sheet.getRange(targetRow, 16).setValue(safeStr(data.preWorkoutTiming));
      sheet.getRange(targetRow, 17).setValue(safeStr(data.preWorkoutCarbs));
      sheet.getRange(targetRow, 18).setValue(safeNum(data.fat));
      sheet.getRange(targetRow, 19).setValue(safeNum(data.carbs));
      sheet.getRange(targetRow, 20).setValue(safeNum(data.fiber));
      sheet.getRange(targetRow, 21).setValue(safeNum(data.salt));
      sheet.getRange(targetRow, 22).setValue(safeNum(data.caffeine));
    }
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ===== 履歴削除：筋トレ種目（1種目＋全セット詳細） =====
function deleteTrainingExercise(exerciseId) {
  try {
    const ss        = getSpreadsheet();
    const trSheet   = ss.getSheetByName('トレーニング記録');
    const setsSheet = ss.getSheetByName('セット詳細');
    if (!trSheet) throw new Error('トレーニング記録シートが見つかりません');

    // トレーニング記録から該当行を削除（逆順で1行だけ）
    const trLastRow = trSheet.getLastRow();
    if (trLastRow > 1) {
      const trIds = trSheet.getRange(2, 1, trLastRow - 1, 1).getValues();
      for (let i = trIds.length - 1; i >= 0; i--) {
        if (String(trIds[i][0]) === String(exerciseId)) {
          trSheet.deleteRow(i + 2);
          break;
        }
      }
    }

    // セット詳細から該当IDに紐づく全行を削除（逆順）
    if (setsSheet && setsSheet.getLastRow() > 1) {
      const sIds = setsSheet.getRange(2, 1, setsSheet.getLastRow() - 1, 2).getValues();
      const toDelete = [];
      sIds.forEach(function(r, i) {
        if (String(r[1]) === String(exerciseId)) toDelete.push(i + 2);
      });
      for (let i = toDelete.length - 1; i >= 0; i--) {
        setsSheet.deleteRow(toDelete[i]);
      }
    }

    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ===== 履歴削除：筋トレセッション（複数種目＋全セット詳細） =====
function deleteTrainingSession(exerciseIds) {
  try {
    const ss        = getSpreadsheet();
    const trSheet   = ss.getSheetByName('トレーニング記録');
    const setsSheet = ss.getSheetByName('セット詳細');
    if (!trSheet) throw new Error('トレーニング記録シートが見つかりません');

    if (!exerciseIds || exerciseIds.length === 0) return { success: true };
    var ids = exerciseIds.map(function(id) { return String(id); }).filter(function(id) { return id; });

    // トレーニング記録から該当行を逆順で削除
    const trLastRow = trSheet.getLastRow();
    if (trLastRow > 1) {
      const trIds = trSheet.getRange(2, 1, trLastRow - 1, 1).getValues();
      const toDelete = [];
      trIds.forEach(function(r, i) {
        if (ids.indexOf(String(r[0])) >= 0) toDelete.push(i + 2);
      });
      for (let i = toDelete.length - 1; i >= 0; i--) {
        trSheet.deleteRow(toDelete[i]);
      }
    }

    // セット詳細から該当IDに紐づく全行を逆順で削除
    if (setsSheet && setsSheet.getLastRow() > 1) {
      const sIds = setsSheet.getRange(2, 1, setsSheet.getLastRow() - 1, 2).getValues();
      const toDelete = [];
      sIds.forEach(function(r, i) {
        if (ids.indexOf(String(r[1])) >= 0) toDelete.push(i + 2);
      });
      for (let i = toDelete.length - 1; i >= 0; i--) {
        setsSheet.deleteRow(toDelete[i]);
      }
    }

    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ===== 履歴編集：身体データ（数値のみ、写真は別関数） =====
function updateBodyRecord(data) {
  try {
    const ss    = getSpreadsheet();
    const sheet = ss.getSheetByName('身体データ記録');
    if (!sheet) throw new Error('身体データ記録シートが見つかりません');
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) throw new Error('記録が見つかりません');
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    let targetRow = -1;
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(data.id)) { targetRow = i + 2; break; }
    }
    if (targetRow === -1) throw new Error('記録ID ' + data.id + ' が見つかりません');
    function safeVal(v) { return (v !== undefined && v !== null) ? v : ''; }
    sheet.getRange(targetRow, 2).setValue(safeVal(data.date));
    sheet.getRange(targetRow, 3).setValue(safeVal(data.weight));
    sheet.getRange(targetRow, 4).setValue(safeVal(data.chest));
    sheet.getRange(targetRow, 5).setValue(safeVal(data.waist));
    sheet.getRange(targetRow, 6).setValue(safeVal(data.hip));
    sheet.getRange(targetRow, 7).setValue(safeVal(data.thigh));
    sheet.getRange(targetRow, 8).setValue(safeVal(data.memo));
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ===== 身体写真：標準部位定義（固定ID、冪等な補完に使用） =====
var _STD_PARTS = [
  { id: 'std-uncategorized-v1', name: '未分類' },
  { id: 'std-wholebody-v1',     name: '全身' },
  { id: 'std-chest-v1',         name: '胸' },
  { id: 'std-back-v1',          name: '背中' },
  { id: 'std-shoulder-v1',      name: '肩' },
  { id: 'std-upperarm-v1',      name: '上腕' },
  { id: 'std-forearm-v1',       name: '前腕' },
  { id: 'std-wrist-v1',         name: '手首' },
  { id: 'std-abs-v1',           name: '腹筋' },
  { id: 'std-glutes-v1',        name: 'お尻' },
  { id: 'std-thigh-v1',         name: '太もも' },
  { id: 'std-hamstring-v1',     name: 'ハムストリング' },
  { id: 'std-calf-v1',          name: 'ふくらはぎ' },
  { id: 'std-ankle-v1',         name: '足首' }
];

// Sheetsが文字列'TRUE'をboolean trueに自動変換するため、両方に対応するヘルパー
function _bodyBool(v) {
  return v === true || (typeof v === 'string' && v.toUpperCase() === 'TRUE');
}

// ===== 身体写真：シート初期化（冪等） =====
function _ensureBodyPhotoSheets(ss) {
  var partSheet = ss.getSheetByName('写真部位マスタ');
  if (!partSheet) {
    partSheet = ss.insertSheet('写真部位マスタ');
    partSheet.appendRow(['partId','partName','isStandard','isActive','createdAt']);
  }

  // 不足している標準部位のみ補完（名前で重複チェック、何度実行しても増殖しない）
  var lastRow = partSheet.getLastRow();
  var existingStdNames = {};
  if (lastRow > 1) {
    var checkRows = partSheet.getRange(2, 1, lastRow - 1, 4).getValues();
    checkRows.forEach(function(r) {
      if (_bodyBool(r[2])) existingStdNames[String(r[1])] = true; // isStandard=true の部位名を収集
    });
  }
  var now = new Date().toISOString();
  _STD_PARTS.forEach(function(part) {
    if (!existingStdNames[part.name]) {
      partSheet.appendRow([part.id, part.name, true, true, now]); // boolean で書き込む（文字列'TRUE'はSheetsが変換するため）
    }
  });

  var photoSheet = ss.getSheetByName('身体写真記録');
  if (!photoSheet) {
    photoSheet = ss.insertSheet('身体写真記録');
    photoSheet.appendRow(['photoId','bodyRecordId','date','partId','fileId','isActive','createdAt']);
  }
  return { partSheet: partSheet, photoSheet: photoSheet };
}

// ===== 写真システムデータ一括取得（部位マスタ＋全写真） =====
function getBodyPhotoSystemData() {
  try {
    var ss = getSpreadsheet();
    var sheets = _ensureBodyPhotoSheets(ss);

    var parts = [];
    var partLastRow = sheets.partSheet.getLastRow();
    if (partLastRow > 1) {
      var partRows = sheets.partSheet.getRange(2, 1, partLastRow - 1, 5).getValues();
      partRows.forEach(function(r) {
        parts.push({
          partId:     String(r[0]),
          partName:   String(r[1]),
          isStandard: _bodyBool(r[2]), // boolean/string 両対応
          isActive:   _bodyBool(r[3])  // boolean/string 両対応
        });
      });
    }

    var photos = {};
    var photoLastRow = sheets.photoSheet.getLastRow();
    if (photoLastRow > 1) {
      var photoRows = sheets.photoSheet.getRange(2, 1, photoLastRow - 1, 7).getValues();
      photoRows.forEach(function(r) {
        if (!_bodyBool(r[5])) return; // isActive（col6, index5）boolean/string 両対応
        var rid = String(r[1]);
        if (!photos[rid]) photos[rid] = [];
        photos[rid].push({ photoId: String(r[0]), date: String(r[2]), partId: String(r[3]), fileId: String(r[4]), isLegacy: false });
      });
    }

    // レガシー写真（身体データ記録.col10）を未分類として追加
    var bodySheet = ss.getSheetByName('身体データ記録');
    if (bodySheet && bodySheet.getLastRow() > 1) {
      var bodyRows = bodySheet.getRange(2, 1, bodySheet.getLastRow() - 1, 10).getValues();
      bodyRows.forEach(function(r) {
        var rid = String(r[0]);
        var fileId = String(r[9] || '');
        if (!fileId) return;
        var hasLegacy = (photos[rid] || []).some(function(p) { return p.isLegacy; });
        if (!hasLegacy) {
          if (!photos[rid]) photos[rid] = [];
          photos[rid].unshift({ photoId: 'LEGACY-' + rid, date: String(r[1] || ''), partId: 'LEGACY', fileId: fileId, isLegacy: true });
        }
      });
    }

    return { success: true, parts: parts, photos: photos };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ===== 部位追加 =====
function addBodyPart(partName) {
  try {
    var trimmed = String(partName || '').trim();
    if (!trimmed) return { success: false, error: '部位名を入力してください' };
    var ss = getSpreadsheet();
    var sheets = _ensureBodyPhotoSheets(ss);
    var lastRow = sheets.partSheet.getLastRow();
    if (lastRow > 1) {
      var names = sheets.partSheet.getRange(2, 2, lastRow - 1, 1).getValues();
      for (var i = 0; i < names.length; i++) {
        if (String(names[i][0]).trim() === trimmed) return { success: false, error: '同じ名前の部位が既に存在します' };
      }
    }
    var partId = Utilities.getUuid();
    sheets.partSheet.appendRow([partId, trimmed, false, true, new Date().toISOString()]); // boolean で書き込む
    return { success: true, partId: partId, partName: trimmed };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ===== 部位名変更（カスタム部位のみ） =====
function updateBodyPartName(partId, newName) {
  try {
    var trimmed = String(newName || '').trim();
    if (!trimmed) return { success: false, error: '部位名を入力してください' };
    var ss = getSpreadsheet();
    var sheets = _ensureBodyPhotoSheets(ss);
    var lastRow = sheets.partSheet.getLastRow();
    if (lastRow <= 1) return { success: false, error: '部位が見つかりません' };
    var rows = sheets.partSheet.getRange(2, 1, lastRow - 1, 4).getValues();
    var targetRow = -1;
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) === String(partId)) {
        if (_bodyBool(rows[i][2])) return { success: false, error: '標準部位は変更できません' }; // isStandard boolean対応
        targetRow = i + 2; break;
      }
    }
    if (targetRow === -1) return { success: false, error: '部位が見つかりません' };
    sheets.partSheet.getRange(targetRow, 2).setValue(trimmed);
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ===== 部位削除（使用中はソフト削除、未使用はハード削除） =====
function deleteBodyPart(partId) {
  try {
    var ss = getSpreadsheet();
    var sheets = _ensureBodyPhotoSheets(ss);
    var lastRow = sheets.partSheet.getLastRow();
    if (lastRow <= 1) return { success: false, error: '部位が見つかりません' };
    var rows = sheets.partSheet.getRange(2, 1, lastRow - 1, 4).getValues();
    var targetRow = -1;
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) === String(partId)) {
        if (_bodyBool(rows[i][2])) return { success: false, error: '標準部位は削除できません' }; // isStandard boolean対応
        targetRow = i + 2; break;
      }
    }
    if (targetRow === -1) return { success: false, error: '部位が見つかりません' };
    // 使用中チェック：col4=partId, col6=isActive → getRange(col4, 3列) = partId,fileId,isActive
    var inUse = false;
    var photoLastRow = sheets.photoSheet.getLastRow();
    if (photoLastRow > 1) {
      var usageRows = sheets.photoSheet.getRange(2, 4, photoLastRow - 1, 3).getValues(); // cols 4,5,6
      for (var j = 0; j < usageRows.length; j++) {
        if (String(usageRows[j][0]) === String(partId) && _bodyBool(usageRows[j][2])) { // partId(col4) & isActive(col6) boolean対応
          inUse = true; break;
        }
      }
    }
    if (inUse) {
      sheets.partSheet.getRange(targetRow, 4).setValue(false); // ソフト削除（boolean で書き込む）
    } else {
      sheets.partSheet.deleteRow(targetRow);
    }
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ===== 写真追加（新システム） =====
function addBodyPhotoRecord(data) {
  try {
    var ss = getSpreadsheet();
    var sheets = _ensureBodyPhotoSheets(ss);
    var uploadRes = uploadBodyPhoto(data.base64, data.mime, data.filename);
    if (!uploadRes.success) return uploadRes;
    var photoId = Utilities.getUuid();
    sheets.photoSheet.appendRow([photoId, String(data.recordId), String(data.date || ''), String(data.partId || ''), uploadRes.fileId, true, new Date().toISOString()]); // boolean
    return { success: true, photoId: photoId, fileId: uploadRes.fileId };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ===== 写真変更（新システム） =====
function changeBodyPhotoRecord(data) {
  try {
    var ss = getSpreadsheet();
    var sheets = _ensureBodyPhotoSheets(ss);
    var lastRow = sheets.photoSheet.getLastRow();
    if (lastRow <= 1) return { success: false, error: '写真が見つかりません' };
    var rows = sheets.photoSheet.getRange(2, 1, lastRow - 1, 5).getValues();
    var targetRow = -1; var oldFileId = '';
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.photoId)) { targetRow = i + 2; oldFileId = String(rows[i][4]); break; }
    }
    if (targetRow === -1) return { success: false, error: '写真が見つかりません' };
    var uploadRes = uploadBodyPhoto(data.base64, data.mime, data.filename);
    if (!uploadRes.success) return uploadRes;
    sheets.photoSheet.getRange(targetRow, 5).setValue(uploadRes.fileId);
    if (oldFileId) { try { Drive.Files.remove(oldFileId); } catch(ex) {} }
    return { success: true, fileId: uploadRes.fileId };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ===== 写真削除（新システム、ソフト削除＋Driveファイル削除） =====
function deleteBodyPhotoRecord(photoId) {
  try {
    var ss = getSpreadsheet();
    var sheets = _ensureBodyPhotoSheets(ss);
    var lastRow = sheets.photoSheet.getLastRow();
    if (lastRow <= 1) return { success: false, error: '写真が見つかりません' };
    var rows = sheets.photoSheet.getRange(2, 1, lastRow - 1, 6).getValues();
    var targetRow = -1; var fileId = '';
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) === String(photoId)) { targetRow = i + 2; fileId = String(rows[i][4]); break; }
    }
    if (targetRow === -1) return { success: false, error: '写真が見つかりません' };
    sheets.photoSheet.getRange(targetRow, 6).setValue(false); // ソフト削除（boolean）
    if (fileId) { try { Drive.Files.remove(fileId); } catch(ex) {} }
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ===== 部位変更（写真の部位を変更） =====
function updateBodyPhotoPartId(photoId, partId) {
  try {
    var ss = getSpreadsheet();
    var sheets = _ensureBodyPhotoSheets(ss);
    var lastRow = sheets.photoSheet.getLastRow();
    if (lastRow <= 1) return { success: false, error: '写真が見つかりません' };
    var rows = sheets.photoSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    var targetRow = -1;
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) === String(photoId)) { targetRow = i + 2; break; }
    }
    if (targetRow === -1) return { success: false, error: '写真が見つかりません' };
    sheets.photoSheet.getRange(targetRow, 4).setValue(String(partId));
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}
