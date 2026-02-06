/**
 * 配送時間指定 管理ウェブアプリ - バックエンド
 *
 * スプレッドシートID は スクリプトプロパティ "SPREADSHEET_ID" に設定してください。
 * シート名: 時間指定DB / マスタ
 */

// ============================================================
// 定数
// ============================================================
var SHEET_DB   = '時間指定DB';
var SHEET_MASTER = 'マスタ';

// 時間指定DB のヘッダー (A~O)
var HEADERS = [
  '登録ID',    // A
  '登録日時',   // B
  '更新日時',   // C
  '配送日',    // D
  '担当者',    // E
  '伝票No',   // F
  '指定区分',   // G
  '指定時刻1',  // H
  '指定時刻2',  // I
  'CS有無',   // J
  '場所',     // K
  '理由',     // L
  '備考',     // M
  '削除フラグ',  // N
  '削除日時'   // O
];

// ============================================================
// ユーティリティ
// ============================================================

/**
 * スプレッドシートを取得
 */
function getSpreadsheet() {
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (id) {
    return SpreadsheetApp.openById(id);
  }
  // プロパティ未設定の場合はバインドされたスプレッドシートを使用
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * 指定シートを取得（なければ作成）
 */
function getOrCreateSheet(name) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (name === SHEET_DB) {
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    }
    if (name === SHEET_MASTER) {
      sheet.getRange(1, 1, 1, 1).setValue('担当者');
      sheet.getRange(1, 1, 1, 1).setFontWeight('bold');
    }
  }
  return sheet;
}

/**
 * UUID v4 生成
 */
function generateUUID() {
  var chars = 'abcdef0123456789';
  var sections = [8, 4, 4, 4, 12];
  var uuid = '';
  for (var s = 0; s < sections.length; s++) {
    if (s > 0) uuid += '-';
    for (var i = 0; i < sections[s]; i++) {
      uuid += chars.charAt(Math.floor(Math.random() * chars.length));
    }
  }
  return uuid;
}

/**
 * 行データをオブジェクトに変換
 */
function rowToObject(row) {
  var obj = {};
  for (var i = 0; i < HEADERS.length; i++) {
    var val = row[i];
    // 日付型を文字列に変換
    if (val instanceof Date) {
      if (HEADERS[i] === '配送日') {
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else if (HEADERS[i] === '指定時刻1' || HEADERS[i] === '指定時刻2') {
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'HH:mm');
      } else {
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
      }
    }
    // boolean 変換
    if (HEADERS[i] === '削除フラグ' || HEADERS[i] === 'CS有無') {
      val = (val === true || val === 'TRUE' || val === 'true');
    }
    obj[HEADERS[i]] = val;
  }
  return obj;
}

/**
 * 現在日時文字列
 */
function now() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

// ============================================================
// doGet - HTML を返却
// ============================================================
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('配送時間指定 管理')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ============================================================
// 登録 (Create)
// ============================================================
function createRecord(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { success: false, message: '他のユーザーが操作中です。しばらくしてから再度お試しください。' };
  }

  try {
    var sheet = getOrCreateSheet(SHEET_DB);
    var id = generateUUID();
    var timestamp = now();

    var row = [
      id,                          // A: 登録ID
      timestamp,                   // B: 登録日時
      timestamp,                   // C: 更新日時
      data.deliveryDate,           // D: 配送日
      data.staff,                  // E: 担当者
      data.slipNo || '',           // F: 伝票No
      data.category,               // G: 指定区分
      data.time1 || '',            // H: 指定時刻1
      data.time2 || '',            // I: 指定時刻2
      data.hasCS ? true : false,   // J: CS有無
      data.place,                  // K: 場所
      data.reason,                 // L: 理由
      data.note || '',             // M: 備考
      false,                       // N: 削除フラグ
      ''                           // O: 削除日時
    ];

    sheet.appendRow(row);
    return { success: true, message: '登録が完了しました。', id: id };
  } catch (e) {
    return { success: false, message: '登録に失敗しました: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// 配送日検索 (Read)
// ============================================================
function getRecordsByDeliveryDate(dateStr) {
  var sheet = getOrCreateSheet(SHEET_DB);
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  var data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  var results = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    // 削除フラグ チェック
    var delFlag = row[13]; // N列
    if (delFlag === true || delFlag === 'TRUE' || delFlag === 'true') continue;

    // 配送日 比較
    var deliveryDate = row[3]; // D列
    var dateVal;
    if (deliveryDate instanceof Date) {
      dateVal = Utilities.formatDate(deliveryDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } else {
      dateVal = String(deliveryDate);
    }

    if (dateVal === dateStr) {
      results.push(rowToObject(row));
    }
  }

  return results;
}

// ============================================================
// 編集 (Update)
// ============================================================
function updateRecord(id, data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { success: false, message: '他のユーザーが操作中です。しばらくしてから再度お試しください。' };
  }

  try {
    var sheet = getOrCreateSheet(SHEET_DB);
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { success: false, message: '対象レコードが見つかりません。' };
    }

    var allData = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
    var targetRowIndex = -1;

    for (var i = 0; i < allData.length; i++) {
      if (String(allData[i][0]) === String(id)) {
        targetRowIndex = i;
        break;
      }
    }

    if (targetRowIndex === -1) {
      return { success: false, message: '対象レコードが見つかりません。' };
    }

    var rowNum = targetRowIndex + 2; // ヘッダー行分 +1、0-index分 +1
    var timestamp = now();

    // 更新日時 (C列)
    sheet.getRange(rowNum, 3).setValue(timestamp);
    // 配送日 (D列)
    sheet.getRange(rowNum, 4).setValue(data.deliveryDate);
    // 担当者 (E列)
    sheet.getRange(rowNum, 5).setValue(data.staff);
    // 伝票No (F列)
    sheet.getRange(rowNum, 6).setValue(data.slipNo || '');
    // 指定区分 (G列)
    sheet.getRange(rowNum, 7).setValue(data.category);
    // 指定時刻1 (H列)
    sheet.getRange(rowNum, 8).setValue(data.time1 || '');
    // 指定時刻2 (I列)
    sheet.getRange(rowNum, 9).setValue(data.time2 || '');
    // CS有無 (J列)
    sheet.getRange(rowNum, 10).setValue(data.hasCS ? true : false);
    // 場所 (K列)
    sheet.getRange(rowNum, 11).setValue(data.place);
    // 理由 (L列)
    sheet.getRange(rowNum, 12).setValue(data.reason);
    // 備考 (M列)
    sheet.getRange(rowNum, 13).setValue(data.note || '');

    return { success: true, message: '更新が完了しました。' };
  } catch (e) {
    return { success: false, message: '更新に失敗しました: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// 削除 (論理削除)
// ============================================================
function deleteRecord(id) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { success: false, message: '他のユーザーが操作中です。しばらくしてから再度お試しください。' };
  }

  try {
    var sheet = getOrCreateSheet(SHEET_DB);
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { success: false, message: '対象レコードが見つかりません。' };
    }

    var allData = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
    var targetRowIndex = -1;

    for (var i = 0; i < allData.length; i++) {
      if (String(allData[i][0]) === String(id)) {
        targetRowIndex = i;
        break;
      }
    }

    if (targetRowIndex === -1) {
      return { success: false, message: '対象レコードが見つかりません。' };
    }

    var rowNum = targetRowIndex + 2;
    var timestamp = now();

    // 更新日時 (C列)
    sheet.getRange(rowNum, 3).setValue(timestamp);
    // 削除フラグ (N列)
    sheet.getRange(rowNum, 14).setValue(true);
    // 削除日時 (O列)
    sheet.getRange(rowNum, 15).setValue(timestamp);

    return { success: true, message: '削除が完了しました。' };
  } catch (e) {
    return { success: false, message: '削除に失敗しました: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// 集計 (期間 x 担当者)
// ============================================================
function getSummary(startDate, endDate, staff) {
  var sheet = getOrCreateSheet(SHEET_DB);
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return { staffCount: 0, totalCount: 0, ratio: 0, lastDate: '' };
  }

  var data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  var staffCount = 0;
  var totalCount = 0;
  var lastDate = '';

  for (var i = 0; i < data.length; i++) {
    var row = data[i];

    // 削除フラグ チェック
    var delFlag = row[13];
    if (delFlag === true || delFlag === 'TRUE' || delFlag === 'true') continue;

    // 登録日時(B列)から日付部分を取得
    var regDate = row[1];
    var dateVal;
    if (regDate instanceof Date) {
      dateVal = Utilities.formatDate(regDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } else {
      dateVal = String(regDate).substring(0, 10);
    }

    // 期間フィルタ（登録日ベース）
    if (dateVal < startDate || dateVal > endDate) continue;

    // 全体件数（期間内の全担当者分）
    totalCount++;

    // 担当者フィルタ
    var rowStaff = String(row[4]); // E列
    if (staff && staff !== '' && staff !== '全員') {
      if (rowStaff === staff) {
        staffCount++;
        if (dateVal > lastDate) {
          lastDate = dateVal;
        }
      }
    } else {
      // 全員の場合
      staffCount++;
      if (dateVal > lastDate) {
        lastDate = dateVal;
      }
    }
  }

  var ratio = totalCount > 0 ? Math.round((staffCount / totalCount) * 10000) / 100 : 0;

  return {
    staffCount: staffCount,
    totalCount: totalCount,
    ratio: ratio,
    lastDate: lastDate
  };
}

// ============================================================
// 担当者一覧取得（DBの過去データからユニークな担当者名を返す）
// ============================================================
function getStaffList() {
  var sheet = getOrCreateSheet(SHEET_DB);
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  var data = sheet.getRange(2, 5, lastRow - 1, 1).getValues(); // E列: 担当者
  var seen = {};
  var list = [];
  for (var i = 0; i < data.length; i++) {
    var val = String(data[i][0]).trim();
    if (val !== '' && !seen[val]) {
      seen[val] = true;
      list.push(val);
    }
  }
  list.sort();
  return list;
}

// ============================================================
// 初期セットアップ（手動実行用）
// ============================================================
function setupSheets() {
  getOrCreateSheet(SHEET_DB);
  getOrCreateSheet(SHEET_MASTER);
  Logger.log('シートのセットアップが完了しました。');
}
