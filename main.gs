/**
 * 本日出力済みフラグをクリアする
 * 動作スケジュール：毎日1回、bot発言が始まるまでに行う。実際の設定は0時台の実行とする。
 * 制約：日付の変わり目と0時台に発言する場合は、何らかの対策が必要。
 */
function clearFlag() {
  let funcName = 'clearFlag';
  try {
    console.log(`${funcName}: start`);
    
    // bot設定シート取得。別スプレッドシートなので、IDと名前で取得する。
    let settingSheet = fetchSheetByIdAndName(SETTING_BOOK_ID, SETTING_SHEET_NAME);
//    console.log(`${funcName}: sheetName=${settingSheet.getSheetName()}`);
    
    // クリア
    let lastRow = settingSheet.getLastRow();
//    console.log(`${ROW_DATA_START}, ${COL_TODAY_DONE}, ${lastRow - ROW_DATA_START + 1}, 1`);
    settingSheet.getRange(ROW_DATA_START, COL_TODAY_DONE, lastRow - ROW_DATA_START + 1, 1).setValue('');
    
  } catch(e) {
    console.log('[ERROR] Error occured.');
    console.log(`[ERROR] fileName:${e.fileName}`);
    console.log(`[ERROR] lineNumber:${e.lineNumber}`);
    console.log(`[ERROR] message:${e.message}`);
  } finally {
    console.log(`${funcName}: term`);
  }
}


/**
 * bot実行処理起点。設定シートを読み込み、発言タイミングに来ているものを発言する。
 * 動作スケジュール：毎分
 */
function startExecBot() {
  let funcName = 'startExecBot';
  try {
    console.log(`${funcName}: start`);
    // bot設定シート取得。別スプレッドシートなので、IDと名前で取得する。
    let settingSheet = fetchSheetByIdAndName(SETTING_BOOK_ID, SETTING_SHEET_NAME);
    
    // トークンシートと祝日シート取得。スクリプトはこのスプレッドシートに紐づいているので、名前で取得する。
    let tokenSheet = fetchSheetByName(ADMIN_TOKEN_SHEET_NAME);
    let holidaySheet = fetchSheetByName(ADMIN_HOLIDAY_SHEET_NAME);
    
    console.log(`settingSheet=${settingSheet.getSheetName()}`);
    console.log(`tokenSheet=${tokenSheet.getSheetName()}`);
    console.log(`holidaySheet=${holidaySheet.getSheetName()}`);

    execBotBySheets(settingSheet, tokenSheet, holidaySheet);
    
  } catch(e) {
    console.log('[ERROR] Error occured.');
    console.log(`[ERROR] fileName:${e.fileName}`);
    console.log(`[ERROR] lineNumber:${e.lineNumber}`);
    console.log(`[ERROR] message:${e.message}`);
  } finally {
    console.log(`${funcName}: term`);
  }
}


/**
 * 対象シートの指定を受け、botを実行する
 */
function execBotBySheets(settingSheet, tokenSheet, holidaySheet) {
  let funcName = 'execBotBySheets';
  console.log(`${funcName}: start`);
  
  // 結果出力時用の日時文字列
  let timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  
  // tokenとholidayを連想配列化する
  let tokens = convertKeyValueSheet(tokenSheet);
  let holidays = convertKeyValueSheet(holidaySheet);
  
  // 最終行を取得し、設定情報と併せて、すべての設定データを取得する
  let lastRow = settingSheet.getLastRow();
  let rows = settingSheet.getRange(ROW_DATA_START, 1, lastRow - ROW_DATA_START + 1, COL_LAST).getValues();
  
  // 設定データごとに処理
  rows.forEach(function(row){
    
    // 設定インスタンスを作成
    let setting = new ReminderSetting(row);
    
    console.log(`No. ${setting.num} start`);
    
    // チェック設定も実行設定もされてない場合、なにもしない
    if(!setting.isDryRun && !setting.isExec) {
      return;
    }
    
    // チェックする。エラーがある場合は結果欄に入力して終了。
    let checkResult = setting.check(tokens);
    if(checkResult.length > 0) {
      setting.writeResult(settingSheet, false, timestamp + ': ' + checkResult)
      return;
    }
    
    // 実行設定でない場合はdryRunなので、結果のみ出力
    if(!setting.isExec) {
      setting.writeResult(settingSheet, false, timestamp + ': dryrun result OK.')
      return;
    }
    
    // 実行判定を行い、実行対象でない場合は終了。
    if(!setting.isToExec(holidays, tokens)) {
      return;
    }
    
    // 実行し、結果が数字のみの文字列であれば成功と判定する。
    let execResult = setting.execBot(tokens);
    if(execResult.match(/[0-9]+/)) {
      setting.writeResult(settingSheet, true, timestamp + ': messageId=' + execResult)
    } else {
      setting.writeResult(settingSheet, false, timestamp + ': result=' + execResult)
    }
    
  });
  console.log(`${funcName}: term`);
}



