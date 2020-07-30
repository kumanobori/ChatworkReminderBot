/**
 * リマインダ設定クラス
 */
class ReminderSetting {
  
  /**
   * コンストラクタ
   */
  constructor(args) {
    
    // 引数の配列用の添字
    let i = 0;
    
    // 通番。シートへの書き込みを行う際に利用する。
    this.num = args[i++];
    // チェックのみを行うフラグ。
    this.isDryRun = args[i++] == CONST_YES ? true : false;
    // bot発言の実行を行うフラグ。
    this.isExec = args[i++] == CONST_YES ? true : false;
    // メモ。このクラスでは使わないが、シートをまるごと取り込む都合上、項目を保有。
    this.note = args[i++];
    // 発言者名。これをもとに、APIトークンを取得する。。
    this.speakerName = args[i++];
    // チャットルームID。
    this.roomId = args[i++];
    
    // 日時指定。
    let strDateTime = args[i++];
    this.detectedDateTime = strDateTime == '' ? null : new Date(strDateTime);
    
    // 曜日時間指定時の、時間指定。引数をもとに、「今日のその時刻」に加工する。
    let strTime = args[i++];
    if(strTime == '') {
      this.detectedTime = null;
    } else {
      
      // これは1899年12月30日の「指定時刻」
      let wkTime = new Date(strTime);　
      
      // 現在日時のDateを作って、時分秒を指定時刻に差し替える
      let wkDateTime = new Date(); // 
      wkDateTime.setHours(wkTime.getHours());
      wkDateTime.setMinutes(wkTime.getMinutes());
      wkDateTime.setSeconds(0);
      this.detectedTime = wkDateTime;
    }
    
    // 曜日時間指定時の、各曜日に実行するか否かのフラグ。
    this.onMonday = args[i++] == CONST_YES ? true : false;
    this.onTuesday = args[i++] == CONST_YES ? true : false;
    this.onWednesday = args[i++] == CONST_YES ? true : false;
    this.onThirsday = args[i++] == CONST_YES ? true : false;
    this.onFriday = args[i++] == CONST_YES ? true : false;
    this.onSaturday = args[i++] == CONST_YES ? true : false;
    this.onSunday = args[i++] == CONST_YES ? true : false;
    // 曜日時間指定時の、祝日には黙るかのフラグ。
    this.offOnHoliday = args[i++] == CONST_YES ? true : false;
    
    // 発言本文。
    this.body = args[i++];
    // 本日既に発言済みかを表すフラグ。
    this.isTodayDone = args[i++] == CONST_YES ? true : false;
    // 結果欄。このクラスでは使わないが、シートをまるごと取り込む都合上、項目を保有。
    this.result = args[i++];
  }
  
  /**
   * インスタンス内容の文字列出力
   */
  toString() {
    return(`ReminderSetting[
      num=${this.num}, isDryRun=${this.isDryRun}, isExec=${this.isExec}, 
      note=${this.note}, speakerName=${this.speakerName}, roomId=${this.roomId},
      detectedDateTime=${this.detectedDateTime}, detectedTime=${this.detectedTime},
      onMonday=${this.onMonday}, onTuesday=${this.onTuesday}, onWednesday=${this.onWednesday},
      onThirsday=${this.onThirsday}, onFriday=${this.onFriday}, onSaturday=${this.onSaturday},
      onSunday=${this.onSunday}, offOnHoliday=${this.offOnHoliday},
      body=${this.body}, isTodayDone=${this.isTodayDone}, result=${this.result}]`);
  }
  
  /**
   * 設定内容のチェックを行う。
   * @return エラーメッセージの配列。 エラー有無の判別は、result.length == 0 で行う
   */
  check(tokens) {
    
    let errors = [];
    
    if(isNaN(this.num)) {
      errors.push(`通番が数字以外です。${this.num}`);
    }
    if(!this.isDryRun && !this.isExec) {
      errors.push('設定チェックと実行のどちらにも設定されていません。');
    }
    if(this.speakerName == '') {
      errors.push('発言者名は必須です。');
    }
    if(this.roomId == '') {
      errors.push('発言部屋IDは必須です。');
    }
    
    let dateTimeValid = this.detectedDateTime != null && this.detectedDateTime.toString() != 'Invalid Date';
    let timeValid = this.detectedTime != null && this.detectedTime.toString() != 'Invalid Date';
    if(!dateTimeValid && !timeValid) {
      errors.push('日時指定と曜日時間指定の時刻の、両方が無効です。');
    }
    
    if(timeValid && !this.onMonday && !this.onTuesday && !this.onWednesday &&
       !this.onThirsday && !this.onFriday && !this.onSaturday && !this.onSunday) {
      errors.push('時間が指定されていますが、曜日がひとつも選択されていません。');
    }
    
    if(this.body == '') {
      errors.push('発言本文は必須です。');
    }
    
    let token = tokens[this.speakerName];
    if(token == undefined) {
      errors.push(`発言者名が無効です。${this.speakerName}`);
    } else {
      let client = ChatWorkClient.factory({token: token});
      let rooms = client.getRooms();
      let targetRoomFound = false
      let targetRoomId = this.roomId;
      rooms.forEach(function(room) {
        if(room.room_id == targetRoomId) {
          targetRoomFound = true;
        }
      });
      if(!targetRoomFound) {
        errors.push(`発言者が指定の部屋のメンバーではありません。${this.speakerName}:${this.roomId}`);
      }
    }
    
    // チェック結果を、発言判定に用いるため、保管する
    this.checkResult = errors.length == 0 ? true : false;
    
    return errors;
  }
  
  /**
   * bot発言を行うか否かを判定する
   * @return bot発言を行う場合はtrue、それ以外はfalse
   */
  isToExec(holidays, tokens) {
    let funcName = 'isToExec';
    // checkに引っかかっていた場合false
    if(!this.checkResult) {
      console.log(`${funcName}: check error false`);
      return false;
    }
    
    // 本日発言済みフラグが立っている場合false
    if(this.isTodayDone) {
      console.log(`${funcName}: TodayDone false`);
      return false;
    }
    
    // 発言すると判定する時間帯
    let targetDateTo = new Date();
    let targetDateFrom = new Date();
    targetDateFrom.setMinutes(targetDateFrom.getMinutes() - MAX_DELAY_MINUTES);
    let targetDateString = Utilities.formatDate(targetDateFrom, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm')
                         + ' - ' + Utilities.formatDate(targetDateTo, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');

    // 日時指定時の判定
    if(this.detectedDateTime != null) {
      let resultByDateTime = (targetDateFrom <= this.detectedDateTime && this.detectedDateTime <= targetDateTo);
      console.log(`${funcName}: detectedDateTime ${resultByDateTime}: ${Utilities.formatDate(this.detectedDateTime, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm')} : ${targetDateString}`);
      return resultByDateTime;
    }
    
    // 注：ここまで処理が来た時点で、曜日時間指定であることは確定している。
    
    // 祝日には黙る判定
    if(this.offOnHoliday) {
      let today = new Date();
      today.setHours(0);
      today.setMinutes(0);
      today.setSeconds(0);
      if(holidays[today] != null) {
        console.log(`${funcName}: holiday false`);
        return false;
      }
    }
    
    // 曜日の判定
    let weekday = targetDateTo.getDay();
    if((weekday == 0 && this.onSunday) ||
       (weekday == 1 && this.onMonday) ||
       (weekday == 2 && this.onTuesday) ||
       (weekday == 3 && this.onWednesday) ||
       (weekday == 4 && this.onThirsday) ||
       (weekday == 5 && this.onFriday) ||
       (weekday == 6 && this.onSaturday))
       {
         // 時間の判定
         let resultByTime = (targetDateFrom <= this.detectedTime && this.detectedTime <= targetDateTo);
         console.log(`${funcName}: weekday matched, detectedTime ${resultByTime}: ${Utilities.formatDate(this.detectedTime, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm')} : ${targetDateString}`);
         return resultByTime;
       } else {
         // 曜日不一致
         console.log(`${funcName}: weekday false`);
         return false;
       }
  }
  
  /**
   * botを実行する
   * @return 送信に成功した場合、メッセージID。そうでない場合、sendMessageの戻り値をそのまま返す。
   */
  execBot(tokens) {
    let client = ChatWorkClient.factory({token: tokens[this.speakerName]});
    let params = {room_id: this.roomId, body: this.body};
    let result = client.sendMessage(params);
    if(result != null && result.message_id != undefined) {
      return result.message_id;
    } else {
      return result;
    }
  }

  /**
   * 設定シートの更新を行う
   * @param sheet 設定シート
   * @param updateDoneFlg trueなら、「本日出力済」の列にフラグを入れる
   * @param logMessage 実行結果を示す文字列
   */
  writeResult(sheet, updateDoneFlg, logMessage) {
    let rowNo = this.num + ROW_DATA_START - 1;
    if(updateDoneFlg) {
      sheet.getRange(rowNo, COL_TODAY_DONE).setValue(CONST_YES);
    }
    sheet.getRange(rowNo, COL_RESULT).setValue(logMessage);
  }
}


/**
 * クラスの構築、チェック、送信判定のテスト
 */
function testConstruct() {
  
  let tokenSheet = fetchSheetByName(ADMIN_TOKEN_SHEET_NAME);
  let tokens = convertKeyValueSheet(tokenSheet);
  let holidaySheet = fetchSheetByName(ADMIN_HOLIDAY_SHEET_NAME);
  let holidays = convertKeyValueSheet(holidaySheet);
  
  let args;
  let obj;
  
  let minutesBefore = new Date();
  minutesBefore.setMinutes(new Date().getMinutes() - 4);
  let strMinutesBefore = Utilities.formatDate(minutesBefore, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  
  let sixMinutesBefore = new Date();
  sixMinutesBefore.setMinutes(new Date().getMinutes() - 6);
  let strSixMinutesBefore = Utilities.formatDate(sixMinutesBefore, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  
  let minutesAfter = new Date();
  minutesAfter.setMinutes(new Date().getMinutes() + 1);
  let strMinutesAfter = Utilities.formatDate(minutesAfter, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  let i = 1;
  args = [i, CONST_YES, CONST_YES, 'note', '稲葉', '195935405',
          strMinutesBefore,
          strMinutesBefore,
              CONST_YES, CONST_YES, CONST_YES, CONST_YES, CONST_YES, CONST_YES, CONST_YES, CONST_YES,
              'body',
              '',
              'log'];
          
  obj = new ReminderSetting(args);
  console.log(`${i} 通過するはず`);
  console.log(`${i} check: ${obj.check(tokens)}`);
  console.log(`${i} isToExec: ${obj.isToExec(holidays, tokens)}`);
  
  i = 'a';
  args = [i, CONST_YES, CONST_YES, 'note', '', '',
              '03:04',
              '12:34',
              CONST_YES, CONST_YES, CONST_YES, CONST_YES, CONST_YES, CONST_YES, CONST_YES, CONST_YES,
              '',
              '',
              'log'];

  obj = new ReminderSetting(args);
  console.log(`${i} チェックエラー１`);
  console.log(`${i} check: ${obj.check(tokens)}`);
  console.log(`${i} isToExec: ${obj.isToExec(holidays, tokens)}`);
  
  i = 2;
  args = [i, CONST_YES, CONST_YES, 'note', '山田', '195935409',
              '2019/01/02 03:04',
              '1899/12/30 12:34',
              '', '', '', '', '', '', '', CONST_YES,
              'body',
              '',
              'log'];
  obj = new ReminderSetting(args);
  console.log(`${i} チェックエラー２`);
  console.log(`${i} check: ${obj.check(tokens)}`);
  console.log(`${i} isToExec: ${obj.isToExec(holidays, tokens)}`);

  i = 3;
  args = [i, CONST_YES, CONST_YES, 'note', '稲葉', '195935405', strMinutesBefore, strMinutesBefore,
           CONST_YES, CONST_YES, CONST_YES, CONST_YES, CONST_YES, CONST_YES, CONST_YES, CONST_YES, 'body',CONST_YES,'log'];
  obj = new ReminderSetting(args);
  console.log(`${i} 本日発言済み`);
  console.log(`${i} check: ${obj.check(tokens)}`);
  console.log(`${i} isToExec: ${obj.isToExec(holidays, tokens)}`);
          
  i = 4;
  args = [i, CONST_YES, CONST_YES, 'note', '稲葉', '195935405', strMinutesAfter, strMinutesBefore,
           CONST_YES, CONST_YES, CONST_YES, CONST_YES, CONST_YES, CONST_YES, CONST_YES, CONST_YES, 'body','','log'];
  obj = new ReminderSetting(args);
  console.log(`${i} 日時指定、未来なので発言しない`);
  console.log(`${i} check: ${obj.check(tokens)}`);
  console.log(`${i} isToExec: ${obj.isToExec(holidays, tokens)}`);
  
  i = 5;
  args = [i, CONST_YES, CONST_YES, 'note', '稲葉', '195935405', strSixMinutesBefore, strMinutesBefore,
           CONST_YES, CONST_YES, CONST_YES, CONST_YES, CONST_YES, CONST_YES, CONST_YES, CONST_YES, 'body','','log'];
  obj = new ReminderSetting(args);
  console.log(`${i} 日時指定、6分遅れたので発言しない`);
  console.log(`${i} check: ${obj.check(tokens)}`);
  console.log(`${i} isToExec: ${obj.isToExec(holidays, tokens)}`);

  i = 6;
  let today = new Date();
  today.setHours(0);
  today.setMinutes(0);
  today.setSeconds(0);
  let testHolidays = {};
  testHolidays[today] = 'テスト用の祝日';
  args = [i, CONST_YES, CONST_YES, 'note', '稲葉', '195935405', '', strMinutesBefore,
           CONST_YES, CONST_YES, CONST_YES, CONST_YES, CONST_YES, CONST_YES, CONST_YES, CONST_YES, 'body','','log'];
  obj = new ReminderSetting(args);
  console.log(`${i} 曜日指定、祝日なので発言しない`);
  console.log(`${i} check: ${obj.check(tokens)}`);
  console.log(`${i} isToExec: ${obj.isToExec(testHolidays, tokens)}`);

  i = 7;
  args = [i, CONST_YES, CONST_YES, 'note', '稲葉', '195935405', '', strMinutesBefore,
           CONST_YES, CONST_YES, CONST_YES, CONST_YES, CONST_YES, CONST_YES, CONST_YES, '', 'body','','log'];
  obj = new ReminderSetting(args);
  console.log(`${i} 曜日指定、祝日だけど祝日黙るフラグがOFFなので発言する`);
  console.log(`${i} check: ${obj.check(tokens)}`);
  console.log(`${i} isToExec: ${obj.isToExec(testHolidays, tokens)}`);

  i = 8;
  args = [i, CONST_YES, CONST_YES, 'note', '稲葉', '195935405', '', strMinutesBefore,
           CONST_YES, '', '', '', '', '', '', '', 'body','','log'];
  obj = new ReminderSetting(args);
  console.log(`${i} 曜日指定、月曜`);
  console.log(`${i} check: ${obj.check(tokens)}`);
  console.log(`${i} isToExec: ${obj.isToExec(holidays, tokens)}`);

  i = 9;
  args = [i, CONST_YES, CONST_YES, 'note', '稲葉', '195935405', '', strMinutesBefore,
           '', CONST_YES, '', '', '', '', '', '', 'body','','log'];
  obj = new ReminderSetting(args);
  console.log(`${i} 曜日指定、火曜`);
  console.log(`${i} check: ${obj.check(tokens)}`);
  console.log(`${i} isToExec: ${obj.isToExec(holidays, tokens)}`);
  i = 10;
  args = [i, CONST_YES, CONST_YES, 'note', '稲葉', '195935405', '', strMinutesBefore,
           '', '', CONST_YES, '', '', '', '', '', 'body','','log'];
  obj = new ReminderSetting(args);
  console.log(`${i} 曜日指定、水曜`);
  console.log(`${i} check: ${obj.check(tokens)}`);
  console.log(`${i} isToExec: ${obj.isToExec(holidays, tokens)}`);
  i = 11;
  args = [i, CONST_YES, CONST_YES, 'note', '稲葉', '195935405', '', strMinutesBefore,
           '', '', '', CONST_YES, '', '', '', '', 'body','','log'];
  obj = new ReminderSetting(args);
  console.log(`${i} 曜日指定、木曜`);
  console.log(`${i} check: ${obj.check(tokens)}`);
  console.log(`${i} isToExec: ${obj.isToExec(holidays, tokens)}`);
  i = 12;
  args = [i, CONST_YES, CONST_YES, 'note', '稲葉', '195935405', '', strMinutesBefore,
           '', '', '', '', CONST_YES, '', '', '', 'body','','log'];
  obj = new ReminderSetting(args);
  console.log(`${i} 曜日指定、金曜`);
  console.log(`${i} check: ${obj.check(tokens)}`);
  console.log(`${i} isToExec: ${obj.isToExec(holidays, tokens)}`);
  i = 13;
  args = [i, CONST_YES, CONST_YES, 'note', '稲葉', '195935405', '', strMinutesBefore,
           '', '', '', '', '', CONST_YES, '', '', 'body','','log'];
  obj = new ReminderSetting(args);
  console.log(`${i} 曜日指定、土曜`);
  console.log(`${i} check: ${obj.check(tokens)}`);
  console.log(`${i} isToExec: ${obj.isToExec(holidays, tokens)}`);
  i = 14;
  args = [i, CONST_YES, CONST_YES, 'note', '稲葉', '195935405', '', strMinutesBefore,
           '', '', '', '', '', '', CONST_YES, '', 'body','','log'];
  obj = new ReminderSetting(args);
  console.log(`${i} 曜日指定、日曜`);
  console.log(`${i} check: ${obj.check(tokens)}`);
  console.log(`${i} isToExec: ${obj.isToExec(holidays, tokens)}`);

  i = 15;
  args = [i, CONST_YES, CONST_YES, 'note', '稲葉', '195935405', strMinutesBefore, '',
           '', '', '', '', '', '', '', '', 'test message','','log'];
  obj = new ReminderSetting(args);
  console.log(`${i} 日時指定、発言する`);
//  console.log(`${i} ${obj.toString()}`);
  console.log(`${i} check: ${obj.check(tokens)}`);
  console.log(`${i} isToExec: ${obj.isToExec(holidays, tokens)}`);
  let result = obj.execBot(tokens);
  console.log(`${i} execBot: ${result}`);

}
