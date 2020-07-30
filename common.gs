/**
 * スプレッドシートIDとシート名により、シートオブジェクトを取得する
 */
function fetchSheetByIdAndName(bookId, sheetName) {
  var book;
  var sheet;
  
  // ブックを取得
  try {
    book = SpreadsheetApp.openById(bookId);
  }
  catch(e) {
    throw new Error(`bookId=${bookId} のopenByIdに失敗しました。` + e)
  }
  
  // シートを取得
  // TODO: ここでシート名が不正の場合、catch節に入らずに実行終了してしまう。理由不明。
  try {
    sheet = book.getSheetByName(sheetName);
  }
  catch(e) {
    throw new Error(`bookId=${bookId} sheetName=${sheetName} のgetSheetByNameに失敗しました。` + e)
  }
  
  // 取得したシートを返す
  return sheet;
}

/**
 * アクティブなスプレッドシートから、シート名により、シートオブジェクトを取得する
 */
function fetchSheetByName(sheetName) {

  var book;
  var sheet;
  
  // ブックを取得
  try {
    book = SpreadsheetApp.getActiveSpreadsheet();
  }
  catch(e) {
    throw new Error(`getActiveSpreadsheetに失敗しました。` + e)
  }
  
  // シートを取得
  // TODO: ここでシート名が不正の場合、catch節に入らずに実行終了してしまう。理由不明。
  try {
    sheet = book.getSheetByName(sheetName);
  }
  catch(e) {
    throw new Error(`bookId=${bookId} sheetName=${sheetName} のgetSheetByNameに失敗しました。` + e)
  }
  
  // 取得したシートを返す
  return sheet;
}

/**
 * key-value構成のシートの内容を連想配列に変換する
 */
function convertKeyValueSheet(sheet) {
  let funcName = 'convertKeyValueSheet';
  let lastRowNo = sheet.getLastRow();
  let values = sheet.getRange(1,1,lastRowNo,2).getValues();
  let result = {};
//  console.log(`${funcName}: values: ${values}`);
  values.forEach(function(val){
    if(val[0] != '' && val[1] != '') {
      result[val[0]] = val[1];
    }
  });
  
//  for(let key in result) {
//    console.log(`${funcName}: ${key}: ${result[key]}`);
//  }
  return result;
}
