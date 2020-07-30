// bot設定スプレッドシートのIDとシート名
const SETTING_BOOK_ID = '11Mq3wPyn42JubZFGF0HHLnLfwGKeWESHMZf5mYYij-Q';
const SETTING_SHEET_NAME = 'bot設定';

// トークンシートと祝日シートのシート名
const ADMIN_TOKEN_SHEET_NAME = 'TOKEN';
const ADMIN_HOLIDAY_SHEET_NAME = 'HOLIDAY';

// この分数以上遅延した場合は発言しない。言い換えると、設定が過去〇分以内のときに発言をする。毎分実行する前提なので、たぶん1か2くらいで十分
const MAX_DELAY_MINUTES = 5;

// シート内で、フラグtrueを表す文字列
const CONST_YES = 'Y';

// 行情報
const ROW_DATA_START = 4; // データ開始行

// 列情報
const COL_TODAY_DONE = 18; // 本日出力済み列
const COL_RESULT = 19; // 結果出力列
const COL_LAST = 19; // 最終列