const sheets_id = 'ID of Spreadsheet';
const calendar_id = 'ID of Calendar';
const num_of_row = 200

function main() {

  // Spreadsheets "バイヨ" を開く  
  let sheets = SpreadsheetApp.openById(sheets_id);


  // セル E1 から年を取得
  let sheet = SpreadsheetApp.openById(sheets_id).getSheetByName('設定');
  const year = sheet.getRange('E1').getValue();

  // 設定したの年のシートを参照する or 作成 (例: 2023)
  sheet = set_sheet(year.toString());

  // シート内をすべてクリア
  sheet.clear();

  // シートの初期設定
  init(sheet)

  // カレンダーから予定を取得しセルに代入
  getCalendarEvents(year, sheet)

  // バイト先名称を一覧で並べる
  let office_names = [];
  j = 2
  
  for(let i = 2; i < num_of_row; i++){

    // A 列に並んでいるバイト先名称を順番に取得
    tmp_name = sheet.getRange(i, 1).getValue();
    
    // 入力済み以外のバイト先名称が来たら
    if(office_names.includes(tmp_name) == false){

      // 入力済みリストに push
      office_names.push(tmp_name);

      // セルに代入
      sheet.getRange(j, 6).setValue(tmp_name);

      j += 1   

    }

  }

  // 何も入っていない配列が入ってしまうので削除
  office_names.pop();

  // 給与入力用
  let salaries = [];
  let office_ids = {};
  let i = 0;
  for(let office_name of office_names){
    salaries.push([office_name, 0, 0]);  // [名称, 勤務時間, 給料]
    office_ids[office_name] = i;
    i += 1
  }

  // 給与をバイト先別に計算する

  // 表からバイト先の名称や勤務時間などを 2 次元配列として格納する
  let shifts = [];
  i = 2

  // 空白が来るまでシートからアルバイト情報を取得
  while(sheet.getRange(i, 1).isBlank() == false){
    shift = sheet.getRange(i, 1, 1, 4).getValues();
    shifts.push(shift[0]);
    i += 1    
  }
  
  // シフトごとに分析
  for(let shift of shifts){
    
    // バイト先名称、出勤時刻、退勤時刻、勤務時間。
    let arbeit_name = shift[0];
    let arbeit_start = shift[1];
    let arbeit_end = shift[2];
    let arbeit_time = shift[3];

    // 出勤および退勤時刻のうち hour のみ
    arbeit_start_hour = arbeit_start.getHours();
    arbeit_end_hour = arbeit_end.getHours();

    // 時間外労働時間
    let overtime = 0;

    // Logger.log(arbeit_start instanceof Date);

    // 給与計算

    // 深夜労働時間の算出
    night_shift_time = night_shift(arbeit_start_hour, arbeit_end_hour);

    // 時間外労働
    if(arbeit_time > 8){
      overtime = arbeit_time - 8
    }

    // 基本時給の取得
    sheet = SpreadsheetApp.openById(sheets_id).getSheetByName('設定')
    let hourly_wage = sheet.getRange(office_ids[arbeit_name] + 2, 2).getValue();

    // 給与の加算
    salaries[office_ids[arbeit_name]][2] += hourly_wage * (arbeit_time + night_shift_time * 0.25 + overtime * 0.25);

    // Logger.log("勤務先: %s, 勤務時間: %s, 時給: %s, 深夜勤務: %s, 時間外勤務: %s", arbeit_name, arbeit_time, hourly_wage, night_shift_time, overtime);

    // 勤務時間の加算
    salaries[office_ids[arbeit_name]][1] += arbeit_time;

  }

  // Logger.log(salaries);

  // 選択した年のシートを参照
  sheet = SpreadsheetApp.openById(sheets_id).getSheetByName(year.toString());
  
  // 年間合計給与
  sum_salary = 0

  for(let salary of salaries){
    
    // セルに格納
    sheet.getRange(office_ids[salary[0]] + 2, 7).setValue(salary[1]);
    sheet.getRange(office_ids[salary[0]] + 2, 8).setValue(salary[2]);
    
    // 年間合計給与を加算
    sum_salary += salary[2];
  
  }

  // セルに格納
  sheet.getRange('G10').setValue(sum_salary);

}


function set_sheet(sheet_name){

  // sheet_name という名前のシートがあるかどうか
  let sheet = SpreadsheetApp.openById(sheets_id).getSheetByName(sheet_name)

  // あれば OK
  if(sheet){
    return sheet;
  }

  // なければ作成して返す
  sheet = SpreadsheetApp.openById(sheets_id).insertSheet();
  sheet.setName(sheet_name);
  return sheet;

}

function init(sheet){

  // 行と列の幅を調整
  sheet.setRowHeights(1, num_of_row, 25);  // 1 から 200 行で幅 25
  sheet.setColumnWidths(1, 10, 150);  // A から J 列までで幅 150

  // 交互色の設定
  sheet.getRange(1, 1, num_of_row, 4).applyRowBanding(SpreadsheetApp.BandingTheme.CYAN)  // A1 から D200 まで

  // 時刻の書式設定
  sheet.getRange(2, 2, num_of_row, 2).setNumberFormat('yyyy/MM/dd hh:mm');

  // 各種入力
  sheet.getRange('A1').setValue('名称');
  sheet.getRange('B1').setValue('開始');
  sheet.getRange('C1').setValue('終了');
  sheet.getRange('D1').setValue('時間');
  sheet.getRange('F1').setValue('名称');
  sheet.getRange('G1').setValue('合計時間');
  sheet.getRange('H1').setValue('合計給与');
  sheet.getRange('F10').setValue('年間合計収入');

}


function getCalendarEvents(year, sheet){

  // Calendar "バイト" を開く
  let calendar = CalendarApp.getCalendarById(calendar_id);

  // 開始年月日と終了年月日
    // 何故か 1 月 1 日からだとうまくいかないので、昨年の 12 月 1 日からにしたらうまくいった。
  const start = new Date(year - 1, 12, 1);
  const end = new Date(year, 12, 31);

  // イベントを取得
  let events = calendar.getEvents(start, end);

  // イベントを取得しセルに入力
  i = 2
  for(let event of events){
    
    // タイトル
    sheet.getRange(i, 1).setValue(event.getTitle())
    
    // 開始日時
    // let start_date = Utilities.formatDate(event.getStartTime(), 'JST', 'yyyy/MM/dd HH:mm')
    let start_date = event.getStartTime();
    sheet.getRange(i, 2).setValue(start_date)
    
    // 終了日時
    // let end_date = Utilities.formatDate(event.getEndTime(), 'JST', 'yyyy/MM/dd HH:mm')
    let end_date = event.getEndTime();
    sheet.getRange(i, 3).setValue(end_date)
    
    // 勤務時間
    sheet.getRange(i, 4).setValue((event.getEndTime() - event.getStartTime()) / 1000 / 60 / 60)

    i += 1

  }

}


function night_shift(arbeit_start_hour, arbeit_end_hour){

  night_shift_time = 0;

  // 勤務時間に深夜労働時間が入っているか

  // 出勤時間が 22 時 ~ 23 時なら
  if(arbeit_start_hour >= 22){
    
    // 24 時以前に退勤しているなら
    if(arbeit_end_hour >= 23){
      night_shift_time += arbeit_end_hour - arbeit_start_hour;
    
    // 日付が変わっても勤務しているなら
    }else{
      
      // 退勤時間が 5 時以降なら
      if(arbeit_end_hour >= 5){
        night_shift_time += 5 + (24 - arbeit_start_hour) % 24;
      
      // 退勤時間が 5 時以前なら
      }else{
        night_shift_time += arbeit_end_hour + (24 - arbeit_start_hour) % 24;
      }
    }
  
  // 出勤時間が 0 時以降なら
  }else if(arbeit_start_hour >= 0 && arbeit_start_hour <= 5){

    // 退勤時間が 5 時以降なら
    if(arbeit_end_hour >= 5){
      night_shift_time += 5 - arbeit_start_hour;
    
    // 退勤時間が 5 時以前なら
    }else{
      night_shift_time += arbeit_end_hour - arbeit_start_hour;

    }

  }

  // Logger.log(night_shift_time);
  
  return night_shift_time;

}