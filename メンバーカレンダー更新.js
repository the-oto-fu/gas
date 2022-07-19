function renewSheetCalendar() {
  const book = SpreadsheetApp.openById('*****');
  const masterSheet = book.getSheetByName('メンバーマスタ');
  const members = masterSheet.getRange(3,2,8).getValues();

  const today = new Date();
  today.setHours(0,0,0,0);

  //カレンダーの開始行を指定
  let calendarRow = 2;

  //先頭のメンバーのを対象としてスクリプト実行当日日付のカレンダー行が存在するかをチェックし、基準日の月を調整する
  let workSheet = book.getSheetByName(members[0]);
  let todayCell = workSheet.getRange(calendarRow,1);
  let lastRow = todayCell.getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  let addMonth = 0;
  while(calendarRow <= lastRow){
    if(todayCell.getValue().getTime() == today.getTime()){
      addMonth = 1;
      break;
    }
    todayCell = todayCell.offset(1,0);
    calendarRow++;
  }

  //新しくカレンダーに追加する日付の開始日と終了日を定義する
  let referenceDate = new Date();
  referenceDate.setDate(1);
  referenceDate.setMonth(referenceDate.getMonth() + addMonth);
  let newMonthStartDate = new Date(referenceDate.getFullYear(), referenceDate.getMonth(), 1);
  let newMonthEndDate = new Date(referenceDate.getFullYear(), referenceDate.getMonth() + 1, 0);

  //メンバー全員分のシートを書き換える
  let currentDateCell = null;
  for(let member in members){
    console.log(members[member] + "の処理開始");
    let workSheet = book.getSheetByName(members[member]);

    //当日日付のセルが見つかった（翌月前に実行した）場合
    if(addMonth == 1){
      let todayRow = todayCell.getRow();
      //end行から今月の最後の行までを2行目まで移動
      let thisMonthRange = workSheet.getRange(todayRow,1,lastRow - todayRow + 1,25);
      workSheet.getRange(2,1,thisMonthRange.getNumRows(),25).setValues(thisMonthRange.getValues());

      //今月の最後の行からカウントアップし、最後の行までクリアしていく
      workSheet.getRange(thisMonthRange.getNumRows() + 2,1,lastRow - thisMonthRange.getNumRows(),25).clearContent();

      currentDateCell = workSheet.getRange(thisMonthRange.getNumRows() + 2,1);
    }else{
      workSheet.getRange(2,1, lastRow,25).clearContent();
      currentDateCell = workSheet.getRange(2,1);
    }

    //新しく追加する開始日から終了日までセルに入れる
    let currentDate = new Date(newMonthStartDate);
    while(currentDate <= newMonthEndDate){
      currentDateCell.setValue(currentDate);
      currentDate.setDate(currentDate.getDate() + 1);
      currentDateCell = currentDateCell.offset(1,0);
    }
  }
  return
}