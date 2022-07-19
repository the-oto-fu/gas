function renewSheetCalendar() {
  const book = SpreadsheetApp.openById('*****');
  const masterSheet = book.getSheetByName('�����o�[�}�X�^');
  const members = masterSheet.getRange(3,2,8).getValues();

  const today = new Date();
  today.setHours(0,0,0,0);

  //�J�����_�[�̊J�n�s���w��
  let calendarRow = 2;

  //�擪�̃����o�[�̂�ΏۂƂ��ăX�N���v�g���s�������t�̃J�����_�[�s�����݂��邩���`�F�b�N���A����̌��𒲐�����
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

  //�V�����J�����_�[�ɒǉ�������t�̊J�n���ƏI�������`����
  let referenceDate = new Date();
  referenceDate.setDate(1);
  referenceDate.setMonth(referenceDate.getMonth() + addMonth);
  let newMonthStartDate = new Date(referenceDate.getFullYear(), referenceDate.getMonth(), 1);
  let newMonthEndDate = new Date(referenceDate.getFullYear(), referenceDate.getMonth() + 1, 0);

  //�����o�[�S�����̃V�[�g������������
  let currentDateCell = null;
  for(let member in members){
    console.log(members[member] + "�̏����J�n");
    let workSheet = book.getSheetByName(members[member]);

    //�������t�̃Z�������������i�����O�Ɏ��s�����j�ꍇ
    if(addMonth == 1){
      let todayRow = todayCell.getRow();
      //end�s���獡���̍Ō�̍s�܂ł�2�s�ڂ܂ňړ�
      let thisMonthRange = workSheet.getRange(todayRow,1,lastRow - todayRow + 1,25);
      workSheet.getRange(2,1,thisMonthRange.getNumRows(),25).setValues(thisMonthRange.getValues());

      //�����̍Ō�̍s����J�E���g�A�b�v���A�Ō�̍s�܂ŃN���A���Ă���
      workSheet.getRange(thisMonthRange.getNumRows() + 2,1,lastRow - thisMonthRange.getNumRows(),25).clearContent();

      currentDateCell = workSheet.getRange(thisMonthRange.getNumRows() + 2,1);
    }else{
      workSheet.getRange(2,1, lastRow,25).clearContent();
      currentDateCell = workSheet.getRange(2,1);
    }

    //�V�����ǉ�����J�n������I�����܂ŃZ���ɓ����
    let currentDate = new Date(newMonthStartDate);
    while(currentDate <= newMonthEndDate){
      currentDateCell.setValue(currentDate);
      currentDate.setDate(currentDate.getDate() + 1);
      currentDateCell = currentDateCell.offset(1,0);
    }
  }
  return
}