// Version.13 - 공백 무시 & 셀에 직접 '메모(노트)' 삽입 및 누적 기능

//  [ Google Spread Sheet → 확장 프로그램 → Apps Script ]  //

function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[0];
    
    var rawData = e.parameter.data || "";
    if (!rawData) return ContentService.createTextOutput("데이터가 없습니다.");

    // 1. 데이터 분리 (항목/금액/메모)
    var match = rawData.match(/^([^\d-]+)\s*(-?\d+)\s*(.*)$/) || rawData.match(/^([^\d-]+)(-?\d+)(.*)$/);
    if (!match) return ContentService.createTextOutput("형식 오류!");

    var category = match[1].trim();
    var amount = parseInt(match[2]);
    var memo = match[3].trim();

    // 2. 날짜 및 위치 찾기
    var today = new Date();
    var yearMonth = Utilities.formatDate(today, "GMT+9", "yy.MM"); 
    sheet.getRange("A1").setValue(yearMonth);

    var rowIndex = today.getDate() + 1; 
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var colIndex = headers.indexOf(category) + 1;

    // 3. 데이터 입력 및 메모 삽입
    if (colIndex > 0) {
      var cell = sheet.getRange(rowIndex, colIndex);
      
      // 금액 합산
      var currentVal = Number(cell.getValue()) || 0;
      cell.setNumberFormat("#,##0"); 
      cell.setValue(currentVal + amount);

      // [핵심] 엑셀 메모(노트) 삽입 기능
      if (memo) {
        var existingNote = cell.getNote(); // 기존에 달린 메모 가져오기
        if (existingNote) {
          cell.setNote(existingNote + ", " + memo); // 있으면 쉼표로 이어붙이기
        } else {
          cell.setNote(memo); // 없으면 새로 만들기
        }
      }
      
      var msg = "✅ " + category + " " + amount + "원 기록 완료!";
      return ContentService.createTextOutput(memo ? msg + " (메모 삽입됨)" : msg);
    } else {
      return ContentService.createTextOutput("에러: '" + category + "' 열을 찾을 수 없습니다.");
    }
  } catch (err) {
    return ContentService.createTextOutput("오류: " + err.message);
  }
}



  // 매달 마지막 날 실행될 자동화 함수

function monthlyBackupAndClear() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var today = new Date();
  var currentMonth = today.getMonth() + 1; // 현재 월 (1~12)

  // 1. 달별 합계 행(33행)의 데이터 복사 (B33:H33)
  // 시트 기준으로 B33부터 H33까지가 합계 영역입니다.
  var monthlySummary = sheet.getRange(33, 2, 1, 7).getValues()[0];

  // 2. 아래쪽 월별 기록장(35~46행) 중 해당 월 행 찾기
  // 35행이 1월, 36행이 2월... 이므로 '34 + 월' 행이 됩니다.
  var targetRow = 34 + currentMonth;

  // 3. 해당 월 행에 데이터 붙여넣기 (B열부터 H열까지)
  sheet.getRange(targetRow, 2, 1, 7).setValues([monthlySummary]);

  // 4. 상단 일별 기록 데이터 초기화 (B2:H32 영역)
  // 다음 달을 위해 깔끔하게 비워줍니다.
  sheet.getRange(2, 2, 31, 7).clearContent();

  console.log(currentMonth + "월 데이터 백업 및 초기화 완료");

}
