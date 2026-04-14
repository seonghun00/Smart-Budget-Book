// Version.9

function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[0];
    
    var rawData = e.parameter.data || "";
    if (!rawData) return ContentService.createTextOutput("데이터가 없습니다.");

    // 1. 문자열에서 항목(한글/영문)과 숫자(마이너스 포함) 분리
    // 숫자와 마이너스 기호만 남김
    var amountText = rawData.replace(/[^0-9-]/g, ""); 
    var amount = parseInt(amountText) || 0;
    
    // 숫자와 기호를 제외한 나머지를 항목으로 취급
    var category = rawData.replace(/[0-9-]/g, "").trim(); 

    if (!category || amount === 0) {
      return ContentService.createTextOutput("형식 오류! (예: 카페 5000 또는 카페 -5000)");
    }

    // 2. 날짜 정보 및 헤더 업데이트
    var today = new Date();
    var yearMonth = Utilities.formatDate(today, "GMT+9", "yy.MM"); 
    sheet.getRange("A1").setValue(yearMonth);

    // 3. 행(Row) 및 열(Column) 찾기
    var rowIndex = today.getDate() + 1; 
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var colIndex = headers.indexOf(category) + 1;

    // 4. 데이터 입력
    if (colIndex > 0) {
      var cell = sheet.getRange(rowIndex, colIndex);
      var currentVal = 0;
      
      try {
        var val = cell.getValue();
        currentVal = (typeof val === 'number' && !isNaN(val)) ? val : 0;
      } catch(e) {
        currentVal = 0;
      }
      
      cell.clearContent(); 
      cell.setNumberFormat("#,##0"); 
      cell.setValue(currentVal + amount); // 마이너스면 자동으로 차감됨
      
      var msg = amount > 0 ? " 기록 성공!" : " 차감 완료!";
      return ContentService.createTextOutput(category + " " + amount + "원" + msg);
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
  // 내 시트 기준으로 B33부터 H33까지가 합계 영역.
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

