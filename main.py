function sendClaimEmailOnUpdate() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  
  // 데이터가 없으면 종료 (1행은 헤더)
  if (lastRow <= 1) return;

  // 가장 마지막에 추가된 행(신규 클레임) 가져오기
  var range = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn());
  var values = range.getValues()[0];

  // 메일 설정
  var recipient = "담당자이메일@example.com"; // 👈 수신자 이메일 입력
  var subject = "[신규 클레임] 새로운 클레임이 등록되었습니다. (ID: " + values[0] + ")";
  
  var body = "새로운 클레임 데이터가 스프레드시트에 등록되었습니다.\n\n" +
             "• 클레임 ID: " + values[0] + "\n" +
             "• 접수일: " + values[1] + "\n" +
             "• 고객명: " + values[2] + "\n" +
             "• 내용: " + values[3] + "\n\n" +
             "자세한 내용은 스프레드시트에서 확인하세요.";

  GmailApp.sendEmail(recipient, subject, body);
}
