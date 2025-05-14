function testSheetName() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  SpreadsheetApp.getUi().alert("คุณกำลังใช้งานชีท: " + sheet.getName());
}
