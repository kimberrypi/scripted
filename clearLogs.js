function clearLogs() {
  let { SHEET_ID } = initialize();

  Logger = BetterLog.useSpreadsheet(SHEET_ID, "Logs");
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  const logSheet = spreadsheet.getSheetByName("Logs");

  logSheet.clearContents();
}
