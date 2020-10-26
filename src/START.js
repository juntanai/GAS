function START() {
  const date = new Date();
  const now_date = Utilities.formatDate(date, "Asia/Tokyo", "yyyy/MM/dd");
  const cell_write = function (value) {
    const urlss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = urlss.getSheetByName("photo");
    sheet.getRange("G2").setValue(value);
  };

  cell_write(now_date);

  Browser.msgBox(
    "スプレッドシートの許可が確認できました。記入を開始してください"
  );
}
