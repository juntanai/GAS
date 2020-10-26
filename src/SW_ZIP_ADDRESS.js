function SW_ZIP_ADD() {
  const getitem = function (sheetname, cellnumber) {
    const ss = SpreadsheetApp.getActiveSpreadsheet(); //スプレッドシート取得 変数に代入
    var sheet = ss.getSheetByName(sheetname); //シート取得　変数に代入
    var getcell = sheet.getRange(cellnumber); //セル取得
    const celltext = getcell.getValue(); //セルの要素取得
    return celltext;
    // console.log(celltext);//ログに出力
  };

  const zipadd = getitem("photo", "A12");

  const address_call = function (zip) {
    const response = UrlFetchApp.fetch(
      "http://zipcloud.ibsnet.co.jp/api/search?zipcode=" + zip
    );
    const results = JSON.parse(response.getContentText()).results;
    return results[0].address1 + results[0].address2 + results[0].address3;
  };

  const returnzip = address_call(zipadd);

  const cell_write = function (value) {
    const urlss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = urlss.getSheetByName("photo");
    sheet.getRange("A13").setValue(value);
  };

  cell_write(returnzip);

  Browser.msgBox("住所の書き換えを完了しました。受注書を確認してください");
}
