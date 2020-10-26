function BGMCOPY() {
  const getitem = function (sheetname, cellnumber) {
    const ss = SpreadsheetApp.getActiveSpreadsheet(); //スプレッドシート取得 変数に代入
    var sheet = ss.getSheetByName(sheetname); //シート取得　変数に代入
    var getcell = sheet.getRange(cellnumber); //セル取得
    const celltext = getcell.getValue(); //セルの要素取得
    return celltext;
    // console.log(celltext);//ログに出力
  };

  const getlastrow = function (url, sn) {
    //BGM管理票記入セル取得
    const urlss = SpreadsheetApp.openByUrl(url);
    var sheet = urlss.getSheetByName(sn);
    const lastrow = sheet
      .getRange(1, 4)
      .getNextDataCell(SpreadsheetApp.Direction.DOWN)
      .getRow();
    const returnrow = lastrow + 1;
    const rowtest = "${re}".replace("${re}", returnrow);
    return rowtest;
  };

  const sv = function (enterrow, entercolum, setitem) {
    //BGM管理票記入スクリプト
    const urlss = SpreadsheetApp.openByUrl(
      "https://docs.google.com/spreadsheets/d/1XRnzgmbMLXXgTldYC9b8kuixTlGPVpnjCHHA4jD_EJQ/edit#gid=1745565909"
    );
    var sheet = urlss.getSheetByName("Sheet1");
    const value = "${setitem}".replace("${setitem}", setitem);
    var cell = sheet.getRange(enterrow, entercolum);
    cell.setNumberFormat("@");
    cell.setValue(value);
  };

  const cell_write = function (cell, value) {
    const urlss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = urlss.getSheetByName("SW");
    sheet.getRange(cell).setValue(value);
  };

  const ps = "photo"; //写真商品のシート名に使用する変数を定義
  const vs = "VTR"; //VTR商品のシート名に使用する変数を定義
  const sw = "SW";

  var DAY = getitem(ps, "B5"); //挙式日を変数に代入
  DAY = Utilities.formatDate(DAY, "JST", "yyyy/MM/dd");

  const men_name = getitem(ps, "B8");
  const women_name = getitem(ps, "G8");
  const m_w_name = men_name + " " + women_name;

  const ed_movie = getitem(vs, "A11");

  const profile_movie = getitem(vs, "A16");

  const company_name = getitem(ps, "K2");

  const gest_telop = getitem(vs, "A26");

  const row = getlastrow(
    "https://docs.google.com/spreadsheets/d/1XRnzgmbMLXXgTldYC9b8kuixTlGPVpnjCHHA4jD_EJQ/edit#gid=1745565909",
    "Sheet1"
  );

  if (m_w_name != "" && DAY != "") {
    sv(row, "3", DAY);
    sv(row, "4", m_w_name);
    sv(row, "5", company_name);
    sv(row, "6", ed_movie);
    sv(row, "7", profile_movie);
    sv(row, "8", gest_telop);

    cell_write("E5", "OK");

    Browser.msgBox("BGM管理票への転記が完了しました。目視でも確認してください");
    sv(row, "1", "1");
  } else {
    Browser.msgBox("空白の部分があります。処理を中止します");
  }
}
