function PROFILEMAIL() {
  const getitem = function (sheetname, cellnumber) {
    const ss = SpreadsheetApp.getActiveSpreadsheet(); //スプレッドシート取得 変数に代入
    var sheet = ss.getSheetByName(sheetname); //シート取得　変数に代入
    var getcell = sheet.getRange(cellnumber); //セル取得
    const celltext = getcell.getValue(); //セルの要素取得
    return celltext;
    // console.log(celltext);//ログに出力
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

  var time = getitem(ps, "B6");
  const fm_time = Utilities.formatDate(time, "JST", "HH:mm");

  const company_name = getitem(ps, "K2");

  var url = getitem(sw, "A22");

  var main_text_option = getitem(sw, "D35");

  const recipient = "ok@ogawakinya.kyoto"; //送信先のメールアドレス
  const subject =
    "P" + DAY + " " + fm_time + "挙式" + " " + m_w_name + "受注書";
  const recipientName = "小川";
  const body = `${recipientName}様\n\n\nお世話になっております。表題のお客様よりプロフィールの受注承りましたので受注書お送り致します。\n何卒よろしくお願い致します。\n${url}\n写真・テキストやBGM素材など、回収できましたら改めてご連絡させていただ
きます。\n\n\n${main_text_option}\n\n\n株式会社BRAINIG PICTURES\n${company_name}`;

  const options = { name: company_name, cc: "brainingapp3.28@gmail.com" };

  const string_check = "https://";

  if (url.indexOf(string_check) == "0") {
    GmailApp.sendEmail(recipient, subject, body, options);

    cell_write("E7", "OK");

    Browser.msgBox(
      "小川さんへのメールが送信されました。Gmailを確認してください"
    );
  } else {
    Browser.msgBox("共有URLが記載されていません。確認してください");
    return;
  }
}
