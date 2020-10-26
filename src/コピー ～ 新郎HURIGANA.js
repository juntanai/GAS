const MEN_HURIGANA_SW = function () {
  var perform = function (output_type, sentence) {
    var endpoint = "https://labs.goo.ne.jp/api/hiragana";
    var payload = {
      app_id:
        "92c51c7a8e684d0c9de07dfe45b7f5272567de9007fbf8e4785c6c0fda3b5829",
      sentence: sentence,
      output_type: output_type,
    };
    var options = {
      method: "post",
      payload: payload,
    };

    var response = UrlFetchApp.fetch(endpoint, options);
    var response_json = JSON.parse(response.getContentText());
    return response_json.converted;
  };

  function KATAKANA(input) {
    return perform("katakana", input);
  }

  const getitem = function (sheetname, cellnumber) {
    const ss = SpreadsheetApp.getActiveSpreadsheet(); //スプレッドシート取得 変数に代入
    var sheet = ss.getSheetByName(sheetname); //シート取得　変数に代入
    var getcell = sheet.getRange(cellnumber); //セル取得
    const celltext = getcell.getValue(); //セルの要素取得
    return celltext;
    // console.log(celltext);//ログに出力
  };

  const men_name = getitem("photo", "B8");

  const men_name_kana = KATAKANA(men_name);

  const cell_write = function (value) {
    const urlss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = urlss.getSheetByName("photo");
    sheet.getRange("B7").setValue(value);
  };

  cell_write(men_name_kana);
};
