function schedule() {
  const getitem = function (sheetname, cellnumber) {
    //特定スプレッドシートの特定セルの値取得関数

    const ss = SpreadsheetApp.getActiveSpreadsheet(); //スプレッドシート取得 変数に代入
    const sheet = ss.getSheetByName(sheetname); //シート取得　変数に代入
    const getcell = sheet.getRange(cellnumber); //セル取得
    const celltext = getcell.getValue(); //セルの要素取得
    return celltext;
  };

  const ps = "photo"; //写真商品のシート名に使用する変数を定義
  const vs = "VTR"; //VTR商品のシート名に使用する変数を定義
  const sw = "SW";

  const men_name = getitem(ps, "B8"); //新郎名前取得
  const women_name = getitem(ps, "G8"); //新婦名前取得
  const m_w_name = men_name + "\t" + women_name; //新郎新婦名前結合
  const place = getitem(ps, "G5"); ////披露宴会場名取得

  const day_get = getitem(ps, "B5");
  const ceremony_day_check = Utilities.formatDate(day_get, "JST", "yyyy/MM/dd");

  const time1 = getitem(ps, "B6"); //挙式開始時間取得
  const time = Utilities.formatDate(time1, "JST", "HHmm"); //挙式開始時間フォーマット変更（時間のみ文字列）

  const ceremony_time = Utilities.formatDate(time1, "JST", "HH:mm");

  const time2 = getitem(ps, "G6"); //披露宴開始時間取得
  const party_time = Utilities.formatDate(time2, "JST", "HH:mm"); //披露宴開始時間フォーマット変更（時間のみ文字列）

  const time_return = ceremony_time + " " + "(" + party_time + ")";

  const photo_item = getitem(ps, "F17"); //挙式当日写真商品取得

  const fm_item = getitem(ps, "A17"); //FM商品取得

  const person = getitem(ps, "F25");

  const VTR_set_item = getitem(vs, "A3"); //動画セット商品を取得
  const VTR_rec_item = getitem(vs, "A8"); //動画記録商品を取得
  const VTR_end_item = getitem(vs, "A11"); //動画エンドロール商品を取得

  const VTR_item_return = function (set, rec, end) {
    //VTR商品の有無判定関数
    if (set != "") {
      return set;
    } else if (rec != "") {
      return rec;
    } else if (end != "") {
      return end;
    } else {
      Browser.msgBox("VTR商品はありません");
    }
  };

  const VTR_item = VTR_item_return(VTR_set_item, VTR_rec_item, VTR_end_item); //上記関数実行

  const sheet_serch_day = Utilities.formatDate(day_get, "JST", "yyyy/MM"); //以下施工管理表URL一覧シートより該当の施工管理表URLを検索する
  var col = "A";
  var sche_spred = SpreadsheetApp.openByUrl(
    "https://docs.google.com/spreadsheets/d/1w7QtLBLsvm1q4ZWGDgrdJDuu0xUTz2CzE_EgkgKIFIY/edit#gid=0"
  );
  var sche_sheet = sche_spred.getSheetByName("施工管理表検索");

  function get_array(sh, col) {
    //施工管理表検索シート内のすべての項目を配列に格納する関数
    var last_row = sh.getLastRow();
    var range = sh.getRange(col + "1:" + col + last_row);
    var values = range.getValues();
    var array = [];
    for (var i = 0; i < values.length; i++) {
      array.push(values[i][0]);
    }
    return array;
  }
  const array = get_array(sche_sheet, col);

  function get_row(array_set, key) {
    //施工管理表URLより挙式日に合致するものを検索し、記載されているセルを返す
    var row = array_set.indexOf(key) + 1;
    return row;
  }
  const row_get = get_row(array, sheet_serch_day);
  console.log(row_get);

  function input_value(set_sh, set_row) {
    //URLを実際に管理票シートより取得する
    const set_url = set_sh.getRange(set_row, 2).getValue();
    return set_url;
  }
  const schedule_url = input_value(sche_sheet, row_get);

  const c_day = getitem(ps, "B5"); //挙式日付取得
  const day = Utilities.formatDate(c_day, "JST", "dd"); //挙式日程フォーマット変更（日付のみ文字列）

  const cellreturn = function (
    timezone,
    cell1000,
    cell1100,
    cell1300,
    cell1500,
    cell1630,
    cell1800
  ) {
    //転記の際の記入セル判定関数
    if (timezone === "") {
      Browser.msgBox("挙式時間が記入されていません");
    } else if (timezone <= "1000") {
      return cell1000;
    } else if (timezone <= "1130") {
      return cell1100;
    } else if (timezone <= "1330") {
      return cell1300;
    } else if (timezone <= "1500") {
      return cell1500;
    } else if (timezone <= "1630") {
      return cell1630;
    } else if (timezone <= "1800") {
      return cell1800;
    } else {
      Browser.msgBox("エラーです");
    }
  };

  const p_cell = cellreturn(time, "C7", "C10", "C13", "C16", "C19", "C22");

  const v_cell = cellreturn(time, "G7", "G10", "G13", "G16", "G19", "G22");

  const fm_cell = cellreturn(time, "G5", "G8", "G11", "G14", "G17", "G20");

  const name_cell = cellreturn(time, "C5", "C8", "C11", "C14", "C17", "C20");

  const place_cell = cellreturn(time, "C6", "C9", "C12", "C15", "C18", "C21");

  const person_cell = cellreturn(time, "J5", "J8", "J11", "J14", "J17", "J20");

  const time_cell = cellreturn(time, "A5", "A8", "A11", "A14", "A17", "A20");

  const cell_write = function (url, sn, cell, value) {
    //指定シートの指定セルに記載させるための関数
    const urlss = SpreadsheetApp.openByUrl(url);
    var sheet = urlss.getSheetByName(sn);
    sheet.getRange(cell).setValue(value);
  };

  if (
    Browser.msgBox(
      "管理表転記内容は以下で差異はありませんか？",
      "お客様名:" +
        m_w_name +
        "\\n" +
        "挙式日程:" +
        ceremony_day_check +
        "\\n" +
        "挙式時間:" +
        time +
        "\\n" +
        "写真商品:" +
        photo_item +
        "\\n" +
        "フォーマル商品:" +
        fm_item +
        "\\n" +
        "VTR商品:" +
        VTR_item +
        "\\n" +
        "指名カメラマン" +
        person,
      Browser.Buttons.YES_NO
    ) == "no"
  ) {
    Browser.msgBox("処理を中止します。"); //確認のチェックボックスを表示させ、noの場合は処理を中止する
  } else {
    //以下OKの場合の処理

    const string_check = "https://";

    if (schedule_url.indexOf(string_check) == "0" && day != "") {
      //施工管理表のURLがしっかりと記載されていれば以下の処理

      cell_write(schedule_url, day, p_cell, photo_item);
      cell_write(schedule_url, day, v_cell, VTR_item);
      cell_write(schedule_url, day, fm_cell, fm_item);
      cell_write(schedule_url, day, name_cell, m_w_name);
      cell_write(schedule_url, day, place_cell, place);
      cell_write(schedule_url, day, person_cell, person);
      cell_write(schedule_url, day, time_cell, time_return);

      const sheet_hide = function (sheetname) {
        //スイッチ群が記載されたシートを隠す処理
        var ss = SpreadsheetApp.getActiveSpreadsheet();

        ss.getSheetByName(sheetname).hideSheet();
      };

      sheet_hide("SW");

      const cell_write_check = function (cell, value) {
        //処理が実行されたら確認欄にOKを記載する処理
        const urlss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = urlss.getSheetByName("SW");
        sheet.getRange(cell).setValue(value);
      };

      cell_write_check("E8", "OK");

      const zipadd = getitem("photo", "A12");

      const address_call = function (zip) {
        //お客様の住所をセルにキャッシュさせる処理
        const response = UrlFetchApp.fetch(
          "http://zipcloud.ibsnet.co.jp/api/search?zipcode=" + zip
        );
        const results = JSON.parse(response.getContentText()).results;
        return results[0].address1 + results[0].address2 + results[0].address3;
      };

      const returnzip = address_call(zipadd);

      const cell_write_zip = function (value) {
        const urlss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = urlss.getSheetByName("photo");
        sheet.getRange("A13").setValue(value);
      };

      cell_write_zip(returnzip);

      Browser.msgBox(
        "施工管理表への転記が完了しました。目視でも確認してください"
      );
    } else {
      Browser.msgBox("施工管理表URL等の記載がないため処理を中止します。");
      return;
    }
  }
}
