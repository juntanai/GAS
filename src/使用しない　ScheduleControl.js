function ScheduleControl() {
  class ControlItem {
    //----------------各処理用関数と数値取得変数のまとめオブジェクト------------------------
    constructor() {
      //オブジェクト生成時に実行される処理
      this._getitems();
    }
    //-------------------------------------------基本は変更あってもここ以下のメソッドを触ればOK---------------------------------------
    _getitems() {
      //基本的なシート記載のアイテムを変数に格納する関数
      this.ps = "photo"; //写真商品のシート名に使用する変数を定義
      this.vs = "VTR"; //VTR商品のシート名に使用する変数を定義
      this.sw = "SW"; //SWシート名を変数に格納
      this.cs = "変更履歴"; //変更履歴シート名を変数に格納

      const totalCell = "A1:L80"; //シートの読み取り範囲全体のセル範囲指定

      const cutTotalCell = totalCell.split(""); //全体のセル範囲文字列を１文字ずつ配列格納
      this.lastColumn = cutTotalCell[3]; //セルのcolumの最も使われている範囲のアルファベットのみ抜き出す

      const psgets = this._setCell(this.ps, totalCell); //photoシート要素全取得
      this.psArray = this._margeArray(psgets); //photoシート要素１時配列化

      const vsgets = this._setCell(this.vs, totalCell); //photoシート要素全取得
      this.vsArray = this._margeArray(vsgets); //VTRシート要素１時配列化

      const swgets = this._setCell(this.sw, totalCell); //swシート要素全取得
      this.swArray = this._margeArray(swgets); //swシート要素1時配列化

      this.timeCellArray = ["A5", "A8", "A11", "A14", "A17", "A20"]; //施工管理表に挙式時間を記載する際のセルを挙式時間ごとに配列にまとめたもの
      this.nameCellArray = ["C5", "C8", "C11", "C14", "C17", "C20"]; //施工管理表に両家名を記載する際のセルを挙式時間ごとに配列にまとめたもの
      this.placeCellArray = ["C6", "C9", "C12", "C15", "C18", "C21"]; //施工管理表に披露宴会場を記載する際のセルを挙式時間ごとに配列にまとめたもの
      this.photoItemCellArray = ["C7", "C10", "C13", "C16", "C19", "C22"]; //施工管理表に写真商品を記載する際のセルを挙式時間ごとに配列にまとめたもの
      this.fmItemCellArray = ["G5", "G8", "G11", "G14", "G17", "G20"]; //施工管理表にフォーマル商品を記載する際のセルを挙式時間ごとに配列にまとめたもの
      this.vtrItemCellArray = ["G7", "G10", "G13", "G16", "G19", "G22"]; //施工管理表にVTR商品を記載する際のセルを挙式時間ごとに配列にまとめたもの
      this.personCellArray = ["J5", "J8", "J11", "J14", "J17", "J20"]; //施工管理表に指名カメラマンを記載する際のセルを挙式時間ごとに配列にまとめたもの

      this.ceremonyDay = this._getInfo("B", 5); //挙式日取得変数代入
      this.ceremonyTime = this._getInfo("B", 6); //挙式時間取得変数代入
      this.partyTime = this._getInfo("G", 6); //披露宴時間取得

      this.ceremonyDayFormat = this._dayFormat(this.ceremonyDay, "yyyy/MM/dd"); //挙式日程を年、月、日の形式にフォーマットする
      this.ceremonyTimeFormat = this._dayFormat(this.ceremonyTime, "HH:mm"); //挙式時間を時、分の形式にフォーマットする
      this.partyTimeFormat = this._dayFormat(this.partyTime, "HH:mm"); //披露宴時間を時、分の形式にフォーマットする

      this.staffName = this._getInfo("K", 2); //BP担当者名取得変数代入

      this.menName = this._getInfo("B", 8); //新郎名前取得変数代入
      this.womenName = this._getInfo("G", 8); //新婦名前取得変数代入
      this.MenWomenName = this.menName + "\t" + this.womenName; //施工管理表シートに記載する形式に新郎新婦名をフォーマットする

      this.partyRoomName = this._getInfo("G", 5); //披露宴会場名取得変数代入

      this.zipAdd = this._getInfo("A", 12); //お客様住所郵便番号取得変数代入

      this.photoItem = this._getInfo("F", 17); //写真商品取得変数代入
      this.photographer = this._getInfo("F", 25); //指名カメラマン商品取得変数代入

      this.fmItem = this._getInfo("A", 17); //フォーマル商品取得変数代入
      this.fmItemColor = this._getInfo("D", 17); //フォーマル台紙色取得変数代入
      this.fmPrice = this._getInfo("E", 17); //フォーマル価格取得変数代入

      this.vtrSetItem = this._getInfo("A", 3, this.vsArray); //VTRのセットアイテムを取得変数代入
      this.vtrRecItem = this._getInfo("A", 8, this.vsArray); //記録映像を取得変数代入
      this.vtrEndItem = this._getInfo("A", 11, this.vsArray); //エンドロールを取得変数代入
      this.vtrProfileItem = this._getInfo("A", 16, this.vsArray); //プロフィールを取得変数代入

      this.vtrItem = this._vtrItemCheck(); //施工管理表に記載するVTR商品のフィルタリング

      this.changeDay = this._getInfo("A", 63, this.swArray); //挙式変更日を変数に格納
      this.changeCeremonyTime = this._getInfo("B", 63, this.swArray); //挙式変更時間を変数に格納
      this.changePartyTime = this._getInfo("C", 63, this.swArray); //披露宴変更時間を変数に格納
      this.changePerson = this._getInfo("D", 63, this.swArray); //変更を受けた担当者名を変数に格納

      this.change = this._changecheck(); //挙式日の変更を検知するための数値を代入　０で初回記載、1で変更を感知
    }

    writeAnotherSheet(change = this.change) {
      //施工管理表に記載するための関数----------------------------------------

      this.CeremonyPartyTime =
        this.ceremonyTimeFormat + " " + "(" + this.partyTimeFormat + ")"; //施工管理表シートに記載する形式に時間をフォーマットする

      this.checkTime = this._dayFormat(this.ceremonyTime, "HHmm"); //挙式開始時間フォーマット変更（時間のみ文字列）

      this._cellGenerater(); //記載セルを取得

      this.sn = this._dayFormat(this.ceremonyDay, "dd"); //シート記入の際のシート名生成
      this._serchScheduleSheet(); //this.scheduleUrlの生成メソッド（記載する施工管理表を検索し、URLを検知する）

      if (change === "" || !change || change === 0) {
        //----------------------------------------初回記載時の処理--------------------------
        if (
          Browser.msgBox(
            "管理表転記内容は以下で差異はありませんか？",
            "記載管理表:" +
              this.setUrlSheetName +
              "\\n" +
              "お客様名:" +
              this.MenWomenName +
              "\\n" +
              "挙式日程:" +
              this.ceremonyDayFormat +
              "\\n" +
              "挙式時間:" +
              this.ceremonyTimeFormat +
              "\\n" +
              "披露宴時間:" +
              this.partyTimeFormat +
              "\\n" +
              "写真商品:" +
              this.photoItem +
              "\\n" +
              "フォーマル商品:" +
              this.fmItem +
              "\\n" +
              "VTR商品:" +
              this.vtrItem +
              "\\n" +
              "指名カメラマン" +
              this.photographer,
            Browser.Buttons.YES_NO
          ) == "no"
        ) {
          Browser.msgBox("処理を中止します。"); //確認のチェックボックスを表示させ、noの場合は処理を中止する
          return;
        }
        const urlStringCheck = "https://"; //urlをチェックするための変数

        if (
          this.scheduleUrl.indexOf(urlStringCheck) == "0" &&
          this.sn != "" &&
          change != 1
        ) {
          //施工管理表のURLがしっかりと記載されていれば以下の処理

          this._cellWriter();

          this._endProcessing();
        }
      } else if (change === 1) {
        //-------------------挙式日程変更時の処理------------------------------
        if (
          Browser.msgBox(
            "以下日程の施工管理表を消去し、挙式日変更を適用しますか？",
            "記載管理表:" +
              this.setUrlSheetName +
              "\\n" +
              "お客様名:" +
              this.MenWomenName +
              "\\n" +
              "挙式日程:" +
              this.ceremonyDayFormat +
              "\\n" +
              "挙式時間:" +
              this.ceremonyTimeFormat +
              "\\n" +
              "披露宴時間:" +
              this.partyTimeFormat +
              "\\n" +
              "写真商品:" +
              this.photoItem +
              "\\n" +
              "フォーマル商品:" +
              this.fmItem +
              "\\n" +
              "VTR商品:" +
              this.vtrItem +
              "\\n" +
              "指名カメラマン" +
              this.photographer,
            Browser.Buttons.YES_NO
          ) == "no"
        ) {
          Browser.msgBox("処理を中止します。"); //確認のチェックボックスを表示させ、noの場合は処理を中止する
          return;
        }
        this._cellErase(); //施工管理表にすでに入っている予定を消去する。

        this.changeDayFormat = this._dayFormat(this.changeDay, "yyyy/MM/dd"); //挙式変更日を年月日形式にフォーマット
        this.changeCeremonyTimeFormat = this._dayFormat(
          this.changeCeremonyTime,
          "HH:mm"
        ); //挙式時間を時、分の形式にフォーマットする
        this.changePartyTimeFormat = this._dayFormat(
          this.changePartyTime,
          "HH:mm"
        ); //披露宴時間を時、分の形式にフォーマットする

        this.CeremonyPartyTime =
          this.changeCeremonyTimeFormat +
          " " +
          "(" +
          this.changePartyTimeFormat +
          ")"; //施工管理表シートに記載する形式に変更時間をフォーマットする
        this.checkChangeTime = this._dayFormat(this.changeCeremonyTime, "HHmm"); //挙式開始時間フォーマット変更（時間のみ文字列）

        this._cellWriteActive("B5", this.changeDayFormat, this.ps); //変更後の挙式日に書き換える
        this._cellWriteActive("B6", this.changeCeremonyTimeFormat, this.ps); //変更後の挙式時間に書き換える
        this._cellWriteActive("G6", this.changePartyTimeFormat, this.ps); //変更後の披露宴時間に書き換える

        this.sn = this._dayFormat(this.changeDay, "dd"); //シート記入の際のシート名生成
        this._serchScheduleSheet(this.changeDay); //this.scheduleUrlの生成メソッド（記載する施工管理表を検索し、URLを検知する）

        console.log(this.scheduleUrl);
        this._cellGenerater(this.checkChangeTime); //変更後の挙式の時間で再度記入セルを生成する

        if (
          Browser.msgBox(
            "管理表転記内容は以下で差異はありませんか？",
            "記載管理表:" +
              this.setUrlSheetName +
              "\\n" +
              "お客様名:" +
              this.MenWomenName +
              "\\n" +
              "挙式日程:" +
              this.changeDayFormat +
              "\\n" +
              "挙式時間:" +
              this.changeCeremonyTimeFormat +
              "\\n" +
              "披露宴時間:" +
              this.changePartyTimeFormat +
              "\\n" +
              "写真商品:" +
              this.photoItem +
              "\\n" +
              "フォーマル商品:" +
              this.fmItem +
              "\\n" +
              "VTR商品:" +
              this.vtrItem +
              "\\n" +
              "指名カメラマン" +
              this.photographer,
            Browser.Buttons.YES_NO
          ) == "no"
        ) {
          Browser.msgBox("処理を中止します。"); //確認のチェックボックスを表示させ、noの場合は処理を中止する
          return;
        }

        const urlStringCheck = "https://";

        if (this.scheduleUrl.indexOf(urlStringCheck) == "0" && this.sn != "") {
          //施工管理表のURLがしっかりと記載されていれば以下の処理

          this._cellWriter(); //施工管理表に変更内容転記関数

          const changeWriteRow = this._getLastRow(this.cs, "1", "1"); //変更履歴シートに変更履歴を記入する行を特定する関数
          const ChangeRecordDayCell = "A" + changeWriteRow; //変更した日付を記入するセルを変数に格納
          const ChangeRecordTextCell = "B" + changeWriteRow; //変更した詳細を記載するセルを変数に格納
          const ChangeRecordText =
            "日程を" +
            "" +
            "挙式日" +
            " " +
            this.ceremonyDayFormat +
            " " +
            "挙式時間" +
            this.ceremonyTimeFormat +
            " " +
            "披露宴時間" +
            this.partyTimeFormat +
            " " +
            "から" +
            " " +
            "挙式日程" +
            " " +
            this.changeDayFormat +
            " " +
            "挙式時間" +
            this.changeCeremonyTimeFormat +
            "　" +
            "披露宴時間" +
            this.changePartyTimeFormat +
            "　" +
            "に変更しました";
          //変更履歴の詳細レポートを作成
          this._cellWriteActive(
            ChangeRecordDayCell,
            this._getNowDate(),
            this.cs
          ); //変更日程を記載
          this._cellWriteActive(
            ChangeRecordTextCell,
            ChangeRecordText,
            this.cs
          ); //変更詳細を記載

          //変更日程記載欄を初期化
          this._cellWriteActive("A63", "", this.sw);
          this._cellWriteActive("B63", "", this.sw);
          this._cellWriteActive(
            "C63",
            '=if(B63="","",B63+TIME(1,30,0))',
            this.sw
          );

          //swシートを隠す。
          this._sheet_hide();

          Browser.msgBox(
            "施工管理表への転記が完了しました。目視でも確認してください"
          );
        } else {
          Browser.msgBox("施工管理表URL等の記載がないため処理を中止します。");
          return;
        }
      }
    }

    _endProcessing() {
      //-------------------------------処理終了後の終了関数--------------------------------
      const customerAddress = this._address_call(); //郵便番号よりお客様の住所を取得する処理
      this._cellWriteActive("A13", customerAddress, this.ps); //お客様住所をセルにキャッシュさせる処理
      this._cellWriteActive("E8", "OK", this.sw); //記載チェックシートへのOK記載
      this._sheet_hide();

      Browser.msgBox(
        "施工管理表への転記が完了しました。目視でも確認してください"
      );
    }

    //-------------------------------------------------------------------------以下の部分は処理用関数のため基本修正必要なし------------------------------------------------------------------

    _setCell(sheet, range) {
      //読み取るシートとセルを選択する関数
      return SpreadsheetApp.getActiveSpreadsheet()
        .getSheetByName(sheet)
        .getRange(range)
        .getValues();
    }

    _margeArray(array) {
      //２次配列を１次配列に格納し直す関数
      return Array.prototype.concat.apply([], array);
    }
    _dayFormat(target, formatType) {
      //日付け、時間の変換関数
      return Utilities.formatDate(target, "JST", formatType);
    }

    _serchArrayPlus(array, key) {
      //keyに該当するセルを見つけ、その右隣のセルの情報を読みための関数
      return array[1 + array.indexOf(key)];
    }

    _getArrayNumber(column, row, column2 = "") {
      //スプレッドシートすべて取りこんだ配列よりセル番号で情報を引き出すための変換関数
      const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

      const coNumber = alphabet.indexOf(column);

      //      if(column2 != ""){
      //      const coNumber2 = alphabet.indexOf(column2);
      //      }
      const lastColumnNumber = alphabet.indexOf(this.lastColumn) + 1;
      return (row - 1) * lastColumnNumber + coNumber;
    }

    _getInfo(column, row, array = this.psArray) {
      //----------------配列よりスプレッドシートのセル番号から情報を取り出す関数------------------
      return array[this._getArrayNumber(column, row)];
    }

    _vtrItemCheck(
      set = this.vtrSetItem,
      rec = this.vtrRecItem,
      end = this.vtrEndItem
    ) {
      //-------------------VTR商品の有無判定関数------------------------
      if (set != "") {
        return set;
      } else if (rec != "") {
        return rec;
      } else if (end != "") {
        return end;
      } else if (!set && !rec && !end) {
        Browser.msgBox("VTR商品はありません");
      } else {
        Browser.msgBox("エラーを検知しました。VTRを確認してください");
      }
    }

    _serchScheduleSheet(day = this.ceremonyDay) {
      //以下施工管理表URL一覧シートより該当の施工管理表URLを検索する関数------------------------------------------
      const sheet_serch_day = this._dayFormat(day, "yyyy/MM");
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

      function input_value(set_sh, set_row) {
        //URLを実際に管理票シートより取得する
        var set_url = set_sh.getRange(set_row, 2).getValue();
        if (set_url === "" || !set_url) {
          Browser.msgBox("施工管理表が見つかりません");
        }
        return set_url;
      }
      function input_sheet_name_value(set_sh, set_row) {
        //記載する管理表名を実際に管理票シートより取得する
        var setSheetName = set_sh.getRange(set_row, 1).getValue();
        return setSheetName;
      }
      this.scheduleUrl = input_value(sche_sheet, row_get);
      this.setUrlSheetName = input_sheet_name_value(sche_sheet, row_get);
    }

    _cellreturn(cellset, timezone) {
      //-----------------------転記の際の記入セル判定関数--------------------------------------------------
      if (timezone === "" || !timezone) {
        Browser.msgBox("挙式時間が記入されていません");
        return;
      } else if (timezone <= "1000" && cellset[0] != "") {
        return cellset[0];
      } else if (timezone <= "1130" && cellset[1] != "") {
        return cellset[1];
      } else if (timezone <= "1330") {
        return cellset[2];
      } else if (timezone <= "1500") {
        return cellset[3];
      } else if (timezone <= "1630") {
        return cellset[4];
      } else if (timezone <= "1800") {
        return cellset[5];
      } else {
        Browser.msgBox("エラーです");
        return;
      }
    }

    _cellGenerater(timezone = this.checkTime) {
      //----------------------------------施工管理表の各記入セルの数値を生成する関数-------------------------
      this.timeCell = this._cellreturn(this.timeCellArray, timezone); //施工管理表に挙式時間を入れるセルの確定

      this.nameCell = this._cellreturn(this.nameCellArray, timezone); //施工管理表に両家名を入れるセルの確定

      this.placeCell = this._cellreturn(this.placeCellArray, timezone); //施工管理表に披露宴会場を入れるセルの確定

      this.photoCell = this._cellreturn(this.photoItemCellArray, timezone); //施工管理表に写真商品を入れるセルの確定

      this.fmCell = this._cellreturn(this.fmItemCellArray, timezone); //施工管理表にフォーマル商品を入れるセルの確定

      this.vtrCell = this._cellreturn(this.vtrItemCellArray, timezone); //施工管理表にVTR商品を入れるセルの確定

      this.personCell = this._cellreturn(this.personCellArray, timezone); //施工管理表に指名カメラマンを入れるセルの確定
    }

    _cellWrite(cell, value, url = this.scheduleUrl, sheetName = this.sn) {
      //-------------------指定シートの指定セルに記載させるための関数-----------------
      const urlss = SpreadsheetApp.openByUrl(url);
      var sheet = urlss.getSheetByName(sheetName);
      sheet.getRange(cell).setValue(value);
    }

    _cellWriter() {
      //----------------------施工管理表に各値をセルに記入する関数------------------------------
      this._cellWrite(this.timeCell, this.CeremonyPartyTime); //挙式時間と披露宴時間を記載

      this._cellWrite(this.nameCell, this.MenWomenName); //両家名を記載

      this._cellWrite(this.placeCell, this.partyRoomName); //披露宴場所の記載

      this._cellWrite(this.photoCell, this.photoItem); //写真商品の記載

      this._cellWrite(this.fmCell, this.fmItem); //フォーマル商品の記載

      this._cellWrite(this.vtrCell, this.vtrItem); //VTR商品の記載

      this._cellWrite(this.personCell, this.photographer); //指名フォトグラファーの記載
    }

    _cellErase() {
      //-------------------施工管理表各値を空白にする関数------------------------------------
      this._cellWrite(this.timeCell, ""); //挙式時間と披露宴時間を空白にする

      this._cellWrite(this.nameCell, ""); //両家名を空白にする

      this._cellWrite(this.placeCell, ""); //披露宴場所を空白にする

      this._cellWrite(this.photoCell, ""); //写真商品を空白にする

      this._cellWrite(this.fmCell, ""); //フォーマル商品を空白にする

      this._cellWrite(this.vtrCell, ""); //VTR商品を空白にする

      this._cellWrite(this.personCell, ""); //指名フォトグラファーを空白にする
    }

    _cellWriteActive(cell, value, name) {
      //------------------------アクティブなシートに値を記載する関数-------------------
      const urlss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = urlss.getSheetByName(name);
      sheet.getRange(cell).setValue(value);
    }

    _sheet_hide(sheetname = this.sw) {
      //------------------スイッチ群が記載されたシートを隠す処理----------------
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      ss.getSheetByName(sheetname).hideSheet();
    }

    _address_call(zip = this.zipAdd) {
      //-------------------お客様の住所を郵便番号より検索する処理--------------------------
      const response = UrlFetchApp.fetch(
        "http://zipcloud.ibsnet.co.jp/api/search?zipcode=" + zip
      );
      const results = JSON.parse(response.getContentText()).results;
      return results[0].address1 + results[0].address2 + results[0].address3;
    }

    _changecheck() {
      //---------------------挙式日の変更を検知する関数---------------------
      if (
        this.changeDay != "" ||
        this.changeCeremonyTime != "" ||
        this.changePartyTime != ""
      ) {
        return 1;
      } else {
        return 0;
      }
    }

    _getLastRow(sn, row, column) {
      //-----------------------------------BGM管理票記入セル取得---------------------------------
      const lastRowGet = SpreadsheetApp.getActiveSpreadsheet()
        .getSheetByName(sn)
        .getRange(row, column)
        .getNextDataCell(SpreadsheetApp.Direction.DOWN)
        .getRow();
      const returnRow = lastRowGet + 1;
      const lastRow = "${re}".replace("${re}", returnRow);
      return lastRow;
    }

    //-----------------------------引数に指定したアイテムを文字列に変換する関数---------------------------------
    _changeStr(setitem) {
      "${setitem}".replace("${setitem}", setitem);
    }

    //---------------------------この関数を実行した現在日時をかえす関数----------------------------
    _getNowDate() {
      const date = new Date();
      const now_date = Utilities.formatDate(date, "Asia/Tokyo", "yyyy/MM/dd");
      return now_date;
    }
  } //--------------------以上オブジェクト---------------------------------------------

  const getItemTest = new ControlItem();
  getItemTest.writeAnotherSheet();
}
