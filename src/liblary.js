function Liblary() {
  class ControlItem {
    //----------------各処理用関数と数値取得変数のまとめオブジェクト------------------------
    constructor() {
      //オブジェクト生成時に実行される処理
      //-------------------------------------------基本は変更あってもここ以下のセル等指定部分を触ればOK---------------------------------------
      this.activeSheet = SpreadsheetApp.getActiveSpreadsheet();

      this.ps = "photo"; //写真商品のシート名に使用する変数を定義
      this.vs = "VTR"; //VTR商品のシート名に使用する変数を定義
      this.sw = "SW"; //SWシート名を変数に格納
      this.cs = "変更履歴"; //変更履歴シート名を変数に格納
      this.manageSheetUrl =
        "https://docs.google.com/spreadsheets/d/1w7QtLBLsvm1q4ZWGDgrdJDuu0xUTz2CzE_EgkgKIFIY/edit#gid=0"; //施工管理表URL一覧シートのURLを変数に格納

      this.bgmSheetUrl =
        "https://docs.google.com/spreadsheets/d/1XRnzgmbMLXXgTldYC9b8kuixTlGPVpnjCHHA4jD_EJQ/edit#gid=1745565909"; //BGM管理表URLを変数に格納

      const totalCell = "A1:L80"; //シートの読み取り範囲全体のセル範囲指定

      const cutTotalCell = totalCell.split(""); //全体のセル範囲文字列を１文字ずつ配列格納
      this.lastColumn = cutTotalCell[3]; //セルのcolumの最も使われている範囲のアルファベットのみ抜き出す

      const psgets = this._setCell(this.ps, totalCell); //photoシート要素全取得
      this.psArray = this._margeArray(psgets); //photoシート要素１時配列化

      const vsgets = this._setCell(this.vs, totalCell); //photoシート要素全取得
      this.vsArray = this._margeArray(vsgets); //VTRシート要素１時配列化

      const swgets = this._setCell(this.sw, totalCell); //swシート要素全取得
      this.swArray = this._margeArray(swgets); //swシート要素1時配列化

      this.ItemGetCellArray = {
        //各情報が記載されているセルの配列
        //photoシート記載項目------------------------------

        //日時関連
        ceremonyDay: "B5", //挙式日記載のセルを設定
        ceremonyTime: "B6", //挙式開始時間のセルを設定
        partyTime: "G6", //披露宴開始時間のセルを設定
        acceptDay: "G2", //受付日のセルを設定

        //名前関連
        staffName: "K2", //打ち合わせ担当BPスタッフが記載されているセルを設定
        menName: "B8", //新郎の名前記載のセルを設定
        womenName: "G8", //新婦の名前記載のセルを設定
        partyRoomName: "G5", //披露宴会場名記載のセルを設定

        //写真商品関連
        photoItem: "F17", //挙式当日写真商品記載のセルを設定
        photoItemPrice: "K17", //挙式当日写真商品価格記載のセルを設定

        photoOption1: "F19", //挙式当日写真オプション１記載のセルを設定
        photoOption1Price: "K17", //挙式当日写真オプション１価格記載のセルを設定

        photoOption2: "F21", //挙式当日写真オプション２記載のセルを設定

        photographer: "F25", //指名カメラマン記載のセルを設定

        //フォーマル商品関係
        fmItem: "A17",
        fmItemColor: "D17",
        fmPrice: "E17",

        //住所関連
        zipAdd: "A12",
        Adress1: "A13",
        Adress2: "E13",

        //VTRシート記載項目-------------------------------------------

        //vtr商品関係
        vtrSetItem: "A3",
        vtrRecItem: "A8",
        vtrEndItem: "A11",
        vtrProfileItem: "A16",

        //SWシート記載項目--------------------------------------------

        //変更関係
        changeDay: "A63", //変更後の挙式日程記載のセルを設定
        changeCeremonyTime: "B63", //変更後の挙式時間記載のセルを設定
        changePartyTime: "C63", //変更後の披露宴時間記載のセルを設定
        changePerson: "D63", //変更処理の担当者名記載のセルを設定
        changeWriteRowSet: "A1", //変更履歴シートの記載検知用セルを設定
        ChangeRecordDayCellSet: "A", //変更が行われた日程を記載する行を設定
        ChangeRecordTextCellSet: "B", //変更の詳細を記載する行を設定

        //swシートのチェック項目関係
        checkVtrMail: "E6",
        checkProfileMail: "E7",
        checkSchedule: "E8",

        activeSheetUrl: "A22", //共有URLが記載するセルを設定

        vtrMailTextOption: "D26",
        profileMailTextOption: "D35",
      };

      //施工管理表に記載する際の配列まとめ
      this.timeCellArray = ["A5", "A8", "A11", "A14", "A17", "A20"]; //施工管理表に挙式時間を記載する際のセルを挙式時間ごとに配列にまとめたもの
      this.nameCellArray = ["C5", "C8", "C11", "C14", "C17", "C20"]; //施工管理表に両家名を記載する際のセルを挙式時間ごとに配列にまとめたもの
      this.placeCellArray = ["C6", "C9", "C12", "C15", "C18", "C21"]; //施工管理表に披露宴会場を記載する際のセルを挙式時間ごとに配列にまとめたもの
      this.photoItemCellArray = ["C7", "C10", "C13", "C16", "C19", "C22"]; //施工管理表に写真商品を記載する際のセルを挙式時間ごとに配列にまとめたもの
      this.fmItemCellArray = ["G5", "G8", "G11", "G14", "G17", "G20"]; //施工管理表にフォーマル商品を記載する際のセルを挙式時間ごとに配列にまとめたもの
      this.vtrItemCellArray = ["G7", "G10", "G13", "G16", "G19", "G22"]; //施工管理表にVTR商品を記載する際のセルを挙式時間ごとに配列にまとめたもの
      this.personCellArray = ["J5", "J8", "J11", "J14", "J17", "J20"]; //施工管理表に指名カメラマンを記載する際のセルを挙式時間ごとに配列にまとめたもの

      //以上のセル指定部分を変更すれば簡易的な変更は対応可能（受注書のレイアウト変更など）

      //---------------------------------------各種処理を実行する際に必要な値を変数に格納------------------------------------
      this.ceremonyDay = this._getInfo("ceremonyDay"); //挙式日取得変数代入
      this.ceremonyTime = this._getInfo("ceremonyTime"); //挙式時間取得変数代入
      this.partyTime = this._getInfo("partyTime"); //披露宴時間取得

      this.ceremonyDayFormat = this._dayFormat(this.ceremonyDay, "yyyy/MM/dd"); //挙式日程を年、月、日の形式にフォーマットする
      this.ceremonyTimeFormat = this._dayFormat(this.ceremonyTime, "HH:mm"); //挙式時間を時、分の形式にフォーマットする
      this.partyTimeFormat = this._dayFormat(this.partyTime, "HH:mm"); //披露宴時間を時、分の形式にフォーマットする
      this.timeNameFormat = this._dayFormat(this.ceremonyDay, "yyyyMMdd"); //ファイル名として使用する形に時間をフォーマットする

      this.staffName = this._getInfo("staffName"); //BP担当者名取得変数代入

      this.menName = this._getInfo("menName"); //新郎名前取得変数代入
      this.womenName = this._getInfo("womenName"); //新婦名前取得変数代入
      this.MenWomenName = this.menName + "\t" + this.womenName; //施工管理表シートに記載する形式に新郎新婦名をフォーマットする

      this.partyRoomName = this._getInfo("partyRoomName"); //披露宴会場名取得変数代入

      this.zipAdd = this._getInfo("zipAdd"); //お客様住所郵便番号取得変数代入

      this.photoItem = this._getInfo("photoItem"); //写真商品取得変数代入
      this.photographer = this._getInfo("photographer"); //指名カメラマン商品取得変数代入

      this.fmItem = this._getInfo("fmItem"); //フォーマル商品取得変数代入
      this.fmItemColor = this._getInfo("fmItemColor"); //フォーマル台紙色取得変数代入
      this.fmPrice = this._getInfo("fmPrice"); //フォーマル価格取得変数代入

      this.vtrSetItem = this._getInfo("vtrSetItem", this.vsArray); //VTRのセットアイテムを取得変数代入
      this.vtrRecItem = this._getInfo("vtrRecItem", this.vsArray); //記録映像を取得変数代入
      this.vtrEndItem = this._getInfo("vtrEndItem", this.vsArray); //エンドロールを取得変数代入
      this.vtrProfileItem = this._getInfo("vtrProfileItem", this.vsArray); //プロフィールを取得変数代入

      this.vtrItem = this._vtrItemCheck(); //施工管理表に記載するVTR商品のフィルタリング

      this.changeDay = this._getInfo("changeDay", this.swArray); //挙式変更日を変数に格納
      this.changeCeremonyTime = this._getInfo(
        "changeCeremonyTime",
        this.swArray
      ); //挙式変更時間を変数に格納
      this.changePartyTime = this._getInfo("changePartyTime", this.swArray); //披露宴変更時間を変数に格納
      this.changePerson = this._getInfo("changePerson", this.swArray); //変更を受けた担当者名を変数に格納
      this.change = this._changecheck(); //挙式日の変更を検知するための数値を代入　０で初回記載、1で変更を感知

      this.vtrMailTextOption = this._getInfo("vtrMailTextOption", this.swArray); //VTR担当者へのメールに追記する部分を変数に格納
      this.profileMailTextOption = this._getInfo(
        "profileMailTextOption",
        this.swArray
      ); //PROFILE担当者へのメールに追記する部分を変数に格納
      this.MailOptions = {
        name: this.staffName,
        cc: "brainingapp3.28@gmail.com",
      };

      this.activeSheetUrl = this._getInfo("activeSheetUrl", this.swArray); //お客様受注書の共有URLを変数に格納
      this.vtrMailRecipent = "filmj0222@gmail.com"; //VTR担当者のメールアドレスを変数に格納
      this.vtrMailSubject = //VTR担当者へ送るメールの題名を変数に格納
        "V" +
        this.ceremonyDayFormat +
        " " +
        this.ceremonyTimeFormat +
        "挙式" +
        " " +
        this.MenWomenName +
        "受注書";
      this.vtrRecipientName = "田中　陣"; //VTR担当者の名前を変数に格納
      this.vtrMailBody = `${this.vtrRecipientName}様\nお世話になっております。表題のお客様よりVTR受注承りましたので受注書お送り致します。\n${this.activeSheetUrl}\n\n\n${this.vtrMailTextOption}\n\n何卒よろしくお願い致します。\n\n株式会社BRAINIG PICTURES\n${this.staffName}`;
      //VTR担当者へ送るメール本文を変数に格納

      this.profileMailRecipent = "ok@ogawakinya.kyoto"; //PROFILE担当者のメールアドレスを変数に格納
      this.profileMailSubject = //PROFILE担当者へ送るメールの題名を変数に格納
        "P" +
        this.ceremonyDayFormat +
        " " +
        this.ceremonyTimeFormat +
        "挙式" +
        " " +
        this.MenWomenName +
        "受注書";

      this.profileRecipientName = "小川"; //PROFILE担当者の名前を変数に格納
      this.profileMailBody = `${this.profileRecipientName}様\n\n\nお世話になっております。表題のお客様よりプロフィールの受注承りましたので受注書お送り致します。\n何卒よろしくお願い致します。\n${this.activeSheetUrl}\n写真・テキストやBGM素材など、回収できましたら改めてご連絡させていただ
        きます。\n\n\n${this.profileMailTextOption}\n\n\n株式会社BRAINIG PICTURES\n${this.staffName}`; //PROFILE担当者へ送るメール本文を変数に格納

      this.fileName = this.timeNameFormat + this.MenWomenName;
    }
    //以上必要な値の変数格納終了

    
    start() {//記入開始時に必ず使用する関数
      if (
        this.MenWomenName === "" ||
        !this.MenWomenName ||
        this.timeNameFormat === "" ||
        !this.timeNameFormat
      ) {
        Browser.msgBox("両家名、挙式日を記載してください");
        return;
      } else {
        this.activeSheet.rename(this.fileName);
        const nowDay = new Date();
        const NowDate = this._dayFormat(nowDay, "yyMMdd");
        this._cellWriteActive(
          this.ItemGetCellArray["acceptDay"],
          NowDate,
          this.ps
        );
      }
    }

    writeAnotherSheet(change = this.change) {//--------------------------------施工管理表に記載するための関数----------------------------------------
      this.CeremonyPartyTime =
        this.ceremonyTimeFormat + " " + "(" + this.partyTimeFormat + ")"; //施工管理表シートに記載する形式に時間をフォーマットする

      this.checkTime = this._dayFormat(this.ceremonyTime, "HHmm"); //挙式開始時間フォーマット変更（時間のみ文字列）

      this._cellGenerater(); //記載セルを取得

      this.sn = this._dayFormat(this.ceremonyDay, "dd"); //シート記入の際のシート名生成
      this._serchScheduleSheet(); //this.scheduleUrlの生成メソッド（記載する施工管理表を検索し、URLを検知する）

      if (change === "" || !change || change === 0) {
        //初回記載時の処理--------------------------
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

        this._cellWriteActive(
          this.ItemGetCellArray["ceremonyDay"],
          this.changeDayFormat,
          this.ps
        ); //変更後の挙式日に書き換える
        this._cellWriteActive(
          this.ItemGetCellArray["ceremonyTime"],
          this.changeCeremonyTimeFormat,
          this.ps
        ); //変更後の挙式時間に書き換える
        this._cellWriteActive(
          this.ItemGetCellArray["partyTime"],
          this.changePartyTimeFormat,
          this.ps
        ); //変更後の披露宴時間に書き換える

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

          const changeWriteRow = this._getLastRow(
            this.cs,
            this.ItemGetCellArray["changeWriteRowSet"]
          ); //変更履歴シートに変更履歴を記入する行を特定する関数
          const ChangeRecordDayCell =
            this.ItemGetCellArray["ChangeRecordDayCellSet"] + changeWriteRow; //変更した日付を記入するセルを変数に格納
          const ChangeRecordTextCell =
            this.ItemGetCellArray["ChangeRecordTextCellSet"] + changeWriteRow; //変更した詳細を記載するセルを変数に格納
          const ChangeRecordText =
            "日程を" +
            " " +
            "挙式日" +
            " " +
            this.ceremonyDayFormat +
            " " +
            "挙式時間" +
            " " +
            this.ceremonyTimeFormat +
            " " +
            "披露宴時間" +
            " " +
            this.partyTimeFormat +
            " " +
            "から" +
            " " +
            "挙式日程" +
            " " +
            this.changeDayFormat +
            " " +
            "挙式時間" +
            " " +
            this.changeCeremonyTimeFormat +
            "　" +
            "披露宴時間" +
            " " +
            this.changePartyTimeFormat +
            "　" +
            "に変更しました" +
            "　" +
            "担当者" +
            "　" +
            this.changePerson;
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
          this._cellWriteActive(
            this.ItemGetCellArray["changeDay"],
            "",
            this.sw
          );
          this._cellWriteActive(
            this.ItemGetCellArray["changeCeremonyTime"],
            "",
            this.sw
          );
          this._cellWriteActive(
            this.ItemGetCellArray["changePartyTime"],
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

    sendMailToProfile() {
      //--------------------PROFILE担当者へメールを送る関数--------------------------
      const string_check = "https://";

      if (this.activeSheetUrl.indexOf(string_check) == "0") {
        GmailApp.sendEmail(
          this.profileMailRecipent,
          this.profileMailSubject,
          this.profileMailBody,
          this.MailOptions
        );

        cell_write(this.ItemGetCellArray.checkProfileMail, "OK");

        Browser.msgBox(
          "小川さんへのメールが送信されました。Gmailを確認してください"
        );
      } else {
        Browser.msgBox("共有URLが記載されていません。確認してください");
        return;
      }
    }

    sendMailToVtr() {
      //---------------------------VTR担当者へのメール送信メソッド---------------------------
      const string_check = "https://";

      if (this.activeSheetUrl.indexOf(string_check) == "0") {
        GmailApp.sendEmail(
          this.vtrMailRecipent,
          this.vtrMailSubject,
          this.vtrMailBody,
          this.MailOptions
        );

        cell_write(this.ItemGetCellArray.checkVtrMail, "OK");

        Browser.msgBox(
          "陣さんへのメールが送信されました。Gmailを確認してください"
        );
      } else {
        Browser.msgBox("共有URLが記載されていません。確認してください");
        return;
      }
    }

    profitSheetWrite(){


    }


    _endProcessing() {
      //-------------------------------処理終了後の終了関数--------------------------------
      this._cellWriteActive(
        this.ItemGetCellArray["Adress1"],
        this._address_call(),
        this.ps
      ); //お客様住所をセルにキャッシュさせる処理
      this._cellWriteActive(
        this.ItemGetCellArray["checkSchedule"],
        "OK",
        this.sw
      ); //記載チェックシートへのOK記載
      this._sheet_hide();

      Browser.msgBox(
        "施工管理表への転記が完了しました。目視でも確認してください"
      );
    }

    //-------------------------------------------------------------------------以下の部分は処理用関数のため基本修正必要なし------------------------------------------------------------------

    _setCell(sheet, range) {
      //アクティブなスプレッドシートから読み取るシートとセルを選択する関数
      return SpreadsheetApp.getActiveSpreadsheet()
        .getSheetByName(sheet)
        .getRange(range)
        .getValues();
    }

    _margeArray(array) {
      //--------------------２次配列を１次配列に格納し直す関数
      return Array.prototype.concat.apply([], array);
    }
    _dayFormat(target, formatType) {
      //----------------------日付け、時間の文字列変換関数
      if (target != "") {
        return Utilities.formatDate(target, "JST", formatType);
      } else {
        return;
      }
    }
    _FormatSellSet(url, sh, cell, formattype) {
      //cellの書式を設定する関数
      SpreadsheetApp.openByUrl(url)
        .getSheetByName(sh)
        .getRange(cell)
        .setNumberFormat(formattype);
    }

    _serchArrayPlus(array, key) {
      //---------------------keyに該当するセルを見つけ、その右隣のセルの情報を読みための関数
      return array[1 + array.indexOf(key)];
    }

    _getArrayNumber(cell) {
      //------------------スプレッドシートすべて取りこんだ配列よりセル番号で情報を引き出すための変換関数
      const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

      const strCell = cell.split("");

      const strCellRow = Number(cell.slice(1));
      const coNumber = alphabet.indexOf(strCell[0]);

      const lastColumnNumber = alphabet.indexOf(this.lastColumn) + 1;
      return (strCellRow - 1) * lastColumnNumber + coNumber;
    }

    _getInfo(key1, array = this.psArray) {
      //----------------配列よりスプレッドシートのセル番号から情報を取り出す関数------------------
      return array[this._getArrayNumber(this.ItemGetCellArray[key1])];
    }

    _serchScheduleSheet(day = this.ceremonyDay,url = this.manageSheetUrl,sheetname = "施工管理表検索") {
      //以下施工管理表URL一覧シートより該当の施工管理表URLを検索する関数------------------------------------------
      const sheet_serch_day = this._dayFormat(day, "yyyy/MM");
      var col = "A";
      var sche_spred = SpreadsheetApp.openByUrl(url);
      var sche_sheet = sche_spred.getSheetByName(sheetname);

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

    _vtrItemCheck( //-------------------VTR商品の有無判定関数------------------------
      set = this.vtrSetItem,
      rec = this.vtrRecItem,
      end = this.vtrEndItem
    ) {
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

    _getLastRow(sn, cell, sheetGet = this.activeSheet) {
      //-----------------------------------行全体を検索し、空白の記入セル取得---------------------------------
      const lastRowGet = sheetGet
        .getSheetByName(sn)
        .getRange(cell)
        .getNextDataCell(SpreadsheetApp.Direction.DOWN)
        .getRow();
      const returnRow = lastRowGet + 1;
      const lastRow = "${re}".replace("${re}", returnRow);
      return lastRow;
    }

    _changeStr(setitem) {
      //-----------------------------引数に指定したアイテムを文字列に変換する関数---------------------------------
      "${setitem}".replace("${setitem}", setitem);
    }

    _getNowDate() {
      //---------------------------この関数を実行した現在日時をかえす関数----------------------------
      const date = new Date();
      const now_date = Utilities.formatDate(date, "Asia/Tokyo", "yyyy/MM/dd");
      return now_date;
    }

    _sheet_hide(sheetname = this.sw) {
      //------------------スイッチ群が記載されたシートを隠す処理----------------
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      ss.getSheetByName(sheetname).hideSheet();
    }

    _setCellUrl(url, sheet, range) {
      //スプレッドシートから読み取るシートとセルを選択する関数
      return SpreadsheetApp.openByUrl(url)
        .getSheetByName(sheet)
        .getRange(range)
        .getValues();
    }

    _getRowSerch(url, sn, cell, target) {
      //配列より検索用語に一致する要素のみ配列に格納する
      const namearray = this._setCellUrl(url, sn, cell);
      const rowNumberArray = namearray.filter(target);
      return rowNumberArray;
    }

    _serchIndex(array, key) {
      //配列内のkeyと一致した要素を新しい配列に格納する。
      const indexArray = [];
      array.forEach(function (el, index) {
        if (el === key) {
          indexArray.push(index + 1);
        }
      });
      return indexArray;
    }

    _totalValueStaff(url, columnArray, column) {
      //列番号の配列を利用して利用して同じ行の異なった列のデータを配列に格納する。
      const valueArray = [];
      columnArray.forEach(function (el) {
        const cell = column + String(el);
        valueArrray.push(this._setCellUrl(url, sheet, cell));
      });
      return valueArray;
    }
    _arrayTotal(valueArray) {
      //配列の要素を合計する関数
      return valueArray.reduce(function (a, b) {
        return a + b;
      });
    }
  }

  //--------------------以上オブジェクト---------------------------------------------
  const createMethods = function () {
    const methods = new ControlItem();
    return methods;
  };
  return createMethods();
}
