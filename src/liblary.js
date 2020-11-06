function Liblary() {
  class ControlItem {
    //----------------各処理用関数と数値取得変数のまとめオブジェクト------------------------
    constructor() {
      //オブジェクト生成時に実行される処理
      //-------------------------------------------基本は変更あってもここ以下のセル等指定部分を触ればOK---------------------------------------
      this.activeSheet = SpreadsheetApp.getActiveSpreadsheet();
      //受注書関連シート名
      this.ps = "photo"; //写真商品のシート名に使用する変数を定義
      this.vs = "VTR"; //VTR商品のシート名に使用する変数を定義
      this.sw = "SW"; //SWシート名を変数に格納
      this.cs = "変更履歴"; //変更履歴シート名を変数に格納

      //売上管理表関連シート名
      this.profitSheetName = {
        当日写真商品: "当日写真商品",
        当日フォーマル商品: "当日フォーマル商品",
        当日VTR商品: "当日VTR商品",
        南青山ルアンジェ当日商品総売上: "南青山ル・アンジェ当日商品総売上",
        ルアンジェ前撮り商品: "ル・アンジェ前撮り商品 ",
        南青山ルアンジェ担当売上前撮り除く:
          "南青山ル・アンジェ担当売上(前撮り除く)",
        青山店前撮り商品: "青山店前撮り商品 ",
        テラス前撮り商品: "テラス前撮り商品 ",
      };

      this.manageSheetUrl =
        "https://docs.google.com/spreadsheets/d/1w7QtLBLsvm1q4ZWGDgrdJDuu0xUTz2CzE_EgkgKIFIY/edit#gid=0"; //施工管理表URL一覧シートのURLを変数に格納

      this.bgmSheetUrl =
        "https://docs.google.com/spreadsheets/d/1XRnzgmbMLXXgTldYC9b8kuixTlGPVpnjCHHA4jD_EJQ/edit#gid=1745565909"; //BGM管理表URLを変数に格納

      this.profitSheetUrl =
        "https://docs.google.com/spreadsheets/d/1oiLscyxbECQ52rtl7MgVWhhDfS1WYYL1gY3Yr6OR4Ss/edit#gid=0"; //売上管理表URLまとめシートURLを変数に格納

      const totalCell = "A1:L90"; //シートの読み取り範囲全体のセル範囲指定
      const profitSheetTotalCell = "A1:AZ600"; //売上管理表で使用しているセル範囲を指定

      const cutTotalCell = totalCell.split(""); //全体のセル範囲文字列を１文字ずつ配列格納
      this.lastColumn = cutTotalCell[3]; //セルのcolumの最も使われている範囲のアルファベットのみ抜き出す

      const cutProfitTotalCell = profitSheetTotalCell.split("");
      this.profitSheetLastColumn =
        cutProfitTotalCell[3] + cutProfitTotalCell[4]; //売上管理表columの最大で使われている範囲のアルファベットのみ抜き出す

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
        plannerName: "K5", //担当プランナー名記載のセルを設定

        //サンプル関連
        sampleCheck: "K14",
        //確定関連
        confirmCheck:"L3",

        //写真商品関連
        photoItem: "F17", //挙式当日写真商品記載のセルを設定
        photoItemPrice: "K17", //挙式当日写真商品価格記載のセルを設定
        photoItemPlan: "L17", //プラン内の有無記載のセルを設定

        photoOption1: "F19", //挙式当日写真オプション１記載のセルを設定
        photoOption1Price: "K19", //挙式当日写真オプション１価格記載のセルを設定

        photoOption2: "F21", //挙式当日写真オプション２記載のセルを設定
        photoOption2Price: "K21", //挙式当日写真オプション2価格記載のセルを設定

        photographer: "F25", //指名カメラマン記載のセルを設定
        photographerPrice: "K25", //指名フォトグラファー価格記載のセルを設定

        //フォーマル商品関係
        fmItem: "A17", //フォーマル商品記載のセルを設定
        fmItemColor: "D17", //フォーマル商品台紙色記載のセルを設定
        fmPrice: "E17", //フォーマル価格記載のセルを設定

        fmOptionItem1: "A19", //フォーマルオプション記載のセルを設定
        fmOptionColor1: "B19", //フォーマルオプション色記載のセルを設定
        fmOptionPorse1: "C19", //フォーマルオプションポーズ記載のセルを設定
        fmOptionNumber1: "D19", //フォーマルオプション冊数記載のセルを設定
        fmOptionPrice1: "E19", //フォーマルオプション価格記載のセルを設定

        fmOptionItem2: "A21", //フォーマルオプション記載のセルを設定
        fmOptionColor2: "B21", //フォーマルオプション色記載のセルを設定
        fmOptionPorse2: "C21", //フォーマルオプションポーズ記載のセルを設定
        fmOptionNumber2: "D21", //フォーマルオプション冊数記載のセルを設定
        fmOptionPrice2: "E21", //フォーマルオプション価格記載のセルを設定

        fmOptionItem3: "A25", //フォーマルオプション記載のセルを設定
        fmOptionColor3: "B25", //フォーマルオプション色記載のセルを設定
        fmOptionPorse3: "C25", //フォーマルオプションポーズ記載のセルを設定
        fmOptionNumber3: "D25", //フォーマルオプション冊数記載のセルを設定
        fmOptionPrice3: "E25", //フォーマルオプション価格記載のセルを設定

        //前撮り商品関係
        prephotoItem1: "A28", //前撮り商品記載のセルを設定
        prephotoItem2: "A30", //前撮り商品記載のセルを設定
        prephotoItem3: "A34", //前撮り商品記載のセルを設定

        prephotoItem1Price: "E28", //前撮り商品価格記載のセルを設定
        prephotoItem2Price: "E30", //前撮り商品価格記載のセルを設定
        prephotoItem3Price: "E34", //前撮り商品価格記載のセルを設定

        prephotoOption1: "A37", //前撮り商品オプション記載のセルを設定
        prephotoOption2: "A38", //前撮り商品オプション記載のセルを設定
        prephotoOption3: "A39", //前撮り商品オプション記載のセルを設定

        prephotoOption1Price: "E37", //前撮り商品オプション価格記載のセルを設定
        prephotoOption2Price: "E38", //前撮り商品オプション価格記載のセルを設定
        prephotoOption3Price: "E39", //前撮り商品オプション価格記載のセルを設定

        //住所関連
        zipAdd: "A12", //郵便番号記載セルを設定
        Adress1: "A13", //住所記載セルを設定
        Adress2: "E13", //住所記載セルを設定

        //VTRシート記載項目-------------------------------------------

        //vtr商品関係
        vtrSetItem: "A3", //VTRセット商品記載セルを設定
        vtrSetItemPrice: "D3", //VTRセット商品価格記載のセルを設定
        vtrSetItemPricePlan: "E3", //VTRセット商品プラン内判定セルを設定
        vtrSetItemBgm: "A5", //VTRセット商品BGM記載のセルを設定
        vtrSetItemBgmPrice: "E5", //vtrセット商品BGM価格記載のセルを設定

        vtrRecItem: "A8", //VTR記録商品記載のセルを設定
        vtrRecItemPrice: "D8", //VTR記録商品価格記載のセルを設定
        vtrRecItemPricePlan: "E8", //VTR記録商品プラン内判定セルを設定

        vtrEndItem: "A11", //VTRエンド商品記載のセルを設定
        vtrEndItemPrice: "D11", //VTRエンド商品価格記載のセルを設定
        vtrEndItemPricePlan: "E11", //VTRエンド商品プラン判定記載セルを設定
        vtrEndItemBgm: "A13", //VTRエンド商品BGM記載セルを設定
        vtrEndItemBgmPrice: "E13", //VTRエンド商品BGM価格記載セルを設定

        vtrProfileItem: "A16", //VTRプロフィール商品記載セルを設定
        vtrProfileItemPrice: "D16", //VTRプロフィール商品価格記載セルを設定
        vtrProfileItemBgm: "A18", //VTRプロフィール商品BGM記載セルを設定
        vtrProfileItemBgmPrice: "D18", //VTRプロフィール商品BGM価格記載セルを設定

        vtrOption1: "A22", //VTRオプション商品記載のセルを設定
        vtrOption1Price: "D22", //VTRオプション商品価格記載のセルを設定
        vtrOption2: "A23", //VTRオプション商品記載のセルを設定
        vtrOption2Price: "D23", //VTRオプション商品価格記載のセルを設定
        vtrOption3: "A24", //VTRオプション商品記載のセルを設定
        vtrOption3Price: "A24", //VTRオプション商品価格記載のセルを設定
        vtrTelop: "A26", //VTRテロップ記載のセルを設定
        vtrTelopPrice: "D26", //VTRテロップ価格記載のセルを設定

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
        checkVtrMail: "E6", //VTRメール可否チェック欄を設定
        checkProfileMail: "E7", //プロフィールメールチェック欄を設定
        checkSchedule: "E8", //施工管理表チェック欄を設定

        activeSheetUrl: "A22", //共有URLが記載するセルを設定

        vtrMailTextOption: "D26", //VTRにメールを送る際の備考欄を設定
        profileMailTextOption: "D35", //プロフィールにメールを送る際の備考欄を設定

        firstPhotoItem: "A78",//初期見積もり写真商品記入セルを設定
        firstPhotoItemPrice: "B78",//初期見積もり写真商品価格記入セルを設定
        
        firstFmItem:"A80",//初期見積もりfm商品記入セルを設定
        firstFmItemPrice:"B80",//初期見積もりfm商品価格を設定

        firstEndVtrItem: "A82",//初期見積もりエンドロール記入セルを設定
        firstEndVtrItemPrice: "B82",//初期見積もりエンドロール価格セルを設定
        
        firstRecVtrItem:"A84",//初期見積もり記録商品記入セルを設定
        firstRecVtrItemPrice:"B84",//初期見積もり記録商品価格記入セルを設定

        firstSetVtrItem:"A86",//初期見積もりセットVTR商品記入セルを設定
        firstSetVtrItem:"B86"//初期見積もりセットVTR価格記入セルを設定
      };

      //施工管理表に記載する際の配列まとめ
      this.timeCellArray = ["A5", "A8", "A11", "A14", "A17", "A20"]; //施工管理表に挙式時間を記載する際のセルを挙式時間ごとに配列にまとめたもの
      this.nameCellArray = ["C5", "C8", "C11", "C14", "C17", "C20"]; //施工管理表に両家名を記載する際のセルを挙式時間ごとに配列にまとめたもの
      this.placeCellArray = ["C6", "C9", "C12", "C15", "C18", "C21"]; //施工管理表に披露宴会場を記載する際のセルを挙式時間ごとに配列にまとめたもの
      this.photoItemCellArray = ["C7", "C10", "C13", "C16", "C19", "C22"]; //施工管理表に写真商品を記載する際のセルを挙式時間ごとに配列にまとめたもの
      this.fmItemCellArray = ["G5", "G8", "G11", "G14", "G17", "G20"]; //施工管理表にフォーマル商品を記載する際のセルを挙式時間ごとに配列にまとめたもの
      this.vtrItemCellArray = ["G7", "G10", "G13", "G16", "G19", "G22"]; //施工管理表にVTR商品を記載する際のセルを挙式時間ごとに配列にまとめたもの
      this.personCellArray = ["J5", "J8", "J11", "J14", "J17", "J20"]; //施工管理表に指名カメラマンを記載する際のセルを挙式時間ごとに配列にまとめたもの

      //売上管理表の項目ごとの列番号配列
      this.profitSheetItemGet = {
        //挙式当日商品共通
        受注日: "B",
        施行日: "C",
        挙式場: "D",
        披露宴会場: "E",
        新郎新婦名: "F",
        サンプル: "G",
        プランナー: "H",
        BP担当: "I",

        //当日写真シート
        初期見積もり商品: "J",
        初期見積もり上代: "K",
        初期見積もり下代: "L",

        打ち合わせ時当日撮影商品: "M",
        確定時当日撮影商品: "N",
        打ち合わせ撮影商品上代: "O",
        打ち合わせ撮影商品下代: "P",
        確定撮影商品上代: "Q",
        確定撮影商品下代: "R",
        撮影商品変動率: "S",

        打ち合わせ時オプション1商品: "T",
        確定オプション1商品: "U",
        オプション1打ち合わせ上代: "V",
        オプション1打ち合わせ下代: "W",
        オプション1確定上代: "X",
        オプション1確定下代: "Y",

        打ち合わせ時オプション2商品: "Z",
        確定オプション2商品: "AA",
        オプション2打ち合わせ上代: "AB",
        オプション2打ち合わせ下代: "AC",
        オプション2確定上代: "AD",
        オプション2確定下代: "AE",

        打ち合わせ時指名商品: "AF",
        確定指名商品: "AG",
        指名打ち合わせ上代: "AH",
        指名打ち合わせ下代: "AI",
        指名確定上代: "AJ",
        指名確定下代: "AK",

        打ち合わせ時オプション総額上代: "AL",
        打ち合わせオプション総額下代: "AM",
        確定オプション総額上代: "AN",
        確定オプション総額下代: "AO",
        オプション変動率: "AP",

        写真打ち合わせ総額上代: "AQ",
        写真打ち合わせ総額下代: "AR",
        写真確定総額上代: "AS",
        写真確定総額下代: "AT",
        写真総額変動率: "AU",

        //FMシート
        FM初期見積もり商品: "J",
        FM初期見積もり上代: "K",
        FM初期見積もり下代: "L",

        FM打ち合わせ時当日撮影商品: "M",
        FM確定時当日撮影商品: "N",
        FM打ち合わせ撮影商品上代: "O",
        FM打ち合わせ撮影商品下代: "P",
        FM確定撮影商品上代: "Q",
        FM確定撮影商品下代: "R",
        FM撮影商品変動率: "S",

        FM打ち合わせ時オプション1商品: "T",
        FM確定オプション1商品: "U",
        FMオプション1打ち合わせ上代: "V",
        FMオプション1打ち合わせ下代: "W",
        FMオプション1確定上代: "X",
        FMオプション1確定下代: "Y",

        FM打ち合わせ時オプション2商品: "Z",
        FM確定オプション2商品: "AA",
        FMオプション2打ち合わせ上代: "AB",
        FMオプション2打ち合わせ下代: "AC",
        FMオプション2確定上代: "AD",
        FMオプション2確定下代: "AE",

        FM打ち合わせ時オプション3商品: "AF",
        FM確定オプション3商品: "AG",
        FMオプション3打ち合わせ上代: "AH",
        FMオプション3打ち合わせ下代: "AI",
        FMオプション3確定上代: "AJ",
        FMオプション3確定下代: "AK",

        FM打ち合わせ時オプション総額上代: "AL",
        FM打ち合わせオプション総額下代: "AM",
        FM確定オプション総額上代: "AN",
        FM確定オプション総額下代: "AO",
        FMオプション変動率: "AP",

        FM打ち合わせ総額上代: "AQ",
        FM打ち合わせ総額下代: "AR",
        FM確定総額上代: "AS",
        FM確定総額下代: "AT",
        FM総額変動率: "AU",

        //VTRシート
        VTR初期見積もり商品: "J",
        VTR初期見積もり上代: "K",
        VTR初期見積もり下代: "L",

        VTRエンド打ち合わせ時商品: "M",
        VTRエンド確定時商品: "N",
        VTRエンド打ち合わせ上代: "O",
        VTRエンド打ち合わせ下代: "P",
        VTRエンド確定上代: "Q",
        VTRエンド確定下代: "R",
        VTRエンド変動率: "S",

        VTR記録打ち合わせ時商品: "T",
        VTR記録確定商品: "U",
        VTR記録打ち合わせ上代: "V",
        VTR記録打ち合わせ下代: "W",
        VTR記録確定上代: "X",
        VTR記録確定下代: "Y",
        VTR記録商品変動率: "Z",

        VTRセット打ち合わせ時商品: "AA",
        VTRセット確定商品: "AB",
        VTRセット打ち合わせ上代: "AC",
        VTRセット打ち合わせ下代: "AD",
        VTRセット確定上代: "AE",
        VTRセット確定下代: "AF",
        VTRセット変動率: "AG",

        VTRプロフィール打ち合わせ時商品: "AH",
        VTRプロフィール確定時商品: "AI",
        VTRプロフィール打ち合わせ上代: "AJ",
        VTRプロフィール打ち合わせ下代: "AK",
        VTRプロフィールプロフィール確定上代: "AL",
        VTRプロフィールプロフィール確定下代: "AM",
        VTRプロフィール変動率: "AN",

        VTR打ち合わせ総額上代: "AO",
        VTR打ち合わせ総額下代: "AP",
        VTR確定総額上代: "AQ",
        VTR確定総額下代: "AR",
        VTR変動率: "AS",

        //前撮り共通項目
        前撮りお客様挙式日: "B",
        前撮り撮影日: "C",
        前撮り撮影場所: "D",
        前撮り新郎新婦名: "E",
        前撮りサンプル: "F",
        前撮りプランナー: "G",
        前撮りBP担当: "H",

        前撮り打ち合わせ商品: "I",
        前撮り確定商品: "J",
        前撮り打ち合わせ上代: "K",
        前撮り打ち合わせ下代: "L",
        前撮り確定上代: "M",
        前撮り確定下代: "N",
        前撮り変動率: "O",

        前撮りオプション1打ち合わせ商品: "P",
        前撮りオプション1確定商品: "Q",
        前撮りオプション1打ち合わせ上代: "R",
        前撮りオプション1打ち合わせ下代: "S",
        前撮りオプション1確定上代: "T",
        前撮りオプション1確定下代: "U",

        前撮りオプション2打ち合わせ商品: "V",
        前撮りオプション2確定商品: "W",
        前撮りオプション2打ち合わせ上代: "X",
        前撮りオプション2打ち合わせ下代: "Y",
        前撮りオプション2確定上代: "Z",
        前撮りオプション2確定下代: "AA",

        前撮りオプション3打ち合わせ商品: "AB",
        前撮りオプション3確定商品: "AC",
        前撮りオプション3打ち合わせ上代: "AD",
        前撮りオプション3打ち合わせ下代: "AE",
        前撮りオプション3確定上代: "AF",
        前撮りオプション3確定下代: "AG",

        前撮りオプション打ち合わせ総額上代: "AH",
        前撮りオプション打ち合わせ総額下代: "AI",
        前撮りオプション確定総額上代: "AJ",
        前撮りオプション確定総額下代: "AK",
        前撮りオプション変動率: "AL",

        前撮り打ち合わせ総額上代: "AM",
        前撮り打ち合わせ総額下代: "AN",
        前撮り確定総額上代: "AO",
        前撮り確定総額下代: "AP",
        前撮り変動率: "AQ",

        //担当売上票
        担当初期見積もり総額: "C",
        担当写真商品総額: "D",
        担当FM商品総額: "E",
        担当VTR商品総額: "F",
        担当打ち合わせ件数: "I",
      };

      //以上のセル指定部分を変更すれば簡易的な変更は対応可能（受注書のレイアウト変更など）

      this.aoyamaChapel = "南青山ル・アンジェ教会";

      //---------------------------------------各種処理を実行する際に必要な値を変数に格納------------------------------------

      this.ceremonyDay = this._getInfo("ceremonyDay"); //挙式日取得変数代入
      this.ceremonyTime = this._getInfo("ceremonyTime"); //挙式時間取得変数代入
      this.partyTime = this._getInfo("partyTime"); //披露宴時間取得
      this.acceptDay = this._getInfo("acceptDay"); //受付日取得

      this.confirmCheck = this._getInfo("confirmCheck");//確定判定用

      this.ceremonyDayFormat = this._dayFormat(this.ceremonyDay, "yyyy/MM/dd"); //挙式日程を年、月、日の形式にフォーマットする
      this.ceremonyTimeFormat = this._dayFormat(this.ceremonyTime, "HH:mm"); //挙式時間を時、分の形式にフォーマットする
      this.partyTimeFormat = this._dayFormat(this.partyTime, "HH:mm"); //披露宴時間を時、分の形式にフォーマットする
      this.timeNameFormat = this._dayFormat(this.ceremonyDay, "yyyyMMdd"); //ファイル名として使用する形に時間をフォーマットする
      this.acceptDayFormat = this._dayFormat(this.acceptDay, "yyyy/MM/dd"); //受付日を記入する形にフォーマットする

      this.staffName = this._getInfo("staffName"); //BP担当者名取得変数代入
      this.plannerName = this._getInfo("plannerName"); //プランナー担当者取得変数代入

      this.menName = this._getInfo("menName"); //新郎名前取得変数代入
      this.womenName = this._getInfo("womenName"); //新婦名前取得変数代入
      this.MenWomenName = this.menName + "\t" + this.womenName; //施工管理表シートに記載する形式に新郎新婦名をフォーマットする

      this.partyRoomName = this._getInfo("partyRoomName"); //披露宴会場名取得変数代入

      this.zipAdd = this._getInfo("zipAdd"); //お客様住所郵便番号取得変数代入

      this.sampleCheck = this._getInfo("sampleCheck"); //サンプルOKかどうかを判定する

      this.photoItem = this._getInfo("photoItem"); //写真商品取得変数代入
      this.photoItemPrice = this._getInfo("photoItemPrice");
      this.photoItemPlan = this._getInfo("photoItemPlan");
      this.photoOption1 = this._getInfo("photoOption1");
      this.photoOption1Price = this._getInfo("photoOption1Price");
      this.photoOption2 = this._getInfo("photoOption2");
      this.photoOption2Price = this._getInfo("photoOption2Price");

      this.photographer = this._getInfo("photographer"); //指名カメラマン商品取得変数代入
      this.photographerPrice = this._getInfo("photographerPrice"); //指名カメラマン価格取得変数代入

      //フォーマル商品関係
      this.fmItem = this._getInfo("fmItem");
      this.fmItemColor = this._getInfo("fmItemColor");
      this.fmPrice = this._getInfo("fmPrice");
      //フォーマルオプション
      this.fmOptionItem1 = this._getInfo("fmOptionItem1");
      this.fmOptionColor1 = this._getInfo("fmOptionColor1");
      this.fmOptionPorse1 = this._getInfo("fmOptionPorse1");
      this.fmOptionNumber1 = this._getInfo("fmOptionNumber1");
      this.fmOptionPrice1 = this._getInfo("fmOptionPrice1");

      this.fmOptionItem2 = this._getInfo("fmOptionItem2");
      this.fmOptionColor2 = this._getInfo("fmOptionColor2");
      this.fmOptionPorse2 = this._getInfo("fmOptionPorse2");
      this.fmOptionNumber2 = this._getInfo("fmOptionNumber2");
      this.fmOptionPrice2 = this._getInfo("fmOptionPrice2");

      this.fmOptionItem3 = this._getInfo("fmOptionItem3");
      this.fmOptionColor3 = this._getInfo("fmOptionColor3");
      this.fmOptionPorse3 = this._getInfo("fmOptionPorse3");
      this.fmOptionNumber3 = this._getInfo("fmOptionNumber3");
      this.fmOptionPrice3 = this._getInfo("fmOptionPrice3");

      this.vtrSetItem = this._getInfo("vtrSetItem", this.vsArray); //VTRのセットアイテムを取得変数代入
      this.vtrRecItem = this._getInfo("vtrRecItem", this.vsArray); //記録映像を取得変数代入
      this.vtrEndItem = this._getInfo("vtrEndItem", this.vsArray); //エンドロールを取得変数代入
      this.vtrProfileItem = this._getInfo("vtrProfileItem", this.vsArray); //プロフィールを取得変数代入
      this.vtrOption1 = this._getInfo("vtrOption1", this.vsArray);
      this.vtrOption2 = this._getInfo("vtrOption2", this.vsArray);
      this.vtrTelop = this._getInfo("vtrTelop", this.vsArray);
      

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
      this.vtrMailRecipent = "filmj0222@gmail.com"; //VTR担当者のメールアドレスを変数に格納filmj0222@gmail.com
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

      this.firstPhotoItem = this._getInfo("firstPhotoItem", this.swArray);
      this.firstPhotoItem = this._getInfo("firstPhotoItemPrice", this.swArray);

      this.firstFmItem = this._getInfo("firstFmItem",this.swArray);
      this.firstFmItem = this._getInfo("firstFmItemPrice",this.swArray);

      this.firstEndVtrItem = this._getInfo("firstEndVtrItem", this.swArray);
      this.firstEndVtrItemPrice = this._getInfo("firstEndVtrItemPrice", this.swArray);
      this.firstRecVtrItem = this._getInfo("firstRecVtrItem",this.swArray);
      this.firstRecVtrItemPrice = this._getInfo("firstRecVtrItemPrice",this.swArray);
      this.firstSetVtrItem = this._getInfo("firstSetVtrItem",this.swArray);
      this.firstSetVtrItemPrice = this._getInfo("firstSetVtrItemPrice",this.swArray);
    }
    //--------------------------------------------以上必要な値の変数格納終了-----------------------------

    start() {
      //記入開始時に必ず使用する関数
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

    writeAnotherSheet(change = this.change) {
      //--------------------------------施工管理表に記載するための関数----------------------------------------
      this.vtrItem = this._vtrItemCheck(); //施工管理表に記載するVTR商品のフィルタリング
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

    writeMinamiAoyamaProfitSheet() {//南青山ル・アンジェ教会の売上管理表記載関数
      this._serchScheduleSheet(
        this.ceremonyDay,
        this.profitSheetUrl,
        "売上管理表検索"
      );
      const profitStandardCell = this.profitSheetItemGet.新郎新婦名 + "1";

      const writeProfitRow = this._getLastRow(
        this.profitSheetName.当日写真商品,
        profitStandardCell,
        SpreadsheetApp.openByUrl(this.scheduleUrl)
      );
      const rowStrChange = String(writeProfitRow);

      

      if(this.confirmCheck === "" || !this.confirmCheck){

        const firstVtrItem =  this._vtrItemCheck(this.firstSetVtrItem,this.firstRecVtrItem,this.firstEndVtrItem);

         if(firstVtrItem===this.firstSetVtrItem){
          var firstVtrItemPrice = this.firstSetVtrItemPrice
        }else if(firstVtrItem === this.firstRecVtrItem){
          var firstVtrItemPrice = this.firstRecVtrItemPrice
        }else{
          var firstVtrItemPrice = this.firstEndVtrItemPrice
        }


      this._cellWrite(
        this.profitSheetItemGet.受注日 + rowStrChange,
        this.acceptDay,
        this.scheduleUrl,
        this.profitSheetName.当日写真商品
      );
      this._cellWrite(
        this.profitSheetItemGet.施行日 + rowStrChange,
        this.ceremonyDayFormat,
        this.scheduleUrl,
        this.profitSheetName.当日写真商品
      );
      this._cellWrite(
        this.profitSheetItemGet.挙式場 + rowStrChange,
        this.aoyamaChapel,
        this.scheduleUrl,
        this.profitSheetName.当日写真商品
      );
      this._cellWrite(
        this.profitSheetItemGet.披露宴会場 + rowStrChange,
        this.partyRoomName,
        this.scheduleUrl,
        this.profitSheetName.当日写真商品
      );
      this._cellWrite(
        this.profitSheetItemGet.新郎新婦名 + rowStrChange,
        this.MenWomenName,
        this.scheduleUrl,
        this.profitSheetName.当日写真商品
      );
      this._cellWrite(
        this.profitSheetItemGet.サンプル + rowStrChange,
        this.sampleCheck,
        this.scheduleUrl,
        this.profitSheetName.当日写真商品
      );
      this._cellWrite(
        this.profitSheetItemGet.プランナー + rowStrChange,
        this.plannerName,
        this.scheduleUrl,
        this.profitSheetName.当日写真商品
      );
      this._cellWrite(
        this.profitSheetItemGet.BP担当 + rowStrChange,
        this.staffName,
        this.scheduleUrl,
        this.profitSheetName.当日写真商品
      );
      this._cellWrite(
        this.profitSheetItemGet.初期見積もり商品 + rowStrChange,
        this.firstPhotoItem,
        this.scheduleUrl,
        this.profitSheetName.当日写真商品
      );
      this._cellWrite(
        this.profitSheetItemGet.打ち合わせ時当日撮影商品 + rowStrChange,
        this.photoItem,
        this.scheduleUrl,
        this.profitSheetName.当日写真商品
      );
      this._cellWrite(
        this.profitSheetItemGet.打ち合わせ時オプション1商品 + rowStrChange,
        this.photoOption1,
        this.scheduleUrl,
        this.profitSheetName.当日写真商品
      );
      this._cellWrite(
        this.profitSheetItemGet.打ち合わせ時オプション2商品 + rowStrChange,
        this.photoOption2,
        this.scheduleUrl,
        this.profitSheetName.当日写真商品
      );
      this._cellWrite(
        this.profitSheetItemGet.打ち合わせ時指名商品 + rowStrChange,
        this.photographer,
        this.scheduleUrl,
        this.profitSheetName.当日写真商品
      );
      this._cellWrite(
        this.profitSheetItemGet.FM初期見積もり商品 + rowStrChange,
        this.firstFmItem,
        this.scheduleUrl,
        this.profitSheetName.当日フォーマル商品
      );
      this._cellWrite(
        this.profitSheetItemGet.FM打ち合わせ時当日撮影商品 + rowStrChange,
        this.fmItem,
        this.scheduleUrl,
        this.profitSheetName.当日フォーマル商品
      );
      this._cellWrite(
        this.profitSheetItemGet.FM打ち合わせ時オプション1商品 + rowStrChange,
        this.fmOptionItem1+this.fmOptionColor1+this.fmOptionPorse1+this.fmOptionNumber1,
        this.scheduleUrl,
        this.profitSheetName.当日フォーマル商品
      );
      this._cellWrite(
        this.profitSheetItemGet.FMオプション1打ち合わせ上代 + rowStrChange,
        this.fmOptionPrice1,
        this.scheduleUrl,
        this.profitSheetName.当日フォーマル商品
      );
      this._cellWrite(
        this.profitSheetItemGet.FM打ち合わせ時オプション2商品 + rowStrChange,
        this.fmOptionItem2+this.fmOptionColor2+this.fmOptionPorse2+this.fmOptionNumber2,
        this.scheduleUrl,
        this.profitSheetName.当日フォーマル商品
      );
      this._cellWrite(
        this.profitSheetItemGet.FMオプション2打ち合わせ上代 + rowStrChange,
        this.fmOptionPrice2,
        this.scheduleUrl,
        this.profitSheetName.当日フォーマル商品
      );
      this._cellWrite(
        this.profitSheetItemGet.FM打ち合わせ時オプション3商品 + rowStrChange,
        this.fmOptionItem3+this.fmOptionColor3+this.fmOptionPorse3+this.fmOptionNumber3,
        this.scheduleUrl,
        this.profitSheetName.当日フォーマル商品
      );
      this._cellWrite(
        this.profitSheetItemGet.FMオプション3打ち合わせ上代 + rowStrChange,
        this.fmOptionPrice3,
        this.scheduleUrl,
        this.profitSheetName.当日フォーマル商品
      );
      this._cellWrite(
        this.profitSheetItemGet.VTR初期見積もり商品 + rowStrChange,
        firstVtrItem,
        this.scheduleUrl,
        this.profitSheetName.当日VTR商品
      );
      this._cellWrite(
        this.profitSheetItemGet.VTR初期見積もり上代 + rowStrChange,
        firstVtrItemPrice,
        this.scheduleUrl,
        this.profitSheetName.当日VTR商品
      );
      this._cellWrite(
        this.profitSheetItemGet.VTRエンド打ち合わせ時商品 + rowStrChange,
        this.vtrEndItem,
        this.scheduleUrl,
        this.profitSheetName.当日VTR商品
      );
      this._cellWrite(
        this.profitSheetItemGet.VTR記録打ち合わせ時商品 + rowStrChange,
        this.vtrRecItem,
        this.scheduleUrl,
        this.profitSheetName.当日VTR商品
      );
      this._cellWrite(
        this.profitSheetItemGet.VTRセット打ち合わせ時商品 + rowStrChange,
        this.vtrSetItem,
        this.scheduleUrl,
        this.profitSheetName.当日VTR商品
      );
      this._cellWrite(
        this.profitSheetItemGet.VTRプロフィール打ち合わせ時商品 + rowStrChange,
        this.vtrProfileItem,
        this.scheduleUrl,
        this.profitSheetName.当日VTR商品
      );
    

      }else if(this.confirmCheck != ""){

        const findRow = String(this._findRow());


        this._cellWrite(
          this.profitSheetItemGet.確定時当日撮影商品 + findRow,
          this.photoItem,
          this.scheduleUrl,
          this.profitSheetName.当日写真商品
        );
        this._cellWrite(
          this.profitSheetItemGet.確定オプション1商品 + findRow,
          this.photoOption1,
          this.scheduleUrl,
          this.profitSheetName.当日写真商品
        );
        this._cellWrite(
          this.profitSheetItemGet.確定オプション2商品 + findRow,
          this.photoOption2,
          this.scheduleUrl,
          this.profitSheetName.当日写真商品
        );
        this._cellWrite(
          this.profitSheetItemGet.確定指名商品 + findRow,
          this.photographer,
          this.scheduleUrl,
          this.profitSheetName.当日写真商品
        );
        this._cellWrite(
          this.profitSheetItemGet.FM確定時当日撮影商品 + findRow,
          this.fmItem,
          this.scheduleUrl,
          this.profitSheetName.当日フォーマル商品
        );
        this._cellWrite(
          this.profitSheetItemGet.FM確定オプション1商品 + findRow,
          this.fmOptionItem1+this.fmOptionColor1+this.fmOptionPorse1+this.fmOptionNumber1,
          this.scheduleUrl,
          this.profitSheetName.当日フォーマル商品
        );
        this._cellWrite(
          this.profitSheetItemGet.FMオプション1確定上代 + findRow,
          this.fmOptionPrice1,
          this.scheduleUrl,
          this.profitSheetName.当日フォーマル商品
        );
        this._cellWrite(
          this.profitSheetItemGet.FM確定オプション2商品 + findRow,
          this.fmOptionItem2+this.fmOptionColor2+this.fmOptionPorse2+this.fmOptionNumber2,
          this.scheduleUrl,
          this.profitSheetName.当日フォーマル商品
        );
        this._cellWrite(
          this.profitSheetItemGet.FMオプション2確定上代 + findRow,
          this.fmOptionPrice2,
          this.scheduleUrl,
          this.profitSheetName.当日フォーマル商品
        );
        this._cellWrite(
          this.profitSheetItemGet.FM確定オプション3商品 + findRow,
          this.fmOptionItem3+this.fmOptionColor3+this.fmOptionPorse3+this.fmOptionNumber3,
          this.scheduleUrl,
          this.profitSheetName.当日フォーマル商品
        );
        this._cellWrite(
          this.profitSheetItemGet.FMオプション3確定上代 + findRow,
          this.fmOptionPrice3,
          this.scheduleUrl,
          this.profitSheetName.当日フォーマル商品
        );
        this._cellWrite(
          this.profitSheetItemGet.VTRエンド確定時商品 + findRow,
          this.vtrEndItem,
          this.scheduleUrl,
          this.profitSheetName.当日VTR商品
        );
        this._cellWrite(
          this.profitSheetItemGet.VTR記録確定商品 + findRow,
          this.vtrRecItem,
          this.scheduleUrl,
          this.profitSheetName.当日VTR商品
        );
        this._cellWrite(
          this.profitSheetItemGet.VTRセット確定商品 + findRow,
          this.vtrSetItem,
          this.scheduleUrl,
          this.profitSheetName.当日VTR商品
        );
        this._cellWrite(
          this.profitSheetItemGet.VTRプロフィール確定時商品 + findRow,
          this.vtrProfileItem,
          this.scheduleUrl,
          this.profitSheetName.当日VTR商品
        );

      }else{
        msgBox("エラーのため処理を中止します");
        return
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

        this._cellWriteActive(
          this.ItemGetCellArray.checkProfileMail,
          "OK",
          this.sw
        );

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

        this._cellWriteActive(
          this.ItemGetCellArray.checkVtrMail,
          "OK",
          this.sw
        );

        Browser.msgBox(
          "陣さんへのメールが送信されました。Gmailを確認してください"
        );
      } else {
        Browser.msgBox("共有URLが記載されていません。確認してください");
        return;
      }
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

    _getArrayNumber(cell, lastColumn = this.lastColumn) {
      //------------------スプレッドシートすべて取りこんだ配列よりセル番号で情報を引き出すための変換関数

      const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
      const alphabetArray = alphabet.split("");
      const Numbers = "1234567890";
      const NumbersArray = Numbers.split("");
      const strCell = cell.split("");
      const lastColumnArray = lastColumn.split("");

      if (
        alphabetArray.includes(strCell[0]) &&
        NumbersArray.includes(strCell[1]) &&
        lastColumnArray.length === 1
      ) {
        //読み取り指定のセルが英字１桁それ以降数字構成（例A12,B4など）で、最終列が英字一文字のとき（例Jなど）
        var strCellRow = Number(cell.slice(1));
        var totalCoNumber = alphabet.indexOf(strCell[0]);

        var totalLastColumnNumber = alphabet.indexOf(lastColumnArray[0]) + 1;
      } else if (
        alphabetArray.includes(strCell[0]) &&
        alphabetArray.includes(strCell[1]) &&
        lastColumnArray.length === 2
      ) {
        //読み取り指定のセルが英字２数字１から２で構成（例AA12,AB4など）で、最終列が英字2文字のとき（例ABなど）
        var strCellRow = Number(cell.slice(2));
        const coNumber = alphabet.indexOf(strCell[0]) + 1;
        const coNumber2 = alphabet.indexOf(strCell[1]);
        var totalCoNumber = coNumber * 26 + coNumber2;

        var lastColumnNumber = alphabet.indexOf(lastColumnArray[0]) + 1;
        var lastColumnNumber2 = alphabet.indexOf(lastColumnArray[1]) + 1;
        var totalLastColumnNumber = lastColumnNumber * 26 + lastColumnNumber2;
      } else if (
        alphabetArray.includes(strCell[0]) &&
        NumbersArray.includes(strCell[1]) &&
        lastColumnArray.length === 2
      ) {
        //読み取り指定のセルが英字1数字１以上で構成（例AA12,AB4など）で、最終列が英字2文字のとき（例ABなど）
        var strCellRow = Number(cell.slice(1));
        var totalCoNumber = alphabet.indexOf(strCell[0]);

        var lastColumnNumber = alphabet.indexOf(lastColumnArray[0]) + 1;
        var lastColumnNumber2 = alphabet.indexOf(lastColumnArray[1]) + 1;
        var totalLastColumnNumber = lastColumnNumber * 26 + lastColumnNumber2;
      }
      return (strCellRow - 1) * totalLastColumnNumber + totalCoNumber;
    }

    _getInfo(key1, array = this.psArray) {
      //----------------配列よりスプレッドシートのセル番号から情報を取り出す関数------------------
      return array[this._getArrayNumber(this.ItemGetCellArray[key1])];
    }

    _serchScheduleSheet(
      day = this.ceremonyDay,
      url = this.manageSheetUrl,
      sheetname = "施工管理表検索"
    ) {
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
      //----------------------スプレッドシートから読み取るシートとセルを選択する関数---------------------------
      return SpreadsheetApp.openByUrl(url)
        .getSheetByName(sheet)
        .getRange(range)
        .getValues();
    }

    _getRowSerch(url, sn, cell, target) {
      //-------------------------配列より検索用語に一致する要素のみ配列に格納する------------------------
      const namearray = this._setCellUrl(url, sn, cell);
      const rowNumberArray = namearray.filter(target);
      return rowNumberArray;
    }

    _serchIndex(array, key) {
      //---------------------------配列内のkeyと一致した要素の配列番号を新しい配列に格納する。---------------------
      const indexArray = [];
      array.forEach(function (el, index) {
        if (el === key) {
          indexArray.push(index + 1);
        }
      });
      return indexArray;
    }

    _totalValueStaff(url, sheet, RowArray, column) {
      //----------------------------列番号の配列を利用して利用して同じ行の異なった列のデータを配列に格納する。--------------
      const valueArray = [];
      RowArray.forEach(function (el) {
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

    _profitSheetPersonTotal(
      profitDay,
      targetProfitSheetName,
      personKey,
      ItemGetRow
    ) {
      //売上管理表の担当者ごとの売上合計を求める関数
      const profitRange =
        this.profitSheetItemGet.BP担当 +
        "4" +
        ":" +
        this.profitSheetItemGet.BP担当 +
        "800";

      const profitSheetUrl = this._serchScheduleSheet(
        profitDay,
        this.profitSheetUrl,
        "シート1"
      ); //対象の売上管理表を取得

      const PersonCellArray = this._setCellUrl(
        profitSheetUrl,
        targetProfitSheetName,
        profitRange
      ); //売上管理表の担当者列からデータを取得

      const photoPersonRowArray = this._serchIndex(PersonCellArray, personKey); //特定の担当者のお客様情報が記載された列番号を配列にする

      //該当担当者が打ち合わせしたお客様の各価格を配列に格納する
      const FirstPriceArray = this._totalValueStaff(
        profitSheetUrl,
        targetProfitSheetName,
        photoPersonRowArray,
        ItemGetRow
      );

      //配列に格納したお客様の価格を合計する
      return this._arrayTotal(FirstPriceArray);
    }

    _findRow(url,sheet,range,val,col){

      var dat = SpreadsheetApp.openByUrl(url).getSheetByName(sheet).getDataRange(range).getValues(); //受け取ったシートのデータを二次元配列に取得
    
      for(var i=1;i<dat.length;i++){
        if(dat[i][col-1] === val){
          return i+1;
        }
      }
      return 0;
    }


  }

  //--------------------以上オブジェクト---------------------------------------------
  const createMethods = function () {
    const methods = new ControlItem();
    return methods;
  };
  return createMethods();
}
