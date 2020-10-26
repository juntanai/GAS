function delivary_day(getday) {
  const ps = "photo"; //写真商品のシート名に使用する変数を定義

  var row_day = getday;

  const day = Utilities.formatDate(row_day, "JST", "d");
  row_day.setMonth(row_day.getMonth() + 2);
  const month = Utilities.formatDate(row_day, "JST", "M");

  const month2 = "${month}".replace("${month}", month);

  if (day <= 10) {
    return month2 + "月" + "上旬";
  } else if (day <= 20) {
    return month2 + "月" + "中旬";
  } else if (day <= 31) {
    return month2 + "月" + "下旬";
  } else {
    Browser.msgBox("エラーです");
  }
}
