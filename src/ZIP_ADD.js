function ZIP_ADD(zip) {
  const response = UrlFetchApp.fetch(
    "http://zipcloud.ibsnet.co.jp/api/search?zipcode=" + zip
  );
  const results = JSON.parse(response.getContentText()).results;
  return results[0].address1 + results[0].address2 + results[0].address3;
}
