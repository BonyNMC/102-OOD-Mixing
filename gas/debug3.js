function testAugmentirAPI() {
  const options = {
    method: "POST",
    payload: JSON.stringify({
      procedureName: "10000133467-04 Bảng Kiểm Máy Móc Hàng Ca Khu Vực Mixing",
      startDate: "2026-05-15T00:00:00.000Z",
      endDate: "",
      status: "",
      excludeArchived: true,
      limit: 10,
      pagingupdate: true,
    }),
    followRedirects: true,
    muteHttpExceptions: true,
    headers: { "X-aug-api-key": "69091b790253df048876a1e8", "Content-Type": "application/json" }
  };
  
  var res = UrlFetchApp.fetch('https://app.augmentir.com/rest/v1/GetJobsFromProcedure', options);
  Logger.log(res.getResponseCode());
  Logger.log(res.getContentText());
}
