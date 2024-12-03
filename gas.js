function sendSpreadsheetAsExcel(){
  /* Google GAS sample code
   * Send google spreadsheet in xlsx format
   * as an attached file using multipart/form-data in Webex Teams
   */

  // Webex related
  const webexToken = "<bot token>"; 
  const webexRoom = "<roomId>"; 
  const webexUrl = 'https://webexapis.com/v1/messages/';

  // google spreadsheet id, just pick out /d/<this value>/edit? from the URL
  const spreadsheetId = 'abcabcabcabc';

  // URL to get spreadsheet as xlsx
  const fetchUrl = "https://docs.google.com/feeds/download/spreadsheets/Export?key="+
    spreadsheetId+"&amp;exportFormat=xlsx";

  // OAuth2 related
  const fetchOpt = {
    "headers": {Authorization: "Bearer " + ScriptApp.getOAuthToken()},
    "muteHttpExceptions": true
  };

  // Obtain xlsx as blob
  const xlsxFile = UrlFetchApp.fetch(fetchUrl, fetchOpt).getBlob();

  // bundary and header for multipart/form-data
  const boundary = "myBoundary";

  const headers = {
    'Authorization': 'Bearer ' + webexToken,
    'Content-Type': "multipart/form-data; boundary=" + boundary
  };
  
  // Create multipart/form-data
  const payload = Utilities.newBlob(
    '--'+boundary+'\r\nContent-Disposition: form-data; name="roomId"\r\n\r\n'+webexRoom+'\r\n' +
    '--'+boundary+'\r\nContent-Disposition: form-data; name="text"\r\n\r\nファイルです！\r\n' +
    '--'+boundary+'\r\nContent-Disposition: form-data; name="files"; filename="myexcel.xlsx"\r\n'+  
        'Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n').getBytes()
    .concat(xlsxFile.getBytes())
    .concat(Utilities.newBlob('\r\n\r\n--'+boundary+'--\r\n').getBytes()
  );

  const options = {
    'method': 'POST',
    'muteHttpExceptions': true,
    'headers': headers,
    'payload': payload
  };

  res = UrlFetchApp.fetch(webexUrl, options);
  console.log(res.toString());
}
