/**
* Send a newsletter
* @param none
* @return none
*/
function sendNewsLetter(){
  const inputSheetName = 'ニュースレター送信';
  const fileIdAddress = 'C1';
  const sendMailOptionsRange = 'A2:B4';
  const testToAddress = 'B5';
  const mainToAddress = 'B6';
  const resString = 'ok';
  const ui = SpreadsheetApp.getUi();
  const inputSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(inputSheetName);
  const fileId = inputSheet.getRange(fileIdAddress).getValue();
  const htmlString = getHtmlFromFile(fileId);
  let sendMailInfo  = Object.fromEntries(inputSheet.getRange(sendMailOptionsRange).getValues());  
  sendMailInfo.noReply = true;
  sendMailInfo.to = inputSheet.getRange(testToAddress).getValue();
  sendMailInfo.htmlBody = htmlString;
  /** send test */
  const saveSubject = sendMailInfo.subject; 
  sendMailInfo.subject = '（テスト送信）' + sendMailInfo.subject; 
  let resSendMail = sendMail(sendMailInfo);
  if (!resSendMail){
    return;
  }
  sendMailInfo.subject = saveSubject;
  /** production */
  const res = SpreadsheetApp.getUi().prompt('本番送信する場合は半角小文字で"' + resString + '"と入力し、OKをクリックしてください。それ以外の操作をすると処理を終了します。', ui.ButtonSet.OK_CANCEL);
  if (res.getResponseText() == resString && res.getSelectedButton() == ui.Button.OK){
    sendMailInfo.to = inputSheet.getRange(mainToAddress).getValue();
    sendMail(sendMailInfo);
  } else {
    ui.alert('送信をキャンセルしました。');
  }
}
/**
* Get the HTML source from a file in Google Drive
* @param {string} File ID of the HTML file
* @return {string} The HTML source
*/
function getHtmlFromFile(fileId){
  const htmlFile = DriveApp.getFileById(fileId).getBlob();
  const htmlContent = HtmlService.createHtmlOutput(htmlFile).getContent();
  return htmlContent;
}
/**
* Send email
* @param {Object} Information such as address and title
* @return {boolean} OK: true, CANCEL: false 
*/
function sendMail(sendMailInfo){
  const ui = SpreadsheetApp.getUi();
  const alertInfo = editAlertInfoStrings(sendMailInfo);
  if (!alertInfo){
    return;
  }
  const res = ui.alert(alertInfo, ui.ButtonSet.OK_CANCEL);
  if (res == ui.Button.OK){
    MailApp.sendEmail(sendMailInfo);
    ui.alert('メールを送信しました。');
    return true;
  } else {
    ui.alert('送信をキャンセルしました。');
    return false;
  }
}
/**
* Create a pop-up menu for confirmation before sending an email
* @param {Object} Information such as address and title
* @return {string} Pop-up menu contents
*/
function editAlertInfoStrings(sendMailInfo){
  var errorCheck = Object.assign({}, sendMailInfo);
  var res = '';
  if (errorCheck.htmlBody){
    delete errorCheck.htmlBody;
  } else {
    SpreadsheetApp.getUi().alert('error:htmlBodyが空白\n送信をキャンセルします。');
    return null;
  };
  if (errorCheck.body){
    delete errorCheck.body;
  }
  Object.keys(errorCheck).forEach(function(key, idx, array){
    res = res + key + ':' + sendMailInfo[key];
    if (idx != array.length){
      res = res + '\n';
    }
  });
  res = 'OKをクリックするとメールが送信されます。キャンセルをクリックすると送信キャンセルします。\n\n' + res;
  return res;
}
/**
* Processing at file open
* @param none
* @return none
*/
function onOpen(){
  SpreadsheetApp.getActiveSpreadsheet().addMenu('メールマガジン送信', [{name: '送信', functionName:'sendNewsLetter'}]);
}
