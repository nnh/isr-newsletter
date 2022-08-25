/**
* Send a newsletter
* @param none
* @return none
*/
function sendNewsLetter(){
  const commonSettings = new CommonSettings();
  if (!commonSettings.sheetStatus){
    return;
  };
  const inputSheet = commonSettings.sheets.sendNewsLetter;
  const resString = 'ok';
  const ui = SpreadsheetApp.getUi();
  const fileIdAddress = 'C1';
  const keyIdx = 0;
  const valueIdx = 1;
  const inputValueColNumber = 2;
  let addressRow = {};
  const options = 'options';
  addressRow[options] = {};
  addressRow.subject = 2;
  addressRow[options].name = 3;
  addressRow.body = 4;
  addressRow.testTo = 5;
  addressRow.mainTo = 6;
  let sendMailInfo = {};
  sendMailInfo[options] = {};
  Object.entries(addressRow).forEach(x => {
    const keyValue = x[keyIdx] !== options ? x : Object.entries(x[valueIdx])[0];
    const key = keyValue[keyIdx];
    const value = inputSheet.getRange(keyValue[valueIdx], inputValueColNumber).getValue();
    if (x[keyIdx] !== options){
      sendMailInfo[key] = value;
    } else {
      sendMailInfo[options][key] = value;
    };
  });
  const fileId = inputSheet.getRange(fileIdAddress).getValue();
  const htmlString = getHtmlFromFile(fileId);
  sendMailInfo.to = sendMailInfo.testTo;
  sendMailInfo.options.htmlBody = htmlString;
  sendMailInfo.options.noReply = sendMailInfo.options.name === '' ? true : false;
  if (sendMailInfo.options.noReply){
    delete sendMailInfo.options.name;
  };
  /** send test */
  sendMailInfo.subject = '（テスト送信）' + sendMailInfo.subject; 
  let resSendMail = sendMail(sendMailInfo);
  if (!resSendMail){
    return;
  };
  sendMailInfo.subject = inputSheet.getRange(addressRow.subject, inputValueColNumber).getValue();
  /** production */
  const res = ui.prompt('本番送信する場合は半角小文字で"' + resString + '"と入力し、OKをクリックしてください。それ以外の操作をすると処理を終了します。', ui.ButtonSet.OK_CANCEL);
  if (res.getResponseText() === resString && res.getSelectedButton() === ui.Button.OK){
    sendMailInfo.to = inputSheet.getRange(mainToAddress).getValue();
    const bccSenders = getBccAddress_();
    if (bccSenders === null){
      sendMail(sendMailInfo);
    } else {
      bccSenders.some(bcc => {
        sendMailInfo.options.bcc = bcc;
        console.log(sendMailInfo.options.bcc);
        const resSendMail = sendMail(sendMailInfo);
        // The process ends when Cancel is clicked.
        if (!resSendMail){
          return true;
        } else {
          Utilities.sleep(2000);
        };
      });
    };
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
  };
  const res = ui.alert(alertInfo, ui.ButtonSet.OK_CANCEL);
  if (res === ui.Button.OK){
    GmailApp.sendEmail(sendMailInfo.to, 
                       sendMailInfo.subject, 
                       sendMailInfo.body,
                       sendMailInfo.options);
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
* @return {String} Pop-up menu contents
*/
function editAlertInfoStrings(sendMailInfo){
  let errorCheck = JSON.parse(JSON.stringify(sendMailInfo));
  if (errorCheck.options.htmlBody){
    delete errorCheck.options.htmlBody;
  } else {
    SpreadsheetApp.getUi().alert('error:htmlBodyが空白\n送信をキャンセルします。');
    return null;
  };
  const hideSendMailInfoKey = ['body', 'testTo', 'mainTo'];
  hideSendMailInfoKey.forEach(x => delete errorCheck[x]);
  const res = createAlertStrings(errorCheck);
  return 'OKをクリックするとメールが送信されます。キャンセルをクリックすると送信キャンセルします。\n\n' + res;
}
/**
 * Form output messages.
 * @param {Object}
 * @return {String}
 */
function createAlertStrings(inputObjects){
  const keyIdx = 0;
  const valueIdx = 1;
  const test = Object.entries(inputObjects).map(x => {
    if (Object.getPrototypeOf(x[valueIdx]).constructor.name === 'Object'){
      return createAlertStrings(x[valueIdx]);
    };
    return x[keyIdx] + ':' + x[valueIdx];
  });
  return test.flat().join('\n');
}
/**
 * Converts the email address entered in column A of the Bcc Destination List sheet into a comma-delimited string.
 * @param none.
 * @return {String} Returns a comma-separated string; null if column A is empty.
 */
function getBccAddress_(){
  const commonSettings = new CommonSettings();
  const inputSheet = commonSettings.sheets.bccSenders;
  const inputColumnNumber = 1;
  const lastRow = inputSheet.getLastRow();
  const bccSenders = lastRow > 0 ? inputSheet.getRange(1, inputColumnNumber, lastRow, 1).getValues().join(',') : null;
  return bccSenders;
}
/**
* Processing at file open
* @param none
* @return none
*/
function onOpen(){
  SpreadsheetApp.getActiveSpreadsheet().addMenu('メールマガジン送信', [{name: '送信', functionName:'sendNewsLetter'}]);
}
