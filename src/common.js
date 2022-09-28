/**
 * Class for setting common items.
 */
class CommonSettings{
  constructor(){
    this.sheets = {};
    this.sheets.createContent = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('コンテンツ作成');
    this.sheets.sendNewsLetter = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ニュースレター送信');
    this.sheets.bccSenders = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Bcc送信先一覧');
    this.sheets.headerAndFooter = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ヘッダー・フッター情報');
    const valueIdx = 1;
    this.sheetStatus = Object.entries(this.sheets).every(x => x[valueIdx]);
    if (!this.sheetStatus){
      SpreadsheetApp.getUi().alert('必要なシートが不足しているため処理を終了します');
    };
  }
}
