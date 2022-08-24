// Working temporary scripts
function myFunction(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetname = 'ヘッダー・フッター情報'
  const checkAddSheet = ss.getSheetByName(sheetname);
  if (!checkAddSheet){
    ss.insertSheet(ss.getSheets().length);
    ss.getActiveSheet().setName(sheetname);
  };
  const targetSheet = ss.getSheetByName(sheetname);
  const targetValues = [['タイトル', '名古屋医療センター臨床研究センター情報システム研究室ニュースレター'], 
                        ['ヘッダーのタイトル', '情報システム研究室ニュースレター'],
                        ['ヘッダーのURL', 'https://crc.nnh.go.jp/'],
                        ['フッターのタイトル', '臨床研究センターポータルサイト（info.nnh.go.jp）へ'],
                        ['フッターのURL', 'http://info.nnh.go.jp'],
                        ['フッターのテキスト１','本ニュースレターは、名古屋医療センター臨床研究センターに勤務する皆様にお届けしています。<br>  '],
                        ['メールアドレス', 'information.system@nnh.go.jp'],
                        ['発行者名', '情報システム研究室'],
                        ['フッターのテキスト２','独立行政法人国立病院機構 名古屋医療センター 臨床研究センター 情報システム研究室 <br>〒460-0001 愛知県名古屋市中区三の丸 4-1-1<br>※NMC外部への掲載内容の無断転載を禁じます。<br> Copyright©  NHO Nagoya Medical Center, Clinical Research Center All Rights Reserved.']];
  targetSheet.clear();
  targetSheet.getRange(1, 1, targetValues.length, targetValues[0].length).setValues(targetValues);

}