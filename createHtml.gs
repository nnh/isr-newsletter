/**
 * Create html files from the content of the document.
 * @param none.
 * @return none.
 */
function createHtmlFile(){
  const commonSettings = new CommonSettings();
  if (!commonSettings.sheetStatus){
    return;
  };
  const inputSheet = commonSettings.sheets.createContent;
  const urlOutputSheet = commonSettings.sheets.sendNewsLetter;
  const inputDocId = inputSheet.getRange('C1').getValue();
  const inputDoc = DocumentApp.openById(inputDocId);
  const outputFolder = inputSheet.getRange('C2').getValue();
  const folder = DriveApp.getFolderById(outputFolder);
  const outputFilename = inputSheet.getRange('C3').getValue();
  const paragraphs = inputDoc.getBody().getParagraphs();
  const paragraphItems = paragraphs.map(x => [x.getText(), x.getType().toString(), x.getHeading().toString()]);
  let paragraphItemsIndex = {};
  paragraphItemsIndex.text = 0;
  paragraphItemsIndex.heading = 2;
  const newsletterTitle = paragraphItems.filter(x => x[paragraphItemsIndex.heading] === 'TITLE').map(x => x[paragraphItemsIndex.text]);
  if (newsletterTitle.length < 0){
    return;
  }
  let htmlInfo = {};
  htmlInfo.titleUrl = 'https://crc.nnh.go.jp/';
  htmlInfo.titleText = '情報システム研究室ニュースレター';
  htmlInfo.tabTitle = '名古屋医療センター臨床研究センター情報システム研究室ニュースレター';
  htmlInfo.heading = newsletterTitle[0];
  htmlInfo.footerUrl = 'http://info.nnh.go.jp';
  htmlInfo.footerUrlText = '臨床研究センターポータルサイト（info.nnh.go.jp）へ';
  htmlInfo.footerTargetText = '本ニュースレターは、名古屋医療センター臨床研究センターに勤務する皆様にお届けしています。<br>  ';
  htmlInfo.mailAddress = 'information.system@nnh.go.jp';
  htmlInfo.mailName = '情報システム研究室';
  htmlInfo.senderInformation = '独立行政法人国立病院機構 名古屋医療センター 臨床研究センター 情報システム研究室 <br>〒460-0001 愛知県名古屋市中区三の丸 4-1-1<br>※NMC外部への掲載内容の無断転載を禁じます。<br> Copyright©  NHO Nagoya Medical Center, Clinical Research Center All Rights Reserved.';
  const editHtml = new EditHtml(htmlInfo);
  const newsletterBodyItems = paragraphItems.filter(x => x[paragraphItemsIndex.heading] !== 'TITLE').concat([['dummy', null, 'HEADING1']]);
  let saveTitle = '';
  let title = '';
  let body = '';
  let htmlText = '';
  newsletterBodyItems.forEach(item => {
    title = item[paragraphItemsIndex.heading] === 'HEADING1' ? item[paragraphItemsIndex.text] : title;
    if (item[paragraphItemsIndex.heading] === 'NORMAL'){
      saveTitle = title;
      let targetText = item[paragraphItemsIndex.text];
      targetText = targetText.substring(0, 4) === 'http' ? '<a href="' + targetText + '">' + targetText + '</a>' : targetText;
      body = body != '' ? body + '<br>' + targetText : targetText;
    } else {
      if (body != ''){
        htmlText = htmlText + editHtml.setBodyText(saveTitle, body);
        body = '';
      }
    }
  });
  const output = editHtml.header + htmlText + editHtml.footer;
  const contentType = 'text/plain';
  const charset = 'utf-8';
  const blob = Utilities.newBlob('', contentType, outputFilename).setDataFromString(output, charset);
  const newFile = folder.createFile(blob);
  urlOutputSheet.getRange('B1').setValue(newFile.getUrl());
}
class EditHtml{
  constructor(htmlInfo){
    this.header = '';
    this.header = this.header + '<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">';
    this.header = this.header + '<html>';
    this.header = this.header +   this.setHeader(htmlInfo.tabTitle);
    this.header = this.header +   '<body>';
    this.header = this.header +     '<table cellpadding="0" cellspacing="10" border="0" width="100%" bgcolor="#f4f4f4" align="center"><tbody>';
    this.header = this.header +       '<tr style="margin:0;padding:0">';
    this.header = this.header +         '<td align="center">';
    this.header = this.header +           this.setTitle(htmlInfo.titleUrl, htmlInfo.titleText);
    this.header = this.header +           this.setHeading(htmlInfo.heading);
    this.header = this.header +           '<table width="100%" cellspacing="0" cellpadding="0" bgcolor="#fff" align="center"style="margin:0;padding:0;font-size:100%;line-height:1.5;font-weight:normal;font-family:&#39;Lucida Grande&#39;,&#39;Hiragino Kaku Gothic ProN&#39;,&#39;\\0030d2\\0030e9\\0030ae\\0030ce\\0089d2\\0030b4  ProN W3&#39;,Meiryo,メイリオ,sans-serif;border-bottom:solid 10px #f4f4f4"><tbody>';
    this.footer = '';
    this.footer = this.footer +           '</tbody></table>';
    this.footer = this.footer +           this.setFooterUrl(htmlInfo.footerUrl, htmlInfo.footerUrlText);
    this.footer = this.footer +           this.setFooter(htmlInfo.footerTargetText, htmlInfo.mailAddress, htmlInfo.mailName, htmlInfo.senderInformation);
    this.footer = this.footer +         '</td>';
    this.footer = this.footer +       '</tr>'
    this.footer = this.footer +     '</tbody></table>';
    this.footer = this.footer +   '</body>';
    this.footer = this.footer + '</html>';
  }
  setHeader(tabTitle){
    let res = '';
    res = res + '<head>';
    res = res +   '<meta http-equiv="Content-Type" content="text/html; charset=UTF-8p" />';
    res = res +   '<meta http-equiv="content-style-type" content="text/css">';
    res = res +   '<meta http-equiv="content-language" content="ja">';
    res = res +   '<title>' + tabTitle + '</title>';
    res = res + '</head>';
    return res;
  }
  setTitle(url, titleText){
    let res = '';
    res = res +     '<table width="100%" cellspacing="0" cellpadding="0" bgcolor="#ffffff"align="center"style="margin:0;padding:0;width:100%;border-top:1px solid #d8d8d8;border-bottom:solid 10px #f4f4f4;font-size:100%;font-weight:normal;font-family:&#39;Lucida Grande&#39;,&#39;Hiragino Kaku Gothic ProN&#39;,&#39;\\0030d2\\0030e9\\0030ae\\0030ce\\0089d2\\0030b4ProN W3&#39;,Meiryo,メイリオ,sans-serif;">';
    res = res +       '<tbody>';
    res = res +         '<tr style="margin:0;padding:0;">';
    res = res +           '<td align="left" style="padding:10px 10px 10px 10px; border-bottom:solid 1px #d8d8d8;border-right:solid 1px #d8d8d8;border-left:solid 1px #d8d8d8">';
    res = res +             '<a href="' + url + '" target="_blank"style="float:left;vertical-align: middle;font-weight:700;font-size:200%;background:transparent;color:#00629d;text-decoration:none;" >';
    res = res +               titleText;
    res = res +             '</a>';
    res = res +           '</td>';
    res = res +         '</tr>';
    res = res +       '</tbody>';
    res = res +     '</table>';
    return res;
  }
  setHeading(strHeading){
    let res = '';
    res = res + '<table width="100%" cellspacing="0" cellpadding="0" bgcolor="#00629d"align="center"style="margin:0;padding:0;font-size:100%;line-height:1.5;font-weight:normal;font-family:&#39;Lucida Grande&#39;,&#39;Hiragino Kaku Gothic ProN&#39;,&#39;\\0030d2\\0030e9\\0030ae\\0030ce\\0089d2\\0030b4  ProN W3&#39;,Meiryo,メイリオ,sans-serif">';
    res = res +   '<tbody><tr><td style="margin:0;padding:10px 10px 10px;border-bottom:#d8d8d8 1px solid;color:#fff;text-decoration:none;font-weight:700;font-size:120%;display:block" valign="top" align="left">'
    res = res +     strHeading;
    res = res +   '</td></tr></tbody>';
    res = res + '</table>';
    return res;
  }
  setFooterUrl(url, urlText){
    let res = '';
    res = res + '<table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" style="margin:0;padding-top:10px">';
    res = res +   '<tbody><tr><td>';
    res = res +     '<p style="border-radius:5px;margin:0;padding:0;line-height:1.5;background-color:#616c72;text-align:center;font-size:100%;font-family:&#39;Lucida Grande&#39;,&#39;Hiragino Kaku Gothic ProN&#39;,&#39;\\0030d2\\0030e9\\0030ae\\0030ce\\0089d2\\0030b4  ProN W3&#39;,Meiryo,メイリオ,sans-serif">';
    res = res +       '<a href="' + url + '" style="margin:0;padding:10px;display:block;color:#fff;text-decoration:none" target="_blank">';
    res = res +         urlText;
    res = res +       '</a>';
    res = res +     '</p>';
    res = res +   '</td></tr></tbody>';
    res = res + '</table>';
    return res;
  }
  setFooter(target, mailAddress, mailName, senderInformation){
    let res = '';
    res = res + '<table width="100%" border="0" cellspacing="0" cellpadding="0"style="margin:0;padding:0;font-size:84%;line-height:1.5;font-family:&#39;Lucida Grande&#39;,&#39;Hiragino Kaku Gothic ProN&#39;,&#39;\\0030d2\\0030e9\\0030ae\\0030ce\\0089d2\\0030b4  ProN W3&#39;,Meiryo,メイリオ,sans-serif">';
    res = res +   '<tbody>';
    res = res +     '<tr><td align="left" style="margin:0;padding:10px 0;color:#333333;font-size:100%;border-bottom:#333 2px solid">';
    res = res +       '<p>';
    res = res +          target;
    res = res +         'ご質問・ご意見などは、';
    res = res +         '<a href="mailto:' + mailAddress + '"target="_blank" style="color:#00629d;">';
    res = res +           mailName;
    res = res +         '</a> 宛てにご連絡お願いいたします。';
    res = res +       '</p>';
    res = res +     '</td></tr>';
    res = res +     '<tr><td align="left" style="margin:0;padding:10px 0;color:#333333;font-size:100%">';
    res = res +       '<p>';
    res = res +         senderInformation;
    res = res +       '</p>';
    res = res +     '</td></tr>';
    res = res +   '</tbody>';
    res = res + '</table>';
  return res;
  }
  setBodyText(bodyTitle, bodyText){
    let res = '';
    res = res + '<tr>';
    res = res +   '<td style="margin:0;padding:10px;border-bottom:1px solid #d8d8d8;border-right:1px solid #d8d8d8;border-left:1px solid #d8d8d8" valign="top" align="left">';
    res = res +     '<div style="overflow:hidden;margin:0;padding:0">';
    res = res +       '<p style="margin:0;padding:0 0 10px;color:#333333;text-decoration:none;font-weight:700;font-size:120%;display:block">';
    res = res +         bodyTitle;
    res = res +       '</p>';
    res = res +       '<div style="margin:0;padding:0;color:#333333">';
    res = res +         bodyText;
    res = res +       '</div>';
    res = res +     '</div>';
    res = res +    '</td>';
    res = res + '</tr>';
    return res;
  } 
}