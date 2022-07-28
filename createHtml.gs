function createHtmlFile() {
  const inputSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('コンテンツ作成');
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
  const editHtml = new EditHtml(newsletterTitle[0]);
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
  folder.createFile(blob);
}

class EditHtml{
  constructor(strTitle){
    this.header = this.setHeader() + this.setTitle(strTitle);
    this.footer = this.setFooter();

  }
  setHeader(){
    const res = '<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd"><html><head><meta http-equiv="Content-Type" content="text/html; charset=UTF-8p" /><meta http-equiv="content-style-type" content="text/css"><meta http-equiv="content-language" content="ja"><title>名古屋医療センター臨床研究センター情報システム研究室ニュースレター</title></head>';
    return res;
  }
  setTitle(strTitle){
    const res = '<body><!--最外のテーブル--><!--背景色設定--><table cellpadding="0" cellspacing="10" border="0" width="100%" bgcolor="#f4f4f4" align="center"><tbody><tr style="margin:0;padding:0"><td align="center"><!--ヘッダー：ロゴ--><table width="100%" cellspacing="0" cellpadding="0" bgcolor="#ffffff" align="center"style="margin:0;padding:0;width:100%;border-top:1px solid #d8d8d8;border-bottom:solid 10px #f4f4f4;font-size:100%;font-weight:normal;font-family:&#39;Lucida Grande&#39;,&#39;Hiragino Kaku Gothic ProN&#39;,&#39;\\0030d2\\0030e9\\0030ae\\0030ce\\0089d2\\0030b4ProN W3&#39;,Meiryo,メイリオ,sans-serif;"><tbody><tr style="margin:0;padding:0;"><td align="left" style="padding:10px 10px 10px 10px; border-bottom:solid 1px #d8d8d8;border-right:solid 1px #d8d8d8;border-left:solid 1px #d8d8d8"><a href="https://crc.nnh.go.jp/" target="_blank"style="float:left;vertical-align: middle;font-weight:700;font-size:200%;background:transparent;color:#00629d;text-decoration:none;" >情報システム研究室ニュースレター</a></td></tr></tbody></table><!--ヘッダー：ロゴ--> <!--ここから上段の記事--> <!--上段の記事のヘッダー--> <table width="100%" cellspacing="0" cellpadding="0" bgcolor="#00629d" align="center"style="margin:0;padding:0;font-size:100%;line-height:1.5;font-weight:normal;font-family:&#39;Lucida Grande&#39;,&#39;Hiragino Kaku Gothic ProN&#39;,&#39;\\0030d2\\0030e9\\0030ae\\0030ce\\0089d2\\0030b4  ProN W3&#39;,Meiryo,メイリオ,sans-serif"><tbody>  <tr> <td style="margin:0;padding:10px 10px 10px;border-bottom:#d8d8d8 1px solid;color:#fff;text-decoration:none;font-weight:700;font-size:120%;display:block" valign="top" align="left">' + strTitle + '</td>  </tr></tbody> </table> <!--ここから上段の記事の本体--> <table width="100%" cellspacing="0" cellpadding="0" bgcolor="#fff" align="center"style="margin:0;padding:0;font-size:100%;line-height:1.5;font-weight:normal;font-family:&#39;Lucida Grande&#39;,&#39;Hiragino Kaku Gothic ProN&#39;,&#39;\\0030d2\\0030e9\\0030ae\\0030ce\\0089d2\\0030b4  ProN W3&#39;,Meiryo,メイリオ,sans-serif;border-bottom:solid 10px #f4f4f4"><tbody>';
    return res;
  }
  setFooter(){
    const res = '</tbody></table><!--ここまで上段の記事の本体--><!--ここからFooter--> <table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" style="margin:0;padding-top:10px"><tbody>  <tr> <td><p  style="border-radius:5px;margin:0;padding:0;line-height:1.5;background-color:#616c72;text-align:center;font-size:100%;font-family:&#39;Lucida Grande&#39;,&#39;Hiragino Kaku Gothic ProN&#39;,&#39;\\0030d2\\0030e9\\0030ae\\0030ce\\0089d2\\0030b4  ProN W3&#39;,Meiryo,メイリオ,sans-serif">  <a href="http://info.nnh.go.jp" style="margin:0;padding:10px;display:block;color:#fff;text-decoration:none" target="_blank">臨床研究センターポータルサイト（info.nnh.go.jp）へ</a></p> </td>  </tr></tbody> </table> <table width="100%" border="0" cellspacing="0" cellpadding="0"style="margin:0;padding:0;font-size:84%;line-height:1.5;font-family:&#39;Lucida Grande&#39;,&#39;Hiragino Kaku Gothic ProN&#39;,&#39;\\0030d2\\0030e9\\0030ae\\0030ce\\0089d2\\0030b4  ProN W3&#39;,Meiryo,メイリオ,sans-serif"><tbody>  <tr> <td align="left" style="margin:0;padding:10px 0;color:#333333;font-size:100%;border-bottom:#333 2px solid"><p>本ニュースレターは、名古屋医療センター臨床研究センターに勤務する皆様にお届けしています。<br>  ご質問・ご意見などは、<a href="mailto:information.system@nnh.go.jp" target="_blank" style="color:#00629d;">情報システム研究室</a> 宛てにご連絡お願いいたします。</p> </td>  </tr>  <tr> <td align="left" style="margin:0;padding:10px 0;color:#333333;font-size:100%"><p>独立行政法人国立病院機構 名古屋医療センター 臨床研究センター 情報システム研究室 <br>〒460-0001 愛知県名古屋市中区三の丸 4-1-1<br>※NMC外部への掲載内容の無断転載を禁じます。<br> Copyright©  NHO Nagoya Medical Center, Clinical Research Center All Rights Reserved.</p> </td>  </tr></tbody> </table> <!--ここまでFooter-->  </td></tr> </tbody>  </table>  <!--最外のテーブル--></body></html>'
    return res;
  }
  setBodyText(bodyTitle, bodyText){
    const bodyHead = '<!--個別の記事-->  <tr> <td style="margin:0;padding:10px;border-bottom:1px solid #d8d8d8;border-right:1px solid #d8d8d8;border-left:1px solid #d8d8d8" valign="top" align="left"><div style="overflow:hidden;margin:0;padding:0">  <!--記事タイトル-->  <p style="margin:0;padding:0 0 10px;color:#333333;text-decoration:none;font-weight:700;font-size:120%;display:block">' + bodyTitle + '</p>  <!--記事タイトル-->  <!--本文-->  <div style="margin:0;padding:0;color:#333333">';
    const bodyFoot = '</div> <!--本文--> </div> </td> </tr> <!--個別の記事-->';
    return bodyHead + bodyText + bodyFoot;
  } 
}