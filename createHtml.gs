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
  const publisher = inputSheet.getRange('B4').getValue();
  let htmlInfo = getHeaderFooter_(publisher, commonSettings.sheets.headerAndFooter);
  if (htmlInfo === null){
    return;
  };
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
  htmlInfo.heading = newsletterTitle[0];
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
      targetText = targetText.substring(0, 4) === 'http' ? '<a href="' + targetText + '" target="_blank">' + targetText + '</a>' : targetText;
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
    this.header = this.header + '<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">' + '\n' + '\n';
    this.header = this.header + '<html>' + '\n' + '\n';
    this.header = this.header + '<!--色指定' + '\n';
    this.header = this.header + '枠線：#d8d8d8' + '\n';
    this.header = this.header + '' + '\n';
    this.header = this.header + '大外背景：#fffaf5' + '\n';
    this.header = this.header + 'ヘッダー背景：#fff' + '\n';
    this.header = this.header + 'ヘッダー文字：#00629d' + '\n';
    this.header = this.header + '' + '\n';
    this.header = this.header + 'イントロ背景：#fff' + '\n';
    this.header = this.header + 'イントロ文字：#333333' + '\n';
    this.header = this.header + '' + '\n';
    this.header = this.header + '全体表題背景：#00629d' + '\n';
    this.header = this.header + '全体表題文字：#fff' + '\n';
    this.header = this.header + '' + '\n';
    this.header = this.header + '本文セクション表題文字：#333333' + '\n';
    this.header = this.header + '本文背景:#fff' + '\n';
    this.header = this.header + '本文文字：#333333' + '\n';
    this.header = this.header + 'リンク文字：#00629d' + '\n';
    this.header = this.header + '' + '\n';
    this.header = this.header + 'フッターボタン：#616c72' + '\n';
    this.header = this.header + '-->' + '\n';
    this.header = this.header +   this.setHeader(htmlInfo.tabTitle);
    this.header = this.header + '<body>' + '\n';
    this.header = this.header + '  <!--最外のテーブル-->' + '\n';
    this.header = this.header + '  <!--背景色設定-->' + '\n';
    this.header = this.header + '  <table cellpadding="0" cellspacing="10" border="0" width="100%" bgcolor="#f4f4f4" align="center">' + '\n';
    this.header = this.header + '    <tbody>' + '\n';
    this.header = this.header + '      <tr style="margin:0;padding:0">' + '\n';
    this.header = this.header + '        <td align="center">' + '\n';
    this.header = this.header + '' + '\n';


    this.header = this.header +           this.setTitle(htmlInfo.titleUrl, htmlInfo.titleText);
    this.header = this.header +           this.setHeading(htmlInfo.heading);
    this.header = this.header + '          <!--ここから上段の記事の本体-->' + '\n';
    this.header = this.header + '          <table width="100%" cellspacing="0" cellpadding="0" bgcolor="#fff" align="center"' + '\n';
    this.header = this.header + '            style="margin:0;padding:0;font-size:100%;line-height:1.5;font-weight:normal;font-family:&#39;Lucida Grande&#39;,&#39;Hiragino Kaku Gothic ProN&#39;,&#39;\\0030d2\\0030e9\\0030ae\\0030ce\\0089d2\\0030b4  ProN W3&#39;,Meiryo,メイリオ,sans-serif;border-bottom:solid 10px #f4f4f4">' + '\n';
    this.header = this.header + '            <tbody>' + '\n';
    this.header = this.header + '' + '\n';
    this.footer = '';
    this.footer = this.footer + '            </tbody>' + '\n';
    this.footer = this.footer + '          </table>' + '\n';
    this.footer = this.footer + '          <!--ここまで上段の記事の本体-->' + '\n';
    this.footer = this.footer +           this.setFooterUrl(htmlInfo.footerUrlText, htmlInfo.footerUrl);
    this.footer = this.footer +           this.setFooter(htmlInfo.footerTargetText, htmlInfo.mailAddress, htmlInfo.mailName, htmlInfo.senderInformation);
    this.footer = this.footer + '        </td>' + '\n';
    this.footer = this.footer + '      </tr>' + '\n';
    this.footer = this.footer + '    </tbody>' + '\n';
    this.footer = this.footer + '  </table>' + '\n';
    this.footer = this.footer + '  <!--最外のテーブル-->' + '\n';
    this.footer = this.footer + '</body>' + '\n';
    this.footer = this.footer + '' + '\n';
    this.footer = this.footer + '</html>' + '\n';
    this.footer = this.footer + '' + '\n';
  }
  setHeader(tabTitle){
    let res = '';
    res = res + '<head>' + '\n';
    res = res + '  <meta http-equiv="Content-Type" content="text/html; charset=UTF-8p" />' + '\n';
    res = res + '  <meta http-equiv="content-style-type" content="text/css">' + '\n';
    res = res + '  <meta http-equiv="content-language" content="ja">' + '\n';
    res = res + '  <title>' + tabTitle + '</title>' + '\n';
    res = res + '</head>' + '\n';
    res = res + '' + '\n';
    return res;
  }
  setTitle(url, titleText){
    let res = '';
    res = res + '          <!--ヘッダー：ロゴ-->' + '\n';
    res = res + '          <table width="100%" cellspacing="0" cellpadding="0" bgcolor="#ffffff" align="center"' + '\n';
    res = res + '            style="margin:0;padding:0;width:100%;border-top:1px solid #d8d8d8;border-bottom:solid 10px #f4f4f4;font-size:100%;font-weight:normal;font-family:&#39;Lucida Grande&#39;,&#39;Hiragino Kaku Gothic ProN&#39;,&#39;\\0030d2\\0030e9\\0030ae\\0030ce\\0089d2\\0030b4  ProN W3&#39;,Meiryo,メイリオ,sans-serif;">' + '\n';
    res = res + '            <tbody>' + '\n';
    res = res + '              <tr style="margin:0;padding:0;">' + '\n';
    res = res + '                <td align="left" style="padding:10px 10px 10px 10px; border-bottom:solid 1px #d8d8d8;border-right:solid 1px #d8d8d8;border-left:solid 1px #d8d8d8">' + '\n';
    res = res + '                  <a href="' + url + '" target="_blank"' + '\n';
    res = res + '                  style="float:left;vertical-align: middle;font-weight:700;font-size:200%;background:transparent;color:#00629d;text-decoration:none;" >' + '\n';
    res = res + '                  ' + titleText + '\n';
    res = res + '                  </a>' + '\n';
    res = res + '                </td>' + '\n';
    res = res + '              </tr>' + '\n';
    res = res + '            </tbody>' + '\n';
    res = res + '          </table>' + '\n';
    res = res + '          <!--ヘッダー：ロゴ-->' + '\n';
    res = res + '' + '\n';
    return res;
  }
  setHeading(strHeading){
    let res = '';
    res = res + '          <!--ここから上段の記事-->' + '\n';
    res = res + '          <!--上段の記事のヘッダー-->' + '\n';
    res = res + '          <table width="100%" cellspacing="0" cellpadding="0" bgcolor="#00629d" align="center"' + '\n';
    res = res + '            style="margin:0;padding:0;font-size:100%;line-height:1.5;font-weight:normal;font-family:&#39;Lucida Grande&#39;,&#39;Hiragino Kaku Gothic ProN&#39;,&#39;\\0030d2\\0030e9\\0030ae\\0030ce\\0089d2\\0030b4  ProN W3&#39;,Meiryo,メイリオ,sans-serif">' + '\n';
    res = res + '            <tbody>' + '\n';
    res = res + '              <tr>' + '\n';
    res = res + '                <td style="margin:0;padding:10px 10px 10px;border-bottom:#d8d8d8 1px solid;color:#fff;text-decoration:none;font-weight:700;font-size:120%;display:block" valign="top" align="left">' + strHeading + '</td>' + '\n';
    res = res + '              </tr>' + '\n';
    res = res + '            </tbody>' + '\n';
    res = res + '          </table>' + '\n';
    res = res + '' + '\n';
    return res;
  }
  setFooterUrl(url, urlText){
    let res = '';
    res = res + '          <!--ここからFooter-->' + '\n';
    res = res + '          <table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" style="margin:0;padding-top:10px">' + '\n';
    res = res + '            <tbody>' + '\n';
    res = res + '              <tr>' + '\n';
    res = res + '                <td>' + '\n';
    res = res + '                  <p' + '\n';
    res = res + '                    style="border-radius:5px;margin:0;padding:0;line-height:1.5;background-color:#616c72;text-align:center;font-size:100%;font-family:&#39;Lucida Grande&#39;,&#39;Hiragino Kaku Gothic ProN&#39;,&#39;\\0030d2\\0030e9\\0030ae\\0030ce\\0089d2\\0030b4  ProN W3&#39;,Meiryo,メイリオ,sans-serif">' + '\n';
    res = res + '                    <a href="' + url + '" style="margin:0;padding:10px;display:block;color:#fff;text-decoration:none" target="_blank">' + urlText + '</a>' + '\n';
    res = res + '                  </p>' + '\n';
    res = res + '                </td>' + '\n';
    res = res + '              </tr>' + '\n';
    res = res + '            </tbody>' + '\n';
    res = res + '          </table>' + '\n';
    return res;
  }
  setFooter(target, mailAddress, mailName, senderInformation){
    let res = '';
    res = res + '          <table width="100%" border="0" cellspacing="0" cellpadding="0"' + '\n';
    res = res + '            style="margin:0;padding:0;font-size:84%;line-height:1.5;font-family:&#39;Lucida Grande&#39;,&#39;Hiragino Kaku Gothic ProN&#39;,&#39;\\0030d2\\0030e9\\0030ae\\0030ce\\0089d2\\0030b4  ProN W3&#39;,Meiryo,メイリオ,sans-serif">' + '\n';
    res = res + '            <tbody>' + '\n';
    res = res + '              <tr>' + '\n';
    res = res + '                <td align="left" style="margin:0;padding:10px 0;color:#333333;font-size:100%;border-bottom:#333 2px solid">' + '\n';
    res = res + '                  <p>' + target + 'ご質問・ご意見などは、<a href="mailto:' + mailAddress + '" target="_blank" style="color:#00629d;">' + mailName + '</a> 宛てにご連絡お願いいたします。</p>' + '\n';
    res = res + '                </td>' + '\n';
    res = res + '              </tr>' + '\n';
    res = res + '              <tr>' + '\n';
    res = res + '                <td align="left" style="margin:0;padding:10px 0;color:#333333;font-size:100%">' + '\n';
    res = res + '                  <p>' + senderInformation + '</p>' + '\n';
    res = res + '                </td>' + '\n';
    res = res + '              </tr>' + '\n';
    res = res + '            </tbody>' + '\n';
    res = res + '          </table>' + '\n';
    res = res + '          <!--ここまでFooter-->' + '\n';
    res = res + '' + '\n';
  return res;
  }
  setBodyText(bodyTitle, bodyText){
    let res = '';
    res = res + '              <!--個別の記事-->' + '\n';
    res = res + '              <tr>' + '\n';
    res = res + '                <td style="margin:0;padding:10px;border-bottom:1px solid #d8d8d8;border-right:1px solid #d8d8d8;border-left:1px solid #d8d8d8" valign="top" align="left">' + '\n';
    res = res + '                  <div style="overflow:hidden;margin:0;padding:0">' + '\n';
    res = res + '                    <!--記事タイトル-->' + '\n';
    res = res + '                    <p style="margin:0;padding:0 0 10px;color:#333333;text-decoration:none;font-weight:700;font-size:120%;display:block">' + bodyTitle + '</p>' + '\n';
    res = res + '                    <!--記事タイトル-->' + '\n';
    res = res + '                    <!--本文-->' + '\n';
    res = res + '                    <div style="margin:0;padding:0;color:#333333;word-break:break-all;">' + '\n';
    res = res + '                      ' + bodyText + '\n';
    res = res + '                    </div>' + '\n';
    res = res + '                    <!--本文-->' + '\n';
    res = res + '                  </div>' + '\n';
    res = res + '                </td>' + '\n';
    res = res + '              </tr>' + '\n';
    res = res + '              <!--個別の記事-->' + '\n';
    res = res + '' + '\n';
    return res;
  } 
}
/**
 * Header and footer information is obtained from the publisher's name.
 * @param {String} The publisher's name.
 * @param {Object} Data source sheet objects for input rules.
 * @return {Object} Header and footer information.
 */
function getHeaderFooter_(publisher, sourceDataSheet){
  const keyList = ['mailName', 'tabTitle', 'titleText', 'titleUrl', 'footerUrl', 'footerUrlText', 'footerTargetText', 'mailAddress', 'senderInformation'];
  const sourceDataPublisherRowIndex = 0;
  const sourceData = sourceDataSheet.getRange(1, 2, sourceDataSheet.getLastRow(), sourceDataSheet.getLastColumn() -1).getValues();
  const targetIndex = sourceData[sourceDataPublisherRowIndex].indexOf(publisher);
  const targetArray = targetIndex > -1 ? sourceData.map(x => x[targetIndex]) : null;
  if (targetArray === null){
    return;
  };
  let headerAndFooterInfo = {};
  keyList.forEach((key, idx) => headerAndFooterInfo[key] = targetArray[idx]);
  return headerAndFooterInfo;
}