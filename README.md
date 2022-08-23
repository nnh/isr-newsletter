# isr-newsletter
## 概要
ニュースレター用のHTMLファイルを作成し、送信するためのGoogleスプレッド用スクリプトです。
## 使い方
Googleドキュメントで記事を作成し、コンテンツ作成シートに入出力情報を設定します。  
拡張機能 > Apps Script > createHtml.gs > createHtmlFileを実行してHTMLファイルを出力します。  
スプレッドシートのメニュー > メールマガジン送信 > 送信をクリックしてニュースレターを配信します。  
## 前提条件
- スプレッドシートに「ニュースレター送信」「Bcc送信先一覧」「コンテンツ作成」の3シートが存在する必要があります。  
- 「ニュースレター送信」シートには下記の情報が設定されている必要があります。
  - C1セルにHTMLファイルのIDが設定されている。
  - B2セルにメールのタイトルが設定されている。
  - B3セルに空白、またはメールの送信者名が設定されている。  
  空白の場合、送信者がnoreply@...になる。空白でない場合、スクリプト実行者のメールアドレスで送信される。  
  - B4セルにHTMLメールが表示できない場合の代替情報が設定されている。
  - B5セルにテスト用の送信先メールアドレスが設定されている。
  - B6セルに本番送信用の送信先メールアドレスが設定されている。
- 「Bcc送信先一覧」には下記の情報が設定されている必要があります。
  - Bccでの送信を行う場合、A列にBcc送信先のメールアドレスが、一セルにつき一アドレスの形式で設定されている。  
    Bccでの送信を行わない場合、A列のセルをすべて空白に設定する。  
- 「コンテンツ作成」シートには下記の情報が設定されている必要があります。
  - C1セルにコンテンツ作成用のGoogleドキュメントのIDが設定されている。
  - C2セルに出力フォルダのIDが設定されている。
  - C3セルに出力するHTMLファイルの名前が設定されている。
- コンテンツ作成用のGoogleドキュメントは下記のように記載されている必要があります。
  - タイトルの段落スタイルが「タイトル」になっている。
  - コンテンツの見出しの段落スタイルが「見出し１」になっている。
  - コンテンツのテキストの段落スタイルが「標準テキスト」になっている。
