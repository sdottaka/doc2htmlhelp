doc2htmlhelp.vbs Version 1.1

[目的]
  WordドキュメントからMicrosoft HTML Help 形式のヘルプファイル (.chm)を
  作成するソフトウェアです。VBScriptで記述されています。
  変換元のWordドキュメントに目次があれば、生成される HTML Helpにも目次が
  作成されます。

[対応OS]
・Windows XP, Vista
  作者は上記の環境で確認していますが、Windows 98, 2000等でも動くと思います。

[必要なソフトウェア]
・Microsoft Word (2000, XP, 2003, 2007)

・Microsoft HTML Help Workshop 
  (以下のMicrosoftのサイトから無料でダウンロードできます)
  http://www.microsoft.com/japan/office/ork/appndx/appa06.mspx
  のhtmlhelp.exeのリンクをクリックしてください。

[使い方]
・簡単な使い方

  展開された doc2htmlhelp.vbsの上に、変換したいWord ファイルをドラッグ＆
  ドロップするか、doc2htmlhelp.vbsをダブルクリックし、Wordの[ファイルを開く]
  ウインドウにて変換したWordファイルを指定してください。

  しばらくすると、変換したいWordファイルが存在するフォルダ内のWordファイル
  と同じ名前のフォルダにHTML Help ファイルが生成されます。
  ※ドキュメントの大きさによっては、分単位の時間がかかることがあります。

・コマンドライン指定

  上記の方法では、詳細な指定(生成先フォルダや分割するドキュメントレベル等)が
  できませんが、下記のように引数に指定することで生成先フォルダなどを指定するこ
  とができます。

  [コマンドライン]
  cscript.exe doc2htmlhelp.vbs [Wordドキュメントファイルパス] [/?] [/Title:HTML Help タイトル] [/DestDir:HTML Help生成先フォルダ] [/DivDocLevel:HTML分割するドキュメントレベル] [/MarginLeft:左余白の幅]

  /?
		ヘルプを表示

  /Title:タイトル
		HTML Helpのタイトルを指定します。
		省略された場合、ドキュメントのプロパティのタイトルが使われます。
		もし、そのプロパティのタイトルが空白だった場合、Wordドキュメン
		トファイル名と同じ名称がつかわれます。

  /DestDir:生成先フォルダ
		生成先フォルダを指定します。
		省略された場合、生成先フォルダは、Wordドキュメントファイル
		が存在するフォルダ内のファイル名と同じフォルダになります。

  /DivDocLevel:分割するドキュメントレベル
		分割するドキュメントレベルを指定します。
		たとえば、1を指定すると見出し1のレベルでHTMLファイルが分割
		され、2を指定すると見出し2のレベルでHTMLファイルが分割
		されます。
		※目次の最大レベルより大きい値を指定しないでください。
		省略された場合、1となります。

  /MarginLeft:左余白の幅
		左余白の幅をピクセル単位で指定します。
		-9999を指定すると見出しなどが表示範囲内に収まるように自動調整します。
		省略された場合、-9999となります。

  [指定例]
  cscript c:\tools\doc2htmlhelp.vbs c:\tools\doc2htmlhelp.doc /Title:doc2htmlhelp説明書 /DivDocLevel:1

[使用条件]
  本ソフトウェアは、MITライセンスです。商用でもフリーでもご利用いただけます。
  MITライセンスについては以下をご覧ください。
  http://ja.wikipedia.org/wiki/MIT_License
  http://www.opensource.org/licenses/mit-license.php

[履歴]
  2008-05-30 1.1 HTMLファイルがUTF-8で出力された場合、目次等が文字化けしてしまう問題の修正
  2008-05-17 1.0 初回リリース

[連絡先]
  mailto:s7taka@gmail.com

