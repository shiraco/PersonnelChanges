## What's this?
横に並んだ兼務の人の所属情報を縦に並べるように変換するExcel VBA ツール

## 対象ファイルと説明
* 人事異動_兼務者レイアウト変換ツール_全社員用.xlsm
全社員の現在（新）所属のレイアウトを変換します。

* 人事異動_兼務者レイアウト変換ツール_異動者用.xlsm
移動対象の社員の現在（新）所属のレイアウト、ならびに、過去（旧）所属のレイアウトを変換します。

## 使い方（全社員用・異動者用共通）
※処理時間の短縮と、万が一の予期せぬ異常終了に備えて、事前に他の不要なExcel ブックは閉じておいてください。
* 1．Excelを開く。
* 2．"source" シートの対応する列に人事異動データをコピペ（値貼り付け）する。
* 3．"target" シートを開く。(2回連続実行などで、シートの"target_table" にデータがある場合は、全データを削除しておく。)
* 4．設定でマクロをを有効にした上で、開発タブ＞マクロから"ConvertFor*()" を実行する。
* 5．列幅などレイアウトを整える。（主に以下の項目）
  * 文字がセル内に収まっているか確認
  * 改ページで表示範囲設定
  * ページ跨りの兼務者がいる場合の改ページ位置調整
  * 複数者へ出向している人の調整
* 6．"target" シートをPDFとして出力する。

## 動作環境
おそらく Excel 2007 以上。
（Excel2010 では動作確認済み。）

## 参考
* [開発] タブの表示の仕方（Excel 2007 以降）
リボンの [ファイル]タブ ＞ [オプション] ＞ [リボンのユーザー設定] ＞ [開発] チェック ボックスをオン
http://office.microsoft.com/ja-jp/excel-help/HA101819080.aspx

* マクロの許可方法（Excel 2007 以降）
リボンの「開発」タブ ＞ [マクロのセキュリティ] ＞[ VBAプロジェクトオブジェクトモデルへのアクセスを信頼する] をオン
http://support.microsoft.com/kb/282830/ja

## 未対応事項
* マクロ実行のボタン化
* 事前にデータが入っている場合のデータのクリア処理

## （開発者向け）開発版の補足・注意事項
（配布版ではマクロをExcel ブックに埋め込んだ上で配布しているので、開発者以外のユーザーはこの項目は無視して下さい。）

* Excel 本体からのVBAスクリプトの分離について
コードと、Excel 本体を分離させるために、以下の参考記事の（２）を参考にExcel VBA のモジュールを .bas ファイルとして外出ししています。
そして、Excel (xlsm) を開いたときに、自動的に、モジュールとして読み込むように設定しています。
この自動で読み込む際に既に標準モジュールが登録されていたらクリア（すなわち削除）するようにしています。
要するに、Excel VB Editor 上で行った修正は、Excelブック上に保存しても、次回起動時には削除されてしまいます。

* 配布方法について
Excel 本体に、.bas ファイルを読み込んで、マクロを埋め込んだ上、ThisWorkBook の自動.bas ファイル読み込みマクロを削除（ThisWorkBookの全スクリプトを削除）して配布して下さい。

* このツールのリポジトリ(origin)
ツールのオリジナルは、以下にあります。
https://github.com/tech-sketch/PersonnelChanges (private repository)

* 参考記事
http://d.hatena.ne.jp/language_and_engineering/20090731/p1

## 変更履歴

1.0 2014/03/31 初版作成
1.1 2014/04/21 新旧所属のある異動者用のフォーマットにも対応
1.2 2014/05/02 新旧職種列の追加、日付表記を漢字表記化、ヘッダー行の下罫線の設定


