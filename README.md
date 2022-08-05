# エクセルのテーブルを SQLで読み出します
<br>

## 特徴
* Where 句が使えます（プレースホルダは ? です）
* パラーメタは AddParam で追加します
* OpenRecordset でクエリを開きます
* Fields("カラム名") でカラムのデータを取り出します
* MoveNext で次のレコードを読み出します
 
<br>
 
## 使い方

1. ソースコードをダウンロードします
   ``` bash
   git clone https://github.com/uhoogohan/uquery.git
   ```
1. エクセルを開き Alt + F11 で、コードエディタを開きます
 
1. Ctrl + M でつぎの３つのファイルを順に読み込みます
   * UConnection
   * UQuery
   * UStringlist
1. ツール(T) -> 参照設定(R)...を開き、次のものをチェックします
   * Microsoft ActiveX Data Objects 6.1 Library
1. "ごはん" という名前のシートに、つぎのテーブルを貼り付けます
   |id|メニュー|種類|価格|
   |---|---|---|---|
   |1|そば|麺|280|
   |2|うどん|麺|250|
   |3|親子丼|米|500|
   |4|ラーメン|麺|700|
   |5|炒飯|米|550|
   |6|ホッピーセット|酒|600|
   |7|鶏ももの唐揚げ|おかず|350|
1. コードエディタに戻り、標準モジュールを追加し、見本コードを貼り付けます
1. いったんエクセルをマクロ有効ブックで保存します
1. 開発 -> 挿入 -> ボタン（フォームコントロール）の順にクリックします
1. シート上に任意の矩形をドラッグ＆ドロップします
1. マクロの登録画面から ボタン_Click を選択しＯＫをクリックします
1. ESC でボタンの選択を解除し、ボタンを押します
 
<br>
 
## 見本コード
 
```vba
Option Explicit

Public con As New UConnection


Sub ボタン_Click()
    Dim q As New UQuery

    q.SetCon con
    q.Sql.Add "SELECT * FROM [ごはん$]"
    q.Sql.Add "WHERE 価格 >= ?"
    q.AddParam "価格", 300
    q.OpenRecordset

    Do Until q.hasNext
        MsgBox q.Fields("メニュー") & " ￥" & q.Fields("価格")
        q.MoveNext
    Loop
End Sub
```

<br><br>
## 見本イメージ
![sample1](https://cdn-ak.f.st-hatena.com/images/fotolife/u/uhoo/20181123/20181123162004.gif)
 
 
<br><br>
## クラスの説明
 
<br>
 
### UConnection
 

ADODB.Connection をラップします。
コネクションは、インスタンス作成時に開かれ、消滅時に閉じられます。
なので、明示的にコネクションを開いたり閉じたりしなくてもいいです。
 
 
|関数とプロシージャ|説明|
|---|---|
|OpenConnection|コネクションを開きます|	
|CloseConnection|コネクションを閉じます|
|GetCon|コネクションを返します|
|Class_Initialize|OpenConnectionを実行します|
|Class_Terminate|CloseConnectionを実行します|
 
<br><br>
 
### UQuery
 

ADODB.RecordsetをラップしますAdd でSQL文を追加していきます。テーブル名は [シート名$] のように $ 記号が必要です。
 
``` vb
    q.SetCon con
    q.Sql.Add "SELECT * FROM [ごはん$]"
    q.Sql.Add "WHERE 価格 >= ?"
    q.AddParam "価格", 300
```
 
クエリは、インスタンス作成時に開かれ、消滅時に閉じられます。
なので、明示的にクエリを開いたり閉じたりしなくていいです。
唯一DELETEが使えません！！！
 
 
|関数とプロシージャ|説明|
|---|---|
|SetCon|コネクションをセットします|
|AddParam|文字列のパラメータを追加します|
|AddParamByInt|整数のパラメータを指定します|
|AddParamByDouble|実数のパラメータを指定します|
|OpenRecordset|クエリを開きます|
|MoveNext|クエリを１件読み出します|
|CloseRecordset|クエリを閉じます|
|clear_param|パラメータをクリアします|
|hasNext|次のレコードを読み出します|
|RecordCount%|レコード数を返します|
|fields|カラムのデータを返します|
|get_fields|フィールドの一覧をカンマで区切って返します|
|execute|実行します。UPDATEやINSERT時に使います|
|Class_Terminate|クエリを閉じます|
 
<br><br>
 
 
おわり
