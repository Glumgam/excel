# エクセルに関すること  
## 初めに  
ブック ⇒ エクセルファイルそのもの  
シート ⇒ ブック内に生成された作業場  
  
ブック内のシートは左から **0から数える**  
> **Note**  
実際のシート名は Sheet1 となっていますが、内部では 0番目のシート となります。  
  
## [1]パスワードが設定されたシートの保護を解除する方法  
> **Warning**  
解除自体には違法性はないものの、組織内的にアウトの場合もあります。  
使用に関しては自己責任でお願いします。  
  
* 解除したいシートが**何番目**にあるか確認する  
* 拡張子を確認する(以下、例として xlsm とする)
* エクセルファイルの拡張子を zip にする  
* 解凍する  
  
  * _rels  
    docProps  
    xl  
    [content_types].xml  
  
    という構成になっている  
* xl/worksheets/sheet*.xmlをメモ帳等で開く  

> **Note**  
*は数字が入ります。シートの数だけ生成されます  
先ほど確認した**何番目**の数字のファイルをメモ帳等で開いてください  
  
ここから2通りの方法があります  
### パスワードを変更する  
* ctrl+f で sheetProtection と検索
* `<sheetProtection password="12AB" sheet="1" objects="1" scenarios="1"/>`  という記載を見つける
* 12AB を 83AF に書き換える。するとパスワードが password に変更される  
*  _rels  
    docProps  
    xl  
    [content_types].xml  
    の4つ全てを選択し、左クリック⇒送る⇒圧縮(zip形式)フォルダでzipにする
*  拡張子を元に戻す(例 .zip⇒.xlsm)
*  エクセルを開き、シートの保護を解除する。パスワードはpasswordである。
  
### パスワードを削除する  
* ctrl+f で sheetProtection と検索
* `<sheetProtection password="12AB" sheet="1" objects="1" scenarios="1"/>`  を削除する
*  _rels  
   docProps  
   xl  
   [content_types].xml  
   の4つ全てを選択し、左クリック⇒送る⇒圧縮(zip形式)フォルダでzipにする
*  拡張子を元に戻す(例 .zip⇒.xlsm)
  
  
## [2]ダブルクリックで画像を張り付ける
> **Note**  
> このコードはセルの中央に画像を張り付けます  
> サイズはセルに収まるように縦横比を保ったまま、自動調整されます  
  
開発⇒Visual Basicで次のコードを張り付ける  
```
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, _
                                        Cancel As Boolean)
    'Glumgam
    Dim i As Long
    Dim FileName As Variant
    Dim X As Double
    Dim Y As Double
    
    FileName = Application.GetOpenFilename( _
        FileFilter:="画像ファイル,*.bmp;*.jpg;*.gif;*png", _
        MultiSelect:=True)
    If Not IsArray(FileName) Then
        Exit Sub
    End If
    
    For i = LBound(FileName) To UBound(FileName)
        With ActiveSheet.Shapes.AddPicture( _
            FileName:=FileName(i), _
            LinkToFile:=False, _
            SaveWithDocument:=True, _
            Left:=Target.Left, _
            Top:=Target.Top, _
            Width:=-1, _
            Height:=-1)
            .LockAspectRatio = msoTrue
            'glumgam
            X = Target.Width / .Width
            Y = Target.Height / .Height
            
            If X > Y Then
                .Height = .Height * Y
            Else
                .Width = .Width * X
            End If
            
            .Left = Target.Left + (Target.Width - .Width) / 2
            .Top = Target.Top + (Target.Height - .Height) / 2
        End With
    Next i
    Application.ScreenUpdating = True
    Cancel = True
End Sub
```  
