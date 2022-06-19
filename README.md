# エクセルに関する情報  
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
レベルの高低は問いませんが、過去にプロテクトを解除した事がある人向けです。  
(やり方を書くのでノンスキルの方でも解除できますが・・・)  
  
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
