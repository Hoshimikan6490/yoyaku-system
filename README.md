# ３Dプリンター予約システム

## 概要
　余裕があるときに、書いておきます。

## 制作手順
### 大まかな流れ
1. Formsの内容をコピー
 2. ①のFormsからExcelを開く
 3. Excelの列の順番を「タイムスタンプ、４桁番号、氏名、メールアドレス、予約日、補助の必要性、備考」の順番に変更する。
※その他は備考欄よりも右側においてあればOK
 5. 拡張機能＞app scriptから、エディターを起動
 ⑤「コード.gs」にプログラムをコピー
  ※既存の内容を書き換える
 ⑥Googleの共有カレンダーを作成
 ⑦Googleカレンダーのタイムゾーン設定を「Asia/Tokyo」に設定
 ⑧共有カレンダーのカレンダーIDを設定
 ⑨Apps Scriptの設定から、「『appsscript.json』マニフェスト ファイルをエディタで表示する」をONにする。
 ⑩「appsscript.json」とGoogleカレンダーの標準時刻を”Asia/Tokyo”にする。
 ⑪以下の設定でトリガーを設定して、保存
  - 実行する関数：onFormSubmit
  - 実行するデプロイ：Head
  - イベントのソース：スプレッドシートから
  - イベントの種類：フォーム送信時
  - エラー通知：１週間おきに通知を受け取る

### 詳細な制作方法
　仕様上、ここで説明するのは困難なため、[こちら](https://docs.google.com/presentation/d/1Z6X8r1lDuS5SBVQIaswEVqARF-M6jq_h/edit?usp=sharing&ouid=106480420577465092683&rtpof=true&sd=true)をご参照ください。
