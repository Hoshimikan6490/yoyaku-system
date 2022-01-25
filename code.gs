/////////////////////////////////////////////////////////////////////////////////////////
// ★本番バージョン制作手順
//　　①Formsの内容をコピー
//　　②①のFormsからExcelを開く
//　　③Excelの列の順番を「タイムスタンプ、４桁番号、氏名、メールアドレス、予約日、補助の必要性、備考」の順番に変更する。
//　　　　※その他は備考欄よりも右側においてあればOK
//　　④拡張機能＞app scriptから、エディターを起動
//　　⑤「コード.gs」にプログラムをコピー
//　　　　※既存の内容を書き換える
//　　⑥Googleの共有カレンダーを作成
//　　⑦Googleカレンダーのタイムゾーン設定を「Asia/Tokyo」に設定
//　　⑧共有カレンダーのカレンダーIDを設定
//　　⑨Apps Scriptの設定から、「『appsscript.json』マニフェスト ファイルをエディタで表示する」をONにする。
//　　⑩「appsscript.json」とGoogleカレンダーの標準時刻を”Asia/Tokyo”にする。
//　　⑪以下の設定でトリガーを設定して、保存
// 　　　・実行する関数：onFormSubmit
// 　　　・実行するデプロイ：Head
// 　　　・イベントのソース：スプレッドシートから
// 　　　・イベントの種類：フォーム送信時
// 　　　・エラー通知：１週間おきに通知を受け取る
/////////////////////////////////////////////////////////////////////////////////////////

function onFormSubmit(e) {
 //　手動初期設定
   // カレンダーIDを指定
    let Calendar_ID = "○○@group.calendar.google.com";

   // GoogleカレンダーのURLを指定
    let Calendar_URL = "https://calendar.google.com/calendar/～～～";
  
   // FormsのURLを指定
    let Forms_URL = "https://forms.gle/～～～";

 // 自動初期設定
   // タイムスタンプを変数「TimeStamp」に代入
     let TimeStamp = e.values[0];

   // ４桁番号を変数「CLNO」に代入
     let CLNO = e.values[1];
    
   // 氏名を変数「Name」に代入
     let Name = e.values[2];
    
   // メールアドレスを変数「Email」に代入
     let Email = e.values[3];
    
   // 予約日を変数「Yoyaku_day」に代入
     let Yoyaku_day = e.values[4];
    
   // 予約物の詳細を変数「About_Thing」に代入
     let About_Thing = e.values[5];
    
   //　補助についてを変数「Hojo」に代入
     let Hojo = e.values[6];
    
   // 備考などを変数「Other_info」に代入
     let Other_info = e.values[7];

   // 順位設定用の変数を作成
     let Ranking = "error";
    
   // ログを書く
     console.log("変数設定完了");

   // カレンダーオブジェクトを取得(赤い文字がカレンダーID。これは予定を入れる先のカレンダーIDを手入力)
   //  ※「Calendar」が実行されたときに動く内容を設定
     let Calendar = CalendarApp.getCalendarById(Calendar_ID);

   // タイムゾーン設定
     Calendar.setTimeZone("Asia/Tokyo");

   // ログを書く
     console.log("初期設定完了");

  //初期設定終了



  // 「Mainprogram」が実行されたときに動く内容を設定
  //  ※「\n」は、改行を表しています。詳細は、こちら
  //   　→ https://www.javadrive.jp/javascript/string/index3.html#section2
   //  Mainprogramが実行されたときに、Ranking変数を持ってきて、「順位：名前」のタイトルの終日イベントを
   // Yoyaku_dayに作成し、詳細欄には、descriptionの中身を入れる
     let Mainprogram = function (Ranking) { Calendar.createAllDayEvent( Ranking + ":" 
     + Name, new Date(Yoyaku_day), {description: "Formsを入力した日時：" + TimeStamp + "\n" 
     + "印刷物：" + About_Thing + "\n" + "補助の必要性：" + Hojo + "\n" + "備考：" + Other_info});   
    
   // イベントの色を設定
     // イベントを取得
        let Newevent = Calendar.getEventsForDay(new Date (Yoyaku_day), {search: Ranking});
     // 色設定
        Newevent[0].setColor('2');

   // 予約番号を書く
     // 予約番号を決定
       //このプログラムが所属するスプレッドシートのうち、「フォームの回答　1」という名前のシートを取得
         let Yoyaku_Spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
         let Yoyaku_Sheet = Yoyaku_Spreadsheet.getSheetByName("フォームの回答 1");
       //１列目の最終行を取得
         let Old_yoyaku_NO_plas = Yoyaku_Sheet.getRange(Yoyaku_Sheet.getMaxRows(), 1).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
     　// １列目の最終行の１つ上の行番号を取得(今回予約するところ)
         let Yoyaku_NO =Old_yoyaku_NO_plas -1
       // ログを書く
        console.log("予約番号設定完了")
     
   // 予約番号を別シートにコピー
     // このプログラムが所属するスプレッドシートのうち、「統計データ」という名前のシートを取得
       let Yoyaku_list_sheet = Yoyaku_Spreadsheet.getSheetByName("統計データ");
     // 予約番号を統計データシートのＡ列に代入
       Yoyaku_list_sheet.getRange(Yoyaku_NO,1).setValue(Yoyaku_NO);
     // 予約日を統計データシートのB列に代入
       // 予約日を「YY/MM/DD」の形から「YY年MM月DD日」の形に変換
         // 「/」を境目にして、一番左側の文字列を変数「Yoyaku_year」に代入
           let Yoyaku_year = Yoyaku_day.split("/")[0];
         // 「/」を境目にして、中央の文字列を変数「Yoyaku_month」に代入
           let Yoyaku_month = Yoyaku_day.split("/")[1]
         // 「/」を境目にして、一番右側の文字列を変数「Yoyaku_day2」に代入
           let Yoyaku_day2 = Yoyaku_day.split("/")[2];
         // 「YY年MM月DD日」の形に変換して、変数「Moji_yoyaku_day」に代入
           let Moji_yoyaku_day = Yoyaku_year + "年" + Yoyaku_month + "月" + Yoyaku_day2 + "日"
         // 予約日を統計データシートのB列に代入
           Yoyaku_list_sheet.getRange(Yoyaku_NO,2).setValue(Moji_yoyaku_day);
       // 氏名を統計データシートのC列に代入
         Yoyaku_list_sheet.getRange(Yoyaku_NO,3).setValue(Name);
       // メールアドレスを統計データシートのD列に代入
         Yoyaku_list_sheet.getRange(Yoyaku_NO,4).setValue(Email);
       // Rankingを統計データシートのE列に代入
         Yoyaku_list_sheet.getRange(Yoyaku_NO,5).setValue(Ranking);
     // ログを書く
       console.log("予約情報のコピー完了") 

   // 確認メール送信
     // 自動返信メール件名を変数「Subject」に代入
       let Subject ="【" + Name + "様へ】　工学院大学付属中学・高等学校　３Dプリンター　予約確認メール";
    
     // 自動返信メール本文を変数「Body」に代入
       let Body = 
         Name + "様" +"\n"
         + "\n"
         + "３Dプリンター　管理者です。" + "\n"
         + "\n"
         + "このたびは、３Dプリンターの利用予約をいただき、誠にありがとうございます。" + "\n"
         + "下記URLの先に記載されている通り、カレンダーへの適応が完了しました。確認をお願い致します。" + "\n"
         + "\n"
         + "★予約の確認はこちら↓★" + "\n"
         + Calendar_URL + "\n"
         + "\n"
         + "---------------------------------------------"+ "\n"
         + "＜予約の詳細＞" + "\n"
         + "\n"
         + "　予約順：" + "　" + Ranking + "\n"
         + "　予約番号：" + "　" + Yoyaku_NO + "\n"
         + "\n"
         + "※予約番号は予約をキャンセルする際に本人確認で使用しますので、厳重に保管してください。" + "\n"
         + "\n"
         + "---------------------------------------------"+ "\n"
         + "＜注意事項＞"+ "\n"
         + " ※お心当たりがない場合は、メールの破棄をお願いいたします。" + "\n"
         + "　※予約は、１日ごとに管理しております。そのため、お手数をおかけしますが、３Dプリンターを使うたびにご予約ください。" + "\n"
         + "　※取り消す際は、専用の「キャンセル用Forms」をご利用ください。" + "\n"
         + "　※何かご不明な点がございましたら、このメールアドレスまでご返信ください。" +"\n"
         + "\n"
         + "---------------------------------------------"+ "\n"
         + "　ご回答された内容：" + "\n"
         + "\n"
         + "　４桁番号：" + "　" + CLNO + "\n"
         + "　　お名前：" + "　" + Name + "\n"
         + "　　予約日：" + "　"+ Yoyaku_day + "\n"
         + "　　印刷物：" + "　" + About_Thing + "\n"
         + "　　　補助：" + "　" + Hojo + "\n"
         + "　　　備考：" + "　" + Other_info + "\n"
         + "\n"
         + "　　Formsが送信された日時："　+ "　" + TimeStamp + "\n"
         + "\n"
         + "============================================" + "\n"
         + "３Dプリンター　管理者" + "\n"
         + "============================================";
    
     // メール送信
       MailApp.sendEmail(Email,Subject,Body);

     // ログを書く
       console.log("予約番号：" + Yoyaku_NO + "、　予約日：" + Yoyaku_day + "に実行完了")

   //Mainprogramの終了を示す
     return Ranking;
    
   }
 // 「Mainprogram」が実行されたときに動く内容を設定終了



 // メインプロセス
   // 予約日に【予約不可】があるときに
     if (Calendar.getEventsForDay(new Date(Yoyaku_day), {search: '【予約不可】'}).length) {
  
   // 自動返信メール件名を変数「Subject」に代入
     let Subject ="【" + Name + "様へ】　予約失敗！！！";
    
   // 自動返信メール本文を変数「Body」に代入
     let Body = 
       Name + "様" +"\n"
       + "\n"
       + "３Dプリンター　管理者です。" + "\n"
       + "\n"
       + "申し訳ございません。ご指定いただいた日程は、予約不可日に指定されているため、予約できません。" + "\n"
       + "ご不明な点がございましたら、デジクリモノづくり班までお問い合わせください。" + "\n"
       + "\n"
       + "以下のURLから予約が空いている日程に再度ご予約をお願いいたします。"　+ "\n"
       + "お手数おかけして、申し訳ございません。" +"\n"
       + "\n"
       + "空いている日程はこちらで確認↓"+ "\n"
       + Calendar_URL + "\n"
       + "再度回答する場合はこちらから↓"+ "\n"
       + Forms_URL + "\n"
       + "\n"
       + "---------------------------------------------" + "\n"
       + "＜注意事項＞"+ "\n"
       + " ※お心当たりがない場合は、メールの破棄をお願いいたします。" + "\n"
       + " ※何かご不明な点がございましたら、このメールアドレスまでご返信ください。" + "\n"
       + "\n"
       + "============================================" + "\n"
       + "３Dプリンター　管理者" + "\n"
       + "============================================";

   // メール送信
     MailApp.sendEmail(Email,Subject,Body);

   // ログを書く
       console.log("予約不可日通知メール送信完了")


 // 「予約不可がない」なら、
   // Yoyaku_dayに①で始まるイベントがないときに、
     } else if (!Calendar.getEventsForDay(new Date(Yoyaku_day), {search: '①'}).length) {

   // 変数「Ranking」を①に設定
     let Ranking = '①';

   // 変数「Ranking」を用いて、「Mainprogram」を実行
     Mainprogram(Ranking);

   // ログを書く
      console.log("予約完了メール送信完了")
      

 // 「予約不可」も①もないなら、
   // Yoyaku_dayに②で始まるイベントがないときに、
   } else if (!Calendar.getEventsForDay(new Date(Yoyaku_day), {search: '②'}).length) {

   // 変数「Ranking」を②に設定
     let Ranking = '②';

   // 変数「Ranking」を用いて、「Mainprogram」を実行
     Mainprogram(Ranking);

   // ログを書く
      console.log("予約完了メール送信完了")


 // 「予約不可」も①も②もないなら、
   // Yoyaku_dayに③で始まるイベントがないときに、
   } else if (!Calendar.getEventsForDay(new Date(Yoyaku_day), {search: '③'}).length) {

   // 変数「Ranking」を③に設定
     let Ranking = '③';

   // 変数「Ranking」を用いて、「Mainprogram」を実行
     Mainprogram(Ranking);

   // ログを書く
      console.log("予約完了メール送信完了")
      

 // 「予約不可」も①も②も③もないなら、
   // Yoyaku_dayに④で始まるイベントがないときに、
   } else if (!Calendar.getEventsForDay(new Date(Yoyaku_day), {search: '④'}).length) {

   // 変数「Ranking」を④に設定
     let Ranking = '④';

   // 変数「Ranking」を用いて、「Mainprogram」を実行
     Mainprogram(Ranking);

   // ログを書く
      console.log("予約完了メール送信完了")


 // 既に④まで予約が入っているときに
   } else {
   // 自動返信メール件名を変数「Subject」に代入
     let Subject ="【" + Name + "様へ】　予約失敗！！！";
   
   // 自動返信メール本文を変数「Body」に代入
     let Body = 
       Name + "様" +"\n"
       + "\n"
       + "３Dプリンター　管理者です。" + "\n"
       + "\n"
       + "申し訳ございません。ご指定いただいた日程は、定員の人数を満たしたため、予約できません。" + "\n"
       + "以下のURLから予約が空いている日程に再度ご予約をお願いいたします。"　+ "\n"
       + "お手数おかけして、申し訳ございません。" +"\n"
       + "\n"
       + "空いている日程はこちらで確認↓"+ "\n"
       + Calendar_URL + "\n"
       + "再度回答する場合はこちらから↓"+ "\n"
       + Forms_URL + "\n"
       + "\n"
       + "---------------------------------------------" + "\n"
       + "＜注意事項＞"+ "\n"
       + " ※お心当たりがない場合は、メールの破棄をお願いいたします。" + "\n"
       + "　※何かご不明な点がございましたら、このメールアドレスまでご返信ください。" + "\n"
       + "\n"
       + "============================================" + "\n"
       + "３Dプリンター　管理者" + "\n"
       + "============================================";

   // メール送信
     MailApp.sendEmail(Email,Subject,Body);

   // ログを書く
      console.log("定員オーバーの予約不可メール送信完了")
   }
 // メインプロセス終了
}
