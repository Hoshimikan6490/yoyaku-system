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
    let Calendar_URL = "https://calendar.google.com/calendar/~~";
  
   // FormsのURLを指定
    let Forms_URL = "https://forms.gle/~~";

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

   // 自動返信メール件名を変数「Subject」に代入
     let Subject ="【" + Name + "様へ】　工学院大学付属中学・高等学校　３Dプリンター　予約確認メール";
    
   // 自動返信メール本文を変数「Body」に代入
     let Body = 
     "---------------------------------------------" + "\n"
     + "※このメールは自動送信しておりますので、このメールにご返信いただいてもお答えできませんのでご了承ください。"+ "\n"
     + "※お心当たりがない場合は、メールの破棄をお願いいたします。" + "\n"
     + "---------------------------------------------" + "\n\n"
     + Name + "様"+"\n\n"
     + "工学院大学附属中学校・高等学校　デジタルクリエイター育成同好会　ものづくり班です。"+ "\n\n"
     + "このたびは、３Dプリンターの利用予約をいただき、誠にありがとうございます。" + "\n"
     + "下記URLの先に記載されている通り、カレンダーへの適応が完了しました。確認をお願い致します。"+ "\n"
     + "あなたの" + Yoyaku_day + "の予約順は、" + Ranking + "番目です。" + "\n\n"
     + Calendar_URL + "\n\n"
     + "＜注意事項＞"+ "\n"
     + "　※予約は、１日ごとに管理しております。そのため、お手数をおかけしますが、３Dプリンターを使うたびにご予約ください。" + "\n"
     + "　※このメールは送信専用です。返信は出来ません。" + "\n"
     + "　※取り消す際は、デジクリの３Dプリンターメンバー（放課後に３Dプリンターをいじっている人は大体そうです。笑）にご一報ください。"
     + "\n\n"
     + "---------------------------------------------"+ "\n"
     + "　ご回答された内容："+ "\n\n"
     + "　４桁番号：" + "　" + CLNO + "\n"
     + "　　お名前：" + "　" + Name + "\n"
     + "　　予約日：" + "　"+ Yoyaku_day + "\n"
     + "　　印刷物：" + "　" + About_Thing + "\n"
     + "　　　補助：" + "　" + Hojo + "\n"
     + "　　　備考：" + "　" + Other_info + "\n\n"
     + "　　Formsが送信された日時："　+ "　" + TimeStamp + "\n\n"
     + "============================================" + "\n"
     + "工学院大学附属中学校・高等学校　デジタルクリエイター育成同好会　ものづくり班" + "\n"
     + "============================================";
    
   // メール送信
     MailApp.sendEmail(Email,Subject,Body);

   // ログを書く
     console.log("予約番号：" + Ranking + "、　予約日：" + new Date(Yoyaku_day) + "に実行完了")

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
     "---------------------------------------------" + "\n"
     + "※このメールは自動送信しておりますので、このメールにご返信いただいてもお答えできませんのでご了承ください。"+ "\n"
     + "※お心当たりがない場合は、メールの破棄をお願いいたします。" + "\n"
     + "---------------------------------------------" + "\n\n"
     + Name + "様"+"\n\n"
     + "工学院大学附属中学校・高等学校　デジタルクリエイター育成同好会　ものづくり班です。"+ "\n\n"
     + "申し訳ございません。ご指定いただいた日程は、予約不可日に指定されているため、予約できません。" + "\n"
     + "ご不明な点がございましたら、デジクリモノづくり班までお問い合わせください。" + "\n\n"
     + "以下のURLから予約が空いている日程に再度ご予約をお願いいたします。"　+ "\n"
     + "お手数おかけして、申し訳ございません。" +"\n\n"
     + "空いている日程はこちらで確認↓"+ "\n"
     + Calendar_URL + "\n"
     + "再度回答する場合はこちらから↓"+ "\n"
     + Forms_URL + "\n\n"
     + "＜注意事項＞"+ "\n"
     + "　※このメールは送信専用です。返信は出来ません。" + "\n\n"
     + "============================================" + "\n"
     + "工学院大学附属中学校・高等学校　デジタルクリエイター育成同好会　ものづくり班" + "\n"
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
     "---------------------------------------------" + "\n"
     + "※このメールは自動送信しておりますので、このメールにご返信いただいてもお答えできませんのでご了承ください。"+ "\n"
     + "※お心当たりがない場合は、メールの破棄をお願いいたします。" + "\n"
     + "---------------------------------------------" + "\n\n"
     + Name + "様"+"\n\n"
     + "工学院大学附属中学校・高等学校　デジタルクリエイター育成同好会　ものづくり班です。"+ "\n\n"
     + "申し訳ございません。ご指定いただいた日程は、定員の人数を満たしたため、予約できません。" + "\n"
     + "以下のURLから予約が空いている日程に再度ご予約をお願いいたします。"　+ "\n"
     + "お手数おかけして、申し訳ございません。" +"\n\n"
     + "空いている日程はこちらで確認↓"+ "\n"
     + Calendar_URL + "\n"
     + "再度回答する場合はこちらから↓"+ "\n"
     + Forms_URL + "\n\n"
     + "＜注意事項＞"+ "\n"
     + "　※このメールは送信専用です。返信は出来ません。" + "\n\n"
     + "============================================" + "\n"
     + "工学院大学附属中学校・高等学校　デジタルクリエイター育成同好会　ものづくり班" + "\n"
     + "============================================";

   // メール送信
     MailApp.sendEmail(Email,Subject,Body);

   // ログを書く
      console.log("定員オーバーの予約不可メール送信完了")
   }
 // メインプロセス終了
}
