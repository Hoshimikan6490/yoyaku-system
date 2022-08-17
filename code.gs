function onFormSubmit(e) {
 //　手動初期設定
   // カレンダーIDを指定
    let Calendar_ID =　"○○@group.calendar.google.com";
   // GoogleカレンダーのURLを指定
    let Calendar_URL = "https://calendar.google.com/calendar/～～～";
   // FormsのURLを指定
    let Forms_URL = "https://forms.gle/～～～";
   // 登録用スプレッドシートの統計データが記録されたシートの名前を設定
     let Database_sheet = "○○○○";
 // 手動初期設定完了


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
   //  Mainprogramが実行されたときに、Ranking変数を持ってきて、「順位：名前」のタイトルの終日イベントをYoyaku_dayに作成し、詳細欄には、descriptionの中身を入れる
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
         console.log("予約番号：" + Yoyaku_NO + "、　予約日：" + Yoyaku_day + "に実行完了");
      // ログを書く
         console.log("予約番号設定完了")
     
   // 予約番号を別シートにコピー
     // このプログラムが所属するスプレッドシートのうち、「統計データ」という名前のシートを取得
       let Yoyaku_list_sheet = Yoyaku_Spreadsheet.getSheetByName(Database_sheet);
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
     // HTMLを取得
       doGet_success();
     // 取得したHTMLに変数を代入
       let html = HtmlService.createHtmlOutputFromFile("success").getContent();
       // 取得したHTMLの一番最初の「Name」を変数の「Name」の中身に置き換え（以降同様）
       let html2 = html.replace("Name", Name);
       let html3 = html2.replace("Calender_URL", Calendar_URL);
       let html4 = html3.replace("Ranking", Ranking);
       let html5 = html4.replace("Yoyaku_NO", Yoyaku_NO);
       let html6 = html5.replace("CLNO", CLNO);
       let html7 = html6.replace("Name", Name);
       let html8 = html7.replace("Yoyaku_day", Yoyaku_day);
       let html9 = html8.replace("About_Thing", About_Thing);
       let html10 = html9.replace("Hojo", Hojo);
       let html11 = html10.replace("Other_info", Other_info);
       let html_perfect = html11.replace("TimeStamp", TimeStamp);
     // ログを書く
       console.log(html_perfect)
     // 件名を変数「Subject」に代入
       let Subject ="【" + Name + "様へ】　３Dプリンター　予約確認メール";
     // bodyは何も書く必要が無いので、無し（ただし、これがないと送信できないので、変数は設定）
       let Body = "";
     // オプションの設定
       let Options ={
         // 送信者を「３Dプリンター　管理者」にする
         "name": "３Dプリンター　管理者",
         // メールの本文に変数代入の終わったHTMLデータを挿入
         "htmlBody": html_perfect
         };     
     // メール送信
       MailApp.sendEmail(Email,Subject,Body,Options);
     // ログを書く
       console.log(Email + "に確認メールを送信完了")

   //Mainprogramの終了を示す
     return Ranking;
   }
 // 「Mainprogram」が実行されたときに動く内容を設定終了


 // メインプロセス
   // 予約日に【予約不可】があるときに
     if (Calendar.getEventsForDay(new Date(Yoyaku_day), {search: '【予約不可】'}).length) {
     // 確認メール送信
       // HTMLを取得
         doGet_canot();
       // 取得したHTMLに変数を代入
         let html = HtmlService.createHtmlOutputFromFile("canot").getContent();
       // 取得したHTMLの一番最初の「Name」を変数の「Name」の中身に置き換え（以降同様）
         let html2 = html.replace("Name", Name);
         let html3 = html2.replace("Calendar_URL", Calendar_URL);
         let html_perfect = html3.replace("Forms_URL", Forms_URL);
       // ログを書く
         console.log(html_perfect)
       // 自動返信メール件名を変数「Subject」に代入
         let Subject ="【" + Name + "様へ】　予約失敗！！！";
       // bodyは何も書く必要が無いので、無し（ただし、これがないと送信できないので、変数は設定）
         let Body = "";

       // 追加設定
         let Options ={
         // 送信者を「３Dプリンター　管理者」にする
          "name": "３Dプリンター　管理者",
         // メールの本文に変数代入の終わったHTMLデータを挿入
          "htmlBody": html_perfect
         };
                 
       // メール送信
         MailApp.sendEmail(Email,Subject,Body,Options);

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

   // 「予約不可」も①もないなら、Yoyaku_dayに②で始まるイベントがないときに、
     } else if (!Calendar.getEventsForDay(new Date(Yoyaku_day), {search: '②'}).length) {
       // 変数「Ranking」を②に設定
         let Ranking = '②';
       // 変数「Ranking」を用いて、「Mainprogram」を実行
         Mainprogram(Ranking);
       // ログを書く
         console.log("予約完了メール送信完了")

   // 「予約不可」も①も②もないなら、Yoyaku_dayに③で始まるイベントがないときに、
     } else if (!Calendar.getEventsForDay(new Date(Yoyaku_day), {search: '③'}).length) {
       // 変数「Ranking」を③に設定
         let Ranking = '③';
       // 変数「Ranking」を用いて、「Mainprogram」を実行
         Mainprogram(Ranking);
       // ログを書く
         console.log("予約完了メール送信完了")
      
   // 「予約不可」も①も②も③もないなら、Yoyaku_dayに④で始まるイベントがないときに、
     } else if (!Calendar.getEventsForDay(new Date(Yoyaku_day), {search: '④'}).length) {
       // 変数「Ranking」を④に設定
         let Ranking = '④';
       // 変数「Ranking」を用いて、「Mainprogram」を実行
         Mainprogram(Ranking);
       // ログを書く
         console.log("予約完了メール送信完了")

   // 既に④まで予約が入っているときに
     } else {
       // HTMLを取得
         doGet_full();
       // 取得したHTMLに変数を代入
         let html = HtmlService.createHtmlOutputFromFile("full").getContent();
       // 取得したHTMLの一番最初の「Name」を変数の「Name」の中身に置き換え（以降同様）
         let html2 = html.replace("Name", Name);
         let html3 = html2.replace("Calendar_URL", Calendar_URL);
         let html_perfect = html3.replace("Forms_URL", Forms_URL);
       //　ログを書く
         console.log(html_perfect)
       // 自動返信メール件名を変数「Subject」に代入
         let Subject ="【" + Name + "様へ】　予約失敗！！！";
       // bodyは何も書く必要が無いので、無し（ただし、これがないと送信できないので、変数は設定）
         let Body = "";
       // 追加設定
         let Options ={
           // 送信者を「３Dプリンター　管理者」にする
           "name": "３Dプリンター　管理者",
           // メールの本文に変数代入の終わったHTMLデータを挿入
           "htmlBody": html_perfect
          };
       // メール送信
         MailApp.sendEmail(Email,Subject,Body,Options);
       // ログを書く
         console.log("定員オーバーの予約不可メール送信完了")
     }
 // メインプロセス終了
}


function doGet_success() {
 // 「doGet_success」が実行されたときに動く内容を設定
   var t = HtmlService.createTemplateFromFile('success.html');
   return t.evaluate();
}

function doGet_canot() {
 // 「doGet_canot」が実行されたときに動く内容を設定
   var t = HtmlService.createTemplateFromFile('canot.html');
   return t.evaluate();
}

function doGet_full() {
 // 「doGet_full」が実行されたときに動く内容を設定
   var t = HtmlService.createTemplateFromFile('full.html');
   return t.evaluate();
}
