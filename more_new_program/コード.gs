function onFormSubmit(e) {
  //　手動初期設定 /////////////////////////////////////////////////////////////////////////////////
  // カレンダーIDを指定
  let Calendar_ID = "○○○○@group.calendar.google.com";
  // GoogleカレンダーのURLを指定
  let Calendar_URL = "https://calendar.google.com/calendar/○○○○";
  // FormsのURLを指定
  let Forms_URL = "https://forms.gle/○○○○";
  // 登録用スプレッドシートの統計データが記録されたシートの名前を設定
  let Database_sheet = "データベース";
  // １日当たりの予約数の最大値を設定
  let Limit = 4;
  // 予約システムのオーナーのメールアドレスで予約できないようにする（trueの場合は出来なくなり、falseの場合は出来るようになります。）
  let Owner_cannnot_reserve = true;
  // 手動初期設定完了 /////////////////////////////////////////////////////////////////////////////

  // 自動初期設定 /////////////////////////////////////////////////////////////////////////////////
  // タイムスタンプを変数「TimeStamp」に代入
  let TimeStamp = e.values[0];
  // ４桁番号を変数「CLNO」に代入
  let CLNO = e.values[1];
  // 氏名を変数「Name」に代入
  let Name = e.values[2];
  // メールアドレスを変数「Email」に代入
  let Email = e.values[3];

  // 予約するかキャンセルするかを変数「Choice」に代入
  let Choice = e.values[4];

  // 予約日を変数「Yoyaku_day」に代入
  let Yoyaku_day = e.values[6];
  // 予約物の詳細を変数「About_Thing」に代入
  let About_Thing = e.values[7];
  //　補助についてを変数「Hojo」に代入
  let Hojo = e.values[8];
  // 備考などを変数「Other_info」に代入
  let Other_info = e.values[9];

  // キャンセルする日を変数「Cancel_day」に代入
  let Cancel_day = e.values[10];
  // キャンセルしたい予約番号を変数「Cancel_NO」に代入
  let Cancel_NO = e.values[11];
  // ログを書く
  console.log("変数設定完了");

  // カレンダーオブジェクトを取得(赤い文字がカレンダーID。これは予定を入れる先のカレンダーIDを手入力)
  //  ※「Calendar」が実行されたときに動く内容を設定
  let Calendar = CalendarApp.getCalendarById(Calendar_ID);

  // 基本Spreadsheetの取得を取得
  // このプログラムに接続されているSpreadsheetに接続
  let Yoyaku_Spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // データベース用のシートを取得
  let Yoyaku_list_sheet = Yoyaku_Spreadsheet.getSheetByName(Database_sheet);

  // タイムゾーン設定
  Calendar.setTimeZone("Asia/Tokyo");

  // ログを書く
  console.log("初期設定完了");
  //初期設定終了 /////////////////////////////////////////////////////////////////////////////

  //予約関係のプログラム　//////////////////////////////////////////////////////////////////////
  let Yoyaku_program = function (Yoyaku_day) {
    // 「Mainprogram」が実行されたときに動く内容を設定
    //  Mainprogramが実行されたときに、Ranking変数を持ってきて、「順位：名前」のタイトルの終日イベントをYoyaku_dayに作成し、詳細欄には、descriptionの中身を入れる
    let Mainprogram = function (Ranking) {
      Calendar.createAllDayEvent(Ranking + ":" + Name, new Date(Yoyaku_day), {
        description:
          "Formsを入力した日時：" +
          TimeStamp +
          "\n" +
          "印刷物：" +
          About_Thing +
          "\n" +
          "補助の必要性：" +
          Hojo +
          "\n" +
          "備考：" +
          Other_info,
      });

      // イベントの色を設定
      // イベントを取得
      let Newevent = Calendar.getEventsForDay(new Date(Yoyaku_day), {
        search: Ranking,
      });
      // 色設定
      Newevent[0].setColor("2");

      // 予約番号を書く
      // 予約番号を決定
      //「フォームの回答　1」という名前のシートを取得
      let Yoyaku_Sheet = Yoyaku_Spreadsheet.getSheetByName("フォームの回答 1");
      //１列目の最終行を取得
      let Old_yoyaku_NO_plas = Yoyaku_Sheet.getRange(
        Yoyaku_Sheet.getMaxRows(),
        1
      )
        .getNextDataCell(SpreadsheetApp.Direction.UP)
        .getRow();
      // １列目の最終行の１つ上の行番号を取得(今回予約するところ)
      let Yoyaku_NO = Old_yoyaku_NO_plas - 1;
      // ログを書く
      console.log(
        "予約番号：" + Yoyaku_NO + "、 予約日：" + Yoyaku_day + "に実行完了"
      );
      // ログを書く
      console.log("予約番号設定完了");

      // 予約番号を別シートにコピー
      // 予約番号を統計データシートのＡ列に代入
      Yoyaku_list_sheet.getRange(Yoyaku_NO, 1).setValue(Yoyaku_NO);
      // 予約日を統計データシートのB列に代入
      // 予約日を「YY/MM/DD」の形から「YY年MM月DD日」の形に変換
      // 「/」を境目にして、一番左側の文字列を変数「Yoyaku_year」に代入
      let Yoyaku_year = Yoyaku_day.split("/")[0];
      // 「/」を境目にして、中央の文字列を変数「Yoyaku_month」に代入
      let Yoyaku_month = Yoyaku_day.split("/")[1];
      // 「/」を境目にして、一番右側の文字列を変数「Yoyaku_day2」に代入
      let Yoyaku_day2 = Yoyaku_day.split("/")[2];
      // 「YY年MM月DD日」の形に変換して、変数「Moji_yoyaku_day」に代入
      let Moji_yoyaku_day =
        Yoyaku_year + "年" + Yoyaku_month + "月" + Yoyaku_day2 + "日";
      // 予約日を統計データシートのB列に代入
      Yoyaku_list_sheet.getRange(Yoyaku_NO, 2).setValue(Moji_yoyaku_day);
      // 氏名を統計データシートのC列に代入
      Yoyaku_list_sheet.getRange(Yoyaku_NO, 3).setValue(Name);
      // メールアドレスを統計データシートのD列に代入
      Yoyaku_list_sheet.getRange(Yoyaku_NO, 4).setValue(Email);
      // Rankingを統計データシートのE列に代入
      Yoyaku_list_sheet.getRange(Yoyaku_NO, 5).setValue(Ranking);
      // ログを書く
      console.log("予約情報のコピー完了");

      // 確認メール送信
      // HTMLを取得
      doGet_yoyaku_success();
      // 取得したHTMLに変数を代入
      let html =
        HtmlService.createHtmlOutputFromFile("yoyaku_success").getContent();
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
      console.log(html_perfect);
      // 件名を変数「Subject」に代入
      let Subject = "【" + Name + "様へ】 3Dプリンター 予約確認メール";
      // bodyは何も書く必要が無いので、無し（ただし、これがないと送信できないので、変数は設定）
      let Body = "";
      // オプションの設定
      let Options = {
        // 送信者を「3Dプリンター 管理者」にする
        name: "3Dプリンター 管理者",
        // メールの本文に変数代入の終わったHTMLデータを挿入
        htmlBody: html_perfect,
      };
      // メール送信
      MailApp.sendEmail(Email, Subject, Body, Options);
      // ログを書く
      console.log(Email + "に確認メールを送信完了");

      //Mainprogramの終了を示す
      return Ranking;
    };
    // 「Mainprogram」が実行されたときに動く内容を設定終了

    // メインプロセス
    // 予約に使ったメールアドレスがこのプログラムを実行してるユーザーだった場合
    if (Owner_cannnot_reserve) {
      let ownerID = Session.getActiveUser().getEmail();
      if (ownerID == Email) {
        // HTMLを取得
        doGet_canot();

        // エラー理由を設定
        let Reason =
          "ご利用されたEメールアドレスは、このフォームのオーナーの物である";

        // 取得したHTMLに変数を代入
        let html = HtmlService.createHtmlOutputFromFile("canot").getContent();
        // 取得したHTMLの一番最初の「Name」を変数の「Name」の中身に置き換え（以降同様）
        let html2 = html.replace("Name", Name);
        let html3 = html2.replace("reason", Reason);
        let html4 = html3.replace("Calendar_URL", Calendar_URL);
        let html_perfect = html4.replace("Forms_URL", Forms_URL);

        // ログを書く
        console.log(html_perfect);

        // 自動返信メール件名を変数「Subject」に代入
        let Subject = "【" + Name + "様へ】 予約失敗！！！";

        // bodyは何も書く必要が無いので、無し（ただし、これがないと送信できないので、変数は設定）
        let Body = "";

        // 追加設定
        let Options = {
          // 送信者を「3Dプリンター 管理者」にする
          name: "3Dプリンター 管理者",
          // メールの本文に変数代入の終わったHTMLデータを挿入
          htmlBody: html_perfect,
        };

        // メール送信
        MailApp.sendEmail(Email, Subject, Body, Options);

        // ログを書く
        console.log("Eメールアドレスエラー通知メール送信完了");
        return;
      }
    }

    // 予約日に【予約不可】があるときに
    if (
      Calendar.getEventsForDay(new Date(Yoyaku_day), { search: "【予約不可】" })
        .length
    ) {
      // 確認メール送信
      // HTMLを取得
      doGet_yoyaku_canot();
      // エラー理由を設定
      let Reason = "ご指定いただいた日程は、予約不可日に指定されている";
      // 取得したHTMLに変数を代入
      let html =
        HtmlService.createHtmlOutputFromFile("yoyaku_canot").getContent();
      // 取得したHTMLの一番最初の「Name」を変数の「Name」の中身に置き換え（以降同様）
      let html2 = html.replace("Name", Name);
      let html3 = html2.replace("reason", Reason);
      let html4 = html3.replace("Calendar_URL", Calendar_URL);
      let html_perfect = html4.replace("Forms_URL", Forms_URL);
      // ログを書く
      console.log(html_perfect);
      // 自動返信メール件名を変数「Subject」に代入
      let Subject = "【" + Name + "様へ】 予約失敗！！！";
      // bodyは何も書く必要が無いので、無し（ただし、これがないと送信できないので、変数は設定）
      let Body = "";

      // 追加設定
      let Options = {
        // 送信者を「3Dプリンター　管理者」にする
        name: "3Dプリンター 管理者",
        // メールの本文に変数代入の終わったHTMLデータを挿入
        htmlBody: html_perfect,
      };

      // メール送信
      MailApp.sendEmail(Email, Subject, Body, Options);

      // ログを書く
      console.log("予約不可日通知メール送信完了");
    } else {
      if (Calendar.getEventsForDay(new Date(Yoyaku_day)).length !== Limit) {
        for (Count = 1; Count < Limit + 1; Count++) {
          let Count_character = String(Count);
          let events = Calendar.getEventsForDay(new Date(Yoyaku_day), {
            search: Count_character,
          });
          if (!events.length) {
            events.filter(function (event) {
              if (!event.getTitle().startsWith(Count_character)) return;
            });
            // Limit数を予約順に変換
            let Ranking = Count_character;
            // 変数「Ranking」を用いて、「Mainprogram」を実行
            Mainprogram(Ranking);
            // ログを書く
            console.log("予約完了メール送信完了");
            // ループを切る
            break;
          }
        }
      } else {
        // HTMLを取得
        doGet_yoyaku_full();
        // 取得したHTMLに変数を代入
        let html =
          HtmlService.createHtmlOutputFromFile("yoyaku_full").getContent();
        // 取得したHTMLの一番最初の「Name」を変数の「Name」の中身に置き換え（以降同様）
        let html2 = html.replace("Name", Name);
        let html3 = html2.replace("Calendar_URL", Calendar_URL);
        let html_perfect = html3.replace("Forms_URL", Forms_URL);
        //　ログを書く
        console.log(html_perfect);
        // 自動返信メール件名を変数「Subject」に代入
        let Subject = "【" + Name + "様へ】 予約失敗！！！";
        // bodyは何も書く必要が無いので、無し（ただし、これがないと送信できないので、変数は設定）
        let Body = "";
        // 追加設定
        let Options = {
          // 送信者を「3Dプリンター　管理者」にする
          name: "3Dプリンター 管理者",
          // メールの本文に変数代入の終わったHTMLデータを挿入
          htmlBody: html_perfect,
        };
        // メール送信
        MailApp.sendEmail(Email, Subject, Body, Options);
        // ログを書く
        console.log("定員オーバーの予約不可メール送信完了");
      }
    }
  };
  //　メインプロセス終了
  //予約関係のプログラム終了////////////////////////////////////////////////////////////////////

  //キャンセル関係のプログラム///////////////////////////////////////////////////////////////////
  let Cancel_program = function (Cancel_NO) {
    // 指定された予約番号に当てはまる統計データシートの予約番号を変数「database_Cancel_NO」に代入
    let Database_Cancel_NO = Yoyaku_list_sheet.getRange(
      Cancel_NO,
      1
    ).getValue();
    // 指定された予約番号に当てはまる統計データシートのメールアドレスを変数「Database_cancel_Email」に代入
    let Database_cancel_Email = Yoyaku_list_sheet.getRange(
      Cancel_NO,
      4
    ).getValue();
    // 指定された予約番号に当てはまる統計データシートの予約順を変数「Database_cancel_Ranking」に代入
    let Database_cancel_Ranking = Yoyaku_list_sheet.getRange(
      Cancel_NO,
      5
    ).getValue();
    //ログを書く
    console.log("登録用Formsに入っている予約番号：" + Database_Cancel_NO);
    console.log(
      "登録用Formsに入っているメールアドレス：" + Database_cancel_Email
    );
    console.log("キャンセルされた予約の予約順：" + Database_cancel_Ranking);

    // OKパターン
    let Can_Cleaning = function (Name) {
      let Database_cancel_Ranking_character = String(Database_cancel_Ranking);
      let Find_events = Calendar.getEventsForDay(new Date(Cancel_day), {
        search: Database_cancel_Ranking_character,
      });
      // もし、キャンセル日に予約者の名前があったときに、
      if (Find_events.length) {
        for (const event of Find_events) {
          //イベントタイトルにRankingの数字が含まれていれば、削除
          const eventTitle = event.getTitle();
          if (eventTitle.startsWith(Database_cancel_Ranking) !== -1) {
            event.deleteEvent();
          }
        }
        // ログを書く
        console.log(
          Name +
            "さんの" +
            Cancel_day +
            "の" +
            Database_cancel_Ranking +
            "番の予約を削除完了"
        );
        //　自動返信メール
        // HTMLを取得
        doGet_cancel_success();
        // 取得したHTMLに変数を代入
        let html =
          HtmlService.createHtmlOutputFromFile("cancel_success").getContent();
        // 取得したHTMLの一番最初の「Name」を変数の「Name」の中身に置き換え（以降同様）
        let html2 = html.replace("Name", Name);
        let html3 = html2.replace("Cancel_day", Cancel_day);
        let html4 = html3.replace("Calendar_URL", Calendar_URL);
        let html5 = html4.replace("CLNO", CLNO);
        let html6 = html5.replace("Name", Name);
        let html7 = html6.replace("Cancel_NO", Cancel_NO);
        let html8 = html7.replace("Cancel_day", Cancel_day);
        let html_perfect = html8.replace("TimeStamp", TimeStamp);
        // ログを書く
        console.log(html_perfect);
        // 件名を変数「Subject」に代入
        let Subject = "【" + Name + "様へ】 削除完了しました";
        // 本文を変数「Body」に代入
        let Body = "";
        // 追加設定
        let Options = {
          // 送信者を「3Dプリンター 管理者」にする
          name: "3Dプリンター 管理者",
          // メールの本文に変数代入の終わったHTMLデータを挿入
          htmlBody: html_perfect,
        };
        // メール送信
        MailApp.sendEmail(Email, Subject, Body, Options);
        // ログを書く
        console.log("完了メール送信完了");
        // データベース用シートに「削除済みと書く」
        Yoyaku_list_sheet.getRange(Cancel_NO, 6).setValue("削除済み");
        // ログを書く
        console.log("データベース書き換え完了");
      }
      // functionの終了
      return Name;
    };

    // 失敗時のメール送信
    let Can_Not_Cleaning = function (Error_reason) {
      //　自動返信メール
      // HTMLを取得
      doGet_cancel_canot();
      // 取得したHTMLに変数を代入
      let html =
        HtmlService.createHtmlOutputFromFile("cancel_canot").getContent();
      // 取得したHTMLの一番最初の「Name」を変数の「Name」の中身に置き換え（以降同様）
      let html2 = html.replace("Name", Name);
      let html3 = html2.replace("Cancel_day", Cancel_day);
      let html4 = html3.replace("Error_reason", Error_reason);
      let html5 = html4.replace("Calendar_URL", Calendar_URL);
      let html6 = html5.replace("CLNO", CLNO);
      let html7 = html6.replace("Name", Name);
      let html8 = html7.replace("Cancel_day", Cancel_day);
      let html_perfect = html8.replace("TimeStamp", TimeStamp);
      // ログを書く
      console.log(html_perfect);
      // 件名を変数「Subject」に代入
      let Subject = "【" + Name + "様へ】 削除失敗！！";
      // 本文を変数「Body」に代入
      let Body = "";
      // 追加設定
      let Options = {
        // 送信者を「3Dプリンター 管理者」にする
        name: "3Dプリンター 管理者",
        // メールの本文に変数代入の終わったHTMLデータを挿入
        htmlBody: html_perfect,
      };
      // メール送信
      MailApp.sendEmail(Email, Subject, Body, Options);
      // ログを書く
      console.log("削除エラーメールを送信完了");
      // functionの終了
      return Error_reason;
    };

    //キャンセルのメインプロセス
    // もし、キャンセル番号と登録番号が同じなら、
    if (Number(Cancel_NO) === Number(Database_Cancel_NO)) {
      // キャンセル時のメールアドレスと登録時のメールアドレスが同じなら、
      if (Email === Database_cancel_Email) {
        let Cancel_finished = Yoyaku_list_sheet.getRange(
          Cancel_NO,
          6
        ).getValue();
        console.log(Cancel_finished);
        if (Cancel_finished !== "削除済み") {
          // OKパターン
          Can_Cleaning(Name);
          console.log("削除完了");
        } else {
          //予定がなかった時に
          let Error_reason =
            "指定された予約は削除済みであったため、削除できませんでした。";
          Can_Not_Cleaning(Error_reason);
        }
      } else {
        // メールアドレスが違った場合
        let Error_reason = "メールアドレスが予約時と異なります。";
        Can_Not_Cleaning(Error_reason);
        console.log("予約削除失敗");
      }
    } else {
      // ダメパターン
      let Error_reason = "あなたのメールアドレスでその予約は行われていません。";
      Can_Not_Cleaning(Error_reason);
      console.log("予約削除失敗");
    }
  };
  // キャンセルのメインプロセス終了
  //キャンセル関係のプログラム終了////////////////////////////////////////////////////////////////

  //予約かキャンセルかを判断　////////////////////////////////////////////////////////////////////
  if (Choice === "予約する") {
    console.log("予約プログラムを作動させます");
    Yoyaku_program(Yoyaku_day);
  } else {
    console.log("キャンセルプログラムを動作させます");
    Cancel_program(Cancel_NO);
  }
  //予約かキャンセルか判断完了///////////////////////////////////////////////////////////////////
}

function doGet_yoyaku_success() {
  // 「doGet_yoyaku_success」が実行されたときに動く内容を設定
  var t = HtmlService.createTemplateFromFile("yoyaku_success.html");
  return t.evaluate();
}

function doGet_yoyaku_canot() {
  // 「doGet_yoyaku_canot」が実行されたときに動く内容を設定
  var t = HtmlService.createTemplateFromFile("yoyaku_canot.html");
  return t.evaluate();
}

function doGet_yoyaku_full() {
  // 「doGet_yoyaku_full」が実行されたときに動く内容を設定
  var t = HtmlService.createTemplateFromFile("yoyaku_full.html");
  return t.evaluate();
}

function doGet_cancel_success() {
  // 「doGet_cancel_success」が実行されたときに動く内容を設定
  var t = HtmlService.createTemplateFromFile("cancel_success.html");
  return t.evaluate();
}

function doGet_cancel_canot() {
  // 「doGet_cancel_canot」が実行されたときに動く内容を設定
  var t = HtmlService.createTemplateFromFile("cancel_canot.html");
  return t.evaluate();
}
