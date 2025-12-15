Cybozuのガルーン(Garoon)はよくできたソフトウェアなのだが、他のサービスとの連携、という点では使いにくかった。

以前からSOAP APIはあったものの、繰り返し予定が難関で、実際の日にちを都度計算しなければならず、同期ツールの作成を断念していた。

だがREST APIが登場し、繰り返し予定が1件ずつ返されるようになってとても処理しやすくなったので、GASで動作するガルーンからGoogleカレンダーへの一方向同期スクリプトを作成した。

## **作成手順**

### **作成〜タイムゾーンの設定**

まずはGoogleドライブからGoogle Apps Scriptを作成する。

プロジェクト名はお好みで(以下の例では`Garoon GCalendar Sync`)。

またIDEは新IDEを使用した場合で説明する。

まずはタイムゾーンを設定するため、「プロジェクトの設定」から、『「appsscript.json」マニフェストファイルをエディタで表示する』にチェックを入れる。

![](https://www.330k.info/essay/sync-garoon-google-calendar/001.webp)

エディタで「appsscript.json」を開き、`timeZone`の値として`Asia/Tokyo`を設定して保存する。

![](https://www.330k.info/essay/sync-garoon-google-calendar/002.webp)

またスクリプトエンジンとしてV8エンジンを使用するよう、「Chrome V8ランタイムを有効にする」にチェックが入っていることを確認。

### **メインスクリプト**
https://script.google.com/home
次にスクリプト本体として、以下のコードを入れる。

```jsx
/**
 * ガルーンの予定をREST APIを使用して取得しGoogle Calendarに同期
 */
function syncGroonSchedule() {
  const TAG_GAROON_UNIQUE_EVENT_ID = "GAROON_UNIQUE_EVENT_ID";
  const TAG_GAROON_SYNC_DATETIME = "GAROON_SYNC_DATETIME";

  const userproperties = PropertiesService.getUserProperties().getProperties();
  const cybozuuser = userproperties["CybozuUser"];
  const cybozupass = userproperties["CybozuPassword"];
  const syncdaysbefore = JSON.parse(userproperties["SyncDaysBefore"]);
  const syncdaysafter = JSON.parse(userproperties["SyncDaysAfter"]);
  
  const today = new Date();
  const datestart = new Date(today);
  const dateend = new Date(today);
  
  datestart.setDate(today.getDate() - syncdaysbefore);
  datestart.setHours(0, 0, 0, 0);
  dateend.setDate(today.getDate() + syncdaysafter);
  dateend.setHours(23, 59, 59, 0);
  console.log([datestart, dateend]);

  // Garoonの予定一覧をREST APIを使って取得(とりあえず最大200件。GASの6分の実行制限時間があるため、360件以上に対応するならスプレッドシートに書き出しておく、などの処理が必要)
  const result = UrlFetchApp.fetch("https://dym.cybozu.com/g/api/v1/schedule/events?rangeStart="
   + encodeURIComponent(formatISODateTime(datestart)) + "&rangeEnd=" + encodeURIComponent(formatISODateTime(dateend)) + "&orderBy=start%20asc&limit=200", {
    method: "get",
    headers: {
      "X-Cybozu-Authorization": Utilities.base64Encode(cybozuuser + ":" + cybozupass),
      "Content-Type": "application/json"
    }
  });
  //console.log(result.getContentText("UTF-8"));
  const events = JSON.parse(result.getContentText("UTF-8")).events;
  
  // 期間内のGoogleカレンダー予定を取得
  const calendar = getSyncCalendar();
  const gcalexistingevents = calendar.getEvents(datestart, dateend);

  for(const event of events){
    let uniqueid = getGaroonUniqueEventID(event);
    console.log(["GAROON EVENT", event.subject, event.start.dateTime, event.updatedAt, uniqueid]);

    const gcalexistingevent = gcalexistingevents.find((e) => e.getTag(TAG_GAROON_UNIQUE_EVENT_ID) === uniqueid);
    if(gcalexistingevent){
      // ガルーンの予定がすでにGoogle Calendarに存在する
      if((new Date(event.updatedAt)).getTime() > new Date(gcalexistingevent.getTag(TAG_GAROON_SYNC_DATETIME)).getTime()){
        // 最終同期日時よりも更新日時が新しいときは既存のイベントを削除
        gcalexistingevent.deleteEvent();
        console.log("UPDATED. DELETED EXISTING EVENT");

      }else{
        // 更新がない場合はスキップ
        console.log("SKIPPED");
        continue;
      }
    }

    let gcalevent;
    const eventtitle = (event.eventMenu ? (event.eventMenu + " ") : "") + event.subject;
    // eventtitleに「誕生日、メモ、めも、休み、FC東京」が含まれている場合はスキップ
    if(/誕生日|メモ|めも|休み|FC東京/.test(eventtitle)){
      console.log("SKIPPED BY TITLE");
      continue;
    }
    const eventoptions = {
      "description": event.notes
    }

    if(event.isAllDay){
      let eventstart = new Date(event.start.dateTime);
      let eventend = new Date(event.end.dateTime);
      // eventstartとeventendのあいだが23:59:59以上の場合はスキップ
      if((eventend.getTime() - eventstart.getTime()) >= 86399000){
        console.log("SKIPPED BY DURATION");
        continue;
      }

      //終日予定の終了時刻はガルーンは当日の23:59:59で返ってくるが、Google Calendarは翌日00:00:00にする
      eventend.setSeconds(eventend.getSeconds() + 1);
      gcalevent = calendar.createAllDayEvent(eventtitle, eventstart, eventend, eventoptions);

    }else{
      let eventstart = new Date(event.start.dateTime);
      let eventend;
      if(event.isStartOnly){
        eventend = new Date(event.start.dateTime);
      }else{
        eventend = new Date(event.end.dateTime);
      }
      gcalevent = calendar.createEvent(eventtitle, eventstart, eventend, eventoptions);
    }

    gcalevent.setTag(TAG_GAROON_UNIQUE_EVENT_ID, uniqueid);
    gcalevent.setTag(TAG_GAROON_SYNC_DATETIME, today.toISOString());
    console.log("CREATED");

    // sleepなしで連続追加するとすぐにエラーが出るので暫定対処
    Utilities.sleep(1000);
  }

  // ガルーンで削除された予定をGoogle Calendarから削除
  const garoonuniqueeventids = events.map((e) => getGaroonUniqueEventID(e));
  for(const e of gcalexistingevents){
    if(!garoonuniqueeventids.includes(e.getTag(TAG_GAROON_UNIQUE_EVENT_ID))){
      console.log("DELETED: " + e.getTitle());
      e.deleteEvent();
    }
  }
}

/**
 * Google CalendarのTagに設定する予定ID(GaroonのeventIdとreapeatIdの組み合わせて一意にする)
 * @return {string}
 */
function getGaroonUniqueEventID(garoonevent){
  return garoonevent.id + (garoonevent.repeatId ? ("-" + garoonevent.repeatId) : "");
}

/**
 * 同期対象のGoogle Calendarを取得(存在しなければ作成する)
 * @return {CalendarApp.Calendar}
 */
function getSyncCalendar(){
  const calendarname = PropertiesService.getUserProperties().getProperty("CalendarName");
  const calendars = CalendarApp.getOwnedCalendarsByName(calendarname);

  if(calendars.length){
    return calendars[0];
  }else{
    const calendar = CalendarApp.createCalendar(calendarname);
    return calendar;
  }
}

/**
 * 日付をyyyy-MM-ddTHH:mm:ss+hh:mm形式にする
 * @param {Date} d
 * @return {string}
 */
function formatISODateTime(d){
  return d.getFullYear() + "-" + ("0" + (d.getMonth() + 1)).slice(-2) + "-" + ("0" + d.getDate()).slice(-2)
   + "T" + ("0" + d.getHours()).slice(-2) + ":" + ("0" + d.getMinutes()).slice(-2) + ":" + ("0" + d.getSeconds()).slice(-2)
   + (d.getTimezoneOffset() <= 0 ? "+" : "-") + ("0" + Math.floor(Math.abs(d.getTimezoneOffset()) / 60)).slice(-2) + ":" + ("0" + (Math.abs(d.getTimezoneOffset()) % 60)).slice(-2);
}
```

コード中、`https://*****.cybozu.com/g/api/v1/schedule/`の部分は使っている環境のリクエストURIに書き換える。

参照) [**Garoon REST APIの共通仕様 – cybozu developer network**](https://developer.cybozu.io/hc/ja/articles/360000503306-Garoon-REST-API%E3%81%AE%E5%85%B1%E9%80%9A%E4%BB%95%E6%A7%98#step4)

仕組みとしては、GoogleカレンダーにタグとしてガルーンのeventIdとrepeatIdを組み合わせたものを保存することで、ガルーンの予定との対応を保存している。 repeatIdに関してはまだドキュメントに載っていないので、将来の仕様変更はあるかもしれないが、スクリプトを作れたのはこのrepeatIdのおかげ。

また、各予定にタグとしてガルーンとの同期日時を記録することで、途中でエラー(「短時間に作成したカレンダーまたはカレンダーの予定の数が多すぎます。しばらくしてからもう一度お試しください。」など)が発生したあとに再実行する際に、予定を作り直す回数を少なくしている。

### **初期設定**

UserPropertiesに設定値を保存するため、適当にスクリプトファイルを作成し、下記内容の関数を作成。

Cybozuのユーザ名とパスワードを自分のものに書き換えて(他のところもお好みで)、IDEから関数`initialize`を実行しておく。

```jsx
function initialize(){
  PropertiesService.getUserProperties().setProperties({
    "CybozuUser": "****",
    "CybozuPassword": "****",
    "SyncDaysBefore": 7,
    "SyncDaysAfter": 30,
    "CalendarName": "Garoon"
  })
}
```

一度実行したら不要なので削除しておく(面倒だし、機密情報をスクリプトにベタ書きしたくないからUserPropertiesに保存してるのに。何とかなりませんかね、Google先生)。

### **初回起動およびトリガーの設定**

一度`syncGroonSchedule`をIDEから実行し、動作することを確認。

IDEから毎回`syncGroonSchedule`を実行してもよいが、せっかくなのでトリガーを設定して自動的に同期するようにしておく。

「トリガー」から「トリガーを追加」とし、以下の画像のように設定する。

![](https://www.330k.info/essay/sync-garoon-google-calendar/003.webp)

Googleカレンダーのほうが意外とすぐに制限に引っかかるので、とりあえず6時間に一度同期するようにしておく。

### **Googleカレンダーで表示**

うまく動いていればGoogleカレンダーに新たなカレンダーが加わり、そこにガルーンの予定が見れるはず。

## **備考**

## **サイボウズ公式のプログラム**

サイボウズ公式で、Javaで動作するサンプルプログラム([**Googleカレンダー連携 - Garoonの予定をGoogleカレンダーに表示 - – cybozu developer network**](https://developer.cybozu.io/hc/ja/articles/204426680-Google%E3%82%AB%E3%83%AC%E3%83%B3%E3%83%80%E3%83%BC%E9%80%A3%E6%90%BA-Garoon%E3%81%AE%E4%BA%88%E5%AE%9A%E3%82%92Google%E3%82%AB%E3%83%AC%E3%83%B3%E3%83%80%E3%83%BC%E3%81%AB%E8%A1%A8%E7%A4%BA-))は当然うまく動作する。

しかしながら

- 動作させるまでの手順が多い
- 自動で同期させるためには24時間稼働のマシンを用意して、cronなどでタスクを設定する必要がある
- サービスアカウントはドメインのメンバーではない(参考: [**サービス アカウント | Cloud IAM のドキュメント | Google Cloud**](https://cloud.google.com/iam/docs/service-accounts?hl=ja))ため、Google Workspaceの設定でカレンダーの「予備カレンダーの外部共有オプション」を「すべての情報を共有する (外部ユーザーにカレンダーの変更を許可する)」に設定する必要がある
    
といった欠点がある(その代わりオンプレ版のガルーンでも動作可能という長所あり)。

本プログラムであれば実行するのは自分自身なので、予備カレンダーの外部共有オプションを変更することなくGaroonからGoogleカレンダーへの同期ができ、またGAS上で動くので自分でマシンを用意する必要がない。

ちなみに社内ではユーザ向けに、上記のスクリプトに加えてUserPropertiesの設定とトリガーの設定を行う画面を作成し、ウェブアプリとして公開した(「ウェブアプリケーションにアクセスしているユーザーとして実行」および「ドメイン内の全ユーザがアクセス可能」と設定)。

この場合でも実行するのは同じドメインのユーザなので、予備カレンダーの外部共有オプションは変える必要はない。

ただし、Google Apps Scriptのファイル自体を実行するユーザ(=アクセスするユーザ)が読み取り権限を持っていないとトリガーの実行ができなかったので、スクリプトを全員に閲覧者として共有しておく必要がある。

## **References**

- [**Garoon REST APIの共通仕様 – cybozu developer network**](https://developer.cybozu.io/hc/ja/articles/360000503306)
    
- [**予定の取得（GET） – cybozu developer network**](https://developer.cybozu.io/hc/ja/articles/360000440583)
    
- [**スケジュールオブジェクト – cybozu developer network**](https://developer.cybozu.io/hc/ja/articles/115005314266)
