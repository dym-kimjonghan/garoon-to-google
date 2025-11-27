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

