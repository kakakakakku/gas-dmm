var CALENDAR_NAME = 'Calendar';
var SEARCH_QUERY = 'is:unread subject:"【DMM英会話】レッスン予約完了のお知らせ"';

function DmmCalendar() {
  var calendars = CalendarApp.getCalendarsByName(CALENDAR_NAME);
  var calendar;
  if (calendars.length === 0) {
    Logger.log(CALENDAR_NAME + "カレンダーが見つかりません");
    return;
  } else {
    calendar = calendars[0];
  }
  var threads = GmailApp.search(SEARCH_QUERY, 0, 1);
  if (threads.length === 0) {
    Logger.log("メールが見つかりません");
    return;
  }

  var thread = threads[0]
  var messages = thread.getMessages();
  var message = messages[messages.length-1];
  var body = message.getBody();
  body = body.replace(/<br \/>/g, "");

  var teacher = null;
  var teacherName = null;
  var startAt = null;
  var date = null;
  var startTime = null;

  var rows = body.split("\n");
  for (var i in rows) {
    var row = rows[i];
    if (row.indexOf("講師名") === 0) {
      teacher = row;
      teacherName = teacher.replace("講師名：", "");
    } else if (row.indexOf("ご予約日") === 0) {
      date = row.replace("ご予約日：", "");
      date = date.replace("年", "/");
      date = date.replace("月", "/");
      date = date.replace("日", "/");
    } else if (row.indexOf('開始時間') === 0) {
      startAt = row;
      startTime = startAt.replace("開始時間：", "");
      startTime = startTime.replace("時", ":");
      startTime = startTime.replace("分", "");
    }
  }

  var startDateStr = date + ' ' + startTime;
  var startDate = new Date(startDateStr);
  var endDate = new Date(startDate);
  endDate.setMinutes(endDate.getMinutes() + 25);

  if (teacher && startDate) {
    // カレンダーに予定を登録する
    // https://developers.google.com/apps-script/reference/calendar/calendar-app
    var title = 'DMM 英会話 (' + teacherName + ')';
    calendar.createEvent(title, startDate, endDate);

    // メールを既読にする
    // https://developers.google.com/apps-script/reference/gmail/gmail-thread
    thread.markRead();
  }
}
