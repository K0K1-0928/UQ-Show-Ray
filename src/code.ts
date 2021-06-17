function UQShowRay() {
  // スプレッドシートから設定情報を取得
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('config');
  const userId = sh.getRange(2, 1).getValue();
  const afterMonths = sh.getRange(2, 2).getValue();

  // 実行対象のカレンダーを取得
  const calendarId = userId || Session.getActiveUser().getEmail();
  const calendar = CalendarApp.getCalendarById(calendarId);

  // 実行日及び実行対象期間終了日を取得
  const start = new Date();
  const end = new Date(start);
  end.setMonth(start.getMonth() + afterMonths);

  // 実行対象期間のUQ Show Ray登録をリセットする
  removeUQShowRayToCalendar(calendar, start, end);

  // 実行対象期間の休暇候補日を判定し処理を行う
  for (let d = start; d < end; d.setDate(d.getDate() + 1)) {
    // 平日のため処理をスキップ
    if (!isHoliday(d)) {
      continue;
    }

    const nextDay = new Date(d);
    nextDay.setDate(d.getDate() + 1);
    // 翌日も休日のため処理をスキップ
    if (isHoliday(nextDay)) {
      continue;
    }

    const nextDayButOne = new Date(d);
    nextDayButOne.setDate(d.getDate() + 2);
    // 翌々日が休日の場合、翌日を休暇候補日として処理を行う
    if (isHoliday(nextDayButOne)) {
      // カレンダー登録の処理
      setUQShowRayToCalendar(calendar, nextDay);
    }
  }
}

/**
 * 日本の祝日カレンダーのオブジェクトを取得する
 *
 * @return {*}  {GoogleAppsScript.Calendar.Calendar}
 */
function getJapanPublicHolidays(): GoogleAppsScript.Calendar.Calendar {
  const calendarId = 'ja.japanese#holiday@group.v.calendar.google.com';
  return CalendarApp.getCalendarById(calendarId);
}

/**
 * 日付が土日祝かどうか判定する
 *
 * @param {Date} date - 判定する日付
 * @return {*}  {boolean}
 */
function isHoliday(date: Date): boolean {
  const weekday = date.getDay();
  if (weekday === 0 || weekday === 6) {
    return true;
  }
  const calendar = getJapanPublicHolidays();
  const events = calendar.getEventsForDay(date);
  return Boolean(events.length);
}

/**
 * Googleカレンダーから、対象期間に存在する休暇候補日イベントを削除する
 *
 * @param {GoogleAppsScript.Calendar.Calendar} calendar - イベントを削除するGoogleカレンダー
 * @param {Date} start - 対象期間の開始日
 * @param {Date} end - 対象期間の終了日
 */
function removeUQShowRayToCalendar(
  calendar: GoogleAppsScript.Calendar.Calendar,
  start: Date,
  end: Date
) {
  const events = calendar.getEvents(start, end, {
    search: '休暇候補日 (UQ Show Ray)',
  });
  events.forEach((event) => event.deleteEvent());
}

/**
 * Googleカレンダーに休暇候補日イベントを追加する
 *
 * @param {GoogleAppsScript.Calendar.Calendar} calendar - イベントを追加するGoogleカレンダー
 * @param {Date} date
 */
function setUQShowRayToCalendar(
  calendar: GoogleAppsScript.Calendar.Calendar,
  date: Date
) {
  const title = '休暇候補日 (UQ Show Ray)';
  const description = 'この予定はUQ Show Rayで作成されました';
  const eventObject = calendar.createAllDayEvent(title, date, {
    description: description,
  });
  setReminder(eventObject);
}

/**
 * スプレッドシートの設定値に基づいてリマインダーを設定する
 *
 * @param {GoogleAppsScript.Calendar.CalendarEvent} event - リマインダーを設定するイベント
 */
function setReminder(event: GoogleAppsScript.Calendar.CalendarEvent) {
  // 既存のリマインダー情報を削除
  event.removeAllReminders();

  // スプレッドシートから設定情報を取得
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('config');
  const days = sh.getRange(2, 3).getValue();
  const hour = sh.getRange(2, 4).getValue();
  const minutes = sh.getRange(2, 5).getValue();

  // 設定時間を分単位に換算してリマインダーを追加
  const minutesBefore = days * 24 * 60 - hour * 60 - minutes;
  event.addPopupReminder(minutesBefore);
}
