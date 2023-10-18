
/**
 * カレンダーID　を以下のダブルクオーテーション内に入れてください。以下のURLで方法を確認してください。
 * https://blog-and-destroy.com/41932
 */
const calender_id = "";

function myFunction() {
  // 学期の開始日と終了日を取得
  const lastRow = 26;
  const ss_timetable = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("時間割").getRange(2, 1, lastRow - 1, 11).getValues();
  const ss_school_hours = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("授業時間").getRange(2, 2, 5, 2).getValues();
  for (let lecture of ss_timetable) {
    lecture_info = { "day": lecture[0], "time": lecture[1], "name": lecture[2], "teacher_name": lecture[3], "room_number": lecture[4], "moodle": lecture[5], "start_date": lecture[6], "end_date": getNextDate(lecture[7]), }
    console.log(lecture[0]+ lecture[1] +": "+lecture[2])
    if (lecture[2] != "") {
      lecture_info.time = parseInt(lecture_info.time)
      let day = lecture_info.day;
      let title = `📝【${lecture_info.room_number}】${lecture_info.name}`
      let startTime = time_formatting(lecture_info.start_date, ss_school_hours[lecture_info.time - 1][0])
      let endTime = time_formatting(lecture_info.start_date, ss_school_hours[lecture_info.time - 1][1])
      let description = `${lecture_info.teacher_name}\n${lecture_info.moodle}\n`
      let endDate = new Date(lecture_info.end_date)
      console.log(startTime)
      set_calender(day, title, startTime, endTime, endDate, description)
    }
  }
}
function time_formatting(day, time) {
  const d = new Date
  d.setFullYear(parseInt(Utilities.formatDate(day, 'Asia/Tokyo', 'yyyy')), parseInt(Utilities.formatDate(day, 'Asia/Tokyo', 'MM')) - 1, parseInt(Utilities.formatDate(day, 'Asia/Tokyo', 'dd')))
  d.setHours(parseInt(Utilities.formatDate(time, 'Asia/Tokyo', 'H')), parseInt(Utilities.formatDate(time, 'Asia/Tokyo', 'm')), 0)
  return d
}
function getNextDate(date) {
  const d = new Date(date);
  d.setDate(d.getDate() + 1)
  return Utilities.formatDate(d, 'JST', 'yyyy/MM/dd');
}

function set_calender(day, title, startTime, endTime, endDate, description) {
  const cal = CalendarApp.getCalendarById(calender_id)
  const recurrence = CalendarApp.newRecurrence().addWeeklyRule().onlyOnWeekday(CalendarApp.Weekday[day]).until(endDate)
  const options = {
    description: description
  }
  const event = cal.createEventSeries(title, startTime, endTime, recurrence, options)
  console.log(event.getId())
}