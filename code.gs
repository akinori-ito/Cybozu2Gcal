function doGet() {
  return HtmlService.createHtmlOutputFromFile("index");
}


const SYNC_TAG = "[Garoon]";

function existsDuplicate(calendar, ev) {
  // 少し広めに検索（APIの仕様対策）
  const searchStart = new Date(ev.start.getTime() - 60 * 1000);
  const searchEnd   = new Date(ev.end.getTime() + 60 * 1000);

  const events = calendar.getEvents(searchStart, searchEnd);

  return events.some(e => {
    const sameTitle = e.getTitle() === ev.summary;
    const hasTag = (e.getDescription() || "").includes(SYNC_TAG);
    return sameTitle && hasTag;
  });
}

/**
 * サイボウズ形式テキストをパースして登録
 */
function syncSchedule(text,year) {
  const events = parseSchedule(text,year);
  const calendar = CalendarApp.getDefaultCalendar();
  // または同期専用カレンダー
  // const calendar = CalendarApp.getCalendarById("xxxx@group.calendar.google.com");

  let added = 0;
  let skipped = 0;

  events.forEach(ev => {
    if (existsDuplicate(calendar, ev)) {
      skipped++;
      return;
    }

    const description = `${SYNC_TAG}Garoon予定表から同期`;

    if (ev.type === "allday") {
      calendar.createAllDayEvent(
        ev.summary,
        ev.start,
        ev.end,
        { description }
      );
    } else {
      calendar.createEvent(
        ev.summary,
        ev.start,
        ev.end,
        { description }
      );
    }
    added++;
  });

  return `登録: ${added} 件 / スキップ（重複）: ${skipped} 件`;
}

function zeropad(x,n) {
  if (x.length != n) 
    return "0"+x;
  return x;
}

function dateStr(year,month,day,time) {
  let dstr = year+"-";
  dstr += zeropad(month,2)+"-";
  dstr += zeropad(day,2)+"T";
  dstr += zeropad(time,5)+":00";
  return dstr;
}
/**
 * Pythonロジックの直移植
 */
function parseSchedule(text,year) {
  const lines = text.trim().split(/\r?\n/);
  const events = [];
  let currentDate = null;

  const datePattern = /(\d+)\s*月\s*(\d+)\s*日/;
  const eventPattern = /(\d{1,2}:\d{2})-(\d{1,2}:\d{2})\s+(.*)/;
  //const year = new Date().getFullYear();

  for (const lineRaw of lines) {
    const line = lineRaw.trim();
    if (line.startsWith("-----")) break;

    const dateMatch = line.match(datePattern);
    if (dateMatch) {
      currentDate = {
        month: dateMatch[1],
        day: dateMatch[2]
      };
      continue;
    }

    const eventMatch = line.match(eventPattern);
    if (eventMatch && currentDate) {
      const [, startTime, endTime, summary] = eventMatch;

      if (startTime === "0:00" && endTime === "24:00") {
        const start = new Date(year, currentDate.month - 1, currentDate.day);
        const end = new Date(start);
        end.setDate(end.getDate() + 1);

        events.push({ type: "allday", start, end, summary });
      } else {
        const start = new Date(
          dateStr(year,currentDate.month,currentDate.day,startTime)
        );
        const end = new Date(
          dateStr(year,currentDate.month,currentDate.day,endTime)
        );

        events.push({ type: "normal", start, end, summary });
      }
    }
  }
  return events;
}

