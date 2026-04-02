function doGet() {
  return HtmlService.createHtmlOutputFromFile("index");
}

/**
 * サイボウズ形式テキストをパースして登録
 */
function syncSchedule(text) {
  const events = parseSchedule(text);
  const calendar = CalendarApp.getDefaultCalendar();
  // 専用カレンダーを使う場合：
  // const calendar = CalendarApp.getCalendarById("xxxx@group.calendar.google.com");

  events.forEach(ev => {
    if (ev.type === "allday") {
      calendar.createAllDayEvent(
        ev.summary,
        ev.start,
        ev.end
      );
    } else {
      calendar.createEvent(
        ev.summary,
        ev.start,
        ev.end
      );
    }
  });

  return `${events.length} 件の予定を登録しました`;
}

/**
 * Pythonロジックの直移植
 */
function parseSchedule(text) {
  const lines = text.trim().split(/\r?\n/);
  const events = [];
  let currentDate = null;

  const datePattern = /(\d+)\s*月\s*(\d+)\s*日/;
  const eventPattern = /(\d{1,2}:\d{2})-(\d{1,2}:\d{2})\s+(.*)/;
  const year = new Date().getFullYear();

  for (const lineRaw of lines) {
    const line = lineRaw.trim();
    if (line.startsWith("-----")) break;

    const dateMatch = line.match(datePattern);
    if (dateMatch) {
      currentDate = {
        month: Number(dateMatch[1]),
        day: Number(dateMatch[2])
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
          `${year}-${currentDate.month}-${currentDate.day}T${startTime}:00`
        );
        const end = new Date(
          `${year}-${currentDate.month}-${currentDate.day}T${endTime}:00`
        );

        events.push({ type: "normal", start, end, summary });
      }
    }
  }
  return events;
}

