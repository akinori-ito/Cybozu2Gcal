function doGet() {
  return HtmlService.createHtmlOutputFromFile("index");
}

const SYNC_TAG = "[Garoon]";

function existsDuplicate(calendar, ev) {
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
 * メイン関数：形式を自動判別して同期
 */
function syncSchedule(text, year) {
  let events = [];
  
  // 1行目に「開始日付」が含まれていればCSVとみなす
  if (text.trim().startsWith("開始日付") || text.includes('"開始日付"')) {
    events = parseCsvSchedule(text);
  } else {
    events = parseSchedule(text, year);
  }

  const calendar = CalendarApp.getDefaultCalendar();
  let added = 0;
  let skipped = 0;

  events.forEach(ev => {
    if (existsDuplicate(calendar, ev)) {
      skipped++;
      return;
    }

    const description = `${SYNC_TAG} Garoonから同期\n${ev.description || ""}`;

    if (ev.type === "allday") {
      calendar.createAllDayEvent(ev.summary, ev.start, ev.end, { description });
    } else {
      calendar.createEvent(ev.summary, ev.start, ev.end, { description });
    }
    added++;
  });

  return `登録: ${added} 件 / スキップ: ${skipped} 件`;
}

/**
 * CSV形式をパース（空時刻を終日予定として扱う修正版）
 * 項目順: 開始日付, 開始時刻, 終了日付, 終了時刻, 予定, 予定詳細, メモ...
 */
function parseCsvSchedule(csvText) {
  const data = Utilities.parseCsv(csvText.trim());
  const events = [];

  // 1行目はヘッダーなので i=1 から開始
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row.length < 6) continue; // 予定詳細まで届かない行はスキップ

    const startDateStr = row[0]; // 年/月/日
    const startTimeStr = row[1]; // 時:分:秒 または 空文字
    const endDateStr   = row[2];
    const endTimeStr   = row[3]; // 時:分:秒 または 空文字
    const summary      = row[5]; // 「予定詳細」
    const memo         = row[6]; // 「メモ」

    let start, end, type;

    // 時刻が空文字列、または 00:00〜00:00(24:00) の場合は終日予定とする
    const isStartTimeEmpty = (!startTimeStr || startTimeStr.trim() === "");
    const isEndTimeEmpty = (!endTimeStr || endTimeStr.trim() === "");
    const isFullDayTime = (startTimeStr === "00:00:00" && (endTimeStr === "00:00:00" || endTimeStr === "24:00:00"));

    if (isStartTimeEmpty || isEndTimeEmpty || isFullDayTime) {
      type = "allday";
      // 終日予定の場合、Dateオブジェクトは時刻なし(00:00:00)で作成
      start = new Date(startDateStr);
      end = new Date(endDateStr);
      // Googleカレンダーの終日予定は「終了日の翌日」を終了値とする必要がある
      // 例：1月1日の1日のみの予定なら、start=1/1, end=1/2 と設定する
      end.setDate(end.getDate() + 1);
    } else {
      type = "normal";
      // 通常の予定（時刻あり）
      start = new Date(`${startDateStr} ${startTimeStr}`);
      end = new Date(`${endDateStr} ${endTimeStr}`);
    }

    // 無効な日付のチェック
    if (isNaN(start.getTime())) continue;

    events.push({
      type: type,
      start: start,
      end: end,
      summary: summary || "(無題)",
      description: memo
    });
  }
  return events;
}

// --- 従来のテキストパース用補助関数 ---
function zeropad(x,n) {
  return x.toString().padStart(n, "0");
}

function dateStr(year,month,day,time) {
  return `${year}-${zeropad(month,2)}-${zeropad(day,2)}T${zeropad(time,5)}:00`;
}


function parseSchedule(text,year) {
  const lines = text.trim().split(/\r?\n/);
  const events = [];
  let currentDate = null;
  const datePattern = /(\d+)\s*月\s*(\d+)\s*日/;
  const eventPattern = /(\d{1,2}:\d{2})-(\d{1,2}:\d{2})\s+(.*)/;

  for (const lineRaw of lines) {
    const line = lineRaw.trim();
    if (line.startsWith("-----")) break;
    const dateMatch = line.match(datePattern);
    if (dateMatch) {
      currentDate = { month: dateMatch[1], day: dateMatch[2] };
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
        const start = new Date(dateStr(year,currentDate.month,currentDate.day,startTime));
        const end = new Date(dateStr(year,currentDate.month,currentDate.day,endTime));
        events.push({ type: "normal", start, end, summary });
      }
    }
  }
  return events;
}