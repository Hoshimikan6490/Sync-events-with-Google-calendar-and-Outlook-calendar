/**
 * RRULE 文字列を解析して RRULE オブジェクトに変換する。
 * 例: "FREQ=WEEKLY;UNTIL=20260331T083000Z;INTERVAL=1;BYDAY=TU;WKST=SU"
 * @param {string} rruleString RRULE 文字列
 * @returns {{freq: string, until: Date|null, interval: number, byDay: string[], count: number|null}} RRULE オブジェクト
 */
function parseRRULE(rruleString) {
  const rrule = {
    freq: 'DAILY',
    until: null,
    interval: 1,
    byDay: [],
    count: null,
  };

  if (!rruleString) {
    return rrule;
  }

  const pairs = rruleString.split(';');
  pairs.forEach((pair) => {
    const [key, value] = pair.split('=');
    if (!key || !value) return;

    switch (key.toUpperCase()) {
      case 'FREQ':
        rrule.freq = value.toUpperCase();
        break;
      case 'UNTIL':
        const untilDate = parseICSDateValue(value, {});
        if (untilDate) {
          rrule.until = untilDate.date;
        }
        break;
      case 'INTERVAL':
        rrule.interval = parseInt(value, 10) || 1;
        break;
      case 'BYDAY':
        rrule.byDay = value.split(',').map((d) => d.trim());
        break;
      case 'COUNT':
        rrule.count = parseInt(value, 10) || null;
        break;
    }
  });

  return rrule;
}

/**
 * RRULE に基づいて繰り返しイベントを個別イベントに展開する。
 * @param {{uid: string, title: string, start: Date, end: Date, isAllDay: boolean, transparency: GoogleAppsScript.Calendar.EventTransparency, visibility: GoogleAppsScript.Calendar.Visibility}} baseEvent ベースイベント
 * @param {string} rruleString RRULE 文字列
 * @returns {{uid: string, title: string, start: Date, end: Date, isAllDay: boolean, transparency: GoogleAppsScript.Calendar.EventTransparency, visibility: GoogleAppsScript.Calendar.Visibility}[]} 展開されたイベント配列
 */
function expandRecurringEvent(baseEvent, rruleString) {
  const rrule = parseRRULE(rruleString);
  const expanded = [];
  const instances = getRecurrenceInstances(baseEvent.start, rrule);
  const duration = baseEvent.end.getTime() - baseEvent.start.getTime();

  instances.forEach((startTime, index) => {
    const eventStart = new Date(startTime);
    const eventEnd = new Date(eventStart.getTime() + duration);

    expanded.push({
      uid: baseEvent.uid + '-' + index,
      title: baseEvent.title,
      start: eventStart,
      end: eventEnd,
      isAllDay: baseEvent.isAllDay,
      transparency: baseEvent.transparency,
      visibility: baseEvent.visibility,
    });
  });

  return expanded;
}

/**
 * RRULE に基づいて、イベント発生日時の配列を返す。
 * @param {Date} startDate ベースの開始日時
 * @param {{freq: string, until: Date|null, interval: number, byDay: string[], count: number|null}} rrule RRULE オブジェクト
 * @returns {Date[]} 発生日時の配列
 */
function getRecurrenceInstances(startDate, rrule) {
  const instances = [];
  const maxCount = rrule.count || 1000; // 無制限の場合は最大 1000 件
  const until = rrule.until || new Date('2099-12-31');

  if (rrule.freq === 'WEEKLY' && rrule.byDay.length > 0) {
    // WEEKLY + BYDAY の場合は効率的に処理
    let current = new Date(startDate);

    // 最初の対象曜日を探す
    while (!isTargetWeekday(current, rrule.byDay) && current <= until) {
      current = addDays(current, 1);
    }

    // 対象曜日のみ追加
    while (current <= until && instances.length < maxCount) {
      instances.push(new Date(current));
      current = addDays(current, 7 * rrule.interval);
    }
  } else {
    // その他の場合
    let current = new Date(startDate);
    let count = 0;

    while (current <= until && count < maxCount) {
      instances.push(new Date(current));
      count++;

      switch (rrule.freq) {
        case 'DAILY':
          current = addDays(current, rrule.interval);
          break;
        case 'WEEKLY':
          current = addDays(current, 7 * rrule.interval);
          break;
        case 'MONTHLY':
          current = new Date(
            current.getFullYear(),
            current.getMonth() + rrule.interval,
            current.getDate(),
          );
          break;
        case 'YEARLY':
          current = new Date(
            current.getFullYear() + rrule.interval,
            current.getMonth(),
            current.getDate(),
          );
          break;
        default:
          current = addDays(current, rrule.interval);
      }
    }
  }

  return instances;
}

/**
 * 指定日付が対象曜日に含まれるか判定する。
 * @param {Date} date 判定対象の日付
 * @param {string[]} targetDays 対象曜日の配列（例: ['TU']）
 * @returns {boolean} 対象曜日なら true
 */
function isTargetWeekday(date, targetDays) {
  const dayMap = {
    SU: 0,
    MO: 1,
    TU: 2,
    WE: 3,
    TH: 4,
    FR: 5,
    SA: 6,
  };

  const targetDayNumbers = targetDays.map((d) => dayMap[d] || 0);
  return targetDayNumbers.includes(date.getDay());
}

/**
 * 指定された曜日の次の発生日を取得する。
 * 例: TU = 火曜日, WE = 水曜日
 * @param {Date} current 現在の日付
 * @param {string[]} targetDays 対象曜日の配列（例: ['TU']）
 * @param {number} interval 週の間隔
 * @returns {Date} 次の発生日
 */

/**
 * ICS テキストを解析してイベント配列を返す。
 * 繰り返しイベント (RRULE) は個別のイベントに展開される。
 * @param {string} ics ICS 生テキスト
 * @returns {{uid: string, title: string, start: Date, end: Date, isAllDay: boolean, transparency: GoogleAppsScript.Calendar.EventTransparency, visibility: GoogleAppsScript.Calendar.Visibility}[]} 解析済みイベント
 */
function parseICS(ics) {
  const lines = unfoldICSLines(ics);
  const events = [];

  let current = null;

  lines.forEach((line) => {
    if (line === 'BEGIN:VEVENT') {
      current = {};
      return;
    }

    if (line === 'END:VEVENT' && current) {
      const parsed = toParsedEvent(current);
      if (parsed) {
        // RRULEがある場合は、繰り返しイベントを展開
        if (current.rrule) {
          const expanded = expandRecurringEvent(parsed, current.rrule);
          events.push(...expanded);
        } else {
          events.push(parsed);
        }
      }
      current = null;
      return;
    }

    if (!current) {
      return;
    }

    const property = parsePropertyLine(line);
    if (!property) {
      return;
    }

    if (property.name === 'UID') {
      current.uid = property.value.trim();
    }
    if (property.name === 'SUMMARY') {
      current.title = property.value.trim();
    }
    if (property.name === 'DTSTART') {
      current.startInfo = parseICSDateValue(property.value, property.params);
    }
    if (property.name === 'DTEND') {
      current.endInfo = parseICSDateValue(property.value, property.params);
    }
    if (property.name === 'TRANSP') {
      current.transparency = parseTransparency(property.value);
    }
    if (property.name === 'CLASS') {
      current.visibility = parseVisibility(property.value);
    }
    if (property.name === 'RRULE') {
      current.rrule = property.value.trim();
    }
    if (
      property.name === 'X-MICROSOFT-CDO-BUSYSTATUS' &&
      !current.transparency
    ) {
      current.transparency = parseBusyStatus(property.value);
    }
  });

  return events;
}

/**
 * ICS の CLASS 値を Google Calendar の公開設定に変換する。
 * @param {string} value CLASS 値
 * @returns {GoogleAppsScript.Calendar.Visibility} 公開設定
 */
function parseVisibility(value) {
  const normalized = String(value || '')
    .trim()
    .toUpperCase();

  if (normalized === 'PUBLIC') {
    return CalendarApp.Visibility.PUBLIC;
  }

  if (normalized === 'PRIVATE') {
    return CalendarApp.Visibility.PRIVATE;
  }

  if (normalized === 'CONFIDENTIAL') {
    return CalendarApp.Visibility.CONFIDENTIAL;
  }

  return CalendarApp.Visibility.DEFAULT;
}

/**
 * ICS の TRANSP 値を Google Calendar の透過設定に変換する。
 * @param {string} value TRANSP 値
 * @returns {GoogleAppsScript.Calendar.EventTransparency} 透過設定
 */
function parseTransparency(value) {
  const normalized = String(value || '')
    .trim()
    .toUpperCase();
  if (normalized === 'TRANSPARENT') {
    return CalendarApp.EventTransparency.TRANSPARENT;
  }
  return CalendarApp.EventTransparency.OPAQUE;
}

/**
 * BUSYSTATUS 値を Google Calendar の透過設定に変換する。
 * @param {string} value BUSYSTATUS 値
 * @returns {GoogleAppsScript.Calendar.EventTransparency} 透過設定
 */
function parseBusyStatus(value) {
  const normalized = String(value || '')
    .trim()
    .toUpperCase();
  if (normalized === 'FREE') {
    return CalendarApp.EventTransparency.TRANSPARENT;
  }
  return CalendarApp.EventTransparency.OPAQUE;
}

/**
 * 折り返し行を展開して ICS 行配列を返す。
 * @param {string} ics ICS 生テキスト
 * @returns {string[]} 展開後の行配列
 */
function unfoldICSLines(ics) {
  const rawLines = ics.split(/\r?\n/);
  const lines = [];

  rawLines.forEach((line) => {
    if (/^[ \t]/.test(line) && lines.length > 0) {
      lines[lines.length - 1] += line.slice(1);
    } else {
      lines.push(line);
    }
  });

  return lines;
}

/**
 * ICS の 1 行をプロパティ情報に分解する。
 * @param {string} line ICS の 1 行
 * @returns {{name: string, params: Object, value: string}|null} 分解結果
 */
function parsePropertyLine(line) {
  const separatorIndex = line.indexOf(':');
  if (separatorIndex === -1) {
    return null;
  }

  const left = line.substring(0, separatorIndex);
  const value = line.substring(separatorIndex + 1);
  const chunks = left.split(';');
  const name = chunks[0];
  const params = {};

  for (let i = 1; i < chunks.length; i++) {
    const eq = chunks[i].indexOf('=');
    if (eq === -1) {
      continue;
    }
    const key = chunks[i].substring(0, eq).toUpperCase();
    const paramValue = chunks[i].substring(eq + 1);
    params[key] = paramValue;
  }

  return { name: name.toUpperCase(), params: params, value: value };
}

/**
 * ICS の日付文字列を Date に変換する。
 * @param {string} value 例: 20260326T120000Z / 20251227
 * @param {Object} params ICS パラメータ
 * @returns {{date: Date, isAllDay: boolean}|null} 変換結果
 */
function parseICSDateValue(value, params) {
  const isAllDay = params.VALUE === 'DATE' || /^\d{8}$/.test(value);

  const m = value.match(
    /^(\d{4})(\d{2})(\d{2})(?:T(\d{2})(\d{2})(\d{2})?)?(Z)?$/,
  );
  if (!m) {
    return null;
  }

  const year = Number(m[1]);
  const month = Number(m[2]) - 1;
  const day = Number(m[3]);
  const hour = Number(m[4] || 0);
  const minute = Number(m[5] || 0);
  const second = Number(m[6] || 0);
  const isUtc = Boolean(m[7]);

  if (isAllDay) {
    return { date: new Date(year, month, day), isAllDay: true };
  }

  if (isUtc) {
    return {
      date: new Date(Date.UTC(year, month, day, hour, minute, second)),
      isAllDay: false,
    };
  }

  return {
    date: new Date(year, month, day, hour, minute, second),
    isAllDay: false,
  };
}

/**
 * 解析途中のイベント情報を同期用イベントに変換する。
 * @param {{uid?: string, title?: string, startInfo?: {date: Date, isAllDay: boolean}, endInfo?: {date: Date, isAllDay: boolean}, transparency?: GoogleAppsScript.Calendar.EventTransparency, visibility?: GoogleAppsScript.Calendar.Visibility}} source 中間イベント
 * @returns {{uid: string, title: string, start: Date, end: Date, isAllDay: boolean, transparency: GoogleAppsScript.Calendar.EventTransparency, visibility: GoogleAppsScript.Calendar.Visibility}|null} 同期用イベント
 */
function toParsedEvent(source) {
  if (!source || !source.uid || !source.startInfo || !source.startInfo.date) {
    return null;
  }

  const isAllDay = source.startInfo.isAllDay;
  const start = source.startInfo.date;
  let end = source.endInfo && source.endInfo.date ? source.endInfo.date : null;

  if (!end) {
    end = isAllDay
      ? addDays(start, 1)
      : new Date(start.getTime() + 60 * 60 * 1000);
  }

  return {
    uid: source.uid,
    title: source.title || '(無題)',
    start: start,
    end: end,
    isAllDay: isAllDay,
    transparency: source.transparency || CalendarApp.EventTransparency.OPAQUE,
    visibility: source.visibility || CalendarApp.Visibility.DEFAULT,
  };
}

/**
 * 日付に日数を加算する。
 * @param {Date} date 元の日付
 * @param {number} days 加算日数
 * @returns {Date} 加算後の日付
 */
function addDays(date, days) {
  return new Date(date.getTime() + days * 24 * 60 * 60 * 1000);
}

/**
 * 対象期間に重なるイベントのみ返す。
 * @param {{uid: string, title: string, start: Date, end: Date, isAllDay: boolean, transparency: GoogleAppsScript.Calendar.EventTransparency, visibility: GoogleAppsScript.Calendar.Visibility}[]} events イベント配列
 * @param {Date} rangeStart 範囲開始
 * @param {Date} rangeEnd 範囲終了
 * @returns {{uid: string, title: string, start: Date, end: Date, isAllDay: boolean, transparency: GoogleAppsScript.Calendar.EventTransparency, visibility: GoogleAppsScript.Calendar.Visibility}[]} フィルタ後イベント
 */
function filterEventsByRange(events, rangeStart, rangeEnd) {
  return events.filter(
    (event) => event.end > rangeStart && event.start < rangeEnd,
  );
}

/**
 * 管理対象(outlook_id付き)の Google Calendar イベントを UID マップ化する。
 * @param {GoogleAppsScript.Calendar.CalendarEvent[]} events Google Calendar イベント
 * @returns {Object.<string, GoogleAppsScript.Calendar.CalendarEvent>} UID -> 予定
 */
function buildManagedEventMap(events) {
  const map = {};
  events.forEach((event) => {
    const description = event.getDescription() || '';
    const match = description.match(/(?:^|\n)outlook_id:([^\n\r]+)/);
    if (match) {
      map[match[1].trim()] = event;
    }
  });
  return map;
}

/**
 * ICS 側イベントを UID マップ化する。
 * @param {{uid: string, title: string, start: Date, end: Date, isAllDay: boolean, transparency: GoogleAppsScript.Calendar.EventTransparency, visibility: GoogleAppsScript.Calendar.Visibility}[]} events ICS イベント
 * @returns {Object.<string, {uid: string, title: string, start: Date, end: Date, isAllDay: boolean, transparency: GoogleAppsScript.Calendar.EventTransparency, visibility: GoogleAppsScript.Calendar.Visibility}>} UID -> イベント
 */
function buildIncomingEventMap(events) {
  const map = {};
  events.forEach((event) => {
    map[event.uid] = event;
  });
  return map;
}

/**
 * 既存と新規の差分を作成する。
 * @param {Object.<string, GoogleAppsScript.Calendar.CalendarEvent>} existingMap 既存 UID マップ
 * @param {Object.<string, {uid: string, title: string, start: Date, end: Date, isAllDay: boolean, transparency: GoogleAppsScript.Calendar.EventTransparency, visibility: GoogleAppsScript.Calendar.Visibility}>} incomingMap 新規 UID マップ
 * @returns {{toCreate: Object[], toUpdate: Object[], toDelete: GoogleAppsScript.Calendar.CalendarEvent[]}} 差分
 */
function buildDiff(existingMap, incomingMap) {
  const toCreate = [];
  const toUpdate = [];
  const toDelete = [];

  Object.keys(incomingMap).forEach((uid) => {
    const incoming = incomingMap[uid];
    const existing = existingMap[uid];

    if (!existing) {
      toCreate.push(incoming);
      return;
    }

    if (!isSameEvent(existing, incoming)) {
      toUpdate.push({ existing: existing, incoming: incoming });
    }
  });

  Object.keys(existingMap).forEach((uid) => {
    if (!incomingMap[uid]) {
      toDelete.push(existingMap[uid]);
    }
  });

  return { toCreate: toCreate, toUpdate: toUpdate, toDelete: toDelete };
}

/**
 * Google Calendar イベントと ICS イベントが同一内容か判定する。
 * @param {GoogleAppsScript.Calendar.CalendarEvent} existing Google Calendar イベント
 * @param {{title: string, start: Date, end: Date, isAllDay: boolean, transparency: GoogleAppsScript.Calendar.EventTransparency, visibility: GoogleAppsScript.Calendar.Visibility}} incoming ICS イベント
 * @returns {boolean} 同一なら true
 */
function isSameEvent(existing, incoming) {
  if (existing.getTitle() !== incoming.title) {
    return false;
  }

  if (existing.isAllDayEvent() !== incoming.isAllDay) {
    return false;
  }

  if (existing.getTransparency() !== incoming.transparency) {
    return false;
  }

  if (existing.getVisibility() !== incoming.visibility) {
    return false;
  }

  if (incoming.isAllDay) {
    return (
      existing.getAllDayStartDate().getTime() === incoming.start.getTime() &&
      existing.getAllDayEndDate().getTime() === incoming.end.getTime()
    );
  }

  return (
    existing.getStartTime().getTime() === incoming.start.getTime() &&
    existing.getEndTime().getTime() === incoming.end.getTime()
  );
}

/**
 * 差分を Google Calendar に反映する。
 * @param {GoogleAppsScript.Calendar.Calendar} calendar 対象カレンダー
 * @param {{toCreate: Object[], toUpdate: Object[], toDelete: GoogleAppsScript.Calendar.CalendarEvent[]}} diff 差分
 * @returns {void}
 */
function applyDiff(calendar, diff) {
  diff.toCreate.forEach((event) => {
    createManagedEvent(calendar, event);
  });

  diff.toUpdate.forEach((item) => {
    updateManagedEvent(calendar, item.existing, item.incoming);
  });

  diff.toDelete.forEach((event) => {
    event.deleteEvent();
  });
}

/**
 * Outlook から取り込んだイベントを除いた Google Calendar イベントを返す。
 * @param {GoogleAppsScript.Calendar.CalendarEvent[]} events Google Calendar イベント
 * @returns {GoogleAppsScript.Calendar.CalendarEvent[]} 同期対象イベント
 */
function buildGoogleSyncCandidates(events) {
  return events.filter((event) => !hasOutlookSyncMarker(event));
}

/**
 * Google Calendar イベントが Outlook 由来か判定する。
 * @param {GoogleAppsScript.Calendar.CalendarEvent} event Google Calendar イベント
 * @returns {boolean} Outlook 由来なら true
 */
function hasOutlookSyncMarker(event) {
  const description = event.getDescription() || '';
  return /(?:^|\n)outlook_id:[^\n\r]+/.test(description);
}

/**
 * 管理対象イベントを新規作成する。
 * @param {GoogleAppsScript.Calendar.Calendar} calendar 対象カレンダー
 * @param {{uid: string, title: string, start: Date, end: Date, isAllDay: boolean, transparency: GoogleAppsScript.Calendar.EventTransparency, visibility: GoogleAppsScript.Calendar.Visibility}} event 作成イベント
 * @returns {GoogleAppsScript.Calendar.CalendarEvent} 作成したイベント
 */
function createManagedEvent(calendar, event) {
  let created;
  if (event.isAllDay) {
    created = calendar.createAllDayEvent(event.title, event.start, event.end, {
      description: 'outlook_id:' + event.uid,
    });
  } else {
    created = calendar.createEvent(event.title, event.start, event.end, {
      description: 'outlook_id:' + event.uid,
    });
  }

  applyEventSettings(created, event.transparency, event.visibility);
  return created;
}

/**
 * イベントの通知設定と空き時間/予定あり設定を反映する。
 * @param {GoogleAppsScript.Calendar.CalendarEvent} event Google Calendar イベント
 * @param {GoogleAppsScript.Calendar.EventTransparency} transparency 透過設定
 * @param {GoogleAppsScript.Calendar.Visibility} visibility 公開設定
 * @returns {void}
 */
function applyEventSettings(event, transparency, visibility) {
  event.removeAllReminders();
  event.setTransparency(transparency);
  event.setVisibility(visibility);
}

/**
 * 管理対象イベントを更新する。
 * 型（終日/時間指定）が変わった場合は再作成する。
 * @param {GoogleAppsScript.Calendar.Calendar} calendar 対象カレンダー
 * @param {GoogleAppsScript.Calendar.CalendarEvent} existing 既存イベント
 * @param {{uid: string, title: string, start: Date, end: Date, isAllDay: boolean, transparency: GoogleAppsScript.Calendar.EventTransparency, visibility: GoogleAppsScript.Calendar.Visibility}} incoming 新しいイベント内容
 * @returns {void}
 */
function updateManagedEvent(calendar, existing, incoming) {
  if (existing.isAllDayEvent() !== incoming.isAllDay) {
    existing.deleteEvent();
    createManagedEvent(calendar, incoming);
    return;
  }

  existing.setTitle(incoming.title);
  existing.setDescription('outlook_id:' + incoming.uid);

  if (incoming.isAllDay) {
    existing.setAllDayDates(incoming.start, incoming.end);
  } else {
    existing.setTime(incoming.start, incoming.end);
  }

  applyEventSettings(existing, incoming.transparency, incoming.visibility);
}
