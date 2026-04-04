const MANAGED_OUTLOOK_SUBJECT_PREFIX = '[研究室] ';
const ONE_DAY_IN_MILLISECONDS = 24 * 60 * 60 * 1000;

/**
 * Outlook の ICS を Google Calendar に同期する。
 * 対象期間は「前日から 1 か月後まで」。
 * @returns {{created: number, updated: number, deleted: number}} 同期結果
 */
function syncOutlookToGoogle() {
  const range = getSyncRange();
  const icsUrl = getICSUrl();
  const icsText = UrlFetchApp.fetch(icsUrl).getContentText();
  const icsEvents = filterEventsByRange(
    parseICS(icsText),
    range.start,
    range.end,
  ).filter((event) => !hasManagedOutlookSubjectPrefix(event.title));

  const calendar = CalendarApp.getDefaultCalendar();
  const existingEvents = calendar.getEvents(range.start, range.end);

  const existingMap = buildManagedEventMap(existingEvents);
  const incomingMap = buildIncomingEventMap(icsEvents);
  const diff = buildDiff(existingMap, incomingMap);

  applyDiff(calendar, diff);

  Logger.log(
    'Google sync finished. created=%d updated=%d deleted=%d',
    diff.toCreate.length,
    diff.toUpdate.length,
    diff.toDelete.length,
  );

  return {
    created: diff.toCreate.length,
    updated: diff.toUpdate.length,
    deleted: diff.toDelete.length,
  };
}

/**
 * Google Calendar の未同期イベントを Outlook に作成する。
 * Outlook から取り込んだイベントは description の outlook_id で除外する。
 * @returns {{created: number, skipped: number}} 同期結果
 */
function syncGoogleToOutlook() {
  const range = getSyncRange();
  const calendar = CalendarApp.getDefaultCalendar();
  const events = calendar.getEvents(range.start, range.end);
  const candidates = buildGoogleSyncCandidates(events);
  const existingOutlookEvents = listOutlookEventsInRange(
    range.start,
    range.end,
  );
  const existingGoogleIds = buildOutlookGoogleIdSet(existingOutlookEvents);
  const accessToken = refreshAccessToken();

  let created = 0;
  let skipped = 0;

  candidates.forEach((event) => {
    const googleEventId = toCanonicalGoogleEventId(event.getId());
    const matched = Boolean(existingGoogleIds[googleEventId]);

    if (matched) {
      skipped += 1;
      return;
    }

    createEventToOutlook(
      convertGoogleEventToOutlookOptions(event),
      accessToken,
    );
    created += 1;
  });

  Logger.log('Outlook sync finished. created=%d skipped=%d', created, skipped);

  return {
    created: created,
    skipped: skipped,
  };
}

/**
 * 1 か月分の双方向同期を実行する。
 * まず Outlook -> Google を反映し、その後 Google -> Outlook を反映する。
 * @returns {{googleToOutlook: {created: number, skipped: number}, outlookToGoogle: {created: number, updated: number, deleted: number}}}
 */
function syncMonthlyCalendars() {
  const outlookToGoogle = syncOutlookToGoogle();
  const googleToOutlook = syncGoogleToOutlook();

  return {
    outlookToGoogle: outlookToGoogle,
    googleToOutlook: googleToOutlook,
  };
}

/**
 * 同期対象期間を返す。
 * @returns {{start: Date, end: Date}} 開始日時と終了日時
 */
function getSyncRange() {
  const now = new Date();
  return {
    start: new Date(now.getTime() - ONE_DAY_IN_MILLISECONDS),
    end: addMonths(now, 1),
  };
}

/**
 * 日付に月数を加算する。
 * @param {Date} date 元の日付
 * @param {number} months 加算月数
 * @returns {Date} 加算後の日付
 */
function addMonths(date, months) {
  const result = new Date(date.getTime());
  result.setMonth(result.getMonth() + months);
  return result;
}

/**
 * スプレッドシートからICS URLを取得する
 * @returns {string} Outlook ICS URL
 */
function getICSUrl() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const icsUrl = String(sheet.getRange('A1').getValue() || '').trim();

  if (!icsUrl) {
    throw new Error('セルA1にICS URLが見つかりません');
  }

  return icsUrl;
}

/**
 * 件名が管理用プレフィックス付きか判定する。
 * @param {string} title 件名
 * @returns {boolean} 管理用プレフィックス付きなら true
 */
function hasManagedOutlookSubjectPrefix(title) {
  return String(title || '').indexOf(MANAGED_OUTLOOK_SUBJECT_PREFIX) === 0;
}
