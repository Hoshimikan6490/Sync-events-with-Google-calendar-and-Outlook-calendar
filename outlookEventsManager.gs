const outlookCalendarId = 'YOUR_OUTLOOK_CALENDAR_ID';
const clientId = 'YOUR_CLIENT_ID';
const clientSecret = 'YOUR_CLIENT_SECRET';
const tenantId = 'consumers'; // DO NOT CHANGE
const redirectUri =
  'https://login.microsoftonline.com/common/oauth2/nativeclient'; // DO NOT CHANGE
const tokenUrl =
  'https://login.microsoftonline.com/' + tenantId + '/oauth2/v2.0/token'; // DO NOT CHANGE
const authCode = 'YOUR_ACCESS_CODE'; //write CODE using setup function

/**
 * OAuth 認可 URL を生成し、PKCE 用の code_verifier を保存する。
 * @returns {void}
 */
function setup() {
  var codeVerifier = generateCodeVerifier();
  var codeChallenge = generateCodeChallenge(codeVerifier);

  PropertiesService.getScriptProperties().setProperty(
    'oauth_code_verifier',
    codeVerifier,
  );

  var url =
    'https://login.microsoftonline.com/' +
    tenantId +
    '/oauth2/v2.0/authorize' +
    '?client_id=' +
    clientId +
    '&response_type=code' +
    '&redirect_uri=' +
    encodeURIComponent(redirectUri) +
    '&scope=' +
    encodeURIComponent(
      'offline_access https://graph.microsoft.com/Calendars.ReadWrite',
    ) +
    '&response_mode=query' +
    '&code_challenge=' +
    encodeURIComponent(codeChallenge) +
    '&code_challenge_method=S256';

  Logger.log('このURLを開いて認証して👇');
  Logger.log(url);
}

/**
 * 認可コードをアクセストークンと交換し、ScriptProperties に保存する。
 * @returns {void}
 */
function authCallback() {
  var codeVerifier = PropertiesService.getScriptProperties().getProperty(
    'oauth_code_verifier',
  );

  if (!codeVerifier) {
    throw new Error(
      'oauth_code_verifier がありません。先に setup() を実行して認可URLを再生成してください。',
    );
  }

  var payload = {
    client_id: clientId,
    code: authCode,
    redirect_uri: redirectUri,
    grant_type: 'authorization_code',
    code_verifier: codeVerifier,
  };

  var options = {
    method: 'post',
    payload: payload,
    muteHttpExceptions: true,
  };

  var res = UrlFetchApp.fetch(tokenUrl, options);
  var body = res.getContentText();
  var status = res.getResponseCode();

  if (status >= 400) {
    throw new Error('Token exchange failed (' + status + '): ' + body);
  }

  var data = JSON.parse(body);

  Logger.log(data);

  // refresh_token は毎回返るとは限らないため、存在時のみ更新する。
  if (data.refresh_token) {
    PropertiesService.getScriptProperties().setProperty(
      'refresh_token',
      data.refresh_token,
    );
  }

  PropertiesService.getScriptProperties().setProperty(
    'access_token',
    data.access_token,
  );
}

/**
 * PKCE 用の code_verifier を生成する。
 * @returns {string} 64 文字の code_verifier
 */
function generateCodeVerifier() {
  var bytes =
    Utilities.getUuid().replace(/-/g, '') +
    Utilities.getUuid().replace(/-/g, '');
  return bytes.slice(0, 64);
}

/**
 * code_verifier から code_challenge を生成する。
 * @param {string} codeVerifier PKCE の code_verifier
 * @returns {string} base64url 形式の code_challenge
 */
function generateCodeChallenge(codeVerifier) {
  var digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    codeVerifier,
    Utilities.Charset.UTF_8,
  );
  return base64UrlEncode(digest);
}

function base64UrlEncode(bytes) {
  return Utilities.base64Encode(bytes)
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/, '');
}

/**
 * 保存済み refresh_token を使って access_token を更新する。
 * @returns {string} 更新後の access_token
 */
function refreshAccessToken() {
  var url = 'https://login.microsoftonline.com/consumers/oauth2/v2.0/token';

  var refreshToken =
    PropertiesService.getScriptProperties().getProperty('refresh_token');

  if (!refreshToken) {
    throw new Error(
      'refresh_token がありません。先に authCallback() を実行してトークンを保存してください。',
    );
  }

  var payload = {
    client_id: clientId,
    refresh_token: refreshToken,
    grant_type: 'refresh_token',
  };

  var options = {
    method: 'post',
    payload: payload,
    muteHttpExceptions: true,
  };

  var res = UrlFetchApp.fetch(url, options);
  var body = res.getContentText();
  var status = res.getResponseCode();

  if (status >= 400) {
    throw new Error('Refresh token failed (' + status + '): ' + body);
  }

  var data = JSON.parse(body);

  // トークン更新保存
  PropertiesService.getScriptProperties().setProperty(
    'access_token',
    data.access_token,
  );

  if (data.refresh_token) {
    PropertiesService.getScriptProperties().setProperty(
      'refresh_token',
      data.refresh_token,
    );
  }

  return data.access_token;
}

/**
 * Graph API へイベントを作成する。
 * @param {Object} [eventOptions] 作成するイベントの設定
 * @param {string} [eventOptions.subject] 件名
 * @param {string} [eventOptions.title] 件名の別名
 * @param {string|Date} [eventOptions.start] 開始日時
 * @param {string|Date} [eventOptions.end] 終了日時
 * @param {string} [eventOptions.timeZone='Asia/Tokyo'] タイムゾーン
 * @param {boolean} [eventOptions.isAllDay=false] 終日イベントかどうか
 * @param {{contentType: string, content: string}} [eventOptions.body] 本文
 * @param {string} [eventOptions.location] 場所
 * @param {{email: string, name?: string, type?: string}[]} [eventOptions.attendees] 参加者
 * @param {string[]} [eventOptions.categories] カテゴリ
 * @param {string} [eventOptions.calendarId] 作成先カレンダーID
 * @param {string} [eventOptions.showAs] 表示状態
 * @param {string} [eventOptions.sensitivity] 機密度
 * @param {string} [eventOptions.importance] 重要度
 * @param {boolean} [eventOptions.isReminderOn] リマインダー有効化
 * @param {number} [eventOptions.reminderMinutesBeforeStart] リマインダー分数
 * @returns {Object} Microsoft Graph の作成結果
 */
function createEventToOutlook(eventOptions, accessToken) {
  var token = accessToken || refreshAccessToken();

  var normalized = normalizeEventOptions(eventOptions);
  var calendarId = normalized.calendarId || outlookCalendarId;
  var url = calendarId
    ? 'https://graph.microsoft.com/v1.0/me/calendars/' +
      encodeURIComponent(calendarId) +
      '/events'
    : 'https://graph.microsoft.com/v1.0/me/events';

  var payload = {
    subject: normalized.subject,
    start: {
      dateTime: formatGraphDateTime(normalized.start, normalized.timeZone),
      timeZone: normalized.timeZone,
    },
    end: {
      dateTime: formatGraphDateTime(normalized.end, normalized.timeZone),
      timeZone: normalized.timeZone,
    },
  };

  if (normalized.body) {
    payload.body = {
      contentType: normalized.body.contentType,
      content: normalized.body.content,
    };
  }

  if (normalized.location) {
    payload.location = {
      displayName: normalized.location,
    };
  }

  if (normalized.isAllDay) {
    payload.isAllDay = true;
  }

  if (normalized.attendees.length > 0) {
    payload.attendees = normalized.attendees.map(function (attendee) {
      return {
        emailAddress: {
          address: attendee.email,
          name: attendee.name || attendee.email,
        },
        type: attendee.type || 'required',
      };
    });
  }

  if (normalized.categories.length > 0) {
    payload.categories = normalized.categories;
  }

  if (normalized.showAs) {
    payload.showAs = normalized.showAs;
  }

  if (normalized.sensitivity) {
    payload.sensitivity = normalized.sensitivity;
  }

  if (normalized.importance) {
    payload.importance = normalized.importance;
  }

  if (typeof normalized.isReminderOn === 'boolean') {
    payload.isReminderOn = normalized.isReminderOn;
  }

  if (typeof normalized.reminderMinutesBeforeStart === 'number') {
    payload.reminderMinutesBeforeStart = normalized.reminderMinutesBeforeStart;
  }

  var requestOptions = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + token,
    },
    payload: JSON.stringify(cleanPayload(payload)),
    muteHttpExceptions: true,
  };

  var response = UrlFetchApp.fetch(url, requestOptions);
  var responseBody = response.getContentText();

  if (response.getResponseCode() >= 400) {
    throw new Error(
      'Event creation failed (' +
        response.getResponseCode() +
        '): ' +
        responseBody,
    );
  }

  return JSON.parse(responseBody);
}

/**
 * Outlook の指定期間イベントを取得する。
 * @param {Date} rangeStart 取得開始日時
 * @param {Date} rangeEnd 取得終了日時
 * @returns {Object[]} Outlook イベント一覧
 */
function listOutlookEventsInRange(rangeStart, rangeEnd) {
  var token = refreshAccessToken();
  var targetCalendarId =
    typeof outlookCalendarId === 'string' ? outlookCalendarId : '';
  var calendarPath = targetCalendarId
    ? '/me/calendars/' + encodeURIComponent(targetCalendarId) + '/calendarView'
    : '/me/calendarView';
  var url =
    'https://graph.microsoft.com/v1.0' +
    calendarPath +
    '?startDateTime=' +
    encodeURIComponent(formatGraphUtcDateTime(rangeStart)) +
    '&endDateTime=' +
    encodeURIComponent(formatGraphUtcDateTime(rangeEnd)) +
    '&$select=id,subject,start,end,isAllDay,body,location,showAs';

  var events = [];

  while (url) {
    var response = fetchGraphJson(url, token);
    if (response.value && response.value.length > 0) {
      events = events.concat(response.value);
    }

    url = response['@odata.nextLink'] || null;
  }

  return events;
}

/**
 * Google 由来の Outlook イベント ID を集める。
 * @param {Object[]} outlookEvents Outlook イベント一覧
 * @returns {Object.<string, boolean>} Google イベント ID -> true
 */
function buildOutlookGoogleIdSet(outlookEvents) {
  var set = {};

  outlookEvents.forEach(function (event) {
    var googleEventId = extractGoogleEventIdFromOutlookEvent(event);
    if (googleEventId) {
      set[googleEventId] = true;
    }
  });

  return set;
}

/**
 * Outlook の body から Google イベント ID を取り出す。
 * @param {Object} outlookEvent Outlook イベント
 * @returns {string|null} Google イベント ID
 */
function extractGoogleEventIdFromOutlookEvent(outlookEvent) {
  var bodyContent =
    outlookEvent && outlookEvent.body && outlookEvent.body.content
      ? String(outlookEvent.body.content)
      : '';

  var match = findGoogleIdMarker(bodyContent);
  return match ? match : null;
}

/**
 * Outlook 本文(HTML/テキスト)から google_id マーカーを抽出する。
 * @param {string} content 本文
 * @returns {string|null} Google イベント ID
 */
function findGoogleIdMarker(content) {
  if (!content) {
    return null;
  }

  var normalized = String(content)
    .replace(/<br\s*\/?\s*>/gi, '\n')
    .replace(/<\/(?:div|p|li|tr|h[1-6])>/gi, '\n')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/\r/g, '');

  var match = normalized.match(/google_id:([^\s<]+@google\.com)\b/i);
  var googleEventId = match ? match[1].trim() : null;

  return googleEventId;
}

/**
 * Google イベント ID を google.com サフィックス付きへ正規化する。
 * @param {string} googleEventId Google イベント ID
 * @returns {string} google_id 用の統一ID
 */
function toCanonicalGoogleEventId(googleEventId) {
  var id = String(googleEventId || '').trim();
  if (!id) {
    return id;
  }

  return /@google\.com$/i.test(id) ? id : id + '@google.com';
}

/**
 * Outlook 由来の Google イベント ID を集める。
 * @param {GoogleAppsScript.Calendar.CalendarEvent[]} googleEvents Google Calendar イベント一覧
 * @returns {Object.<string, boolean>} Outlook イベント ID -> true
 */
function buildGoogleOutlookIdSet(googleEvents) {
  var set = {};

  googleEvents.forEach(function (event) {
    var outlookEventId = extractOutlookEventIdFromGoogleEvent(event);
    if (outlookEventId) {
      set[outlookEventId] = true;
    }
  });

  return set;
}

/**
 * Google Calendar イベントの description から Outlook イベント ID を取り出す。
 * @param {GoogleAppsScript.Calendar.CalendarEvent} googleEvent Google Calendar イベント
 * @returns {string|null} Outlook イベント ID
 */
function extractOutlookEventIdFromGoogleEvent(googleEvent) {
  var description =
    googleEvent && googleEvent.getDescription
      ? googleEvent.getDescription()
      : '';
  if (!description) {
    return null;
  }

  var match = String(description).match(/(?:^|\n)outlook_id:([^\n\r]+)/);
  return match ? match[1].trim() : null;
}

/**
 * Google Calendar のイベントを Outlook 作成用のパラメータへ変換する。
 * @param {GoogleAppsScript.Calendar.CalendarEvent} event Google Calendar イベント
 * @returns {Object} createEventToOutlook 用のオプション
 */
function convertGoogleEventToOutlookOptions(event) {
  var calendar = CalendarApp.getDefaultCalendar();
  var timeZone =
    typeof calendar.getTimeZone === 'function'
      ? calendar.getTimeZone()
      : Session.getScriptTimeZone();
  var isAllDay = event.isAllDayEvent();
  var start = isAllDay ? event.getAllDayStartDate() : event.getStartTime();
  var end = isAllDay ? event.getAllDayEndDate() : event.getEndTime();
  var description = event.getDescription() || '';
  var bodyContent = description ? description + '\n\n' : '';

  bodyContent += 'google_id:' + toCanonicalGoogleEventId(event.getId());

  return {
    subject: ensureManagedOutlookSubjectPrefix(event.getTitle()),
    start: start,
    end: end,
    timeZone: timeZone,
    isAllDay: isAllDay,
    body: {
      contentType: 'text',
      content: bodyContent,
    },
    location: event.getLocation() || '',
    showAs: mapGoogleTransparencyToOutlookShowAs(event.getTransparency()),
  };
}

/**
 * 管理用プレフィックスが未付与なら件名先頭に付与する。
 * @param {string} title 元の件名
 * @returns {string} プレフィックス付与後の件名
 */
function ensureManagedOutlookSubjectPrefix(title) {
  var normalizedTitle = title ? String(title) : '(無題)';

  if (
    typeof MANAGED_OUTLOOK_SUBJECT_PREFIX === 'string' &&
    MANAGED_OUTLOOK_SUBJECT_PREFIX.length > 0 &&
    normalizedTitle.indexOf(MANAGED_OUTLOOK_SUBJECT_PREFIX) !== 0
  ) {
    return MANAGED_OUTLOOK_SUBJECT_PREFIX + normalizedTitle;
  }

  return normalizedTitle;
}

/**
 * Google Calendar の transparency を Outlook の showAs に変換する。
 * @param {GoogleAppsScript.Calendar.EventTransparency} transparency Google Calendar の透過設定
 * @returns {string} Outlook の表示状態
 */
function mapGoogleTransparencyToOutlookShowAs(transparency) {
  if (transparency === CalendarApp.EventTransparency.TRANSPARENT) {
    return 'free';
  }

  return 'busy';
}

/**
 * Graph API へ GET リクエストを送り、JSON を返す。
 * @param {string} url Graph API URL
 * @param {string} token アクセストークン
 * @returns {Object} 解析済みレスポンス
 */
function fetchGraphJson(url, token) {
  var response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: {
      Authorization: 'Bearer ' + token,
    },
    muteHttpExceptions: true,
  });

  var body = response.getContentText();
  if (response.getResponseCode() >= 400) {
    throw new Error(
      'Graph request failed (' + response.getResponseCode() + '): ' + body,
    );
  }

  return JSON.parse(body);
}

/**
 * Graph API 向けに UTC の日時文字列へ整形する。
 * @param {Date} date 対象日時
 * @returns {string} Graph API 用の UTC 文字列
 */
function formatGraphUtcDateTime(date) {
  if (!(date instanceof Date) || isNaN(date.getTime())) {
    throw new Error(
      '日時が不正です。Date オブジェクトか有効な日付を指定してください。',
    );
  }

  return Utilities.formatDate(date, 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'");
}

/**
 * createEventToOutlook 用に入力値を正規化する。
 * @param {Object} [eventOptions] 元のイベント設定
 * @returns {{subject: string, body: Object|null, location: string, start: Date, end: Date, timeZone: string, isAllDay: boolean, attendees: Object[], categories: string[], calendarId: string, showAs: string, sensitivity: string, importance: string, isReminderOn: (boolean|undefined), reminderMinutesBeforeStart: (number|undefined)}} 正規化後の設定
 */
function normalizeEventOptions(eventOptions) {
  var now = new Date();
  var start =
    eventOptions && eventOptions.start ? new Date(eventOptions.start) : now;
  var end =
    eventOptions && eventOptions.end
      ? new Date(eventOptions.end)
      : new Date(start.getTime() + 60 * 60 * 1000);

  return {
    subject: eventOptions && (eventOptions.subject || eventOptions.title),
    body: eventOptions && eventOptions.body ? eventOptions.body : null,
    location:
      eventOptions && eventOptions.location
        ? String(eventOptions.location)
        : '',
    start: start,
    end: end,
    timeZone: (eventOptions && eventOptions.timeZone) || 'Asia/Tokyo',
    isAllDay: Boolean(eventOptions && eventOptions.isAllDay),
    attendees:
      eventOptions && Array.isArray(eventOptions.attendees)
        ? eventOptions.attendees
        : [],
    categories:
      eventOptions && Array.isArray(eventOptions.categories)
        ? eventOptions.categories
        : [],
    calendarId:
      eventOptions && eventOptions.calendarId
        ? String(eventOptions.calendarId)
        : '',
    showAs: eventOptions && eventOptions.showAs ? eventOptions.showAs : '',
    sensitivity:
      eventOptions && eventOptions.sensitivity ? eventOptions.sensitivity : '',
    importance:
      eventOptions && eventOptions.importance ? eventOptions.importance : '',
    isReminderOn:
      eventOptions && typeof eventOptions.isReminderOn === 'boolean'
        ? eventOptions.isReminderOn
        : undefined,
    reminderMinutesBeforeStart:
      eventOptions &&
      typeof eventOptions.reminderMinutesBeforeStart === 'number'
        ? eventOptions.reminderMinutesBeforeStart
        : undefined,
  };
}

/**
 * Graph API 向けに Date を yyyy-MM-dd'T'HH:mm:ss へ整形する。
 * @param {Date} date 対象日時
 * @param {string} timeZone 出力に使うタイムゾーン
 * @returns {string} Graph API 用の日時文字列
 */
function formatGraphDateTime(date, timeZone) {
  if (!(date instanceof Date) || isNaN(date.getTime())) {
    throw new Error(
      'イベント日時が不正です。Date オブジェクトか有効な日付を指定してください。',
    );
  }

  return Utilities.formatDate(date, timeZone, "yyyy-MM-dd'T'HH:mm:ss");
}

/**
 * JSON シリアライズ前に undefined / null / 空文字を取り除く。
 * @param {*} value 入力値
 * @returns {*} 取り除き済みの値
 */
function cleanPayload(value) {
  if (Array.isArray(value)) {
    return value.map(cleanPayload);
  }

  if (value && typeof value === 'object') {
    var cleaned = {};
    Object.keys(value).forEach(function (key) {
      if (
        value[key] !== undefined &&
        value[key] !== null &&
        value[key] !== ''
      ) {
        cleaned[key] = cleanPayload(value[key]);
      }
    });
    return cleaned;
  }

  return value;
}
