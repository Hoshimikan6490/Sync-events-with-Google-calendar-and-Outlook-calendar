/**
 * 指定期間の Google カレンダーイベントを取得して正規化して返す。
 * occurrence 単位取得のため、singleEvents は true に設定。
 * @param {Date} startDate 取得開始日時
 * @param {Date} endDate 取得終了日時
 * @returns {Array<Object>} 正規化されたイベント配列（occurrence 単位）
 */
function getGoogleEvents(startDate, endDate) {
	const calendarId = CalendarApp.getDefaultCalendar().getId();
	const timeMin = buildGoogleApiQueryDateTime_(startDate);
	const timeMax = buildGoogleApiQueryDateTime_(endDate);
	const events =
		Calendar.Events.list(calendarId, {
			timeMin: timeMin,
			timeMax: timeMax,
			singleEvents: true,
			orderBy: 'startTime',
		}).items || [];

	return events
		.map(normalizeGoogleCalendarEvent_)
		.filter((event) => googleEventOverlapsWindow_(event, startDate, endDate));
}

/**
 * Google イベントが指定ウィンドウと重複するか判定する。
 * @param {Object} event Google イベント
 * @param {Date} startDate ウィンドウ開始
 * @param {Date} endDate ウィンドウ終了
 * @returns {boolean} 重複する場合は true
 */
function googleEventOverlapsWindow_(event, startDate, endDate) {
	const start = event.start && (event.start.dateTime || event.start.date);
	const end = event.end && (event.end.dateTime || event.end.date);
	if (!start || !end) {
		return false;
	}

	const eventStart = new Date(start).getTime();
	const eventEnd = new Date(end).getTime();
	return eventEnd >= startDate.getTime() && eventStart <= endDate.getTime();
}

/**
 * Google カレンダーに新しいイベントを作成する。
 * @param {Object} eventData 作成するイベントデータ（summary, start, end 等）
 * @returns {Object} 作成されたイベントのリソース
 */
function createGoogleEvent(eventData) {
	const calendarId = CalendarApp.getDefaultCalendar().getId();
	const resource = buildGoogleCalendarResource_(eventData);
	return Calendar.Events.insert(resource, calendarId);
}

/**
 * 既存の Google イベントを更新する。
 * @param {string} eventId 更新対象のイベント ID
 * @param {Object} eventData 更新内容のイベントデータ
 * @returns {Object} 更新後のイベントリソース
 */
function updateGoogleEvent(eventId, eventData) {
	const calendarId = CalendarApp.getDefaultCalendar().getId();
	const resource = buildGoogleCalendarResource_(eventData);
	return Calendar.Events.update(resource, calendarId, eventId);
}

/**
 * Google カレンダーからイベントを削除する。
 * @param {string} eventId 削除対象のイベント ID
 * @returns void
 */
function deleteGoogleEvent(eventId) {
	const calendarId = CalendarApp.getDefaultCalendar().getId();
	return Calendar.Events.remove(calendarId, eventId);
}

/**
 * 説明文から埋め込まれた google_id を抽出する。
 * @param {string} description イベント説明文
 * @returns {string} 抽出された google_id（存在しない場合は空文字）
 */
function extractGoogleEventId(description) {
	const ids = parseIds(description);
	return ids.googleId;
}

/**
 * Google イベントの説明文を組み立てる（outlook_id と outlookSyncKey を含める）。
 * @param {Object} event 元イベントオブジェクト（description を使用）
 * @param {string} outlookId Outlook のイベント ID
 * @param {string} outlookSyncKey Outlook 側の同期キー
 * @returns {string} 組み立てた説明文
 */
function buildGoogleDescription(event, outlookId, outlookSyncKey) {
	const lines = [];
	const descriptionLines = String(event.description || '')
		.split(/\r?\n/)
		.filter(
			(line) =>
				line && !/^outlook_id:/i.test(line) && !/^outlooksynckey:/i.test(line),
		);

	if (descriptionLines.length > 0) {
		lines.push(descriptionLines.join('\n').trim());
	}
	if (outlookSyncKey) {
		lines.push(`outlookSyncKey:${outlookSyncKey}`);
	}
	if (outlookId) {
		lines.push(`outlook_id:${outlookId}`);
	}
	return lines.join('\n');
}

/**
 * Google Calendar API のイベントリソースを内部で扱う正規化形式（UTC + timeZone）に変換する。
 * Google のdateTime（RFC3339 with offset）をUTC + timeZoneに正規化する。
 * @param {Object} event API のイベントオブジェクト
 * @returns {Object} 正規化されたイベントオブジェクト
 */
function normalizeGoogleCalendarEvent_(event) {
	const startNormalized = normalizeGoogleDateTime_(event.start);
	const endNormalized = normalizeGoogleDateTime_(event.end);

	// 内部表現は Outlook 仕様をメインとする（subject, showAs, sensitivity など）
	const isAllDay = Boolean(startNormalized.start && startNormalized.start.date);
	return {
		id: event.id,
		subject: event.summary || '',
		description: event.description || '',
		// location は Outlook 側で使う文字列形式を採用
		location:
			event.location || (event.location && event.location.displayName) || '',
		// start/end は既に正規化済みのオブジェクトをそのまま使う
		start: startNormalized.start || {},
		end: endNormalized.start || {},
		isAllDay: isAllDay,
		// Google の透明性/表示設定を Outlook の showAs/sensitivity に変換
		showAs: mapTransparencyToShowAs(event.transparency || 'opaque'),
		sensitivity: mapVisibilityToSensitivity(event.visibility || 'default'),
		// occurrence 識別用フィールド
		recurringEventId: event.recurringEventId || null,
		originalStartTime: event.originalStartTime || null,
		occurrenceDate: normalizeOccurrenceDateText_(
			event.originalStartTime || event.start || event.startDateTime || null,
		),
		raw: event,
	};
}

/**
 * Google Calendar API の datetime オブジェクトをUTC + timeZone形式に正規化する。
 * RFC3339 with offset の形式をUTC+タイムゾーン形式に変換する。
 * @param {Object} googleDateTime Google の start/end オブジェクト {dateTime, date, timeZone}
 * @returns {Object} 正規化されたオブジェクト {dateTime（UTC）, timeZone, start}
 */
function normalizeGoogleDateTime_(googleDateTime) {
	if (!googleDateTime) {
		return { dateTime: '', timeZone: SYNC_TIMEZONE, start: {} };
	}

	// 全日イベント: date のみ
	if (googleDateTime.date && !googleDateTime.dateTime) {
		return {
			dateTime: '',
			timeZone: SYNC_TIMEZONE,
			start: { date: googleDateTime.date },
		};
	}

	// 時間指定イベント: dateTime（RFC3339 with offset）をUTCに正規化
	if (googleDateTime.dateTime) {
		const utcDateTime = convertRfc3339ToUtc_(googleDateTime.dateTime);
		return {
			dateTime: utcDateTime,
			timeZone: googleDateTime.timeZone || SYNC_TIMEZONE,
			start: {
				dateTime: utcDateTime,
				timeZone: googleDateTime.timeZone || SYNC_TIMEZONE,
			},
		};
	}

	return { dateTime: '', timeZone: SYNC_TIMEZONE, start: {} };
}

/**
 * Google Calendar API に渡すリソースオブジェクトを構築する。
 * 内部のUTC + timeZone 形式をGoogle形式に変換する。
 * @param {Object} eventData 内部表現のイベントデータ（UTC + timeZone）
 * @returns {Object} API に渡すリソースオブジェクト（Google形式）
 */
function buildGoogleCalendarResource_(eventData) {
	// eventData は Outlook 仕様（subject, showAs, sensitivity, description, location）で渡される前提
	const resource = {
		summary: eventData.subject || '',
		description: eventData.description || '',
		transparency: mapShowAsToTransparency(eventData.showAs) || 'opaque',
		visibility: mapSensitivityToVisibility(eventData.sensitivity) || 'default',
	};

	const timeZone = eventData.timeZone || SYNC_TIMEZONE;

	// location (Google は単純な文字列)。内部は Outlook 仕様の文字列を想定する
	if (eventData.location) {
		if (typeof eventData.location === 'string') {
			resource.location = eventData.location;
		} else if (eventData.location.displayName) {
			resource.location = eventData.location.displayName;
		}
	}
	if (eventData.start && eventData.start.date) {
		resource.start = { date: eventData.start.date, timeZone: timeZone };
	} else if (eventData.start && eventData.start.dateTime) {
		resource.start = {
			dateTime: convertUtcToLocalDateTime_(eventData.start.dateTime, timeZone),
			timeZone: timeZone,
		};
	} else if (eventData.startDateTime) {
		resource.start = {
			dateTime: convertUtcToLocalDateTime_(eventData.startDateTime, timeZone),
			timeZone: timeZone,
		};
	}

	if (eventData.end && eventData.end.date) {
		resource.end = { date: eventData.end.date, timeZone: timeZone };
	} else if (eventData.end && eventData.end.dateTime) {
		resource.end = {
			dateTime: convertUtcToLocalDateTime_(eventData.end.dateTime, timeZone),
			timeZone: timeZone,
		};
	} else if (eventData.endDateTime) {
		resource.end = {
			dateTime: convertUtcToLocalDateTime_(eventData.endDateTime, timeZone),
			timeZone: timeZone,
		};
	}

	if (!resource.end && resource.start) {
		resource.end = buildDefaultGoogleEndFromStart_(resource.start, timeZone);
	}

	return resource;
}

/**
 * 指定した開始からデフォルトの終了を構築する（終日の場合は翌日、そうでなければ +1 時間）。
 * @param {Object} start 開始情報（date または dateTime を想定）
 * @param {string} timeZone タイムゾーン（デフォルト: SYNC_TIMEZONE）
 * @returns {Object} 終了情報オブジェクト（Google形式、ローカル時刻）
 */
function buildDefaultGoogleEndFromStart_(start, timeZone) {
	timeZone = timeZone || SYNC_TIMEZONE;

	if (!start) {
		return buildDefaultGoogleEndFromStart_(
			buildDefaultGoogleStart_(),
			timeZone,
		);
	}

	if (start.date) {
		const startDate = new Date(`${start.date}T00:00:00Z`);
		startDate.setUTCDate(startDate.getUTCDate() + 1);
		return {
			date: Utilities.formatDate(startDate, timeZone, 'yyyy-MM-dd'),
			timeZone: timeZone,
		};
	}

	if (start.dateTime) {
		// start.dateTime は内部形式（UTC）なので、UTC時刻を1時間後に計算
		const startDateTime = new Date(start.dateTime);
		const endDateTime = new Date(startDateTime.getTime() + 60 * 60 * 1000);
		// ローカルタイムに変換して返す
		return {
			dateTime: convertUtcToLocalDateTime_(
				Utilities.formatDate(endDateTime, 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'"),
				timeZone,
			),
			timeZone: timeZone,
		};
	}

	return {};
}

/**
 * 指定日時を同期タイムゾーンで Google API クエリ用（RFC3339 with offset）にフォーマットする。
 * @param {Date|string} value Date オブジェクトまたは日時文字列
 * @returns {string} RFC3339 形式の日時文字列（例: 2026-05-08T11:40:00+09:00）
 */
function buildGoogleApiQueryDateTime_(value) {
	const date = value instanceof Date ? value : new Date(value);
	return convertDateToGoogleDateTime_(date, SYNC_TIMEZONE);
}

/**
 * Date オブジェクトをGoogle形式（RFC3339 with offset）に変換する。
 * @param {Date} date Date オブジェクト
 * @param {string} timeZone タイムゾーン
 * @returns {string} RFC3339 形式の日時文字列（例: 2026-05-08T11:40:00+09:00）
 */
function convertDateToGoogleDateTime_(date, timeZone) {
	const dateTimeText = Utilities.formatDate(
		date,
		timeZone,
		"yyyy-MM-dd'T'HH:mm:ss",
	);

	// タイムゾーンオフセットを計算
	const offset = getTimezoneOffset_(date, timeZone);
	const offsetStr = formatOffset_(offset);

	return dateTimeText + offsetStr;
}

/**
 * タイムゾーンのUTCオフセット（分単位）を取得する。
 * @param {Date} date Date オブジェクト
 * @param {string} timeZone タイムゾーン（例: "Asia/Tokyo"）
 * @returns {number} UTCからのオフセット（分単位、例: 540 は +09:00）
 */
function getTimezoneOffset_(date, timeZone) {
	// 指定タイムゾーンでのフォーマットとUTC でのフォーマットの差分からオフセットを計算
	const tzFormatted = Utilities.formatDate(
		date,
		timeZone,
		'yyyy-MM-dd HH:mm:ss',
	);
	const utcFormatted = Utilities.formatDate(date, 'UTC', 'yyyy-MM-dd HH:mm:ss');

	const tzDate = new Date(tzFormatted);
	const utcDate = new Date(utcFormatted);

	const offsetMs = tzDate.getTime() - utcDate.getTime();
	return Math.round(offsetMs / 60000); // ミリ秒から分に変換
}

/**
 * オフセット（分単位）を "+HH:mm" または "-HH:mm" 形式の文字列に変換する。
 * @param {number} offsetMinutes オフセット（分単位、例: 540 は +09:00）
 * @returns {string} オフセット文字列（例: "+09:00", "-05:00"）
 */
function formatOffset_(offsetMinutes) {
	const sign = offsetMinutes >= 0 ? '+' : '-';
	const absOffset = Math.abs(offsetMinutes);
	const hours = Math.floor(absOffset / 60);
	const minutes = absOffset % 60;
	return (
		sign +
		String(hours).padStart(2, '0') +
		':' +
		String(minutes).padStart(2, '0')
	);
}

/**
 * RFC3339 with offset形式（"2026-05-08T11:40:00+09:00"）をUTC形式（"2026-05-08T02:40:00Z"）に変換する。
 * @param {string} rfc3339DateTime RFC3339形式の日時文字列
 * @returns {string} UTC形式の日時文字列
 */
function convertRfc3339ToUtc_(rfc3339DateTime) {
	if (!rfc3339DateTime) {
		return '';
	}

	// Date として解析（JavaScriptはRFC3339 with offsetを自動的にパース）
	const date = new Date(rfc3339DateTime);
	if (isNaN(date.getTime())) {
		return '';
	}

	// UTC形式でフォーマット
	return Utilities.formatDate(date, 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'");
}

/**
 * UTC形式の日時（"2026-05-08T02:40:00Z"）をローカル日時（オフセットなし）に変換する。
 * @param {string} utcDateTime UTC形式の日時文字列
 * @param {string} timeZone タイムゾーン
 * @returns {string} ローカル日時文字列（例: 2026-05-08T11:40:00）
 */
function convertUtcToLocalDateTime_(utcDateTime, timeZone) {
	if (!utcDateTime) {
		return '';
	}

	// UTC時刻を Date オブジェクトに変換
	const date = new Date(utcDateTime);
	if (isNaN(date.getTime())) {
		return '';
	}

	// 指定タイムゾーンでのローカル時刻にフォーマット
	return Utilities.formatDate(date, timeZone, "yyyy-MM-dd'T'HH:mm:ss");
}

/**
 * map helpers: Outlook <-> Google
 */
function mapTransparencyToShowAs(transparency) {
	// Google: 'transparent' (free) / 'opaque' (busy)
	if ((transparency || '').toLowerCase() === 'transparent') {
		return 'free';
	}
	return 'busy';
}

function mapShowAsToTransparency(showAs) {
	// Outlook showAs: 'free','busy','tentative','oof' => Google transparency
	if (!showAs) return undefined;
	const s = String(showAs).toLowerCase();
	if (s === 'free') return 'transparent';
	return 'opaque';
}

function mapVisibilityToSensitivity(visibility) {
	// Google visibility: 'default','public','private','confidential'
	// Outlook sensitivity: 'normal','personal','private','confidential'
	const v = (visibility || '').toLowerCase();
	if (v === 'private') return 'private';
	if (v === 'confidential') return 'confidential';
	if (v === 'public') return 'normal';
	return 'normal';
}

function mapSensitivityToVisibility(sensitivity) {
	const s = (sensitivity || '').toLowerCase();
	if (s === 'private') return 'private';
	if (s === 'confidential') return 'confidential';
	if (s === 'personal') return 'default';
	return 'default';
}

/**
 * Google イベントから recurrence 配列を抽出する。
 * @param {Object} googleEvent Google Calendar API のイベントオブジェクト
 * @returns {Array<string>|null} recurrence 配列（例: ["RRULE:FREQ=DAILY"]）またはnull
 */
function extractRecurrenceFromGoogleEvent(googleEvent) {
	if (!googleEvent || !googleEvent.recurrence) {
		return null;
	}

	// Google API の recurrence は配列で、RRULE形式の文字列を含む
	const recurrence = googleEvent.recurrence;
	if (Array.isArray(recurrence) && recurrence.length > 0) {
		return recurrence;
	}

	return null;
}

/**
 * Outlook 形式の recurrence オブジェクトから Google 形式の recurrence 配列を構築する。
 * @param {Object} outlookRecurrence Outlook Graph API の recurrence オブジェクト
 * @returns {Array<string>|null} Google 形式の recurrence 配列またはnull
 */
function buildGoogleRecurrenceFromOutlook(outlookRecurrence) {
	if (!outlookRecurrence || !outlookRecurrence.pattern) {
		return null;
	}

	const recurrenceRules = [];
	const pattern = outlookRecurrence.pattern;
	const range = outlookRecurrence.range || {};

	// pattern.type から FREQ を取得
	const freq = mapOutlookTypeToFreq_(pattern.type);
	if (!freq) {
		return null;
	}

	// RRULEを構築
	let rrule = `FREQ=${freq}`;

	// interval を追加
	if (pattern.interval && pattern.interval > 1) {
		rrule += `;INTERVAL=${pattern.interval}`;
	}

	// daysOfWeek を追加（BYDAY）
	if (pattern.daysOfWeek && Array.isArray(pattern.daysOfWeek)) {
		const bydays = pattern.daysOfWeek
			.map((day) => mapOutlookDayToRruleFormat_(day))
			.filter(Boolean)
			.join(',');
		if (bydays) {
			rrule += `;BYDAY=${bydays}`;
		}
	}

	// COUNT または UNTIL を追加
	if (range.type === 'numbered' && range.numberOfOccurrences) {
		rrule += `;COUNT=${range.numberOfOccurrences}`;
	} else if (range.type === 'endDate' && range.endDate) {
		// UNTIL は YYYYMMDD形式
		const endDateFormatted = range.endDate.replace(/-/g, '');
		rrule += `;UNTIL=${endDateFormatted}`;
	}

	recurrenceRules.push(`RRULE:${rrule}`);

	return recurrenceRules;
}

/**
 * Outlook の recurrenceType を RRULE の FREQ にマッピングする。
 * @param {string} outlookType Outlook のrecurrenceType（daily, weekly等）
 * @returns {string|null} RRULEのFREQ値
 */
function mapOutlookTypeToFreq_(outlookType) {
	const type = String(outlookType || '').toLowerCase();
	switch (type) {
		case 'daily':
			return 'DAILY';
		case 'weekly':
			return 'WEEKLY';
		case 'absolutemonthly':
		case 'relativeMonthly':
			return 'MONTHLY';
		case 'absoluteyearly':
		case 'relativeYearly':
			return 'YEARLY';
		default:
			return null;
	}
}

/**
 * Outlook の dayOfWeek 値を RRULE の BYDAY 形式にマッピングする。
 * @param {string} outlookDay Outlook の dayOfWeek 値（sunday, monday等）
 * @returns {string|null} RRULE形式の曜日コード（SU, MO等）
 */
function mapOutlookDayToRruleFormat_(outlookDay) {
	const day = String(outlookDay || '').toLowerCase();
	const mapping = {
		sunday: 'SU',
		monday: 'MO',
		tuesday: 'TU',
		wednesday: 'WE',
		thursday: 'TH',
		friday: 'FR',
		saturday: 'SA',
	};
	return mapping[day] || null;
}
