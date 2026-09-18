/**
 * [点検済み] 指定期間の Google カレンダーイベントを取得して正規化して返す。
 * occurrence 単位取得のため、singleEvents は true に設定。
 * @param {Date} startDate 取得開始日時
 * @param {Date} endDate 取得終了日時
 * @returns {Array<Object>} 正規化されたイベント配列（occurrence 単位）
 */
function _getGoogleEvents(startDate, endDate) {
	// カレンダーを取得
	const calendarId = CalendarApp.getDefaultCalendar().getId();

	// DateをGoogle API用のRFC3339形式に変換
	const timeMin = convertDateToGoogleDateTime(startDate);
	const timeMax = convertDateToGoogleDateTime(endDate);

	// Google Calendar APIを使用してイベントを取得
	const events =
		Calendar.Events.list(calendarId, {
			timeMin: timeMin,
			timeMax: timeMax,
			singleEvents: true,
			orderBy: 'startTime',
		}).items || [];

	const parentCache = new Map();

	return events
		.map((event) => normalizeGoogleCalendarEvent(event, parentCache))
		.filter((event) => googleEventOverlapsWindow(event, startDate, endDate));
}

/**
 * [点検済み] Google イベントが指定ウィンドウと重複するか判定する。
 * @param {Object} event Google イベント
 * @param {Date} startDate ウィンドウ開始
 * @param {Date} endDate ウィンドウ終了
 * @returns {boolean} 重複する場合は true
 */
function googleEventOverlapsWindow(event, startDate, endDate) {
	const start = event.start;
	const end = event.end;
	if (!start || !end) {
		return false;
	}

	const eventStart = new Date(start).getTime();
	const eventEnd = new Date(end).getTime();
	return eventEnd >= startDate.getTime() && eventStart <= endDate.getTime();
}

/**
 * [点検済み] Google Calendar API のイベントリソースを内部で扱う正規化形式（UTC + timeZone）に変換する。
 * Google のdateTime（RFC3339 with offset）をUTC + timeZoneに正規化する。
 * @param {Object} event API のイベントオブジェクト
 * @param {Map} parentCache 親イベントのキャッシュ
 * @returns {Object} 正規化されたイベントオブジェクト
 */
function normalizeGoogleCalendarEvent(event, parentCache) {
	const startNormalized = normalizeGoogleDateTime(event.start);
	const endNormalized = normalizeGoogleDateTime(event.end);

	// 内部表現は Outlook 仕様をメインとする（subject, showAs, sensitivity など）
	const isAllDay = Boolean(startNormalized.isAllDay && endNormalized.isAllDay);
	const normalizedEvent = {
		id: event.id,

		subject: event.summary || '',
		description: event.description || '',
		// location は Outlook 側で使う文字列形式を採用
		location:
			event.location || (event.location && event.location.displayName) || '',
		// start/end は既に正規化済みのオブジェクトをそのまま使う
		isAllDay: isAllDay,
		timezone: SYNC_TIMEZONE,
		start: startNormalized.dateTime,
		end: endNormalized.dateTime,
		// Google の空き状況および公開状況の設定を Outlook の showAs/sensitivity に変換
		showAs: mapTransparencyToShowAs(event.transparency),
		sensitivity: mapVisibilityToSensitivity(event.visibility),
		updatedAt: event.updated || null,

		// 繰り返しイベントの場合に、親イベントから変更された例外イベントを格納するフィールド
		exception: null,

		// イベントの生データも一応保持しておく
		raw: event,
	};

	// recurringEventIdが無い場合は、通常のイベントとして返す
	if (!event.recurringEventId) {
		return normalizedEvent;
	}

	// recurring eventの親を取得
	// 同じrecurringEventIdが複数来るので、キャッシュを使って親イベントを取得する
	const normalizedParent = getGoogleRecurringEventParent(
		event.recurringEventId,
		parentCache,
	);

	// 削除された繰り返しイベントのインスタンス（status=cancelled）を制御
	if (event.status === 'cancelled') {
		normalizedEvent.exception = createGoogleUpdatedOrDeletedException(
			event,
			'deleted',
		);
		return normalizedEvent;
	}

	// 親イベントと比較して、このインスタンスで変更された項目だけを取得する
	const diff = getGoogleEventDiff(normalizedParent, normalizedEvent);

	// 親と完全に同じなら通常の繰り返しイベントとして返し、exceptionは生成しない
	if (Object.keys(diff).length === 0) {
		return normalizedEvent;
	}

	// 親との差分がある場合には変更されたexceptionを生成して返す
	normalizedEvent.exception = createGoogleUpdatedOrDeletedException(
		event,
		'updated',
		{ diff },
	);

	return normalizedEvent;
}

/**
 * [点検済み] Google Calendar API の datetime オブジェクトを正規化する。
 * @param {Object} googleDateTime Google の start/end オブジェクト {dateTime, date, timeZone}
 * @returns {Object} 正規化されたオブジェクト {isAllDay:boolean, dateTime:string, timeZone:string}
 */
function normalizeGoogleDateTime(googleDateTime) {
	if (!googleDateTime) {
		throw new Error(
			"Failed to format Google's DateTime: No value was provided.",
		);
	}

	// 全日イベント: date のみ
	if (googleDateTime.date && !googleDateTime.dateTime) {
		return {
			isAllDay: true,
			dateTime: _convertTimeZone(googleDateTime, SYNC_TIMEZONE),
			timeZone: SYNC_TIMEZONE,
		};
	}

	// 時間指定イベント: dateTime が存在する場合
	if (googleDateTime.dateTime) {
		return {
			isAllDay: false,
			dateTime: _convertTimeZone(googleDateTime, SYNC_TIMEZONE),
			timeZone: googleDateTime.timeZone ?? SYNC_TIMEZONE,
		};
	}

	throw new Error(
		`Failed to format Google's DateTime: Unknown format was provided. Value: ${JSON.stringify(googleDateTime)}`,
	);
}

/**
 * [点検済み] Googleカレンダーの空き状況（transparency）をOutlookのshowAsにマッピングする。
 * @param {string} transparency Googleのtransparency値（'transparent' または 'opaque'）
 * @returns {string} OutlookのshowAs値（'free' または 'busy'）
 */
function mapTransparencyToShowAs(transparency) {
	// transparencyがtransparentの場合はfree、opaqueまたは未定義の場合はbusyにマッピングする
	try {
		if (transparency?.toLowerCase() === 'transparent') {
			return 'free';
		}
		return 'busy';
	} catch (error) {
		console.log(`given transparency: ${transparency}`);
		console.error('Error in mapTransparencyToShowAs:', error);
	}
}

/**
 * [点検済み] Googleカレンダーの可視性（visibility）をOutlookのsensitivityにマッピングする。
 * @param {string} visibility Googleのvisibility値
 * @returns {string} Outlookのsensitivity値
 */
function mapVisibilityToSensitivity(visibility) {
	// 明示的に指定されている場合はその値を返す
	// [重要] Googleのvisibilityは、public, private, defaultのいずれかであり、defaultはそのGoogleカレンダーの共有設定に従うという意味である。今回は、defaultをpublicとして扱う。
	if (visibility?.toLocaleLowerCase() === 'private') {
		return 'private';
	} else {
		return 'public';
	}
}

/**
 * [点検済み] recurringEventIdから親イベントを取得する。
 * @param {string} recurringEventId 親イベントのID
 * @param {Map} parentCache 親イベントのキャッシュ
 * @returns {Object|null} 親イベントオブジェクトまたはnull
 */
function getGoogleRecurringEventParent(recurringEventId, parentCache) {
	// 親が取得済みならAPIを呼び出さない
	if (parentCache.has(recurringEventId)) {
		return parentCache.get(recurringEventId);
	}

	// recurringEventIdは親イベント自身のIDなので、そっから取得
	const calendarId = CalendarApp.getDefaultCalendar().getId();
	const parentEvent = Calendar.Events.get(calendarId, recurringEventId);
	const startNormalized = normalizeGoogleDateTime(parentEvent.start);
	const endNormalized = normalizeGoogleDateTime(parentEvent.end);
	const parentNormalizedEvent = {
		id: parentEvent.id,
		subject: parentEvent.summary || '',
		description: parentEvent.description || '',
		location:
			parentEvent.location ||
			(parentEvent.location && parentEvent.location.displayName) ||
			'',
		isAllDay: Boolean(startNormalized.isAllDay && endNormalized.isAllDay),
		timezone: SYNC_TIMEZONE,
		start: startNormalized.dateTime,
		end: endNormalized.dateTime,
		showAs: mapTransparencyToShowAs(parentEvent.transparency),
		sensitivity: mapVisibilityToSensitivity(parentEvent.visibility),
		updatedAt: parentEvent.updated || null,
		exception: null,
		raw: parentEvent,
	};

	parentCache.set(recurringEventId, parentNormalizedEvent);

	return parentNormalizedEvent;
}

/**
 * [点検済み] Googleカレンダーの変更または削除された例外イベントを内部表現に変換する。
 * @param {Object} event Googleの削除された例外イベントオブジェクト
 * @param {string} type 例外イベントのタイプ（'updated' または 'deleted'）
 * @param {Object} diff 変更された項目の差分オブジェクト
 * @returns {Object} 内部表現の削除された例外イベントオブジェクト
 */
function createGoogleUpdatedOrDeletedException(event, type, { diff } = {}) {
	let diffEventInfo;
	if (type === 'updated') {
		diffEventInfo = diff;
	} else if (type === 'deleted') {
		diffEventInfo = null;
	}

	return {
		originalStart: {
			dateTime: _convertTimeZone(event.originalStartTime, SYNC_TIMEZONE),
			timeZone: SYNC_TIMEZONE,
		},
		status: type,
		id: event.id,
		updatedAt: event.updated || null,
		event: diffEventInfo,
	};
}

/**
 * [点検済み] Googleカレンダーの親イベントとインスタンスを比較し、変更された項目だけを返す。
 * @param {Object} parent Googleの親イベントオブジェクト
 * @param {Object} instance Googleのインスタンスイベントオブジェクト
 * @returns {Object} 変更された項目の差分オブジェクト
 */
function getGoogleEventDiff(parent, instance) {
	const diff = {};

	// Boolean型(isAllDay)の比較
	const parentIsAllDay = parent.isAllDay;
	const instanceIsAllDay = instance.isAllDay;

	if (parentIsAllDay !== instanceIsAllDay) {
		diff.isAllDay = instanceIsAllDay;
	}

	// 日付型(start,end)の比較
	const compareDateParameters = ['start', 'end'];
	compareDateParameters.forEach((param) => {
		const parentDate = parent[param];
		const instanceDate = instance[param];

		if (
			(parentIsAllDay !== instanceIsAllDay || parentDate.dateTime) !==
			instanceDate.dateTime
		) {
			diff[param] = instanceDate.dateTime;
			diff['timeZone'] = SYNC_TIMEZONE; // normalizeした時点で、timeZoneはSYNC_TIMEZONEに統一されているので、ここでもSYNC_TIMEZONEを返す
		}
	});

	// 文字列型(subject,description,location,showAs,sensitivity)の比較
	const compareStringParameters = [
		'subject',
		'description',
		'location',
		'showAs',
		'sensitivity',
		'updatedAt',
	];
	compareStringParameters.forEach((param) => {
		const parentValue = parent[param] ?? '';
		const instanceValue = instance[param] ?? '';

		if (parentValue !== instanceValue) {
			diff[param] = instanceValue;
		}
	});

	return diff;
}

/**
 * [未点検] Google カレンダーに新しいイベントを作成する。
 * @param {Object} eventData 作成するイベントデータ（summary, start, end 等）
 * @returns {Object} 作成されたイベントのリソース
 */
function createGoogleEvent(eventData) {
	const calendarId = CalendarApp.getDefaultCalendar().getId();
	const resource = buildGoogleCalendarResource_(eventData);
	return Calendar.Events.insert(resource, calendarId);
}

/**
 * [未点検]既存の Google イベントを更新する。
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
 * [未点検] Google カレンダーからイベントを削除する。
 * @param {string} eventId 削除対象のイベント ID
 * @returns void
 */
function deleteGoogleEvent(eventId) {
	const calendarId = CalendarApp.getDefaultCalendar().getId();
	return Calendar.Events.remove(calendarId, eventId);
}

/**
 * [未点検] Google イベントの説明文を組み立てる（outlookSyncKey を含める）。
 * @param {Object} event 元イベントオブジェクト（description を使用）
 * @param {string} outlookSyncKey Outlook のイベント同期キー
 * @returns {string} 組み立てた説明文
 */
function buildGoogleDescription(event, outlookSyncKey) {
	const lines = [];
	const descriptionLines = String(event.description || '')
		.split(/\r?\n/)
		.filter((line) => line && !/^outlookSyncKey:/i.test(line));

	if (descriptionLines.length > 0) {
		lines.push(descriptionLines.join('\n').trim());
	}
	if (outlookSyncKey) {
		lines.push(`outlookSyncKey:${outlookSyncKey}`);
	}
	return lines.join('\n');
}

/**
 * [未点検] Google Calendar API に渡すリソースオブジェクトを構築する。
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
 * [未点検] 指定した開始からデフォルトの終了を構築する（終日の場合は翌日、そうでなければ +1 時間）。
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
 * [未点検] Date をGoogle形式（RFC3339 with offset）に変換する。
 * @param {Date|string} date Date オブジェクトまたは日時文字列
 * @returns {string} RFC3339 形式の日時文字列（例: 2026-05-08T11:40:00+09:00）
 */
function convertDateToGoogleDateTime(value) {
	const date = value instanceof Date ? value : new Date(value);

	const dateTimeText = Utilities.formatDate(
		date,
		SYNC_TIMEZONE,
		"yyyy-MM-dd'T'HH:mm:ss",
	);

	// タイムゾーンオフセットを計算
	const offset = getTimezoneOffset_(date, SYNC_TIMEZONE);
	const offsetStr = formatOffset_(offset);

	return dateTimeText + offsetStr;
}

/**
 * [未点検] タイムゾーンのUTCオフセット（分単位）を取得する。
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
 * [未点検] オフセット（分単位）を "+HH:mm" または "-HH:mm" 形式の文字列に変換する。
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
 * [未点検] UTC形式の日時（"2026-05-08T02:40:00Z"）をローカル日時（オフセットなし）に変換する。
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
 * [未点検] OutlookのshowAsをGoogleのtransparencyにマッピングする。
 * @param {string} showAs OutlookのshowAs値（'free' または 'busy'）
 * @returns {string} Googleのtransparency値（'transparent' または 'opaque'）
 */
function mapShowAsToTransparency(showAs) {
	// Outlook showAs: 'free','busy','tentative','oof' => Google transparency
	if (!showAs) return undefined;
	const s = String(showAs).toLowerCase();
	if (s === 'free') return 'transparent';
	return 'opaque';
}

/**
 * [未点検] OutlookのsensitivityをGoogleのvisibilityにマッピングする。
 * @param {string} sensitivity Outlookのsensitivity値
 * @returns {string} Googleのvisibility値
 */
function mapSensitivityToVisibility(sensitivity) {
	const s = (sensitivity || '').toLowerCase();
	if (s === 'private') return 'private';
	if (s === 'confidential') return 'confidential';
	if (s === 'personal') return 'default';
	return 'default';
}

/**
 * [未点検] Google イベントから recurrence 配列を抽出する。
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
