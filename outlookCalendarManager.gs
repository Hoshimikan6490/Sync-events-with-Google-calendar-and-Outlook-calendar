/**
 * [点検中] Outlook イベントを取得する。
 * @param {Date} startDate 取得開始日時
 * @param {Date} endDate 取得終了日時
 * @returns {Array<Object>} 正規化されたイベント配列（occurrence 単位）
 */
async function getOutlookEvents(startDate, endDate) {
	const path = getOutlookCalendarBasePath_() + '/calendarView';
	const startDateTime = buildApiDateTimeInSyncTimezone_(startDate);
	const endDateTime = buildApiDateTimeInSyncTimezone_(endDate);
	const query = [
		`startDateTime=${encodeURIComponent(startDateTime)}`,
		`endDateTime=${encodeURIComponent(endDateTime)}`,
		'$orderby=start/dateTime',
	].join('&');

	const response = fetchOutlookGraph_(`${O365_GRAPH_BASE}${path}?${query}`, {
		method: 'get',
		headers: {
			Authorization: `Bearer ${getAccessToken()}`,
			Accept: 'application/json',
		},
		muteHttpExceptions: true,
	});

	const payload = JSON.parse(response.getContentText() || '{}');
	const events = payload.value || [];
	return events
		.map(normalizeOutlookCalendarEvent_)
		.filter((event) => eventOverlapsWindow_(event, startDate, endDate));
}

/**
 * [未点検] マスターイベントから単一の発生日の発生イベントを構築する。
 * @param {Object} master シリーズのマスターイベント
 * @param {Date} startDate 発生日の開始日時
 * @param {number} durationMs 期間（ミリ秒）
 * @param {boolean} allDay 終日フラグ
 * @returns {Object} 正規化された発生日イベント
 */
function buildRecurringOccurrence_(master, startDate, durationMs, allDay) {
	const normalized = normalizeOutlookEvent_(master);
	if (allDay) {
		const startDateOnly = Utilities.formatDate(
			startDate,
			SYNC_TIMEZONE,
			'yyyy-MM-dd',
		);
		const endDateOnly = Utilities.formatDate(
			new Date(startDate.getTime() + durationMs),
			SYNC_TIMEZONE,
			'yyyy-MM-dd',
		);
		normalized.start = { date: startDateOnly, timeZone: SYNC_TIMEZONE };
		normalized.end = { date: endDateOnly, timeZone: SYNC_TIMEZONE };
		normalized.occurrenceDate = startDateOnly;
		return normalized;
	}

	const startString = Utilities.formatDate(
		startDate,
		SYNC_TIMEZONE,
		"yyyy-MM-dd'T'HH:mm:ss",
	);
	const endString = Utilities.formatDate(
		new Date(startDate.getTime() + durationMs),
		SYNC_TIMEZONE,
		"yyyy-MM-dd'T'HH:mm:ss",
	);
	normalized.start = { dateTime: startString, timeZone: SYNC_TIMEZONE };
	normalized.end = { dateTime: endString, timeZone: SYNC_TIMEZONE };
	normalized.occurrenceDate = normalizeOccurrenceDateText_(startString);
	return normalized;
}

/**
 * [未点検] 発生日の開始日時を生成する（終日の場合は時間を除去）。
 * @param {Date} candidateDate 発生日候補の Date
 * @param {Date} masterStartDate マスターの開始日時
 * @param {boolean} allDay 終日フラグ
 * @returns {Date} 発生日の開始日時
 */
function buildOccurrenceStart_(candidateDate, masterStartDate, allDay) {
	if (allDay) {
		return stripTime_(candidateDate);
	}

	const occurrence = new Date(candidateDate.getTime());
	occurrence.setHours(
		masterStartDate.getHours(),
		masterStartDate.getMinutes(),
		masterStartDate.getSeconds(),
		masterStartDate.getMilliseconds(),
	);
	return occurrence;
}

/**
 * [未点検] イベントオブジェクトから発生日キーを生成する。
 * @param {Object} event イベントオブジェクト
 * @returns {string} 発生日キー
 */
function buildOccurrenceKeyFromEvent_(event) {
	if (event.start && event.start.date) {
		return `date:${event.start.date}`;
	}
	const dateTime =
		event.start && event.start.dateTime ? event.start.dateTime : '';
	return `dateTime:${dateTime}`;
}

/**
 * [未点検] Date から発生日キーを生成するユーティリティ。
 * @param {Date} date 発生日の Date
 * @param {boolean} allDay 終日フラグ
 * @returns {string} 発生日キー
 */
function buildOccurrenceKey_(date, allDay) {
	if (!date) {
		return '';
	}
	return allDay
		? `date:${Utilities.formatDate(date, SYNC_TIMEZONE, 'yyyy-MM-dd')}`
		: `dateTime:${Utilities.formatDate(date, SYNC_TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss")}`;
}

/**
 * [未点検] 発生日キーで比較してソート用の比較値を返す。
 * @param {Object} left 左側イベント
 * @param {Object} right 右側イベント
 * @returns {number} 比較結果 -1/0/1
 */
function compareOccurrenceStart_(left, right) {
	const leftKey = buildOccurrenceKeyFromEvent_(left);
	const rightKey = buildOccurrenceKeyFromEvent_(right);
	if (leftKey < rightKey) {
		return -1;
	}
	if (leftKey > rightKey) {
		return 1;
	}
	return 0;
}

/**
 * [未点検] イベントが指定ウィンドウと重複するか判定する。
 * @param {Object} event イベント（start, end を含む）
 * @param {Date} startDate ウィンドウ開始
 * @param {Date} endDate ウィンドウ終了
 * @returns {boolean} 重複する場合は true
 */
function eventOverlapsWindow_(event, startDate, endDate) {
	const start = event.start && (event.start.dateTime || event.start.date);
	const end = event.end && (event.end.dateTime || event.end.date);
	if (!start || !end) {
		return false;
	}
	const startKey = new Date(start).getTime();
	const endKey = new Date(end).getTime();
	return endKey >= startDate.getTime() && startKey <= endDate.getTime();
}

/**
 * [未点検] 指定日付に日数を加算した新しい Date を返す。
 * @param {Date} date ベース日付
 * @param {number} days 加算する日数（負数可）
 * @returns {Date} 計算結果の Date
 */
function addDays_(date, days) {
	const next = new Date(date.getTime());
	next.setDate(next.getDate() + days);
	return next;
}

/**
 * [未点検] Date から時刻情報を除去して日付のみの Date を返す。
 * @param {Date} date 入力の Date
 * @returns {Date} 時刻を除去した Date
 */
function stripTime_(date) {
	return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

/**
 * [未点検] Outlook Graph API を呼び出し、401 ならトークンリフレッシュして再試行する。
 * @param {string} url API エンドポイント URL
 * @param {Object} options UrlFetchApp に渡すオプション
 * @returns {HTTPResponse} UrlFetchApp のレスポンスオブジェクト
 */
function fetchOutlookGraph_(url, options) {
	const initialResponse = UrlFetchApp.fetch(url, options);
	if (initialResponse.getResponseCode() !== 401) {
		return initialResponse;
	}

	const refreshedToken = refreshAccessToken();
	const retryOptions = Object.assign({}, options, {
		headers: Object.assign({}, options.headers, {
			Authorization: `Bearer ${refreshedToken}`,
		}),
	});
	return UrlFetchApp.fetch(url, retryOptions);
}

/**
 * [未点検] Outlook にイベントを作成する (Graph API POST)。
 * @param {Object} eventData Outlook 用のイベントデータ
 * @returns {Object} 作成されたイベントのレスポンス（JSON）
 */
function createOutlookEvent(eventData) {
	const response = fetchOutlookGraph_(
		`${O365_GRAPH_BASE}${getOutlookCalendarBasePath_()}/events`,
		{
			method: 'post',
			contentType: 'application/json',
			headers: {
				Authorization: `Bearer ${getAccessToken()}`,
				Accept: 'application/json',
			},
			payload: JSON.stringify(buildOutlookCalendarResource_(eventData)),
			muteHttpExceptions: true,
		},
	);

	return JSON.parse(response.getContentText() || '{}');
}

/**
 * [未点検] Outlook のイベントを更新する (Graph API PATCH)。
 * @param {string} eventId 更新対象のイベント ID
 * @param {Object} eventData 更新データ
 * @returns {string|HTTPResponse} API の応答ボディまたはレスポンス
 */
function updateOutlookEvent(eventId, eventData) {
	const response = fetchOutlookGraph_(
		`${O365_GRAPH_BASE}${getOutlookCalendarBasePath_()}/events/${encodeURIComponent(eventId)}`,
		{
			method: 'patch',
			contentType: 'application/json',
			headers: {
				Authorization: `Bearer ${getAccessToken()}`,
				Accept: 'application/json',
			},
			payload: JSON.stringify(buildOutlookCalendarResource_(eventData)),
			muteHttpExceptions: true,
		},
	);

	return response.getContentText();
}

/**
 * [未点検] Outlook のイベントを削除する (Graph API DELETE)。
 * @param {string} eventId 削除対象のイベント ID
 * @returns {number} HTTP ステータスコード
 */
function deleteOutlookEvent(eventId) {
	const response = fetchOutlookGraph_(
		`${O365_GRAPH_BASE}${getOutlookCalendarBasePath_()}/events/${encodeURIComponent(eventId)}`,
		{
			method: 'delete',
			headers: {
				Authorization: `Bearer ${getAccessToken()}`,
			},
			muteHttpExceptions: true,
		},
	);

	return response.getResponseCode();
}

/**
 * [未点検] Outlook API のベースパス（カレンダー ID を含む）を返す。
 * @param void
 * @returns {string} ベースパス文字列
 */
function getOutlookCalendarBasePath_() {
	const calendarId = getScriptPropertyValue(OUTLOOK_PROPERTY_KEYS.calendarId);
	if (!calendarId) {
		return '/me';
	}

	return `/me/calendars/${encodeURIComponent(calendarId)}`;
}

/**
 * [未点検] スクリプトプロパティからアクセストークンを取得し、無ければリフレッシュを試みる。
 * @param void
 * @returns {string} 利用可能なアクセストークン
 */
function getAccessToken() {
	const accessToken = getScriptPropertyValue(OUTLOOK_PROPERTY_KEYS.accessToken);
	if (accessToken) {
		return accessToken;
	}

	const refreshToken = getScriptPropertyValue(
		OUTLOOK_PROPERTY_KEYS.refreshToken,
	);
	if (!refreshToken) {
		throw new Error(
			'ACCESS_TOKEN または REFRESH_TOKEN が設定されていません。outlookOauth2.gs の setup() と authCallback() を実行してください。',
		);
	}

	return refreshAccessToken();
}

/**
 * [未点検] Outlook 用の説明文を組み立てる（`googleSyncKey` を付加）。
 * @param {Object} event 元イベントオブジェクト（description を使用）
 * @param {string} googleSyncKey Google のイベント同期キー
 * @returns {string} 組み立てた説明文
 */
function buildOutlookDescription(event, googleSyncKey) {
	const lines = [];
	const descriptionLines = String(event.description || '')
		.split(/\r?\n/)
		.filter((line) => line && !/^googleSyncKey:/i.test(line));

	if (descriptionLines.length > 0) {
		lines.push(descriptionLines.join('\n').trim());
	}
	if (googleSyncKey) {
		lines.push(`googleSyncKey:${googleSyncKey}`);
	}
	return lines.join('\n');
}

/**
 * [未点検] Graph API のイベントを内部で扱う正規化形式に変換する。
 * @param {Object} event Graph API のイベントオブジェクト
 * @returns {Object} 正規化されたイベントオブジェクト
 */
function normalizeOutlookCalendarEvent_(event) {
	const startDateTime =
		event.start && (event.start.dateTime || event.start.date);
	return {
		id: event.id,
		subject: event.subject || '',
		description: event.body && event.body.content ? event.body.content : '',
		location:
			(event.location && event.location.displayName) || event.location || '',
		start: event.start || {},
		end: event.end || {},
		isAllDay: Boolean(event.isAllDay),
		showAs: event.showAs || 'busy',
		sensitivity: event.sensitivity || 'normal',
		recurrence: event.recurrence || null,
		// occurrence 識別用フィールド
		uid: event.id,
		occurrenceDate: normalizeOccurrenceDateText_(startDateTime || ''),
		raw: event,
	};
}

/**
 * [未点検] Outlook Graph API に渡すイベントリソースを構築する。
 * @param {Object} eventData 内部表現のイベントデータ
 * @returns {Object} Graph API 用のイベントオブジェクト
 */
function buildOutlookCalendarResource_(eventData) {
	return {
		subject: eventData.subject || '',
		body: eventData.body || {
			contentType: 'text',
			content: eventData.description || '',
		},
		// Graph API expects location object
		location: eventData.location
			? typeof eventData.location === 'string'
				? { displayName: eventData.location }
				: { displayName: eventData.location.displayName || '' }
			: undefined,
		start: eventData.start || {
			dateTime: eventData.startDateTime || '',
			timeZone: SYNC_TIMEZONE,
		},
		end: eventData.end || {
			dateTime: eventData.endDateTime || '',
			timeZone: SYNC_TIMEZONE,
		},
		isAllDay: Boolean(eventData.isAllDay),
		showAs:
			eventData.showAs ||
			mapTransparencyToShowAs(eventData.transparency) || //※後ほどTODO: これはGoogleのtransparencyを変換する関数なので、内容の確認が必要
			'busy',
		sensitivity:
			eventData.sensitivity ||
			mapVisibilityToSensitivity(eventData.visibility) || //※後ほどTODO: これはGoogleのvisibilityを変換する関数なので、内容の確認が必要
			'normal',
	};
}

/**
 * [未点検] 指定日時を同期タイムゾーンで API 用日時（ISO 形式、オフセットなし）にフォーマットする。
 * @param {Date|string} value Date オブジェクトまたは日時文字列
 * @returns {string} API 用の日時文字列
 */
function buildApiDateTimeInSyncTimezone_(value) {
	const date = value instanceof Date ? value : new Date(value);
	const dateTimeText = Utilities.formatDate(
		date,
		SYNC_TIMEZONE,
		"yyyy-MM-dd'T'HH:mm:ss",
	);
	return dateTimeText;
}

/**
 * [未点検] FREQ値をOutlook形式の recurrenceType にマッピングする。
 * @param {string} freq FREQ値（DAILY, WEEKLY, MONTHLY, YEARLY等）
 * @returns {string} Outlook の recurrenceType（daily, weekly, absoluteMonthly等）
 */
function mapFreqToOutlookRecurrenceType_(freq) {
	const f = String(freq || '').toUpperCase();
	switch (f) {
		case 'DAILY':
			return 'daily';
		case 'WEEKLY':
			return 'weekly';
		case 'MONTHLY':
			return 'absoluteMonthly';
		case 'YEARLY':
			return 'absoluteYearly';
		default:
			return 'daily';
	}
}

/**
 * [未点検] 曜日番号（0=Sun...6=Sat）を Outlook の dayOfWeek 値にマッピングする。
 * @param {number} weekday 曜日番号
 * @returns {string} Outlook の dayOfWeek 値（sunday, monday等）
 */
function mapWeekdayToOutlookDay_(weekday) {
	const days = [
		'sunday',
		'monday',
		'tuesday',
		'wednesday',
		'thursday',
		'friday',
		'saturday',
	];
	return days[weekday] || 'monday';
}
