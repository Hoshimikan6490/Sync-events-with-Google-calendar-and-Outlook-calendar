const O365_GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

/**
 * [点検中] Outlook イベントを取得する。
 * @param {Date} startDate 取得開始日時
 * @param {Date} endDate 取得終了日時
 * @returns {Array<Object>} 正規化されたイベント配列（occurrence 単位）
 */
async function getOutlookEvents(startDate, endDate) {
	const path = `${getOutlookCalendarBasePath()}/calendarView`;
	const startDateTime = convertTimeFromInternalFormatToApiFormat(startDate);
	const endDateTime = convertTimeFromInternalFormatToApiFormat(endDate);
	const query = [
		`startDateTime=${encodeURIComponent(startDateTime)}`,
		`endDateTime=${encodeURIComponent(endDateTime)}`,
		'$orderby=start/dateTime',
	].join('&');

	const response = await fetchToOutlookAPI(
		`${O365_GRAPH_BASE}${path}?${query}`,
		{
			method: 'get',
			headers: {
				Authorization: `Bearer ${_getAccessToken()}`,
				Accept: 'application/json',
			},
			muteHttpExceptions: true,
		},
	);

	const payload = JSON.parse(response.getContentText() || '{}');
	const events = payload.value || [];
	return events
		.map(normalizeOutlookCalendarEvent_)
		.filter((event) => eventOverlapsWindow_(event, startDate, endDate));
}

/**
 * [点検済み] Outlook API のベースパス（カレンダー ID を含む）を返す。
 * @param void
 * @returns {string} ベースパス文字列
 */
function getOutlookCalendarBasePath() {
	const calendarId = _getScriptPropertyValue(OUTLOOK_PROPERTY_KEYS.calendarId);
	if (!calendarId) {
		return '/me';
	}

	return `/me/calendars/${encodeURIComponent(calendarId)}`;
}

/**
 * [点検済み] 内部表現の日時を Outlook API 用の日時文字列に変換する。
 * @param {Date|string} value Date オブジェクトまたは日時文字列
 * @returns {string} API 用の日時文字列
 */
function convertTimeFromInternalFormatToApiFormat(value) {
	const date = value instanceof Date ? value : new Date(value);
	const dateTimeText = date.toISOString().split('.')[0] + 'Z';
	return dateTimeText;
}

/**
 * [点検済み] Outlook Graph API を呼び出し、401 ならトークンリフレッシュして再試行する。
 * @param {string} url API エンドポイント URL
 * @param {Object} options UrlFetchApp に渡すオプション
 * @returns {HTTPResponse} UrlFetchApp のレスポンスオブジェクト
 */
function fetchToOutlookAPI(url, options) {
	const initialResponse = UrlFetchApp.fetch(url, options);
	if (initialResponse.getResponseCode() !== 401) {
		return initialResponse;
	}

	const refreshedToken = refreshAccessToken();
	options.headers.Authorization = `Bearer ${refreshedToken}`;

	return UrlFetchApp.fetch(url, options);
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
 * [未点検] Outlook にイベントを作成する (Graph API POST)。
 * @param {Object} eventData Outlook 用のイベントデータ
 * @returns {Object} 作成されたイベントのレスポンス（JSON）
 */
function createOutlookEvent(eventData) {
	const response = fetchToOutlookAPI(
		`${O365_GRAPH_BASE}${getOutlookCalendarBasePath()}/events`,
		{
			method: 'post',
			contentType: 'application/json',
			headers: {
				Authorization: `Bearer ${_getAccessToken()}`,
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
	const response = fetchToOutlookAPI(
		`${O365_GRAPH_BASE}${getOutlookCalendarBasePath()}/events/${encodeURIComponent(eventId)}`,
		{
			method: 'patch',
			contentType: 'application/json',
			headers: {
				Authorization: `Bearer ${_getAccessToken()}`,
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
	const response = fetchToOutlookAPI(
		`${O365_GRAPH_BASE}${getOutlookCalendarBasePath()}/events/${encodeURIComponent(eventId)}`,
		{
			method: 'delete',
			headers: {
				Authorization: `Bearer ${_getAccessToken()}`,
			},
			muteHttpExceptions: true,
		},
	);

	return response.getResponseCode();
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
