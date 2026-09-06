/**
 * スクリプトプロパティに設定された ICS URL を取得して内容を返す。
 * @param void
 * @returns {string} ICS テキスト（取得失敗時は空文字）
 */
function fetchIcs() {
	const properties = PropertiesService.getScriptProperties();
	const icsUrl = properties.getProperty('OUTLOOK_ICS_URL');

	if (!icsUrl) {
		return '';
	}

	const response = UrlFetchApp.fetch(icsUrl, {
		method: 'get',
		muteHttpExceptions: true,
	});

	return response.getContentText();
}

/**
 * ICS の折り返し行を復元して 1 行ずつにするユーティリティ。
 * @param {Array<string>} lines ICS の行配列
 * @returns {Array<string>} 折り返しを展開した行配列
 */
function unfoldIcsLines_(lines) {
	const unfolded = [];
	for (const line of lines) {
		if (
			(line.startsWith(' ') || line.startsWith('\t')) &&
			unfolded.length > 0
		) {
			unfolded[unfolded.length - 1] += line.slice(1);
			continue;
		}
		unfolded.push(line);
	}
	return unfolded;
}

/**
 * 受け取った VEVENT の配列から繰り返し展開を行いフラットなイベント配列を返す。
 * @param {Array<Object>} events ICS から抽出したイベントオブジェクト配列
 * @param {Date} startDate 展開ウィンドウ開始
 * @param {Date} endDate 展開ウィンドウ終了
 * @returns {Array<Object>} 展開されたイベント配列
 */
function expandRecurringEvents(events, startDate, endDate) {
	const grouped = new Map();
	for (const event of events || []) {
		const uid = event.uid || event.id || Utilities.getUuid();
		if (!grouped.has(uid)) {
			grouped.set(uid, []);
		}
		grouped.get(uid).push(event);
	}

	const expanded = [];
	for (const seriesEvents of grouped.values()) {
		const seriesExpanded = expandRecurringSeries_(
			seriesEvents,
			startDate,
			endDate,
		);
		expanded.push.apply(expanded, seriesExpanded);
	}

	return expanded;
}

/**
 * 単一のシリーズ（同一 UID の複数 VEVENT）を繰り返しルールに基づき展開する。
 * @param {Array<Object>} seriesEvents シリーズに属する VEVENT 配列
 * @param {Date} startDate 展開ウィンドウ開始
 * @param {Date} endDate 展開ウィンドウ終了
 * @returns {Array<Object>} 展開されたイベント配列
 */
function expandRecurringSeries_(seriesEvents, startDate, endDate) {
	const master =
		seriesEvents.find((event) => !getIcsFieldValue_(event, 'recurrence-id')) ||
		seriesEvents[0];
	const masterStartInfo = getIcsDateTimeInfo_(master, 'dtstart');
	const masterEndInfo = getIcsDateTimeInfo_(master, 'dtend');
	const masterStartDate = toDateFromIcsValue_(
		masterStartInfo.raw,
		masterStartInfo.allDay,
	);
	const masterEndDate = toDateFromIcsValue_(
		masterEndInfo.raw,
		masterEndInfo.allDay,
	);
	const durationMs =
		masterStartDate && masterEndDate
			? masterEndDate.getTime() - masterStartDate.getTime()
			: masterStartInfo.allDay
				? 24 * 60 * 60 * 1000
				: 60 * 60 * 1000;
	const recurrenceRule = parseRrule_(getIcsFieldValue_(master, 'rrule'));
	const exdateSet = buildExdateSet_(seriesEvents);
	const overrideMap = buildOverrideMap_(seriesEvents);
	const windowStart = startDate || new Date(0);
	const windowEnd = endDate || new Date(Date.now() + 31 * 24 * 60 * 60 * 1000);

	if (!recurrenceRule.freq) {
		const single = normalizeOutlookEvent_(master);
		return eventOverlapsWindow_(single, windowStart, windowEnd) ? [single] : [];
	}

	const occurrences = [];
	let repeat = 0;
	for (
		let cursor = new Date(windowStart.getTime());
		cursor <= windowEnd;
		cursor = addDays_(cursor, 1)
	) {
		if (!masterStartDate || cursor < masterStartDate) {
			continue;
		}
		if (!matchesRrule_(cursor, masterStartDate, recurrenceRule)) {
			continue;
		}

		const occurrenceStart = buildOccurrenceStart_(
			cursor,
			masterStartDate,
			masterStartInfo.allDay,
		);
		const occurrenceKey = buildOccurrenceKey_(
			occurrenceStart,
			masterStartInfo.allDay,
		);
		if (exdateSet.has(occurrenceKey)) {
			continue;
		}

		const override = overrideMap.get(occurrenceKey);
		if (override) {
			occurrences.push(normalizeOutlookEvent_(override));
			repeat += 1;
			continue;
		}

		occurrences.push(
			buildRecurringOccurrence_(
				master,
				occurrenceStart,
				durationMs,
				masterStartInfo.allDay,
			),
		);
		repeat += 1;

		if (recurrenceRule.count && repeat >= recurrenceRule.count) {
			break;
		}
	}

	for (const override of overrideMap.values()) {
		const occurrence = normalizeOutlookEvent_(override);
		if (eventOverlapsWindow_(occurrence, windowStart, windowEnd)) {
			const occurrenceKey = buildOccurrenceKeyFromEvent_(occurrence);
			if (
				!occurrences.some(
					(item) => buildOccurrenceKeyFromEvent_(item) === occurrenceKey,
				)
			) {
				occurrences.push(occurrence);
			}
		}
	}

	occurrences.sort((left, right) => compareOccurrenceStart_(left, right));
	return occurrences;
}

/**
 * マスターイベントから単一の発生日の発生イベントを構築する。
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
 * 発生日の開始日時を生成する（終日の場合は時間を除去）。
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
 * 指定日が RRULE にマッチするか判定する。
 * @param {Date} candidateDate 判定対象の日付
 * @param {Date} masterStartDate マスター開始日
 * @param {Object} rule 解析済みの RRULE オブジェクト
 * @returns {boolean} マッチする場合は true
 */
function matchesRrule_(candidateDate, masterStartDate, rule) {
	const freq = (rule.freq || '').toUpperCase();
	const interval = Math.max(Number(rule.interval || 1), 1);
	const candidateStart = stripTime_(candidateDate);
	const masterStart = stripTime_(masterStartDate);

	if (candidateDate < masterStartDate) {
		return false;
	}

	if (rule.until && candidateDate > rule.until) {
		return false;
	}

	if (freq === 'DAILY') {
		const days = Math.floor(
			(candidateStart.getTime() - masterStart.getTime()) / 86400000,
		);
		return days % interval === 0 && matchesByDay_(candidateDate, rule.byday);
	}

	if (freq === 'WEEKLY') {
		const weeks = Math.floor(
			(candidateStart.getTime() - masterStart.getTime()) / (7 * 86400000),
		);
		return (
			weeks % interval === 0 &&
			matchesByDay_(candidateDate, rule.byday, masterStartDate.getDay())
		);
	}

	if (freq === 'MONTHLY') {
		const months =
			(candidateDate.getFullYear() - masterStartDate.getFullYear()) * 12 +
			(candidateDate.getMonth() - masterStartDate.getMonth());
		if (months % interval !== 0) {
			return false;
		}
		if (rule.byday.length > 0) {
			return matchesByDay_(candidateDate, rule.byday);
		}
		return candidateDate.getDate() === masterStartDate.getDate();
	}

	if (freq === 'YEARLY') {
		const years = candidateDate.getFullYear() - masterStartDate.getFullYear();
		if (years % interval !== 0) {
			return false;
		}
		if (rule.byday.length > 0) {
			return matchesByDay_(candidateDate, rule.byday);
		}
		return (
			candidateDate.getMonth() === masterStartDate.getMonth() &&
			candidateDate.getDate() === masterStartDate.getDate()
		);
	}

	return false;
}

/**
 * BYDAY 条件にマッチするかを判定するヘルパー。
 * @param {Date} candidateDate 判定対象の日付
 * @param {Array<Object>} bydayTokens BYDAY トークン配列
 * @param {number} [fallbackDay] フォールバックの曜日番号
 * @returns {boolean} マッチする場合は true
 */
function matchesByDay_(candidateDate, bydayTokens, fallbackDay) {
	if (!bydayTokens || bydayTokens.length === 0) {
		if (fallbackDay === undefined) {
			return true;
		}
		return candidateDate.getDay() === fallbackDay;
	}

	return bydayTokens.some((token) => token.weekday === candidateDate.getDay());
}

/**
 * RRULE テキストを解析してオブジェクトに変換する。
 * @param {string} rruleText RRULE の生テキスト
 * @returns {Object} 解析結果のルールオブジェクト
 */
function parseRrule_(rruleText) {
	const rule = {
		freq: '',
		interval: 1,
		count: 0,
		until: null,
		byday: [],
	};

	for (const part of String(rruleText || '').split(';')) {
		if (!part) {
			continue;
		}
		const equalsIndex = part.indexOf('=');
		if (equalsIndex === -1) {
			continue;
		}
		const key = part.slice(0, equalsIndex).toUpperCase();
		const value = part.slice(equalsIndex + 1);
		if (key === 'FREQ') {
			rule.freq = value;
			continue;
		}
		if (key === 'INTERVAL') {
			rule.interval = Math.max(Number(value) || 1, 1);
			continue;
		}
		if (key === 'COUNT') {
			rule.count = Math.max(Number(value) || 0, 0);
			continue;
		}
		if (key === 'UNTIL') {
			rule.until =
				toDateFromIcsValue_(value, false) || toDateFromIcsValue_(value, true);
			continue;
		}
		if (key === 'BYDAY') {
			rule.byday = parseByDayList_(value);
		}
	}

	return rule;
}

/**
 * BYDAY リスト文字列を解析してトークン配列を返す。
 * @param {string} value BYDAY のカンマ区切り文字列
 * @returns {Array<Object>} {ordinal:number,weekday:number} の配列
 */
function parseByDayList_(value) {
	return String(value || '')
		.split(',')
		.map((token) => token.trim())
		.filter(Boolean)
		.map((token) => {
			const weekdayToken = token.slice(-2).toUpperCase();
			return {
				ordinal: token.length > 2 ? Number(token.slice(0, -2)) || 0 : 0,
				weekday: mapWeekdayToken_(weekdayToken),
			};
		})
		.filter((token) => token.weekday >= 0);
}

/**
 * BYDAY のトークンを曜日番号にマッピングする。
 * @param {string} token 曜日トークン（SU,MO,..）
 * @returns {number} 曜日番号（0=Sun ... 6=Sat）、不明なら -1
 */
function mapWeekdayToken_(token) {
	switch (token) {
		case 'SU':
			return 0;
		case 'MO':
			return 1;
		case 'TU':
			return 2;
		case 'WE':
			return 3;
		case 'TH':
			return 4;
		case 'FR':
			return 5;
		case 'SA':
			return 6;
		default:
			return -1;
	}
}

/**
 * シリーズ中の EXDATE を解析して Set を構築する。
 * @param {Array<Object>} seriesEvents シリーズに属するイベント配列
 * @returns {Set<string>} 発生日キーの集合
 */
function buildExdateSet_(seriesEvents) {
	const set = new Set();
	for (const event of seriesEvents) {
		const exdates = getIcsFieldValue_(event, 'exdate');
		if (!exdates) {
			continue;
		}
		for (const value of exdates.split(',')) {
			const date =
				toDateFromIcsValue_(value.trim(), false) ||
				toDateFromIcsValue_(value.trim(), true);
			if (date) {
				set.add(
					buildOccurrenceKey_(
						date,
						/;value=date/i.test(findIcsFieldKey_(event, 'exdate')),
					),
				);
			}
		}
	}
	return set;
}

/**
 * シリーズ中のオーバーライド（RECURRENCE-ID を持つイベント）をマップ化する。
 * @param {Array<Object>} seriesEvents シリーズに属するイベント配列
 * @returns {Map<string,Object>} 発生日キーをキーとするオーバーライドマップ
 */
function buildOverrideMap_(seriesEvents) {
	const map = new Map();
	for (const event of seriesEvents) {
		const recurrenceId = getIcsFieldValue_(event, 'recurrence-id');
		if (!recurrenceId) {
			continue;
		}
		const key = buildOccurrenceKeyFromIcsValue_(
			recurrenceId,
			/;value=date/i.test(findIcsFieldKey_(event, 'recurrence-id')),
		);
		map.set(key, event);
	}
	return map;
}

/**
 * ICS の日時値から発生日キーを生成する。
 * @param {string} value ICS の日時値
 * @param {boolean} allDay 終日フラグ
 * @returns {string} 発生日キー
 */
function buildOccurrenceKeyFromIcsValue_(value, allDay) {
	const date = toDateFromIcsValue_(value, allDay);
	return buildOccurrenceKey_(date, allDay);
}

/**
 * イベントオブジェクトから発生日キーを生成する。
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
 * Date から発生日キーを生成するユーティリティ。
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
 * 発生日キーで比較してソート用の比較値を返す。
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
 * イベントが指定ウィンドウと重複するか判定する。
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
 * 指定日付に日数を加算した新しい Date を返す。
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
 * Date から時刻情報を除去して日付のみの Date を返す。
 * @param {Date} date 入力の Date
 * @returns {Date} 時刻を除去した Date
 */
function stripTime_(date) {
	return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

/**
 * ICS イベントオブジェクトから指定フィールドの値を取得する。
 * @param {Object} event ICS のイベントオブジェクト
 * @param {string} fieldName フィールド名（例: 'exdate'）
 * @returns {string} フィールド値（存在しなければ空文字）
 */
function getIcsFieldValue_(event, fieldName) {
	const key = findIcsFieldKey_(event, fieldName);
	return key ? event[key] || '' : '';
}

/**
 * ICS イベントオブジェクトのキー一覧から指定フィールドに該当する実キーを見つける。
 * @param {Object} event ICS のイベントオブジェクト
 * @param {string} fieldName 欲しいフィールド名
 * @returns {string} 実際のキー名（無ければ空文字）
 */
function findIcsFieldKey_(event, fieldName) {
	const lowerField = String(fieldName || '').toLowerCase();
	for (const key of Object.keys(event || {})) {
		if (key === lowerField || key.startsWith(`${lowerField};`)) {
			return key;
		}
	}
	return '';
}

/**
 * ICS の日時フィールドの生値と終日判定を返す。
 * @param {Object} event ICS のイベントオブジェクト
 * @param {string} fieldName フィールド名（'dtstart' 等）
 * @returns {{raw:string,allDay:boolean}} 生値と終日フラグ
 */
function getIcsDateTimeInfo_(event, fieldName) {
	const key = findIcsFieldKey_(event, fieldName);
	const raw = key ? event[key] || '' : '';
	return {
		raw,
		allDay: /;value=date/i.test(key) || /^\d{8}$/.test(raw),
	};
}

/**
 * ICS の日時表記（YYYYMMDD, YYYYMMDDTHHMMSSZ 等）を Date に変換する。
 * @param {string} value ICS の日時文字列
 * @param {boolean} allDay 終日フラグ
 * @returns {Date|null} 変換結果の Date、失敗時は null
 */
function toDateFromIcsValue_(value, allDay) {
	const text = String(value || '').trim();
	if (!text) {
		return null;
	}

	if (/^\d{8}$/.test(text)) {
		const year = Number(text.slice(0, 4));
		const month = Number(text.slice(4, 6)) - 1;
		const day = Number(text.slice(6, 8));
		return new Date(Date.UTC(year, month, day));
	}

	if (allDay) {
		const datePart = text.slice(0, 8);
		if (/^\d{8}$/.test(datePart)) {
			const year = Number(datePart.slice(0, 4));
			const month = Number(datePart.slice(4, 6)) - 1;
			const day = Number(datePart.slice(6, 8));
			return new Date(Date.UTC(year, month, day));
		}
	}

	const basicMatch = text.match(
		/^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})(Z)?$/i,
	);
	if (basicMatch) {
		const iso = `${basicMatch[1]}-${basicMatch[2]}-${basicMatch[3]}T${basicMatch[4]}:${basicMatch[5]}:${basicMatch[6]}${basicMatch[7] ? 'Z' : ''}`;
		return new Date(iso);
	}

	if (/^\d{4}-\d{2}-\d{2}T/.test(text)) {
		return new Date(text);
	}

	return new Date(text);
}

/**
 * Outlook イベントを取得する。ICS URL が設定されていれば ICS をパースして返す。
 * 新仕様: RRULE を展開して occurrence を生成する。
 * @param {Date} startDate 取得開始日時
 * @param {Date} endDate 取得終了日時
 * @returns {Array<Object>} 正規化されたoccurrence 配列
 */
function fetchOutlookEvents(startDate, endDate) {
	const icsText = fetchIcs();
	if (icsText) {
		return parseIcs(icsText, startDate, endDate);
	}

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
 * Outlook Graph API を呼び出し、401 ならトークンリフレッシュして再試行する。
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
 * Outlook にイベントを作成する (Graph API POST)。
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
 * Outlook のイベントを更新する (Graph API PATCH)。
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
 * Outlook のイベントを削除する (Graph API DELETE)。
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
 * Outlook API のベースパス（カレンダー ID を含む）を返す。
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
 * スクリプトプロパティからアクセストークンを取得し、無ければリフレッシュを試みる。
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
 * Outlook 用の説明文を組み立てる（`googleSyncKey` を付加）。
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
 * Graph API のイベントを内部で扱う正規化形式に変換する。
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
 * ICS 由来または内部形式の Outlook イベントを正規化する。
 * @param {Object} event イベントオブジェクト
 * @returns {Object} 正規化されたイベントオブジェクト
 */
function normalizeOutlookEvent_(event) {
	const startInfo = getIcsDateTimeInfo_(event, 'dtstart');
	const endInfo = getIcsDateTimeInfo_(event, 'dtend');
	const recurrenceId = getIcsFieldValue_(event, 'recurrence-id');
	const startValue = startInfo.raw || '';
	const endValue = endInfo.raw || '';
	const isAllDay = startInfo.allDay || endInfo.allDay;

	// ICS の TRANSP と CLASS を抽出してマッピング
	const transp = getIcsFieldValue_(event, 'transp');
	const classValue = getIcsFieldValue_(event, 'class');
	const showAs = mapIcsTranspToShowAs_(transp);
	const sensitivity = mapIcsClassToSensitivity_(classValue);

	return {
		id: event.uid || event.id || Utilities.getUuid(),
		subject: event.summary || '',
		location: event.location || '',
		description: event.description || '',
		start: isAllDay
			? {
					date: normalizeIcsDateTime_(startValue),
					timeZone: SYNC_TIMEZONE,
				}
			: {
					dateTime: normalizeIcsDateTime_(startValue),
					timeZone: SYNC_TIMEZONE,
				},
		end: isAllDay
			? {
					date: normalizeIcsDateTime_(endValue),
					timeZone: SYNC_TIMEZONE,
				}
			: {
					dateTime: normalizeIcsDateTime_(endValue),
					timeZone: SYNC_TIMEZONE,
				},
		showAs: showAs,
		sensitivity: sensitivity,
		isAllDay: Boolean(isAllDay),
		recurrenceId: recurrenceId ? normalizeIcsDateTime_(recurrenceId) : '',
		// occurrence 識別用フィールド
		uid: event.uid || event.id || Utilities.getUuid(),
		occurrenceDate: normalizeOccurrenceDateText_(
			recurrenceId
				? normalizeIcsDateTime_(recurrenceId)
				: normalizeIcsDateTime_(startValue),
		),
		raw: event,
	};
}

/**
 * ICS の日時表現を正規化して文字列にするユーティリティ。
 * @param {string} value ICS の日時値
 * @returns {string} 正規化済み日時文字列
 */
function normalizeIcsDateTime_(value) {
	if (!value) {
		return '';
	}

	if (/^\d{8}$/.test(value)) {
		return `${value.slice(0, 4)}-${value.slice(4, 6)}-${value.slice(6, 8)}`;
	}

	const parsed = toDateFromIcsValue_(value, false);
	if (!parsed) {
		return value;
	}

	if (value.endsWith('Z')) {
		return Utilities.formatDate(parsed, SYNC_TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
	}

	return Utilities.formatDate(parsed, SYNC_TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
}

/**
 * Outlook Graph API に渡すイベントリソースを構築する。
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
			mapTransparencyToShowAs(eventData.transparency) ||
			'busy',
		sensitivity:
			eventData.sensitivity ||
			mapVisibilityToSensitivity(eventData.visibility) ||
			'normal',
	};
}

/**
 * 指定日時を同期タイムゾーンで API 用日時（ISO 形式、オフセットなし）にフォーマットする。
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
 * ICS の TRANSP 値を Outlook Graph API の showAs 値にマッピングする。
 * @param {string} transp ICS の TRANSP 値（'TRANSPARENT' or 'OPAQUE'）
 * @returns {string} Outlook の showAs 値（'free' or 'busy'）
 */
function mapIcsTranspToShowAs_(transp) {
	const value = String(transp || '').toUpperCase();
	if (value === 'TRANSPARENT') {
		return 'free';
	}
	return 'busy'; // OPAQUE がデフォルト、不明な値も busy
}

/**
 * ICS の CLASS 値を Outlook Graph API の sensitivity 値にマッピングする。
 * @param {string} classValue ICS の CLASS 値（'PUBLIC', 'PRIVATE', 'CONFIDENTIAL'）
 * @returns {string} Outlook の sensitivity 値（'normal', 'private', 'confidential'）
 */
function mapIcsClassToSensitivity_(classValue) {
	const value = String(classValue || '').toUpperCase();
	if (value === 'PRIVATE') {
		return 'private';
	}
	if (value === 'CONFIDENTIAL') {
		return 'confidential';
	}
	return 'normal'; // PUBLIC がデフォルト、不明な値も normal
}

/**
 * Outlook イベント（ICS由来）から Google 形式の recurrence 配列を抽出する。
 * RRULEやEXDATEをGoogle形式に変換する。
 * @param {Object} outlookEvent Outlook のイベントオブジェクト（ICS由来）
 * @returns {Array<string>|null} Google 形式の recurrence 配列（RRULE形式）またはnull
 */
function extractRecurrenceFromOutlookEvent(outlookEvent) {
	if (!outlookEvent) {
		return null;
	}

	const recurrenceRules = [];

	// RRULE を取得
	const rrule = getIcsFieldValue_(outlookEvent, 'rrule');
	if (rrule) {
		recurrenceRules.push(`RRULE:${rrule}`);
	}

	// EXDATE を取得（複数の場合も処理）
	const exdateFieldKey = findIcsFieldKey_(outlookEvent, 'exdate');
	if (exdateFieldKey) {
		const exdateValue = outlookEvent[exdateFieldKey];
		if (exdateValue) {
			// EXDATE は複数行ある可能性があるため、分割して処理
			const exdates = Array.isArray(exdateValue) ? exdateValue : [exdateValue];
			for (const exdate of exdates) {
				if (exdate) {
					recurrenceRules.push(`EXDATE:${exdate}`);
				}
			}
		}
	}

	return recurrenceRules.length > 0 ? recurrenceRules : null;
}

/**
 * Google イベント（recurrence情報を含む）から Outlook 形式の recurrence オブジェクトを構築する。
 * Google の recurrence 配列（RRULE形式）を Outlook の recurrence オブジェクトに変換する。
 * @param {Array<string>} googleRecurrence Google 形式の recurrence 配列（例: ["RRULE:FREQ=DAILY"]）
 * @returns {Object|null} Outlook の recurrence オブジェクトまたは null
 */
function buildOutlookRecurrenceFromGoogle(googleRecurrence) {
	if (!googleRecurrence || googleRecurrence.length === 0) {
		return null;
	}

	// Google の recurrence は RRULE や EXDATE 形式の文字列配列
	// これを Outlook Graph API の recurrence オブジェクトに変換する
	const rruleLines = googleRecurrence
		.filter((line) => line && line.toUpperCase().startsWith('RRULE'))
		.map((line) => line.substring(6)); // "RRULE:" を削除

	const exdateLines = googleRecurrence
		.filter((line) => line && line.toUpperCase().startsWith('EXDATE'))
		.map((line) => line.substring(6)); // "EXDATE:" を削除

	if (rruleLines.length === 0) {
		return null;
	}

	// RRULEを解析してOutlook形式に変換
	const rrule = rruleLines[0];
	const recurrenceObj = parseRruleForOutlook_(rrule);

	// EXDATE を recurrenceObj に追加
	if (exdateLines.length > 0) {
		recurrenceObj.recurrenceTimeZone = SYNC_TIMEZONE;
		// Outlook Graph API の recurrence 形式では、exceptionsフィールドで例外日を指定する
		// ここでは EXDATE の日付を exceptions として設定（簡易実装）
		// 詳細は Outlook Graph API ドキュメント参照
	}

	return recurrenceObj;
}

/**
 * RRULEを解析してOutlook Graph API 互換の recurrence オブジェクトを構築する。
 * @param {string} rruleText RRULE文字列（例: "FREQ=DAILY;INTERVAL=1"）
 * @returns {Object} Outlook の recurrence オブジェクト
 */
function parseRruleForOutlook_(rruleText) {
	const pattern = parseRrule_(rruleText);

	// Outlook 形式に変換
	const outlookRecurrence = {
		pattern: {
			type: mapFreqToOutlookRecurrenceType_(pattern.freq),
			interval: pattern.interval || 1,
		},
		range: {
			type: 'endDate',
			startDate: new Date().toISOString().split('T')[0],
		},
	};

	// COUNT がある場合
	if (pattern.count && pattern.count > 0) {
		outlookRecurrence.range.type = 'numbered';
		outlookRecurrence.range.numberOfOccurrences = pattern.count;
	}

	// UNTIL がある場合
	if (pattern.until) {
		outlookRecurrence.range.type = 'endDate';
		outlookRecurrence.range.endDate = Utilities.formatDate(
			pattern.until,
			SYNC_TIMEZONE,
			'yyyy-MM-dd',
		);
	}

	// BYDAY 情報を追加
	if (pattern.byday && pattern.byday.length > 0) {
		outlookRecurrence.pattern.daysOfWeek = pattern.byday
			.map((day) => mapWeekdayToOutlookDay_(day.weekday))
			.filter(Boolean);
	}

	return outlookRecurrence;
}

/**
 * FREQ値をOutlook形式の recurrenceType にマッピングする。
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
 * 曜日番号（0=Sun...6=Sat）を Outlook の dayOfWeek 値にマッピングする。
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

/**
 * ICS テキストをパースしてイベント配列に変換する（RRULE 展開なし）。
 * マスターイベント（RRULE付き）をそのまま返す。
 * @param {string} icsText ICS の生テキスト
 * @param {Date} startDate 取得開始日時（ウィンドウ）
 * @param {Date} endDate 取得終了日時（ウィンドウ）
 * @returns {Array<Object>} マスターイベント配列（RRULE情報を保持）
 */
function parseIcs(icsText, startDate, endDate) {
	const lines = unfoldIcsLines_(String(icsText || '').split(/\r?\n/));
	const events = [];
	let current = null;

	for (const line of lines) {
		if (line === 'BEGIN:VEVENT') {
			current = {};
			continue;
		}

		if (line === 'END:VEVENT') {
			if (current) {
				events.push(current);
			}
			current = null;
			continue;
		}

		if (!current) {
			continue;
		}

		const colonIndex = line.indexOf(':');
		if (colonIndex === -1) {
			continue;
		}

		const key = line.slice(0, colonIndex).toLowerCase();
		const value = line.slice(colonIndex + 1);
		current[key] = value;
	}

	// RRULE ごとにグループ化：同じ UID で RECURRENCE-ID がない（マスター）イベントでグループ化
	const groupedByUid = {};
	for (const event of events) {
		const uid = event.uid || Utilities.getUuid();
		if (!groupedByUid[uid]) {
			groupedByUid[uid] = [];
		}
		groupedByUid[uid].push(event);
	}

	// 各グループを展開
	const occurrences = [];
	for (const uid in groupedByUid) {
		const group = groupedByUid[uid];
		const expanded = expandRecurringSeries_(group, startDate, endDate);
		occurrences.push(...expanded);
	}

	return occurrences;
}

/**
 * ICS のイベントオブジェクトを Outlook イベント形式に正規化する（RRULE保持版）。
 * @param {Object} event ICS から抽出したイベントオブジェクト
 * @returns {Object} 正規化された Outlook イベントオブジェクト
 */
function normalizeOutlookEventFromIcs_(event) {
	const startInfo = getIcsDateTimeInfo_(event, 'dtstart');
	const endInfo = getIcsDateTimeInfo_(event, 'dtend');
	const startValue = startInfo.raw || '';
	const endValue = endInfo.raw || '';
	const isAllDay = startInfo.allDay || endInfo.allDay;

	// ICS の TRANSP と CLASS を抽出してマッピング
	const transp = getIcsFieldValue_(event, 'transp');
	const classValue = getIcsFieldValue_(event, 'class');
	const showAs = mapIcsTranspToShowAs_(transp);
	const sensitivity = mapIcsClassToSensitivity_(classValue);

	// RRULE を抽出
	const rrule = getIcsFieldValue_(event, 'rrule');

	return {
		id: event.uid || event.id || Utilities.getUuid(),
		subject: event.summary || '',
		location: event.location || '',
		description: event.description || '',
		start: isAllDay
			? {
					date: normalizeIcsDateTime_(startValue),
					timeZone: SYNC_TIMEZONE,
				}
			: {
					dateTime: normalizeIcsDateTime_(startValue),
					timeZone: SYNC_TIMEZONE,
				},
		end: isAllDay
			? {
					date: normalizeIcsDateTime_(endValue),
					timeZone: SYNC_TIMEZONE,
				}
			: {
					dateTime: normalizeIcsDateTime_(endValue),
					timeZone: SYNC_TIMEZONE,
				},
		showAs: showAs,
		sensitivity: sensitivity,
		isAllDay: Boolean(isAllDay),
		recurrence: rrule ? [`RRULE:${rrule}`] : null,
		raw: event,
	};
}
