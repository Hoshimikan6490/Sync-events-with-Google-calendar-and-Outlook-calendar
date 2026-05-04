const SYNC_TIMEZONE = 'Asia/Tokyo';
const LOOKBACK_MONTHS = 1;
const O365_GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
const O365_AUTH_BASE = 'https://login.microsoftonline.com';

/**
 * カレンダー間の同期処理をエントリポイントとして実行する。
 * @param void
 * @returns void
 */
function syncCalendars() {
	const windowRange = getSyncWindow_();
	const googleEvents = getGoogleEvents(windowRange.start, windowRange.end);
	const outlookEvents = fetchOutlookEvents(windowRange.start, windowRange.end);

	const googleMaps = buildGoogleMaps(googleEvents);
	const outlookMaps = buildOutlookMaps(outlookEvents);
	const syncedIdSets = buildSyncedIdSets_(googleMaps, outlookMaps);

	const stats = {
		googleToOutlook: { create: 0, update: 0, delete: 0 },
		outlookToGoogle: { create: 0, update: 0, delete: 0 },
	};

	syncGoogleToOutlook(googleEvents, outlookMaps, syncedIdSets, stats);
	syncOutlookToGoogle(outlookEvents, googleMaps, syncedIdSets, stats);

	outputSummaryLog(stats);
}

/**
 * Google カレンダーから Outlook へイベントを同期する。
 * @param {Array<Object>} googleEvents Google 側のイベント配列
 * @param {Object} outlookMaps Outlook 側の参照マップ (byId, byGoogleId)
 * @param {Object} syncedIdSets 同期済みイベント ID の集合オブジェクト
 * @param {Object} stats 同期集計オブジェクト
 * @returns void
 */
function syncGoogleToOutlook(googleEvents, outlookMaps, syncedIdSets, stats) {
	const syncedGoogleEventIds = syncedIdSets.googleInOutlook;
	const sourceGoogleIds = new Set(googleEvents.map((event) => event.id));

	for (const googleEvent of googleEvents) {
		const ids = parseIds(googleEvent.description);

		// 同期元の概要欄に outlook_id があるものは、元が Outlook から同期されてきたイベント
		// なので Google -> Outlook に対しては処理しない（同期し返さない）
		if (ids.outlookId) {
			continue;
		}
		const payload = buildOutlookPayloadFromGoogleEvent_(
			googleEvent,
			ids.repeat,
		);
		const hasSyncedOutlookEvent = syncedGoogleEventIds.has(googleEvent.id);
		const targetEvent = hasSyncedOutlookEvent
			? resolveOutlookTargetEvent_(googleEvent, ids, outlookMaps)
			: null;

		if (!hasSyncedOutlookEvent || !targetEvent) {
			const created = createOutlookEvent(payload);
			if (created && created.id) {
				logAction('Google → Outlook', 'create', googleEvent.summary);
				stats.googleToOutlook.create += 1;
				syncedGoogleEventIds.add(googleEvent.id);
			}
			continue;
		}

		// 対象 Outlook イベントに google_id が設定されていなければ設定する
		const targetIds = parseIds(targetEvent.description);
		if (targetIds.googleId !== googleEvent.id) {
			updateOutlookEventWithGoogleId_(targetEvent, googleEvent.id, ids.repeat);
		}

		const comparePayload = Object.assign({}, payload);
		delete comparePayload.body;
		const nextPayload = mergeOutlookEventPayload_(targetEvent, comparePayload);
		const nextDescription = String(nextPayload.description || '');
		const newDescription = String(
			payload.body && payload.body.content ? payload.body.content : '',
		);
		const shouldUpdate =
			shouldUpdateOutlookEvent_(targetEvent, nextPayload) ||
			normalizeDescriptionText_(nextDescription) !==
				normalizeDescriptionText_(newDescription);

		if (shouldUpdate) {
			nextPayload.body = {
				contentType: 'text',
				content: newDescription,
			};
			updateOutlookEvent(targetEvent.id, nextPayload);
			logAction('Google → Outlook', 'update', googleEvent.summary);
			stats.googleToOutlook.update += 1;
		}
	}

	for (const [
		googleId,
		linkedOutlookEvents,
	] of outlookMaps.byGoogleId.entries()) {
		if (sourceGoogleIds.has(googleId)) {
			continue;
		}
		for (const outlookEvent of linkedOutlookEvents) {
			deleteOutlookEvent(outlookEvent.id);
			logAction(
				'Google → Outlook',
				'delete',
				outlookEvent.subject || outlookEvent.summary || 'イベント',
			);
			stats.googleToOutlook.delete += 1;
		}
	}
}

/**
 * Outlook から Google へイベントを同期する。
 * @param {Array<Object>} outlookEvents Outlook 側のイベント配列
 * @param {Object} googleMaps Google 側の参照マップ (byId, byOutlookId)
 * @param {Object} syncedIdSets 同期済みイベント ID の集合オブジェクト
 * @param {Object} stats 同期集計オブジェクト
 * @returns void
 */
function syncOutlookToGoogle(outlookEvents, googleMaps, syncedIdSets, stats) {
	const syncedOutlookEventIds = syncedIdSets.outlookInGoogle;
	const sourceOutlookIds = new Set(outlookEvents.map((event) => event.id));

	for (const outlookEvent of outlookEvents) {
		const ids = parseIds(outlookEvent.description);

		// 同期元の概要欄に google_id があるものは、元が Google から同期されてきたイベント
		// なので Outlook -> Google に対しては処理しない（同期し返さない）
		if (ids.googleId) {
			continue;
		}
		const payload = buildGooglePayloadFromOutlookEvent_(
			outlookEvent,
			ids.repeat,
		);
		const hasSyncedGoogleEvent = syncedOutlookEventIds.has(outlookEvent.id);
		const targetEvent = hasSyncedGoogleEvent
			? resolveGoogleTargetEvent_(outlookEvent, ids, googleMaps)
			: null;

		if (!hasSyncedGoogleEvent || !targetEvent) {
			const created = createGoogleEvent(payload);
			if (created && created.id) {
				logAction(
					'Outlook → Google',
					'create',
					outlookEvent.subject || outlookEvent.summary || 'イベント',
				);
				stats.outlookToGoogle.create += 1;
				syncedOutlookEventIds.add(outlookEvent.id);
			}
			continue;
		}

		// 対象 Google イベントに outlook_id が設定されていなければ設定する
		const targetIds = parseIds(targetEvent.description);
		if (targetIds.outlookId !== outlookEvent.id) {
			updateGoogleEventWithOutlookId_(targetEvent, outlookEvent.id, ids.repeat);
		}

		const nextPayload = mergeGoogleEventPayload_(targetEvent, payload);
		if (shouldUpdateGoogleEvent_(targetEvent, nextPayload)) {
			updateGoogleEvent(targetEvent.id, nextPayload);
			logAction(
				'Outlook → Google',
				'update',
				outlookEvent.subject || outlookEvent.summary || 'イベント',
			);
			stats.outlookToGoogle.update += 1;
		}
	}

	for (const [
		outlookId,
		linkedGoogleEvents,
	] of googleMaps.byOutlookId.entries()) {
		if (sourceOutlookIds.has(outlookId)) {
			continue;
		}
		for (const googleEvent of linkedGoogleEvents) {
			deleteGoogleEvent(googleEvent.id);
			logAction(
				'Outlook → Google',
				'delete',
				googleEvent.subject || googleEvent.summary || 'イベント',
			);
			stats.outlookToGoogle.delete += 1;
		}
	}
}

/**
 * 同期済みイベント ID の一覧セットを構築する。
 * @param {Object} googleMaps Google 側の参照マップ
 * @param {Object} outlookMaps Outlook 側の参照マップ
 * @returns {{googleInOutlook:Set<string>,outlookInGoogle:Set<string>}} 同期済み ID セット
 */
function buildSyncedIdSets_(googleMaps, outlookMaps) {
	return {
		googleInOutlook: new Set(outlookMaps.byGoogleId.keys()),
		outlookInGoogle: new Set(googleMaps.byOutlookId.keys()),
	};
}

/**
 * Google イベント配列から ID をキーにした Map を構築する。
 * @param {Array<Object>} googleEvents Google 側のイベント配列
 * @returns {{byId:Map,byOutlookId:Map}} 生成したマップオブジェクト
 */
function buildGoogleMaps(googleEvents) {
	const byId = new Map();
	const byOutlookId = new Map();

	for (const event of googleEvents) {
		byId.set(event.id, event);
		const ids = parseIds(event.description);
		if (ids.outlookId) {
			if (!byOutlookId.has(ids.outlookId)) {
				byOutlookId.set(ids.outlookId, []);
			}
			byOutlookId.get(ids.outlookId).push(event);
		}
	}

	return { byId, byOutlookId };
}

/**
 * Outlook イベント配列から ID をキーにした Map を構築する。
 * @param {Array<Object>} outlookEvents Outlook 側のイベント配列
 * @returns {{byId:Map,byGoogleId:Map}} 生成したマップオブジェクト
 */
function buildOutlookMaps(outlookEvents) {
	const byId = new Map();
	const byGoogleId = new Map();

	for (const event of outlookEvents) {
		byId.set(event.id, event);
		const ids = parseIds(event.description);
		if (ids.googleId) {
			if (!byGoogleId.has(ids.googleId)) {
				byGoogleId.set(ids.googleId, []);
			}
			byGoogleId.get(ids.googleId).push(event);
		}
	}

	return { byId, byGoogleId };
}

/**
 * イベントの説明文から同期用の ID 情報を抽出する。
 * @param {string} description イベントの説明テキスト
 * @returns {{outlookId:string,googleId:string,repeat:number}} 抽出した ID 情報
 */
function parseIds(description) {
	let text = normalizeDescriptionText_(description) || '';

	// convert literal backslash-n sequences to real newlines (some sources escape newlines)
	text = text.replace(/\\n/g, '\n');

	function safeSplitField(sourceText, variant) {
		const lower = sourceText.toLowerCase();
		const v = String(variant || '').toLowerCase();

		const idx = lower.indexOf(v);
		if (idx === -1) return '';

		const tail = sourceText.slice(idx + v.length);
		// split on actual newline or literal "\\n" if present

		return tail.split(/\r?\n|\\n/)[0].trim();
	}

	const outlookId = safeSplitField(text, 'outlook_id:');
	const googleId = safeSplitField(text, 'google_id:');
	const repeatRaw = safeSplitField(text, 'Repeat:');
	const repeat = repeatRaw ? Number(repeatRaw.replace(/[^0-9]/g, '')) || 0 : 0;

	return { outlookId, googleId, repeat };
}

/**
 * イベント説明文を ID 抽出しやすいプレーンテキストへ整形する。
 * @param {string} description イベント説明文（HTML/テキスト）
 * @returns {string} 整形後のテキスト
 */
function normalizeDescriptionText_(description) {
	return String(description || '')
		.replace(/\\r\\n/g, '\n')
		.replace(/\\n/g, '\n')
		.replace(/<br\s*\/?\s*>/gi, '\n')
		.replace(/<\/(?:div|p|li|tr|h[1-6])>/gi, '\n')
		.replace(/<[^>]+>/g, ' ')
		.replace(/&nbsp;/gi, ' ')
		.replace(/&amp;/gi, '&')
		.replace(/&lt;/gi, '<')
		.replace(/&gt;/gi, '>')
		.replace(/\r/g, '')
		.replace(/：/g, ':')
		.replace(/＝/g, '=')
		.replace(/＿/g, '_');
}

/**
 * Google イベントに対応する Outlook 側のターゲットイベントを解決する。
 * @param {Object} googleEvent Google 側のイベントオブジェクト
 * @param {{outlookId:string,googleId:string,repeat:number}} ids 抽出済み ID 情報
 * @param {Object} outlookMaps Outlook 側の参照マップ
 * @returns {Object|null} 対応する Outlook イベント、無ければ null
 */
function resolveOutlookTargetEvent_(googleEvent, ids, outlookMaps) {
	if (ids.outlookId && outlookMaps.byId.has(ids.outlookId)) {
		return outlookMaps.byId.get(ids.outlookId);
	}

	if (outlookMaps.byGoogleId.has(googleEvent.id)) {
		const candidates = outlookMaps.byGoogleId.get(googleEvent.id);
		if (Array.isArray(candidates) && candidates.length > 0) {
			return candidates[0];
		}
	}

	return null;
}

/**
 * Outlook イベントに対応する Google 側のターゲットイベントを解決する。
 * @param {Object} outlookEvent Outlook 側のイベントオブジェクト
 * @param {{outlookId:string,googleId:string,repeat:number}} ids 抽出済み ID 情報
 * @param {Object} googleMaps Google 側の参照マップ
 * @returns {Object|null} 対応する Google イベント、無ければ null
 */
function resolveGoogleTargetEvent_(outlookEvent, ids, googleMaps) {
	if (ids.googleId && googleMaps.byId.has(ids.googleId)) {
		return googleMaps.byId.get(ids.googleId);
	}

	if (googleMaps.byOutlookId.has(outlookEvent.id)) {
		const candidates = googleMaps.byOutlookId.get(outlookEvent.id);
		if (Array.isArray(candidates) && candidates.length > 0) {
			return candidates[0];
		}
	}

	return null;
}

/**
 * 同期アクションをコンソール出力するユーティリティ。
 * @param {string} direction 同期方向のラベル（例: 'Google → Outlook'）
 * @param {string} action アクション（'create'|'update'|'delete'）
 * @param {string} eventName イベント名
 * @returns void
 */
function logAction(direction, action, eventName) {
	const label =
		action === 'create' ? '作成' : action === 'update' ? '更新' : '削除';
	console.log(`「${eventName}」が${direction} に${label}されました`);
}

/**
 * 同期結果のサマリをコンソールに出力する。
 * @param {Object} stats 同期の集計情報オブジェクト
 * @returns void
 */
function outputSummaryLog(stats) {
	console.log('=== 同期結果 ===');
	console.log(
		`Google → Outlook: ${stats.googleToOutlook.create}件の作成、${stats.googleToOutlook.update}件の更新、${stats.googleToOutlook.delete}件の削除を行いました。`,
	);
	console.log(
		`Outlook → Google: ${stats.outlookToGoogle.create}件の作成、${stats.outlookToGoogle.update}件の更新、${stats.outlookToGoogle.delete}件の削除を行いました。`,
	);
}

/**
 * `syncCalendars` を30分毎に実行するトリガーをセットする。
 * @param void
 * @returns void
 */
function installThirtyMinuteTrigger() {
	const handlerFunction = 'syncCalendars';
	const triggers = ScriptApp.getProjectTriggers();
	for (const trigger of triggers) {
		if (trigger.getHandlerFunction() === handlerFunction) {
			ScriptApp.deleteTrigger(trigger);
		}
	}

	ScriptApp.newTrigger(handlerFunction).timeBased().everyMinutes(30).create();
}

/**
 * 同期対象の日時ウィンドウ（開始と終了）を生成する。
 * @param void
 * @returns {{start:Date,end:Date}} 同期ウィンドウの開始日時と終了日時
 */
function getSyncWindow_() {
	const start = new Date();
	start.setHours(0, 0, 0, 0);
	const end = new Date(start);
	end.setMonth(end.getMonth() + LOOKBACK_MONTHS);
	end.setHours(23, 59, 59, 999);
	return { start, end };
}

/**
 * Google イベントから Outlook 作成用のペイロードを構築する。
 * @param {Object} googleEvent Google のイベントオブジェクト
 * @param {number} repeat 繰り返しインデックス（0 or positive）
 * @returns {Object} Outlook API 用のイベントペイロード
 */
function buildOutlookPayloadFromGoogleEvent_(googleEvent, repeat) {
	const isAllDay = Boolean(
		(googleEvent.start && googleEvent.start.date) ||
		(googleEvent.end && googleEvent.end.date),
	);
	const startValue =
		googleEvent.startDateTime ||
		(googleEvent.start &&
			(googleEvent.start.dateTime || googleEvent.start.date));
	const endValue =
		googleEvent.endDateTime ||
		(googleEvent.end && (googleEvent.end.dateTime || googleEvent.end.date));
	const start = isAllDay
		? {
				dateTime: toOutlookAllDayDateTime_(startValue),
				timeZone: SYNC_TIMEZONE,
			}
		: {
				dateTime: buildOutlookDateTimeString_(startValue, false),
				timeZone: SYNC_TIMEZONE,
			};
	const end = isAllDay
		? {
				dateTime: toOutlookAllDayDateTime_(endValue),
				timeZone: SYNC_TIMEZONE,
			}
		: {
				dateTime: buildOutlookDateTimeString_(endValue, false),
				timeZone: SYNC_TIMEZONE,
			};

	return {
		subject: googleEvent.summary || '',
		body: {
			contentType: 'text',
			content: buildOutlookDescription(googleEvent, repeat, googleEvent.id),
		},
		start,
		end,
		isAllDay: isAllDay,
		showAs: mapGoogleTransparencyToOutlook_(googleEvent.transparency),
		sensitivity: mapGoogleVisibilityToOutlook_(googleEvent.visibility),
	};
}

/**
 * Outlook イベントから Google 作成用のペイロードを構築する。
 * @param {Object} outlookEvent Outlook のイベントオブジェクト
 * @param {number} repeat 繰り返しインデックス（0 or positive）
 * @returns {Object} Google Calendar API 用のイベントペイロード
 */
function buildGooglePayloadFromOutlookEvent_(outlookEvent, repeat) {
	const start = normalizeOutlookCalendarDateTime_(
		outlookEvent.start,
		outlookEvent,
	);
	const end =
		normalizeOutlookCalendarDateTime_(outlookEvent.end, outlookEvent) ||
		buildDefaultGoogleEndFromStart_(start);
	const safeStart = start || buildDefaultGoogleStart_();
	const safeEnd =
		end && (end.date || end.dateTime)
			? end
			: buildDefaultGoogleEndFromStart_(safeStart);

	return {
		summary: outlookEvent.subject || '',
		description: buildGoogleDescription(outlookEvent, repeat, outlookEvent.id),
		start: safeStart,
		end: safeEnd,
		transparency: mapOutlookShowAsToGoogle_(outlookEvent.showAs),
		visibility: mapOutlookSensitivityToGoogle_(outlookEvent.sensitivity),
	};
}

/**
 * デフォルトの Google イベント開始日時を現在時刻で構築する。
 * @param void
 * @returns {{dateTime:string,timeZone:string}} Google イベントの開始情報
 */
function buildDefaultGoogleStart_() {
	const startDateTime = new Date();
	return {
		dateTime: Utilities.formatDate(
			startDateTime,
			SYNC_TIMEZONE,
			"yyyy-MM-dd'T'HH:mm:ss",
		),
		timeZone: SYNC_TIMEZONE,
	};
}

/**
 * Outlook 形式の日時文字列を構築する（終日の場合も対応）。
 * @param {string|Date} value 日時または日付文字列/Date オブジェクト
 * @param {boolean} allDay 終日フラグ
 * @returns {string} Outlook 用の日時文字列
 */
function buildOutlookDateTimeString_(value, allDay) {
	if (!value) {
		return '';
	}

	if (allDay) {
		if (/^\d{4}-\d{2}-\d{2}$/.test(value)) {
			return `${value}T00:00:00`;
		}
		if (value instanceof Date) {
			return Utilities.formatDate(
				value,
				SYNC_TIMEZONE,
				"yyyy-MM-dd'T'00:00:00",
			);
		}
	}

	return toIsoStringInTimeZone_(value);
}

/**
 * 終日イベント用の Outlook 日付時刻文字列を作成する。
 * @param {string|Date} value 日付または日時
 * @returns {string} 終日用の日時文字列（例: YYYY-MM-DDT00:00:00）
 */
function toOutlookAllDayDateTime_(value) {
	if (!value) {
		return '';
	}

	if (/^\d{4}-\d{2}-\d{2}$/.test(String(value))) {
		return `${value}T00:00:00`;
	}

	if (value instanceof Date) {
		return Utilities.formatDate(value, SYNC_TIMEZONE, "yyyy-MM-dd'T'00:00:00");
	}

	const datePart = String(value).slice(0, 10);
	if (/^\d{4}-\d{2}-\d{2}$/.test(datePart)) {
		return `${datePart}T00:00:00`;
	}

	return '';
}

/**
 * Outlook API 由来の日時情報を Google 用に正規化する。
 * @param {Object|string} value Outlook の日時フィールドまたは文字列
 * @param {Object} outlookEvent 該当 Outlook イベント（オプション）
 * @returns {Object|null} 正規化された日時オブジェクトまたは null
 */
function normalizeOutlookCalendarDateTime_(value, outlookEvent) {
	if (!value) {
		return null;
	}

	const isAllDay = Boolean(
		(outlookEvent && outlookEvent.isAllDay) ||
		(value && value.date) ||
		(value && value.dateTime && /T00:00:00/.test(value.dateTime)),
	);

	if (typeof value === 'string') {
		if (!value) {
			return null;
		}
		if (isAllDay) {
			return { date: String(value).slice(0, 10), timeZone: SYNC_TIMEZONE };
		}
		return { dateTime: value, timeZone: SYNC_TIMEZONE };
	}

	if (value.dateTime) {
		if (isAllDay) {
			return {
				date: value.dateTime.slice(0, 10),
				timeZone: value.timeZone || SYNC_TIMEZONE,
			};
		}
		return {
			dateTime: value.dateTime,
			timeZone: value.timeZone || SYNC_TIMEZONE,
		};
	}

	if (value.date) {
		return {
			date: value.date,
			timeZone: value.timeZone || SYNC_TIMEZONE,
		};
	}

	return null;
}

/**
 * Outlook イベントのマージを行い次のペイロードを作る。
 * @param {Object} currentEvent 現在の Outlook イベントオブジェクト
 * @param {Object} nextPayload 更新用のペイロード
 * @returns {Object} マージ後のイベントオブジェクト
 */
function mergeOutlookEventPayload_(currentEvent, nextPayload) {
	return Object.assign({}, currentEvent, nextPayload, {
		body: nextPayload.body,
	});
}

/**
 * Google イベントのマージを行い次のペイロードを作る。
 * @param {Object} currentEvent 現在の Google イベントオブジェクト
 * @param {Object} nextPayload 更新用のペイロード
 * @returns {Object} マージ後のイベントオブジェクト
 */
function mergeGoogleEventPayload_(currentEvent, nextPayload) {
	return Object.assign({}, currentEvent, nextPayload);
}

/**
 * Outlook イベントが更新対象かどうかを判定する。
 * @param {Object} currentEvent 現在の Outlook イベント
 * @param {Object} nextPayload 比較対象の次ペイロード
 * @returns {boolean} 更新が必要なら true
 */
function shouldUpdateOutlookEvent_(currentEvent, nextPayload) {
	console.log('Comparing Outlook Event:');
	console.log(
		'Current Event:',
		JSON.stringify(normalizeOutlookEvent_(currentEvent), null, 2),
	);
	console.log(
		'Next Payload:',
		JSON.stringify(normalizeOutlookEvent_(nextPayload), null, 2),
	);

	return (
		JSON.stringify(normalizeOutlookEvent_(currentEvent)) !==
		JSON.stringify(normalizeOutlookEvent_(nextPayload))
	);
}

/**
 * Google イベントが更新対象かどうかを判定する。
 * @param {Object} currentEvent 現在の Google イベント
 * @param {Object} nextPayload 比較対象の次ペイロード
 * @returns {boolean} 更新が必要なら true
 */
function shouldUpdateGoogleEvent_(currentEvent, nextPayload) {
	return (
		JSON.stringify(normalizeGoogleEvent_(currentEvent)) !==
		JSON.stringify(normalizeGoogleEvent_(nextPayload))
	);
}

/**
 * Outlook イベントを比較しやすい形に正規化する。
 * @param {Object} event Outlook イベントオブジェクト
 * @returns {Object} 正規化されたイベント情報
 */
function normalizeOutlookEvent_(event) {
	const startValue =
		event.start && (event.start.dateTime || event.start.date)
			? event.start.dateTime || event.start.date
			: '';
	const endValue =
		event.end && (event.end.dateTime || event.end.date)
			? event.end.dateTime || event.end.date
			: '';

	return {
		subject: event.subject || '',
		body: event.body && event.body.content ? event.body.content : '',
		start: startValue,
		end: endValue,
		isAllDay: Boolean(event.isAllDay || (event.start && event.start.date)),
		showAs: event.showAs || '',
		sensitivity: event.sensitivity || '',
	};
}

/**
 * Google イベントを比較しやすい形に正規化する。
 * @param {Object} event Google イベントオブジェクト
 * @returns {Object} 正規化されたイベント情報
 */
function normalizeGoogleEvent_(event) {
	return {
		summary: event.summary || '',
		description: event.description || '',
		start:
			event.startDateTime ||
			(event.start && event.start.dateTime) ||
			(event.start && event.start.date) ||
			'',
		end:
			event.endDateTime ||
			(event.end && event.end.dateTime) ||
			(event.end && event.end.date) ||
			'',
		transparency: event.transparency || '',
		visibility: event.visibility || '',
	};
}

/**
 * Google の透明性(transparency) を Outlook の showAs にマッピングする。
 * @param {string} transparency Google の透明性値
 * @returns {string} Outlook 用の showAs 値
 */
function mapGoogleTransparencyToOutlook_(transparency) {
	return transparency === 'transparent' ? 'free' : 'busy';
}

/**
 * Outlook の showAs を Google の transparency にマッピングする。
 * @param {string} showAs Outlook の showAs 値
 * @returns {string} Google 用の transparency 値
 */
function mapOutlookShowAsToGoogle_(showAs) {
	return showAs === 'free' ? 'transparent' : 'opaque';
}

/**
 * Google の visibility を Outlook の sensitivity にマッピングする。
 * @param {string} visibility Google の visibility 値
 * @returns {string} Outlook 用の sensitivity 値
 */
function mapGoogleVisibilityToOutlook_(visibility) {
	return visibility === 'private' ? 'private' : 'normal';
}

/**
 * Outlook の sensitivity を Google の visibility にマッピングする。
 * @param {string} sensitivity Outlook の sensitivity 値
 * @returns {string} Google 用の visibility 値
 */
function mapOutlookSensitivityToGoogle_(sensitivity) {
	return sensitivity === 'private' ? 'private' : 'default';
}

/**
 * 指定タイムゾーンで ISO 風文字列を生成するユーティリティ。
 * @param {string|Date} value 日時または Date オブジェクト
 * @returns {string} 変換後の文字列
 */
function toIsoStringInTimeZone_(value) {
	if (!value) {
		return '';
	}
	if (value instanceof Date) {
		return Utilities.formatDate(value, 'UTC', "yyyy-MM-dd'T'HH:mm:ss");
	}

	const text = String(value);
	const parsed = new Date(text);
	if (!isNaN(parsed.getTime())) {
		// UTC(+Z) / RFC3339(+09:00) を UTC のまま Outlook 向け形式へ正規化する
		return Utilities.formatDate(parsed, 'UTC', "yyyy-MM-dd'T'HH:mm:ss");
	}

	// 解析できない場合のみ元の文字列を返す
	return text;
}

/**
 * Google イベントの説明に Outlook の ID を埋め込みて更新する。
 * @param {Object} googleEvent Google イベントオブジェクト
 * @param {string} outlookId Outlook のイベント ID
 * @param {number} repeat 繰り返しインデックス
 * @returns void
 */
function updateGoogleEventWithOutlookId_(googleEvent, outlookId, repeat) {
	const description = buildGoogleDescription(
		{ description: googleEvent.description },
		repeat,
		outlookId,
	);
	updateGoogleEvent(
		googleEvent.id,
		Object.assign({}, googleEvent, {
			description,
		}),
	);
}

/**
 * Outlook イベントの説明に Google の ID を埋め込みて更新する。
 * @param {Object} outlookEvent Outlook イベントオブジェクト
 * @param {string} googleId Google のイベント ID
 * @param {number} repeat 繰り返しインデックス
 * @returns void
 */
function updateOutlookEventWithGoogleId_(outlookEvent, googleId, repeat) {
	const bodyContent = buildOutlookDescription(
		{ description: outlookEvent.description },
		repeat,
		googleId,
	);
	updateOutlookEvent(
		outlookEvent.id,
		Object.assign({}, outlookEvent, {
			body: {
				contentType: 'text',
				content: bodyContent,
			},
		}),
	);
}
