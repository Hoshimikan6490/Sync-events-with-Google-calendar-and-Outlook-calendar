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

	const stats = {
		googleToOutlook: { create: 0, update: 0, delete: 0 },
		outlookToGoogle: { create: 0, update: 0, delete: 0 },
	};

	const googleToOutlookTasks = buildGoogleToOutlookSyncTasks_(
		googleEvents,
		outlookMaps,
	);
	const outlookToGoogleTasks = buildOutlookToGoogleSyncTasks_(
		outlookEvents,
		googleMaps,
	);

	executeGoogleToOutlookTasks_(googleToOutlookTasks, stats.googleToOutlook);
	executeOutlookToGoogleTasks_(outlookToGoogleTasks, stats.outlookToGoogle);

	outputSummaryLog(stats);
}

/**
 * Google カレンダーから Outlook へイベントを同期する（occurrence 単位）。
 * @param {Array<Object>} googleEvents Google 側の occurrence 配列
 * @param {Object} outlookMaps Outlook 側の参照マップ (byGoogleSyncKey, byOutlookSyncKey)
 * @param {Object} stats 同期集計オブジェクト
 * @returns void
 */

function syncGoogleToOutlook(googleEvents, outlookMaps, stats) {
	const tasks = buildGoogleToOutlookSyncTasks_(googleEvents, outlookMaps);
	executeGoogleToOutlookTasks_(
		tasks,
		stats || { create: 0, update: 0, delete: 0 },
	);
}

/**
 * Google カレンダーから Outlook への同期タスクを組み立てる。
 * @param {Array<Object>} googleEvents Google 側の occurrence 配列
 * @param {Object} outlookMaps Outlook 側の参照マップ
 * @returns {Array<Object>} 同期タスク配列
 */
function buildGoogleToOutlookSyncTasks_(googleEvents, outlookMaps) {
	const tasks = [];
	const googleSyncKeys = new Set();
	for (const googleEvent of googleEvents) {
		const googleSyncKey = generateGoogleSyncKey(googleEvent);
		googleSyncKeys.add(googleSyncKey);
		// 1. syncKey で直接検索
		let outlookEvent = outlookMaps.byGoogleSyncKey.get(googleSyncKey);

		// 2. 見つからない場合、description から outlookSyncKey を抽出して検索
		if (!outlookEvent) {
			const ids = parseIds(googleEvent.description);
			if (ids.googleSyncKey && ids.googleSyncKey !== googleSyncKey) {
				googleSyncKeys.add(ids.googleSyncKey);
			}
			if (ids.outlookSyncKey) {
				// description に記録されている Outlook Sync Key で検索
				outlookEvent = outlookMaps.byOutlookSyncKey?.get(ids.outlookSyncKey);
			}
			if (!outlookEvent && ids.googleSyncKey) {
				outlookEvent = outlookMaps.byGoogleSyncKey.get(ids.googleSyncKey);
			}
		}

		if (outlookEvent) {
			// Outlook に対応する occurrence が存在 → 更新
			const payload = buildOutlookPayloadFromGoogleEvent_(googleEvent);
			const shouldUpdate = shouldUpdateOutlookEvent_(outlookEvent, payload);

			if (shouldUpdate) {
				tasks.push({
					action: 'update',
					direction: 'Google → Outlook',
					sourceEvent: googleEvent,
					targetEvent: outlookEvent,
					payload: payload,
				});
			}
		} else {
			// Outlook に対応する occurrence がない → 作成
			const payload = buildOutlookPayloadFromGoogleEvent_(googleEvent);
			tasks.push({
				action: 'create',
				direction: 'Google → Outlook',
				sourceEvent: googleEvent,
				payload: payload,
			});
		}
	}

	for (const [
		googleSyncKey,
		outlookEvent,
	] of outlookMaps.byGoogleSyncKey.entries()) {
		if (googleSyncKeys.has(googleSyncKey)) {
			continue;
		}
		tasks.push({
			action: 'delete',
			direction: 'Google → Outlook',
			targetEvent: outlookEvent,
		});
	}

	return tasks;
}

/**
 * Outlook から Google へイベントを同期する（occurrence 単位）。
 * @param {Array<Object>} outlookEvents Outlook 側の occurrence 配列
 * @param {Object} googleMaps Google 側の参照マップ (byOutlookSyncKey, byGoogleSyncKey)
 * @param {Object} stats 同期集計オブジェクト
 * @returns void
 */
function syncOutlookToGoogle(outlookEvents, googleMaps, stats) {
	const tasks = buildOutlookToGoogleSyncTasks_(outlookEvents, googleMaps);
	executeOutlookToGoogleTasks_(
		tasks,
		stats || { create: 0, update: 0, delete: 0 },
	);
}

/**
 * Outlook から Google への同期タスクを組み立てる。
 * @param {Array<Object>} outlookEvents Outlook 側の occurrence 配列
 * @param {Object} googleMaps Google 側の参照マップ
 * @returns {Array<Object>} 同期タスク配列
 */
function buildOutlookToGoogleSyncTasks_(outlookEvents, googleMaps) {
	const tasks = [];
	const outlookSyncKeys = new Set();
	for (const outlookEvent of outlookEvents) {
		const outlookSyncKey = generateOutlookSyncKey(outlookEvent);
		outlookSyncKeys.add(outlookSyncKey);
		// 1. syncKey で直接検索
		let googleEvent = googleMaps.byOutlookSyncKey.get(outlookSyncKey);

		// 2. 見つからない場合、description から googleSyncKey を抽出して検索
		if (!googleEvent) {
			const ids = parseIds(outlookEvent.description);
			if (ids.outlookSyncKey && ids.outlookSyncKey !== outlookSyncKey) {
				outlookSyncKeys.add(ids.outlookSyncKey);
			}
			if (ids.googleSyncKey) {
				// description に記録されている Google Sync Key で検索
				googleEvent = googleMaps.byGoogleSyncKey?.get(ids.googleSyncKey);
			}
			if (!googleEvent && ids.outlookSyncKey) {
				googleEvent = googleMaps.byOutlookSyncKey.get(ids.outlookSyncKey);
			}
		}

		if (googleEvent) {
			// Google に対応する occurrence が存在 → 更新
			const payload = buildGooglePayloadFromOutlookEvent_(outlookEvent);
			const shouldUpdate = shouldUpdateGoogleEvent_(googleEvent, payload);

			if (shouldUpdate) {
				tasks.push({
					action: 'update',
					direction: 'Outlook → Google',
					sourceEvent: outlookEvent,
					targetEvent: googleEvent,
					payload: payload,
				});
			}
		} else {
			// Google に対応する occurrence がない → 作成
			const payload = buildGooglePayloadFromOutlookEvent_(outlookEvent);
			tasks.push({
				action: 'create',
				direction: 'Outlook → Google',
				sourceEvent: outlookEvent,
				payload: payload,
			});
		}
	}

	for (const [
		outlookSyncKey,
		googleEvent,
	] of googleMaps.byOutlookSyncKey.entries()) {
		if (outlookSyncKeys.has(outlookSyncKey)) {
			continue;
		}
		tasks.push({
			action: 'delete',
			direction: 'Outlook → Google',
			targetEvent: googleEvent,
		});
	}

	return tasks;
}

/**
 * Google → Outlook の同期タスクを実行する。
 * @param {Array<Object>} tasks 同期タスク配列
 * @param {Object} stats 同期集計オブジェクト
 * @returns void
 */
function executeGoogleToOutlookTasks_(tasks, stats) {
	const actionPriority = { delete: 0, update: 1, create: 2 };
	for (const task of tasks.slice().sort((left, right) => {
		return actionPriority[left.action] - actionPriority[right.action];
	})) {
		if (task.action === 'update') {
			updateOutlookEvent(task.targetEvent.id, task.payload);
			logAction(
				'Google → Outlook',
				'update',
				task.sourceEvent.subject || 'イベント',
			);
			stats.update += 1;
			continue;
		}

		if (task.action === 'create') {
			const created = createOutlookEvent(task.payload);
			if (created && created.id) {
				updateGoogleEvent(
					task.sourceEvent.id,
					Object.assign({}, task.sourceEvent, {
						description: buildGoogleDescription(
							task.sourceEvent,
							generateOutlookSyncKey(created),
						),
					}),
				);

				logAction(
					'Google → Outlook',
					'create',
					task.sourceEvent.subject || 'イベント',
				);
				stats.create += 1;
			}
			continue;
		}

		if (task.action === 'delete') {
			deleteOutlookEvent(task.targetEvent.id);
			logAction(
				'Google → Outlook',
				'delete',
				task.targetEvent.subject || 'イベント',
			);
			stats.delete += 1;
		}
	}
}

/**
 * Outlook → Google の同期タスクを実行する。
 * @param {Array<Object>} tasks 同期タスク配列
 * @param {Object} stats 同期集計オブジェクト
 * @returns void
 */
function executeOutlookToGoogleTasks_(tasks, stats) {
	const actionPriority = { delete: 0, update: 1, create: 2 };
	for (const task of tasks.slice().sort((left, right) => {
		return actionPriority[left.action] - actionPriority[right.action];
	})) {
		if (task.action === 'update') {
			updateGoogleEvent(task.targetEvent.id, task.payload);
			logAction(
				'Outlook → Google',
				'update',
				task.sourceEvent.subject || 'イベント',
			);
			stats.update += 1;
			continue;
		}

		if (task.action === 'create') {
			const created = createGoogleEvent(task.payload);
			if (created && created.id) {
				const newOutlookPayload = {
					body: {
						contentType: 'text',
						content: buildOutlookDescription(
							task.sourceEvent,
							created.id,
							generateGoogleSyncKey(created),
						),
					},
				};
				updateOutlookEvent(task.sourceEvent.id, newOutlookPayload);

				logAction(
					'Outlook → Google',
					'create',
					task.sourceEvent.subject || 'イベント',
				);
				stats.create += 1;
			}
			continue;
		}

		if (task.action === 'delete') {
			deleteGoogleEvent(task.targetEvent.id);
			logAction(
				'Outlook → Google',
				'delete',
				task.targetEvent.subject || 'イベント',
			);
			stats.delete += 1;
		}
	}
}

/**
 * Google イベント（occurrence を含む）の syncKey を生成する。
 * occurrence: google:<recurringEventId>:<originalStartTime>
 * single event: google:<eventId>
 * @param {Object} event 正規化された Google イベント
 * @returns {string} syncKey
 */
function generateGoogleSyncKey(event) {
	const occurrenceDate = normalizeOccurrenceDateText_(
		event.originalStartTime || event.occurrenceDate || event.start || null,
	);
	if (event.recurringEventId && occurrenceDate) {
		return `google:${event.recurringEventId}:${occurrenceDate}`;
	}
	return `google:${event.id}`;
}

/**
 * Outlook イベント（occurrence を含む）の syncKey を生成する。
 * occurrence: outlook:<uid>:<occurrenceDate>
 * @param {Object} event 正規化された Outlook イベント
 * @returns {string} syncKey
 */
function generateOutlookSyncKey(event) {
	const occurrenceDate = normalizeOccurrenceDateText_(
		event.recurrenceId || event.occurrenceDate || event.start || null,
	);
	if (event.uid && occurrenceDate) {
		return `outlook:${event.uid}:${occurrenceDate}`;
	}
	return `outlook:${event.id}`;
}

/**
 * Google イベント配列から syncKey をキーにした Map を構築する。
 * @param {Array<Object>} googleEvents Google 側のイベント配列
 * @returns {Object} { bySyncKey, byGoogleSyncKey, byGoogleSyncKey } のマップオブジェクト
 */
function buildGoogleMaps(googleEvents) {
	const bySyncKey = new Map();
	const byOutlookSyncKey = new Map();
	const byGoogleSyncKey = new Map();

	for (const event of googleEvents) {
		const syncKey = generateGoogleSyncKey(event);
		bySyncKey.set(syncKey, event);

		// description から googleSyncKey を抽出してセカンダリマップを構築
		const ids = parseIds(event.description);
		if (ids.outlookSyncKey) {
			byOutlookSyncKey.set(ids.outlookSyncKey, event);
		}
		if (ids.googleSyncKey) {
			byGoogleSyncKey.set(ids.googleSyncKey, event);
		}
	}

	return { bySyncKey, byOutlookSyncKey, byGoogleSyncKey };
}

/**
 * Outlook イベント配列から syncKey をキーにした Map を構築する。
 * @param {Array<Object>} outlookEvents Outlook 側のイベント配列
 * @returns {Object} { bySyncKey, byOutlookSyncKey } のマップオブジェクト
 */
function buildOutlookMaps(outlookEvents) {
	const bySyncKey = new Map();
	const byGoogleSyncKey = new Map();
	const byOutlookSyncKey = new Map();

	for (const event of outlookEvents) {
		const syncKey = generateOutlookSyncKey(event);
		bySyncKey.set(syncKey, event);

		// description から outlookSyncKey を抽出してセカンダリマップを構築
		const ids = parseIds(event.description);
		if (ids.googleSyncKey) {
			byGoogleSyncKey.set(ids.googleSyncKey, event);
		}
		if (ids.outlookSyncKey) {
			byOutlookSyncKey.set(ids.outlookSyncKey, event);
		}
	}

	return { bySyncKey, byGoogleSyncKey, byOutlookSyncKey };
}

/**
 * イベントの説明文から同期用の ID 情報を抽出する。
 * @param {string} description イベントの説明テキスト
 * @returns {{outlookSyncKey:string,googleSyncKey:string}} 抽出した ID 情報
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

	// 新仕様: description 内の同期メタは `googleSyncKey:` / `outlookSyncKey:` を使用する
	const outlookSyncKey = safeSplitField(text, 'outlookSyncKey:');
	const googleSyncKey = safeSplitField(text, 'googleSyncKey:');

	return { outlookSyncKey, googleSyncKey };
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
 * @param {{outlookSyncKey:string,googleSyncKey:string}} ids 抽出済み ID 情報
 * @param {Object} outlookMaps Outlook 側の参照マップ
 * @returns {Object|null} 対応する Outlook イベント、無ければ null
 */
function resolveOutlookTargetEvent_(googleEvent, ids, outlookMaps) {
	if (
		ids.outlookSyncKey &&
		outlookMaps.byOutlookSyncKey.has(ids.outlookSyncKey)
	) {
		return outlookMaps.byOutlookSyncKey.get(ids.outlookSyncKey);
	}

	if (outlookMaps.byGoogleSyncKey.has(ids.googleSyncKey)) {
		const candidates = outlookMaps.byGoogleSyncKey.get(ids.googleSyncKey);
		if (Array.isArray(candidates) && candidates.length > 0) {
			return candidates[0];
		}
	}

	return null;
}

/**
 * Outlook イベントに対応する Google 側のターゲットイベントを解決する。
 * @param {Object} outlookEvent Outlook 側のイベントオブジェクト
 * @param {{outlookSyncKey:string,googleSyncKey:string}} ids 抽出済み ID 情報
 * @param {Object} googleMaps Google 側の参照マップ
 * @returns {Object|null} 対応する Google イベント、無ければ null
 */
function resolveGoogleTargetEvent_(outlookEvent, ids, googleMaps) {
	if (ids.googleSyncKey && googleMaps.byGoogleSyncKey.has(ids.googleSyncKey)) {
		return googleMaps.byGoogleSyncKey.get(ids.googleSyncKey);
	}

	if (googleMaps.byOutlookSyncKey.has(outlookEvent.id)) {
		const candidates = googleMaps.byOutlookSyncKey.get(outlookEvent.id);
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
	const end = new Date(start);
	end.setMonth(end.getMonth() + LOOKBACK_MONTHS);
	end.setHours(23, 59, 59, 999);
	return { start, end };
}

/**
 * 同期キー用の occurrence 日時文字列へ正規化する。
 * @param {Object|string|Date} value 日時情報
 * @returns {string} 正規化済み日時文字列
 */
function normalizeOccurrenceDateText_(value) {
	if (!value) {
		return '';
	}

	if (value instanceof Date) {
		return Utilities.formatDate(value, 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'");
	}

	if (typeof value === 'string') {
		if (/^\d{4}-\d{2}-\d{2}$/.test(value)) {
			return value;
		}
		const normalizedValue = /[zZ]$|[+-]\d{2}:\d{2}$/.test(value)
			? value
			: `${value}+09:00`;
		const parsed = new Date(normalizedValue);
		if (!isNaN(parsed.getTime())) {
			return Utilities.formatDate(parsed, 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'");
		}
		return value;
	}

	if (value.date) {
		return String(value.date);
	}

	if (value.dateTime) {
		const parsed = new Date(value.dateTime);
		if (!isNaN(parsed.getTime())) {
			return Utilities.formatDate(parsed, 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'");
		}
		return String(value.dateTime);
	}

	return '';
}

/**
 * Google イベントから Outlook 作成用のペイロードを構築する。
 * @param {Object} googleEvent Google のイベントオブジェクト
 * @returns {Object} Outlook API 用のイベントペイロード
 */
function buildOutlookPayloadFromGoogleEvent_(googleEvent) {
	const isAllDay = Boolean(
		(googleEvent.start && googleEvent.start.date) ||
		(googleEvent.end && googleEvent.end.date),
	);
	const startValue =
		(googleEvent.start &&
			(googleEvent.start.dateTime || googleEvent.start.date)) ||
		'';
	const endValue =
		(googleEvent.end && (googleEvent.end.dateTime || googleEvent.end.date)) ||
		'';
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

	const payload = {
		subject: googleEvent.subject || '',
		body: {
			contentType: 'text',
			content: buildOutlookDescription(
				googleEvent,
				googleEvent.id,
				generateGoogleSyncKey(googleEvent),
			),
		},
		start,
		end,
		isAllDay: isAllDay,
		showAs:
			googleEvent.showAs ||
			mapGoogleTransparencyToOutlook_(googleEvent.transparency),
		sensitivity:
			googleEvent.sensitivity ||
			mapGoogleVisibilityToOutlook_(googleEvent.visibility),
		location: googleEvent.location || '',
	};

	// recurrence 情報を抽出して追加
	const googleRecurrence = extractRecurrenceFromGoogleEvent(googleEvent.raw);
	if (googleRecurrence) {
		const outlookRecurrence =
			buildOutlookRecurrenceFromGoogle(googleRecurrence);
		if (outlookRecurrence) {
			payload.recurrence = outlookRecurrence;
		}
	}

	return payload;
}

/**
 * Outlook イベントから Google 作成用のペイロードを構築する。
 * @param {Object} outlookEvent Outlook のイベントオブジェクト
 * @returns {Object} Google Calendar API 用のイベントペイロード
 */
function buildGooglePayloadFromOutlookEvent_(outlookEvent) {
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

	const payload = {
		subject: outlookEvent.subject || '',
		description: buildGoogleDescription(
			outlookEvent,
			generateOutlookSyncKey(outlookEvent),
		),
		start: safeStart,
		end: safeEnd,
		transparency: mapOutlookShowAsToGoogle_(outlookEvent.showAs),
		visibility: mapOutlookSensitivityToGoogle_(outlookEvent.sensitivity),
	};

	// recurrence 情報を抽出して追加
	const recurrence = extractRecurrenceFromOutlookEvent(outlookEvent);
	if (recurrence) {
		payload.recurrence = recurrence;
	}

	return payload;
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
	console.log('Event Name: ' + (currentEvent.subject || 'イベント'));
	const currentNormalized = normalizeOutlookEventForCompare_(currentEvent);
	const nextNormalized = normalizeOutlookEventForCompare_(nextPayload);
	const isDifferent =
		JSON.stringify(currentNormalized) !== JSON.stringify(nextNormalized);

	if (isDifferent) {
		const diffEntries = [];
		const keys = new Set([
			...Object.keys(currentNormalized),
			...Object.keys(nextNormalized),
		]);

		for (const key of keys) {
			if (
				JSON.stringify(currentNormalized[key]) !==
				JSON.stringify(nextNormalized[key])
			) {
				diffEntries.push({
					field: key,
					current: currentNormalized[key],
					next: nextNormalized[key],
				});
			}
		}

		console.log('Diff: ' + JSON.stringify(diffEntries, null, 2));
	} else {
		console.log('Diff: none');
	}

	return JSON.stringify(currentNormalized) !== JSON.stringify(nextNormalized);
}

/**
 * Outlook イベントの更新判定用に、同期メタデータを除いた比較形式へ正規化する。
 * @param {Object} event Outlook イベントオブジェクト
 * @returns {Object} 比較用に正規化されたイベント情報
 */
function normalizeOutlookEventForCompare_(event) {
	const normalized = normalizeOutlookEvent_(event);
	return Object.assign({}, normalized, {
		body: stripSyncMetadataFromText_(normalized.body),
	});
}

/**
 * 同期用メタデータ行を本文から取り除く。
 * @param {string} text 対象テキスト
 * @returns {string} メタデータを除去したテキスト
 */
function stripSyncMetadataFromText_(text) {
	return String(text || '')
		.split(/\r?\n/)
		.filter((line) => {
			const trimmed = line.trim();
			return (
				!/^googleSyncKey:/i.test(trimmed) && !/^outlookSyncKey:/i.test(trimmed)
			);
		})
		.join('\n')
		.trim();
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
	const locationValue = event.location
		? typeof event.location === 'string'
			? event.location
			: event.location.displayName || ''
		: '';

	return {
		subject: event.subject || '',
		body: event.body && event.body.content ? event.body.content : '',
		location: locationValue,
		start: startValue,
		end: endValue,
		isAllDay: Boolean(event.isAllDay || (event.start && event.start.date)),
		showAs: event.showAs || '',
		sensitivity: event.sensitivity || '',
		recurrence: event.recurrence || null,
	};
}

/**
 * Google イベントを比較しやすい形に正規化する。
 * @param {Object} event Google イベントオブジェクト
 * @returns {Object} 正規化されたイベント情報
 */
function normalizeGoogleEvent_(event) {
	const isAllDay = Boolean(
		(event.start && event.start.date) || (event.end && event.end.date),
	);

	return {
		subject: event.subject || '',
		description: event.description || '',
		location: event.location || '',
		start:
			(event.start && (event.start.dateTime || event.start.date)) ||
			event.startDateTime ||
			'',
		end:
			(event.end && (event.end.dateTime || event.end.date)) ||
			event.endDateTime ||
			'',
		isAllDay: isAllDay,
		showAs: event.showAs || event.transparency || '',
		sensitivity: event.sensitivity || event.visibility || '',
		recurrence: event.recurrence || null,
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
 * @param {string} outlookSyncKey Outlook のイベント同期キー
 * @returns void
 */
function updateGoogleEventWithOutlookSyncKey_(googleEvent, outlookSyncKey) {
	const description = buildGoogleDescription(
		{ description: googleEvent.description },
		outlookSyncKey,
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
 * @param {string} googleSyncKey Google のイベント同期キー
 * @returns void
 */
function updateOutlookEventWithGoogleSyncKey_(outlookEvent, googleSyncKey) {
	const bodyContent = buildOutlookDescription(
		{ description: outlookEvent.description },
		googleSyncKey,
		'',
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
