/**
 * [点検済み] 同期対象の日時ウィンドウ（開始と終了）を生成する。
 * @param void
 * @returns {{start:Date,end:Date}} 同期ウィンドウの開始日時と終了日時
 */
function getSyncWindow() {
	const start = new Date();
	const end = new Date(start);
	end.setMonth(end.getMonth() + LOOKBACK_MONTHS);
	end.setHours(23, 59, 59, 999);
	return { start, end };
}

/**
 * [点検済み] タイムゾーンを調整し、YYYY-MM-DDThh:mm:ss形式で返す。
 * @param {object} dateTimeAndTimeZone 変換対象の日時とタイムゾーン
 * @param {string} dateTimeAndTimeZone.dateTime 変換対象の日時（YYYY-MM-DDThh:mm:ss）
 * @param {string} dateTimeAndTimeZone.date 変換対象の日時（YYYY-MM-DD）
 * @param {string} dateTimeAndTimeZone.timeZone 変換対象の日時のタイムゾーン
 * @param {string} toTimeZone 変換先のタイムゾーン
 * @returns {string} 変換後の日時（YYYY-MM-DDThh:mm:ss）
 */
function convertTimeZone(dateTimeAndTimeZone, toTimeZone) {
	let dateTimeWithOffset =
		dateTimeAndTimeZone.dateTime || dateTimeAndTimeZone.date;
	const fromTimeZone =
		dateTimeAndTimeZone.timeZone ||
		CalendarApp.getDefaultCalendar().getTimeZone(); // googleカレンダーのデフォルトタイムゾーンを取得する

	// オフセットがない場合
	if (!/[+-]\d{2}:\d{2}$|Z$/.test(dateTimeWithOffset)) {
		/*
		 * fromTimeZoneの現地時刻として解釈する。
		 *
		 * GASのDateは文字列を実行環境のタイムゾーンで
		 * 解釈する可能性があるため、明示的にオフセットを付ける。
		 */
		const offset = getTimeZoneOffset(dateTimeWithOffset, fromTimeZone);

		dateTimeWithOffset = `${dateTimeWithOffset}${offset}`;
	}

	const date = new Date(dateTimeWithOffset);

	if (isNaN(date.getTime())) {
		throw new Error(`Invalid dateTime: ${JSON.stringify(dateTimeAndTimeZone)}`);
	}

	return Utilities.formatDate(date, toTimeZone, "yyyy-MM-dd'T'HH:mm:ss");
}

/**
 * [点検済み] 指定された日時のタイムゾーンオフセットを取得する。
 * @param {string} dateTime 変換対象の日時
 * @param {string} timeZone 変換対象の日時のタイムゾーン
 * @returns {string} +09:00 / -04:00 など
 */
function getTimeZoneOffset(dateTime, timeZone) {
	const date = new Date(`${dateTime}Z`);

	const formatted = Utilities.formatDate(date, timeZone, 'Z');

	const sign = formatted.startsWith('-') ? '-' : '+';
	const value = formatted.replace(/[+-]/, '');

	return `${sign}${value.slice(0, 2)}:${value.slice(2, 4)}`;
}
