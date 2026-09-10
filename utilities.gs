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
 * [点検済み] タイムゾーンを比較し、必要に応じて変換する。
 * @param {string} dateTime 変換対象の日時
 * @param {string} fromTimeZone 変換対象の日時のタイムゾーン
 * @param {string} toTimeZone 変換先のタイムゾーン
 * @returns {string} 変換後の日時
 */
function convertTimeZone(dateTime, fromTimeZone, toTimeZone) {
	if (fromTimeZone === toTimeZone) {
		// タイムゾーンが同じ場合は変換不要
		return dateTime;
	}

	const zonedDateTime = Temporal.ZonedDateTime.from(
		`${dateTime}[${fromTimeZone}]`,
	);

	return zonedDateTime.withTimeZone(toTimeZone).toPlainDateTime().toString();
}
