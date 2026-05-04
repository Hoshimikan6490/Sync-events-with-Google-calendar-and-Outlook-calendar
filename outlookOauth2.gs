// ===== Outlook OAuth2 設定 =====
const OUTLOOK_AUTH_BASE_URL = 'https://login.microsoftonline.com';
const OUTLOOK_GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
const OUTLOOK_GRAPH_SCOPE =
	'offline_access https://graph.microsoft.com/Calendars.ReadWrite';

const OUTLOOK_CONFIG = {
	clientId: '', // Script Properties から読み込み
	tenantId: 'consumers', //基本的には 'consumers' で問題ありませんが、必要に応じてテナントIDを指定してください
	redirectUri: 'https://login.microsoftonline.com/common/oauth2/nativeclient',
};

const OUTLOOK_PROPERTY_KEYS = {
	clientId: 'CLIENT_ID',
	tenantId: 'TENANT_ID',
	calendarId: 'OUTLOOK_CALENDAR_ID',
	authCode: 'AUTH_CODE',
	accessToken: 'ACCESS_TOKEN',
	refreshToken: 'REFRESH_TOKEN',
	codeVerifier: 'OUTLOOK_CODE_VERIFIER',
};

// ===== ヘルパー関数 =====
/**
 * スクリプトプロパティからキーの値を取得するヘルパー。
 * @param {string} key プロパティキー
 * @returns {string|null} プロパティ値または null
 */
function getScriptPropertyValue(key) {
	return PropertiesService.getScriptProperties().getProperty(key);
}

/**
 * スクリプトプロパティにキーと値を保存するヘルパー。
 * @param {string} key プロパティキー
 * @param {string} value 保存する値
 * @returns void
 */
function setScriptPropertyValue(key, value) {
	PropertiesService.getScriptProperties().setProperty(key, value);
}

/**
 * Base64 URL encode（RFC 4648 のセクション 5）
 * @param {string|Byte[]} input
 * @returns {string}
 */
/**
 * Base64 URL エンコードを行う（RFC4648 section5）。
 * @param {string|Byte[]} input 入力バイト列または文字列
 * @returns {string} base64url 形式の文字列
 */
function base64UrlEncode(input) {
	if (typeof input === 'string') {
		input = Utilities.newBlob(input).getBytes();
	}
	let base64 = Utilities.base64Encode(input);
	return base64.replace(/\+/g, '-').replace(/\//g, '_').replace(/=/g, '');
}

// ===== 認証フロー =====
/**
 * Outlook のトークンエンドポイント URL を返す。
 * @param void
 * @returns {string} トークンエンドポイント URL
 */
function getOutlookAuthTokenUrl() {
	const clientId = getScriptPropertyValue(OUTLOOK_PROPERTY_KEYS.clientId);
	const tenantId =
		getScriptPropertyValue(OUTLOOK_PROPERTY_KEYS.tenantId) || 'consumers';
	return OUTLOOK_AUTH_BASE_URL + '/' + tenantId + '/oauth2/v2.0/token';
}

/**
 * Outlook の認可エンドポイント URL を返す。
 * @param void
 * @returns {string} 認可エンドポイント URL
 */
function getOutlookAuthAuthorizeUrl() {
	const clientId = getScriptPropertyValue(OUTLOOK_PROPERTY_KEYS.clientId);
	const tenantId =
		getScriptPropertyValue(OUTLOOK_PROPERTY_KEYS.tenantId) || 'consumers';
	return OUTLOOK_AUTH_BASE_URL + '/' + tenantId + '/oauth2/v2.0/authorize';
}

/**
 * OAuth 認可 URL を生成し、PKCE 用の code_verifier を保存する。
 * @returns {void}
 */
/**
 * 認可用 URL を生成して PKCE の code_verifier を保存するセットアップ処理。
 * @param void
 * @returns void
 */
function setup() {
	const clientId = getScriptPropertyValue(OUTLOOK_PROPERTY_KEYS.clientId);
	if (!clientId) {
		throw new Error('Script Properties に CLIENT_ID が設定されていません。');
	}

	const codeVerifier = generateCodeVerifier();
	const codeChallenge = generateCodeChallenge(codeVerifier);

	setScriptPropertyValue(OUTLOOK_PROPERTY_KEYS.codeVerifier, codeVerifier);

	const url =
		getOutlookAuthAuthorizeUrl() +
		'?client_id=' +
		encodeURIComponent(clientId) +
		'&response_type=code' +
		'&redirect_uri=' +
		encodeURIComponent(OUTLOOK_CONFIG.redirectUri) +
		'&scope=' +
		encodeURIComponent(OUTLOOK_GRAPH_SCOPE) +
		'&response_mode=query' +
		'&code_challenge=' +
		encodeURIComponent(codeChallenge) +
		'&code_challenge_method=S256';

	Logger.log('このURLを開いて認証してください👇');
	Logger.log(url);
}

/**
 * 認可コードをアクセストークンと交換し、ScriptProperties に保存する。
 * 手順: setup() で表示された URL を開き、code を取得して引数に渡す。
 * @param {string} code 認可コード
 * @returns {void}
 */
/**
 * 認可コードをトークンに交換して ScriptProperties に保存する。
 * @param void
 * @returns void
 */
function authCallback() {
	const codeVerifier = getScriptPropertyValue(
		OUTLOOK_PROPERTY_KEYS.codeVerifier,
	);

	if (!codeVerifier) {
		throw new Error(
			'oauth_code_verifier がありません。先に setup() を実行して認可URLを再生成してください。',
		);
	}

	const clientId = getScriptPropertyValue(OUTLOOK_PROPERTY_KEYS.clientId);
	if (!clientId) {
		throw new Error('Script Properties に CLIENT_ID が設定されていません。');
	}

	const payload = {
		client_id: clientId,
		code: decodeURIComponent(
			getScriptPropertyValue(OUTLOOK_PROPERTY_KEYS.authCode),
		),
		redirect_uri: OUTLOOK_CONFIG.redirectUri,
		grant_type: 'authorization_code',
		code_verifier: codeVerifier,
	};

	const options = {
		method: 'post',
		payload: payload,
		muteHttpExceptions: true,
	};

	const res = UrlFetchApp.fetch(getOutlookAuthTokenUrl(), options);
	const body = res.getContentText();
	const status = res.getResponseCode();

	if (status >= 400) {
		throw new Error('Token exchange failed (' + status + '): ' + body);
	}

	const data = JSON.parse(body);

	Logger.log('認証成功');

	// refresh_token は毎回返るとは限らないため、存在時のみ更新する。
	if (data.refresh_token) {
		setScriptPropertyValue(
			OUTLOOK_PROPERTY_KEYS.refreshToken,
			data.refresh_token,
		);
	}

	setScriptPropertyValue(OUTLOOK_PROPERTY_KEYS.accessToken, data.access_token);
	Logger.log('トークンを保存しました。');
}

/**
 * PKCE 用の code_verifier を生成する。
 * @returns {string} 64 文字の code_verifier
 */
/**
 * PKCE 用の code_verifier を生成する。
 * @param void
 * @returns {string} 生成された code_verifier
 */
function generateCodeVerifier() {
	const bytes =
		Utilities.getUuid().replace(/-/g, '') +
		Utilities.getUuid().replace(/-/g, '');
	return bytes.slice(0, 64);
}

/**
 * code_verifier から code_challenge を生成する。
 * @param {string} codeVerifier PKCE の code_verifier
 * @returns {string} base64url 形式の code_challenge
 */
/**
 * code_verifier から SHA-256 を用いて code_challenge を生成する。
 * @param {string} codeVerifier PKCE の code_verifier
 * @returns {string} base64url 形式の code_challenge
 */
function generateCodeChallenge(codeVerifier) {
	const digest = Utilities.computeDigest(
		Utilities.DigestAlgorithm.SHA_256,
		codeVerifier,
		Utilities.Charset.UTF_8,
	);
	return base64UrlEncode(digest);
}

/**
 * 保存済み refresh_token を使って access_token を更新する。
 * @returns {string} 更新後の access_token
 */
/**
 * 保存済みの refresh_token を使いアクセス トークンを更新して返す。
 * @param void
 * @returns {string} 更新後の access_token
 */
function refreshAccessToken() {
	const url = getOutlookAuthTokenUrl();

	const refreshToken = getScriptPropertyValue(
		OUTLOOK_PROPERTY_KEYS.refreshToken,
	);

	if (!refreshToken) {
		throw new Error(
			'refresh_token がありません。先に authCallback() を実行してトークンを保存してください。',
		);
	}

	const clientId = getScriptPropertyValue(OUTLOOK_PROPERTY_KEYS.clientId);
	if (!clientId) {
		throw new Error('Script Properties に CLIENT_ID が設定されていません。');
	}

	const payload = {
		client_id: clientId,
		refresh_token: refreshToken,
		grant_type: 'refresh_token',
	};

	const options = {
		method: 'post',
		payload: payload,
		muteHttpExceptions: true,
	};

	const res = UrlFetchApp.fetch(url, options);
	const body = res.getContentText();
	const status = res.getResponseCode();

	if (status >= 400) {
		throw new Error('Refresh token failed (' + status + '): ' + body);
	}

	const data = JSON.parse(body);

	// トークン更新保存
	setScriptPropertyValue(OUTLOOK_PROPERTY_KEYS.accessToken, data.access_token);

	if (data.refresh_token) {
		setScriptPropertyValue(
			OUTLOOK_PROPERTY_KEYS.refreshToken,
			data.refresh_token,
		);
	}

	return data.access_token;
}
