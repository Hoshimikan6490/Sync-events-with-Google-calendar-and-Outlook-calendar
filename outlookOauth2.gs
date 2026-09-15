// ===== Outlook OAuth2 設定 =====
const OUTLOOK_AUTH_BASE_URL = 'https://login.microsoftonline.com';
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

// ===== 認証フロー =====
/**
 * [点検済み] 認可コードを取得するための URL を生成し、ログに出力する。
 * @param void
 * @returns void
 */
function _setup() {
	const clientId = _getScriptPropertyValue(OUTLOOK_PROPERTY_KEYS.clientId);
	if (!clientId) {
		throw new Error('Script Properties に CLIENT_ID が設定されていません。');
	}

	const codeVerifier = generateCodeVerifier();
	const codeChallenge = generateCodeChallenge(codeVerifier);

	_setScriptPropertyValue(OUTLOOK_PROPERTY_KEYS.codeVerifier, codeVerifier);

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
 * [点検済み] PKCE 用の code_verifier を生成する。
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
 * [点検済み] code_verifier から SHA-256 を用いて code_challenge を生成する。
 * @param {string} codeVerifier PKCE の code_verifier
 * @returns {string} base64url 形式の code_challenge
 */
function generateCodeChallenge(codeVerifier) {
	const digest = Utilities.computeDigest(
		Utilities.DigestAlgorithm.SHA_256,
		codeVerifier,
		Utilities.Charset.UTF_8,
	);
	return _base64UrlEncode(digest);
}

/**
 * [点検済み] Outlook の認可エンドポイント URL を返す。
 * @param void
 * @returns {string} 認可エンドポイント URL
 */
function getOutlookAuthAuthorizeUrl() {
	const tenantId =
		_getScriptPropertyValue(OUTLOOK_PROPERTY_KEYS.tenantId) || 'consumers';
	return `${OUTLOOK_AUTH_BASE_URL}/${tenantId}/oauth2/v2.0/authorize`;
}

/**
 * [点検済み] 認可コードをトークンに交換して ScriptProperties に保存する。
 * @param void
 * @returns void
 */
function _authenticate() {
	const codeVerifier = _getScriptPropertyValue(
		OUTLOOK_PROPERTY_KEYS.codeVerifier,
	);

	if (!codeVerifier) {
		throw new Error(
			'oauth_code_verifier がありません。先に setup() を実行して認可URLを再生成してください。',
		);
	}

	const clientId = _getScriptPropertyValue(OUTLOOK_PROPERTY_KEYS.clientId);
	if (!clientId) {
		throw new Error('Script Properties に CLIENT_ID が設定されていません。');
	}

	const options = {
		method: 'post',
		payload: {
			client_id: clientId,
			code: decodeURIComponent(
				_getScriptPropertyValue(OUTLOOK_PROPERTY_KEYS.authCode),
			),
			redirect_uri: OUTLOOK_CONFIG.redirectUri,
			grant_type: 'authorization_code',
			code_verifier: codeVerifier,
		},
		muteHttpExceptions: true,
	};

	const res = UrlFetchApp.fetch(getOutlookAuthTokenUrl(), options);
	const body = res.getContentText();
	const status = res.getResponseCode();

	if (status >= 400) {
		throw new Error(`Token exchange failed (${status}): ${body}`);
	}

	const data = JSON.parse(body);

	Logger.log('認証成功');

	// refresh_token は毎回返るとは限らないため、存在時のみ更新する。
	if (data.refresh_token) {
		_setScriptPropertyValue(
			OUTLOOK_PROPERTY_KEYS.refreshToken,
			data.refresh_token,
		);
	}

	_setScriptPropertyValue(OUTLOOK_PROPERTY_KEYS.accessToken, data.access_token);
	Logger.log('アクセストークンを保存しました。');
}

/**
 * [点検済み] Outlook のトークンエンドポイント URL を返す。
 * @param void
 * @returns {string} トークンエンドポイント URL
 */
function getOutlookAuthTokenUrl() {
	const tenantId =
		_getScriptPropertyValue(OUTLOOK_PROPERTY_KEYS.tenantId) || 'consumers';
	return `${OUTLOOK_AUTH_BASE_URL}/${tenantId}/oauth2/v2.0/token`;
}

/**
 * [点検済み] 保存済みの refresh_token を使いアクセス トークンを更新して返す。
 * @param void
 * @returns {string} 更新後の access_token
 */
function refreshAccessToken() {
	const url = getOutlookAuthTokenUrl();

	const refreshToken = _getScriptPropertyValue(
		OUTLOOK_PROPERTY_KEYS.refreshToken,
	);

	if (!refreshToken) {
		throw new Error(
			'refresh_token がありません。先に _authenticate() を実行してトークンを保存してください。',
		);
	}

	const clientId = _getScriptPropertyValue(OUTLOOK_PROPERTY_KEYS.clientId);
	if (!clientId) {
		throw new Error('Script Properties に CLIENT_ID が設定されていません。');
	}

	const options = {
		method: 'post',
		payload: {
			client_id: clientId,
			refresh_token: refreshToken,
			grant_type: 'refresh_token',
		},
		muteHttpExceptions: true,
	};

	const res = UrlFetchApp.fetch(url, options);
	const body = res.getContentText();
	const status = res.getResponseCode();

	if (status >= 400) {
		throw new Error(`Refresh token failed (${status}): ${body}`);
	}

	const data = JSON.parse(body);

	Logger.log('アクセストークンを更新しました。');
	// トークン更新保存
	_setScriptPropertyValue(OUTLOOK_PROPERTY_KEYS.accessToken, data.access_token);

	if (data.refresh_token) {
		_setScriptPropertyValue(
			OUTLOOK_PROPERTY_KEYS.refreshToken,
			data.refresh_token,
		);
	}

	return data.access_token;
}

/**
 * [点検済み] スクリプトプロパティからアクセストークンを取得し、無ければリフレッシュを試みる。
 * @param void
 * @returns {string} 利用可能なアクセストークン
 */
function _getAccessToken() {
	const accessToken = _getScriptPropertyValue(
		OUTLOOK_PROPERTY_KEYS.accessToken,
	);
	if (accessToken) {
		return accessToken;
	}

	const refreshToken = _getScriptPropertyValue(
		OUTLOOK_PROPERTY_KEYS.refreshToken,
	);
	if (!refreshToken) {
		throw new Error(
			'ACCESS_TOKEN または REFRESH_TOKEN が設定されていません。outlookOauth2.gs の setup() と _authenticate() を実行してください。',
		);
	}

	return refreshAccessToken();
}
