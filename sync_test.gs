const clientId = 'YOUR_CLIENT_ID';
const clientSecret = 'YOUR_CLIENT_SECRET';
const tenantId = 'consumers'; // DO NOT CHANGE
const redirectUri =
  'https://login.microsoftonline.com/common/oauth2/nativeclient'; // DO NOT CHANGE
const tokenUrl =
  'https://login.microsoftonline.com/' + tenantId + '/oauth2/v2.0/token'; // DO NOT CHANGE
const authCode =
  'YOUR_ACCESS_CODE'; //write CODE using setup function

function setup() {
  var codeVerifier = generateCodeVerifier();
  var codeChallenge = generateCodeChallenge(codeVerifier);

  PropertiesService.getScriptProperties().setProperty(
    'oauth_code_verifier',
    codeVerifier,
  );

  var url =
    'https://login.microsoftonline.com/' +
    tenantId +
    '/oauth2/v2.0/authorize' +
    '?client_id=' +
    clientId +
    '&response_type=code' +
    '&redirect_uri=' +
    encodeURIComponent(redirectUri) +
    '&scope=' +
    encodeURIComponent(
      'offline_access https://graph.microsoft.com/Calendars.ReadWrite',
    ) +
    '&response_mode=query' +
    '&code_challenge=' +
    encodeURIComponent(codeChallenge) +
    '&code_challenge_method=S256';

  Logger.log('このURLを開いて認証して👇');
  Logger.log(url);
}

function authCallback() {
  var codeVerifier = PropertiesService.getScriptProperties().getProperty(
    'oauth_code_verifier',
  );

  if (!codeVerifier) {
    throw new Error(
      'oauth_code_verifier がありません。先に setup() を実行して認可URLを再生成してください。',
    );
  }

  var payload = {
    client_id: clientId,
    code: authCode,
    redirect_uri: redirectUri,
    grant_type: 'authorization_code',
    code_verifier: codeVerifier,
  };

  var options = {
    method: 'post',
    payload: payload,
    muteHttpExceptions: true,
  };

  var res = UrlFetchApp.fetch(tokenUrl, options);
  var body = res.getContentText();
  var status = res.getResponseCode();

  if (status >= 400) {
    throw new Error('Token exchange failed (' + status + '): ' + body);
  }

  var data = JSON.parse(body);

  Logger.log(data);

  // refresh_token は毎回返るとは限らないため、存在時のみ更新する。
  if (data.refresh_token) {
    PropertiesService.getScriptProperties().setProperty(
      'refresh_token',
      data.refresh_token,
    );
  }

  PropertiesService.getScriptProperties().setProperty(
    'access_token',
    data.access_token,
  );
}

function generateCodeVerifier() {
  var bytes =
    Utilities.getUuid().replace(/-/g, '') +
    Utilities.getUuid().replace(/-/g, '');
  return bytes.slice(0, 64);
}

function generateCodeChallenge(codeVerifier) {
  var digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    codeVerifier,
    Utilities.Charset.UTF_8,
  );
  return base64UrlEncode(digest);
}

function base64UrlEncode(bytes) {
  return Utilities.base64Encode(bytes)
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/, '');
}

function createTestEvent() {
  var token =
    PropertiesService.getScriptProperties().getProperty('access_token');

  var url = 'https://graph.microsoft.com/v1.0/me/events';

  var payload = {
    subject: 'テスト成功！',
    start: {
      dateTime: '2026-04-05T10:00:00',
      timeZone: 'Asia/Tokyo',
    },
    end: {
      dateTime: '2026-04-05T11:00:00',
      timeZone: 'Asia/Tokyo',
    },
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + token,
    },
    payload: JSON.stringify(payload),
  };

  var res = UrlFetchApp.fetch(url, options);
  Logger.log(res.getContentText());
}

function refreshAccessToken() {
  var url = 'https://login.microsoftonline.com/consumers/oauth2/v2.0/token';

  var refreshToken =
    PropertiesService.getScriptProperties().getProperty('refresh_token');

  if (!refreshToken) {
    throw new Error(
      'refresh_token がありません。先に authCallback() を実行してトークンを保存してください。',
    );
  }

  var payload = {
    client_id: clientId,
    refresh_token: refreshToken,
    grant_type: 'refresh_token',
  };

  var options = {
    method: 'post',
    payload: payload,
    muteHttpExceptions: true,
  };

  var res = UrlFetchApp.fetch(url, options);
  var body = res.getContentText();
  var status = res.getResponseCode();

  if (status >= 400) {
    throw new Error('Refresh token failed (' + status + '): ' + body);
  }

  var data = JSON.parse(body);

  // トークン更新保存
  PropertiesService.getScriptProperties().setProperty(
    'access_token',
    data.access_token,
  );

  if (data.refresh_token) {
    PropertiesService.getScriptProperties().setProperty(
      'refresh_token',
      data.refresh_token,
    );
  }

  return data.access_token;
}

function createEventAuto() {
  var token = refreshAccessToken();

  var url = 'https://graph.microsoft.com/v1.0/me/events';

  var payload = {
    subject: '完全自動イベント',
    start: {
      dateTime: new Date().toISOString(),
      timeZone: 'Asia/Tokyo',
    },
    end: {
      dateTime: new Date(Date.now() + 60 * 60 * 1000).toISOString(),
      timeZone: 'Asia/Tokyo',
    },
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + token,
    },
    payload: JSON.stringify(payload),
  };

  UrlFetchApp.fetch(url, options);
}
