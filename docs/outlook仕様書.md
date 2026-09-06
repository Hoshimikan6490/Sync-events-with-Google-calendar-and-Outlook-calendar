# Outlook 側実装仕様

## 概要

このファイルは、Outlook カレンダー側の取得・認証・変換ロジックを実装に合わせて整理したものです。同期全体の流れは [動作仕様書](動作仕様書.md) を参照してください。

同期の起点は `syncCalendars()` で、Outlook 側の取得・正規化・変換はこの仕様に従います。

## 認証関連

### スクリプトプロパティーに保存される内容

現在の実装で参照する主なスクリプトプロパティーは次の通りです。

- `CLIENT_ID`
- `TENANT_ID`
- `OUTLOOK_CALENDAR_ID`
- `AUTH_CODE`
- `OUTLOOK_CODE_VERIFIER`
- `REFRESH_TOKEN`
- `ACCESS_TOKEN`

`OUTLOOK_CALENDAR_ID` が未設定なら既定の Outlook カレンダーを使います。

### アクセストークンの取得と更新
- `setup()`

PKCE 用の `code_verifier` を生成し、認可 URL をログへ出力します。`code_challenge` と `code_challenge_method=S256` を使います。

- `authCallback()`

保存済みの認可コードをトークンに交換し、`ACCESS_TOKEN` と `REFRESH_TOKEN` を保存します。

- `refreshAccessToken()`

`REFRESH_TOKEN` を使って `ACCESS_TOKEN` を更新します。401 が返った場合の再取得にも使われます。

- `generateCodeVerifier()` / `generateCodeChallenge()`

PKCE の `S256` 用コードを生成します。

認証トークンなどの機密情報は Logger に出力しません。

## 同期キー

Outlook 側の同期キー( syncKey )は `outlook_id:` で始めます。

- **単一イベント**: `outlook_id:<eventId>`
- **繰り返し occurrence**: `outlook_id:repeat_<recurringEventId>_<occurrenceDateTime>`
  - `eventId`: Graph API のイベントID
  - `recurringEventId`: Graph API の繰り返しイベント ID
  - `occurrenceDateTime`: occurrence の開始日時（UTC、RFC3339 形式）

Outlookのイベントは各種パラメータからこの「同期キー」を作成し、Googleのイベントを作成する際の概要欄の一番最後の行に設定します。これを設定・読み取りすることで、同一イベントを重複して作成してしまう問題を防止し、正しくイベントの比較・更新を行えるようになります。

## イベントデータの正規化

### Outlookのイベントから、本システムの内部データ形式への正規化

`normalizeOutlookEventToLocalFormat(event)`

Graph API から取得したイベントを内部表現に変換します。
対応フォーマットは、[イベント仕様の対応表](動作仕様書.md#内部正規化)を参考にしてください。

### 本システムの内部データ形式から、Outlookのイベントへの正規化

`normalizeOutlookEventFromLocalFormat(event)`

内部表現のイベントデータを、outlookに渡す形式に変換します。
対応フォーマットは、[イベント仕様の対応表](動作仕様書.md#内部正規化)を参考にしてください。

### 日時の管理に関して

## イベント管理

### イベント取得

`fetchOutlookEvents(startDate, endDate)`

Outlook 側のイベントは Microsoft Graph API の `calendarView` で取得します。

Graph API の設定は以下の通りです。

- ベースパスは `OUTLOOK_CALENDAR_ID` があれば `/me/calendars/{id}`
- 未設定なら `/me`
- `startDateTime` / `endDateTime` は UTC ベースの ISO 文字列で送る
- `orderby=start/dateTime` を使う

取得後は内部表現へ正規化し、同期ウィンドウとの重なりで再フィルタします。

### イベント作成

`createOutlookEvent(eventData)`

Microsoft Graph の `POST /events` でイベントを作成します。作成時には、自動的に `google_id:`で始まるgoogleカレンダーのイベントのsyncKeyが追加されるようにします。

### イベント編集

`updateOutlookEvent(eventId, eventData)`

Microsoft Graph の `PATCH /events/{id}` で更新します。更新時には、概要欄の一番最後に`google_id:`で始まるgoogleカレンダーのイベントのsyncKeyがgoogleのイベント情報に基づいて最新の状態に更新されるようにします。ただし、syncKeyの更新の際は、概要欄にはsyncKeyは必ず最大１つだけになるようにし、２つ以上にはならないように適切に置換処理をします。

### イベント削除

`deleteOutlookEvent(eventId)`

Microsoft Graph の `DELETE /events/{id}` で削除します。現在の同期ロジックでは削除タスクも実行し、処理順は `delete -> update -> create` です。

## 留意事項

- 同期の開始点は日始まりではなく実行時点です
- 同期ウィンドウの範囲内において、同期元のカレンダーでイベントが削除されると、同期先のカレンダーでも削除されます
- 認証トークンなどの機密情報は Logger に出力してはいけません
- このドキュメントを含むすべてのドキュメントに書かれた関数は一例であり、実装上必要な関数は任意で作成・編集・削除してください。ただし、それらの変更は原則 google と outlook の仕様書に反映させてください
- 各関数は、以下の形式で関数の解説を付けてください
```js
/**
 * 関数の解説
 * @param {引数の型} parameter_name 引数の説明 
 * @returns {返り値の型} 返り値の説明
 */
function FUNCTION_NAME(PARAMETERS) {
	return RETURNS
}
```

## 参考

- 全体の同期フローは [動作仕様書](動作仕様書.md)
- Google 側のリソース変換は [Google 仕様書](Google仕様書.md)

以上
