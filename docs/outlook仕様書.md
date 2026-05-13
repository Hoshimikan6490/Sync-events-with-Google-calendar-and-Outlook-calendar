# Outlook 側実装仕様

## 概要

このファイルは、Outlook カレンダー側の取得・認証・変換ロジックを実装に合わせて整理したものです。同期全体の流れは [動作仕様書](動作仕様書.md) を参照してください。

## 取得経路

### `fetchOutlookEvents(startDate, endDate)`

Outlook 側のイベントは Microsoft Graph API の `calendarView` で取得します。

Graph API の設定は以下の通りです。

- ベースパスは `OUTLOOK_CALENDAR_ID` があれば `/me/calendars/{id}`
- 未設定なら `/me`
- `startDateTime` / `endDateTime` は UTC ベースの ISO 文字列で送る
- `orderby=start/dateTime` を使う

取得後は内部表現へ正規化し、同期ウィンドウとの重なりで再フィルタします。

## 正規化

### `normalizeOutlookCalendarEvent(event)`

Graph API のイベントを内部表現に変換します。

主な対応は次の通りです。

- `subject` → `subject`
- `body.content` → `description`
- `location.displayName` → `location`
- `start` / `end` → そのまま保持
- `showAs` / `sensitivity` を保持
- `uid` は Graph の `id` を使う
- `occurrenceDate` を開始時刻から生成する

### `normalizeOutlookCalendarDateTime()`

Outlook から来た日時を Google 側へ渡しやすい形に整えます。

## イベント ID の取得

**統一ルール**: Outlook イベント ID の取得方法はコード内で統一します。

**Graph API から取得する場合（必須パターン）**:
```javascript
// Graph API レスポンスの id フィールドを使用
var id = event.id;
```

- **常にこのパターンを使用** してコード内で統一する
- 他の方法での ID 抽出は禁止

これらの ID は以降の同期キーや補助情報（`outlook_id:` 記載）に使用されます。

## 作成・更新・削除

### `createOutlookEvent(eventData)`

Microsoft Graph の `POST /events` でイベントを作成します。

### `updateOutlookEvent(eventId, eventData)`

Microsoft Graph の `PATCH /events/{id}` で更新します。

### `deleteOutlookEvent(eventId)`

Microsoft Graph の `DELETE /events/{id}` で削除します。現在の同期ロジックでは削除タスクも実行します。

## リソース構築

### `buildOutlookCalendarResource(eventData)`

内部表現から Graph API 用のイベントリソースを組み立てます。

- `subject` を設定する
- `body.contentType` は `text`
- `location` は `displayName` 付きオブジェクトにする
- `start` / `end` は `dateTime` と `timeZone` を持つ形にする
- `isAllDay` を設定する
- `showAs` と `sensitivity` を設定する

### `buildOutlookDescription(event, googleId)`

Outlook 側の description に Google 側との関連付け情報を追記します。

- 既存の `google_id:` / `Repeat:` を除去する
- `google_id:` を追記する

## OAuth2

### `setup()`

PKCE 用の `code_verifier` を生成し、認可 URL をログへ出力します。

### `authCallback()`

保存済みの認可コードをトークンに交換し、`ACCESS_TOKEN` と `REFRESH_TOKEN` を保存します。

### `refreshAccessToken()`

`REFRESH_TOKEN` を使って `ACCESS_TOKEN` を更新します。401 が返った場合の再取得にも使われます。

### `generateCodeVerifier()` / `generateCodeChallenge()`

PKCE の `S256` 用コードを生成します。

## Script Properties

現在の実装で参照する主なキーは次の通りです。

- `CLIENT_ID`
- `TENANT_ID`
- `OUTLOOK_CALENDAR_ID`
- `AUTH_CODE`
- `OUTLOOK_CODE_VERIFIER`
- `REFRESH_TOKEN`
- `ACCESS_TOKEN`

`OUTLOOK_CALENDAR_ID` が未設定なら既定の Outlook カレンダーを使います。

## 変換ルール

### `showAs` の扱い

- Graph API の `showAs` を内部表現の空き状況として保持する
- 同期時は Google 側の `transparency` と相互変換する

### `sensitivity` の扱い

- Graph API の `sensitivity` を内部表現に保持する
- 同期時は Google 側の `visibility` と相互変換する

### 日時の UTC 正規化

Graph API から取得したイベントの `start` / `end` は `dateTime` + `timeZone` の組み合わせなので、UTC の RFC3339 Z 形式（例: `2026-05-09T10:00:00Z`）に変換してから内部処理を行います。

## 参考

- 全体の同期フローは [動作仕様書](動作仕様書.md)
- Google 側のリソース変換は [Google 仕様書](Google仕様書.md)
