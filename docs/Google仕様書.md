# Google 側実装仕様

## 概要

このファイルは、Google カレンダー側の入出力と変換ロジックを実装に合わせて整理したものです。実際の同期フロー全体は [動作仕様書](動作仕様書.md) を参照してください。

## 取得

### `getGoogleEvents(startDate, endDate)`

Google の既定カレンダーからイベントを取得し、同期用の occurrence 配列に正規化して返します。

現在の実装は以下の通りです。

- `Calendar.Events.list(calendarId, ...)` を使う
- `singleEvents: true` を指定して繰り返し予定を展開済みの occurrence として取得する
- `orderBy: 'startTime'` を指定する
- `timeMin` / `timeMax` は UTC ベースの RFC3339 with offset (+09:00) 文字列で送る
- 取得後に、開始・終了が同期ウィンドウに重なるものだけを残す

### イベント ID の取得

**統一ルール**: Google イベント ID の取得方法はコード内で統一します。

Google Calendar API の `event.id` はカレンダー内で一意ですが、ConfidentialID など内部形式です。UI で見える `eid=` パラメータの方が参照値として適切なため、以下の方法で抽出します。

**実装パターン（必須）**:
```javascript
var link = event.htmlLink;
var event_id = link.split("eid=")[1];
```

- **常にこのパターンを使用** してコード内で統一する
- `event.id` を直接使わない
- 他の方法での ID 抽出は禁止

こうして取得した `event_id` を以降の同期キーや補助情報に使用します。

### 正規化

`normalizeGoogleCalendarEvent(event)` は、Google API のイベントを内部表現に変換します。

主な対応は次の通りです。

- `summary` → `subject`
- `description` → `description`
- `location` → `location`
- `start` / `end` → 正規化済みの開始・終了情報
- `transparency` → `showAs`
- `visibility` → `sensitivity`
- `recurringEventId` / `originalStartTime` / `occurrenceDate` を保持

開始・終了が欠けているイベントは同期対象外です。

## 作成・更新・削除

### `createGoogleEvent(eventData)`

Google Calendar API の `insert` でイベントを作成します。

### `updateGoogleEvent(eventId, eventData)`

Google Calendar API の `update` で既存イベントを更新します。

### `deleteGoogleEvent(eventId)`

Google Calendar API の `remove` でイベントを削除します。現在の同期ロジックでは削除タスクも実行します。

## リソース構築

### `buildGoogleCalendarResource(eventData)`

内部表現から Google API 用のリソースを組み立てます。

変換内容は次の通りです。

- `subject` → `summary`
- `description` → `description`
- `showAs` → `transparency`
- `sensitivity` → `visibility`
- `location` は文字列として送る
- `start` / `end` は Google 形式に戻して送る

終日イベントは `date`、時刻付きイベントは `dateTime` を使います。終了値が無い場合は、終日なら翌日、時刻付きなら 1 時間後のデフォルトを補完します。

## 説明文メタデータ

### `buildGoogleDescription(event, outlookId)`

Google 側の description 末尾に、Outlook 側との関連付け情報を追記します。

- 既存の `outlook_id:` は除去する
- `outlook_id:` を追記する

形式は `outlook_id:<outlookId>` です（ただし単一イベントのみ、繰り返しは個々の occurrence で管理）。

### `extractGoogleEventId(description)`

description から `google_id:` を取り出します。現在は legacy 互換の補助関数です。

## タイムゾーンと変換

Google 側は RFC3339 with offset (+09:00) で外部入出力を行いますが、内部処理はすべて UTC に正規化します。

- API クエリ (`timeMin` / `timeMax`): UTC ベースの RFC3339 with offset (+09:00) を送信
- 取得イベント: RFC3339 with offset → UTC 形式に正規化
- 送信イベント: UTC 形式 → RFC3339 with offset に変換

## 変換ヘルパー

### `mapTransparencyToShowAs(transparency)`

- `transparent` → `free`
- それ以外 → `busy`

### `mapShowAsToTransparency(showAs)`

- `free` → `transparent`
- それ以外 → `opaque`

### `mapVisibilityToSensitivity(visibility)`

- `private` → `private`
- `confidential` → `confidential`
- `public` / その他 → `normal`

### `mapSensitivityToVisibility(sensitivity)`

- `private` → `private`
- `confidential` → `confidential`
- `personal` → `default`
- それ以外 → `default`

## recurrence 関連

Google 側の recurrence を扱う補助関数もこのファイルに置かれています。

- `extractRecurrenceFromGoogleEvent(googleEvent)`
- `buildGoogleRecurrenceFromOutlook(outlookRecurrence)`
- `mapOutlookTypeToFreq()`
- `mapOutlookDayToRruleFormat()`

現在の同期フローでは、Outlook 側から Google への作成時に recurrence 情報を引き継ぐために使われます。

## 参考

- 入力の起点は [動作仕様書](動作仕様書.md)
- Outlook 側の取得・認証は [Outlook 仕様書](outlook仕様書.md)
