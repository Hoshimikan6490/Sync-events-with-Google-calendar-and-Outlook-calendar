# Google ↔ Outlook カレンダー同期

Google Apps Script で Google カレンダーと Outlook カレンダーを双方向同期するプロジェクトです。仕様の詳細は [仕様書.md](仕様書.md) を参照してください。

## ファイル

- [code.gs](code.gs) - 同期の入口、差分判定、集計ログ
- [googleCalendarManager.gs](googleCalendarManager.gs) - Google カレンダーの取得・作成・更新・削除
- [outlookCalendarManager.gs](outlookCalendarManager.gs) - Outlook カレンダーの取得・作成・更新・削除、ICS 展開
- [outlookOauth2.gs](outlookOauth2.gs) - Outlook OAuth2 認証
- [appsscript.json](appsscript.json) - GAS マニフェスト

## 必要な Script Properties

- `CLIENT_ID`
- `TENANT_ID`
- `OUTLOOK_CALENDAR_ID`
- `OUTLOOK_ICS_URL`
- `REFRESH_TOKEN`
- `ACCESS_TOKEN`

`OUTLOOK_CALENDAR_ID` を設定した場合はそのカレンダーを使い、未設定なら Outlook の既定カレンダーを使います。`OUTLOOK_ICS_URL` は取得できる場合に優先して使われます。

## 認証

1. `setup()` を実行して認証 URL を生成します。
2. 表示された URL を開いて Microsoft アカウントで認証します。
3. 返ってきた `code` を使って `authCallback()` を実行し、トークンを Script Properties に保存します。
4. 以後は `refresh_token` による自動更新で運用します。

## 同期実行

- `syncCalendars()` を実行すると、JST 基準の「当日 00:00:00 から 1 か月後 23:59:59.999」までのイベントを双方向同期します。
- 30 分間隔で回す場合は `installThirtyMinuteTrigger()` を 1 回実行します。

## 運用メモ

- 予定の関連付けは description 内の `googleSyncKey:` / `outlookSyncKey:` で行います。
- 逆流防止のため、同期元の ID を持つイベントは元カレンダーへ戻しません。
- タイムゾーンは `Asia/Tokyo` に統一しています。

## セットアップ方法

### [ステップ1]　MicrosoftのAPIの準備

1. https://portal.azure.com/ にアクセスし、自分のMicrosoftアカウントでログインする。
2. (任意)ログイン時に、右上の自分のアカウントのところに、自分の個人アカウントで参加している別プロジェクトの名前等が表示された場合は、アカウントアイコンをクリックして、「ディレクトリの切り替え」をクリックする。切り替わった画面にて、「既定のディレクトリ」に切り替えボタンを押す。
3. 左上の三本線をクリックし、「Microsoft Entra ID」をクリックし、画面中央上の「＋追加」から「アプリを登録」をクリックする。
4. アプリケーションの登録画面において、「名前」は任意の名前、「サポートされているアカウントの種類」を「任意のEntra ID テナント+個人用Microsoftアカウント」または「個人用アカウントのみ」に設定し、したの登録ボタンを押す。なお、リダイレクトURIは後程設定するので今は不要。
5. 左側のメニューから「APIのアクセス許可」を選択し、画面中央の「+アクセス許可の追加」を押して、出てきた画面の「Microsoft Graph」を選択し、「Calendars」＞「Calendars.ReadWrite」にチェックを入れて、下の「アクセス許可を追加」ボタンを押す。
6. 左側のメニューから、「概要」を選択し、以下の２項目をメモする。
  - アプリケーション (クライアント) ID
  - ディレクトリ (テナント) ID
7. 左側のメニューから、「証明書とシークレット」を選択し、「クライアントシークレット」モードになっていることを確認したうえで「+新しいクライアントシークレット」ボタンから資格情報を作成する。説明と有効期限は任意の内容。
8. 作成したクライアントシークレットの「値」をメモする。

### [ステップ2] Outlook カレンダーのICS URLの準備

1. https://outlook.live.com/calendar/view/month にアクセスし、自分のMicrosoftアカウントでログインする。
2. 画面左上の「表示」タブに移動し、一番右にお「⚙予定表の設定」をクリックする。ただし、画面幅によっては歯車のみ表示されるため、注意。
3. 開いた設定画面の「予定表＞共有予定表」を開く。
4. 「予定表を共有する」から、Googleカレンダーと同期したいカレンダーを選択し、「全ての詳細を閲覧可能」にして公開する。
5. HTMLとICSのURLが作られるので、ICSのURLをメモする。

### [ステップ3] Outlook カレンダーIDの準備

1. https://developer.microsoft.com/en-us/graph/graph-explorer にアクセスし、自分のMicrosoftアカウントでログインする。この際に、アクセス許可が求められた場合は許可する。
2. 画面右上のリクエスト入力画面で、「GET v1.0 https://graph.microsoft.com/v1.0/me/calendars」と入力する。
3. すぐ下の「Modify Permissions」をクリックし、「Calendars.Read」の横にある「Consent」をクリックして権限を許可する。
4. 2で入力したリクエストURLの右にある「▷Run query」をクリックし、下に緑色で「OK - 200 - xxxms」などと表示されれば、OK。さらに下の「Response preview」に自分のカレンダーの一覧が出るので、nameパラメータから、予定を管理したいカレンダー(基本的にはICSのURLを作成したカレンダー)のidパラメータの内容をメモする。

### [ステップ4] Google Apps Scriptの準備

1. Google スプレッドシートを 1 つ新規作成し、そのスプレッドシートに紐づく形で Google Apps Script を開く。
2. このリポジトリの `code.gs`、`googleEventsManager.gs`、`outlookEventsManager.gs` の内容を、GAS プロジェクトにそれぞれコピーする。
3. `outlookEventsManager.gs` を開き、`OUTLOOK_CONFIG.calendarId`、`OUTLOOK_CONFIG.clientId` を自分の値に置き換える。他はそのままでよい。
4. `code.gs` の `getICSUrl()` はスプレッドシートのアクティブシート A1 を見にいくので、ステップ2でメモした ICS URL をスプレッドシートの A1 に貼り付ける。
5. まず `setup()` を実行して、ログに出力された認可 URL を開く。
6. Microsoft アカウントでサインインし、表示されたリダイレクト先 URL の `code=` 以降の値を `outlookEventsManager.gs` の `OUTLOOK_CONFIG.authCode` に貼り付ける。恐らくフルスクリーンでも4行ぐらいある長文文字列のはず。
7. `authCallback()` を実行してリフレッシュトークンを含めた認証情報を「スクリプトトークン」に保存する。
8. 以後は `syncMonthlyCalendars()` を実行すれば、Outlook -> Google -> Outlook の順で 1 か月分の同期ができる。