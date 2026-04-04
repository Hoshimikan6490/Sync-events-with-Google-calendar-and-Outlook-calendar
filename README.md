# outlookカレンダーにGASからイベントを登録する方法
## 概要
Google Calendar と Outlook Calendar の 1 か月分同期用の Google Apps Script を含む。

現在の同期は次の順で動作する。

1. Outlook の ICS を取得して Google Calendar に反映する
2. Google Calendar の中で Outlook 由来ではないイベントを Outlook に作成する

主な実行関数は以下の 3 つ。

- `syncOutlookToGoogle()`
- `syncGoogleToOutlook()`
- `syncMonthlyCalendars()`

Outlook から取り込んだイベントは Google 側で `outlook_id` を持つため、Google → Outlook の再作成対象から除外される。

Outlook へのイベント作成では、スクリプト先頭の `outlookCalendarId` を設定すると、既定の予定表ではなく指定したカレンダーに作成できる。`createEventToOutlook()` に `calendarId` を渡した場合は、それが優先される。

```javascript
const outlookCalendarId = 'AAMkAGI2T...AAAAAA==';

createEventToOutlook({
  subject: '会議',
  start: new Date('2026-04-10T10:00:00'),
  end: new Date('2026-04-10T11:00:00'),
});
```

## 注意事項
- あんまり回数が多いとレートリミットに引っかかるので、程よい頻度で
- こちらは、外部ライブラリを使用しない場合のコードになります

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
