# google-form-invitation
Google Form で出席登録するときに使えるスクリプト.
送信をトリガーに、送信者のメールアドレスを作成済みのGoogle カレンダーイベントに招待する.

## 設定方法
```
# clasp をインストール https://github.com/google/clasp
sudo npm i -g clasp
# ドキュメントに紐づけする(すでに紐付いているスクリプトが上書きされるので注意)
clasp create --parentId <紐付けるGoogle Form, Spreadsheet のID>
```

1. カレンダーにイベントを作成
2. `clasp open` してスクリプトエディタを開く
3. スクリプトのプロパティに以下を設定
   1. `CALENDAR_ID` : イベントを作成したカレンダーのID
      1. Google カレンダーのサイドバーから、イベントを作成したカレンダーの「設定と共有」を開く
      2. カレンダーのIDをコピー
   2. `EVENT_NAME` : イベント名(部分一致)
   3. `EVENT_DATE` : イベント日付
4. フォーム送信時に実行されるようにする
   1. フォームに紐づけている場合 `createFormTrigger` を実行
   2. スプレッドシートに紐付けている場合 `createSpreadsheetTrigger` を実行