# OutlookPauseMan(アウトルック ポーズマン)

新しい Outlook 用の誤送信防止アドイン(Office Add-in / Event-based activation)。
メール送信前に宛先・件名・本文・外部ドメイン・宛先人数・禁止ワードをチェックします。

- 作成者: TATSURO HIGUCHI
- ターゲット: Microsoft 365 の新しい Outlook(Outlook on the Web / 新しい Outlook デスクトップ / Outlook on Mac)
- イベント: `OnMessageSend`(SendMode: PromptUser)

## 機能

| カテゴリ | 動作 |
|---|---|
| 宛先未入力 / 禁止ワード検出 | **エラー** — 送信ブロック |
| 件名なし / 本文なし / 外部ドメイン / 宛先多数 | **警告** — 送信前に確認ダイアログ |
| 全項目 OK | そのまま送信 |

設定はメール作成リボンの「**誤送信防止**」グループ → 「**設定**」ボタンから編集できます。設定は `Office.context.roamingSettings` に保存されるため、ユーザーのアカウントに紐づいて各端末で利用できます。

## ホスティング

GitHub Pages: <https://6nfdpbrwqt-ship-it.github.io/outlook-pause-man/>

## ローカル開発

```bash
npm install
npm run dev-server   # https://localhost:3000 で起動
npm run validate     # manifest.xml を検証
npm run build        # production ビルド (dist/ に出力、URL は本番のものに置換)
npm run deploy       # dist/ を gh-pages ブランチに push(GitHub Pages 反映)
```

## 配信(Microsoft 365 admin center)

1. `npm run deploy` で GitHub Pages に最新版を反映
2. <https://admin.microsoft.com> → 統合アプリ → カスタム アプリのアップロード
3. `dist/manifest.xml`(本番 URL に置換済み)を選択
4. ユーザー / グループを指定して展開

## ライセンス

Copyright (C) TATSURO HIGUCHI
