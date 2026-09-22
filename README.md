# 社内タイミー シフト応募

シフト応募フォーム（静的HTML）＋ Google Apps Script（GAS）バックエンドの一式。

## 構成

- `index.html` — スタッフ向けシフト応募フォーム（GitHub Pages 等で静的配信）。
  GAS Web App を `?page=apply_data&token=...` で叩いて職種・日程を取得し、
  応募は `action:'apply'` を POST する。
- `gas/` — Google Apps Script プロジェクト「社内タイミー」のソースミラー
  （Drive 上の Apps Script プロジェクトから取得）。管理画面・LINE連携・
  勤怠・給与計算などのバックエンド一式。

## 開発環境セットアップ

```bash
npm install        # clasp をインストール
npm run gas:login  # 初回のみ。ブラウザでGoogleアカウント認可
npm run gas:pull    # Apps Script側の最新コードを gas/ に取得
npm run gas:push    # gas/ の変更を Apps Script に反映
npm run gas:open    # Apps Script エディタをブラウザで開く
```

`.clasp.json` は Drive 上の Apps Script プロジェクト（ID:
`17A1-vlnDXJJMusaIxbtxiBSLBTLO7W14ahF6UBpha1Fy6VLf-clucyjm`）を指しています。

## 注意

- `index.html` は GAS の `ApplyForm.html`（iframe配信テンプレート）とは別の、
  fetch ベースの実装（実機のビューポート問題を回避するため）。GAS側の
  `ApplyForm.html` は旧方式のテンプレートとして残っている。
- `gas/` への変更を Apps Script 本番に反映する `clasp push` は本番影響が
  あるため、実行前によく確認すること。
