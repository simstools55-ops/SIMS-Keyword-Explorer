# SIMS Keyword Explorer v0.3.1

## Purpose
SBMの現行「記事一覧」を、加工せずSKE Article Masterへ取り込めるようにします。

## Main changes
- Added: 追加の操作 → SBM記事一覧からArticle Masterを取り込む
- SBMのコピー＆ペースト（TSV）を自動解析
- 必須項目を以下の3つに変更
  - 記事URL
  - 記事タイトル / H1タイトル
  - メインクエリ
- 任意項目
  - ArticleID
  - SearchIntent
  - 状態 / 作業状態
- 不要列（記事ランク、クリック、表示回数、CTR、掲載順位、更新日など）は自動無視
- ArticleIDが無い場合はURLからSKE内部IDを生成
- 外部探索結果のカニバリ判定で関連URLをCandidate Registryへ保存

## Apps Script
Replace:
- Code.gs

Add:
- none

## Recommended workflow
1. SBM → 記事一覧を開く
2. 見出し行を含めて記事一覧をコピー
3. SKE → 追加の操作 → SBM記事一覧からArticle Masterを取り込む
4. 貼り付けて登録
5. SKE → 6. Doctor外部探索結果を取り込む
