# SIMS Keyword Explorer v0.1.2

## この版について
P1（内部探索）プロトタイプです。SIMS Site Collector Evidence ZIPを読み込み、GSCクエリから候補を抽出し、Candidate Registryへ登録してDoctor用ZIPを生成します。

## 初回手順
1. 新しいGoogleスプレッドシートを作成します。
2. Apps Scriptへ `Code.gs` と `appsscript.json` を反映します。
3. Apps Scriptを保存し、スプレッドシートを再読み込みします。
4. 上部に `SIMS Keyword Explorer` メニューが表示されたら、`1. 初期設定 / 対象ブログを準備` を実行します。
5. `2. Evidenceを読み込む` を実行し、Windows風のGoogle Drive選択画面からSIMS Site Collectorが生成したEvidence ZIPを選びます。
6. `3. 新しいキーワード候補を探す` を実行します。
7. `4. 候補を確認する` で内容を確認します。
8. Doctorへ送りたい候補にチェックを入れ、`5. 処置を進める` または追加操作からDoctor用ZIPを生成します。

## Evidenceの選択
Google Drive URLやファイルIDの手入力は不要です。フォルダーを移動し、対象のEvidence ZIPをクリックするとサイト名・URL・作成日時・収集期間を確認してから読み込めます。

## メニューが表示されない場合
`onOpen()` はメニュー生成だけを行う設計です。コード保存後にスプレッドシートを再読み込みしてください。初期設定でエラーが起きても、次回再読み込み時にメニュー自体は表示できる構造です。

## Article Master（任意）
既存記事との重複判定精度を高めるため、`追加の操作` → `Article Masterの使い方` から `_SKE_ARTICLE_MASTER` を開き、SBM等の既存記事データを貼り付けられます。

## P1で未実装
- 外部Web Discovery
- SERP Gap自動判定
- Demand Maturityの外部シグナル判定
- Article Lifespanの自動判定
- Doctor/Creatorの完全自動連携

## 生成ZIP名
利用者が見て用途を識別できるよう、Doctor用は次の形式です。

`SKE-[ブログ名]-Doctor用-[CandidateID]-[日時].zip`
