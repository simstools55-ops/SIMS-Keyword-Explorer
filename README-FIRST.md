# SIMS Keyword Explorer v0.1.0

## この版について
P1（内部探索）プロトタイプです。SIMS Site Collector Evidence ZIPを読み込み、GSCクエリから候補を抽出し、Candidate Registryへ登録してDoctor用ZIPを生成します。

## 初回手順
1. 新しいGoogleスプレッドシートを作成します。
2. Apps Scriptへ `Code.gs` と `appsscript.json` を反映します。
3. スプレッドシートを再読み込みします。
4. `SIMS Keyword Explorer` → `1. 対象ブログ / Evidenceを読み込む` を実行します。
5. Collectorが生成したEvidence ZIPのDrive URLまたはIDを貼り付けます。
6. `2. 新しいキーワード候補を探す` を実行します。
7. `キーワード候補` でDoctorへ送りたい候補にチェックを入れ、Doctor用ZIPを生成します。

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
