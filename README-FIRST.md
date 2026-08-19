# SIMS Keyword Explorer v0.3.0

## Major feature: External Discovery

Search Audience Profileを使い、
「このブログに来る検索者が次に困りそうな新しい外部変化」を
SIMS DoctorのWeb検索へ渡すルートを実装しました。

## Main workflow
1. 初期設定
2. Evidenceを読み込む
3. 検索オーディエンスを分析する
4. 外部探索テーマを作る
5. 外部探索Packageを作る
6. Doctor外部探索結果を取り込む
7. 候補を確認する
8. 処置を進める

## External Discovery Package
Example:
SKE-ガジェット探検記-外部探索-YYYYMMDD-HHmm.zip

Contents:
- external_discovery_request.json
- search_audience_profile.csv
- article_master.csv (when available)
- README-FIRST.md

## Doctor result import
Doctorの回答全文またはJSONを貼り付け可能。
Required format:
SIMS_DOCTOR_EXTERNAL_DISCOVERY_RESULT_V1

SKE側でArticle Masterと照合し、
- existing owner strong -> WRITER_REDIRECT
- no strong owner -> DOCTOR_REVIEW
へ振り分けます。

## Apps Script
Replace:
- Code.gs

Add:
- none

New spreadsheet sheets are created automatically:
- 外部探索
- _SKE_EXTERNAL_RESULTS
