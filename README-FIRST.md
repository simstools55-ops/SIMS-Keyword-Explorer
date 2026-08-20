# SIMS Keyword Explorer v0.4.1

## 更新内容

- Doctor最終診断の `SIMS_DOCTOR_CASE_RESULT_V2` + `SIMS_DOCTOR_CREATOR_SERP_RESULT_V1` 契約を受理できるよう修正
- Doctor判定を `diagnosis.verdict` からも取得できるよう修正
- Homeの「次の操作」を処理状況に応じて判定するよう修正
- `EARLY_OPPORTUNITY` などDoctor再診対象ではない候補をDoctor Package生成から除外
- Package内 `contract_version` と画面上の製品バージョンを v0.4.1 に統一

## 入れ替え対象

- `Code.gs` : 置換

その他のApps Scriptファイルは変更不要です。

## 推奨コミットメッセージ

`fix: release SIMS Keyword Explorer v0.4.1 doctor import and workflow routing`
