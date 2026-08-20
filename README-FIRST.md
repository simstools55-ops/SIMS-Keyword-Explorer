# SIMS Keyword Explorer v0.5.0

## 新機能
- GREEN / CREATOR_READY候補からCreator用ZIPを自動生成
- 「9. 処置を進める」で候補状態に応じてDoctor用/Creator用Packageへ自動ルーティング
- Creator Packageに以下を同梱
  - creator_referral.json
  - candidate_evidence.csv
  - article_master.csv
  - README-FIRST.md

## Creator Packageの役割
DoctorがGREENと確定した新記事候補を、SIMS Article Creatorへ正式に引き継ぎます。
SKE Candidate IDを保持し、公開後のSKE登録・SBMモニターへ追跡可能にします。

## Apps Scriptで変更するファイル
- Code.gs : 置換
- その他 : 変更なし

## 実運用再開
1. Code.gsをv0.5.0へ置換して保存
2. スプレッドシートを再読み込み
3. `SKE-20260820-2DC528CF`（line メッセージ編集 いつから）だけを選択
4. 「9. 処置を進める」
5. 生成されたCreator用ZIPをSIMS Article Creatorへ渡す

## 推奨コミットメッセージ
feat: release SKE v0.5.0 Creator referral package workflow
