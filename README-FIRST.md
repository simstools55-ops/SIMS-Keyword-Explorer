# SIMS Keyword Explorer v0.5.1

## 修正内容
Creator連携v0.5.0の完走処理として、公開済み新記事をSKEへ戻す登録機能を追加しました。

- Candidate ID
- 公開ArticleID
- 公開URL

を登録すると、対象候補を `PUBLISHED` に更新し、選択チェックを解除してHomeを再集計します。

## Apps Scriptで変更するファイル
- Code.gs : 置換
- その他 : 変更なし

## 今回の実用試験3件目で入力する値
- Candidate ID: SKE-20260820-2DC528CF
- 公開ArticleID: A000430
- 公開URL: https://tonbos55.hatenablog.com/entry/2026/08/20/155116

## 操作
1. v0.5.1のCode.gsへ置換して保存
2. スプレッドシートを再読み込み
3. SIMS Keyword Explorer → 追加の操作 → Creator公開結果を登録
4. 上記3項目を入力して登録
5. Homeで Creator紹介可能 0 / 公開済み 1 への更新を確認

## 推奨コミットメッセージ
fix: complete SKE v0.5.1 Creator publication return workflow
