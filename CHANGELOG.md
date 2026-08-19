# CHANGELOG

## v0.1.2 - 2026-08-20
- Evidence読込をGoogle Drive URL/ファイルID手入力から、Windows風のDriveファイル選択ダイアログへ変更。
- フォルダー移動、Evidence ZIP選択、manifest情報の事前確認を追加。
- 読込後にサイト名・Query行数・次操作を同一ダイアログへ表示。

## v0.1.1
- 初回起動時にメニューが表示されない不具合を修正。
- `onOpen()` から初期化処理を分離し、メニュー生成を最優先化。
- メニューへ `.addToUi()` を追加。
- `onInstall()` と初回セットアップメニューを追加。
- 利用手順のメニュー番号を整理。

## v0.1.0
- SIMS Keyword Explorer 初版P1プロトタイプ。
- Collector Evidence ZIP読込。
- 内部GSCクエリ探索。
- Article Master任意照合。
- Candidate Registry。
- Doctor用ZIP生成。
- SKE短縮ファイル命名規則を採用。
