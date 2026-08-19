# SIMS Keyword Explorer v0.1.1

## 修正内容
- 初回起動時のメニュー表示不具合を修正。
- `onOpen()` では初期化処理を実行せず、メニューを最優先で表示するよう変更。
- メニューへ `.addToUi()` を明示し、スプレッドシート上へ確実に追加するよう修正。
- 初回利用者向けに `1. 初期設定 / 対象ブログを準備` を追加。
- Evidence読込以降のメニュー番号を整理。
- `onInstall()` を追加し、インストール時も同じメニュー生成経路を利用。

## 変更ファイル
- Code.gs
- VERSION
- README-FIRST.md
- CHANGELOG.md
- RELEASE-NOTES-v0.1.1.md
