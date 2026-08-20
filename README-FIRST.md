# SIMS Keyword Explorer v0.5.2

## 修正内容
v0.5.1で追加した「Creator公開結果を登録」ダイアログを開くと、
`ReferenceError: skeHtml_ is not defined`
が発生する不具合を修正しました。

原因:
- ダイアログHTMLへ候補名・Candidate IDを安全に埋め込むため `skeHtml_()` を呼び出していた
- しかし、そのHTMLエスケープ関数本体がCode.gsへ追加されていなかった

修正:
- `skeHtml_()` を追加
- 製品バージョンを v0.5.2 へ更新

## Apps Scriptで変更するファイル
- Code.gs : 置換
- その他 : 変更なし

## 実用試験の再開
1. Code.gsをv0.5.2へ全置換
2. 保存してスプレッドシートを再読み込み
3. SIMS Keyword Explorer → 追加の操作 → Creator公開結果を登録
4. 以下を登録
   - Candidate ID: SKE-20260820-2DC528CF
   - 公開ArticleID: A000430
   - 公開URL: https://tonbos55.hatenablog.com/entry/2026/08/20/155116

## 推奨コミットメッセージ
fix: release SKE v0.5.2 publication dialog html escape helper
