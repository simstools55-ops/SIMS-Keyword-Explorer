# SIMS Keyword Explorer v0.2.2

## 目的
v0.2.1実データ検証で残った誤分類を修正するPATCHです。

## 修正内容
- `ami bios` / `american megatrends bios` が Apple に入る誤判定を修正
- `password` 内の `word` を Microsoft Word と誤判定する問題を修正
- 日本語表記 `チャットGPT` / `クロード` を生成AIへ追加
- 先頭欠落した `omething went wrong...` も生成AIエラーとして吸収
- `we experienced an error...` / `rate limit exceeded` を生成AIエラーとして認識
- 大規模OTHERに多かった「重い・遅い」「ログイン履歴」「通知」「場所・復元」などの意図分類を追加
- BIOS・UEFI判定をAppleより優先

## Apps Script
置換: Code.gs
新規追加: なし

## 次の確認
「3. 検索オーディエンスを分析する」を再実行し、
1. AppleにBIOSが混ざらない
2. Windowsパスワード系がMicrosoftサービスへ誤分類されない
3. ChatGPT日本語表記が未知テーマから生成AIへ移る
ことを確認してください。
