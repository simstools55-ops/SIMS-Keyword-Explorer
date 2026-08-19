# SIMS Keyword Explorer v0.2.1

## 今回の目的
v0.2.0の実データ検証で判明した大量OTHERと誤分類を修正します。

## 主な修正
- ChatGPT/生成AIの英語エラーメッセージを生成AIへ分類
- Teams / Microsoft 365系を「Microsoftサービス」へ独立
- American Megatrends / AMI BIOSをAppleから「BIOS・UEFI」へ分離
- DNS / DHCP / VPN / Wi-Fiを「ネットワーク」へ独立
- Chrome等のGoogleサービス、制作・業務アプリを追加
- 障害・稼働状況、制限・容量を意図軸へ追加
- OTHER対象を「辞書不足候補」と「未知テーマ」に分離
- 固有サービスを広いOS/PC分類より優先

## Apps Script
置換: Code.gs
新規追加: なし

## 次の試験
1. Code.gsを置換
2. 保存・スプレッドシート再読み込み
3. 「3. 検索オーディエンスを分析する」
4. 上位15グループと「辞書不足候補」「未知テーマ」を確認
