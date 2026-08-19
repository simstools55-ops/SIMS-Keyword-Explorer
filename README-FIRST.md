# SIMS Keyword Explorer v0.2.0

## Purpose
Search Audience Profile を導入し、GSC Evidenceから
「誰が来ているか」を個人属性ではなく検索行動として分析します。

## Major change
旧: 1クエリ -> 1検索ペルソナ
新: 対象軸 × 検索意図軸

例:
- Windows × エラー・不具合解決
- 生成AI × 意味・仕組み理解
- 動画・SNS × 設定変更・解除
- Apple × 同期・接続

## OTHER analysis
対象軸または意図軸で分類できないQueryも捨てません。
OTHERとして表示し、未知の検索オーディエンス発見材料にします。

## Replace
- Code.gs

## New files in Apps Script
- none

## Unchanged
- Evidence picker
- Evidence import format
- Article Master gate for new-article candidate discovery
- OWNED_QUERY gate
- Candidate novelty gate
- Doctor ZIP workflow

## Test sequence
1. Code.gsを置換
2. 保存
3. スプレッドシート再読み込み
4. 「3. 検索オーディエンスを分析する」
5. 「検索オーディエンス」シートを確認
