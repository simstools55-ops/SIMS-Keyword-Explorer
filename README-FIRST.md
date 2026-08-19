# SIMS Keyword Explorer v0.2.3

## Purpose
v0.2.2 Search Audience classification final patch before External Discovery.

## Fixes
- Prevent Japanese パスワード from matching Microsoft Word (ワード).
- Explicit Windows context now wins unless a real Microsoft app name is present.
- Normalize ChatGPT error variants: wasn't / wasn’t / wasn t.
- Recognize common Claude slow/incomplete-response Japanese error text as Generative AI.

## Apps Script
Replace:
- Code.gs

Add:
- none

## Verification
Run "3. 検索オーディエンスを分析する" and confirm:
- Windows password/ZIP password queries are not classified as Microsoftサービス.
- BIOS is not classified as Apple.
- ChatGPT/Claude error variants are mostly under 生成AI.
