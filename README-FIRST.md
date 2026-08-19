# SIMS Keyword Explorer v0.3.4

## Purpose
UI / record consistency patch after the first successful External Discovery roundtrip.

## Fixes
- When `既存記事判定 = VERIFIED_NO_STRONG_OWNER`:
  - `関連ArticleID` is blank
  - `関連URL` is blank
- Import result wording changed:
  - old: `既存更新`
  - new: `候補台帳更新`
  This avoids confusion with "update an existing article".

## Apps Script
Replace:
- Code.gs

Add:
- none

## Retest
Import the same LINE Lab Doctor result again.

Expected:
- 処理候補: 1件（新規 0 / 候補台帳更新 1）
- Writer振替: 0件
- Doctor候補: 1件
- Doctor見送り: 2件
- LINEラボ row:
  - 既存記事判定 = VERIFIED_NO_STRONG_OWNER
  - 関連ArticleID = blank
  - 関連URL = blank
