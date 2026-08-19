# SIMS Keyword Explorer v0.3.3

## Critical fix: false cannibalization match

The first External Discovery roundtrip mapped:
- Candidate: LINEラボ 新機能 使い方
to:
- Existing article: Claude（クロード）の使い方・最新機能

Cause:
Generic intent words such as 「新機能」「使い方」 were counted as strong title overlap,
even when the topic/entity 「LINEラボ」 did not match.

## Fix
- Added topic/entity anchor gate to Article Master ownership matching.
- Generic intent tokens no longer establish ownership by themselves.
- Existing article must match at least one non-generic topic anchor in title or main query.
- GSC observed URL cannot override missing topic anchor for external new-topic ownership.

## Apps Script
Replace:
- Code.gs

Add:
- none

## Retest
Import the same Doctor External Discovery result again.

Expected for LINEラボ:
- It must NOT map to the Claude article.
- If no LINEラボ/LINE Lab owner exists in Article Master:
  - Doctor candidate: 1
  - Writer redirect: 0
- Doctor rejected/deprioritized: 2
