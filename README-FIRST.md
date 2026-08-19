# SIMS Keyword Explorer v0.3.2

## Fix
v0.3.1 Doctor external result import could show:
- 外部候補 0
- Writer振替 0
- Doctor候補 0

when the same Candidate ID already existed in Candidate Registry.

The previous code silently skipped duplicate Candidate IDs.

## v0.3.2 behavior
- New candidate -> add
- Existing same Candidate ID -> update in place
- Preserve progress fields:
  - selection
  - Doctor final verdict/confidence
  - recheck date
  - published ArticleID / URL
- Show new vs updated counts
- Show Doctor rejected/deprioritized count
- Mark external themes:
  - DOCTOR_IMPORTED
  - DOCTOR_REJECTED
- Accept SERP gap values MODERATE / STRONG in addition to MODERATE_GAP / STRONG_GAP
- Never silently return an unexplained all-zero result when Doctor actually supplied a candidate

## Apps Script
Replace:
- Code.gs

Add:
- none

## Retest
Paste the same Doctor external discovery result again.
Expected for the current LINE Lab case:
- Processed candidates: 1
- Doctor rejected: 2
- Candidate is either newly added or existing-updated
