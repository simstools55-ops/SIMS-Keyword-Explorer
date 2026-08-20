# SIMS Keyword Explorer v0.4.0

## Major feature
Doctor final diagnosis return path.

External Discovery and final Doctor diagnosis are now separate operations.

## Menu
6. Doctor外部探索結果を取り込む
   - accepts SIMS_DOCTOR_EXTERNAL_DISCOVERY_RESULT_V1

7. Doctor診断結果を取り込む
   - accepts SIMS_DOCTOR_CREATOR_SERP_RESULT_V1
   - GREEN / YELLOW / BLOCK

## Final diagnosis mapping

### GREEN
- Doctor判定: GREEN
- 状態: CREATOR_READY
- Engine判定: CREATOR_REFERRAL
- 次回確認日: blank

### YELLOW
- Doctor判定: YELLOW
- 状態: EARLY_OPPORTUNITY
- Engine判定: MONITOR
- Doctor確信度: stored
- 次回確認日: stored from treatment_plan.recheck_date or review_after_days

### BLOCK
- default:
  - 状態: BLOCKED
  - Engine判定: BLOCK
- if Doctor explicitly redirects to Writer:
  - 状態: WRITER_REDIRECT
  - Engine判定: WRITER_REDIRECT

## Current LINE Lab expected result
- Doctor判定: YELLOW
- Doctor確信度: 70
- 次回確認日: 2026-09-15
- 状態: EARLY_OPPORTUNITY
- Engine判定: MONITOR

## Apps Script
Replace:
- Code.gs

Add:
- none
