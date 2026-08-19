# SIMS Keyword Explorer v0.1.3

## Replace
- Code.gs

## Unchanged
- Spreadsheet itself
- Existing Evidence ZIP
- Existing Collector / Diagnosis / SBM

## Main changes
- Query Cluster is executed before candidate generation.
- Existing Article Gate is strengthened.
- Practical candidates are limited to 10.
- Doctor candidates are limited to 3.
- Regeneration replaces only unprocessed P1 candidates for the current site.
- Candidate sheet hides internal columns that are not normally needed.
- v0.1.2 Windows-style Evidence picker/import flow is preserved.

## Test
1. Replace the current Apps Script Code.gs with this Code.gs.
2. Save and reload the spreadsheet.
3. Confirm the menu shows normally.
4. Run "3. 新しいキーワード候補を探す".
5. Confirm practical candidates are 10 or fewer and Doctor candidates are 3 or fewer.
