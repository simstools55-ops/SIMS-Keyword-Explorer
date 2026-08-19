# SIMS Keyword Explorer v0.1.4

## Replace
- Code.gs

## Main changes
- Adds Search Persona Profile generated from aggregated GSC query behavior.
- Treats personas as search-need clusters, not personal identification.
- Article Master is now required before internal Blue Ocean discovery.
- Adds OWNED_QUERY Gate: queries strongly owned by existing articles are excluded.
- Adds novelty/intent-drift gate before a query can become a new-article candidate.
- Keeps the verified Windows-style Evidence picker/import flow.
- Keeps a maximum of 10 internal Blue Ocean candidates.

## New visible sheet
- 検索ペルソナ

## Test
1. Replace Code.gs and reload the spreadsheet.
2. Ensure Article Master has article data.
3. Run "3. 新しいキーワード候補を探す".
4. Confirm the completion dialog reports persona groups, OWNED_QUERY exclusions, novelty exclusions, and candidate count.
5. Open "4. 検索ペルソナを確認する".
6. Confirm existing ace keywords no longer dominate the new-article candidates.
