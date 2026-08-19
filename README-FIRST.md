# SIMS Keyword Explorer v0.1.5

## Replace
- Code.gs

## Fix
v0.1.4 checked Article Master before Search Persona analysis.
That prevented persona analysis even though persona analysis only needs Collector Evidence.

## Workflow
1. 初期設定
2. Evidenceを読み込む
3. 検索ペルソナを分析する（Article Master不要）
4. 新しいキーワード候補を探す（Article Master必須）
5. 検索ペルソナを確認する
6. 候補を確認する
7. 処置を進める

## Preserved
- Windows-style Evidence picker
- Evidence import
- Persona clustering logic
- OWNED_QUERY gate
- Candidate novelty gate
