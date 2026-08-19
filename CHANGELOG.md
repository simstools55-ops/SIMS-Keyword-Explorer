# CHANGELOG

## v0.3.1
- Added SBM article-list compatible Article Master import.
- ArticleID and SearchIntent are now optional for cannibalization checks.
- Minimum Article Master fields are URL, title/H1, and main query.
- Generates deterministic internal ArticleID when SBM ArticleID is unavailable.
- Preserves related URL when external discovery candidate overlaps an existing article.
