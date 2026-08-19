# SIMS Keyword Explorer v0.3.2 Release Notes

This patch fixes the zero-result ambiguity found in the first real External Discovery roundtrip.

A repeated Doctor result is now safe to import.
SKE updates the same candidate instead of silently skipping it, and explicitly tracks
Doctor-rejected exploration themes.
