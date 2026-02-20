# Deep Scrub Maintenance Note (2026-02-20)

## What was done
- Rewrote repository history for `main` in a dedicated rewrite clone.
- Removed sync-only paths from full history:
  - `docs/upstream-sync/`
  - `scripts/upstream_sync.py`
- Applied global text replacement rules to remove legacy upstream branding tokens.
- Added hard guardrails in current `main`:
  - `.github/workflows/brand-hygiene.yml` (fails CI on forbidden tokens)
  - `.gitignore` rules for local-only sync artifacts

## Backup and safety
- Mirror backup and bundle backup were created before rewrite.
- A local snapshot branch was created before cleanup:
  - `wip/pre-deep-scrub-20260220`

## Push result
- `main` was force-updated successfully.
- Tags `v0.4.2` and `v0.4.3` were force-updated successfully.
- Tags `v0.4.4`, `v0.4.5`, `v0.4.6` were rejected by GitHub protected-ref rules and need admin action to rewrite.

## Team action
- Recommended: fresh clone after history rewrite.
- If protected tags must also be scrubbed, temporarily allow tag updates for those refs, then re-push rewritten tags.
