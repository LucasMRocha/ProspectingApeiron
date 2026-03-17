# Prospecting Local Sync Workflow

This folder follows a single merge pipeline (local data only):

1. Confirm source file(s) are present in `Zoho data/` or local datasets.
2. Run dry-run sync (no file changes).
3. Run apply sync (updates DB + dashboard data + single-file build).

## One-time setup (local)

1. Ensure `Zoho data/` CSV files are available (if migrating from historical data).
2. Run:
```powershell
.\tools\get-zoho-token.ps1 -Code "1000.xxxxx.xxxxx"
```

3. This creates `tools\zoho-env.ps1` locally.

## Daily sync

Dry-run first:

```bat
SYNC_ZOHO_DRY_RUN.bat
```

If summary looks correct, apply:

```bat
SYNC_ZOHO_APPLY.bat
```

## Canonical script

- `tools\sync-zoho-to-db.py`

## Legacy scripts

- `tools\archive_legacy\`
