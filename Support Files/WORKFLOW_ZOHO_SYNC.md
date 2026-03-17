# Prospecting Zoho Sync Workflow

This folder follows a single merge pipeline:

1. Load Zoho credentials into environment.
2. Run dry-run sync (no file changes).
3. Run apply sync (updates DB + dashboard data + single-file build).

## One-time credential setup

1. Get grant code in Zoho.
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
