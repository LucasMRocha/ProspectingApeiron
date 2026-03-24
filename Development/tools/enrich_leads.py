#!/usr/bin/env python3
"""Lead enrichment script — 3 sources, zero cost.

1. CNPJ → BrasilAPI: sector (CNAE), porte, location, phone, situação
2. LinkedIn search URLs: generated from Full Name + Company
3. Hunter.io (free tier): email finder by domain + name

Usage:
    python enrich_leads.py                    # runs CNPJ + LinkedIn (no API key needed)
    python enrich_leads.py --hunter-key KEY   # also runs Hunter.io email finder
    python enrich_leads.py --dry-run          # preview without saving
"""
import argparse
import io
import json
import re
import shutil
import sys
import time
import urllib.request
import urllib.parse
from datetime import datetime
from pathlib import Path

import openpyxl

# Fix Windows console encoding
if sys.stdout and hasattr(sys.stdout, "buffer"):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
if sys.stderr and hasattr(sys.stderr, "buffer"):
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

# ── Helpers ──

def cell_text(ws, row, col):
    v = ws.cell(row, col).value
    if v is None:
        return ""
    return str(v).strip()


def clean_cnpj(raw):
    return re.sub(r"[^\d]", "", str(raw).strip())


def clean_domain(website):
    if not website:
        return ""
    d = re.sub(r"https?://", "", website.strip())
    d = re.sub(r"^www\.", "", d)
    d = re.sub(r"/.*", "", d)
    return d.lower()


def split_name(full_name):
    parts = full_name.strip().split()
    if len(parts) < 2:
        return full_name.strip(), ""
    return parts[0], parts[-1]


# ── CNPJ Enrichment via BrasilAPI ──

def fetch_cnpj_data(cnpj):
    url = f"https://brasilapi.com.br/api/cnpj/v1/{cnpj}"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            return json.loads(resp.read())
    except Exception as e:
        return {"error": str(e)}


CNPJ_FIELD_MAP = {
    "cnae_fiscal_descricao": "Account_Sector",
    "porte": "Account_Size",
    "ddd_telefone_1": "Account_Phone",
}


def enrich_cnpj_batch(ws, headers, dry_run=False):
    """Enrich company data from CNPJ via BrasilAPI."""
    col_cnpj = headers.get("Account_CNPJ")
    if not col_cnpj:
        print("  [CNPJ] Column Account_CNPJ not found, skipping.")
        return 0

    # Collect unique CNPJs that need enrichment
    col_sector = headers.get("Account_Sector")
    col_size = headers.get("Account_Size")
    col_phone = headers.get("Account_Phone")
    col_city = headers.get("Account_City_State")
    col_website = headers.get("Account_Website")

    # Group rows by CNPJ
    cnpj_rows = {}
    for row in range(2, ws.max_row + 1):
        raw = cell_text(ws, row, col_cnpj)
        cnpj = clean_cnpj(raw)
        if len(cnpj) < 11:
            continue
        if cnpj not in cnpj_rows:
            cnpj_rows[cnpj] = []
        cnpj_rows[cnpj].append(row)

    print(f"  [CNPJ] {len(cnpj_rows)} CNPJs únicos para consultar")
    updated = 0
    errors = 0

    for i, (cnpj, rows) in enumerate(cnpj_rows.items()):
        if i > 0 and i % 50 == 0:
            print(f"  [CNPJ] Progresso: {i}/{len(cnpj_rows)} ({errors} erros)")

        data = fetch_cnpj_data(cnpj)
        if "error" in data:
            errors += 1
            time.sleep(0.5)
            continue

        situacao = data.get("descricao_situacao_cadastral", "")
        cnae = data.get("cnae_fiscal_descricao", "")
        porte = data.get("porte", "")
        phone_raw = data.get("ddd_telefone_1", "")
        uf = data.get("uf", "")
        municipio = data.get("municipio", "")
        city_state = f"{municipio}/{uf}" if municipio and uf else ""

        for row in rows:
            changed = False

            # Account_Sector: only fill if empty
            if col_sector and cnae and not cell_text(ws, row, col_sector):
                if not dry_run:
                    ws.cell(row, col_sector).value = cnae
                changed = True

            # Account_Size: only fill if empty
            if col_size and porte and not cell_text(ws, row, col_size):
                if not dry_run:
                    ws.cell(row, col_size).value = porte
                changed = True

            # Account_Phone: only fill if empty
            if col_phone and phone_raw and not cell_text(ws, row, col_phone):
                phone = phone_raw.strip()
                if len(phone) >= 8:
                    if not dry_run:
                        ws.cell(row, col_phone).value = phone
                    changed = True

            # Account_City_State: only fill if empty
            if col_city and city_state and not cell_text(ws, row, col_city):
                if not dry_run:
                    ws.cell(row, col_city).value = city_state
                changed = True

            if changed:
                updated += 1

        # Rate limit: ~3 req/sec
        time.sleep(0.35)

    print(f"  [CNPJ] Concluído: {updated} campos atualizados, {errors} erros de API")
    return updated


# ── LinkedIn URL Generation ──

def enrich_linkedin_urls(ws, headers, dry_run=False):
    """Generate LinkedIn search URLs for leads without LinkedIn."""
    col_li = headers.get("LinkedIn")
    col_name = headers.get("Full Name")
    col_company = headers.get("Company")

    if not col_li or not col_name:
        print("  [LinkedIn] Columns LinkedIn/Full Name not found, skipping.")
        return 0

    updated = 0
    for row in range(2, ws.max_row + 1):
        current_li = cell_text(ws, row, col_li)
        if current_li:
            continue

        name = cell_text(ws, row, col_name)
        if not name:
            continue

        company = cell_text(ws, row, col_company) if col_company else ""
        query = f"{name} {company}".strip()
        encoded = urllib.parse.quote(query)
        url = f"https://www.linkedin.com/search/results/people/?keywords={encoded}"

        if not dry_run:
            ws.cell(row, col_li).value = url
        updated += 1

    print(f"  [LinkedIn] {updated} URLs geradas")
    return updated


# ── Hunter.io Email Finder ──

def fetch_hunter_email(domain, first_name, last_name, api_key):
    params = urllib.parse.urlencode({
        "domain": domain,
        "first_name": first_name,
        "last_name": last_name,
        "api_key": api_key,
    })
    url = f"https://api.hunter.io/v2/email-finder?{params}"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            data = json.loads(resp.read())
        result = data.get("data", {})
        email = result.get("email", "")
        score = result.get("score", 0)
        return email, score
    except Exception:
        return "", 0


def enrich_hunter_emails(ws, headers, api_key, dry_run=False):
    """Find emails via Hunter.io free tier (25 searches/month)."""
    col_email = headers.get("Email")
    col_name = headers.get("Full Name")
    col_website = headers.get("Account_Website")

    if not col_email or not col_name or not col_website:
        print("  [Hunter] Columns Email/Full Name/Account_Website not found, skipping.")
        return 0

    candidates = []
    for row in range(2, ws.max_row + 1):
        current_email = cell_text(ws, row, col_email)
        if current_email:
            continue
        name = cell_text(ws, row, col_name)
        website = cell_text(ws, row, col_website)
        if not name or not website:
            continue
        domain = clean_domain(website)
        if not domain:
            continue
        first, last = split_name(name)
        if not first or not last:
            continue
        candidates.append((row, domain, first, last))

    # Hunter free tier: 25 searches/month — limit to 25
    limit = min(25, len(candidates))
    print(f"  [Hunter] {len(candidates)} candidatos, processando {limit} (free tier limit)")

    updated = 0
    for i, (row, domain, first, last) in enumerate(candidates[:limit]):
        email, score = fetch_hunter_email(domain, first, last, api_key)
        if email and score >= 50:
            if not dry_run:
                ws.cell(row, col_email).value = email
            updated += 1
            name = cell_text(ws, row, col_name)
            print(f"    Found: {name} -> {email} (score: {score})")
        time.sleep(1)  # Rate limit

    print(f"  [Hunter] {updated} emails encontrados")
    return updated


# ── Main ──

def main():
    parser = argparse.ArgumentParser(description="Lead enrichment script")
    parser.add_argument("--hunter-key", default="", help="Hunter.io API key (free tier)")
    parser.add_argument("--dry-run", action="store_true", help="Preview without saving")
    args = parser.parse_args()

    root = Path(__file__).resolve().parents[2]
    bd_path = root / "Apeiron_BR_Gestao_Comercial.xlsx"
    if not bd_path.exists():
        print(f"ERROR: BD not found: {bd_path}")
        return 1

    # Backup
    if not args.dry_run:
        backup_dir = root / "Support Files" / "Backups"
        backup_dir.mkdir(parents=True, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = backup_dir / f"Apeiron_BR_Gestao_Comercial_backup_{ts}.xlsx"
        shutil.copy2(bd_path, backup_path)
        print(f"Backup: {backup_path.name}")

    wb = openpyxl.load_workbook(bd_path)
    ws = wb["Leads"]

    # Build column index
    headers = {}
    for col in range(1, ws.max_column + 1):
        h = cell_text(ws, 1, col)
        if h:
            headers[h] = col

    print(f"\nBD: {ws.max_row - 1} leads | Modo: {'DRY RUN' if args.dry_run else 'LIVE'}\n")

    total = 0

    # 1. CNPJ enrichment
    print("═══ CNPJ → BrasilAPI ═══")
    total += enrich_cnpj_batch(ws, headers, args.dry_run)

    # 2. LinkedIn URLs
    print("\n═══ LinkedIn Search URLs ═══")
    total += enrich_linkedin_urls(ws, headers, args.dry_run)

    # 3. Hunter.io
    if args.hunter_key:
        print("\n═══ Hunter.io Email Finder ═══")
        total += enrich_hunter_emails(ws, headers, args.hunter_key, args.dry_run)
    else:
        print("\n═══ Hunter.io ═══")
        print("  Skipped (use --hunter-key KEY to enable)")

    # Save
    if not args.dry_run and total > 0:
        wb.save(bd_path)
        print(f"\n✓ Salvo: {total} campos atualizados no BD")
    elif args.dry_run:
        print(f"\n[DRY RUN] {total} campos seriam atualizados")
    else:
        print("\nNenhuma atualização necessária")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
