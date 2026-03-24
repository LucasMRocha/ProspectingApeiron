#!/usr/bin/env python3
"""Enrich Account_Summary (briefing) for leads in the BD.

Generates structured briefings from existing fields:
Account_Sector, Account_Size, Account_Revenue, Account_City_State,
Account_Website, Account_Phone, Account_CNPJ, Account_Brand.

Also appends an Apeiron fit indicator when sector/role match.
Only overwrites briefings that are empty or just the company name repeated.
Use --refresh to re-generate when the new briefing would be richer.
"""
import argparse
import re
import shutil
from datetime import datetime
from pathlib import Path

import openpyxl

APEIRON_SECTOR_RE = re.compile(
    r"oil|gas|o&g|petr[oó]leo|petroqui?m|energy|energia|mining|minera[cç][ãa]o|"
    r"siderur|chemicals?|qui?m|offshore|fpso|epc|refin|metalurgi|metals?|"
    r"celulose|papel|pulp|cimento|cement|port|terminal|utilit",
    re.IGNORECASE,
)
APEIRON_ROLE_RE = re.compile(
    r"instrument|automa[tcç]|el[eé]tric|electr|maintenance|manuten[cç]|"
    r"icss|scada|plc|clp|dcs|safety|seguran[cç]a|cybersec|ciberseg|"
    r"hmi|control|engenheiro|engineer",
    re.IGNORECASE,
)


def cell_text(ws, row, col):
    v = ws.cell(row, col).value
    if v is None:
        return ""
    return str(v).strip()


def is_poor_briefing(summary, company):
    """True if briefing is empty or just the company name repeated."""
    if not summary:
        return True
    s = summary.strip().lower()
    c = company.strip().lower()
    if not s:
        return True
    # briefing is just the company name (possibly with minor diffs)
    if s == c or s.replace(" ", "") == c.replace(" ", ""):
        return True
    # briefing is shorter than company name + 10 chars (no real info added)
    if len(s) <= len(c) + 5 and c[:20].lower() in s:
        return True
    return False


def build_briefing(company, sector, city_state, size, revenue, website, phone, cnpj, brand):
    """Build a structured briefing string from available fields."""
    parts = [company.upper()]

    info = []
    if sector:
        info.append(sector)
    if city_state:
        info.append(city_state)

    if info:
        parts[0] += " — " + ", ".join(info)

    details = []
    if size:
        details.append(f"~{size} func.")
    if revenue:
        details.append(f"Revenue: {revenue}")
    if website:
        details.append(website)
    if phone:
        details.append(phone)
    if cnpj:
        details.append(f"CNPJ: {cnpj}")
    if brand and brand.lower() != company.lower():
        details.append(f"Brand: {brand}")

    if details:
        parts.append(" | ".join(details))

    return " | ".join(parts) if len(parts) > 1 else parts[0]


def apeiron_fit_label(sector, roles):
    """Return Apeiron fit string if sector or roles match."""
    fits = []
    if sector and APEIRON_SECTOR_RE.search(sector):
        fits.append("sector match")
    if any(APEIRON_ROLE_RE.search(r) for r in roles if r):
        fits.append("role match")
    if fits:
        return " [Apeiron fit: " + " + ".join(fits) + "]"
    return ""


def briefing_richness(text):
    """Count info segments in a briefing (pipes and dashes separate them)."""
    if not text:
        return 0
    return text.count("|") + text.count(" — ") + 1


def main():
    parser = argparse.ArgumentParser(description="Enrich Account_Summary briefings")
    parser.add_argument("--refresh", action="store_true",
                        help="Re-generate briefings when new version is richer")
    args = parser.parse_args()

    root = Path(__file__).resolve().parents[2]
    bd_path = root / "Apeiron_BR_Gestao_Comercial.xlsx"
    if not bd_path.exists():
        print(f"ERROR: BD not found: {bd_path}")
        return 1

    # Backup
    backup_dir = root / "Support Files" / "Backups"
    backup_dir.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = backup_dir / f"Apeiron_BR_Gestao_Comercial_backup_{ts}.xlsx"
    shutil.copy2(bd_path, backup_path)
    print(f"Backup: {backup_path.name}")

    wb = openpyxl.load_workbook(bd_path)
    ws = wb["Leads"]

    # Build column index from header row
    headers = {}
    for col in range(1, ws.max_column + 1):
        h = cell_text(ws, 1, col)
        if h:
            headers[h] = col

    required = ["Company", "Account_Summary"]
    for r in required:
        if r not in headers:
            print(f"ERROR: Column '{r}' not found. Available: {list(headers.keys())}")
            return 1

    col_company = headers["Company"]
    col_summary = headers["Account_Summary"]
    col_sector = headers.get("Account_Sector")
    col_city = headers.get("Account_City_State")
    col_size = headers.get("Account_Size")
    col_revenue = headers.get("Account_Revenue")
    col_website = headers.get("Account_Website")
    col_phone = headers.get("Account_Phone")
    col_cnpj = headers.get("Account_CNPJ")
    col_brand = headers.get("Account_Brand")
    col_role = headers.get("Role")

    # First pass: collect roles per company for Apeiron fit
    company_roles = {}
    for row in range(2, ws.max_row + 1):
        company = cell_text(ws, row, col_company)
        if not company:
            continue
        key = company.strip().lower()
        role = cell_text(ws, row, col_role) if col_role else ""
        if key not in company_roles:
            company_roles[key] = []
        if role:
            company_roles[key].append(role)

    # Second pass: enrich briefings
    updated = 0
    skipped = 0
    for row in range(2, ws.max_row + 1):
        company = cell_text(ws, row, col_company)
        if not company:
            continue

        current_summary = cell_text(ws, row, col_summary)
        poor = is_poor_briefing(current_summary, company)

        if not poor and not args.refresh:
            skipped += 1
            continue

        sector = cell_text(ws, row, col_sector) if col_sector else ""
        city = cell_text(ws, row, col_city) if col_city else ""
        size = cell_text(ws, row, col_size) if col_size else ""
        revenue = cell_text(ws, row, col_revenue) if col_revenue else ""
        website = cell_text(ws, row, col_website) if col_website else ""
        phone = cell_text(ws, row, col_phone) if col_phone else ""
        cnpj = cell_text(ws, row, col_cnpj) if col_cnpj else ""
        brand = cell_text(ws, row, col_brand) if col_brand else ""

        briefing = build_briefing(company, sector, city, size, revenue, website, phone, cnpj, brand)

        # Add Apeiron fit
        key = company.strip().lower()
        roles = company_roles.get(key, [])
        fit = apeiron_fit_label(sector, roles)
        if fit:
            briefing += fit

        # In refresh mode, only update if new briefing is richer
        if not poor and args.refresh:
            if briefing_richness(briefing) <= briefing_richness(current_summary):
                skipped += 1
                continue

        ws.cell(row, col_summary).value = briefing
        updated += 1

    wb.save(bd_path)
    print(f"Done: {updated} briefings enriched, {skipped} already rich (kept)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
