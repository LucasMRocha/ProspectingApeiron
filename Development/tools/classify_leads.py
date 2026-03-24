#!/usr/bin/env python3
"""Classify leads by outreach readiness and generate gap report.

Categories:
  - contact-ready: has email + phone + linkedin + sector
  - needs-enrichment: has company + name but missing key outreach fields
  - incomplete: missing fundamental fields (name or company)

Usage:
    python classify_leads.py              # report only
    python classify_leads.py --save       # write Readiness column to BD
"""
import argparse
import io
import shutil
import sys
from datetime import datetime
from pathlib import Path

import openpyxl

if sys.stdout and hasattr(sys.stdout, "buffer"):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")


def cell_text(ws, row, col):
    v = ws.cell(row, col).value
    if v is None:
        return ""
    return str(v).strip()


REQUIRED_FIELDS = ["Full Name", "Company"]
OUTREACH_FIELDS = ["Email", "Account_Phone", "LinkedIn"]
ENRICHMENT_FIELDS = ["Account_Sector", "Account_Size", "Account_City_State",
                     "Role", "Account_Summary"]


def classify_lead(fields_present):
    """Return readiness category and missing fields."""
    # Check fundamentals
    for f in REQUIRED_FIELDS:
        if f not in fields_present:
            return "incomplete", [f for f in REQUIRED_FIELDS if f not in fields_present]

    # Check outreach readiness
    outreach_missing = [f for f in OUTREACH_FIELDS if f not in fields_present]
    enrich_missing = [f for f in ENRICHMENT_FIELDS if f not in fields_present]

    has_email = "Email" in fields_present
    has_phone = "Account_Phone" in fields_present
    has_linkedin = "LinkedIn" in fields_present
    has_sector = "Account_Sector" in fields_present

    # contact-ready: at least email OR (phone + linkedin), plus sector
    if (has_email or (has_phone and has_linkedin)) and has_sector:
        return "contact-ready", outreach_missing + enrich_missing

    return "needs-enrichment", outreach_missing + enrich_missing


def main():
    parser = argparse.ArgumentParser(description="Classify leads by readiness")
    parser.add_argument("--save", action="store_true",
                        help="Write Readiness column to BD")
    args = parser.parse_args()

    root = Path(__file__).resolve().parents[2]
    bd_path = root / "Apeiron_BR_Gestao_Comercial.xlsx"
    if not bd_path.exists():
        print(f"ERROR: BD not found: {bd_path}")
        return 1

    wb = openpyxl.load_workbook(bd_path)
    ws = wb["Leads"]

    headers = {}
    for col in range(1, ws.max_column + 1):
        h = cell_text(ws, 1, col)
        if h:
            headers[h] = col

    # Classify each lead
    categories = {"contact-ready": [], "needs-enrichment": [], "incomplete": []}
    gap_counter = {}  # field -> count of leads missing it
    all_fields = REQUIRED_FIELDS + OUTREACH_FIELDS + ENRICHMENT_FIELDS

    for f in all_fields:
        gap_counter[f] = 0

    for row in range(2, ws.max_row + 1):
        fields_present = set()
        for field_name in all_fields:
            col = headers.get(field_name)
            if col and cell_text(ws, row, col):
                fields_present.add(field_name)

        cat, missing = classify_lead(fields_present)
        name = cell_text(ws, row, headers.get("Full Name", 0)) or f"Row {row}"
        company = cell_text(ws, row, headers.get("Company", 0))
        categories[cat].append((row, name, company, missing))

        for f in all_fields:
            if f not in fields_present:
                gap_counter[f] += 1

    total = sum(len(v) for v in categories.values())

    # Print report
    print("=== Lead Readiness Report ===")
    print(f"Total leads: {total}\n")

    for cat in ["contact-ready", "needs-enrichment", "incomplete"]:
        leads = categories[cat]
        pct = len(leads) * 100 // total if total else 0
        print(f"  {cat}: {len(leads)} ({pct}%)")

    print(f"\n--- Gap Analysis (fields missing) ---")
    for f in sorted(gap_counter, key=gap_counter.get, reverse=True):
        cnt = gap_counter[f]
        if cnt > 0:
            pct = cnt * 100 // total
            print(f"  {f}: {cnt} leads missing ({pct}%)")

    # Top needs-enrichment leads (closest to contact-ready)
    needs = categories["needs-enrichment"]
    needs_sorted = sorted(needs, key=lambda x: len(x[3]))  # fewest gaps first
    print(f"\n--- Top 10 closest to contact-ready ---")
    for row, name, company, missing in needs_sorted[:10]:
        print(f"  Row {row}: {name} ({company}) -- missing: {', '.join(missing)}")

    # Save Readiness column
    if args.save:
        # Backup
        backup_dir = root / "Support Files" / "Backups"
        backup_dir.mkdir(parents=True, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        shutil.copy2(bd_path, backup_dir / f"Apeiron_BR_Gestao_Comercial_backup_{ts}.xlsx")

        # Find or create Readiness column
        col_readiness = headers.get("Readiness")
        if not col_readiness:
            col_readiness = ws.max_column + 1
            ws.cell(1, col_readiness).value = "Readiness"

        for cat, leads in categories.items():
            for row, name, company, missing in leads:
                ws.cell(row, col_readiness).value = cat

        wb.save(bd_path)
        print(f"\nReadiness column saved to BD ({total} leads classified)")
    else:
        print("\n(use --save to write Readiness column to BD)")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
