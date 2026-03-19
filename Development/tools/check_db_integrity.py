#!/usr/bin/env python3
import json
import re
import sys
from pathlib import Path
from datetime import datetime

STATUS = {"New", "In Contact", "Meeting Scheduled", "Proposal Sent", "Client", "Discarded"}
PRIORITY = {"High", "Medium", "Low"}
CRITICAL_FIELDS = ("id", "company", "name", "status", "priority", "createdDate")


def load_payload(leads_js: Path):
    raw = leads_js.read_text(encoding="utf-8")
    match = re.search(r"window\.INITIAL_DATA\s*=\s*(\[.*\]);\s*$", raw, re.S)
    if not match:
        raise RuntimeError("Unable to parse window.INITIAL_DATA from leads.js")
    return json.loads(match.group(1))


def is_valid_date(value: str) -> bool:
    value = str(value or "").strip()
    if not value:
        return False
    for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            datetime.strptime(value, fmt)
            return True
        except ValueError:
            pass
    return False


def check_rows(rows):
    errors = []
    seen = set()
    for idx, row in enumerate(rows, start=1):
        rid = row.get("id")
        try:
            rid_num = int(rid)
            if rid_num <= 0:
                raise ValueError()
        except Exception:
            errors.append(f"row {idx}: invalid id={rid!r}")
            rid_num = None

        if rid_num is not None:
            if rid_num in seen:
                errors.append(f"row {idx}: duplicate id={rid_num}")
            seen.add(rid_num)

        company = str(row.get("company", "")).strip()
        name = str(row.get("name", "")).strip()
        status = str(row.get("status", "")).strip()
        priority = str(row.get("priority", "")).strip()
        created = str(row.get("createdDate", "")).strip()
        next_date = str(row.get("nextActionDate", "")).strip()

        if not company:
            errors.append(f"row {idx}: missing company")
        if not name:
            errors.append(f"row {idx}: missing name")
        if status not in STATUS:
            errors.append(f"row {idx}: invalid status={status!r}")
        if priority not in PRIORITY:
            errors.append(f"row {idx}: invalid priority={priority!r}")
        if not is_valid_date(created):
            errors.append(f"row {idx}: invalid createdDate={created!r}")
        if next_date and not is_valid_date(next_date):
            errors.append(f"row {idx}: invalid nextActionDate={next_date!r}")

    return errors


def main():
    root = Path(__file__).resolve().parents[2]
    leads_js = root / "src" / "dashboard" / "data" / "leads.js"
    if not leads_js.exists():
        print(f"[integrity] leads.js not found: {leads_js}")
        return 1

    rows = load_payload(leads_js)
    errors = check_rows(rows)
    print(f"[integrity] checked rows={len(rows)} critical_fields={','.join(CRITICAL_FIELDS)}")
    if errors:
        print(f"[integrity] FAILED with {len(errors)} issue(s):")
        for item in errors[:20]:
            print(f"  - {item}")
        if len(errors) > 20:
            print(f"  - ... and {len(errors) - 20} more")
        return 1
    print("[integrity] OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
