#!/usr/bin/env python3
import hashlib
import json
from pathlib import Path

import pandas as pd


def text(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""
    return str(v).strip()


def date_text(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""
    if isinstance(v, pd.Timestamp):
        return v.strftime("%Y-%m-%d")
    raw = str(v).strip()
    if not raw:
        return ""
    parsed = pd.to_datetime(raw, errors="coerce")
    if pd.isna(parsed):
        return raw
    return parsed.strftime("%Y-%m-%d")


def to_int(v):
    if v is None:
        return None
    try:
        if isinstance(v, float) and pd.isna(v):
            return None
        return int(float(v))
    except Exception:
        return None


def dashboard_targets(root: Path) -> list[Path]:
    candidates = [
        root / "Development" / "src" / "dashboard" / "data" / "leads.js",
        root / "src" / "dashboard" / "data" / "leads.js",
    ]
    out = []
    seen = set()
    for p in candidates:
        key = str(p.resolve()).lower() if p.exists() else str(p).lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(p)
    return out


def export_payload(workbook_path: Path):
    xls = pd.ExcelFile(workbook_path)
    available = xls.sheet_names
    if "Leads" in available:
        sheet_name = "Leads"
    elif "Leads_Operations" in available:
        sheet_name = "Leads_Operations"
    else:
        raise ValueError(f"No 'Leads' or 'Leads_Operations' sheet found. Available: {available}")
    df = pd.read_excel(xls, sheet_name=sheet_name)

    id_col = "Contact_ID" if "Contact_ID" in df.columns else "ID"
    rows = []
    for _, row in df.dropna(subset=[id_col]).iterrows():
        iid = to_int(row.get(id_col))
        if iid is None:
            continue
        item = {
            "id": iid,
            "company": text(row.get("Company")),
            "name": text(row.get("Full Name")),
            "role": text(row.get("Role")),
            "status": text(row.get("Status")),
            "priority": text(row.get("Priority")),
            "nextAction": text(row.get("Next Action")),
            "nextActionDate": date_text(row.get("Next Action Date")),
            "createdDate": date_text(row.get("Created Date")),
            "channel": text(row.get("Channel")),
            "owner": text(row.get("Lead Insight")),
            "source": text(row.get("Source")),
            "sharePointId": text(row.get("SharePoint_ID")),
            "possibleDuplicate": text(row.get("Possible Duplicate")),
            "email": text(row.get("Email")),
            "phone": text(row.get("Phone")),
            "linkedin": text(row.get("LinkedIn")),
            "location": text(row.get("Location")),
            "department": text(row.get("Department")),
            "notes": text(row.get("Notes")),
            "accountSummary": text(row.get("Account_Summary")),
            "accountSector": text(row.get("Account_Sector")),
            "accountCityState": text(row.get("Account_City_State")),
            "accountSize": text(row.get("Account_Size")),
            "accountRevenue": text(row.get("Account_Revenue")),
            "accountWebsite": text(row.get("Account_Website")),
            "accountPhone": text(row.get("Account_Phone")),
            "accountCnpj": text(row.get("Account_CNPJ")),
            "accountBrand": text(row.get("Account_Brand") or row.get("Brand")),
            "tier": text(row.get("Tier")),
            "leadScore": text(row.get("Lead_Score")),
        }
        rows.append(item)

    payload = json.dumps(rows, ensure_ascii=False)
    hash12 = hashlib.sha1(payload.encode("utf-8")).hexdigest()[:12]
    mtime = int(workbook_path.stat().st_mtime)
    version = f"db{mtime}-r{len(rows)}-{hash12}"
    return rows, payload, version


def main() -> int:
    root = Path(__file__).resolve().parents[2]
    workbook_path = root / "Apeiron_BR_Gestao_Comercial.xlsx"
    if not workbook_path.exists():
        print(f"ERROR: Workbook not found: {workbook_path}")
        return 1

    rows, payload, version = export_payload(workbook_path)
    content = (
        f'window.INITIAL_DATA_VERSION = "{version}";\n'
        f"window.INITIAL_DATA = {payload};\n"
    )

    targets = dashboard_targets(root)
    for out_js in targets:
        out_js.parent.mkdir(parents=True, exist_ok=True)
        out_js.write_text(content, encoding="utf-8")

    print(f"Refreshed dashboard payload from DB: {len(rows)} rows | version={version}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
