#!/usr/bin/env python3
import os
import re
from datetime import datetime
from pathlib import Path

import pandas as pd


def norm(v):
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    s = str(v).strip()
    return "" if s.lower() in {"nan", "none", "nat"} else s.lower()


def find_project_root() -> Path:
    here = Path(__file__).resolve().parent
    for p in [here, *here.parents]:
        if (p / "Apeiron_BR_Gestao_Comercial.xlsx").exists():
            return p
    return Path(__file__).resolve().parents[1]


def clean_phone(v):
    return re.sub(r"[^0-9+]", "", str(v or ""))


def full_name(first, last):
    return " ".join([str(first or "").strip(), str(last or "").strip()]).strip()


def key_name_company(name, company):
    return f"{norm(name)}|{norm(company)}"


def safe_read_excel(path, sheet, header=0):
    return pd.read_excel(path, sheet_name=sheet, header=header)


def main():
    base = find_project_root()
    support = base / "Support Files"
    wb_path = os.path.join(base, "Apeiron_BR_Gestao_Comercial.xlsx")
    zoho_dir = os.path.join((support / "Zoho data") if (support / "Zoho data").exists() else (base / "Zoho data"))
    out_dir = os.path.join(support, "_output")
    os.makedirs(out_dir, exist_ok=True)

    ops = safe_read_excel(wb_path, "Leads_Operations", header=0)
    det = safe_read_excel(wb_path, "Leads_Details", header=0)

    xls = pd.ExcelFile(wb_path)
    contacts_sheet = next((s for s in xls.sheet_names if "Contacts" in s), None)
    if contacts_sheet is None:
        raise RuntimeError("Contacts sheet not found.")
    con = pd.read_excel(wb_path, sheet_name=contacts_sheet, header=2)

    z_acc = pd.read_csv(os.path.join(zoho_dir, "Accounts_2026_03_17.csv"), keep_default_na=False)
    z_con = pd.read_csv(os.path.join(zoho_dir, "Contacts_2026_03_17.csv"), keep_default_na=False)
    z_lea = pd.read_csv(os.path.join(zoho_dir, "Leads_2026_03_17.csv"), keep_default_na=False)

    # DB indexes
    db_email_set = set(det.get("Email", pd.Series(dtype=str)).fillna("").astype(str).map(norm))
    db_email_set.discard("")
    db_name_company_set = set(
        ops.apply(lambda r: key_name_company(r.get("Full Name", ""), r.get("Company", "")), axis=1).tolist()
    )

    # Zoho contacts normalized
    z_con_n = pd.DataFrame({"record_id": z_con.get("Record Id", "")})
    z_con_n["module"] = "Contacts"
    z_con_n["company"] = z_con.get("Account Name", "")
    z_con_n["name"] = z_con.apply(lambda r: full_name(r.get("First Name", ""), r.get("Last Name", "")), axis=1)
    z_con_n["email"] = z_con.get("Email", "").fillna("").astype(str).map(norm)
    z_con_n["phone"] = z_con.get("Phone", "").fillna("").astype(str).map(clean_phone)
    z_con_n["mobile"] = z_con.get("Mobile", "").fillna("").astype(str).map(clean_phone)

    # Zoho leads normalized (mapped as contacts for audit)
    z_lea_n = pd.DataFrame({"record_id": z_lea.get("Record Id", "")})
    z_lea_n["module"] = "Leads"
    z_lea_n["company"] = z_lea.get("Company", "")
    z_lea_n["name"] = z_lea.apply(lambda r: full_name(r.get("First Name", ""), r.get("Last Name", "")), axis=1)
    z_lea_n["email"] = z_lea.get("Email", "").fillna("").astype(str).map(norm)
    z_lea_n["phone"] = z_lea.get("Phone", "").fillna("").astype(str).map(clean_phone)
    z_lea_n["mobile"] = z_lea.get("Mobile", "").fillna("").astype(str).map(clean_phone)

    z_all = pd.concat([z_con_n, z_lea_n], ignore_index=True)
    z_all["k_name_company"] = z_all.apply(lambda r: key_name_company(r["name"], r["company"]), axis=1)

    z_all["match_email"] = z_all["email"].map(lambda e: e in db_email_set if e else False)
    z_all["match_name_company"] = z_all["k_name_company"].map(lambda k: k in db_name_company_set if k else False)
    z_all["matched"] = z_all["match_email"] | z_all["match_name_company"]

    # DB duplicates
    db_ops = ops.copy()
    db_ops["k_name_company"] = db_ops.apply(lambda r: key_name_company(r.get("Full Name", ""), r.get("Company", "")), axis=1)
    dup_name_company = (
        db_ops["k_name_company"]
        .value_counts()
        .rename_axis("k_name_company")
        .reset_index(name="count")
    )
    dup_name_company = dup_name_company[dup_name_company["count"] > 1]

    db_det = det.copy()
    db_det["email_norm"] = db_det.get("Email", "").fillna("").astype(str).map(norm)
    dup_email = (
        db_det[db_det["email_norm"] != ""]["email_norm"]
        .value_counts()
        .rename_axis("email")
        .reset_index(name="count")
    )
    dup_email = dup_email[dup_email["count"] > 1]

    unmatched = z_all[~z_all["matched"]].copy()
    unmatched_csv = os.path.join(out_dir, "zoho_unmatched_candidates.csv")
    unmatched[["module", "record_id", "company", "name", "email", "phone", "mobile"]].to_csv(unmatched_csv, index=False)

    report_md = os.path.join(out_dir, "merge_audit_latest.md")
    with open(report_md, "w", encoding="utf-8") as f:
        f.write("# Merge Audit - Zoho vs DB\n\n")
        f.write(f"- generated_at: {datetime.now().isoformat(timespec='seconds')}\n")
        f.write(f"- workbook: `{wb_path}`\n")
        f.write(f"- contacts_sheet: `{contacts_sheet}`\n\n")

        f.write("## Inputs\n")
        f.write(f"- Zoho Accounts: {len(z_acc)}\n")
        f.write(f"- Zoho Contacts: {len(z_con)}\n")
        f.write(f"- Zoho Leads: {len(z_lea)}\n")
        f.write(f"- Zoho Contact-like total: {len(z_all)}\n\n")

        f.write("## DB Current Size\n")
        f.write(f"- Leads_Operations rows: {len(ops)}\n")
        f.write(f"- Leads_Details rows: {len(det)}\n")
        f.write(f"- Contacts rows: {len(con)}\n\n")

        f.write("## Matching Summary (Zoho -> DB)\n")
        f.write(f"- matched by email: {int(z_all['match_email'].sum())}\n")
        f.write(f"- matched by company+name: {int(z_all['match_name_company'].sum())}\n")
        f.write(f"- matched total (union): {int(z_all['matched'].sum())}\n")
        f.write(f"- unmatched candidates: {len(unmatched)}\n")
        f.write(f"- unmatched export: `{unmatched_csv}`\n\n")

        f.write("## Data Quality Flags (DB)\n")
        f.write(f"- duplicate company+name keys in operations: {len(dup_name_company)}\n")
        f.write(f"- duplicate emails in details: {len(dup_email)}\n\n")

        if len(dup_name_company):
            f.write("### Top duplicate company+name keys\n")
            for _, r in dup_name_company.head(10).iterrows():
                f.write(f"- {r['k_name_company']}: {int(r['count'])}\n")
            f.write("\n")

        if len(dup_email):
            f.write("### Top duplicate emails\n")
            for _, r in dup_email.head(10).iterrows():
                f.write(f"- {r['email']}: {int(r['count'])}\n")
            f.write("\n")

    print("Audit report:", report_md)
    print("Unmatched CSV:", unmatched_csv)
    print("Summary:", {
        "zoho_total": int(len(z_all)),
        "matched_total": int(z_all["matched"].sum()),
        "unmatched": int(len(unmatched)),
        "dup_name_company": int(len(dup_name_company)),
        "dup_email": int(len(dup_email)),
    })


if __name__ == "__main__":
    main()
