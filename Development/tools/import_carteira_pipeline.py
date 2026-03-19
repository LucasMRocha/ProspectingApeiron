#!/usr/bin/env python3
from __future__ import annotations

import argparse
import ctypes
import json
import os
import re
import shutil
import subprocess
import sys
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd


ROOT = Path(__file__).resolve().parents[2]
DB_PATH = ROOT / "Apeiron_BR_Gestao_Comercial.xlsx"
DEFAULT_SOURCE = ROOT / "Support Files" / "New Contacts" / "CARTEIRA - AREA INDUSTRIAL 2.xls revisado.xls"
STAGING_DIR = ROOT / "Support Files" / "_staging" / "carteira_import"
UNDO_DIR = STAGING_DIR / "_undo"
LATEST_PLAN = STAGING_DIR / "latest_plan.json"
LATEST_APPLY = UNDO_DIR / "latest_apply.json"

OPS_SHEET = "Leads_Operations"
DET_SHEET = "Leads_Details"

IMPORT_SOURCE_LABEL = "CARTEIRA - AREA INDUSTRIAL 2"
DEFAULT_STATUS = "New"
DEFAULT_PRIORITY = "Medium"
DEFAULT_ROLE = "To Map"
DEFAULT_CONTACT = "To Map"
DEFAULT_LEAD_INSIGHT = "Imported from CARTEIRA industrial portfolio"

DETAILS_COLUMNS = [
    "ID",
    "Company",
    "Full Name",
    "Role",
    "Department",
    "Email",
    "Phone",
    "LinkedIn",
    "Location",
    "Status",
    "Priority",
    "Channel Email",
    "Channel Phone",
    "Channel LinkedIn",
    "Channel WhatsApp",
    "Next Action",
    "Next Action Date",
    "Last Contact",
    "Source",
    "Lead Insight",
    "Notes",
    "Relevance",
    "Account_Summary",
    "Account_Sector",
    "Account_City_State",
    "Account_Size",
    "Account_Revenue",
    "Account_Website",
    "Account_Phone",
    "Created Date",
    "SharePoint_ID",
    "Global",
    "Account_CNPJ",
    "Account_Brand",
]

OPS_COLUMNS = [
    "ID",
    "Company",
    "Full Name",
    "Role",
    "Status",
    "Priority",
    "Next Action",
    "Next Action Date",
    "Channel",
    "Channel Email",
    "Channel Phone",
    "Channel LinkedIn",
    "Channel WhatsApp",
    "Lead Insight",
    "Action Due",
    "Possible Duplicate",
    "Last Contact",
    "Source",
    "SharePoint_ID",
]


@dataclass
class SourceRecord:
    source_id: str
    company: str
    cnpj: str
    phone: str
    city: str
    state: str
    region: str
    country: str
    address: str
    neighborhood: str
    contact: str
    sector: str


def now_ts() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def txt(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""
    return str(v).strip()


def normalize_key(v: str) -> str:
    s = unicodedata.normalize("NFKD", txt(v))
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"[^a-zA-Z0-9]+", " ", s).strip().lower()
    return re.sub(r"\s+", " ", s)


def clean_cnpj(v: str) -> str:
    digits = re.sub(r"\D+", "", txt(v))
    if len(digits) == 14:
        return f"{digits[0:2]}.{digits[2:5]}.{digits[5:8]}/{digits[8:12]}-{digits[12:14]}"
    return txt(v)


def first_non_empty(*values: Any) -> str:
    for value in values:
        t = txt(value)
        if t:
            return t
    return ""


def best_text(a: str, b: str) -> str:
    a_t, b_t = txt(a), txt(b)
    if not a_t:
        return b_t
    if not b_t:
        return a_t
    if len(b_t) > len(a_t):
        return b_t
    return a_t


def city_state(city: str, state: str) -> str:
    c = txt(city)
    s = txt(state)
    if c and s:
        return f"{c}/{s}"
    return c or s


def ensure_dirs() -> None:
    STAGING_DIR.mkdir(parents=True, exist_ok=True)
    UNDO_DIR.mkdir(parents=True, exist_ok=True)


def require_xlrd() -> None:
    try:
        import xlrd  # noqa: F401
    except Exception as exc:  # noqa: BLE001
        raise RuntimeError(
            "xlrd is required for .xls import. Install with: python -m pip install --user xlrd"
        ) from exc


def read_source_records(source_path: Path) -> list[SourceRecord]:
    require_xlrd()
    df = pd.read_excel(source_path, sheet_name=0, dtype=object, engine="xlrd")
    expected = {
        "Identificador",
        "Nome completo / Razão social",
        "CNPJ/CPF",
        "Fone",
        "Nome do logradouro",
        "Número",
        "Bairro",
        "Município",
        "UF",
        "Região",
        "País",
        "Contato",
        "grupo_superior",
        "nivel_arvore_artificial",
    }
    missing = [c for c in expected if c not in df.columns]
    if missing:
        raise RuntimeError(f"Source columns missing: {missing}")

    out: list[SourceRecord] = []
    for _, row in df.iterrows():
        ident = txt(row.get("Identificador"))
        company = txt(row.get("Nome completo / Razão social"))
        if ident.lower().startswith("total itens"):
            continue
        if not company:
            continue
        contact = txt(row.get("Contato"))
        number = txt(row.get("Número"))
        street = txt(row.get("Nome do logradouro"))
        addr = " ".join(x for x in [street, number] if x)
        record = SourceRecord(
            source_id=ident,
            company=company,
            cnpj=clean_cnpj(row.get("CNPJ/CPF")),
            phone=txt(row.get("Fone")),
            city=txt(row.get("Município")),
            state=txt(row.get("UF")),
            region=txt(row.get("Região")),
            country=txt(row.get("País")) or "BR",
            address=addr,
            neighborhood=txt(row.get("Bairro")),
            contact=contact,
            sector=first_non_empty(row.get("grupo_superior"), row.get("nivel_arvore_artificial"), row.get("Região"), "Industrial"),
        )
        out.append(record)
    return dedupe_source(out)


def dedupe_source(records: list[SourceRecord]) -> list[SourceRecord]:
    by_key: dict[str, SourceRecord] = {}
    for rec in records:
        key = normalize_key(rec.cnpj) if rec.cnpj else normalize_key(rec.company)
        if not key:
            continue
        if key not in by_key:
            by_key[key] = rec
            continue
        prev = by_key[key]
        by_key[key] = SourceRecord(
            source_id=best_text(prev.source_id, rec.source_id),
            company=best_text(prev.company, rec.company),
            cnpj=best_text(prev.cnpj, rec.cnpj),
            phone=best_text(prev.phone, rec.phone),
            city=best_text(prev.city, rec.city),
            state=best_text(prev.state, rec.state),
            region=best_text(prev.region, rec.region),
            country=best_text(prev.country, rec.country),
            address=best_text(prev.address, rec.address),
            neighborhood=best_text(prev.neighborhood, rec.neighborhood),
            contact=best_text(prev.contact, rec.contact),
            sector=best_text(prev.sector, rec.sector),
        )
    return list(by_key.values())


def load_workbook_sheets(db_path: Path) -> tuple[list[str], dict[str, pd.DataFrame]]:
    xls = pd.ExcelFile(db_path)
    names = list(xls.sheet_names)
    sheets = {name: pd.read_excel(xls, sheet_name=name, dtype=object) for name in names}
    return names, sheets


def ensure_sheet_columns(df: pd.DataFrame, required: list[str]) -> pd.DataFrame:
    out = df.copy()
    for col in required:
        if col not in out.columns:
            out[col] = ""
    return out


def id_int(v: Any) -> int | None:
    t = txt(v)
    if not t:
        return None
    try:
        return int(float(t))
    except Exception:
        return None


def completeness(row: pd.Series | dict[str, Any]) -> int:
    fields = [
        "Company",
        "Full Name",
        "Role",
        "Status",
        "Priority",
        "Phone",
        "Email",
        "LinkedIn",
        "Account_CNPJ",
        "Account_City_State",
        "Account_Sector",
    ]
    score = 0
    for field in fields:
        if txt(row.get(field)):
            score += 1
    return score


def build_plan(source_records: list[SourceRecord], ops: pd.DataFrame, det: pd.DataFrame, dedupe_existing: bool) -> dict[str, Any]:
    ops = ensure_sheet_columns(ops, OPS_COLUMNS)
    det = ensure_sheet_columns(det, DETAILS_COLUMNS)

    ops_by_id: dict[int, pd.Series] = {}
    det_by_id: dict[int, pd.Series] = {}
    for _, row in ops.iterrows():
        iid = id_int(row.get("ID"))
        if iid is not None:
            ops_by_id[iid] = row
    for _, row in det.iterrows():
        iid = id_int(row.get("ID"))
        if iid is not None:
            det_by_id[iid] = row

    all_ids = sorted(set(ops_by_id.keys()) | set(det_by_id.keys()))
    next_id = (max(all_ids) + 1) if all_ids else 1

    idx_cnpj: dict[str, set[int]] = {}
    idx_company: dict[str, set[int]] = {}
    for iid, row in det_by_id.items():
        cnpj = normalize_key(clean_cnpj(row.get("Account_CNPJ")))
        company_key = normalize_key(row.get("Company"))
        if cnpj:
            idx_cnpj.setdefault(cnpj, set()).add(iid)
        if company_key:
            idx_company.setdefault(company_key, set()).add(iid)

    inserts: list[dict[str, Any]] = []
    updates: list[dict[str, Any]] = []
    deletions: list[dict[str, Any]] = []

    def maybe_set(changes: dict[str, Any], field: str, current_value: Any, incoming_value: Any) -> None:
        cur = txt(current_value)
        inc = txt(incoming_value)
        if not inc:
            return
        if not cur:
            changes[field] = inc

    for rec in source_records:
        company_key = normalize_key(rec.company)
        cnpj_key = normalize_key(rec.cnpj)
        candidates: set[int] = set()
        if cnpj_key and cnpj_key in idx_cnpj:
            candidates |= idx_cnpj[cnpj_key]
        if company_key and company_key in idx_company:
            candidates |= idx_company[company_key]

        target_id: int | None = None
        if candidates:
            ranked = sorted(
                candidates,
                key=lambda iid: completeness(det_by_id.get(iid, {})),
                reverse=True,
            )
            target_id = ranked[0]
            if dedupe_existing and len(ranked) > 1:
                keeper = ranked[0]
                for duplicate_id in ranked[1:]:
                    duplicate_row = det_by_id.get(duplicate_id, {})
                    deletions.append(
                        {
                            "kind": "drop_duplicate_id",
                            "id": duplicate_id,
                            "keeper_id": keeper,
                            "company": txt(duplicate_row.get("Company")),
                            "name": txt(duplicate_row.get("Full Name")),
                            "reason": f"Duplicate matched by company/cnpj during import ({rec.company})",
                        }
                    )

        if target_id is None:
            contact_name = rec.contact or DEFAULT_CONTACT
            city_state_value = city_state(rec.city, rec.state)
            inserts.append(
                {
                    "id": next_id,
                    "company": rec.company,
                    "full_name": contact_name,
                    "role": DEFAULT_ROLE if contact_name == DEFAULT_CONTACT else "",
                    "status": DEFAULT_STATUS,
                    "priority": DEFAULT_PRIORITY,
                    "source": IMPORT_SOURCE_LABEL,
                    "lead_insight": DEFAULT_LEAD_INSIGHT,
                    "channel": "Phone" if rec.phone else "No contact",
                    "channel_phone": "Yes" if rec.phone else "",
                    "phone": rec.phone,
                    "location": city_state_value,
                    "account_summary": rec.company,
                    "account_sector": rec.sector,
                    "account_city_state": city_state_value,
                    "account_phone": rec.phone,
                    "account_cnpj": rec.cnpj,
                    "created_date": datetime.now().strftime("%Y-%m-%d"),
                    "notes": "Imported from CARTEIRA industrial portfolio",
                }
            )
            next_id += 1
            continue

        drow = det_by_id.get(target_id, {})
        orow = ops_by_id.get(target_id, {})
        d_changes: dict[str, Any] = {}
        o_changes: dict[str, Any] = {}

        maybe_set(d_changes, "Company", drow.get("Company"), rec.company)
        maybe_set(o_changes, "Company", orow.get("Company"), rec.company)
        maybe_set(d_changes, "Full Name", drow.get("Full Name"), rec.contact or DEFAULT_CONTACT)
        maybe_set(o_changes, "Full Name", orow.get("Full Name"), rec.contact or DEFAULT_CONTACT)
        maybe_set(d_changes, "Role", drow.get("Role"), DEFAULT_ROLE)
        maybe_set(o_changes, "Role", orow.get("Role"), DEFAULT_ROLE)
        maybe_set(d_changes, "Status", drow.get("Status"), DEFAULT_STATUS)
        maybe_set(o_changes, "Status", orow.get("Status"), DEFAULT_STATUS)
        maybe_set(d_changes, "Priority", drow.get("Priority"), DEFAULT_PRIORITY)
        maybe_set(o_changes, "Priority", orow.get("Priority"), DEFAULT_PRIORITY)
        maybe_set(d_changes, "Source", drow.get("Source"), IMPORT_SOURCE_LABEL)
        maybe_set(o_changes, "Source", orow.get("Source"), IMPORT_SOURCE_LABEL)
        maybe_set(d_changes, "Lead Insight", drow.get("Lead Insight"), DEFAULT_LEAD_INSIGHT)
        maybe_set(o_changes, "Lead Insight", orow.get("Lead Insight"), DEFAULT_LEAD_INSIGHT)
        maybe_set(d_changes, "Phone", drow.get("Phone"), rec.phone)
        maybe_set(d_changes, "Location", drow.get("Location"), city_state(rec.city, rec.state))
        maybe_set(d_changes, "Account_Summary", drow.get("Account_Summary"), rec.company)
        maybe_set(d_changes, "Account_Sector", drow.get("Account_Sector"), rec.sector)
        maybe_set(d_changes, "Account_City_State", drow.get("Account_City_State"), city_state(rec.city, rec.state))
        maybe_set(d_changes, "Account_Phone", drow.get("Account_Phone"), rec.phone)
        maybe_set(d_changes, "Account_CNPJ", drow.get("Account_CNPJ"), rec.cnpj)
        maybe_set(d_changes, "Created Date", drow.get("Created Date"), datetime.now().strftime("%Y-%m-%d"))
        maybe_set(d_changes, "Notes", drow.get("Notes"), "Imported from CARTEIRA industrial portfolio")

        if txt(rec.phone):
            maybe_set(o_changes, "Channel", orow.get("Channel"), "Phone")
            maybe_set(o_changes, "Channel Phone", orow.get("Channel Phone"), "Yes")
            maybe_set(d_changes, "Channel Phone", drow.get("Channel Phone"), "Yes")

        if d_changes:
            updates.append(
                {
                    "id": target_id,
                    "target": "details",
                    "changes": d_changes,
                    "reason": f"Enrichment from source company {rec.company}",
                }
            )
        if o_changes:
            updates.append(
                {
                    "id": target_id,
                    "target": "operations",
                    "changes": o_changes,
                    "reason": f"Enrichment from source company {rec.company}",
                }
            )

    return {
        "metadata": {
            "created_at": datetime.now().isoformat(timespec="seconds"),
            "source_label": IMPORT_SOURCE_LABEL,
            "source_record_count": len(source_records),
            "dedupe_existing": bool(dedupe_existing),
            "approved": False,
        },
        "summary": {
            "inserts": len(inserts),
            "updates": len(updates),
            "deletions": len(deletions),
        },
        "inserts": inserts,
        "updates": updates,
        "deletions": deletions,
    }


def save_plan(plan: dict[str, Any], source_records: list[SourceRecord], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")

    raw_df = pd.DataFrame([r.__dict__ for r in source_records])
    if not raw_df.empty:
        raw_df.to_csv(STAGING_DIR / "source_clean.csv", index=False, encoding="utf-8-sig")

    summary_lines = [
        "CARTEIRA IMPORT DRY-RUN",
        f"Created at: {plan['metadata']['created_at']}",
        f"Source records: {plan['metadata']['source_record_count']}",
        f"Inserts: {plan['summary']['inserts']}",
        f"Updates: {plan['summary']['updates']}",
        f"Deletions: {plan['summary']['deletions']}",
        f"Approved: {plan['metadata']['approved']}",
    ]
    (STAGING_DIR / "plan_summary.txt").write_text("\n".join(summary_lines), encoding="utf-8")


def load_plan(path: Path) -> dict[str, Any]:
    if not path.exists():
        raise RuntimeError(f"Plan not found: {path}")
    return json.loads(path.read_text(encoding="utf-8"))


def save_plan_file(plan: dict[str, Any], path: Path) -> None:
    path.write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")


def popup_yes_no_cancel(title: str, message: str) -> str:
    if os.name != "nt":
        while True:
            ans = input(f"{title}\n{message}\n[y]es / [n]o / [c]ancel: ").strip().lower()
            if ans in {"y", "yes"}:
                return "yes"
            if ans in {"n", "no"}:
                return "no"
            if ans in {"c", "cancel"}:
                return "cancel"
    flags = 0x00000003 | 0x00000020 | 0x00001000  # YESNOCANCEL + QUESTION + TOPMOST
    result = ctypes.windll.user32.MessageBoxW(None, message, title, flags)
    if result == 6:
        return "yes"
    if result == 7:
        return "no"
    return "cancel"


def run_backup_script() -> None:
    backup_script = ROOT / "Development" / "tools" / "backup-db.ps1"
    if not backup_script.exists():
        return
    subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(backup_script)],
        check=True,
    )


def refresh_dashboard_payload() -> None:
    refresh_script = ROOT / "Development" / "tools" / "refresh_dashboard_from_db.py"
    if refresh_script.exists():
        subprocess.run([sys.executable, str(refresh_script)], check=True)


def apply_plan(plan: dict[str, Any], plan_path: Path, require_popup: bool) -> dict[str, Any]:
    names, sheets = load_workbook_sheets(DB_PATH)
    if OPS_SHEET not in sheets or DET_SHEET not in sheets:
        raise RuntimeError("Workbook missing required sheets Leads_Operations/Leads_Details")

    ops = ensure_sheet_columns(sheets[OPS_SHEET], OPS_COLUMNS)
    det = ensure_sheet_columns(sheets[DET_SHEET], DETAILS_COLUMNS)

    run_backup_script()

    undo_snapshot = UNDO_DIR / f"Apeiron_BR_Gestao_Comercial_preimport_{now_ts()}.xlsx"
    shutil.copy2(DB_PATH, undo_snapshot)

    ops_id_to_idx = {id_int(v): idx for idx, v in enumerate(ops["ID"]) if id_int(v) is not None}
    det_id_to_idx = {id_int(v): idx for idx, v in enumerate(det["ID"]) if id_int(v) is not None}

    applied_updates = 0
    applied_inserts = 0
    applied_deletions = 0
    skipped_deletions = 0
    deletion_decisions: list[dict[str, Any]] = []

    for action in plan.get("updates", []):
        iid = id_int(action.get("id"))
        if iid is None:
            continue
        target = txt(action.get("target")).lower()
        changes = action.get("changes", {})
        if target == "details":
            idx = det_id_to_idx.get(iid)
            if idx is None:
                continue
            for field, new_value in changes.items():
                det.at[idx, field] = new_value
            applied_updates += 1
        elif target == "operations":
            idx = ops_id_to_idx.get(iid)
            if idx is None:
                continue
            for field, new_value in changes.items():
                ops.at[idx, field] = new_value
            applied_updates += 1

    for ins in plan.get("inserts", []):
        iid = id_int(ins.get("id"))
        if iid is None:
            continue
        detail_row = {c: "" for c in DETAILS_COLUMNS}
        detail_row.update(
            {
                "ID": iid,
                "Company": ins.get("company", ""),
                "Full Name": ins.get("full_name", ""),
                "Role": ins.get("role", ""),
                "Status": ins.get("status", DEFAULT_STATUS),
                "Priority": ins.get("priority", DEFAULT_PRIORITY),
                "Source": ins.get("source", IMPORT_SOURCE_LABEL),
                "Lead Insight": ins.get("lead_insight", DEFAULT_LEAD_INSIGHT),
                "Notes": ins.get("notes", ""),
                "Location": ins.get("location", ""),
                "Phone": ins.get("phone", ""),
                "Channel Phone": ins.get("channel_phone", ""),
                "Next Action": "",
                "Next Action Date": "",
                "Account_Summary": ins.get("account_summary", ""),
                "Account_Sector": ins.get("account_sector", ""),
                "Account_City_State": ins.get("account_city_state", ""),
                "Account_Phone": ins.get("account_phone", ""),
                "Account_CNPJ": ins.get("account_cnpj", ""),
                "Created Date": ins.get("created_date", datetime.now().strftime("%Y-%m-%d")),
                "Account_Brand": "",
            }
        )
        ops_row = {c: "" for c in OPS_COLUMNS}
        ops_row.update(
            {
                "ID": iid,
                "Company": ins.get("company", ""),
                "Full Name": ins.get("full_name", ""),
                "Role": ins.get("role", ""),
                "Status": ins.get("status", DEFAULT_STATUS),
                "Priority": ins.get("priority", DEFAULT_PRIORITY),
                "Channel": ins.get("channel", "No contact"),
                "Channel Phone": ins.get("channel_phone", ""),
                "Lead Insight": ins.get("lead_insight", DEFAULT_LEAD_INSIGHT),
                "Source": ins.get("source", IMPORT_SOURCE_LABEL),
            }
        )
        det = pd.concat([det, pd.DataFrame([detail_row])], ignore_index=True)
        ops = pd.concat([ops, pd.DataFrame([ops_row])], ignore_index=True)
        applied_inserts += 1

    for deletion in plan.get("deletions", []):
        iid = id_int(deletion.get("id"))
        if iid is None:
            continue
        approved = True
        if require_popup:
            title = "Approve Automatic Deletion"
            msg = (
                f"Delete duplicate contact/company data?\n\n"
                f"ID: {iid}\n"
                f"Company: {txt(deletion.get('company'))}\n"
                f"Name: {txt(deletion.get('name'))}\n"
                f"Reason: {txt(deletion.get('reason'))}\n\n"
                f"You can undo this entire apply later."
            )
            decision = popup_yes_no_cancel(title, msg)
            deletion_decisions.append({"id": iid, "decision": decision, "reason": txt(deletion.get("reason"))})
            if decision == "cancel":
                raise RuntimeError("Apply canceled by user during deletion approvals.")
            if decision == "no":
                approved = False
                skipped_deletions += 1
        if not approved:
            continue
        ops = ops[ops["ID"].apply(id_int) != iid].copy()
        det = det[det["ID"].apply(id_int) != iid].copy()
        applied_deletions += 1

    sheets[OPS_SHEET] = ops
    sheets[DET_SHEET] = det

    with pd.ExcelWriter(DB_PATH, engine="openpyxl", mode="w") as writer:
        for sheet_name in names:
            frame = sheets.get(sheet_name, pd.DataFrame())
            frame.to_excel(writer, sheet_name=sheet_name, index=False)

    refresh_dashboard_payload()

    apply_report = {
        "applied_at": datetime.now().isoformat(timespec="seconds"),
        "plan_path": str(plan_path),
        "undo_snapshot": str(undo_snapshot),
        "applied_updates": applied_updates,
        "applied_inserts": applied_inserts,
        "applied_deletions": applied_deletions,
        "skipped_deletions": skipped_deletions,
        "deletion_decisions": deletion_decisions,
    }
    LATEST_APPLY.parent.mkdir(parents=True, exist_ok=True)
    LATEST_APPLY.write_text(json.dumps(apply_report, ensure_ascii=False, indent=2), encoding="utf-8")
    return apply_report


def command_dry_run(args: argparse.Namespace) -> int:
    ensure_dirs()
    source = Path(args.source).resolve()
    if not source.exists():
        raise RuntimeError(f"Source file not found: {source}")
    if not DB_PATH.exists():
        raise RuntimeError(f"DB file not found: {DB_PATH}")

    names, sheets = load_workbook_sheets(DB_PATH)
    if OPS_SHEET not in names or DET_SHEET not in names:
        raise RuntimeError("DB workbook missing Leads_Operations/Leads_Details")

    records = read_source_records(source)
    plan = build_plan(records, sheets[OPS_SHEET], sheets[DET_SHEET], dedupe_existing=bool(args.dedupe_existing))
    plan["metadata"]["source_file"] = str(source)
    plan["metadata"]["plan_file"] = str(LATEST_PLAN)
    save_plan(plan, records, LATEST_PLAN)

    out_name = f"plan_{now_ts()}.json"
    save_plan_file(plan, STAGING_DIR / out_name)
    print(json.dumps(plan["summary"], ensure_ascii=False))
    print(f"Plan saved: {LATEST_PLAN}")
    return 0


def command_approve(args: argparse.Namespace) -> int:
    ensure_dirs()
    plan_path = Path(args.plan).resolve() if args.plan else LATEST_PLAN
    plan = load_plan(plan_path)
    plan.setdefault("metadata", {})
    plan["metadata"]["approved"] = True
    plan["metadata"]["approved_at"] = datetime.now().isoformat(timespec="seconds")
    plan["metadata"]["approved_by"] = os.environ.get("USERNAME", "unknown")
    save_plan_file(plan, plan_path)
    if plan_path != LATEST_PLAN:
        save_plan_file(plan, LATEST_PLAN)
    print(f"Approved plan: {plan_path}")
    return 0


def command_apply(args: argparse.Namespace) -> int:
    ensure_dirs()
    plan_path = Path(args.plan).resolve() if args.plan else LATEST_PLAN
    plan = load_plan(plan_path)
    approved = bool(plan.get("metadata", {}).get("approved"))
    if not approved and not args.force_without_approval:
        raise RuntimeError("Plan is not approved. Run approve command first.")

    report = apply_plan(plan, plan_path=plan_path, require_popup=not bool(args.no_popup))
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


def command_undo(_args: argparse.Namespace) -> int:
    if not LATEST_APPLY.exists():
        raise RuntimeError("No apply history found for undo.")
    info = json.loads(LATEST_APPLY.read_text(encoding="utf-8"))
    snapshot = Path(info.get("undo_snapshot", ""))
    if not snapshot.exists():
        raise RuntimeError(f"Undo snapshot not found: {snapshot}")
    shutil.copy2(snapshot, DB_PATH)
    refresh_dashboard_payload()
    print(f"Undo completed from snapshot: {snapshot}")
    return 0


def parser_build() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Import CARTEIRA industrial data into Apeiron DB.")
    sub = p.add_subparsers(dest="cmd", required=True)

    p_dry = sub.add_parser("dry-run", help="Generate merge plan without changing DB.")
    p_dry.add_argument("--source", default=str(DEFAULT_SOURCE), help="Source .xls file path.")
    p_dry.add_argument("--dedupe-existing", action="store_true", help="Propose duplicate ID deletions in DB.")
    p_dry.set_defaults(func=command_dry_run)

    p_approve = sub.add_parser("approve", help="Approve generated plan.")
    p_approve.add_argument("--plan", default="", help="Plan path (defaults to latest_plan.json).")
    p_approve.set_defaults(func=command_approve)

    p_apply = sub.add_parser("apply", help="Apply approved plan to DB.")
    p_apply.add_argument("--plan", default="", help="Plan path (defaults to latest_plan.json).")
    p_apply.add_argument("--no-popup", action="store_true", help="Disable deletion approval popups.")
    p_apply.add_argument("--force-without-approval", action="store_true", help="Apply even if plan not approved.")
    p_apply.set_defaults(func=command_apply)

    p_undo = sub.add_parser("undo-last", help="Undo last apply using snapshot.")
    p_undo.set_defaults(func=command_undo)
    return p


def main() -> int:
    try:
        args = parser_build().parse_args()
        return int(args.func(args))
    except Exception as exc:  # noqa: BLE001
        print(f"ERROR: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
