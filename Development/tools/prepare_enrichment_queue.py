#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Tuple

import pandas as pd


def text(v) -> str:
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    s = str(v).strip()
    if s.lower() in {"nan", "none", "nat", "<na>"}:
        return ""
    return s


def norm(v) -> str:
    return text(v).lower()


def slug(v) -> str:
    s = norm(v)
    s = re.sub(r"[^a-z0-9]+", "-", s).strip("-")
    return s or "unknown"


def clean_email(v) -> str:
    s = text(v).lower()
    if not s:
        return ""
    if "@" not in s or "." not in s.split("@")[-1]:
        return ""
    return s


def clean_phone(v) -> str:
    s = text(v)
    if not s:
        return ""
    digits = re.sub(r"[^\d+]", "", s)
    if len(re.sub(r"\D", "", digits)) < 8:
        return ""
    return digits


def clean_url(v) -> str:
    s = text(v)
    if not s:
        return ""
    if s.lower().startswith(("http://", "https://")):
        return s
    if "." in s:
        return f"https://{s}"
    return ""


def clean_cnpj(v) -> str:
    s = text(v)
    if not s:
        return ""
    digits = re.sub(r"\D", "", s)
    if len(digits) != 14:
        return ""
    return f"{digits[:2]}.{digits[2:5]}.{digits[5:8]}/{digits[8:12]}-{digits[12:]}"


def cnpj_digits(v) -> str:
    s = text(v)
    if not s:
        return ""
    digits = re.sub(r"\D", "", s)
    return digits if len(digits) == 14 else ""


def company_group_key(company: str, account_cnpj: str) -> str:
    digits = cnpj_digits(account_cnpj)
    if digits:
        return f"cnpj:{digits}"
    return f"name:{slug(company)}"


def uniq_keep_order(values: Iterable[str]) -> List[str]:
    out: List[str] = []
    seen = set()
    for item in values:
        s = text(item)
        k = norm(s)
        if not s or k in seen:
            continue
        seen.add(k)
        out.append(s)
    return out


def best_value(values: Iterable[str]) -> str:
    cleaned = [text(v) for v in values if text(v)]
    if not cleaned:
        return ""
    cleaned.sort(key=lambda x: (len(x), x), reverse=True)
    return cleaned[0]


def find_project_root() -> Path:
    here = Path(__file__).resolve().parent
    for p in [here, *here.parents]:
        if (p / "Apeiron_BR_Gestao_Comercial.xlsx").exists():
            return p
    return Path(__file__).resolve().parents[1]


def write_jsonl(path: Path, rows: Iterable[dict]) -> int:
    path.parent.mkdir(parents=True, exist_ok=True)
    count = 0
    with path.open("w", encoding="utf-8") as f:
        for row in rows:
            f.write(json.dumps(row, ensure_ascii=False) + "\n")
            count += 1
    return count


def chunk_rows(rows: List[dict], size: int) -> List[List[dict]]:
    if size <= 0:
        return [rows]
    return [rows[i : i + size] for i in range(0, len(rows), size)]


@dataclass
class Lead:
    id: int
    company: str
    company_group_key: str
    name: str
    role: str
    department: str
    email: str
    phone: str
    linkedin: str
    location: str
    status: str
    priority: str
    next_action: str
    lead_insight: str
    notes: str
    source: str
    account_summary: str
    account_sector: str
    account_city_state: str
    account_size: str
    account_revenue: str
    account_website: str
    account_phone: str
    account_cnpj: str


def load_leads(workbook_path: Path) -> List[Lead]:
    xls = pd.ExcelFile(workbook_path)
    ops = pd.read_excel(xls, sheet_name="Leads_Operations")
    det = pd.read_excel(xls, sheet_name="Leads_Details")

    ops["ID"] = pd.to_numeric(ops["ID"], errors="coerce")
    det["ID"] = pd.to_numeric(det["ID"], errors="coerce")
    ops = ops.dropna(subset=["ID"]).copy()
    det = det.dropna(subset=["ID"]).copy()
    ops["ID"] = ops["ID"].astype(int)
    det["ID"] = det["ID"].astype(int)

    merged = ops.merge(det, on="ID", how="outer", suffixes=("_ops", "_det"))
    merged["ID"] = pd.to_numeric(merged["ID"], errors="coerce").astype("Int64")
    merged = merged.dropna(subset=["ID"]).copy()
    merged["ID"] = merged["ID"].astype(int)

    out: List[Lead] = []
    for _, r in merged.sort_values("ID").iterrows():
        company = text(r.get("Company_ops")) or text(r.get("Company_det"))
        name = text(r.get("Full Name_ops")) or text(r.get("Full Name_det"))
        role = text(r.get("Role_ops")) or text(r.get("Role_det"))
        department = text(r.get("Department"))
        email = clean_email(r.get("Email"))
        phone = clean_phone(r.get("Phone"))
        linkedin = clean_url(r.get("LinkedIn"))
        location = text(r.get("Location"))
        status = text(r.get("Status_ops")) or text(r.get("Status_det"))
        priority = text(r.get("Priority_ops")) or text(r.get("Priority_det"))
        next_action = text(r.get("Next Action_ops")) or text(r.get("Next Action_det"))
        lead_insight = text(r.get("Lead Insight_ops")) or text(r.get("Lead Insight_det"))
        notes = text(r.get("Notes"))
        source = text(r.get("Source_ops")) or text(r.get("Source_det"))
        account_summary = text(r.get("Account_Summary"))
        account_sector = text(r.get("Account_Sector"))
        account_city_state = text(r.get("Account_City_State"))
        account_size = text(r.get("Account_Size"))
        account_revenue = text(r.get("Account_Revenue"))
        account_website = clean_url(r.get("Account_Website"))
        account_phone = clean_phone(r.get("Account_Phone"))
        account_cnpj = clean_cnpj(r.get("Account_CNPJ"))

        out.append(
            Lead(
                id=int(r["ID"]),
                company=company,
                company_group_key=company_group_key(company, account_cnpj),
                name=name,
                role=role,
                department=department,
                email=email,
                phone=phone,
                linkedin=linkedin,
                location=location,
                status=status,
                priority=priority,
                next_action=next_action,
                lead_insight=lead_insight,
                notes=notes,
                source=source,
                account_summary=account_summary,
                account_sector=account_sector,
                account_city_state=account_city_state,
                account_size=account_size,
                account_revenue=account_revenue,
                account_website=account_website,
                account_phone=account_phone,
                account_cnpj=account_cnpj,
            )
        )
    return out


def build_company_map(leads: List[Lead]) -> Dict[str, dict]:
    grouped: Dict[str, List[Lead]] = {}
    for lead in leads:
        grouped.setdefault(lead.company_group_key, []).append(lead)

    sorted_keys = sorted(grouped.keys())
    company_ids = {k: f"C{idx + 1:04d}" for idx, k in enumerate(sorted_keys)}
    out: Dict[str, dict] = {}
    for key in sorted_keys:
        items = grouped[key]
        out[key] = {
            "company_id": company_ids[key],
            "company": best_value([x.company for x in items]),
            "company_aliases": uniq_keep_order([x.company for x in items])[:4],
            "contacts_count": len(items),
            "known": {
                "account_sector": best_value([x.account_sector for x in items]),
                "account_city_state": best_value([x.account_city_state for x in items]),
                "account_website": best_value([x.account_website for x in items]),
                "account_phone": best_value([x.account_phone for x in items]),
                "account_size": best_value([x.account_size for x in items]),
                "account_revenue": best_value([x.account_revenue for x in items]),
                "account_summary": best_value([x.account_summary for x in items]),
                "account_cnpj": best_value([x.account_cnpj for x in items]),
            },
            "evidence": uniq_keep_order(
                [
                    x.account_summary
                    for x in items
                ]
                + [x.lead_insight for x in items]
                + [x.notes for x in items]
            )[:4],
        }
    return out


def build_company_queue(company_map: Dict[str, dict], include_complete: bool) -> List[dict]:
    required = [
        "account_sector",
        "account_city_state",
        "account_website",
        "account_phone",
        "account_size",
        "account_revenue",
        "account_summary",
        "account_cnpj",
    ]
    rows: List[dict] = []
    for _, item in sorted(company_map.items(), key=lambda x: x[1]["company_id"]):
        known = item["known"]
        missing = [k for k in required if not text(known.get(k))]
        if not include_complete and not missing:
            continue
        rows.append(
            {
                "task": "enrich_company",
                "company_id": item["company_id"],
                "company": item["company"],
                "company_aliases": item.get("company_aliases", []),
                "known": {k: v for k, v in known.items() if text(v)},
                "missing": missing,
                "evidence": item["evidence"],
                "contacts_count": item["contacts_count"],
            }
        )
    return rows


def build_contact_queue(leads: List[Lead], company_map: Dict[str, dict], include_complete: bool) -> List[dict]:
    required = ["role", "department", "email", "phone", "linkedin", "location"]
    rows: List[dict] = []
    for lead in leads:
        known = {
            "name": lead.name,
            "company": lead.company,
            "role": lead.role,
            "department": lead.department,
            "email": lead.email,
            "phone": lead.phone,
            "linkedin": lead.linkedin,
            "location": lead.location,
            "status": lead.status,
            "priority": lead.priority,
            "source": lead.source,
        }
        missing = [k for k in required if not text(known.get(k))]
        if not include_complete and not missing:
            continue

        company_known = company_map.get(lead.company_group_key, {}).get("known", {})
        company_item = company_map.get(lead.company_group_key, {})
        rows.append(
            {
                "task": "enrich_contact",
                "id": lead.id,
                "company_id": company_item.get("company_id", ""),
                "known": {k: v for k, v in known.items() if text(v)},
                "missing": missing,
                "company_context": {
                    "account_cnpj": text(company_known.get("account_cnpj")),
                    "account_website": text(company_known.get("account_website")),
                    "account_city_state": text(company_known.get("account_city_state")),
                    "account_sector": text(company_known.get("account_sector")),
                    "company_aliases": company_item.get("company_aliases", []),
                },
                "evidence": uniq_keep_order([lead.lead_insight, lead.notes, lead.next_action])[:3],
            }
        )
    return rows


def write_prompt_templates(out_dir: Path) -> None:
    company_prompt = """You will enrich company records from a queue with minimal tokens.
Rules:
- Return one JSON object per input line.
- Do not add prose.
- Keep original IDs.
- Fill only requested fields in `updates`.
- If unknown, omit the field.
- Add confidence from 0 to 1.
- If CNPJ is missing, prioritize finding `account_cnpj` first, then use it to validate other company fields.

Input line schema:
{task, company_id, company, company_aliases, known, missing, evidence, contacts_count}

Output line schema:
{"task":"enrich_company","company_id":"C0001","updates":{"account_sector":"","account_city_state":"","account_website":"","account_phone":"","account_size":"","account_revenue":"","account_summary":"","account_cnpj":""},"confidence":0.0,"reason":""}
"""
    contact_prompt = """You will enrich contact records from a queue with minimal tokens.
Rules:
- Return one JSON object per input line.
- Do not add prose.
- Keep original IDs.
- Fill only requested fields in `updates`.
- If unknown, omit the field.
- Add confidence from 0 to 1.

Input line schema:
{task, id, company_id, known, missing, company_context, evidence}

Output line schema:
{"task":"enrich_contact","id":1,"updates":{"role":"","department":"","email":"","phone":"","linkedin":"","location":""},"confidence":0.0,"reason":""}
"""
    (out_dir / "claude_prompt_company.txt").write_text(company_prompt, encoding="utf-8")
    (out_dir / "claude_prompt_contact.txt").write_text(contact_prompt, encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(description="Prepare low-token enrichment queues for Claude Coworker.")
    parser.add_argument("--workbook", type=Path, default=None, help="Path to Apeiron_BR_Gestao_Comercial.xlsx")
    parser.add_argument("--batch-size", type=int, default=25, help="Records per JSONL batch")
    parser.add_argument("--include-complete", action="store_true", help="Include records without missing fields")
    args = parser.parse_args()

    root = find_project_root()
    workbook = args.workbook or (root / "Apeiron_BR_Gestao_Comercial.xlsx")
    out_dir = root / "Support Files" / "_output" / "enrichment"
    out_dir.mkdir(parents=True, exist_ok=True)

    leads = load_leads(workbook)
    company_map = build_company_map(leads)
    company_rows = build_company_queue(company_map, include_complete=args.include_complete)
    contact_rows = build_contact_queue(leads, company_map, include_complete=args.include_complete)

    company_file = out_dir / "company_enrichment_queue.jsonl"
    contact_file = out_dir / "contact_enrichment_queue.jsonl"
    write_jsonl(company_file, company_rows)
    write_jsonl(contact_file, contact_rows)

    for kind, rows in [("company", company_rows), ("contact", contact_rows)]:
        batch_dir = out_dir / f"{kind}_batches"
        batch_dir.mkdir(parents=True, exist_ok=True)
        for f in batch_dir.glob("*.jsonl"):
            f.unlink()
        for idx, chunk in enumerate(chunk_rows(rows, args.batch_size), start=1):
            write_jsonl(batch_dir / f"{kind}_batch_{idx:03d}.jsonl", chunk)

    write_prompt_templates(out_dir)

    summary = {
        "workbook": str(workbook),
        "total_leads": len(leads),
        "total_companies": len(company_map),
        "company_queue": len(company_rows),
        "contact_queue": len(contact_rows),
        "batch_size": args.batch_size,
        "company_batches": len(chunk_rows(company_rows, args.batch_size)),
        "contact_batches": len(chunk_rows(contact_rows, args.batch_size)),
        "outputs": {
            "company_queue": str(company_file),
            "contact_queue": str(contact_file),
            "company_prompt": str(out_dir / "claude_prompt_company.txt"),
            "contact_prompt": str(out_dir / "claude_prompt_contact.txt"),
        },
    }
    (out_dir / "enrichment_queue_summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
