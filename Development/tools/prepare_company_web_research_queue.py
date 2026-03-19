#!/usr/bin/env python3
from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Dict, List

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


def slug(v: str) -> str:
    s = norm(v)
    s = re.sub(r"[^a-z0-9]+", "-", s).strip("-")
    return s or "unknown"


def cnpj_digits(v) -> str:
    s = text(v)
    if not s:
        return ""
    digits = re.sub(r"\D", "", s)
    return digits if len(digits) == 14 else ""


def clean_cnpj(v) -> str:
    digits = cnpj_digits(v)
    if not digits:
        return ""
    return f"{digits[:2]}.{digits[2:5]}.{digits[5:8]}/{digits[8:12]}-{digits[12:]}"


def company_group_key(company: str, account_cnpj: str) -> str:
    digits = cnpj_digits(account_cnpj)
    if digits:
        return f"cnpj:{digits}"
    return f"name:{slug(company)}"


def find_project_root() -> Path:
    here = Path(__file__).resolve().parent
    for p in [here, *here.parents]:
        if (p / "Apeiron_BR_Gestao_Comercial.xlsx").exists():
            return p
    return Path(__file__).resolve().parents[1]


def suspicious_website(v: str) -> bool:
    s = norm(v)
    if not s:
        return True
    if "linkedin.com" in s:
        return True
    if "manual search" in s:
        return True
    if re.search(r"\+?\d[\d\s\-().]{7,}", s):
        return True
    return not bool(re.search(r"\.[a-z]{2,}", s))


def city_state_ok(v: str) -> bool:
    s = text(v)
    if not s:
        return False
    uf = "AC|AL|AP|AM|BA|CE|DF|ES|GO|MA|MT|MS|MG|PA|PB|PR|PE|PI|RJ|RN|RS|RO|RR|SC|SP|SE|TO"
    if re.fullmatch(rf"(?:{uf})", s, re.I):
        return True
    if re.search(rf"^[^/]+/\s*(?:{uf})$", s, re.I):
        return True
    if re.search(rf"^.+,\s*(?:{uf})$", s, re.I):
        return True
    return False


def main() -> None:
    root = find_project_root()
    workbook = root / "Apeiron_BR_Gestao_Comercial.xlsx"
    out_dir = root / "Support Files" / "_output" / "enrichment"
    out_dir.mkdir(parents=True, exist_ok=True)

    xls = pd.ExcelFile(workbook)
    det = pd.read_excel(xls, sheet_name="Leads_Details").dropna(subset=["ID"]).copy()

    grouped: Dict[str, List[dict]] = {}
    for _, row in det.iterrows():
        company = text(row.get("Company"))
        if not company:
            continue
        key = company_group_key(company, text(row.get("Account_CNPJ")))
        grouped.setdefault(key, []).append(row.to_dict())

    rows: List[dict] = []
    for key, group_rows in grouped.items():
        companies = [text(r.get("Company")) for r in group_rows if text(r.get("Company"))]
        companies = [c for i, c in enumerate(companies) if norm(c) not in {norm(x) for x in companies[:i]}]
        primary_company = companies[0] if companies else ""

        def first(col: str) -> str:
            vals = [text(r.get(col)) for r in group_rows if text(r.get(col))]
            return vals[0] if vals else ""

        item = {
            "task": "discover_cnpj_then_enrich_company",
            "company": primary_company,
            "company_aliases": companies[:5],
            "known": {
                "account_cnpj": clean_cnpj(first("Account_CNPJ")),
                "account_website": first("Account_Website"),
                "account_city_state": first("Account_City_State"),
                "account_sector": first("Account_Sector"),
                "account_phone": first("Account_Phone"),
            },
            "contact_samples": [
                {
                    "id": int(text(r.get("ID")) or 0),
                    "name": text(r.get("Full Name")),
                    "role": text(r.get("Role")),
                    "email": text(r.get("Email")),
                }
                for r in group_rows[:3]
            ],
        }
        flags = []
        if not text(item["known"]["account_cnpj"]):
            flags.append("missing_cnpj")
        if suspicious_website(item["known"]["account_website"]):
            flags.append("website_needs_validation")
        if not city_state_ok(item["known"]["account_city_state"]):
            flags.append("city_state_needs_normalization")
        if not text(item["known"]["account_sector"]):
            flags.append("missing_sector")
        if not text(item["known"]["account_phone"]):
            flags.append("missing_company_phone")
        if not flags:
            continue
        item["research_flags"] = flags
        item["group_key"] = key
        rows.append(item)

    queue_path = out_dir / "company_web_research_queue.jsonl"
    with queue_path.open("w", encoding="utf-8") as f:
        for row in rows:
            f.write(json.dumps(row, ensure_ascii=False) + "\n")

    prompt = """You are enriching Brazilian companies for CRM data quality.
Return JSONL only, one output line per input line.
Use reliable public sources (official company site, Receita Federal references, trusted business registries).
Do not invent values.

Input line schema:
{task, company, company_aliases, known, contact_samples, research_flags, group_key}

Output line schema:
{"company":"","updates":{"account_cnpj":"","account_website":"","account_city_state":"","account_sector":"","account_phone":""},"confidence":0.0,"sources":[""],"notes":""}
"""
    (out_dir / "claude_prompt_company_web_research.txt").write_text(prompt, encoding="utf-8")

    print(
        json.dumps(
            {
                "workbook": str(workbook),
                "queue_path": str(queue_path),
                "prompt_path": str(out_dir / "claude_prompt_company_web_research.txt"),
                "companies_in_queue": len(rows),
            },
            ensure_ascii=False,
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
