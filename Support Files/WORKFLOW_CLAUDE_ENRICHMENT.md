# Claude Coworker Enrichment Workflow (Low Token)

This workflow enriches the DB with minimal token usage.

## 1) Prepare queues (only gaps)

```powershell
python Development/tools/ensure_db_schema.py
python Development/tools/prepare_enrichment_queue.py --batch-size 20
```

Outputs:
- `Support Files/_output/enrichment/company_enrichment_queue.jsonl`
- `Support Files/_output/enrichment/contact_enrichment_queue.jsonl`
- batch files in:
  - `Support Files/_output/enrichment/company_batches/`
  - `Support Files/_output/enrichment/contact_batches/`
- prompt templates:
  - `Support Files/_output/enrichment/claude_prompt_company.txt`
  - `Support Files/_output/enrichment/claude_prompt_contact.txt`

Web-research queue (focused on CNPJ + company validation):

```powershell
python Development/tools/prepare_company_web_research_queue.py
```

Outputs:
- `Support Files/_output/enrichment/company_web_research_queue.jsonl`
- `Support Files/_output/enrichment/claude_prompt_company_web_research.txt`
- `Support Files/_output/enrichment/company_web_research_batches/company_web_batch_XXX.jsonl`
- `Support Files/_output/enrichment/CLAUDE_INPUT_company_web_batch_001.txt`

Recommended order:
1. Resolve missing `account_cnpj` from company name (web-research queue).
2. Run standard company/contact enrichment queues with CNPJ already known.

## 2) Send to Claude Coworker in batches

Use one batch file at a time.

### Company prompt
1. Paste `claude_prompt_company.txt`
2. Paste one `company_batch_XXX.jsonl`
3. Ask for JSONL output only
4. Save as:
   - `Support Files/_output/enrichment/company_enrichment_results.jsonl`

### Contact prompt
1. Paste `claude_prompt_contact.txt`
2. Paste one `contact_batch_XXX.jsonl`
3. Ask for JSONL output only
4. Save as:
   - `Support Files/_output/enrichment/contact_enrichment_results.jsonl`

## 3) Validate merge impact (dry-run)

```powershell
python Development/tools/merge_enrichment_results.py --dry-run --skip-build
```

For web-research result file (company-name keyed), use:

```powershell
python Development/tools/merge_enrichment_results.py --company-results "Support Files/_output/enrichment/company_web_research_results.jsonl" --dry-run --skip-build
```

## 4) Apply merge to DB

```powershell
python Development/tools/merge_enrichment_results.py
```

For web-research company result apply:

```powershell
python Development/tools/merge_enrichment_results.py --company-results "Support Files/_output/enrichment/company_web_research_results.jsonl"
```

Notes:
- Creates backup automatically and keeps only 2 versions.
- Updates workbook + `Development/src/dashboard/data/leads.js`.
- Rebuilds single-file dashboard (unless `--skip-build`).

## Token optimization rules

- Always process in small batches (`20` is a good default).
- Send only JSONL input and request JSONL output.
- Do not send full DB context on every prompt.
- Enrich company first, then contacts.
- Avoid overwrite unless needed (`--overwrite` only for correction rounds).
