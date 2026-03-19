# Prospecting Dashboard (Zoho Sync)

Este repositório gerencia o pipeline de integração de dados Zoho -> dashboard local.

## Estrutura

- `src/dashboard/`: aplicação HTML/JS do dashboard de leads
- `tools/`: scripts de sync/backup/transformação
- `Zoho data/`: exportações brutas de Zoho (CSV)
- `_backup/`: backups automáticos (não versionado)
- `_output/`: dados ETL processados (não versionado)

## Setup inicial (executar uma vez)

```powershell
cd "<path-to-your-workspace>"
# inicializar repo (já feito):
# git init
# criar main/dev (já feito)
# git branch -M main && git branch -f dev main

# ajustar identidade (recomendado):
git config --global user.name "Lucas Martins Rocha"
git config --global user.email "seu-email@apeironengg.com"

# adicionar remoto (substitua URL):
# git remote add origin https://github.com/SEU_USUARIO/prospecting.git
# git push -u origin main
```

## Fluxo de trabalho sugerido

1. Crie branch de feature:
   `git checkout -b feature/zoho-sync`
2. Atualize scripts e dados.
3. Commit e teste local:
   `git add . && git commit -m "feat: atualiza pipeline Zoho"`
4. Merge para `main`:
   `git checkout main`
   `git merge feature/zoho-sync --no-ff`
5. Push ao remoto.

## Uso rápido do dashboard

Abrir `src/dashboard/index.html` no navegador.

## Comandos de ETL (exemplo)

```powershell
# 1) gerar staging a partir de CSV Zoho
python tools/etl_zoho_prospecting.py

# 2) atualizar Excel de importação / planejamento
python tools/update_excel.py

# 3) consolidar no CRM_Master
python tools/optimize_structure.py

# 4) verificar novos contatos e aplicar
python tools/calc_new_contacts.py
python tools/finalize_contact_sync.py
```

## Observações

- `.gitignore` exclui `*.xlsx`, `_backup/`, `_output/`, `dist/`, logs.
- Trabalhar com dados no `Zoho data/` e pipeline em `tools/`.
