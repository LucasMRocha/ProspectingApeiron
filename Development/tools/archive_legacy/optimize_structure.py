import pandas as pd, os
from pathlib import Path

base = str(Path(__file__).resolve().parents[3])
wb = os.path.join(base, 'Apeiron_BR_Gestao_Comercial.xlsx')

contatos = pd.read_csv(os.path.join(base, '_output', 'contatos_final.csv'))
empresas = pd.read_csv(os.path.join(base, '_output', 'empresas_final.csv'))

# Create optimized master table with both contact and company details
merged = contatos.merge(empresas[['company_id', 'company_name', 'industry', 'website']], on='company_id', how='left', suffixes=('','_company'))

# Normalize selected columns (drop wide columns)
master = merged[['contact_id', 'contact_name', 'email', 'phone', 'mobile', 'company_id', 'company_name', 'industry', 'website', 'status', 'source', 'source_type', 'zoho_account_id']].copy()

# Write as new optimized sheet
with pd.ExcelWriter(wb, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    master.to_excel(writer, sheet_name='CRM_Master', index=False)

print('CRM_Master sheet created with', len(master), 'rows and', len(master.columns), 'columns')

# optional: dump schema summary of new sheet to file
with open(os.path.join(base, '_output', 'structure_improvement_suggestions.md'), 'a', encoding='utf-8') as f:
    f.write('\n\n## AÃ§Ã£o realizada: estrutura otimizada\n')
    f.write('- Added sheet `CRM_Master` with 13 columns (contact/company consolidated).\n')
    f.write('- Kept original tabs as historical sources.\n')
