import pandas as pd, os, re
from pathlib import Path

base = str(Path(__file__).resolve().parents[3])
# Try CRM_Master first, then fallback to legacy sheet
try:
    cont_existing = pd.read_excel(os.path.join(base,'Apeiron_BR_Gestao_Comercial.xlsx'), sheet_name='CRM_Master')
except Exception:
    cont_existing = pd.read_excel(os.path.join(base,'Apeiron_BR_Gestao_Comercial.xlsx'), sheet_name='ðŸ“‹ Contacts')
cont_new = pd.read_csv(os.path.join(base,'_output','contatos_final.csv'))

def key(df, name_col, company_col):
    names = df[name_col].fillna('').astype(str).str.lower().str.strip()
    companies = df[company_col].fillna('').astype(str).str.lower().str.strip()
    return (names + '|' + companies).str.lower()

# Use contact_name + company_name dedupe
if 'Contact Name' in cont_existing.columns and 'Account Name' in cont_existing.columns:
    existing_keys = set(key(cont_existing, 'Contact Name', 'Account Name'))
else:
    existing_keys = set(cont_existing.index.astype(str))

new_keys = key(cont_new, 'contact_name', 'company_name')
new_to_add = cont_new[~new_keys.isin(existing_keys)]

print('existing contacts rows', len(cont_existing))
print('transformed contact source rows', len(cont_new))
print('new unique rows (not matching existing by email/phone/name):', len(new_to_add))
print('new_to_add sample count', min(5,len(new_to_add)))
print(new_to_add.head(5).to_dict('records'))
