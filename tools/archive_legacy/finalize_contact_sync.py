import os, pandas as pd, re

base = r'C:\Users\LucasMartinsRocha\OneDrive - Apeiron Pte Ltd\Prospecting'
wb = os.path.join(base, 'Apeiron_BR_Gestao_Comercial.xlsx')

master_sheet = 'CRM_Master'
new_contacts_file = os.path.join(base, '_output', 'contatos_final.csv')

master = pd.read_excel(wb, sheet_name=master_sheet)
new = pd.read_csv(new_contacts_file)

# Normalize key for master and new by contact_name + company_name
master['CONTACT_NAME'] = master.get('contact_name', '').fillna('').astype(str).str.lower().str.strip()
master['COMPANY_NAME'] = master.get('company_name', '').fillna('').astype(str).str.lower().str.strip()
master['UNIQUE_KEY'] = (master['CONTACT_NAME'] + '|' + master['COMPANY_NAME']).str.lower()

new['CONTACT_NAME'] = new['contact_name'].fillna('').astype(str).str.lower().str.strip()
new['COMPANY_NAME'] = new['company_name'].fillna('').astype(str).str.lower().str.strip()
new['UNIQUE_KEY'] = (new['CONTACT_NAME'] + '|' + new['COMPANY_NAME']).str.lower()

master_keys = set(master['UNIQUE_KEY'])
to_add = new[~new['UNIQUE_KEY'].isin(master_keys)].copy()

to_add['imported_from_zoho'] = True
if 'imported_from_zoho' not in master.columns:
    master['imported_from_zoho'] = False
else:
    master['imported_from_zoho'] = master['imported_from_zoho'].fillna(False)

if len(to_add) > 0:
    # align columns
    for c in ['contact_id','contact_name','email','phone','mobile','company_id','company_name','industry','website','status','source','source_type','zoho_account_id','imported_from_zoho']:
        if c not in master.columns:
            master[c] = ''
    required = ['contact_id','contact_name','email','phone','mobile','company_id','company_name','industry','website','status','source','source_type','zoho_account_id','imported_from_zoho']
    inserted = to_add[required]
    merged = pd.concat([master, inserted], ignore_index=True, sort=False)
else:
    merged = master

merged.to_excel(wb, sheet_name=master_sheet, index=False)

print('master rows before', len(master), 'new to add', len(to_add), 'master rows after', len(merged))

# write a report file
with open(os.path.join(base, '_output', 'sync_add_report.md'), 'w', encoding='utf-8') as f:
    f.write(f'# Sync add report\n')
    f.write(f'- master sheet: {master_sheet}\n')
    f.write(f'- before rows: {len(master)}\n')
    f.write(f'- inserted rows: {len(to_add)}\n')
    f.write(f'- after rows: {len(merged)}\n')
    f.write('- inserted sample:\n')
    f.write(to_add.head(5).to_csv(index=False))
