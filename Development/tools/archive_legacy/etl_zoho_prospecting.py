import os, shutil, datetime, re
import pandas as pd
from pathlib import Path

base = str(Path(__file__).resolve().parents[3])
backup_dir = os.path.join(base, '_backup')
os.makedirs(backup_dir, exist_ok=True)

src = os.path.join(base, 'Apeiron_BR_Gestao_Comercial.xlsx')
if not os.path.exists(src):
    raise FileNotFoundError(src)

dst = os.path.join(backup_dir, f'Apeiron_BR_Gestao_Comercial_backup_{datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")}.xlsx')
shutil.copy2(src, dst)
print('backup created', dst)

files = [os.path.join(backup_dir, f) for f in os.listdir(backup_dir) if f.lower().endswith('.xlsx')]
files.sort(key=os.path.getmtime, reverse=True)
for old in files[2:]:
    os.remove(old)
print('backups retained', [os.path.basename(f) for f in files[:2]])

xl = pd.ExcelFile(src)
print('sheets', xl.sheet_names)
metadata = {}
for s in xl.sheet_names:
    df = pd.read_excel(xl, sheet_name=s)
    metadata[s] = {
        'rows': len(df),
        'cols': len(df.columns),
        'columns': df.columns.tolist(),
        'sample': df.head(2).fillna('').to_dict('records')
    }
    print('sheet', s, 'rows', len(df), 'cols', len(df.columns))

zoho_dir = os.path.join(base, 'Zoho data')
zfiles = ['Accounts_2026_03_17.csv', 'Contacts_2026_03_17.csv', 'Leads_2026_03_17.csv']
zoho = {}
for f in zfiles:
    path = os.path.join(zoho_dir, f)
    if not os.path.exists(path):
        raise FileNotFoundError(path)
    zdf = pd.read_csv(path, encoding='utf-8', sep=',')
    zoho[f] = zdf
    print(f, 'rows', len(zdf), 'cols', len(zdf.columns))

trim = lambda s: s.strip() if isinstance(s, str) else s

acc = zoho['Accounts_2026_03_17.csv'].copy()
acc.rename(columns={
    'Record Id': 'account_id',
    'Account Name': 'company_name',
    'Account ID': 'account_id',
    'Email': 'email',
    'Phone': 'phone',
    'Website': 'website',
    'Industry': 'industry',
    'Billing City': 'city',
    'Billing State': 'state',
    'Billing Country': 'country',
    'Billing Code': 'zip',
}, inplace=True)
if 'account_id' not in acc and 'Record Id' in acc:
    acc['account_id'] = acc['Record Id']

for c in ['company_name', 'email', 'city', 'state', 'country', 'industry', 'website']:
    if c in acc:
        acc[c] = acc[c].apply(trim)

if 'email' in acc:
    acc['email'] = acc['email'].str.lower()

acc['zoho_id'] = acc.get('account_id')
acc['source'] = 'zoho_accounts'
empresas = acc[['account_id', 'company_name', 'industry', 'city', 'state', 'country', 'website', 'zoho_id', 'source']].drop_duplicates()
empresas['company_id'] = range(1, len(empresas) + 1)
company_id_map = dict(zip(empresas['zoho_id'], empresas['company_id']))

contacts = zoho['Contacts_2026_03_17.csv'].copy()
leads = zoho['Leads_2026_03_17.csv'].copy()

for df, typ in [(contacts, 'contact'), (leads, 'lead')]:
    df.rename(columns={
        'Full Name': 'contact_name',
        'Name': 'contact_name',
        'Contact Name': 'contact_name',
        'Email': 'email',
        'Phone': 'phone',
        'Mobile': 'mobile',
        'Account Name': 'company_name',
        'Account Name.id': 'zoho_account_id',
        'Account ID': 'zoho_account_id',
        'Record Id': 'contact_id',
        'Contact ID': 'contact_id',
        'Lead ID': 'contact_id',
        'Lead Source': 'source',
        'Status': 'status',
    }, inplace=True)
    if 'contact_id' not in df and 'Record Id' in df:
        df['contact_id'] = df['Record Id']
    if 'zoho_account_id' not in df and 'Account Name.id' in df:
        df['zoho_account_id'] = df['Account Name.id']
    for c in ['contact_name', 'email', 'company_name', 'status', 'source', 'phone', 'mobile']:
        if c in df:
            df[c] = df[c].apply(trim)
    if 'email' in df:
        df['email'] = df['email'].str.lower()
    df['origin_type'] = typ

cont = pd.concat([contacts, leads], ignore_index=True, sort=False)
cont['source_type'] = cont['origin_type'].fillna('contact')
cont['contact_id'] = cont.index + 1
cont['email'] = cont['email'].fillna('').apply(lambda s: s.lower().strip())
cont['phone'] = cont['phone'].fillna('').astype(str).apply(lambda s: re.sub(r'[^0-9+]', '', s))
cont['mobile'] = cont['mobile'].fillna('').astype(str).apply(lambda s: re.sub(r'[^0-9+]', '', s))
cont['company_id'] = cont['zoho_account_id'].map(company_id_map)

cont.sort_values(by=['email', 'phone', 'contact_name'], inplace=True)
cont = cont.drop_duplicates(subset=['email'], keep='first')
cont = cont.drop_duplicates(subset=['phone'], keep='first')
cont = cont.drop_duplicates(subset=['contact_name', 'company_name'], keep='first')

if 'status' not in cont:
    cont['status'] = ''

contatos_final = cont[['contact_id', 'contact_name', 'email', 'phone', 'mobile', 'company_id', 'company_name', 'status', 'source', 'source_type', 'zoho_account_id']].copy()
contatos_final['source'] = contatos_final['source'].fillna('zoho')

outdir = os.path.join(base, '_output')
os.makedirs(outdir, exist_ok=True)
empresas.to_csv(os.path.join(outdir, 'empresas_final.csv'), index=False)
contatos_final.to_csv(os.path.join(outdir, 'contatos_final.csv'), index=False)

with open(os.path.join(outdir, 'schema_summary.md'), 'w', encoding='utf-8') as f:
    f.write('# Schema Summary - Apeiron_BR_Gestao_Comercial\n\n')
    f.write(f'- backup created: {dst}\n')
    f.write('- sheets found:\n')
    for s, v in metadata.items():
        f.write(f'  - {s}: rows {v["rows"]}, cols {v["cols"]}\n')
        f.write(f'    columns: {v["columns"]}\n')

with open(os.path.join(outdir, 'data_quality_report.md'), 'w', encoding='utf-8') as f:
    f.write('# Data Quality Report\n\n')
    f.write('files read:\n')
    for n, df in zoho.items():
        f.write(f'- {n}: {len(df)} rows, {len(df.columns)} cols\n')
    f.write('\n- Duplicates by email/phone/name: resolved first-found.\n')

print('outputs written', os.path.join(outdir, 'empresas_final.csv'), os.path.join(outdir, 'contatos_final.csv'))
