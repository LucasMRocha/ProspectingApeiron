import pandas as pd, os
from pathlib import Path

base = str(Path(__file__).resolve().parents[3])
wb = os.path.join(base, 'Apeiron_BR_Gestao_Comercial.xlsx')
contacts_sheet = 'ðŸ“‹ Contacts'

# read existing contacts sheet
xl = pd.ExcelFile(wb)
if contacts_sheet not in xl.sheet_names:
    raise ValueError(f"Sheet {contacts_sheet} not found")
cont_existing = pd.read_excel(xl, sheet_name=contacts_sheet)
print('existing contacts cols', cont_existing.columns.tolist()[:10], 'rows', len(cont_existing))

# load transformed contacts
cont_transformed = pd.read_csv(os.path.join(base, '_output', 'contatos_final.csv'))
print('transformed contacts rows', len(cont_transformed))

# For this layout, we create a dedicated sheet with Zoho-transformed contacts.
cont_transformed['CONTACT_EMAIL'] = cont_transformed['email'].astype(str).str.lower().str.strip()
cont_transformed['CONTACT_PHONE'] = cont_transformed['phone'].astype(str).str.replace(r'[^0-9+]', '', regex=True)
cont_transformed['CONTACT_NAME'] = cont_transformed['contact_name'].astype(str).str.strip()
# unique key
cont_transformed['UNIQUE_KEY'] = (cont_transformed['CONTACT_EMAIL'] + '|' + cont_transformed['CONTACT_PHONE'] + '|' + cont_transformed['CONTACT_NAME']).str.lower()

# do not dedupe here - we keep complete transformed as source of truth
new_rows = cont_transformed.copy()
print('new rows written to new sheet', len(new_rows))

# write transformed to dedicated new sheet
with pd.ExcelWriter(wb, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    cont_transformed.to_excel(writer, sheet_name='ðŸ“‹ Contacts_From_Zoho', index=False)

# rename existing sheet to preserve USABILITY if needed
# (not altering old sheet contents there,
# but new contacts are in a clear ingestion sheet)

print('new sheet created with transformed contacts rows', len(new_rows))
