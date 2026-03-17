import pandas as pd
from openpyxl import load_workbook

base = r'C:\Users\LucasMartinsRocha\OneDrive - Apeiron Pte Ltd\Prospecting'
file_path = base + '\\Apeiron_BR_Gestao_Comercial.xlsx'

empresas = pd.read_csv(base + '\\_output\\empresas_final.csv')
contatos = pd.read_csv(base + '\\_output\\contatos_final.csv')

with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    empresas.to_excel(writer, sheet_name='ToImport_Empresas', index=False)
    contatos.to_excel(writer, sheet_name='ToImport_Contatos', index=False)

print('Excel updated with ToImport_Empresas and ToImport_Contatos')
