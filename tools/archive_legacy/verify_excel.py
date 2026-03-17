import pandas as pd
xl = pd.ExcelFile(r'C:\Users\LucasMartinsRocha\OneDrive - Apeiron Pte Ltd\Prospecting\Apeiron_BR_Gestao_Comercial.xlsx')
print('sheets', xl.sheet_names)
