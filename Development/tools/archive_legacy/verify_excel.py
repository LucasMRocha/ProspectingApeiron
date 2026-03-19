import pandas as pd
from pathlib import Path

base = Path(__file__).resolve().parents[3]
xl = pd.ExcelFile(base / 'Apeiron_BR_Gestao_Comercial.xlsx')
print('sheets', xl.sheet_names)
