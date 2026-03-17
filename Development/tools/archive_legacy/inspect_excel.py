import pandas as pd, os
f='c:\\Users\\LucasMartinsRocha\\OneDrive - Apeiron Pte Ltd\\Prospecting\\Apeiron_BR_Gestao_Comercial.xlsx'
print('exists', os.path.exists(f))
xl=pd.ExcelFile(f)
print('sheets', xl.sheet_names)
for s in xl.sheet_names:
    df=pd.read_excel(xl, sheet_name=s)
    print('sheet', s, 'rows', len(df), 'cols', len(df.columns))
    print('columns:', df.columns.tolist())
    print('sample:', df.head(2).to_dict('records'))
