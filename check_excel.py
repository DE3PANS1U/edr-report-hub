import pandas as pd

xl = pd.ExcelFile('combined.xlsx')
print('Sheet names:', xl.sheet_names)

for sheet_name in xl.sheet_names:
    df = pd.read_excel('combined.xlsx', sheet_name=sheet_name)
    print(f'\nSheet: {sheet_name}')
    print(f'Columns: {df.columns.tolist()}')
    print(f'Shape: {df.shape}')
    print('First few rows:')
    print(df.head())
