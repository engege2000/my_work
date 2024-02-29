import pandas as pd
MS=pd.read_excel(r'./need_file/MS.xlsx',skiprows=3, header=0,sheet_name='MS')
need_columns=pd.read_excel(r'./need_file/MS.xlsx', sheet_name='need_columns')
#将其转换成我要的列表
need_columns=need_columns['need_columns'].values.tolist()
MS=MS.loc[:,need_columns]
MS=MS[(MS['Form Factor']=='3.5')&((MS['Fiscal Year Mth Name']=='FY2025JUL')|(MS['Fiscal Year Mth Name']=='FY2025AUG'))]
with pd.ExcelWriter(r'./need_file/MS.xlsx', engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    MS.to_excel(writer, sheet_name='filter_MS', index=False)