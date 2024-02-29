import pandas as pd

def replace_slash(x):
    if '/' in x:
        return x.split('/')[-1]
    else:
        return x
def process_data(file_name,sheet_name):
    # 读取 excel 文件
    result = pd.read_excel(file_name,skiprows=1, header=0, sheet_name=sheet_name)
    if sheet_name == 'Projection by operation':
        result['PRODUCT_DETAILED_NAME'] = result['PRODUCT_DETAILED_NAME'].replace(
            {'CimarronSATA ': 'CimarronSATA', 'CimarronSAS ': 'CimarronSAS'})
    # 判断 sheet_name 是否为 aa 或 bb
    if sheet_name in ['Projection by operation', 'CIMBP ']:
        # 只保留 PSG 和 ESG
        result = result[result.iloc[:, 0].isin(['PSG', 'ESG'])]
        # 保留我要的列
        # 判断 file_name 中是否包含 Q324
        if 'Q324' in file_name:
            # 如果包含，就用 Q324_need_columns 作为 sheet_name
            column_name = 'Q324_need_columns'
        elif'Q424' in file_name:
            # 如果不包含，就用 Q424_need_columns 作为 sheet_name
            column_name = 'Q424_need_columns'
        else:
            column_name = 'Q125_need_columns'
        # 用 sheet_name 的值来读取 excel 文件
        need_coulmns = pd.read_excel(r'./need_file/match.xlsx', sheet_name=column_name)
        need_coulmns = need_coulmns['columns'].tolist()
        result = result[need_coulmns]
        result = result.rename(columns={'LOCATION': 'LOC', 'MEASURE_TYPE': 'MEAS', 'SLOT_NAME': 'EQPT',
                            'OPER_NAME': 'IND_STEP', 'PRODUCT_DETAILED_NAME': 'DPN', 'CAPACITY': 'CAP'})
    else:
        # 保留我要的列
        if 'Q324' in file_name:
            # 如果包含，就用 Q324_need_columns 作为 sheet_name
            column_name = 'Q324_need_columns'
        elif 'Q424' in file_name:
            # 如果不包含，就用 Q424_need_columns 作为 sheet_name
            column_name = 'Q424_need_columns'
        else:
            column_name = 'Q125_need_columns'
        # 用 sheet_name 的值来读取 excel 文件
        need_coulmns = pd.read_excel(r'./need_file/match.xlsx', sheet_name=column_name)
        need_coulmns = need_coulmns['another_columns'].tolist()
        result = result[need_coulmns]
        result = result.rename(columns={'LOCATION': 'LOC', 'MEASURE_TYPE': 'MEAS', 'RESOURCE_NAME': 'EQPT',
    'OPER_NAME': 'IND_STEP', 'FACTORY_PRODUCT_NAME': 'DPN', 'CAPACITY': 'CAP','HEADS_NUM':'HEADS','DISCS_NUM':'DISCS'})
    """第一列列名为 LOCATION 取名为 LOC    
    第二列列名为 MEASURE_TYPE 取名为 MEAS,筛选列名为 MEAS 的值为 TT 和 YIELD,YIELD 改为 YD    
    第三列列名为 SLOT_NAME 取名为 EQPT,值为 SP 的改为 GEMINI_SP，IO 的改为 GEMINI_IO    
    新增第四列，列名改为 TEST_FLOW，并将该列的值改为 GG    
    第五列列名为 OPER_NAME 取名为 IND_STEP    
    第六列列名为 PRODUCT_DETAILED_NAME 取名为 DPN    
    第七列列名为 CAPACITY 取名为 CAP    
    第八第九列 HEADS DISCS 保留前 22 列"""
    result = result[result['MEAS'].isin(['TT', 'YIELD'])]
    result['MEAS'] = result['MEAS'].replace({'YIELD': 'YD'})
    result['EQPT'] = result['EQPT'].replace({'SP': 'GEMINI_SP', 'IO': 'GEMINI_IO'})
    result['TEST_FLOW'] = 'GG'
    result.insert(3,'TEST_FLOW',result.pop('TEST_FLOW'))
    result.columns = result.columns[:9].tolist() + ([f'2024_FW{col}' for col in range(27,40)] if 'Q324' in file_name
                                                else [f'2024_FW{col}' for col in range(40,53)] if 'Q424' in file_name
                                                else [f'2025_FW{col}' for col in range(1,14)])
    result['HEADS'] = result['HEADS'].astype(str).apply(replace_slash)
    result['DISCS'] = result['DISCS'].astype(str).apply(replace_slash)
    # 删除 IND_STEP 列一些不需要的步骤的行
    # 用 sheet_name 的值来读取 excel 文件
    noneed_step = pd.read_excel(r'./need_file/match.xlsx', sheet_name=column_name)
    noneed_step = noneed_step['noneed_step'].tolist()
    result = result[~result["IND_STEP"].isin(noneed_step)]
    # 读取 DPN,CAP,HEADS,DISCS(detail_C_H_D)，根据 detail_C_H_D 进行条件筛选，保留我们所需要的
    result['CAP'] = result['CAP'].astype(float).astype(int)
    result['HEADS'] = result['HEADS'].astype(float).astype(int)
    result['DISCS'] = result['DISCS'].astype(float).astype(int)
    detail_C_H_D = pd.read_excel(r'./need_file/match.xlsx', sheet_name='detail_C_H_D')
    # 创建一个空的 dataframe 对象 Q，用来存放筛选结果
    Q = pd.DataFrame()
    for index, row in detail_C_H_D.iterrows():
        # 用 no_need 的四列的值作为筛选条件
        condition = ((result['DPN'] == row['DPN']) & (result['CAP'] == row['CAP'])&
                     (result['HEADS'] == row['HEADS']) & (result['DISCS'] == row['DISCS']))
        # 筛选出 Q3 中满足条件的行
        filtered = result[condition]
        # 将筛选结果添加到 Q 中
        Q = Q._append(filtered)
    return Q
def concat_(file_name):
    pbo=process_data(file_name,'Projection by operation')
    cim=process_data(file_name,'CIMBP ')
    ebp=process_data(file_name,'EvansBP')
    lpk=process_data(file_name,'LongsPeak')
    Q=pd.concat([pbo,cim,ebp,lpk],axis=0,ignore_index=True)
    return Q
# Q=concat_()