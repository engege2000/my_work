import pandas as pd
# 定义一个函数，参数为文件名和工作表名
def process_data(filename, sheetname):
    # 读取 Excel 文件
    result = pd.read_excel(filename, sheet_name=sheetname)
    # 将第4行作为列名
    result.columns = result.iloc[3]
    result = result.rename(columns={'PROD MODEL':'Product Model'})
    # 保留我要的列
    if "Q3'24" in filename:
        columns = ['Product Model'] + list(range(27,40))
    elif "Q4'24" in filename:
        columns = ['Product Model'] + list(range(40,53))
    else:
        columns = ['Product Model'] + list(range(1,14))
    result = result[columns]
    # 删除前6行
    result = result.drop([0,1,2,3,4,5])
    # 将空格替换为NAN
    result = result.replace(" ", pd.NA)
    # 删掉列名为Product Model的值有缺失值的所有行
    result = result.dropna(subset=['Product Model'])
    result = result.reset_index(drop=True)
    # 读取映射表
    map = pd.read_excel('./middle_file/map.xlsx', sheet_name='mapPM')
    # 给result创建几列
    result['Design Application'] = None
    result['Form Factor'] = None
    result['Internal Product Name'] = None
    result['DPN'] = None
    result['Capacity'] = None
    result['Heads'] = None
    result['Discs'] = None
    result['MASS/LEGACY'] = None
    # 逐行读取result的列名为Product Model下的所有行
    for i in range(len(result)):
        # 拿出result第i行胡Product Model的值
        result_Product_Model = result.loc[i,'Product Model']
        # 将其与map列明为Product Model下的所有行进行对比
        for j in range(len(map)):
            # 拿出map第i行胡Product Model的值
            map_Product_Model = map.loc[j,'Product Model']
            if result_Product_Model == map_Product_Model:
                map_values = map.loc[j,['Design Application','Form Factor','Internal Product Name',
                                        'DPN','Capacity','Heads','Discs','MASS/LEGACY']]
                result.loc[i, ['Design Application','Form Factor','Internal Product Name',
                               'DPN','Capacity','Heads','Discs','MASS/LEGACY']] = map_values
                break
    # 更改列的位置
    cols = list(result.columns)
    cols = cols[-8:] + cols[:-8]
    result = result[cols]
    result = result.fillna(0)
    result = result.drop(result[result['Design Application']==0].index)
    # 这里由于把不存在的product改为了0，删掉
    # 返回处理后的数据
    return result
# 定义一个函数，参数为两个dataframe对象
def mergeadd(df1, df2,df3):
    # 合并两个dataframe对象，按照Product Model列
    df2 = df2[['Product Model'] + list(range(40,53))]
    df3 = df3[['Product Model'] + list(range(1,14))]
    result = pd.merge(df1, df2, on='Product Model')
    result = pd.merge(result, df3, on='Product Model')
    #删掉全为0的行
    result = result[~(result.iloc[:, -26:].sum(axis=1) == 0)]
    """result_col表示保留最后所需要的result的列名
    result表示保留最后需要的result的27-52周的具体数值
    把result_col和result联合起来的方法是通过CONCAT"""
    sep = ''
    result_col = result.drop_duplicates(subset=['DPN', 'Capacity', 'Heads', 'Discs'], keep='first')
    result_col['CONCAT'] = result_col.apply(lambda x: sep.join(x.loc['DPN':'Discs'].astype(str)), axis=1)
    #按照'DPN_x', 'Capacity_x', 'Heads_x', 'Discs_x'进行分组求和
    result['CONCAT'] = result.apply(lambda x: sep.join(x.loc['DPN':'Discs'].astype(str)), axis=1)
    result = result.groupby('CONCAT').sum()
    result = result.reset_index()
    result_col = result_col.reset_index()
    result = result[['CONCAT'] + list(range(27,53))+list(range(1,14))]
    result_col = result_col[['CONCAT', 'Design Application', 'Form Factor', 'Internal Product Name'
                            ,'DPN', 'Capacity', 'Heads', 'Discs', 'MASS/LEGACY']]
    result = pd.merge(result, result_col, on='CONCAT')
    result = result.drop(columns=['CONCAT'], axis=1)
    result=result[['Design Application', 'Form Factor', 'Internal Product Name'
                            ,'DPN', 'Capacity', 'Heads', 'Discs', 'MASS/LEGACY']+list(range(27,53))+list(range(1,14))]
    result = result.sort_values(by=['DPN', 'Capacity', 'Heads', 'Discs'])
    # 返回处理后的数据
    return result
def MS_WK(wx_sheet_name,k_sheet_name):
    wx_thir_output=pd.read_excel(r'./middle_file/result.xlsx',sheet_name=wx_sheet_name)
    sep = ''
    wx_thir_output['CONCAT'] = wx_thir_output.apply(lambda x: sep.join(x.loc['DPN':'DISCS'].astype(str)), axis=1)
    wx_ms_jul=pd.read_excel(r'./middle_file/add_result.xlsx',sheet_name='wx_jul')
    wx_ms_jul=wx_ms_jul[['CONCAT','JUL']]
    wx_output=pd.merge(wx_thir_output,wx_ms_jul,on='CONCAT')
    wx_output['FW2501']=(wx_output['JUL']/4).round(0).astype(int)
    wx_output['FW2502']=(wx_output['JUL']/4).round(0).astype(int)
    wx_output['FW2503']=(wx_output['JUL']/4).round(0).astype(int)
    wx_output['FW2504']=(wx_output['JUL']/4).round(0).astype(int)
    wx_output=wx_output.drop(columns=['CONCAT','JUL'],axis=1)
    k_thir_output=pd.read_excel(r'./middle_file/result.xlsx',sheet_name=k_sheet_name)
    k_thir_output['CONCAT'] = k_thir_output.apply(lambda x: sep.join(x.loc['DPN':'DISCS'].astype(str)), axis=1)
    k_ms_jul=pd.read_excel(r'./middle_file/add_result.xlsx',sheet_name='k_jul')
    k_ms_aug=pd.read_excel(r'./middle_file/add_result.xlsx',sheet_name='k_aug')
    k_ms_jul=k_ms_jul[['CONCAT','JUL']]
    k_ms_aug=k_ms_aug[['CONCAT','AUG']]
    k_output=pd.merge(k_thir_output,k_ms_jul,on='CONCAT')
    k_output=pd.merge(k_output,k_ms_aug,on='CONCAT')
    k_output[[i for i in range(27,51)]]=k_output[[i for i in range(29,53)]]
    k_output[51]=(k_output['JUL']/4).round(0).astype(int)
    k_output[52]=(k_output['JUL']/4).round(0).astype(int)
    k_output['FW2501']=(k_output['JUL']/4).round(0).astype(int)
    k_output['FW2502']=(k_output['JUL']/4).round(0).astype(int)
    k_output['FW2503']=(k_output['AUG']/4).round(0).astype(int)
    k_output['FW2504']=(k_output['AUG']/4).round(0).astype(int)
    k_output=k_output.drop(columns=['CONCAT','JUL','AUG'],axis=1)
    output=pd.concat([wx_output,k_output],axis=0,ignore_index=True)
    return output
def filter_and_process(loc_code):
    MS = pd.read_excel(r'./need_file/MS.xlsx', sheet_name='filter_MS')
    need_columns = pd.read_excel(r'./need_file/MS.xlsx', sheet_name='need_columns')
    # 将其转换成我要的列表
    need_columns = need_columns['need_columns'].values.tolist()
    result = MS[MS['Location Code'] == loc_code]
    result = result.reindex(columns=need_columns)
    sep = ''
    result['CONCAT'] = result.apply(lambda x: sep.join(x[['Detailed Product Name', 'Capacity',
                                                          'Heads', 'Discs', 'Fiscal Year Mth Name']].astype(str)),
                                    axis=1)
    result = (result.groupby('CONCAT', as_index=False).agg
              ({'Location Code': 'first', 'Design Application': 'first', 'Form Factor': 'first',
                'Internal Product Name': 'first', 'Detailed Product Name': 'first',
                'Capacity': 'first', 'Heads': 'first', 'Discs': 'first', 'Cur MS Qty': 'sum', 'CONCAT': 'first'}))
    result['JUL'] = None
    mask = result['CONCAT'].str.contains('JUL')
    result.loc[mask, 'JUL'] = result.loc[mask, 'Cur MS Qty']
    result['AUG'] = None
    mask = result['CONCAT'].str.contains('AUG')
    result.loc[mask, 'AUG'] = result.loc[mask, 'Cur MS Qty']
    result_JUL = result.drop(columns=['AUG'], axis=1)
    result_JUL = result_JUL.dropna()
    result_AUG = result.drop(columns=['JUL'], axis=1)
    result_AUG = result_AUG.dropna()
    result_JUL['CONCAT'] = result_JUL.apply(lambda x: sep.join(x[['Detailed Product Name', 'Capacity',
                                                                  'Heads', 'Discs']].astype(str)), axis=1)
    result_AUG['CONCAT'] = result_AUG.apply(lambda x: sep.join(x[['Detailed Product Name', 'Capacity',
                                                                  'Heads', 'Discs']].astype(str)), axis=1)
    # 返回处理后的数据框
    return result_JUL,result_AUG