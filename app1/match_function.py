import pandas as pd
def match(Q,sheet_name):
    match_data= pd.read_excel(r'./need_file/match.xlsx', sheet_name=sheet_name)
    match_list = match_data.apply(lambda x: ([x["DPN2"], x["CAP2"], x["HEADS2"], x["DISCS2"]],
                                     [x["DPN1"], x["CAP1"], x["HEADS1"], x["DISCS1"]]), axis=1).tolist()
    for t in match_list:
        match = (Q["DPN"] == t[0][0]) & ((Q['CAP']) == t[0][1]) & (Q["HEADS"] == t[0][2]) & (Q["DISCS"] == t[0][3])
        Q.loc[match, ["DPN", "CAP", "HEADS", "DISCS"]] = t[1]
    return Q
def save(Q,sheet_name):
    with pd.ExcelWriter(r'./middle_file/result.xlsx', engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        Q.to_excel(writer, sheet_name=sheet_name, index=False)
class QUpdater: # Initialize method, accepts one argument: q1_file
    def __init__(self):
        # Read the Q1_file file and save it as a property
        self.Q1 = pd.read_csv('Q1.csv')
    def merge_Q(self, Q2_sheet_name, match_sheet_name):
        # Read Q2_file and match_file files
        Q2 = pd.read_excel( r'./middle_file/result.xlsx',sheet_name=Q2_sheet_name)
        match2 = pd.read_excel(r'./need_file/match.xlsx',sheet_name=match_sheet_name)
        # For each row of match2, find the same DPN1,CAP1,HEADS1,DISCS1 rows from Q2 and add them to the end of Q1
        for i in range(len( match2)):
            row = match2.iloc[i] # Fetch row i of match2
            Q2_matched = Q2[(Q2['DPN'] == row['DPN1']) & (Q2['CAP'] == row['CAP1']) & (Q2['HEADS'] == row['HEADS1' ]) & (Q2['DISCS'] == row['DISCS1']) & (Q2['MEAS'] == row['MEAS1'])] # Find identical rows from Q2
            self.Q1 = pd.concat([self.Q1, Q2_matched], ignore_index=True) # Add the Q2_matched to the end of Q1
        # Sort Q1 by the specified columns
        self.Q1 = self.Q1.sort_values(by=['DPN','MEAS','CAP','HEADS','DISCS'])
        # Return Q1 as a dataframe
        self.Q1.to_csv('Q1.csv',index=False)
def step_add_map(sheet_name):
    # 导入pandas库
    import pandas as pd
    # 读取excel文件的sheet为STEP
    step_map = pd.read_excel(r'./need_file/match.xlsx', sheet_name='step_map')
    # 读取列名为DPN下的值，转换成一个列表
    dpn = step_map['DPN'].tolist()
    # 读取第4列到第14列的值，转换成一个DataFrame
    cols = step_map.iloc[:, 3:14]
    # 创建一个空字典，用来存储DPN和对应的替换规则
    dic = {}
    # 遍历DPN列表的每个元素
    for i in range(len(dpn)):
        # 把DPN列表的第i个元素赋值给key变量
        key = dpn[i]
        # 把cols的第i行的非空值赋值给value变量，转换成一个Series
        value = cols.iloc[i][cols.iloc[i].notnull()]
        # 把value的每个元素加1，转换成一个列表
        value = [j + 1 for j, _ in enumerate(value)]
        # 把cols的第i行的值赋值给cols_list变量，转换成一个列表
        cols_list = list(cols.iloc[i])
        # 把cols_list中的空值去掉，保留非空值
        cols_list = [x for x in cols_list if not pd.isnull(x)]
        # 把key和一个由cols_list和value组成的列表作为键值对，添加到dic字典中
        dic[key] = (cols_list, value)
    # 读取csv文件
    Q1 = pd.read_csv('Q1.csv')
    # 遍历dic字典的每个键值对
    for key, value in dic.items():
        # 找到Q中DPN等于key的行的索引
        index = Q1[Q1['DPN'] == key].index
        # 把这些行的IND_STEP列的值，根据value的第一个元素和第二个元素的对应关系，进行替换
        Q1.loc[index, 'IND_STEP'] = Q1.loc[index, 'IND_STEP'].replace(value[0], value[1])
    # 读取excel文件的sheet为STEP
    add_map = pd.read_excel(r'./need_file/match.xlsx', sheet_name='add_map')
    # 读取列名为DPN下的值，转换成一个列表
    dpn_list = add_map['DPN'].tolist()
    # 把IND_STEP转换为整数类型
    Q1['IND_STEP'] = Q1['IND_STEP'].astype(int, errors='ignore')
    # 重置索引，从0开始
    Q1 = Q1.reset_index(drop=True)
    # 缺失值填充
    Q1 = Q1.fillna(0)
    # 从后往前遍历每一行
    for i in range(len(Q1) - 1, -1, -1):
        # 读取DPN的值
        dpn = Q1.loc[i, 'DPN']
        # 读取IND_STEP的值
        ind_step = Q1.loc[i, 'IND_STEP']
        # 如果DPN是列表的前13个的其中一个，且IND_STEP为7
        if dpn in dpn_list[:13] and ind_step == 7:
            # 把这一行的第10列到最后一列的值加到上一行对应的位置
            Q1.iloc[i - 1, 9:] = Q1.iloc[i - 1, 9:] + Q1.iloc[i, 9:]
            # 删除这一行
            Q1 = Q1.drop(i)
        # 如果DPN是列表的第14个和第15个的其中一个，且IND_STEP为5
        elif dpn in dpn_list[13:15] and ind_step == 5:
            # 把这一行的第10列到最后一列的值加到上一行对应的位置
            Q1.iloc[i - 1, 9:] = Q1.iloc[i - 1, 9:] + Q1.iloc[i, 9:]
            # 删除这一行
            Q1 = Q1.drop(i)
        # 如果DPN是列表的第16个到第20个的其中一个，且IND_STEP为8
        elif dpn in dpn_list[15:20] and ind_step == 8:
            # 把这一行的第10列到最后一列的值加到上一行对应的位置
            Q1.iloc[i - 1, 9:] = Q1.iloc[i - 1, 9:] + Q1.iloc[i, 9:]
            # 删除这一行
            Q1 = Q1.drop(i)
        # 如果DPN是列表的第21，22，23的其中一个，且IND_STEP为10
        elif dpn in dpn_list[20:23] and ind_step == 10:
            # 把这一行的第10列到最后一列的值加到上一行对应的位置
            Q1.iloc[i - 1, 9:] = Q1.iloc[i - 1, 9:] + Q1.iloc[i, 9:]
            # 删除这一行
            Q1 = Q1.drop(i)
    # 重新重置索引，从0开始
    Q1 = Q1.reset_index(drop=True)
    # 读取PSES3.5.csv文件的DPN1列，并转换成一个列表
    ALLMAP = pd.read_excel(r'./need_file/match.xlsx', sheet_name='PSE3.5')
    ALLMAP_list = ALLMAP['DPN1'].tolist()
    # 筛选Q_updated中DPN在ALLMAP_list中的行
    Q1 = Q1[Q1['DPN'].isin(ALLMAP_list)]
    # 添加一个EQPT列，赋值为GEMINI_SP
    Q1['EQPT'] = 'GEMINI_SP'
    # 找到Q_updated中DPN为CIMARRONBPSAS或CIMARRONBPSATA，且IND_STEP为6或7的行的索引
    index = Q1[(Q1['DPN'].isin(['CIMARRONBPSAS', 'CIMARRONBPSATA']))
               & (Q1['IND_STEP'].isin([6, 7]))].index
    # 把这些行的EQPT列的值改为GEMINI_IO
    Q1.loc[index, 'EQPT'] = 'GEMINI_IO'
    # 找到Q_updated中DPN为HEPBURNOASIS或PHARAOHOASIS，且IND_STEP为4或5的行的索引
    index = Q1[(Q1['DPN'].isin(['HEPBURNOASIS', 'PHARAOHOASIS']))
               & (Q1['IND_STEP'].isin([4, 5]))].index
    # 把这些行的EQPT列的值改为GEMINI_IO
    Q1.loc[index, 'EQPT'] = 'GEMINI_IO'
    # 把CAP和IND_STEP列转换为数值类型
    Q1['CAP'] = pd.to_numeric(Q1['CAP'], errors='coerce')
    Q1['IND_STEP'] = pd.to_numeric(Q1['IND_STEP'], errors='coerce')
    # 打开已经存在的excel文件，假设文件名是result.xlsx
    with pd.ExcelWriter(r'./middle_file/result.xlsx', engine='openpyxl', mode='a',if_sheet_exists='replace') as writer:
        # 将Q_updated写入到一个新的工作表，假设工作表名是PSES3.5_new1
        Q1.to_excel(writer, sheet_name=sheet_name, index=False)
def merge_():
    Q324_PSE3_5=pd.read_excel(r'./middle_file/result.xlsx',sheet_name='Q324_PSE3.5')
    Q424_PSE3_5=pd.read_excel(r'./middle_file/result.xlsx',sheet_name='Q424_PSE3.5')
    Q125_PSE3_5 = pd.read_excel(r'./middle_file/result.xlsx', sheet_name='Q125_PSE3.5')
    sep = ''
    Q324_PSE3_5['CONCAT'] = Q324_PSE3_5.apply(lambda x: sep.join(x.loc['LOC':'DISCS'].astype(str)), axis=1)
    Q424_PSE3_5['CONCAT'] = Q424_PSE3_5.apply(lambda x: sep.join(x.loc['LOC':'DISCS'].astype(str)), axis=1)
    Q125_PSE3_5['CONCAT'] = Q125_PSE3_5.apply(lambda x: sep.join(x.loc['LOC':'DISCS'].astype(str)), axis=1)
    Q424_PSE3_5=Q424_PSE3_5[['CONCAT']+[f'2024_FW{col}' for col in range(40,53)]]
    Q125_PSE3_5 = Q125_PSE3_5[['CONCAT'] + [f'2025_FW{col}' for col in range(1, 14)]]
    Q_result=pd.merge(Q324_PSE3_5,Q424_PSE3_5,on='CONCAT')
    Q_result = pd.merge(Q_result, Q125_PSE3_5, on='CONCAT')
    Q_result=Q_result.drop(columns=['Unnamed: 0','CONCAT'])
    return Q_result
#Q=concat_()
# Q1=match(Q,'match1')
# save(Q1,'Q1')