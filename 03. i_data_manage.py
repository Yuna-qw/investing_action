
import util.aSetting as aSetting
from openpyxl import load_workbook

Update = aSetting.Update
EndDate_Str = aSetting.EndDate_Str_1
MasterSheet = aSetting.MasterSheet
MasterOrigin = aSetting.MasterOrigin
BasePath = aSetting.BasePath
RawDataPath = aSetting.RawDataPath
LogPath = aSetting.LogPath
FinalDataPath = aSetting.FinalDataPath

'''
copy over i_management
'''

Name_clean = ('combine_clean_' + str(Update))
Name_I = ('combine_i_' + str(Update))

file1 = (FinalDataPath + Name_clean + '.xlsx')
print("file1的路径是：", file1)
file2 = MasterSheet
file3 = (FinalDataPath + Name_I + '.xlsx')
wbTarget = load_workbook(file1)
wb2 = load_workbook(file2)
WS = wb2.sheetnames  # cat: change from get_sheet_names
select_sheet = ['i_data', 'i_event', 'i_extremeValue', 'i_yield_price']
copyWSList = [x for x in WS if x in select_sheet]

for sheet in copyWSList:
    ws = wbTarget.create_sheet(sheet)
    for row in wb2[sheet]:
        for cell in row:
            ws[cell.coordinate].value = cell.value
wbTarget.save(file3)

'''
create i_indicator 
'''

I_Data = load_workbook(file3)


# cat: all the sheet name, it is a list
# 获取所有工作表名称和对应的工作表对象
sheetname_list = I_Data.sheetnames            # cat: change from get_sheet_names
data_list = [I_Data[x] for x in sheetname_list]   # cat: all the sheet name, it is a list
# 提取“清洗后数据表”的首行（列名）
first_row = list(data_list[0].rows)[0]			# cat: first sheet 'original': first row;
select_col = copyWSList
col_list_originalData = [i.value for i in first_row]      # cat: get all the columns name in sheet 'original'
name_index = [(j, col_list_originalData.index(i)) for i, j in zip(select_col, data_list[1:])]

# 定义关键列
# 主数据表中用于匹配的关键列（如类别、文件路径、日期）
key_col = ['Subcategory', 'FilePath', 'Date']  # select key and useful column at data sheet
# 指标工作表中用于匹配的关键列（如类别、文件路径、规则生效起止日期）
key_col_EachSheet = ['Subcategory', 'FilePath', 'Issue_StartDate', 'Issue_EndDate']  # select key and useful column at indicator sheet
# 记录关键列在主数据表中的索引（列位置）
key_index = [(j, col_list_originalData.index(i)) for i, j in zip(key_col, key_col)]
# print(key_col)
keyList = []  # find out # of selected column at data sheet
for key in key_index:
    keyList.append(key)


# 遍历每个指标工作表
for i in name_index:
    print('\n', i)
    # 提取当前指标工作表的列名和关键列索引
    first_row_EachSheet = list(i[0].rows)[0]			# cat: first sheet 'original': first row;
    col_list_EachSheet = [i.value for i in first_row_EachSheet]      # cat: get all the columns name in sheet 'original'
    key_col_EachSheet_2 = key_col_EachSheet.copy()                   # add i_indicator
    key_col_EachSheet_2.append(i[0].title)   # 新增当前指标列名（如i_data）到关键列
    # 记录指标表中关键列的（名称, 索引）
    key_index_EachSheet = [(j, col_list_EachSheet.index(n)) for n, j in zip(key_col_EachSheet_2, key_col_EachSheet_2)]

    keyList_EachSheet = []  # find out # of selected column at indicator sheet 指标表关键列列表
    for key_EachSheet in key_index_EachSheet:
        keyList_EachSheet.append(key_EachSheet)

    # 计算指标表的有效行数（排除空行）
    for row in i[0].rows:  # write self maxrow program, to fix i[0].max_row count empty rows issue
        values = [cell.value for cell in row]
        if any(values):   # 只要行中有一个非空值，就认为是有效行
            maxrow = row[0].row

    # 遍历指标表中的每条规则
    for j in range(2, maxrow+1):
        print('j=: ', i[0].cell(row=j, column=keyList_EachSheet[0][1]+1).value)  # Subcategory 股票
        # 是否需要标记
        if i[0].cell(row=j, column=keyList_EachSheet[4][1] + 1).value == 'yes':  # i[0].title
            # print("column",keyList_EachSheet[4][1] + 1)    column 5
            start_row = 0
            end_row = 0

            # 遍历主数据表，定位符合规则的日期范围
            for k in range(2, data_list[0].max_row+1):
                # 匹配条件1：文件路径一致 + 规则起始日期与主表日期匹配
                if (i[0].cell(row=j, column=keyList_EachSheet[1][1] + 1).value == data_list[0].cell(row=k, column=keyList[1][1] + 1).value
                        and i[0].cell(row=j, column=keyList_EachSheet[2][1] + 1).value[-6:] == str(data_list[0].cell(row=k, column=keyList[2][1] + 1).value)):
                    if i[0].cell(row=j, column=keyList_EachSheet[2][1] + 1).value[:-6].strip(' ') == ">":
                        start_row = k - 1
                    else:
                        start_row = k
                    print('start_row:', start_row)

                # 匹配条件2：文件路径一致 + 规则结束日期与主表日期匹配，且日期不重复
                if (i[0].cell(row=j, column=keyList_EachSheet[1][1]+1).value == data_list[0].cell(row=k, column=keyList[1][1]+1).value
                        and i[0].cell(row=j, column=keyList_EachSheet[3][1]+1).value[-6:] == str(data_list[0].cell(row=k, column=keyList[2][1]+1).value)
                        and str(data_list[0].cell(row=k, column=keyList[2][1]+1).value) != str(data_list[0].cell(row=k-1, column=keyList[2][1]+1).value)):    # to avoid duplicated date records
                    if i[0].cell(row=j, column=keyList_EachSheet[3][1]+1).value[:-6].strip(' ') == "<":
                        end_row = k + 1
                    else:
                        end_row = k
                    print('end_row:', end_row)  # cat: for debug
            # 在主表中标记符合条件的行（指标列设为1）
            if start_row > 0 & end_row > 0:
                [data_list[0].cell(row=k, column=i[1]+1, value=1) for k in range(end_row, start_row+1)]
        else:
            print('No need to record!')

'''
below is for output
'''
Name_Idata = ('combine_i_data_manage_' + str(Update))
xlsName_Idata = (FinalDataPath + Name_Idata + '.xlsx')
I_Data.save(xlsName_Idata)
