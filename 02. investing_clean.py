import os
import re
import numpy as np
import pandas as pd
import util.aSetting as aSetting

import warnings
warnings.filterwarnings("ignore")

Update = aSetting.Update
EndDate_Str = aSetting.EndDate_Str_1
MasterSheet = aSetting.MasterSheet
MasterOrigin = aSetting.MasterOrigin
BasePath = aSetting.BasePath
RawDataPath = aSetting.RawDataPath
LogPath = aSetting.LogPath
FinalDataPath = aSetting.FinalDataPath

DataSelectColumn = ['Date', 'Price', 'Open', 'High', 'Low', 'Volume_k', 'Change', 'Change_pct']
SplitDate = u' |-,'
SplitVolume = ' |,[A-Z]'
DictRenameCol = {'Month': 'Date',
                 'Chg.': 'Change',
                 'Change %': 'Change_pct',
                 'Volume': 'Volume_k',
                 'Vol.': 'Volume_k',
                 'Last': 'Price'
                 }
DictDate = {u'\xa0': ' '}  # clean date format before split
DictDate_1 = {"/":" "}
DictMonth = {"Jan": "01",
             "Feb": "02",
             "Mar": "03",
             "Apr": "04",
             "May": "05",
             "Jun": "06",
             "Jul": "07",
             "Aug": "08",
             "Sep": "09",
             "Oct": "10",
             "Nov": "11",
             "Dec": "12"
             }  # Convert english month to key-value
DictPrice = {':': '',
             "'": '',
             '-': '.',
             's': '',
             ",": ''
             }  # Clean price

DictVolume = {'B': '&1000000',
              'M': '&1000',
              'K': '&1',
              ':': '',
              "'": '',
              '-': '',
              ",": ''
              }  # Clean volume, ':' can not separate properly

DictChgPct = {'%': '',
              ':': '',
              "'": '',
              'nan': '',
              'NAN%': '',
              ",": ''

              }  # Clean change_pct
numerics = ['int16', 'int32', 'int64', 'float16', 'float32', 'float64', 'bool']

inclean = {
    'NaN%': ':',
    '': ':'
}


def replace_all_df(dataf, dic):
    if dataf.dtypes not in numerics:
        for i, j in dic.items():
            dataf = dataf.astype(str).str.replace(i, j)

    else:
        dataf = dataf.astype(str)
    return dataf


def Negative_N_Replace_df(dic, dataf, col):

    """
    function Negative_N_Replace_df build on function replace_all_df
    deal with negative value, along the data cleaning
    """

    if dataf[col].dtypes in numerics:
        # to avoid turn numeric to scientific number after astype(str) and mishandle it, in else step
        dataf[col] = pd.to_numeric(dataf[col])
    else:
        Negative = (dataf[col].astype(str).str[:1] == '-').sum()   # str[:1]正向输出
        if Negative > 0:  # cat: judge if it is negative no
            print('negative value, it needs to treat it!')
            dataf['Sign'] = (dataf[col].astype(str).str[:1] == '-').map(lambda x: -1 if x == 1 else 1)
            dataf['_n'] = dataf[col].astype(str).str[:1].replace('-', '') + dataf[col].astype(str).str[1:]  # cat: just take the value, remove the sign
            dataf['_n'] = replace_all_df(dataf['_n'], dic)  # cat: call function replace_all_df

            dataf[col] = pd.to_numeric(dataf['_n']) * pd.to_numeric(dataf['Sign'])
            del dataf['Sign'], dataf['_n']
        else:
            print('positive value!')
            # normal replacement
            dataf[col] = replace_all_df(dataf[col], dic)
            dataf[col] = pd.to_numeric(dataf[col], errors='coerce')
    return dataf[col]


def normalize_year(value):
    """Convert year fragments such as ``20`` or ``15, 2020`` to YYYY."""
    text = str(value).strip()
    full_year = re.search(r'(?<!\d)(\d{4})(?!\d)', text)
    if full_year:
        return int(full_year.group(1))

    number = pd.to_numeric(text, errors='coerce')
    if pd.isna(number):
        return np.nan
    number = int(number)
    return 1900 + number if number > 30 else 2000 + number


def parse_raw_dates_to_month(data_series):
    """Parse raw dates from both crawler locales and return YYYYMM integers.

    Supported examples include 2020/5/12, 05/12/2020, May 12, 2020,
    and 2020年5月12日. The downstream model uses month-level YYYYMM dates.
    """
    text = data_series.astype(str).str.strip()
    text = text.str.replace(r'年|月', '/', regex=True).str.replace('日', '', regex=False)
    try:
        parsed = pd.to_datetime(text, errors='coerce', format='mixed')
    except TypeError:
        parsed = pd.to_datetime(text, errors='coerce')
    return (parsed.dt.year * 100 + parsed.dt.month).astype('Int64')


Master = pd.read_excel(MasterSheet, sheet_name='RawData', engine='openpyxl')
print(Master.head())  # cat: check if load in data

namesOrg = Master[['NoSecurity', 'Choose_to_Download', 'Choose_to_DataCombine', 'Choose_to_ModelDev', 'Update_time', 'FilePath',
                   'Subcategory', 'Futures', 'Country', 'Terms_year', 'Terms', 'Macro_Bond', 'Macro_Equity',
                   'Directory', 'Subdirectory1', 'Subdirectory2', 'Name_investing', 'Description', 'StartDate', 'EndDate',
                   ]]   # cat: choose the variables we need
namesOrg['Update_time'] = namesOrg['Update_time'].replace('、', None)

namesOrg['Update_time'] = namesOrg['Update_time'].fillna(0).astype(int)

names = namesOrg[(namesOrg['Choose_to_DataCombine'] == "yes") & (namesOrg['Update_time'].astype(int) == Update)].reset_index(drop=True)
print('MasterSheet shape: ', Master.shape)      # (1286, 31)
print('names shape: ', names.shape)             # (1049, 20)

# create data template to load in data
data_template = pd.DataFrame(data=None, index=None, columns=DataSelectColumn,dtype=object)

# create data template to combine all securities data with price and feature inf
combineCol = np.append(names.columns.values, DataSelectColumn)
combineCol = np.append(combineCol, 'NoCombine')
combined_template = pd.DataFrame(data=None, index=None, columns=combineCol,dtype=object)
combined_item = combined_template

for ind, row in names.iterrows():
    names_2 = names[ind: ind+1].copy()
    names_2.loc[:, 'NoCombine'] = ind + 1
    filePath = str(names_2['FilePath'].values[0])
    filename = RawDataPath + filePath + '.csv'
    print(filename)
    # v1.2: modify due to filename which downloaded fail
    try:
        data = pd.read_csv(filename, encoding="utf_8")
    except:
        continue


    if len(data) > 1:  # if there is data, then proceed following
        data.rename(columns=DictRenameCol, inplace=True)        # rename and standardize col name
        # Volume is optional for some instruments.  Keep a standard numeric
        # column so files with and without Vol. can be combined together.
        if 'Volume_k' not in data.columns:
            data['Volume_k'] = np.nan
        data = data[data['Date'].astype(str).str.contains('Cash') == 0]     # delete row with "cash"
        data = pd.concat([data_template, data], ignore_index=True, sort=True)  # remove sort to avoid warning

        # print(len(data['Date'][1]))

        # Normalize crawler output such as 2020/5/12 to the model's YYYYMM.
        data['Date'] = parse_raw_dates_to_month(data['Date'])

        # clean price 对price、open数据中带有的字符进行处理
        data['Price'] = replace_all_df(data['Price'], DictPrice)
        data['Open'] = replace_all_df(data['Open'], DictPrice)

        data['High'] = Negative_N_Replace_df(DictPrice, data, col='High')  # for TR_CC crb_nickel 200605, it has negative value
        data['Low'] = Negative_N_Replace_df(DictPrice, data, col='Low')    # for TR_CC crb_nickel 200605, it has negative value

        #print(filename, 'negative value, it needs to treat it!')
        data['Change'] = Negative_N_Replace_df(DictPrice, data, col='Change')
        # data['Change_pct'] = replace_all_df(data['Change_pct'], inclean)
        try:
            data['Change_pct'] = Negative_N_Replace_df(DictChgPct, data, col='Change_pct')
        except:
            # 先替换缺失值，再清洗
            data['Change_pct'] = replace_all_df(data['Change_pct'], inclean)
            # data['Change_pct'] = Negative_N_Replace_df(inclean, data, col='Change_pct')
            data['Change_pct'] = Negative_N_Replace_df(DictChgPct, data, col='Change_pct')

        # clean volume: normalize every value to thousands (Volume_k).
        # K -> x1, M -> x1,000, B -> x1,000,000. Missing volume remains NaN.
        data['Volume_k'] = replace_all_df(data['Volume_k'], inclean)
        data['Volume_k'] = replace_all_df(data['Volume_k'], DictVolume)
        # print(data['Volume_k'])
        # print(data['Volume_k'].str.contains('&').sum())
        # 拆分数值和单位，计算实际值
        a = data['Volume_k'].str.contains('&').sum()
        if a > 0:
            # 再判断元素是否为空 一直到不为空的时候
            data['Volume_k'] = replace_all_df(data['Volume_k'], DictVolume)
            # 使用 n 指定最大拆分次数，并通过 expand=True 转为DataFrame
            split_volume = data['Volume_k'].str.split('&', n=1, expand=True)
            data['Volume_n'] = split_volume[0]  # 数值部分
            data['Volume_unit'] = split_volume[1]  # 单位倍数

            try:
                # 计算最终值 数值×单位倍数
                data['Volume_k'] = pd.to_numeric(data['Volume_n']) * pd.to_numeric(data['Volume_unit'])
            except:
                # 无单位的直接转为数值
                data['Volume_k'] = pd.to_numeric(data['Volume_k'], errors='coerce')

        else:
            data['Volume_k'] = pd.to_numeric(data['Volume_k'], errors='coerce')
            # errors='coerce' 只转换能转换的，不能转换的赋值为 NaN(Not a Number) 用NaN替换所有非数字值
        # selected date period 保留日期 <= Update的数据
        # 非期货数据只保留日期≤当前更新时间（Update）的数据，期货数据不限制
        if names_2['Futures'].values[0] != 'yes':
            data = data[data['Date'] <= Update]
        else:
            data = data

        # 筛选出预设的核心字段（DataSelectColumn）
        dataSelect = data[data.columns[data.columns.isin(DataSelectColumn)]]  # take selected column
        dataSelect = dataSelect.reset_index()  # reset index before final merge, for price
        dataSelect.index = dataSelect.index + 1

        # for each security, combine feature inf
        # 仅合并有有效数据的项
        if len(dataSelect) > 0:  # if there is data after cleaning, then proceed following
            names_2 = pd.concat([names_2] * len(dataSelect), ignore_index=True).reset_index(drop=True)   # reset index before final merge, for feature
            names_2.index = names_2.index + 1
            data_w_name = pd.merge(names_2, dataSelect, how='outer', left_index=True, right_index=True)  # combine feature and price inf, with index
            data_w_name_2 = data_w_name[combined_template.columns.values.tolist()]  # standadize the fix column no for each security
            data_w_name_3 = data_w_name_2.reindex(combineCol, axis=1)  # reorder column to fit into combine_template
            data_w_name_3['NoObs'] = data_w_name_3.index
            combined_item = pd.concat([combined_item, data_w_name_3], ignore_index=True, sort=False)  # append to with all other securities

'''
below is for output
'''
# design output format and output 设计输出格式和输出
combineClean_csv = combined_item.copy()
outputCol = ['NoSecurity', 'NoCombine', 'Subcategory', 'NoObs', 'Date', 'Price', 'Open', 'High', 'Low', 'Volume_k', 'Change', 'Change_pct',
             'i_data', 'i_event', 'i_extremeValue', 'i_yield_price', 'Yield_price_reverse',
             # added empty fields at this step for next ste
             'i_missing', 'i_dup',
             # added empty fields at this step for next step
             'Futures', 'Country', 'Terms_year', 'Terms', 'Macro_Bond', 'Macro_Equity',
             'Directory', 'Subdirectory1', 'Subdirectory2', 'Name_investing', 'Description', 'FilePath', 'StartDate', 'EndDate',
             'Choose_to_Download', 'Choose_to_DataCombine', 'Choose_to_ModelDev']
combineClean_csv = combineClean_csv.reindex(outputCol, axis=1)
Name_clean = ('combine_clean_' + str(Update))
# csvName_clean = FinalDataPath + Name_clean + '.csv'
xlsName_clean = FinalDataPath + Name_clean + '.xlsx'

if not os.path.exists(FinalDataPath):
    os.mkdir(FinalDataPath)

# combineClean_csv.to_csv(csvName_clean, index=False)  # 1 min: for quick check and debug
combineClean_csv.to_excel(xlsName_clean, index=False)  # 10 mins: for formal use, since smaller data size and no dif datatype issues
