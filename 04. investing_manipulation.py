import os
import numpy as np
import pandas as pd
import util.aSetting as aSetting

Update = aSetting.Update
UpdateYear = aSetting.UpdateYear
UpdateMonth = aSetting.UpdateMonth
EndDate_Str = aSetting.EndDate_Str_1
MasterSheet = aSetting.MasterSheet
MasterOrigin = aSetting.MasterOrigin
BasePath = aSetting.BasePath
RawDataPath = aSetting.RawDataPath
LogPath = aSetting.LogPath
FinalDataPath = aSetting.FinalDataPath

import warnings
warnings.filterwarnings('ignore')

# read excel data
Name_Idata = ('combine_i_data_manage_' + str(Update))
xlsName_Idata = (FinalDataPath + Name_Idata + '.xlsx')
combined_item_fix = pd.read_excel(xlsName_Idata, engine='openpyxl')
resultCombine = pd.DataFrame()

# Volume is optional on Investing.com (some instruments/periods return it as
# an empty column).  Keep the column in the pipeline so those instruments do
# not make the manipulation step fail with KeyError or string/numeric errors.
if 'Volume_k' not in combined_item_fix.columns:
    combined_item_fix['Volume_k'] = np.nan
else:
    combined_item_fix['Volume_k'] = pd.to_numeric(
        combined_item_fix['Volume_k'], errors='coerce'
    )

# len(combined_item_fix.columns)
'''
Consolidate all the final variables after adding some in this step
'''
# below are new added variables at this step 当前步骤新增的变量列
AddColThisStep = ['StartDate_mani', 'EndDate_mani',
                  'Yield_price', 'Yield', 'Derived_Price', 'Adj_Price', 'Volume_b',
                  'Yield_L12M', 'Adj_Price_L12M', 'Vol_k_L12M', 'Vol_b_L12M',
                  'Yield_Chg', 'Adj_Price_Chg', 'Vol_k_Chg', 'Vol_b_Chg', 'Yield_L12M_Chg', 'Adj_Price_L12M_Chg', 'Vol_L12M_k_Chg', 'Vol_L12M_b_Chg',
                  'Yield_Chg_pct', 'Adj_Price_Chg_pct', 'Vol_k_Chg_pct', 'Vol_b_Chg_pct',
                  'Yield_L12M_Chg_pct', 'Adj_Price_L12M_Chg_pct', 'Vol_L12M_k_Chg_pct', 'Vol_L12M_b_Chg_pct']
# 定义输出列结构
outputCol = ['NoSecurity', 'NoCombine', 'Subcategory', 'NoObs', 'Date', 'Price', 'Open', 'High', 'Low', 'Volume_k', 'Change', 'Change_pct',
             'i_data', 'i_event', 'i_extremeValue', 'i_yield_price', 'Yield_price_reverse',          # added empty fields at 02 step
             'i_missing', 'i_dup',                                                                   # added empty fields at 02 step
             'Futures', 'Country', 'Terms_year', 'Terms',  'Macro_Bond', 'Macro_Equity',
             'Directory', 'Subdirectory1', 'Subdirectory2', 'Name_investing', 'Description', 'FilePath', 'StartDate', 'EndDate',
             'Choose_to_Download', 'Choose_to_DataCombine', 'Choose_to_ModelDev']                    # new variable added here
len(outputCol)  # 36
# outputCol.index('Choose_to_ModelDev')
# 在指定位置插入新增列
outputCol[outputCol.index('Choose_to_ModelDev')+1: outputCol.index('Choose_to_ModelDev')+1] = AddColThisStep
print('Lenths of Added Columns: %', len(AddColThisStep))
print('Lenths of Final Output Columns: %', len(outputCol))

'''
Remove i_data
'''
combined_item_fix_goodData = combined_item_fix[combined_item_fix['i_data'] != 1].copy()
print("combined_item_fix rows: %i", combined_item_fix.shape[0])
print("after drop i_data: %i", combined_item_fix_goodData.shape[0])

'''
FX:  convert FFF_USD to USD_FFF   # 外汇转换
KYD_USD
GAU_USD
XAG_USD
XAU_USD
'''
# for selected FX part 筛选需要倒置的外汇数据
FXList = ['KYD_USD', 'GAU_USD', 'XAG_USD', 'XAU_USD']
# 从有效数据中筛选出这些外汇数据，创建副本
combined_item_fix_FX = combined_item_fix_goodData[combined_item_fix_goodData['Subcategory'].isin(FXList)].copy()
# 再创建一个副本作为“镜像数据”（用于后续反向计算）
combined_item_fix_FX_2 = combined_item_fix_FX.copy()  # take the mirror value

# clean name relevant field 转换外汇货币对的名称信息
combined_item_fix_FX.loc[:, 'Subcategory'] = combined_item_fix_FX_2['Subcategory'].str[-3:] + '_' + combined_item_fix_FX_2['Subcategory'].str[:3]
combined_item_fix_FX.loc[:, 'Name_investing'] = combined_item_fix_FX['Subcategory']
combined_item_fix_FX.loc[:, 'Description'] = combined_item_fix_FX['Subcategory']

# write in data
print(xlsName_Idata)
print(combined_item_fix)
# 计算反向汇率的价格数据 反向汇率 = 1 / 原始汇率
combined_item_fix_FX.loc[:, 'Price'] = 1/combined_item_fix_FX_2['Price'].astype(float)
combined_item_fix_FX.loc[:, 'Open'] = 1/combined_item_fix_FX_2['Open'].astype(float)
combined_item_fix_FX.loc[:, 'High'] = 1/combined_item_fix_FX_2['High'].astype(float)
combined_item_fix_FX.loc[:, 'Low'] = 1/combined_item_fix_FX_2['Low'].astype(float)

# for non-selected FX part: consistent the data format
# 处理非外汇数据（格式统一） 转换为浮点型
combined_item_fix_nonFX = combined_item_fix_goodData[combined_item_fix_goodData['Subcategory'].isin(FXList) == 0].copy()
convertFloat_List = ['Price', 'Open', 'High', 'Low']
# Excel 中个别记录可能包含类似 '.61.00' 的非法数字字符串。
# 转换失败的值按缺失值处理，避免整批数据因单个脏值中断。
for col in convertFloat_List:
    invalid = pd.to_numeric(combined_item_fix_nonFX[col], errors='coerce').isna() & combined_item_fix_nonFX[col].notna()
    if invalid.any():
        print(f'Invalid numeric values in {col}: {invalid.sum()} rows')
        print(combined_item_fix_nonFX.loc[invalid, ['Subcategory', 'Date', col]].head(10))
    combined_item_fix_nonFX.loc[:, col] = pd.to_numeric(
        combined_item_fix_nonFX[col], errors='coerce'
    )

# append FX fix and non-selected FX part back
# 合并转换后的外汇数据和非外汇数据，重置索引
combined_item_fix_FX = pd.concat([combined_item_fix_FX, combined_item_fix_nonFX], ignore_index=True, sort=False)
combined_item_fix_FX.shape

'''
delete missing / 0 value
'''
# 价格既不为 0 且非缺失
nonZero = (combined_item_fix_FX['Price'] != 0)
nonMissing = (combined_item_fix_FX['Price'].isnull() == 0)
combined_item_validValue = combined_item_fix_FX[nonZero & nonMissing]
combined_item_validValue.shape

'''
dedup: double date
'''
# 处理数据中的重复记录
combined_item_dup = combined_item_validValue.sort_values(['NoSecurity', 'Date', 'Price'], ascending=[True, False, False])  # sort by NoSecurity, descending Date, Price; but indicator build on 2nd case
combined_item_dup.loc[combined_item_dup.duplicated(['NoSecurity', 'Date']) == 1, 'i_dup'] = 1  # create i_dup by taking out the duplicated record
combined_item_dedup = combined_item_dup.drop_duplicates(['NoSecurity', 'Date'], keep='last')
combined_item_dedup.shape

# create yield_price / yield / Derived_Price / Adj_Price
'''
bond:  fix yield and price
'''
combined_item_fixBD = combined_item_dedup.copy()

# create yield_price, just for 'Bonds'
combined_item_fixBD['Yield_price'] = 'price'
combined_item_fixBD.loc[(combined_item_fixBD['Price'].astype(float) < 50) & (combined_item_fixBD['Directory'] == 'Bonds') & (combined_item_fixBD['Subdirectory1'] == 'World govern bonds'), 'Yield_price'] = 'yield'

# yield_price special treatment - for those don't apply rule "yield is with price < 50"
Yield_price_reverse = pd.read_excel(MasterSheet, sheet_name='Yield_price_reverse', engine='openpyxl')
Treatments = Yield_price_reverse[Yield_price_reverse['Yield_price_reverse'] != None][['FilePath', 'Yield_price_reverse']]
for i in range(len(Treatments)):
    combined_item_fixBD.loc[(combined_item_fixBD['FilePath'] == Treatments.loc[i][0]), 'Yield_price'] = Treatments.loc[i][1]

# take world_govern_bond == 'yield' and non-world_govern_bond
Yield = ((combined_item_fixBD['Subdirectory1'] == 'World govern bonds') & (combined_item_fixBD['Yield_price'] == 'yield'))
NonYield = (combined_item_fixBD['Subdirectory1'] != 'World govern bonds')
combined_item_fixBD = combined_item_fixBD[Yield | NonYield]
combined_item_fixBD.shape

# create yield   计算债券收益率
combined_item_fixBD['Yield'] = np.nan
combined_item_fixBD.loc[combined_item_fixBD['Yield_price'] == 'yield', 'Yield'] = combined_item_fixBD['Price'].astype('float64')

# create Derived_Price 计算债券派生价格
combined_item_fixBD['Derived_Price'] = np.nan
combined_item_fixBD.loc[(combined_item_fixBD['Yield_price'] == 'yield') & (combined_item_fixBD['Terms_year'].isnull() == 0), 'Derived_Price'] \
    = 100 / ((1 + combined_item_fixBD['Price'].astype(float)/100)**combined_item_fixBD['Terms_year'].astype(float))
combined_item_fixBD.shape

'''
interpolate missing value / yearly freq <= 12:
1. take data <= 2017.12, not consider future contract since they are not for modelling purpose
2. just take World govern bonds & yield part
3. combine back to non-bond set
4. fill out the time gap with empty value yield / derive price
5. extrapolate the price for empty value of derive price
6. i_missing
7. combine back with data > 2017.12
'''
# 处理 “时间序列缺失值” 并计算关键衍生指标（如 12 个月均值、变化量、变化率）
# calculate min and max date 计算每个数据项的时间范围（起止日期）
# 按文件路径（FilePath）分组，统计每个数据项的最小/最大日期
grouped = combined_item_fixBD.groupby('FilePath')
DateRange = grouped['Date'].agg([np.min, np.max])
DateRange['FilePath'] = DateRange.index
DateRange.reset_index(drop=True, inplace=True)
# 步骤 1: 清理所有可能冲突的旧列 (min/max，带后缀的，或未成功重命名的)
cols_to_remove = [col for col in combined_item_fixBD.columns
                  if col in ['min', 'max']
                  or col.endswith('_x')
                  or col.endswith('_y')
                  or col.endswith('_new_range')
                  or col in ['StartDate_mani', 'EndDate_mani']
                 ]

if cols_to_remove:
    combined_item_fixBD.drop(columns=cols_to_remove, inplace=True)
# 步骤 2: 计算每个 FilePath 的最小和最大日期
grouped = combined_item_fixBD.groupby('FilePath')
# 确认聚合结果是 'min' 和 'max'
DateRange = grouped['Date'].agg([np.min, np.max])
DateRange['FilePath'] = DateRange.index
DateRange.reset_index(drop=True, inplace=True)

# 步骤 3: 合并并正确重命名新列
# 使用 suffixes=('', '_new_range') 明确指定右表的新列后缀，以避免 MergeError
combined_item_fixBD = pd.merge(
    combined_item_fixBD,
    DateRange,
    how='left',
    on='FilePath',
    suffixes=('', '_new_range') # 右表新列名为 min_new_range, max_new_range
)

# 步骤 4: 正确重命名新列
combined_item_fixBD.rename(
    columns={
        'min': 'StartDate_mani', # 修正：目标列名是 'min'
        'max': 'EndDate_mani'    # 修正：目标列名是 'max'
    },
    inplace=True
)

# 分离现货数据与期货数据
# separate spot data and <= update / future data
Combine_Spot = combined_item_fixBD[(combined_item_fixBD['Futures'] != 'yes') & (combined_item_fixBD['Date'] <= Update)].copy()
Combine_Future = combined_item_fixBD[combined_item_fixBD['Futures'] == 'yes'].copy()

# create time template 创建 “时间模板”
# TimeTemplate = pd.DataFrame(combined_item_fixBD[combined_item_fixBD['Subcategory'] == 'S&p_500']['Date'])
# if TimeTemplate.count().values[0] == ((UpdateYear - 1980) * 12 + UpdateMonth):
#     TimeTemplate = TimeTemplate
# else:
#     TimeTemplate = pd.DataFrame(data=None, index=None, columns=None)
#     print("ERROR!  TimeTemplate can't be created!")

# create time template
sp500_data = combined_item_fixBD[
    combined_item_fixBD['Subcategory'] == 'S&p_500'
].copy()

print("\n========== TimeTemplate DEBUG ==========")
print("S&p_500 rows:", len(sp500_data))
print("S&p_500 non-null Date:", sp500_data['Date'].notna().sum())
print("S&p_500 unique Date:", sp500_data['Date'].nunique())
print("S&p_500 min Date:", sp500_data['Date'].min())
print("S&p_500 max Date:", sp500_data['Date'].max())

expected_count = (UpdateYear - 1980) * 12 + UpdateMonth

print("Expected Date count:", expected_count)
print("Actual Date count:", sp500_data['Date'].notna().sum())

print("\nDuplicate dates:")
print(
    sp500_data[
        sp500_data['Date'].duplicated(keep=False)
    ][['Date', 'Price', 'FilePath']].sort_values('Date')
)

print("========================================\n")

TimeTemplate = sp500_data[['Date']].dropna().drop_duplicates().copy()

if len(TimeTemplate) != expected_count:
    print(
        f"WARNING! TimeTemplate count mismatch: "
        f"expected={expected_count}, actual={len(TimeTemplate)}"
    )

# 遍历每个现货数据项，补全时间与处理缺失
# create template for each item result
resultTemplate = pd.DataFrame(data=None, index=None, columns=outputCol)
Combine_Spot_MissingFix = resultTemplate.copy()
# 定义“无需插值、仅向前填充”的列
FillMissingCol = ['NoSecurity', 'NoCombine', 'Subcategory',
                  'i_yield_price',  # some i_indicator will need to fill in the blank as well
                  'Yield_price',
                  'Futures', 'Country', 'Terms_year', 'Terms', 'Macro_Bond', 'Macro_Equity',
                  'Directory', 'Subdirectory1', 'Subdirectory2', 'Name_investing', 'Description', 'FilePath', 'StartDate', 'EndDate',
                  'Choose_to_Download', 'Choose_to_DataCombine', 'Choose_to_ModelDev',
                  'StartDate_mani', 'EndDate_mani'
                  ]

# 按证券唯一编号（NoSecurity）遍历每个现货数据项
UniqueSubCat = Combine_Spot.groupby('NoSecurity')['Subcategory'].unique()
for NoSec in UniqueSubCat.index.values.tolist():
    print(UniqueSubCat[UniqueSubCat.index == NoSec])
    # 缺失值和时间间隔处理
    # 提取当前证券的原始数据
    left = Combine_Spot[Combine_Spot['NoSecurity'] == NoSec].copy() # 确保 left 是一个独立副本
    # DEBUG: 检查 Macro Equity 的成交量原始数据
    if left['Macro_Equity'].notna().any():
        print(
            f"[Volume DEBUG] "
            f"NoSecurity={NoSec}, "
            f"Subcategory={left['Subcategory'].iloc[0]}, "
            f"rows={len(left)}, "
            f"Volume_k non-null={left['Volume_k'].notna().sum()}"
        )

    # 与时间模板合并，补全缺失日期
    result = pd.merge(left, TimeTemplate, how='outer', on='Date', indicator=True)  # 填充时间间隔
    result = result.sort_values(['Date'], ascending=[False]) # 按日期降序排列（最新在前）
    # 标记缺失值（i_missing = 1）
    result.loc[result['_merge'] == 'right_only', 'i_missing'] = 1  # 创建 i_missing

    # 检查 result 中有哪些 FillMissingCol 里的列
    FillMissingCol_actual = [col for col in FillMissingCol if col in result.columns]

    # 使用实际存在的列进行 fillna 操作
    # 使用 .loc 确保赋值正确
    result.loc[:, FillMissingCol_actual] = result.loc[:, FillMissingCol_actual].fillna(method='ffill')

    # 筛选“在数据实际时间范围内”的行
    result = result[(result['Date'] >= result['StartDate_mani']) & (result['Date'] <= result['EndDate_mani'])]

    # reset_index
    # 'NoObs'
    result = result.reset_index()  # reset index to do following lead and lag
    resultIndex = result['Yield'].index.values.tolist()
    result['NoObs'] = result.index + 1

    # 标记区间型指标（i_event / i_extremeValue）
    # intrapolate 'i_event', 'i_extremeValue', since these variables are rnage and it can't simply fillna
    if result['i_event'].count() > 0:
        result.iloc[result[result['i_event'] == 1].index.min():result[result['i_event'] == 1].index.max(), result.columns.get_loc("i_event")] = 1
    if result['i_extremeValue'].count() > 0:
        result.iloc[result[result['i_extremeValue'] == 1].index.min():result[result['i_extremeValue'] == 1].index.max(), result.columns.get_loc("i_extremeValue")] = 1

    # 线性插值填充数值型缺失（Derived_Price / Yield / Volume_k）
    # intrapolate 'Derived_Price', 'Yield', 'Volume_k', 'Volume_b'
    result['Derived_Price'] = result['Derived_Price'].interpolate(method='linear')  # fill in missing value for price, yield, volume variables for time gap
    result['Yield'] = result['Yield'].interpolate(method='linear')
    result['Volume_k'] = result['Volume_k'].interpolate(method='linear')
    result['Volume_b'] = result['Volume_k'] / 1000000      # convert unit thousand to billion

    # recalculate
    # create 'Adj_Price' 调整后价格（Adj_Price）
    # interpolate 'Yield_price' for 'Yield_price' == ['Price']
    result['Adj_Price'] = np.nan
    result.loc[result['Yield_price'] == 'yield', 'Adj_Price'] = result['Derived_Price']
    result.loc[result['Yield_price'] != 'yield', 'Adj_Price'] = result['Price']
    result['Adj_Price'] = result['Adj_Price'].interpolate(method='linear')

    # calculation: 12个月均值（L12M）
    # 'Yield_L12M', 'Adj_Price_L12M', 'Vol_k_L12M', 'Vol_b_L12M',
    result['Yield_L12M'] = np.nan
    result['Adj_Price_L12M'] = np.nan
    result['Vol_k_L12M'] = np.nan
    result['Vol_b_L12M'] = np.nan
    for i in resultIndex:
        if i <= resultIndex[-1] + 1 - 12:  # to create missing value if not have enough 12m history
            result.loc[i, 'Yield_L12M'] = result['Yield'][i: i+12].mean()
            result.loc[i, 'Adj_Price_L12M'] = result['Adj_Price'][i: i + 12].mean()
            result.loc[i, 'Vol_k_L12M'] = result['Volume_k'][i: i + 12].mean()
            result.loc[i, 'Vol_b_L12M'] = result['Volume_b'][i: i + 12].mean()

    # calculation: 当期与上期绝对变化量（Chg）
    # 'Yield_Chg', 'Adj_Price_Chg', 'Vol_k_Chg', 'Vol_b_Chg', 'Yield_L12M_Chg', 'Adj_Price_L12M_Chg', 'Vol_L12M_k_Chg', 'Vol_L12M_b_Chg'
    result['Yield_Chg'] = np.nan
    result['Adj_Price_Chg'] = np.nan
    result['Vol_k_Chg'] = np.nan
    result['Vol_b_Chg'] = np.nan
    for i in resultIndex:
        if i <= resultIndex[-1] - 1:
            result.loc[i, 'Yield_Chg'] = (result['Yield'][i] - result['Yield'][i+1])
            result.loc[i, 'Adj_Price_Chg'] = (result['Adj_Price'][i] - result['Adj_Price'][i + 1])
            result.loc[i, 'Vol_k_Chg'] = (result['Volume_k'][i] - result['Volume_k'][i + 1])
            result.loc[i, 'Vol_b_Chg'] = (result['Volume_b'][i] - result['Volume_b'][i + 1])

    result['Yield_L12M_Chg'] = np.nan
    result['Adj_Price_L12M_Chg'] = np.nan
    result['Vol_L12M_k_Chg'] = np.nan
    result['Vol_L12M_b_Chg'] = np.nan
    for i in resultIndex:
        if i <= resultIndex[-1]+1 - 12 - 1:
            result.loc[i, 'Yield_L12M_Chg'] = (result['Yield_L12M'][i] - result['Yield_L12M'][i + 1])
            result.loc[i, 'Adj_Price_L12M_Chg'] = (result['Adj_Price_L12M'][i] - result['Adj_Price_L12M'][i + 1])
            result.loc[i, 'Vol_L12M_k_Chg'] = (result['Vol_k_L12M'][i] - result['Vol_k_L12M'][i + 1])
            result.loc[i, 'Vol_L12M_b_Chg'] = (result['Vol_b_L12M'][i] - result['Vol_b_L12M'][i + 1])

    # calculation: 12个月均值的绝对变化量（L12M_Chg）
    # 'Yield_Chg_pct', 'Adj_Price_Chg_pct', 'Vol_k_Chg_pct', 'Vol_b_Chg_pct','Yield_L12M_Chg_pct', 'Adj_Price_L12M_Chg_pct', 'Vol_L12M_k_Chg_pct', 'Vol_L12M_b_Chg_pct'
    result['Yield_Chg_pct'] = np.nan
    result['Adj_Price_Chg_pct'] = np.nan
    result['Vol_k_Chg_pct'] = np.nan
    result['Vol_b_Chg_pct'] = np.nan
    for i in resultIndex:
        if i <= resultIndex[-1] - 1:
            if result['Yield'][i + 1] != 0:
                result.loc[i, 'Yield_Chg_pct'] = (result['Yield_Chg'][i] / result['Yield'][i + 1]) * 100
            if result['Adj_Price'][i + 1] != 0:
                result.loc[i, 'Adj_Price_Chg_pct'] = (result['Adj_Price_Chg'][i] / result['Adj_Price'][i + 1]) * 100
            if result['Volume_k'][i + 1] != 0:
                result.loc[i, 'Vol_k_Chg_pct'] = (result['Vol_k_Chg'][i] / result['Volume_k'][i + 1]) * 100
            if result['Volume_b'][i + 1] != 0:
                result.loc[i, 'Vol_b_Chg_pct'] = (result['Vol_b_Chg'][i] / result['Volume_b'][i + 1]) * 100

    # 变化率（Chg_pct）
    result['Yield_L12M_Chg_pct'] = np.nan
    result['Adj_Price_L12M_Chg_pct'] = np.nan
    result['Vol_L12M_k_Chg_pct'] = np.nan
    result['Vol_L12M_b_Chg_pct'] = np.nan
    for i in resultIndex:
        if i <= resultIndex[-1]+1 - 12 - 1:
            if result['Yield_L12M'][i + 1] != 0:
                result.loc[i, 'Yield_L12M_Chg_pct'] = (result['Yield_L12M_Chg'][i] / result['Yield_L12M'][i + 1]) * 100
            if result['Adj_Price_L12M'][i + 1] != 0:
                result.loc[i, 'Adj_Price_L12M_Chg_pct'] = (result['Adj_Price_L12M_Chg'][i] / result['Adj_Price_L12M'][i + 1]) * 100
            if result['Vol_k_L12M'][i + 1] != 0:
                result.loc[i, 'Vol_L12M_k_Chg_pct'] = (result['Vol_L12M_k_Chg'][i] / result['Vol_k_L12M'][i + 1]) * 100
            if result['Vol_b_L12M'][i + 1] != 0:
                result.loc[i, 'Vol_L12M_b_Chg_pct'] = (result['Vol_L12M_b_Chg'][i] / result['Vol_b_L12M'][i + 1]) * 100

    # 整合所有现货数据
    # combine each spot item
    if NoSec == 1:
        resultCombine = result.copy()
    else:
        # 修正后的代码
        resultCombine = pd.concat([resultCombine, result], ignore_index=True)

# standardize final spot output
Combine_Spot_MissingFix = resultCombine[outputCol].copy()
Combine_Spot_MissingFix = Combine_Spot_MissingFix.reindex(outputCol, axis=1)  # reorder column to fit into combine_template

# merge back spot_MissingFix with Combine_Future
Combine_Spot_Future = pd.concat([Combine_Spot_MissingFix, Combine_Future], ignore_index=True, sort=False)
Combine_Spot_Future = Combine_Spot_Future.reindex(outputCol, axis=1)  # reorder column to fit into combine_template
print('obs of Combine_Spot_MissingFix is: %i, vs  obs of Combine_Spot_Future is: %i' % (len(Combine_Spot_MissingFix), len(Combine_Spot_Future)))

'''
Final output
'''

Name_Mani = ('combine_manipulation_' + str(Update))
# csvName_Mani = (FinalDataPath + Name_Mani + '.csv')
xlsName_Mani = (FinalDataPath + Name_Mani + '.xlsx')

# Combine_Spot_Future.to_csv(csvName_Mani, index=False)  # 1 min: for quick check and debug
Combine_Spot_Future.to_excel(xlsName_Mani, index=False)  # 2 mins: for formal use, since smaller data size and no dif datatype issues

'''
QC: dataflow output
'''
# 1. 生成QC数据框
QCdata = {'DataFlow': ['combine_i_data_manage', 'after drop i_data', 'delete missing',
                       'dedup double date', 'drop price for bond part', 'fill time gap for spot item'],
          'Obs': [combined_item_fix.shape[0], combined_item_fix_goodData.shape[0], combined_item_validValue.shape[0],
                  combined_item_dedup.shape[0], combined_item_fixBD.shape[0], Combine_Spot_Future.shape[0]],
          'Column': [combined_item_fix.shape[1], combined_item_fix_goodData.shape[1], combined_item_validValue.shape[1],
                     combined_item_dedup.shape[1], combined_item_fixBD.shape[1], Combine_Spot_Future.shape[1]]}
QCframe = pd.DataFrame(QCdata)
QCframe['RemainRate_pct'] = (QCframe['Obs'] / QCframe['Obs'][0] * 100).round(2)
QCframe['DroppingRate_pct'] = 100 - QCframe['RemainRate_pct']
QCCol = ['DataFlow', 'Obs', 'Column', 'RemainRate_pct', 'DroppingRate_pct']
QCframe = QCframe.reindex(QCCol, axis=1)

# 2. 处理QC文件
xlsName_QC = LogPath + '\\QC_' + str(Update) + '.xlsx'
file_exists = os.path.exists(xlsName_QC)

# 空表列名
empty_qc_columns = [
    'NoSecurity', 'FilePath', 'Subcategory',
    'Choose_to_Download', 'QC_download', 'Download_Status'
]

# 读取已有QcDownload表（若文件存在且有效）
if file_exists:
    try:
        # 用openpyxl读取现有文件（避免引擎冲突）
        with pd.ExcelFile(xlsName_QC, engine='openpyxl') as xls:
            QCDownloadSheet = pd.read_excel(xls, sheet_name='QcDownload')
    except:
        # 若文件损坏，强制创建新表
        print(f"文件 {xlsName_QC} 损坏，创建新表...")
        QCDownloadSheet = pd.DataFrame(columns=empty_qc_columns)
else:
    # 新文件：创建空表
    QCDownloadSheet = pd.DataFrame(columns=empty_qc_columns)

# 3. 写入文件（使用openpyxl引擎，避免格式冲突）
with pd.ExcelWriter(xlsName_QC, engine='openpyxl', mode='w') as QCWriter:
    # 写入QcDownload表
    QCDownloadSheet.to_excel(QCWriter, sheet_name='QcDownload', index=False)
    # 写入Dataflow表
    QCframe.to_excel(QCWriter, sheet_name='Dataflow', index=False)

print(f"QC文件已保存至：{xlsName_QC}")

'''
notes:
    it is reasonable for blank 'Adj_Price', if date > Update which is future product and not do data treatment
'''
#
# # create directory for SAS pricing job later on
# pathSASOut1 = 'Data\\SASData\\' + str(Update)
# pathSASOut2 = pathSASOut1 + '\\Investing'    # for sas data later on
# if not os.path.exists(pathSASOut1):
#     os.makedirs(pathSASOut1)
# if not os.path.exists(pathSASOut2):
#     os.makedirs(pathSASOut2)
