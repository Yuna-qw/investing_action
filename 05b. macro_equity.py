import numpy as np
import pandas as pd
import util.aSetting as aSetting

Update = aSetting.Update
FinalDataPath = aSetting.FinalDataPath

import warnings
warnings.filterwarnings('ignore')

# alternative get source data
Name_Mani = ('combine_manipulation_' + str(Update))
xlsName_Mani = (FinalDataPath + Name_Mani + '.xlsx')
Combine_Spot_Future_2 = pd.read_excel(xlsName_Mani, engine='openpyxl')


EquityCol = ['NoSecurity', 'Subcategory', 'NoObs', 'Date',
             'i_data', 'i_event', 'i_extremeValue', 'i_yield_price', 'Yield_price_reverse',
             'i_missing', 'i_dup',
             'Country', 'Macro_Equity',
             'Yield_price', 'Adj_Price', 'Adj_Price_L12M', 'Vol_b_L12M',
             'Adj_Price_Chg', 'Adj_Price_L12M_Chg', 'Vol_L12M_b_Chg',
             'Adj_Price_Chg_pct', 'Adj_Price_L12M_Chg_pct', 'Vol_L12M_b_Chg_pct']

# Volume is not available for every instrument.  Preserve the expected
# schema when 04 produces an empty/missing volume series; equity price
# indicators can still be calculated independently.
for col in EquityCol:
    if col not in Combine_Spot_Future_2.columns:
        Combine_Spot_Future_2[col] = np.nan
for col in ['Vol_b_L12M', 'Vol_L12M_b_Chg', 'Vol_L12M_b_Chg_pct']:
    Combine_Spot_Future_2[col] = pd.to_numeric(
        Combine_Spot_Future_2[col], errors='coerce'
    ).fillna(0)
Combine_Spot_Future_3 = Combine_Spot_Future_2[Combine_Spot_Future_2['Macro_Equity'].isnull() == 0][EquityCol].copy()
print('Macro Equity Data Observation:  %i' % Combine_Spot_Future_3.shape[0])   # 11297


# get yield data by country by date, with select ShortTerm, MidTerm and LongTerm yield
Category = ['Country', 'Date']
ValueEquityColumn = ['Adj_Price', 'Adj_Price_L12M', 'Vol_b_L12M', 'Adj_Price_Chg', 'Adj_Price_L12M_Chg', 'Vol_L12M_b_Chg',
                     'Adj_Price_Chg_pct', 'Adj_Price_L12M_Chg_pct', 'Vol_L12M_b_Chg_pct']
Equity_Transpose = pd.pivot_table(Combine_Spot_Future_3, values=ValueEquityColumn, index=Category, aggfunc=np.mean)

Equity_Transpose_2 = Equity_Transpose.reset_index()
Equity_Transpose_2 = Equity_Transpose_2.sort_values(['Country', 'Date'], ascending=[True, False])
len(Equity_Transpose_2)

UniqueCountryCat = Equity_Transpose_2['Country'].unique()
for EquityObs in range(0, len(UniqueCountryCat)):
    Country = UniqueCountryCat[EquityObs]
    print(Country)

    # missing and time gap
    left = Equity_Transpose_2[Equity_Transpose_2['Country'] == Country].copy()
    result = left.sort_values(['Date'], ascending=[False])

    # reset_index
    # 'NoObs',
    result = result.reset_index()  # reset index to do following lead and lag
    resultIndex = result['Date'].index.values.tolist()
    result['NoObs'] = result.index + 1

    # calculation:不同周期的股价均值
    # L3M / L36M
    result['Adj_Price_L3M'] = np.nan
    result['Adj_Price_L36M'] = np.nan
    for i in resultIndex:
        if i <= resultIndex[-1] + 1 - 3:
            result.loc[i, 'Adj_Price_L3M'] = result['Adj_Price'][i: i + 3].mean()
        if i <= resultIndex[-1] + 1 - 36:
            result.loc[i, 'Adj_Price_L36M'] = result['Adj_Price'][i: i + 36].mean()

    # calculation:不同周期均值的价差
    # Spread / ABS(spread)
    result['Dif_L3M_L12M'] = np.nan
    result['Dif_L12M_L36M'] = np.nan
    result['Dif_L3M_L36M'] = np.nan
    result['Dif_L3M_L12M'] = result['Adj_Price_L3M'] - result['Adj_Price_L12M']
    result['Dif_L12M_L36M'] = result['Adj_Price_L12M'] - result['Adj_Price_L36M']
    result['Dif_L3M_L36M'] = result['Adj_Price_L3M'] - result['Adj_Price_L36M']

    result['Dif_L3M_L12M_abs'] = np.nan
    result['Dif_L12M_L36M_abs'] = np.nan
    result['Dif_L3M_L36M_abs'] = np.nan
    result['Dif_L3M_L12M_abs'] = (result['Adj_Price_L3M'] - result['Adj_Price_L12M']).abs()
    result['Dif_L12M_L36M_abs'] = (result['Adj_Price_L12M'] - result['Adj_Price_L36M']).abs()
    result['Dif_L3M_L36M_abs'] = (result['Adj_Price_L3M'] - result['Adj_Price_L36M']).abs()

    # calculation:股价周期均值的趋势方向  周期价差的趋势
    # Spread level / Trend  # not same as SAS code, it is vs year/month ago while sas is L12M vs month ago
    # count of Spread level / Trend
    result['i_Trend_L3M'] = ''
    result['i_Trend_L12M'] = ''
    result['i_Trend_L36M'] = ''
    result['i_spread_L3M_L12M'] = ''
    result['i_spread_L3M_L36M'] = ''
    result['i_spread_L12M_L36M'] = ''
    for j in resultIndex:
        if j <= resultIndex[-1] - 12:
            if (result['Adj_Price_L3M'][j] - result['Adj_Price_L3M'][j+1]) < 0 \
                    and (result['Adj_Price_L3M'][j+1] != None) \
                    and (result['Adj_Price_L3M'][j] - result['Adj_Price_L3M'][j + 12]) < 0 \
                    and (result['Adj_Price_L3M'][j + 12] != None):
                result.loc[j, 'i_Trend_L3M'] = 'down'
            elif (result['Adj_Price_L3M'][j] - result['Adj_Price_L3M'][j+1]) > 0 \
                    and (result['Adj_Price_L3M'][j+1] != None) \
                    and (result['Adj_Price_L3M'][j] - result['Adj_Price_L3M'][j+12]) > 0 \
                    and (result['Adj_Price_L3M'][j+12] != None):
                result.loc[j, 'i_Trend_L3M'] = 'up'

            if (result['Adj_Price_L12M'][j] - result['Adj_Price_L12M'][j+1]) < 0 \
                    and (result['Adj_Price_L12M'][j+1] != None) \
                    and (result['Adj_Price_L12M'][j] - result['Adj_Price_L12M'][j + 12]) < 0 \
                    and (result['Adj_Price_L12M'][j + 12] != None):
                result.loc[j, 'i_Trend_L12M'] = 'down'
            elif (result['Adj_Price_L12M'][j] - result['Adj_Price_L12M'][j+1]) > 0 and (result['Adj_Price_L12M'][j+1] != None) \
                    and (result['Adj_Price_L12M'][j] - result['Adj_Price_L12M'][j+12]) > 0 and (result['Adj_Price_L12M'][j+12] != None):
                result.loc[j, 'i_Trend_L12M'] = 'up'

            if (result['Adj_Price_L36M'][j] - result['Adj_Price_L36M'][j+1]) < 0 and (result['Adj_Price_L36M'][j+1] != None) \
                    and (result['Adj_Price_L36M'][j] - result['Adj_Price_L36M'][j + 12]) < 0 and (result['Adj_Price_L36M'][j + 12] != None):
                result.loc[j, 'i_Trend_L36M'] = 'down'
            elif (result['Adj_Price_L36M'][j] - result['Adj_Price_L36M'][j+1]) > 0 and (result['Adj_Price_L36M'][j+1] != None) \
                    and (result['Adj_Price_L36M'][j] - result['Adj_Price_L36M'][j+12]) > 0 and (result['Adj_Price_L36M'][j+12] != None):
                result.loc[j, 'i_Trend_L36M'] = 'up'

        if j <= resultIndex[-1] - 12:
            if (result['Dif_L3M_L12M'][j] < 0) and (result['Dif_L3M_L12M'][j+12] < 0) \
                    and (result['Dif_L3M_L12M_abs'][j] < result['Dif_L3M_L12M_abs'][j+12]) \
                    and (result['Dif_L3M_L12M'][j+12] != None) \
                    and (result['Dif_L3M_L12M'][j + 1] < 0) \
                    and (result['Dif_L3M_L12M_abs'][j] < result['Dif_L3M_L12M_abs'][j + 1]) \
                    and (result['Dif_L3M_L12M'][j + 1] != None):
                result.loc[j, 'i_spread_L3M_L12M'] = 'narrow'
            elif (result['Dif_L3M_L12M'][j] < 0 and result['Dif_L3M_L12M'][j+12] < 0) \
                    and (result['Dif_L3M_L12M_abs'][j] > result['Dif_L3M_L12M_abs'][j+12]) \
                    and (result['Dif_L3M_L12M'][j+12] != None) \
                    and (result['Dif_L3M_L12M'][j + 1] < 0) \
                    and (result['Dif_L3M_L12M_abs'][j] > result['Dif_L3M_L12M_abs'][j+1]) \
                    and (result['Dif_L3M_L12M'][j+1] != None):
                result.loc[j, 'i_spread_L3M_L12M'] = 'wider'
            elif (result['Dif_L3M_L12M'][j] > 0 and result['Dif_L3M_L12M'][j+12] > 0) \
                    and (result['Dif_L3M_L12M_abs'][j] > result['Dif_L3M_L12M_abs'][j+12]) \
                    and (result['Dif_L3M_L12M'][j+12] != None) \
                    and (result['Dif_L3M_L12M'][j + 1] > 0) \
                    and (result['Dif_L3M_L12M_abs'][j] > result['Dif_L3M_L12M_abs'][j+1]) \
                    and (result['Dif_L3M_L12M'][j] != None):
                        result.loc[j, 'i_spread_L3M_L12M'] = 'wider'
            elif (result['Dif_L3M_L12M'][j] > 0 and result['Dif_L3M_L12M'][j+12] > 0) \
                    and (result['Dif_L3M_L12M_abs'][j] < result['Dif_L3M_L12M_abs'][j+12]) \
                    and (result['Dif_L3M_L12M'][j+12] != None) \
                    and (result['Dif_L3M_L12M'][j + 1] > 0) \
                    and (result['Dif_L3M_L12M_abs'][j] < result['Dif_L3M_L12M_abs'][j+1]) \
                    and (result['Dif_L3M_L12M'][j+1] != None):
                result.loc[j, 'i_spread_L3M_L12M'] = 'narrow'

            if (result['Dif_L3M_L36M'][j] < 0) and (result['Dif_L3M_L36M'][j+12] < 0) \
                    and (result['Dif_L3M_L36M_abs'][j] < result['Dif_L3M_L36M_abs'][j+12]) \
                    and (result['Dif_L3M_L36M'][j+12] != None) \
                    and (result['Dif_L3M_L36M'][j + 1] < 0) \
                    and (result['Dif_L3M_L36M_abs'][j] < result['Dif_L3M_L36M_abs'][j + 1]) \
                    and (result['Dif_L3M_L36M'][j + 1] != None):
                result.loc[j, 'i_spread_L3M_L36M'] = 'narrow'
            elif (result['Dif_L3M_L36M'][j] < 0 and result['Dif_L3M_L36M'][j+12] < 0) \
                    and (result['Dif_L3M_L36M_abs'][j] > result['Dif_L3M_L36M_abs'][j+12]) \
                    and (result['Dif_L3M_L36M'][j+12] != None) \
                    and (result['Dif_L3M_L36M'][j + 1] < 0) \
                    and (result['Dif_L3M_L36M_abs'][j] > result['Dif_L3M_L36M_abs'][j+1]) \
                    and (result['Dif_L3M_L36M'][j+1] != None):
                result.loc[j, 'i_spread_L3M_L36M'] = 'wider'
            elif (result['Dif_L3M_L36M'][j] > 0 and result['Dif_L3M_L36M'][j+12] > 0) \
                    and (result['Dif_L3M_L36M_abs'][j] > result['Dif_L3M_L36M_abs'][j+12]) \
                    and (result['Dif_L3M_L36M'][j+12] != None) \
                    and (result['Dif_L3M_L36M'][j + 1] > 0) \
                    and (result['Dif_L3M_L36M_abs'][j] > result['Dif_L3M_L36M_abs'][j+1]) \
                    and (result['Dif_L3M_L36M'][j+1] != None):
                result.loc[j, 'i_spread_L3M_L36M'] = 'wider'
            elif (result['Dif_L3M_L36M'][j] > 0 and result['Dif_L3M_L36M'][j+12] > 0) \
                    and (result['Dif_L3M_L36M_abs'][j] < result['Dif_L3M_L36M_abs'][j+12]) \
                    and (result['Dif_L3M_L36M'][j+12] != None) \
                    and (result['Dif_L3M_L36M'][j + 1] > 0) \
                    and (result['Dif_L3M_L36M_abs'][j] < result['Dif_L3M_L36M_abs'][j+1]) \
                    and (result['Dif_L3M_L36M'][j+1] != None):
                result.loc[j, 'i_spread_L3M_L36M'] = 'narrow'

            if (result['Dif_L12M_L36M'][j] < 0) and (result['Dif_L12M_L36M'][j+12] < 0) \
                    and (result['Dif_L12M_L36M_abs'][j] < result['Dif_L12M_L36M_abs'][j+12]) \
                    and (result['Dif_L12M_L36M'][j+12] != None) \
                    and (result['Dif_L12M_L36M'][j + 1] < 0) \
                    and (result['Dif_L12M_L36M_abs'][j] < result['Dif_L12M_L36M_abs'][j + 1]) \
                    and (result['Dif_L12M_L36M'][j + 1] != None):
                result.loc[j, 'i_spread_L12M_L36M'] = 'narrow'
            elif (result['Dif_L12M_L36M'][j] < 0 and result['Dif_L12M_L36M'][j+12] < 0) \
                    and (result['Dif_L12M_L36M_abs'][j] > result['Dif_L12M_L36M_abs'][j+12]) \
                    and (result['Dif_L12M_L36M'][j+12] != None) \
                    and (result['Dif_L12M_L36M'][j + 1] < 0) \
                    and (result['Dif_L12M_L36M_abs'][j] > result['Dif_L12M_L36M_abs'][j+1]) \
                    and (result['Dif_L12M_L36M'][j+1] != None):
                result.loc[j, 'i_spread_L12M_L36M'] = 'wider'
            elif (result['Dif_L12M_L36M'][j] > 0 and result['Dif_L12M_L36M'][j+12] > 0) \
                    and (result['Dif_L12M_L36M_abs'][j] > result['Dif_L12M_L36M_abs'][j+12]) \
                    and (result['Dif_L12M_L36M'][j+12] != None) \
                    and (result['Dif_L12M_L36M'][j + 1] > 0) \
                    and (result['Dif_L12M_L36M_abs'][j] > result['Dif_L12M_L36M_abs'][j+1]) \
                    and (result['Dif_L12M_L36M'][j+1] != None):
                result.loc[j, 'i_spread_L12M_L36M'] = 'wider'
            elif (result['Dif_L12M_L36M'][j] > 0 and result['Dif_L12M_L36M'][j+12] > 0) \
                    and (result['Dif_L12M_L36M_abs'][j] < result['Dif_L12M_L36M_abs'][j+12]) \
                    and (result['Dif_L12M_L36M'][j+12] != None) \
                    and (result['Dif_L12M_L36M'][j + 1] > 0) \
                    and (result['Dif_L12M_L36M_abs'][j] < result['Dif_L12M_L36M_abs'][j+1]) \
                    and (result['Dif_L12M_L36M'][j+1] != None):
                result.loc[j, 'i_spread_L12M_L36M'] = 'narrow'

    # calculation: 统计 “趋势 / 价差的出现次数”
    # count of Spread level / Trend
    result['TrendUp_L3M_cnt_L12M'] = np.nan
    result['TrendUp_L12M_cnt_L12M'] = np.nan
    result['TrendUp_L36M_cnt_L12M'] = np.nan
    result['TrendDn_L3M_cnt_L12M'] = np.nan
    result['TrendDn_L12M_cnt_L12M'] = np.nan
    result['TrendDn_L36M_cnt_L12M'] = np.nan

    result['TrendUp_L3M_cnt_L36M'] = np.nan
    result['TrendUp_L12M_cnt_L36M'] = np.nan
    result['TrendUp_L36M_cnt_L36M'] = np.nan
    result['TrendDn_L3M_cnt_L36M'] = np.nan
    result['TrendDn_L12M_cnt_L36M'] = np.nan
    result['TrendDn_L36M_cnt_L36M'] = np.nan

    result['SpreadNarrow_3v12_cnt_L12M'] = np.nan
    result['SpreadNarrow_3v36_cnt_L12M'] = np.nan
    result['SpreadNarrow_12v36_cnt_L12M'] = np.nan
    result['SpreadWider_3v12_cnt_L12M'] = np.nan
    result['SpreadWider_3v36_cnt_L12M'] = np.nan
    result['SpreadWider_12v36_cnt_L12M'] = np.nan

    result['SpreadNarrow_3v12_cnt_L36M'] = np.nan
    result['SpreadNarrow_3v36_cnt_L36M'] = np.nan
    result['SpreadNarrow_12v36_cnt_L36M'] = np.nan
    result['SpreadWider_3v12_cnt_L36M'] = np.nan
    result['SpreadWider_3v36_cnt_L36M'] = np.nan
    result['SpreadWider_12v36_cnt_L36M'] = np.nan
    for i in resultIndex:
        if i <= resultIndex[-1] - 12 - 12:
            result.loc[i, 'TrendUp_L3M_cnt_L12M'] = (result[i: i+12]['i_Trend_L3M'] == 'up').sum()
            result.loc[i, 'TrendUp_L12M_cnt_L12M'] = (result[i: i+12]['i_Trend_L12M'] == 'up').sum()
            result.loc[i, 'TrendUp_L36M_cnt_L12M'] = (result[i: i+12]['i_Trend_L36M'] == 'up').sum()
            result.loc[i, 'TrendDn_L3M_cnt_L12M'] = (result[i: i+12]['i_Trend_L3M'] == 'down').sum()
            result.loc[i, 'TrendDn_L12M_cnt_L12M'] = (result[i: i+12]['i_Trend_L12M'] == 'down').sum()
            result.loc[i, 'TrendDn_L36M_cnt_L12M'] = (result[i: i+12]['i_Trend_L36M'] == 'down').sum()

            result.loc[i, 'SpreadNarrow_3v12_cnt_L12M'] = (result[i: i+12]['i_spread_L3M_L12M'] == 'narrow').sum()
            result.loc[i, 'SpreadNarrow_3v36_cnt_L12M'] = (result[i: i+12]['i_spread_L3M_L36M'] == 'narrow').sum()
            result.loc[i, 'SpreadNarrow_12v36_cnt_L12M'] = (result[i: i+12]['i_spread_L12M_L36M'] == 'narrow').sum()
            result.loc[i, 'SpreadWider_3v12_cnt_L12M'] = (result[i: i+12]['i_spread_L3M_L12M'] == 'wider').sum()
            result.loc[i, 'SpreadWider_3v36_cnt_L12M'] = (result[i: i+12]['i_spread_L3M_L36M'] == 'wider').sum()
            result.loc[i, 'SpreadWider_12v36_cnt_L12M'] = (result[i: i+12]['i_spread_L12M_L36M'] == 'wider').sum()

        if i <= resultIndex[-1] - 12 - 36:
            result.loc[i, 'TrendUp_L3M_cnt_L36M'] = (result[i: i+36]['i_Trend_L3M'] == 'up').sum()
            result.loc[i, 'TrendUp_L12M_cnt_L36M'] = (result[i: i+36]['i_Trend_L12M'] == 'up').sum()
            result.loc[i, 'TrendUp_L36M_cnt_L36M'] = (result[i: i+36]['i_Trend_L36M'] == 'up').sum()
            result.loc[i, 'TrendDn_L3M_cnt_L36M'] = (result[i: i+36]['i_Trend_L3M'] == 'down').sum()
            result.loc[i, 'TrendDn_L12M_cnt_L36M'] = (result[i: i+36]['i_Trend_L12M'] == 'down').sum()
            result.loc[i, 'TrendDn_L36M_cnt_L36M'] = (result[i: i+36]['i_Trend_L36M'] == 'down').sum()

            result.loc[i, 'SpreadNarrow_3v12_cnt_L36M'] = (result[i: i+36]['i_spread_L3M_L12M'] == 'narrow').sum()
            result.loc[i, 'SpreadNarrow_3v36_cnt_L36M'] = (result[i: i+36]['i_spread_L3M_L36M'] == 'narrow').sum()
            result.loc[i, 'SpreadNarrow_12v36_cnt_L36M'] = (result[i: i+36]['i_spread_L12M_L36M'] == 'narrow').sum()
            result.loc[i, 'SpreadWider_3v12_cnt_L36M'] = (result[i: i+36]['i_spread_L3M_L12M'] == 'wider').sum()
            result.loc[i, 'SpreadWider_3v36_cnt_L36M'] = (result[i: i+36]['i_spread_L3M_L36M'] == 'wider').sum()
            result.loc[i, 'SpreadWider_12v36_cnt_L36M'] = (result[i: i+36]['i_spread_L12M_L36M'] == 'wider').sum()

    # 合并所有国家的分析结果
    if EquityObs == 0:
        resultCombineStk = result.copy()
    else:
        resultCombineStk = pd.concat([resultCombineStk, result], ignore_index=True, axis=0)
# round to 2 decimal
MacroEquity = resultCombineStk.copy()

# pivot_table drops columns whose values are entirely NaN.  Recreate the
# optional volume outputs before rounding/final column selection.
for col in ['Vol_b_L12M', 'Vol_L12M_b_Chg', 'Vol_L12M_b_Chg_pct']:
    if col not in MacroEquity.columns:
        MacroEquity[col] = np.nan

print("\n========== MacroEquity DEBUG ==========")

check_cols = [
    'Vol_b_L12M',
    'Vol_L12M_b_Chg',
    'Vol_L12M_b_Chg_pct'
]

print("MacroEquity columns:")
print(MacroEquity.columns.tolist())

print("\nMissing columns:")
print([col for col in check_cols if col not in MacroEquity.columns])

print("\n原始 Combine_Spot_Future_3:")
for col in check_cols:
    if col in Combine_Spot_Future_3.columns:
        print(
            col,
            "总行数:", len(Combine_Spot_Future_3),
            "非空数量:", Combine_Spot_Future_3[col].notna().sum()
        )
    else:
        print(col, "原始数据中不存在")

print("========================================")

round2Col = ['Adj_Price', 'Adj_Price_L3M', 'Adj_Price_L12M', 'Adj_Price_L36M', 'Vol_b_L12M',
             'Adj_Price_Chg', 'Adj_Price_L12M_Chg',	'Vol_L12M_b_Chg', 'Adj_Price_Chg_pct', 'Adj_Price_L12M_Chg_pct', 'Vol_L12M_b_Chg_pct',
             'Dif_L3M_L12M',	'Dif_L12M_L36M', 'Dif_L3M_L36M', 'Dif_L3M_L12M_abs', 'Dif_L12M_L36M_abs', 'Dif_L3M_L36M_abs']
MacroEquity[round2Col] = MacroEquity[round2Col].round(2)

# standardize final bond output
outputMacroEquityCol = ['Country', 'Date', 'NoObs', 'Adj_Price', 'Adj_Price_L3M', 'Adj_Price_L12M', 'Adj_Price_L36M', 'Vol_b_L12M',
                        'Adj_Price_Chg', 'Adj_Price_L12M_Chg',	'Vol_L12M_b_Chg', 'Adj_Price_Chg_pct', 'Adj_Price_L12M_Chg_pct', 'Vol_L12M_b_Chg_pct',
                        'Dif_L3M_L12M',	'Dif_L12M_L36M', 'Dif_L3M_L36M', 'Dif_L3M_L12M_abs', 'Dif_L12M_L36M_abs', 'Dif_L3M_L36M_abs',
                        'i_Trend_L3M', 'i_Trend_L12M', 'i_Trend_L36M', 'i_spread_L3M_L12M',	'i_spread_L3M_L36M', 'i_spread_L12M_L36M',
                        'TrendUp_L3M_cnt_L12M', 'TrendUp_L12M_cnt_L12M', 'TrendUp_L36M_cnt_L12M', 'TrendDn_L3M_cnt_L12M', 'TrendDn_L12M_cnt_L12M', 'TrendDn_L36M_cnt_L12M',
                        'TrendUp_L3M_cnt_L36M',	'TrendUp_L12M_cnt_L36M', 'TrendUp_L36M_cnt_L36M', 'TrendDn_L3M_cnt_L36M', 'TrendDn_L12M_cnt_L36M', 'TrendDn_L36M_cnt_L36M',
                        'SpreadNarrow_3v12_cnt_L12M', 'SpreadNarrow_3v36_cnt_L12M', 'SpreadNarrow_12v36_cnt_L12M', 'SpreadWider_3v12_cnt_L12M', 'SpreadWider_3v36_cnt_L12M', 'SpreadWider_12v36_cnt_L12M',
                        'SpreadNarrow_3v12_cnt_L36M', 'SpreadNarrow_3v36_cnt_L36M',	'SpreadNarrow_12v36_cnt_L36M', 'SpreadWider_3v12_cnt_L36M',	'SpreadWider_3v36_cnt_L36M', 'SpreadWider_12v36_cnt_L36M'
                        ]
MacroEquity = MacroEquity[outputMacroEquityCol]
MacroEquity = MacroEquity.reindex(outputMacroEquityCol, axis=1)  # reorder column to fit into combine_template  # reorder column to fit into combine_template
print('obs of MacroBond is: %i' % len(MacroEquity))

'''''''''''''''
Final output
'''''''''''''''
Name_Equity = ('Macro_Equity_' + str(Update))
# csvName_Equity = (FinalDataPath + Name_Equity + '.csv')
xlsName_Equity = (FinalDataPath + Name_Equity + '.xlsx')

# MacroEquity.to_csv(csvName_Equity, index=False)  # 1 min: for quick check and debug
MacroEquity.to_excel(xlsName_Equity, index=False)  # 2 mins: for formal use, since smaller data size and no dif datatype issues
