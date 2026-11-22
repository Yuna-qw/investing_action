import numpy as np
import pandas as pd
import util.aSetting as aSetting
import warnings

warnings.filterwarnings('ignore')

# 1. 初始化和数据加载
Update = aSetting.Update
FinalDataPath = aSetting.FinalDataPath

# alternative get source data
Name_Mani = ('combine_manipulation_' + str(Update))
xlsName_Mani = (FinalDataPath + Name_Mani + '.xlsx')

# 检查文件是否存在，防止出错
try:
    Combine_Spot_Future_2 = pd.read_excel(xlsName_Mani, engine='openpyxl')
    print(f"成功加载文件: {xlsName_Mani}")
except FileNotFoundError:
    print(f"错误: 找不到文件 {xlsName_Mani}。请检查路径和文件名。")
    exit()
except Exception as e:
    print(f"加载 Excel 文件时发生错误: {e}")
    exit()

BondCol = ['NoSecurity', 'Subcategory', 'NoObs', 'Date',
           'i_data', 'i_event', 'i_extremeValue', 'i_yield_price', 'Yield_price_reverse',
           'i_missing', 'i_dup',
           'Country', 'Macro_Bond',
           'Yield_price', 'Yield', 'Yield_L12M', 'Yield_Chg', 'Yield_L12M_Chg', 'Yield_Chg_pct', 'Yield_L12M_Chg_pct'
           ]
# 筛选条件：Macro_Bond字段非空（即属于宏观债券范畴的数据）
Combine_Spot_Future_3 = Combine_Spot_Future_2[Combine_Spot_Future_2['Macro_Bond'].isnull() == 0][BondCol].copy()
print('Macro Bond Data Observation:  %i' % Combine_Spot_Future_3.shape[0])

# get yield data by country by date, with select ShortTerm, MidTerm and LongTerm yield
Category = ['Country', 'Date']
# 数值列：需要汇总计算的收益率相关指标
ValueColumn = ['Yield', 'Yield_L12M', 'Yield_Chg', 'Yield_L12M_Chg', 'Yield_Chg_pct', 'Yield_L12M_Chg_pct']
# 生成透视表：按国家+日期分组，将不同期限（Macro_Bond）的收益率指标取均值汇总
Bond_Transpose = pd.pivot_table(Combine_Spot_Future_3, values=ValueColumn, index=Category, columns='Macro_Bond',
                                aggfunc=np.mean)
# 债券转置，重命名列
RenameColumn = [
    'LT_Yield', 'MT_Yield', 'ST_Yield',
    'LT_Yield_Chg', 'MT_Yield_Chg', 'ST_Yield_Chg',
    'LT_Yield_Chg_pct', 'MT_Yield_Chg_pct', 'ST_Yield_Chg_pct',
    'LT_Yield_L12M', 'MT_Yield_L12M', 'ST_Yield_L12M',
    'LT_Yield_L12M_Chg', 'MT_Yield_L12M_Chg', 'ST_Yield_L12M_Chg',
    'LT_Yield_L12M_Chg_pct', 'MT_Yield_L12M_Chg_pct', 'ST_Yield_L12M_Chg_pct'
]
Bond_Transpose.columns = RenameColumn
Bond_Transpose_2 = Bond_Transpose.reset_index()
# 注意：这里按日期降序排序 (Ascending=False)，方便后续的滞后计算，因为索引0是最新数据
Bond_Transpose_2 = Bond_Transpose_2.sort_values(['Country', 'Date'], ascending=[True, False])
print(f"Pivot Table Observation: {len(Bond_Transpose_2)}")

# ----------------------------------------------------
# 2. 按国家逐一计算指标 (主要修复部分)
# ----------------------------------------------------

UniqueCountryCat = Bond_Transpose_2['Country'].unique()

# --- FIX 1: 在循环外初始化一个空列表用于收集每个国家的最终结果 ---
all_country_results = []
# ----------------------------------------------------------------------

for BondObs in range(0, len(UniqueCountryCat)):
    Country = UniqueCountryCat[BondObs]
    print(f"{Country}")

    left = Bond_Transpose_2[Bond_Transpose_2['Country'] == Country].copy()
    result = left.sort_values(['Date'], ascending=[False])

    # reset_index
    # 重置索引并移除旧索引列，确保索引是连续的数字 [0, 1, 2, ...]
    result = result.reset_index(drop=True)
    resultIndex = result.index.values.tolist()
    result['NoObs'] = result.index + 1

    # calculation:计算 “收益率利差”
    # 简化代码，直接计算并赋值
    result['Dif_StMt_Yield_L12M'] = result['ST_Yield_L12M'] - result['MT_Yield_L12M']
    result['Dif_MtLt_Yield_L12M'] = result['MT_Yield_L12M'] - result['LT_Yield_L12M']
    result['Dif_StLt_Yield_L12M'] = result['ST_Yield_L12M'] - result['LT_Yield_L12M']

    result['Dif_StMt_Yield_L12M_abs'] = result['Dif_StMt_Yield_L12M'].abs()
    result['Dif_MtLt_Yield_L12M_abs'] = result['Dif_MtLt_Yield_L12M'].abs()
    result['Dif_StLt_Yield_L12M_abs'] = result['Dif_StLt_Yield_L12M'].abs()

    # calculation: 判断 “收益率趋势方向” 判断 “利差趋势”
    # 初始化趋势和利差列
    result['i_Trend_ST'] = ''
    result['i_Trend_MT'] = ''
    result['i_Trend_LT'] = ''
    result['i_spread_StMt'] = ''
    result['i_spread_StLt'] = ''
    result['i_spread_MtLt'] = ''

    # 初始化计数列（即使在循环内计算，在此处初始化可以避免 Pandas SettingWithCopyWarning）
    count_cols = [
        'TrendUp_cnt_ST_L12M', 'TrendUp_cnt_MT_L12M', 'TrendUp_cnt_LT_L12M',
        'TrendDn_cnt_ST_L12M', 'TrendDn_cnt_MT_L12M', 'TrendDn_cnt_LT_L12M',
        'TrendUp_cnt_ST_L36M', 'TrendUp_cnt_MT_L36M', 'TrendUp_cnt_LT_L36M',
        'TrendDn_cnt_ST_L36M', 'TrendDn_cnt_MT_L36M', 'TrendDn_cnt_LT_L36M',
        'SpreadNarrow_StMt_cnt_L12M', 'SpreadNarrow_StLt_cnt_L12M', 'SpreadNarrow_MtLt_cnt_L12M',
        'SpreadWider_StMt_cnt_L12M', 'SpreadWider_StLt_cnt_L12M', 'SpreadWider_MtLt_cnt_L12M',
        'SpreadNarrow_StMt_cnt_L36M', 'SpreadNarrow_StLt_cnt_L36M', 'SpreadNarrow_MtLt_cnt_L36M',
        'SpreadWider_StMt_cnt_L36M', 'SpreadWider_StLt_cnt_L36M', 'SpreadWider_MtLt_cnt_L36M'
    ]
    for col in count_cols:
        result[col] = np.nan

    # 循环计算 i_Trend 和 i_spread (需要前 12 个月的数据)
    max_index_for_12m_calc = resultIndex[-1] - 12

    for j in resultIndex:
        if j <= max_index_for_12m_calc:
            # 确保当前和滞后数据不是 NaN，否则比较会出错
            st_curr = result.loc[j, 'ST_Yield_L12M']
            st_m1 = result.loc[j + 1, 'ST_Yield_L12M']
            st_m12 = result.loc[j + 12, 'ST_Yield_L12M']
            mt_curr = result.loc[j, 'MT_Yield_L12M']
            mt_m1 = result.loc[j + 1, 'MT_Yield_L12M']
            mt_m12 = result.loc[j + 12, 'MT_Yield_L12M']
            lt_curr = result.loc[j, 'LT_Yield_L12M']
            lt_m1 = result.loc[j + 1, 'LT_Yield_L12M']
            lt_m12 = result.loc[j + 12, 'LT_Yield_L12M']

            # Trend ST
            if pd.notna(st_curr) and pd.notna(st_m1) and pd.notna(st_m12):
                if (st_curr - st_m1) < 0 and (st_curr - st_m12) < 0:
                    result.loc[j, 'i_Trend_ST'] = 'down'
                elif (st_curr - st_m1) > 0 and (st_curr - st_m12) > 0:
                    result.loc[j, 'i_Trend_ST'] = 'up'

            # Trend MT
            if pd.notna(mt_curr) and pd.notna(mt_m1) and pd.notna(mt_m12):
                if (mt_curr - mt_m1) < 0 and (mt_curr - mt_m12) < 0:
                    result.loc[j, 'i_Trend_MT'] = 'down'
                elif (mt_curr - mt_m1) > 0 and (mt_curr - mt_m12) > 0:
                    result.loc[j, 'i_Trend_MT'] = 'up'

            # Trend LT
            if pd.notna(lt_curr) and pd.notna(lt_m1) and pd.notna(lt_m12):
                if (lt_curr - lt_m1) < 0 and (lt_curr - lt_m12) < 0:
                    result.loc[j, 'i_Trend_LT'] = 'down'
                elif (lt_curr - lt_m1) > 0 and (lt_curr - lt_m12) > 0:
                    result.loc[j, 'i_Trend_LT'] = 'up'


            # --------------------- 利差趋势 i_spread ---------------------

            def calculate_spread_trend(df, j, diff_col, abs_col, spread_col):
                curr_diff = df.loc[j, diff_col]
                m1_diff = df.loc[j + 1, diff_col]
                m12_diff = df.loc[j + 12, diff_col]
                curr_abs = df.loc[j, abs_col]
                m1_abs = df.loc[j + 1, abs_col]
                m12_abs = df.loc[j + 12, abs_col]

                # 确保所有涉及的值都不是 NaN
                if pd.isna(curr_diff) or pd.isna(m1_diff) or pd.isna(m12_diff) or \
                        pd.isna(curr_abs) or pd.isna(m1_abs) or pd.isna(m12_abs):
                    return

                # 利差方向一致性
                is_negative_consistent = curr_diff < 0 and m1_diff < 0 and m12_diff < 0
                is_positive_consistent = curr_diff > 0 and m1_diff > 0 and m12_diff > 0

                if is_negative_consistent or is_positive_consistent:
                    # Narrow (绝对值变小)
                    if curr_abs < m1_abs and curr_abs < m12_abs:
                        df.loc[j, spread_col] = 'narrow'
                    # Wider (绝对值变大)
                    elif curr_abs > m1_abs and curr_abs > m12_abs:
                        df.loc[j, spread_col] = 'wider'


            calculate_spread_trend(result, j, 'Dif_StMt_Yield_L12M', 'Dif_StMt_Yield_L12M_abs', 'i_spread_StMt')
            calculate_spread_trend(result, j, 'Dif_StLt_Yield_L12M', 'Dif_StLt_Yield_L12M_abs', 'i_spread_StLt')
            calculate_spread_trend(result, j, 'Dif_MtLt_Yield_L12M', 'Dif_MtLt_Yield_L12M_abs', 'i_spread_MtLt')

    # calculation:统计 “趋势 / 利差的出现次数”

    # L12M 计数需要至少 i+12 存在 (即索引 j 到 j+11，共 12 个月)
    max_index_for_L12M_count = resultIndex[-1] - 12
    # L36M 计数需要至少 i+36 存在 (即索引 j 到 j+35，共 36 个月)
    max_index_for_L36M_count = resultIndex[-1] - 36

    # 计数 i 从 0 开始，统计未来 12 或 36 个月 (从 i+1 开始到 i+N)
    for i in resultIndex:
        # L12M 计数 (统计未来 12 个月)
        if i <= max_index_for_L12M_count:
            # 这里的切片是 result[i+1: i+13]，即未来 12 个月
            result.loc[i, 'TrendUp_cnt_ST_L12M'] = (result[i + 1: i + 13]['i_Trend_ST'] == 'up').sum()
            result.loc[i, 'TrendUp_cnt_MT_L12M'] = (result[i + 1: i + 13]['i_Trend_MT'] == 'up').sum()
            result.loc[i, 'TrendUp_cnt_LT_L12M'] = (result[i + 1: i + 13]['i_Trend_LT'] == 'up').sum()
            result.loc[i, 'TrendDn_cnt_ST_L12M'] = (result[i + 1: i + 13]['i_Trend_ST'] == 'down').sum()
            result.loc[i, 'TrendDn_cnt_MT_L12M'] = (result[i + 1: i + 13]['i_Trend_MT'] == 'down').sum()
            result.loc[i, 'TrendDn_cnt_LT_L12M'] = (result[i + 1: i + 13]['i_Trend_LT'] == 'down').sum()

            result.loc[i, 'SpreadNarrow_StMt_cnt_L12M'] = (result[i + 1: i + 13]['i_spread_StMt'] == 'narrow').sum()
            result.loc[i, 'SpreadNarrow_StLt_cnt_L12M'] = (result[i + 1: i + 13]['i_spread_StLt'] == 'narrow').sum()
            result.loc[i, 'SpreadNarrow_MtLt_cnt_L12M'] = (result[i + 1: i + 13]['i_spread_MtLt'] == 'narrow').sum()
            result.loc[i, 'SpreadWider_StMt_cnt_L12M'] = (result[i + 1: i + 13]['i_spread_StMt'] == 'wider').sum()
            result.loc[i, 'SpreadWider_StLt_cnt_L12M'] = (result[i + 1: i + 13]['i_spread_StLt'] == 'wider').sum()
            result.loc[i, 'SpreadWider_MtLt_cnt_L12M'] = (result[i + 1: i + 13]['i_spread_MtLt'] == 'wider').sum()

        # L36M 计数 (统计未来 36 个月)
        if i <= max_index_for_L36M_count:
            # 这里的切片是 result[i+1: i+37]，即未来 36 个月
            result.loc[i, 'TrendUp_cnt_ST_L36M'] = (result[i + 1: i + 37]['i_Trend_ST'] == 'up').sum()
            result.loc[i, 'TrendUp_cnt_MT_L36M'] = (result[i + 1: i + 37]['i_Trend_MT'] == 'up').sum()
            result.loc[i, 'TrendUp_cnt_LT_L36M'] = (result[i + 1: i + 37]['i_Trend_LT'] == 'up').sum()
            result.loc[i, 'TrendDn_cnt_ST_L36M'] = (result[i + 1: i + 37]['i_Trend_ST'] == 'down').sum()
            result.loc[i, 'TrendDn_cnt_MT_L36M'] = (result[i + 1: i + 37]['i_Trend_MT'] == 'down').sum()
            result.loc[i, 'TrendDn_cnt_LT_L36M'] = (result[i + 1: i + 37]['i_Trend_LT'] == 'down').sum()

            result.loc[i, 'SpreadNarrow_StMt_cnt_L36M'] = (result[i + 1: i + 37]['i_spread_StMt'] == 'narrow').sum()
            result.loc[i, 'SpreadNarrow_StLt_cnt_L36M'] = (result[i + 1: i + 37]['i_spread_StLt'] == 'narrow').sum()
            result.loc[i, 'SpreadNarrow_MtLt_cnt_L36M'] = (result[i + 1: i + 37]['i_spread_MtLt'] == 'narrow').sum()
            result.loc[i, 'SpreadWider_StMt_cnt_L36M'] = (result[i + 1: i + 37]['i_spread_StMt'] == 'wider').sum()
            result.loc[i, 'SpreadWider_StLt_cnt_L36M'] = (result[i + 1: i + 37]['i_spread_StLt'] == 'wider').sum()
            result.loc[i, 'SpreadWider_MtLt_cnt_L36M'] = (result[i + 1: i + 37]['i_spread_MtLt'] == 'wider').sum()

    # --- FIX 2: 将当前国家的结果追加到列表中 (而不是错误的合并/覆盖) ---
    all_country_results.append(result.copy())
    # ----------------------------------------------------------------------

# ----------------------------------------------------
# 3. 最终合并与输出
# ----------------------------------------------------

# --- FIX 3: 在循环结束后，一次性合并所有国家的结果 ---
if all_country_results:
    resultCombineBd = pd.concat(all_country_results, ignore_index=True)
    print(f"\n所有国家数据合并成功，总观察次数: {len(resultCombineBd)}")
else:
    resultCombineBd = pd.DataFrame()
    print("\n警告: 没有数据被合并。")
# ----------------------------------------------------


# round to 2 decimal 保留2位小数
MacroBond = resultCombineBd.copy()
round2Col = ['LT_Yield', 'MT_Yield', 'ST_Yield', 'LT_Yield_L12M', 'MT_Yield_L12M', 'ST_Yield_L12M',
             'LT_Yield_Chg', 'MT_Yield_Chg', 'ST_Yield_Chg', 'LT_Yield_Chg_pct', 'MT_Yield_Chg_pct', 'ST_Yield_Chg_pct',
             'LT_Yield_L12M_Chg', 'MT_Yield_L12M_Chg', 'ST_Yield_L12M_Chg', 'LT_Yield_L12M_Chg_pct',
             'MT_Yield_L12M_Chg_pct', 'ST_Yield_L12M_Chg_pct',
             'Dif_StMt_Yield_L12M', 'Dif_MtLt_Yield_L12M', 'Dif_StLt_Yield_L12M', 'Dif_StMt_Yield_L12M_abs',
             'Dif_MtLt_Yield_L12M_abs', 'Dif_StLt_Yield_L12M_abs']
MacroBond[round2Col] = MacroBond[round2Col].round(2)

# standardize final bond output 规范列顺序
outputMacroCol = ['Country', 'Date', 'NoObs', 'LT_Yield', 'MT_Yield', 'ST_Yield', 'LT_Yield_L12M', 'MT_Yield_L12M',
                  'ST_Yield_L12M',
                  'LT_Yield_Chg', 'MT_Yield_Chg', 'ST_Yield_Chg', 'LT_Yield_Chg_pct', 'MT_Yield_Chg_pct',
                  'ST_Yield_Chg_pct',
                  'LT_Yield_L12M_Chg', 'MT_Yield_L12M_Chg', 'ST_Yield_L12M_Chg', 'LT_Yield_L12M_Chg_pct',
                  'MT_Yield_L12M_Chg_pct', 'ST_Yield_L12M_Chg_pct',
                  'Dif_StMt_Yield_L12M', 'Dif_MtLt_Yield_L12M', 'Dif_StLt_Yield_L12M', 'Dif_StMt_Yield_L12M_abs',
                  'Dif_MtLt_Yield_L12M_abs', 'Dif_StLt_Yield_L12M_abs',
                  'i_Trend_ST', 'i_Trend_MT', 'i_Trend_LT', 'i_spread_StMt', 'i_spread_StLt', 'i_spread_MtLt',
                  'TrendUp_cnt_ST_L12M', 'TrendUp_cnt_MT_L12M', 'TrendUp_cnt_LT_L12M', 'TrendDn_cnt_ST_L12M',
                  'TrendDn_cnt_MT_L12M', 'TrendDn_cnt_LT_L12M',
                  'TrendUp_cnt_ST_L36M', 'TrendUp_cnt_MT_L36M', 'TrendUp_cnt_LT_L36M', 'TrendDn_cnt_ST_L36M',
                  'TrendDn_cnt_MT_L36M', 'TrendDn_cnt_LT_L36M',
                  'SpreadNarrow_StMt_cnt_L12M', 'SpreadNarrow_StLt_cnt_L12M', 'SpreadNarrow_MtLt_cnt_L12M',
                  'SpreadWider_StMt_cnt_L12M', 'SpreadWider_StLt_cnt_L12M', 'SpreadWider_MtLt_cnt_L12M',
                  'SpreadNarrow_StMt_cnt_L36M', 'SpreadNarrow_StLt_cnt_L36M', 'SpreadNarrow_MtLt_cnt_L36M',
                  'SpreadWider_StMt_cnt_L36M', 'SpreadWider_StLt_cnt_L36M', 'SpreadWider_MtLt_cnt_L36M']
MacroBond = MacroBond[outputMacroCol]
MacroBond = MacroBond.reindex(outputMacroCol, axis=1)
# reorder column to fit into combine_template
print('obs of MacroBond is: %i' % len(MacroBond))

'''
Final output
'''
Name_Bond = ('Macro_Bond_' + str(Update))
xlsName_Bond = (FinalDataPath + Name_Bond + '.xlsx')

MacroBond.to_excel(xlsName_Bond, index=False)
print(f"数据已成功输出到文件: {xlsName_Bond}")