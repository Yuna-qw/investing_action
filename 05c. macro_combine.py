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

# Alternative:
Name_Bond = ('Macro_Bond_' + str(Update))
xlsName_Bond = (FinalDataPath + Name_Bond + '.xlsx')
Name_Equity = ('Macro_Equity_' + str(Update))
xlsName_Equity = (FinalDataPath + Name_Equity + '.xlsx')
MacroBond_2 = pd.read_excel(xlsName_Bond, engine='openpyxl')
MacroEquity_2 = pd.read_excel(xlsName_Equity, engine='openpyxl')

for col in MacroBond_2.columns:
    if col not in ['Country', 'Date']:
        newCol = 'bd_'+col
        print(newCol)
        MacroBond_2.rename(columns={col: newCol}, inplace=True)

for col in MacroEquity_2.columns:
    if col not in ['Country', 'Date']:
        newCol = 'stk_'+col
        print(newCol)
        MacroEquity_2.rename(columns={col: newCol}, inplace=True)

MacroBond_2 = MacroBond_2.sort_values(['Country', 'Date'], ascending=[True, False])
MacroEquity_2 = MacroEquity_2.sort_values(['Country', 'Date'], ascending=[True, False])
len(MacroBond_2)
len(MacroEquity_2)

# get bond / equity exist counrtry indictor
BondCountry = pd.DataFrame(MacroBond_2.groupby('Country')['Country'].nunique())
BondCountry.rename(columns={'Country': 'i_BondCountry'}, inplace=True)
BondCountry = BondCountry.reset_index()

EquityCountry = pd.DataFrame(MacroEquity_2.groupby('Country')['Country'].nunique())
EquityCountry.rename(columns={'Country': 'i_EquityCountry'}, inplace=True)
EquityCountry = EquityCountry.reset_index()

i_Country = pd.merge(BondCountry, EquityCountry, how='outer', on='Country')

# merge bond and equity macro data
MacroCombine = pd.merge(MacroBond_2, MacroEquity_2, how='outer', on=['Country', 'Date'], indicator=True)
len(MacroCombine)  # 19720

# merge with i_country data
MacroCombine_2 = pd.merge(MacroCombine, i_Country, how='outer', on='Country')
MacroCombine_2.loc[MacroCombine_2['_merge'] == 'left_only', 'MacroSource'] = 'bond'
MacroCombine_2.loc[MacroCombine_2['_merge'] == 'right_only', 'MacroSource'] = 'equity'
MacroCombine_2.loc[MacroCombine_2['_merge'] == 'both', 'MacroSource'] = 'both'
MacroCombine_2 = MacroCombine_2.drop(['_merge'], axis=1)
MacroCombine_2 = MacroCombine_2.sort_values(['Country', 'Date'], ascending=[True, False])
MacroCombine_2.shape  # (19720, 108)

'''''''''''''''
Final output
'''''''''''''''
Name_Macro = ('Macro_' + str(Update))
# csvName_Macro = (FinalDataPath + Name_Macro + '.csv')
xlsName_Macro = (FinalDataPath + Name_Macro + '.xlsx')

# MacroCombine_2.to_csv(csvName_Macro, index=False)    # 0.5 min: for quick check and debug
MacroCombine_2.to_excel(xlsName_Macro, index=False)   # 1 mins: for formal use, since smaller data size and no dif datatype issues
