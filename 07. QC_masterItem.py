import numpy as np
import pandas as pd
import util.aSetting as aSetting
import warnings
warnings.filterwarnings('ignore')

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
XLSQC = aSetting.XLSQC

combine_manipulation = FinalDataPath + "combine_manipulation_" + str(Update) + ".xlsx"
current_clean = FinalDataPath + "combine_clean_" + str(Update) + ".xlsx"

'''
QC 1.
# compare with test case, if not match for certain %, build indicator and drop it
'''
benchmarkDate = 202610
CurrentRawData = pd.read_excel(current_clean, engine='openpyxl')
CurrentRawData = CurrentRawData[CurrentRawData['Date'] <= benchmarkDate][['FilePath', 'Subcategory', 'Date', 'Price', 'Volume_k']]

benchmark_clean = 'Data\\data_' + str(benchmarkDate) + "\\FinalData\\combine_clean_" + str(benchmarkDate) + ".xlsx"
BenchmarkData = pd.read_excel(benchmark_clean, engine='openpyxl')
BenchmarkData = BenchmarkData[['FilePath', 'Date', 'Price', 'Volume_k']]
BenchmarkData.rename(columns={'Price': 'Price_BM', 'Volume_k': 'Volume_k_BM'}, inplace=True)

Combine = pd.merge(CurrentRawData, BenchmarkData, how='outer', on=['FilePath', 'Date'], indicator=True)
Combine['vs_Price'] = np.nan
Combine['vs_Volume'] = np.nan
Combine.loc[Combine['Price'] == Combine['Price_BM'], 'vs_Price'] = 1
Combine.loc[Combine['Volume_k'] == Combine['Volume_k_BM'], 'vs_Volume'] = 1

matchRate = Combine.groupby('FilePath')['FilePath'].agg(np.sum)

'''
create empty error table
'''
ErrorData = []
column = ['no', 'item', 'details', 'Error_Indicator']
# 初始化 ErrorTable 为一个空的 DataFrame
ErrorTable = pd.DataFrame(columns=column)

column = ['no', 'item', 'details', 'Error_Indicator']
# ErrorTable = pd.concat([ErrorTable, dataError], ignore_index=True)

'''
QC 2a. check if combine key "FilePath" unique
'''
no = '2a'
item = 'FilePath should be unique'
# from 02.investing_clean
MasterSheet = pd.read_excel(MasterSheet, sheet_name='RawData', engine='openpyxl')
print(MasterSheet.head())  # cat: check if load in data

# 'SpotPrice_Futures_Indices',
namesOrg = MasterSheet[['NoSecurity', 'Choose_to_Download', 'Choose_to_DataCombine', 'Choose_to_ModelDev', 'Update_time', 'FilePath',
                        'Subcategory',  'Country', 'Terms_year', 'Terms', 'Directory', 'Subdirectory1', 'Subdirectory2',
                        'Name_investing', 'Description', 'StartDate', 'EndDate']]   # cat: choose the variables we need
Download = namesOrg[(namesOrg['Choose_to_Download'] == "yes") & (namesOrg['Update_time'].astype(str) == "202107")].reset_index(drop=True)
TotalItemNo = MasterSheet.shape[0]
DownloadNo = Download.shape[0]
UniqueDownloadNo = len(Download['FilePath'].unique())
str = ("TotalItemNo: %i, FilePath: %i, unique FilePath: %i " % (TotalItemNo, DownloadNo, UniqueDownloadNo))
if DownloadNo == UniqueDownloadNo:
    ErrorIndicator = 'no'
else:
    ErrorIndicator = 'yes'
dataError = pd.DataFrame([[no, item, str, ErrorIndicator]], columns=column)
ErrorTable = pd.concat([ErrorTable, dataError], ignore_index=True)
print(ErrorTable)


'''
# check update time for each sheet, some issue sheet need to update to most update, some need to monitor and decide later
'''
MasterSheet = aSetting.MasterSheet
SheetList = ['RawData', 'i_data', 'i_event', 'i_extremeValue', 'i_yield_price', 'Yield_price_reverse', 'i_missing', 'i_dup']
# for sheet in SheetList:
#     sheetName = 'Master_' + sheet
#     print(sheetName)
#     sheetName = pd.DataFrame(data=None, index=None, columns=None)
#     data = pd.read_excel(MasterSheet, sheetname=sheet, engine='openpyxl')
#     sheetName = data.copy()
#
#     print(sheetName.head())
#
# Master_Yield_price_reverse.head()

'''
QC 3a.
# check 'RawData' update item
# check 'RawData' Update_EndDate == 'yes' has all the update at column 'Update_time'
'''
no = '3a'
item = 'sheet_RawData should has all the Update_time'
sheetName = SheetList[0]
print(sheetName)
sheetName = pd.read_excel(MasterSheet, sheet_name=sheetName, engine='openpyxl')
TotalItemNo = len(sheetName)
UpdateTimeNo = len(sheetName[sheetName['Update_time'] == Update])
str = ("No of TotalItem: %i, No of Update_time: %i" % (TotalItemNo, UpdateTimeNo))
if TotalItemNo == UpdateTimeNo:
    ErrorIndicator = 'no'
else:
    ErrorIndicator = 'yes'
dataError = pd.DataFrame([[no, item, str, ErrorIndicator]], columns=column)
ErrorTable = pd.concat([ErrorTable, dataError], ignore_index=True)
print(ErrorTable)

'''
QC 3b.
# check successfuly download item
'''
no = '3b'
item = 'QC.csv should has all item downloaded successfully'
Master_Download = pd.read_excel(XLSQC, sheet_name='QcDownload', engine='openpyxl')
ToBeDownloadNo = len(Master_Download[Master_Download['Choose_to_Download'].str.lower() == 'yes'])
SuccessfulDownloadNo = len(Master_Download[Master_Download['QC_download'] == 'successful'])
str = ("No of To Be Downloaded Item: %i, No of Successfully Download Item: %i" % (ToBeDownloadNo, SuccessfulDownloadNo))
if ToBeDownloadNo == SuccessfulDownloadNo:
    ErrorIndicator = 'no'
else:
    ErrorIndicator = 'yes'
dataError = pd.DataFrame([[no, item, str, ErrorIndicator]], columns=column)
ErrorTable = pd.concat([ErrorTable, dataError], ignore_index=True)
print(ErrorTable)


'''
QC 3c.
# check 'i_data' update item
'''
no = '3c'
item = 'sheet i_data should has all the update Issue_EndDate'
sheetName = SheetList[1]
print(sheetName)
sheetName = pd.read_excel(MasterSheet, sheet_name=sheetName, engine='openpyxl')
i_condition = (sheetName['i_data'].str.lower() == 'yes')
ToBeUpdate_condition = (sheetName['Update_EndDate'].astype('str').str.lower() == 'yes')
Update_condition = (sheetName['Issue_EndDate'].str[-6:].astype(int) == Update)
ToBeUpdateNo = len(sheetName[i_condition & ToBeUpdate_condition])
UpdateTimeNo_iData = len(sheetName[Update_condition & i_condition & ToBeUpdate_condition])
str = ("No of ToBeUpdated Item: %i, No of Update Item: %i" % (ToBeUpdateNo, UpdateTimeNo_iData))
if ToBeUpdateNo == UpdateTimeNo_iData:
    ErrorIndicator = 'no'
else:
    ErrorIndicator = 'yes'
dataError = pd.DataFrame([[no, item, str, ErrorIndicator]], columns=column)
ErrorTable = pd.concat([ErrorTable, dataError], ignore_index=True)
print(ErrorTable)

'''
QC 3d.
# check 'i_event' update item
'''
no = '3d'
item = 'sheet i_event should has all the update Issue_EndDate'
sheetName = SheetList[2]
print(sheetName)
sheetName = pd.read_excel(MasterSheet, sheet_name=sheetName, engine='openpyxl')
i_condition = (sheetName['i_event'].str.lower() == 'yes')
ToBeUpdate_condition = (sheetName['Update_EndDate'].astype('str').str.lower() == 'yes')
Update_condition = (sheetName['Issue_EndDate'].str[-6:].astype(int) == Update)
ToBeUpdateNo = len(sheetName[i_condition & ToBeUpdate_condition])
UpdateTimeNo_iEvent = len(sheetName[Update_condition & i_condition & ToBeUpdate_condition])
str = ("No of ToBeUpdated Item: %i, No of Update Item: %i" % (ToBeUpdateNo, UpdateTimeNo_iEvent))
if ToBeUpdateNo == UpdateTimeNo_iData:
    ErrorIndicator = 'no'
else:
    ErrorIndicator = 'yes'
dataError = pd.DataFrame([[no, item, str, ErrorIndicator]], columns=column)
ErrorTable = pd.concat([ErrorTable, dataError], ignore_index=True)
print(ErrorTable)

'''
QC 3e.
# check 'i_extremeValue' update item
'''
no = '3e'
item = 'sheet i_extremeValue should has all the update Issue_EndDate'
sheetName = SheetList[3]
print(sheetName)
sheetName = pd.read_excel(MasterSheet, sheet_name=sheetName, engine='openpyxl')
i_condition = (sheetName['i_extremeValue'].str.lower() == 'yes')
ToBeUpdate_condition = (sheetName['Update_EndDate'].astype('str').str.lower() == 'yes')
Update_condition = (sheetName['Issue_EndDate'].str[-6:].astype(int) == Update)
ToBeUpdateNo = len(sheetName[i_condition & ToBeUpdate_condition])
UpdateTimeNo_iExtremeValue = len(sheetName[Update_condition & i_condition & ToBeUpdate_condition])
str = ("No of ToBeUpdated Item: %i, No of Update Item: %i" % (ToBeUpdateNo, UpdateTimeNo_iExtremeValue))
if ToBeUpdateNo == UpdateTimeNo_iExtremeValue:
    ErrorIndicator = 'no'
else:
    ErrorIndicator = 'yes'
dataError = pd.DataFrame([[no, item, str, ErrorIndicator]], columns=column)
ErrorTable = pd.concat([ErrorTable, dataError], ignore_index=True)
print(ErrorTable)

'''
QC 3f.
# check 'i_yield_price' update item
'''
no = '3f'
item = 'sheet i_yield_price should has all the update Issue_EndDate'
sheetName = SheetList[4]
print(sheetName)
sheetName = pd.read_excel(MasterSheet, sheet_name=sheetName, engine='openpyxl')
i_condition = (sheetName['i_yield_price'].str.lower() == 'yes')
ToBeUpdate_condition = (sheetName['Update_EndDate'].astype('str').str.lower() == 'yes')
Update_condition = (sheetName['Issue_EndDate'].str[-6:].astype(int) == Update)
ToBeUpdateNo = len(sheetName[i_condition & ToBeUpdate_condition])
UpdateTimeNo_iYieldPrice = len(sheetName[Update_condition & i_condition & ToBeUpdate_condition])
str = ("No of ToBeUpdated Item: %i, No of Update Item: %i" % (ToBeUpdateNo, UpdateTimeNo_iYieldPrice))
if ToBeUpdateNo == UpdateTimeNo_iYieldPrice:
    ErrorIndicator = 'no'
else:
    ErrorIndicator = 'yes'
dataError = pd.DataFrame([[no, item, str, ErrorIndicator]], columns=column)
ErrorTable = pd.concat([ErrorTable, dataError], ignore_index=True)
print(ErrorTable)

'''
QC 3g.
# check 'Yield_price_reverse' update item
'''
no = '2g'
item = 'sheet Yield_price_reverse should has all the update Issue_EndDate'
sheetName = SheetList[5]
print(sheetName)
sheetName = pd.read_excel(MasterSheet, sheet_name=sheetName, engine='openpyxl')
i_condition = (sheetName['Yield_price_reverse'].isnull() == False)
ToBeUpdate_condition = (sheetName['Update_EndDate'].astype('str').str.lower() == 'yes')
Update_condition = (sheetName['Issue_EndDate'].str[-6:] == Update)
ToBeUpdateNo = len(sheetName[i_condition & ToBeUpdate_condition])
UpdateTimeNo_YieldPriceReverse = len(sheetName[Update_condition & i_condition & ToBeUpdate_condition])
str = ("No of ToBeUpdated Item: %i, No of Update Item: %i" % (ToBeUpdateNo, UpdateTimeNo_YieldPriceReverse))
if ToBeUpdateNo == UpdateTimeNo_YieldPriceReverse:
    ErrorIndicator = 'no'
else:
    ErrorIndicator = 'yes'
dataError = pd.DataFrame([[no, item, str, ErrorIndicator]], columns=column)
ErrorTable = pd.concat([ErrorTable, dataError], ignore_index=True)
print(ErrorTable)


'''
# below data is for QC 4X
'''
data = pd.read_excel(combine_manipulation, engine='openpyxl')

'''
QC 4a.
# check 'i_missing' update date --> reflect back to mastersheet tab"i_missing"
'''
col = ['Directory', 'Subdirectory1', 'Subdirectory2', 'Subcategory', 'i_missing', 'Issue_StartDate', 'Issue_EndDate']
dataMissing = data[data['i_missing'] == 1].copy()
MinMaxMissing = dataMissing.groupby('FilePath')['Date'].agg([np.min, np.max])

'''
QC 4b.
# check 'i_dup' update date --> reflect back to mastersheet tab"i_dup"
'''
col = ['Directory', 'Subdirectory1', 'Subdirectory2', 'Subcategory', 'i_dup', 'Issue_StartDate', 'Issue_EndDate']
dataDup = data[data['i_dup'] == 1].copy()
MinMaxDup = dataDup.groupby('FilePath')['Date'].agg([np.min, np.max])

'''
QC 4c.
# check 'i_dup' update date --> reflect back to mastersheet tab"i_dup"
'''
MinMaxDate = data.groupby('FilePath')['Date'].agg([np.min, np.max])
