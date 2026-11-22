import datetime
import os
import pandas as pd
'''''''''''''''''''''''''''''''''''''''
Instruction:
1. manually update date
2. manually copy MasterOrigin in running folder / update the 'EndDate' & 'Update_time' in mastersheet
'''''''''''''''''''''''''''''''''''''''
date = datetime.date(2026,11, 1)                        # type in the first day of running month

# 每次跑改个日期
'''''''''''''''''''''''''''''''''''''''
below will be updated automatically
'''''''''''''''''''''''''''''''''''''''
LastPeriod = 202610
Update = 202610

EndDate = date - datetime.timedelta(days=1)			  	# datetime.date(2017, 12, 31) for datetime.date(2018,1,1)
EndDate_Str_1 = EndDate.strftime("%m/%d/%Y")			  	# '31/05/2018'
EndDate_Str_2 = EndDate.strftime("%Y-%m-%d")                # '2022-10-01'
print(EndDate_Str_2)
# Update = EndDate.year * 100 + EndDate.month	  			# 201805
UpdateYear = int(Update/100)							# 2018
UpdateMonth = Update - UpdateYear*100					# 5
# LastPeriod = int(EndDate.year*100 + (EndDate.month-1))

BasePath = 'Data\\data_' + str(Update)
MasterPath = 'Data\\02.Master\\'
MasterOrigin = MasterPath + 'Investing_MasterSheet_v2_' + str(LastPeriod) +'.xlsx'  # origin excel Investing_MasterSheet_v2_202504.xlsx
MasterSheet = MasterPath + 'Investing_MasterSheet_v2_' + str(Update) +'.xlsx'

# excel after convert formula to value, mainly for column 'FilePath'
# 建文件夹 update = 202202 excel里面Update_time也要改成202202
RawDataPath = BasePath + '\\RawData\\'  # set up your own path
FinalDataPath = BasePath + '\\FinalData\\'
LogPath = BasePath + '\\log\\'
XLSQC = LogPath + 'QC_' + str(Update) + '.xlsx'


