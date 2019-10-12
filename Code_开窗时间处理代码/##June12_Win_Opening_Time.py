# -*- coding: utf-8 -*-
import pandas as pd
import datetime
import re
import os
import numpy as np

# month="03"
#reside = ["Z6"]
reside=["Z3","Z4","Z5","Z6"]
# reside="J6"

month=["02","03","04","05","06","07","08","09","10","11"]
month_day=[28,31,30,31,30,31,31,30,31,30]

# month = ["03"]
# month_day = [31]

dictmonth = dict()
i = 0
for m in month:
    dictmonth[m] = month_day[i]
    i = i + 1

# reside="J6"
for res in reside:
    BRWin_data_add = pd.DataFrame([])
    for m in month:
        if int(m) < 10:
            a = str(m)
            m1 = a[- 1]
        else:
            m1 = m

        dire =  'F:\研究生课题\##小论文写作\期刊论文\Source_Data\KM_Data\KM_M{}/KM{}_M_2017-{}.xlsx'.format(m1,res,m)
        A = pd.ExcelFile(dire)
        Sheetname = A.sheet_names
        print(Sheetname)
        # print(a)
        for i in range(len(Sheetname)):
            if (not re.search('客', Sheetname[i]) == None) and (not re.search('窗', Sheetname[i]) == None):
                LRWindex = i
            if (not re.search('卧', Sheetname[i]) == None) and (not re.search('窗', Sheetname[i]) == None):
                BRWindex = i
            if (not re.search('卧', Sheetname[i]) == None) and (not re.search('门', Sheetname[i]) == None):
                BRDindex = i

        BRWin_data = pd.read_excel(dire, sheetname=BRWindex, parse_cols=[11, 12], skiprows=[0])
        BRWin_data.columns = ["Time", "Status"]
        BRWin_data["Time"] = pd.to_datetime(BRWin_data["Time"], format="%Y/%m/%d %H:%M:%S")
        DA = pd.DataFrame(columns=["Status"])

        for bl in range(len(BRWin_data["Status"])):
            if BRWin_data.iloc[bl, 1] == "open":
                a = int(1)
                BRWin_data.iloc[bl, 1] = a
            else:
                b = int(0)
                BRWin_data.iloc[bl, 1] = b
        print(BRWin_data)

        if len(BRWin_data) == 0:
            continue
        BRWin_add = pd.DataFrame([])
        print(len(BRWin_data))
        while i < (len(BRWin_data) - 2):
            if BRWin_data.iloc[i + 1, 1] == BRWin_data.iloc[i, 1]:
                BRWin_data = BRWin_data.drop(BRWin_data.index[i])
            i = i + 1
        print(BRWin_data)

        initial_time = pd.to_datetime("2017-{}-01 00:00:00".format(m), format="%Y/%m/%d %H:%M:%S")
        end_time = pd.to_datetime("2017-{}-{} 23:59:59".format(m, dictmonth[m]), format="%Y/%m/%d %H:%M:%S")

        init = pd.DataFrame([[initial_time], [BRWin_data.iloc[1, 1]]]).transpose()
        print(init)
        init.columns = ["Time", "Status"]
        end = pd.DataFrame([[end_time], [BRWin_data.iloc[-2, 1]]]).transpose()
        end.columns = ["Time", "Status"]
        print(end)

        BRWin_data = init.append(BRWin_data)
        BRWin_data = BRWin_data.append(end)
        BRWin_data = pd.DataFrame(BRWin_data)
        BRWin_data = BRWin_data.reset_index(drop=True)
        print(BRWin_data)
        BRWin_data_add = pd.concat([BRWin_data_add, BRWin_data],axis=0)

    writer = pd.ExcelWriter("F:\研究生课题\##小论文写作\期刊论文\Main_Part/2018_June12_Results_Add\Win_Opening_Time/{}.xlsx".format(res))
    BRWin_data_add.to_excel(writer, "BR_Win")