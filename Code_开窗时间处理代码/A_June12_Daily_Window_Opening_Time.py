import pandas as pd
import datetime
import numpy as np


w_data=pd.read_excel("J:\研究生课题\##小论文写作\期刊论文\Main_Part\May20_2018_Revisions\Main\Results_Add\Daily_Window_Opening_Time\结果\计算单户开窗时间/J1_opening_time_Handled.xlsx")
#print(w_data)
w_data["Time"]=pd.to_datetime(w_data["Time"],format="%Y/%m/%d %H:%M:%S")
statusarray=np.array(["status"])
#statusarray.columns=["Status"]

print(w_data)
for i in range(len(w_data["Time"])-1):
    if w_data.iloc[i,1]==1:
        opent=w_data.iloc[i,0]
        while opent<=w_data.iloc[i+1,0]:
            print(opent)
            opent=opent+datetime.timedelta(minutes=1)
            statusarray=np.row_stack([statusarray,opent])

print(statusarray)
statusarray=pd.DataFrame(statusarray)
writer=pd.ExcelWriter("J:\研究生课题\##小论文写作\期刊论文\Main_Part\May20_2018_Revisions\Main\Results_Add\Daily_Window_Opening_Time\结果\计算单户开窗时间/J1_Mid_Results.xlsx")
statusarray.to_excel(writer,"A")
writer.save()
writer.close()
# print(statusarray)