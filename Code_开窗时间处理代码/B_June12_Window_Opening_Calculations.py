import datetime
import pandas as pd
import numpy as np

data=pd.read_excel("J:\研究生课题\##小论文写作\期刊论文\Main_Part\May20_2018_Revisions\Main\Results_Add\Daily_Window_Opening_Time\结果\计算单户开窗时间/J1_Mid_Results_Test.xlsx")
data["Time"]=pd.to_datetime(data["Time"],format="%Y/%m/%d %H:%M:%S")
datetime1=np.array(["A"])
vtime=datetime.datetime(2017,2,1,0,0,0)
opentime_min=list()
datelist=list()
for t in range(30):
    count=0
    print(vtime)
    for i in range(len(data)):
        if (data.iloc[i, 0]>=vtime) & (data.iloc[i, 0]<=vtime+datetime.timedelta(days=1)):
            count=count+1
            print("Date:{}, Min:{}".format(vtime,count))
    datelist.append(vtime)
    vtime=(vtime+datetime.timedelta(days=1))
    opentime_min.append(count)
    print("Date:{}, Min:{}".format(vtime,opentime_min))
opentime_min=pd.DataFrame(opentime_min)
opentime_min.columns=["Time"]
datelist=pd.DataFrame(datelist)
datelist.columns=["Date"]

opentime_min=pd.concat((datelist,opentime_min),axis=1)
opentime_min=pd.DataFrame(opentime_min)
writer=pd.ExcelWriter("J:\研究生课题\##小论文写作\期刊论文\Main_Part\May20_2018_Revisions\Main\Results_Add\Daily_Window_Opening_Time\结果\计算单户开窗时间/J1_Opening_time_Results.xlsx")
opentime_min.to_excel(writer,"Res_J1")
writer.save()
writer.close()