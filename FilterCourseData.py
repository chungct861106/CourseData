
#-*- coding : utf-8 -*-
# coding: utf-8

year = ['104', '105', '106', '107', '108', '109']
import pandas as pd
data = pd.read_csv("CourseData.csv", encoding="utf-8", index_col=0)
import json

depart = dict()
with open("Department.json", encoding='utf-8') as f:
    depart = json.load(f)


fullname = list(depart.keys())
data = data.dropna(subset=['COUDPTN'])
data = data.loc[data["CREDIT"]>=2]
data = data.loc[data["S_TERM"]==1]




#%%
# check = data.iloc[0:5]

# num_depart ={
#         "1":'文學院',
#         "2": '理學院',
#         "3": '社會科學院',
#         "4": '醫學院',
#         "5": '工學院',
#         "6": '生物資源暨農學院',
#         "7": '管理學院',
#         "8": '公共衛生學院',
#         "9": '電機資訊學院',
#         "A": '法律學院',
#         "B": '生命科學院',
#         "E": '其他',
#         "H": '其他',
#         "Z": '其他' 
#     }

# #%%
# Coukind = list(set(num_depart.values()))
# Coukind.remove("其他")
# data["COUCAT"] = [depart[cat] for cat in data['COUDPTN']]
# data["STDCAT"] = [num_depart[num[3]] for num in data["REG_NO"]]
# data["STDDPTNO"] = [num[3:6] for num in data["REG_NO"]]
# data["COUDPTNO"] = [code.split(" ")[0] if code.find(" ")!=-1 else code[0:3] for code in data["COU_CODE"]]

# #%%


# col = ['COUDPTN', 'COUCAT', 'COUDPTNO', 'DPT_SCNAME', "STDCAT", 'CREDIT', "STDDPTNO", 'COU_CNAME', 'GRADE', 'SCORE_GP', 'S_YEAR']
# df = data.filter(items=col)


# #%%
# df = df.loc[df['STDCAT'] != "其他"]
# mask = [name in Coukind for name in df["COUCAT"]]
# df = df.loc[mask]



# #%%

# df.to_csv("NewCourseData-1.csv")
