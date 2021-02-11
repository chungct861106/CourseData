import pandas as pd
import matplotlib
import matplotlib.pyplot as plt
import numpy as np
import sys
import os
import matplotlib.font_manager as fm
import json


data = pd.read_csv("NewCourseData-2.csv", encoding="utf-8", index_col=0)
sem = "下學期"
#%%

Departs = dict()
with open("Department.json") as f:
    Departs = json.load(f)
depart = dict()
CollegeKind = dict()
for d in Departs:
    if d[-1] != "系":
        continue
    depart[d] = Departs[d]
    if Departs[d] not in CollegeKind:
        CollegeKind[Departs[d]] = [d]
    else:
       CollegeKind[Departs[d]].append(d)
depart["生技系"] = "生命科學院"
depart["生命科學系"] = "生命科學院"
CollegeKind["生命科學院"].append("生技系")
CollegeKind["生命科學院"].append("生命科學系")
mask = [True if name in depart else False for name in data["DPT_SCNAME"]]
data = data.loc[mask]
test = data["DPT_SCNAME"].unique()
#%%
#  ['COUDPTN', 'COUCAT', 'DPT_SCNAME', 'STDCAT', 'CREDIT', 'COU_CNAME', 'GRADE', 'SCORE_GP', 'S_YEAR']


Years = sorted(data["S_YEAR"].unique().tolist())
dirs = dict()
for college in CollegeKind:
    sheets = dict()
    cdata = data.loc[data["STDCAT"] == college]
    StdKind = sorted(cdata["DPT_SCNAME"].unique().tolist())
    StdKind = [name for name in StdKind if name[-1] == "系" and name in CollegeKind[college]]
    for std in StdKind:
        sdata = cdata.loc[cdata["DPT_SCNAME"] == std]
        Coursekind = sorted(sdata["COUDPTN"].unique().tolist(), key=lambda x: -len(sdata.loc[sdata["COUDPTN"] == x]))
        if std in Coursekind:
            Coursekind.remove(std)
        outputdata = [[len(sdata.loc[(sdata["COUDPTN"] == course) & (sdata["S_YEAR"] == year)]) for course in Coursekind] for year in Years]
        output = pd.DataFrame(outputdata, index=Years, columns=Coursekind)
        sheets[std] = output
    dirs[college] = sheets
#%%



def getfig(df, title):
    df = df.sort_index()
    df = df.filter(df.columns[0:10])
    fpath = 'TaipeiSansTCBeta-Bold.ttf'
    prop = fm.FontProperties(fname=fpath)
    labels = df.columns.tolist()
    if len(labels) == 0:
        return None
    pos = np.arange(len(labels))
    labelsdatas = df.values.tolist()
    fig = plt.figure(figsize=[15,6])
    ax = fig.add_subplot(1,1,1)
    width = 0.5
    for i in range(len(df)):
        ax.bar(pos - width/2 + i * width/(len(df.index)-1), labelsdatas[i], width/(len(df.index)-1), label=df.index[i])
    ax.set_title(title, fontproperties=prop)
    ax.set_xticks(pos)
    ax.set_xticklabels(labels, fontproperties=prop)
    ax.legend(ncol=len(labelsdatas), bbox_to_anchor=(0.5, -0.2),
                  loc='lower center', fontsize='small')
    return fig

figures = dict()

for direct in dirs:
    with pd.ExcelWriter('各院學生跨系選修分布趨勢/各系跨學系/{}/{}各系學生跨系選修趨勢-{}.xlsx'.format(direct,direct,sem), engine='xlsxwriter') as writer:
        workbook = writer.book
        for name in dirs[direct]:
            df = dirs[direct][name]
            figure = getfig(df, name)
            figure.savefig('各院學生跨系選修分布趨勢/各系跨學系/{}/{}學生跨系選修趨勢-{}.png'.format(direct, name,sem))
            df.sort_index().to_excel(writer, sheet_name=name)
            worksheet = writer.sheets[name]
            chart = workbook.add_chart({'type': 'column'})
            # Configure the series of the chart from the dataframe data.
            for row_num in range(1, len(Years) + 1):
                chart.add_series({
                    'name':       [name, row_num, 0],
                    'categories': [name, 0, 1, 0, min(len(df.columns), 10)],
                    'values':     [name, row_num, 1, row_num ,min(len(df.columns),10)],
                    'gap':        300,
                })
            chart.set_title({'name': name + "學生跨系前十選修趨勢"})
            chart.set_y_axis({'major_gridlines': {'visible': False}})
            worksheet.insert_chart('A11', chart)
        writer.save()
        

    
