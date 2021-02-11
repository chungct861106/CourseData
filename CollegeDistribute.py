import pandas as pd
import matplotlib
import matplotlib.pyplot as plt
import numpy as np
import sys
import os
import matplotlib.font_manager as fm

data = pd.read_csv("NewCourseData-2.csv", encoding="utf-8", index_col=0)
#%%

StudentCat = data["STDCAT"].unique()
Years = sorted(data["S_YEAR"].unique())
cat_df = dict()
for cat in StudentCat:
    sdata = data.loc[(data["STDCAT"] == cat) & (data["COUCAT"]!=cat)]
    CourseCat = sorted(sdata["COUCAT"].unique())
    outputdata = list()
    for year in Years:
        df = sdata.loc[sdata["S_YEAR"] == year]
        linedata = [len(df.loc[df["COUCAT"]==course]) for course in CourseCat]
        outputdata.append(linedata)
    cat_df[cat] = pd.DataFrame(outputdata, index=Years, columns=CourseCat)
#%%

import sys
import os
import matplotlib.font_manager as fm
import matplotlib
import matplotlib.pyplot as plt
import numpy as np
def getfig(df, title):
    df = df.sort_index()
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
with pd.ExcelWriter('各院學生跨院選修趨勢-下學期.xlsx', engine='xlsxwriter') as writer:
    workbook = writer.book
    for name in cat_df:
        df = cat_df[name]
        figures[name+"學生跨院選修趨勢-下學期"] = getfig(df, name)
        df.sort_index().to_excel(writer, sheet_name=name)
        worksheet = writer.sheets[name]
        chart = workbook.add_chart({'type': 'column'})
        # Configure the series of the chart from the dataframe data.
        for row_num in range(1, len(Years) + 1):
            chart.add_series({
                'name':       [name, row_num, 0],
                'categories': [name, 0, 1, 0, len(df.columns)],
                'values':     [name, row_num, 1, row_num ,len(df.columns)],
                'gap':        300,
            })
        chart.set_title({'name': name + "學生跨院選修趨勢"})
        chart.set_y_axis({'major_gridlines': {'visible': False}})
        worksheet.insert_chart('A11', chart)
    writer.save()
#%%
    
for fig in figures:
    figures[fig].savefig("各院學生跨系選修分布趨勢/" + fig + "-下學期.png")

    
    