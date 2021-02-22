import pandas as pd
import matplotlib
import matplotlib.pyplot as plt
import numpy as np
import sys
import os
import matplotlib.font_manager as fm
fpath = 'TaipeiSansTCBeta-Bold.ttf'
prop = fm.FontProperties(fname=fpath)
data = pd.read_csv("NewCourseData-1.csv", encoding="utf-8", index_col=0)
#%%
mask = [True if name[-1] == "系" else False for name in data["COUDPTN"]]
data = data.loc[mask]
mask = [True if name[-1] == "系" else False for name in data["DPT_SCNAME"]]
data = data.loc[mask]
year = data["S_YEAR"].unique().tolist()
ydf = list()
outdp = list()
outColl = list()
for y in range(len(year)):
    df = data.loc[data["S_YEAR"] == year[y]]
    ydf.append(len(df))
    outdp.append(len(df.loc[df["COUDPTN"]!=df["DPT_SCNAME"]]))
    outColl.append(len(df.loc[df["COUCAT"]!=df["STDCAT"]]))
    

plt.plot(year, ydf, label="ALL")
plt.plot(year, outdp, label="Cross Department")
plt.plot(year, outColl, label="Cross Collage")
plt.legend(ncol=3, bbox_to_anchor=(0.1, -0.2),
                   loc='lower left', fontsize='small')
plt.title("跨院系選修趨勢", fontproperties=prop)


#%%

# #%%
# #  ['COUDPTN', 'COUCAT', 'DPT_SCNAME', 'STDCAT', 'CREDIT', 'COU_CNAME', 'GRADE', 'SCORE_GP', 'S_YEAR']


# Years = data["S_YEAR"].unique().tolist()
# CollegeKind = sorted(data["COUCAT"].unique().tolist())

# sheets = dict()
# for college in CollegeKind:
#     cdata = data.loc[data["COUCAT"] == college]
#     CourseKind = sorted(cdata["COUDPTN"].unique().tolist())
#     CourseKind = [name for name in CourseKind if name[-1] == "系"]


    
    
#     #%%
#     col = ['origin', 'foregin', 'foregin percent']
#     foregin = list()
#     college_diff = list()
#     ratio = list()
#     ratio_col = list()
    
#     #%%
#     for year in Years:
#         df = cdata.loc[cdata["S_YEAR"] == year]
#         fore = list()
#         coldiff = list()
#         r = list()
#         r2 = list()
#         for coll in CourseKind:
#             totalstd = df.loc[df["COUDPTN"] == coll]
#             f = len(totalstd.loc[totalstd["COUDPTNO"] != totalstd["STDDPTNO"]])
#             c = len(totalstd.loc[totalstd["COUCAT"] != totalstd["STDCAT"]])
#             fore.append(f)
#             coldiff.append(c)
#             if len(totalstd)>0:
#                 r.append(f/len(totalstd))
#                 r2.append(c/len(totalstd))
#             else:
#                 r.append(0)
#                 r2.append(0)
#         foregin.append(fore)
#         college_diff.append(coldiff)
#         ratio.append(r)
#         ratio_col.append(r2)
#     #%%
#     sheet = dict()
#     sheet["外系人數趨勢"] = pd.DataFrame(foregin, columns=CourseKind, index=Years)
#     sheet["外系比例趨勢"] = pd.DataFrame(ratio, columns=CourseKind, index=Years)
#     sheet["外院人數趨勢"] = pd.DataFrame(college_diff, columns=CourseKind, index=Years)
#     sheet["外院比例趨勢"] = pd.DataFrame(ratio_col, columns=CourseKind, index=Years)
#     sheets[college] = sheet


# #%%

# def getfig(df, title):
#     df = df.sort_index()
#     fpath = 'TaipeiSansTCBeta-Bold.ttf'
#     prop = fm.FontProperties(fname=fpath)
#     labels = df.columns.tolist()
#     if len(labels) == 0:
#         return None
#     pos = np.arange(len(labels))
#     labelsdatas = df.values.tolist()
#     fig, ax = plt.subplots()
#     width = 0.5
#     for i in range(len(df)):
#         ax.bar(pos - width/2 + i * width/(len(df.index)-1), labelsdatas[i], width/(len(df.index)-1), label=df.index[i])
#     ax.set_title(title, fontproperties=prop)
#     ax.set_xticks(pos)
#     ax.set_xticklabels(labels, fontproperties=prop)
#     ax.legend(ncol=len(labelsdatas), bbox_to_anchor=(0, -0.2),
#                   loc='lower left', fontsize='small')
#     fig.tight_layout()
#     return fig

# figures = dict()
# for sheet in sheets:
#     with pd.ExcelWriter(sheet+'-學期.xlsx', engine='xlsxwriter') as writer:
#         workbook = writer.book
#         for name in sheets[sheet]:
#             df = sheets[sheet][name]
#             figures[sheet+"-"+name] = getfig(df, name)
#             df.sort_index().to_excel(writer, sheet_name=name)
#             worksheet = writer.sheets[name]
#             chart = workbook.add_chart({'type': 'column'})

#             # Configure the series of the chart from the dataframe data.
#             for row_num in range(1, len(Years) + 1):
#                 chart.add_series({
#                     'name':       [name, row_num, 0],
#                     'categories': [name, 0, 1, 0, len(df.columns)],
#                     'values':     [name, row_num, 1, row_num ,len(df.columns)],
#                     'gap':        300,
#                 })
#             chart.set_title({'name': sheet + name[0:4]})
#             chart.set_y_axis({'major_gridlines': {'visible': False}})
#             worksheet.insert_chart('A11', chart)
#         writer.save()


# #%%

# for fig in figures:
#     figures[fig].savefig(fig+"-下學期.png")