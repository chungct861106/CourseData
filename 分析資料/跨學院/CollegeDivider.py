import pandas as pd

data = pd.read_csv("NewCourseData-1.csv", encoding="utf-8", index_col=0)
#%%
#  ['COUDPTN', 'COUCAT', 'DPT_SCNAME', 'STDCAT', 'CREDIT', 'COU_CNAME', 'GRADE', 'SCORE_GP', 'S_YEAR']


CollegeKind = sorted(data["COUCAT"].unique().tolist())
CourseKind = sorted(data["COU_CNAME"].unique().tolist())
StudentKind = sorted(data["DPT_SCNAME"].unique().tolist())
Years = data["S_YEAR"].unique().tolist()


#%%
col = ['origin', 'foregin', 'foregin percent']
foregin = list()
ratio = list()


for year in Years:
    df = data.loc[data["S_YEAR"] == year]
    fore = [len(df.loc[(df["COUCAT"] == coll) & (df["COUCAT"] != df["STDCAT"])]) for coll in CollegeKind]
    total = [len(df.loc[(df["COUCAT"] == coll)]) for coll in CollegeKind]
    foregin.append(fore)
    ratio.append([fore[i]/total[i] for i in range(len(CollegeKind))])

sheet = dict()
#%%
sheet["外院人數趨勢"] = pd.DataFrame(foregin, columns=CollegeKind, index=Years)
sheet["外院比例趨勢"] = pd.DataFrame(ratio, columns=CollegeKind, index=Years)


#%%
with pd.ExcelWriter('學院課程選修趨勢-上學期.xlsx') as writer:
    for name in sheet:
        sheet[name].sort_index().to_excel(writer, sheet_name=name)
    
    