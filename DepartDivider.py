import pandas as pd

data = pd.read_csv("NewCourseData-2.csv", encoding="utf-8", index_col=0)
#%%
#  ['COUDPTN', 'COUCAT', 'DPT_SCNAME', 'STDCAT', 'CREDIT', 'COU_CNAME', 'GRADE', 'SCORE_GP', 'S_YEAR']

Years = data["S_YEAR"].unique().tolist()
CollegeKind = sorted(data["COUCAT"].unique().tolist())

sheets = dict()
for college in CollegeKind:
    cdata = data.loc[data["COUCAT"] == college]
    CourseKind = sorted(cdata["COUDPTN"].unique().tolist())
    # StudentKind = sorted(cdata["DPT_SCNAME"].unique().tolist())
    
    
    
    #%%
    col = ['origin', 'foregin', 'foregin percent']
    foregin = list()
    college_diff = list()
    ratio = list()
    ratio_col = list()
    
    #%%
    for year in Years:
        df = cdata.loc[cdata["S_YEAR"] == year]
        fore = list()
        coldiff = list()
        r = list()
        r2 = list()
        for coll in CourseKind:
            totalstd = df.loc[df["COUDPTN"] == coll]
            f = len(totalstd.loc[totalstd["COUDPTNO"] != totalstd["STDDPTNO"]])
            c = len(totalstd.loc[totalstd["COUCAT"] != totalstd["STDCAT"]])
            fore.append(f)
            coldiff.append(c)
            if len(totalstd)>0:
                r.append(f/len(totalstd))
                r2.append(c/len(totalstd))
            else:
                r.append(0)
                r2.append(0)
        foregin.append(fore)
        college_diff.append(coldiff)
        ratio.append(r)
        ratio_col.append(r2)
    #%%
    sheet = dict()
    sheet["外系人數趨勢"] = pd.DataFrame(foregin, columns=CourseKind, index=Years)
    sheet["外系比例趨勢"] = pd.DataFrame(ratio, columns=CourseKind, index=Years)
    sheet["外院人數趨勢"] = pd.DataFrame(college_diff, columns=CourseKind, index=Years)
    sheet["外院比例趨勢"] = pd.DataFrame(ratio_col, columns=CourseKind, index=Years)
    sheets[college] = sheet


#%%

for sheet in sheets:
    with pd.ExcelWriter(sheet+'-下學期.xlsx') as writer:
        for name in sheets[sheet]:
            sheets[sheet][name].sort_index().to_excel(writer, sheet_name=name)
    