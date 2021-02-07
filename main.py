import pandas as pd

data = pd.read_csv("NewCourseData-1.csv", encoding="utf-8", index_col=0)
#%%
#  ['COUDPTN', 'COUCAT', 'DPT_SCNAME', 'STDCAT', 'CREDIT', 'COU_CNAME', 'GRADE', 'SCORE_GP', 'S_YEAR']

Years = data["S_YEAR"].unique().tolist()
CollegeKind = sorted(data["COUCAT"].unique().tolist())


college = "工學院"
cdata = data.loc[data["COUCAT"] == college]
CourseKind = sorted(cdata["COUDPTN"].unique().tolist())
# StudentKind = sorted(cdata["DPT_SCNAME"].unique().tolist())



#%%
col = ['origin', 'foregin', 'foregin percent']
foregin = list()

ratio = list()

#%%
for year in Years:
    df = cdata.loc[cdata["S_YEAR"] == year]
    fore = list()
    r = list()
    for coll in CourseKind:
        totalstd = df.loc[df["COUDPTN"] == coll]
        f = [name for name in totalstd["DPT_SCNAME"] if name.find(coll)==-1]
        fore.append(len(f))
        if len(totalstd)>0:
            r.append(len(f)/len(totalstd))
        else:
            r.append(0)
            
    foregin.append(fore)
    ratio.append(r)
#%%
sheet = dict()
sheet["外系人數趨勢"] = pd.DataFrame(foregin, columns=CourseKind, index=Years)
sheet["外系比例趨勢"] = pd.DataFrame(ratio, columns=CourseKind, index=Years)



    
    