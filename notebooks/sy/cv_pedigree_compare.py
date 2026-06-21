"""혈통/EBV 파생 "넣은거 vs 안넣은거" 누수안전 CV (메모리 경량판: usecols·dev only·float32)."""
import pandas as pd, numpy as np, time, gc, json
import lightgbm as lgb
from sklearn.metrics import f1_score
from sklearn.model_selection import train_test_split
ROOT="/Users/chosun-nhn04/IdeaProjects/Hanwoo_Contest"; SEED,SUB,M=42,200_000,50
NFOLDS=2
t0=time.time(); log=lambda *a: print(f"[{time.time()-t0:6.0f}s]",*a,flush=True)

FEATS=json.loads('["WEIGHT","AGE","death_count","C2023","C2024","C2025","AREA","days_total","days_양호","days_주의","days_경고","days_위험","rn_day_mean","ws_davg_mean","ta_min_mean","ratio_고온","ratio_강더위","ratio_위험","abatt_year","abatt_month","abatt_quarter","birth_year","birth_month","fattening_days","daily_gain","Age_squared","age_optimal","heat_high","density","death_rate","density_x_heat","KPN_NO_freq","FATHER_CATTLE_NO_freq","MOTHER_ANIMAL_NO_freq","sido_경기도","sido_경상남도","sido_경상북도","sido_광주광역시","sido_대구광역시","sido_대전광역시","sido_부산광역시","sido_서울특별시","sido_세종특별자치시","sido_울산광역시","sido_인천광역시","sido_전라남도","sido_전북특별자치도","sido_제주특별자치도","sido_충청남도","sido_충청북도","abatt_season_겨울","abatt_season_봄","abatt_season_여름","birth_season_겨울","birth_season_봄","birth_season_여름","sex_수","sex_암"]')

log("로드(usecols+float32)…")
use=FEATS+["CATTLE_NO","grade_num","INSFAT","JUDGE_DATE"]
dt={f:"float32" for f in FEATS}
df=pd.read_csv(f"{ROOT}/data/processed/3_eda/step11_features.csv",usecols=use,dtype=dt,encoding="utf-8-sig")
df["grade_num"]=pd.to_numeric(df["grade_num"],errors="coerce").astype("float32")
df["INSFAT"]=pd.to_numeric(df["INSFAT"],errors="coerce").astype("float32")
df["jd"]=pd.to_datetime(df["JUDGE_DATE"],errors="coerce"); df.drop(columns=["JUDGE_DATE"],inplace=True)
sp=pd.read_csv(f"{ROOT}/data/processed/4_model/split_assignment.csv",encoding="utf-8-sig")
df=df.merge(sp,on="CATTLE_NO",how="left")
df=df[df["split"]=="dev"].reset_index(drop=True); del sp; gc.collect()
log(f"dev {len(df):,}행")

lin=pd.read_csv(f"{ROOT}/data/raw/hanwoo_lineage_0612.csv",dtype=str,usecols=["CATTLE_NO","KPN_NO","FATHER_CATTLE_NO","MOTHER_ANIMAL_NO","F_GFATHER_CATTLE_NO"])
for c in lin.columns: lin[c]=lin[c].str.strip()
df=df.merge(lin.drop_duplicates("CATTLE_NO"),on="CATTLE_NO",how="left"); del lin; gc.collect()
g=pd.read_excel(f"{ROOT}/data/raw/KPN 유전능력 자료.xlsx"); g["KPN명호"]=g["KPN명호"].astype(str).str.strip()
g=g[~g["KPN명호"].isin(["nan",""])].drop_duplicates("KPN명호")
emap={"근내지방도 육종가":"ebv_marbling","도체중 육종가":"ebv_cwt","등심단면적 육종가":"ebv_rea","등지방두께 육종가":"ebv_backfat","12개월체중 육종가":"ebv_wt12"}
g2=g[["KPN명호"]+list(emap)].rename(columns={"KPN명호":"KPN_NO",**emap}); EBV=list(emap.values())
for c in EBV: g2[c]=pd.to_numeric(g2[c],errors="coerce").astype("float32")
df=df.merge(g2,on="KPN_NO",how="left"); df["has_ebv"]=df["ebv_marbling"].notna().astype("int8"); del g,g2; gc.collect()
log("준비완료, CV 시작")

fold_idx=[(np.where(df["fold"]!=k)[0],np.where(df["fold"]==k)[0]) for k in range(NFOLDS)]
def te(st,keys,prior,self_val=None):
    s=keys.map(st["sum"]).astype("float64"); c=keys.map(st["count"]).astype("float64")
    if self_val is not None: s=s-self_val.values; c=c-1
    return ((s+M*prior)/(c+M)).fillna(prior)
def build(tr,tgt,self_in,qual):
    f={}; pg=tr["grade_num"].mean(); phi=(tr["grade_num"]>=10).mean(); pm=tr["INSFAT"].mean()
    sv=tgt["grade_num"] if self_in else None; svh=(tgt["grade_num"]>=10).astype(float) if self_in else None; svm=tgt["INSFAT"] if self_in else None
    for key,nm in [("KPN_NO","kpn"),("FATHER_CATTLE_NO","father"),("F_GFATHER_CATTLE_NO","fgf")]:
        f[f"{nm}_prog_grade"]=te(tr.groupby(key)["grade_num"].agg(["sum","count"]),tgt[key],pg,sv).values
    f["kpn_prog_highgrade"]=te(tr.assign(hi=(tr["grade_num"]>=10).astype(float)).groupby("KPN_NO")["hi"].agg(["sum","count"]),tgt["KPN_NO"],phi,svh).values
    f["kpn_freq2"]=tgt["KPN_NO"].map(tr["KPN_NO"].value_counts()).fillna(0).values
    rec=tr.dropna(subset=["CATTLE_NO"]).drop_duplicates("CATTLE_NO").set_index("CATTLE_NO")
    for key,nm in [("MOTHER_ANIMAL_NO","mother"),("FATHER_CATTLE_NO","father")]:
        ajd=tgt[key].map(rec["jd"]); ag=tgt[key].map(rec["grade_num"])
        gt=np.where(ajd.values<tgt["jd"].values,ag.values,np.nan)
        f[f"{nm}_own_grade"]=np.where(np.isnan(gt),pg,gt); f[f"has_{nm}_grade"]=(~np.isnan(gt)).astype("int8")
    for c in EBV: f[c]=tgt[c].fillna(tr[c].mean()).values
    f["has_ebv"]=tgt["has_ebv"].values
    if qual:
        f["kpn_prog_marbling"]=te(tr.groupby("KPN_NO")["INSFAT"].agg(["sum","count"]),tgt["KPN_NO"],pm,svm).values
        ajd=tgt["MOTHER_ANIMAL_NO"].map(rec["jd"]); am=tgt["MOTHER_ANIMAL_NO"].map(rec["INSFAT"])
        gm=np.where(ajd.values<tgt["jd"].values,am.values,np.nan); f["mother_own_marbling"]=np.where(np.isnan(gm),pm,gm)
    return pd.DataFrame(f,index=tgt.index)
P=dict(objective="multiclass",num_class=16,learning_rate=0.05,n_estimators=500,max_depth=8,num_leaves=101,
    min_child_samples=66,subsample=0.888,subsample_freq=1,colsample_bytree=0.813,reg_alpha=0.001,reg_lambda=0.70,
    class_weight="balanced",n_jobs=-1,verbosity=-1,force_col_wise=True,random_state=SEED)
from sklearn.metrics import accuracy_score
EBV_F=EBV+["has_ebv"]
PROG=["kpn_prog_grade","father_prog_grade","fgf_prog_grade","kpn_prog_highgrade","kpn_freq2"]
OWN=["mother_own_grade","has_mother_grade","father_own_grade","has_father_grade"]
CFGS={"C0 기존":[], "A +EBV만":EBV_F, "B +자손TE만":PROG, "C +조상본인등급만":OWN,
      "D +자손TE+EBV":PROG+EBV_F, "E +자손마블링":PROG+EBV_F+["kpn_prog_marbling"]}
res={k:{"f1":[],"acc":[]} for k in CFGS}
for k,(tri,vai) in enumerate(fold_idx):
    tr=df.iloc[tri]; va=df.iloc[vai]
    trs,_=train_test_split(tr,train_size=min(SUB,len(tr)),stratify=tr["grade_num"],random_state=SEED)
    pt=build(tr,trs,True,True).reset_index(drop=True); pv=build(tr,va,False,True).reset_index(drop=True)
    X0t,X0v=trs[FEATS].reset_index(drop=True),va[FEATS].reset_index(drop=True)
    yt,yv=trs["grade_num"].astype(int).values,va["grade_num"].astype(int).values
    for name,ex in CFGS.items():
        Xt=pd.concat([X0t,pt[ex]],axis=1) if ex else X0t; Xv=pd.concat([X0v,pv[ex]],axis=1) if ex else X0v
        m=lgb.LGBMClassifier(**P); m.fit(Xt,yt,eval_set=[(Xv,yv)],eval_metric="multi_logloss",callbacks=[lgb.early_stopping(40,verbose=False),lgb.log_evaluation(0)])
        p=np.asarray(m.predict(Xv)).ravel(); f=f1_score(yv,p,average="macro"); a=accuracy_score(yv,p)
        res[name]["f1"].append(f); res[name]["acc"].append(a); log(f"fold{k} [{name}] F1={f:.4f} Acc={a:.4f}")
    del tr,va,trs,pt,pv; gc.collect()
print("\n===== 격리 결과 (%d-fold 평균) ====="%NFOLDS,flush=True)
base=np.mean(res["C0 기존"]["f1"])
for name in CFGS:
    f=np.mean(res[name]["f1"]); ac=np.mean(res[name]["acc"])
    print(f"  {name:16s} MacroF1={f:.4f} ({f-base:+.4f})  Acc={ac:.4f}",flush=True)
log("완료")
