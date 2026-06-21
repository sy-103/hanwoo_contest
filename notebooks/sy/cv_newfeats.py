"""신규 파생변수 "넣은거 vs 안넣은거" CV (전부 타깃 미사용=누수0, dev 통계로 정적 계산).
A daily_gain_squared(트리 no-op 증명) / B weight_vs_peer(월령+성별 또래체격)
C daily_gain_z_by_sex(성별내 z) / D lineage_missing / E ratio_고온_late(출하전90일 더위) / F 좋은것 결합."""
import pandas as pd, numpy as np, time, gc, json
import lightgbm as lgb
from sklearn.metrics import f1_score, accuracy_score
from sklearn.model_selection import train_test_split
ROOT="/Users/chosun-nhn04/IdeaProjects/Hanwoo_Contest"; SEED,SUB,NF=42,200_000,2
t0=time.time(); log=lambda *a: print(f"[{time.time()-t0:5.0f}s]",*a,flush=True)

FEATS=json.loads('["WEIGHT","AGE","death_count","C2023","C2024","C2025","AREA","days_total","days_양호","days_주의","days_경고","days_위험","rn_day_mean","ws_davg_mean","ta_min_mean","ratio_고온","ratio_강더위","ratio_위험","abatt_year","abatt_month","abatt_quarter","birth_year","birth_month","fattening_days","daily_gain","Age_squared","age_optimal","heat_high","density","death_rate","density_x_heat","KPN_NO_freq","FATHER_CATTLE_NO_freq","MOTHER_ANIMAL_NO_freq","sido_경기도","sido_경상남도","sido_경상북도","sido_광주광역시","sido_대구광역시","sido_대전광역시","sido_부산광역시","sido_서울특별시","sido_세종특별자치시","sido_울산광역시","sido_인천광역시","sido_전라남도","sido_전북특별자치도","sido_제주특별자치도","sido_충청남도","sido_충청북도","abatt_season_겨울","abatt_season_봄","abatt_season_여름","birth_season_겨울","birth_season_봄","birth_season_여름","sex_수","sex_암","ebv_marbling","ebv_marbling_acc","ebv_cwt","ebv_rea","ebv_backfat","ebv_wt12","has_ebv","mother_own_grade","has_mother_grade","father_own_grade","has_father_grade"]')

log(f"로드(track1 {len(FEATS)}개 + 빌드용)…")
use=FEATS+["CATTLE_NO","grade_num","JUDGE_SEX","ABATT_DATE","stn","KPN_NO"]
df=pd.read_csv(f"{ROOT}/data/processed/3_eda/step11_features.csv",usecols=use,
               dtype={f:"float32" for f in FEATS},encoding="utf-8-sig",low_memory=False)
sp=pd.read_csv(f"{ROOT}/data/processed/4_model/split_assignment.csv",encoding="utf-8-sig")
df=df.merge(sp,on="CATTLE_NO",how="left"); df=df[df["split"]=="dev"].reset_index(drop=True); del sp; gc.collect()
df["grade_num"]=pd.to_numeric(df["grade_num"],errors="coerce")
log(f"dev {len(df):,}")

# ── 신규 피처 (dev 통계, 타깃 미사용) ──
df["daily_gain_squared"]=(df["daily_gain"]**2).astype("float32")                                  # A
df["weight_vs_peer"]=(df["WEIGHT"]/df.groupby(["AGE","JUDGE_SEX"])["WEIGHT"].transform("mean")).astype("float32")  # B
sm=df.groupby("JUDGE_SEX")["daily_gain"].transform("mean"); ss=df.groupby("JUDGE_SEX")["daily_gain"].transform("std")
df["daily_gain_z_by_sex"]=((df["daily_gain"]-sm)/ss).astype("float32")                            # C
df["lineage_missing"]=df["KPN_NO"].isnull().astype("int8")                                        # D
# E: 출하 직전 90일 더위(주의+경고+위험) 비율 — 날씨 누적합 + searchsorted
log("late-heat 계산…")
w=pd.read_csv(f"{ROOT}/data/processed/1_merge/step4-3_weather.csv",usecols=["stn","date","THI_grade"])
w["date"]=pd.to_datetime(w["date"]); w["hot"]=w["THI_grade"].isin(["주의","경고","위험"]).astype(int)
ab=pd.to_datetime(df["ABATT_DATE"],errors="coerce").values
lo_d=(pd.to_datetime(df["ABATT_DATE"],errors="coerce")-pd.Timedelta(days=90)).values
hot_late=np.full(len(df),np.nan); dwin=np.full(len(df),np.nan)
for stn,ws in w.groupby("stn"):
    ws=ws.sort_values("date"); dates=ws["date"].values
    csum=np.concatenate([[0],ws["hot"].cumsum().values])
    idx=np.where(df["stn"].values==stn)[0]
    if len(idx)==0: continue
    hi=np.searchsorted(dates,ab[idx],side="right"); lo=np.searchsorted(dates,lo_d[idx],side="left")
    hot_late[idx]=csum[hi]-csum[lo]; dwin[idx]=hi-lo
df["ratio_고온_late"]=(hot_late/np.where(dwin>0,dwin,np.nan)).astype("float32")                   # E
del w; gc.collect()
log("피처 빌드 완료. 결측/통계 확인:")
for c in ["weight_vs_peer","daily_gain_z_by_sex","ratio_고온_late","lineage_missing"]:
    print(f"  {c:22s} 결측 {df[c].isnull().mean()*100:4.1f}%  grade_num 상관 {df[[c,'grade_num']].corr().iloc[0,1]:+.3f}",flush=True)

# ── CV ──
fold=[(np.where(df["fold"]!=k)[0],np.where(df["fold"]==k)[0]) for k in range(NF)]
P=dict(objective="multiclass",num_class=16,learning_rate=0.05,n_estimators=500,max_depth=8,num_leaves=101,
   min_child_samples=66,subsample=0.888,subsample_freq=1,colsample_bytree=0.813,reg_alpha=0.001,reg_lambda=0.70,
   class_weight="balanced",n_jobs=-1,verbosity=-1,force_col_wise=True,random_state=SEED)
GOOD=["weight_vs_peer","daily_gain_z_by_sex","lineage_missing","ratio_고온_late"]
CFGS={"C0 기존(69)":[], "A +gain²(무의미예상)":["daily_gain_squared"], "B +또래체격":["weight_vs_peer"],
      "C +성별내z":["daily_gain_z_by_sex"], "D +혈통미등록":["lineage_missing"], "E +출하전더위":["ratio_고온_late"],
      "F +좋은것전부":GOOD}
res={k:{"f1":[],"acc":[]} for k in CFGS}
for k,(tri,vai) in enumerate(fold):
    tr=df.iloc[tri]; va=df.iloc[vai]
    trs,_=train_test_split(tr,train_size=min(SUB,len(tr)),stratify=tr["grade_num"],random_state=SEED)
    yt,yv=trs["grade_num"].astype(int).values,va["grade_num"].astype(int).values
    for nm,ex in CFGS.items():
        Xt,Xv=trs[FEATS+ex],va[FEATS+ex]
        m=lgb.LGBMClassifier(**P); m.fit(Xt,yt,eval_set=[(Xv,yv)],eval_metric="multi_logloss",
            callbacks=[lgb.early_stopping(40,verbose=False),lgb.log_evaluation(0)])
        p=np.asarray(m.predict(Xv)).ravel(); f=f1_score(yv,p,average="macro"); a=accuracy_score(yv,p)
        res[nm]["f1"].append(f); res[nm]["acc"].append(a); log(f"fold{k} [{nm}] F1={f:.4f} Acc={a:.4f}")
    del tr,va,trs; gc.collect()
print("\n===== 결과 (%d-fold 평균) ====="%NF,flush=True); base=np.mean(res["C0 기존(69)"]["f1"])
for nm in CFGS:
    f=np.mean(res[nm]["f1"]); print(f"  {nm:22s} MacroF1={f:.4f} ({f-base:+.4f})  Acc={np.mean(res[nm]['acc']):.4f}",flush=True)
log("완료")
