"""step11 재생성: 교정 KPN(0612) 반영 + KPN_freq 재계산 + EBV·조상등급 추가.
파이프라인 노트북 경로가 깨져 nbconvert 불가 → 동일 결과를 직접 생성. 원본은 백업됨."""
import pandas as pd, numpy as np, time, gc
ROOT="/Users/chosun-nhn04/IdeaProjects/Hanwoo_Contest"
t0=time.time(); log=lambda *a: print(f"[{time.time()-t0:5.0f}s]",*a,flush=True)

log("step11 로드…")
df=pd.read_csv(f"{ROOT}/data/processed/3_eda/step11_features.csv",encoding="utf-8-sig",low_memory=False)
orig_cols=df.columns.tolist()
log(f"원본 {df.shape}")
has_split="split" in df.columns

# split 병합(dev 마스크용) — 원래 없으면 끝에 제거
sp=pd.read_csv(f"{ROOT}/data/processed/4_model/split_assignment.csv",encoding="utf-8-sig")
df=df.merge(sp[["CATTLE_NO","split"]].rename(columns={"split":"_split"}),on="CATTLE_NO",how="left")
dev=df["_split"].eq("dev").values

# ① 교정 KPN 교체
lin=pd.read_csv(f"{ROOT}/data/raw/hanwoo_lineage_0612.csv",dtype=str,usecols=["CATTLE_NO","KPN_NO"]).drop_duplicates("CATTLE_NO")
lin["KPN_NO"]=lin["KPN_NO"].str.strip()
df=df.drop(columns=["KPN_NO"]).merge(lin,on="CATTLE_NO",how="left"); del lin; gc.collect()
# ② KPN_NO_freq 재계산 (파이프라인 17_파생변수와 동일: 전체 value_counts)
df["KPN_NO_freq"]=df["KPN_NO"].map(df["KPN_NO"].value_counts())
log(f"KPN 교체+freq 재계산 (KPN 결측 {df['KPN_NO'].isnull().mean()*100:.1f}%)")

# ③ EBV 조인 (static 외부값)
g=pd.read_excel(f"{ROOT}/data/raw/KPN 유전능력 자료.xlsx"); g["KPN명호"]=g["KPN명호"].astype(str).str.strip()
g=g[~g["KPN명호"].isin(["nan",""])].drop_duplicates("KPN명호")
emap={"근내지방도 육종가":"ebv_marbling","근내지방도 정확도":"ebv_marbling_acc","도체중 육종가":"ebv_cwt",
      "등심단면적 육종가":"ebv_rea","등지방두께 육종가":"ebv_backfat","12개월체중 육종가":"ebv_wt12"}
g2=g[["KPN명호"]+list(emap)].rename(columns={"KPN명호":"KPN_NO",**emap})
for c in emap.values(): g2[c]=pd.to_numeric(g2[c],errors="coerce")
df=df.merge(g2,on="KPN_NO",how="left"); df["has_ebv"]=df["ebv_marbling"].notna().astype(int); del g,g2; gc.collect()
log(f"EBV 조인 (개체 커버 {df['has_ebv'].mean()*100:.1f}%)")

# ④ 조상 본인등급 (dev-only lookup + 날짜 게이트)
jd=pd.to_datetime(df["JUDGE_DATE"],errors="coerce")
gn=pd.to_numeric(df["grade_num"],errors="coerce")
rec=pd.DataFrame({"jd":jd[dev].values,"grade_num":gn[dev].values},index=df.loc[dev,"CATTLE_NO"].values)
rec=rec[~rec.index.duplicated()]
for idcol,nm in [("MOTHER_ANIMAL_NO","mother"),("FATHER_CATTLE_NO","father")]:
    a_jd=df[idcol].map(rec["jd"]); a_g=df[idcol].map(rec["grade_num"])
    gated=np.where(a_jd.values<jd.values, a_g.values, np.nan)   # 조상이 송아지보다 먼저 판정된 것만
    df[f"{nm}_own_grade"]=gated; df[f"has_{nm}_grade"]=(~np.isnan(gated)).astype(int)
    log(f"{nm}_own_grade 커버 {df[f'has_{nm}_grade'].mean()*100:.1f}%")

# 저장 (원래 컬럼 + 신규; split 임시컬럼 제거)
new_cols=["ebv_marbling","ebv_marbling_acc","ebv_cwt","ebv_rea","ebv_backfat","ebv_wt12","has_ebv",
          "mother_own_grade","has_mother_grade","father_own_grade","has_father_grade"]
out_cols=orig_cols+new_cols
df=df[out_cols]
log(f"저장 중… 최종 {df.shape} (원본 {len(orig_cols)} + 신규 {len(new_cols)})")
df.to_csv(f"{ROOT}/data/processed/3_eda/step11_features.csv",index=False,encoding="utf-8-sig")
log("완료")
print("신규 컬럼:",new_cols,flush=True)
