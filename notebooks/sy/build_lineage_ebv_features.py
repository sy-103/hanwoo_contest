"""
혈통(KPN 교정본) + 유전능력(EBV) → 누수안전 파생변수 사이드파일 생성.
출력: data/processed/4_model/lineage_ebv_features.csv  (CATTLE_NO 키로 step11에 머지)

누수 차단 원칙
- freq / EBV 결측대치 / TE 전역평균 : 전부 dev(train)에서만 계산
- TE(target-encoding) : dev는 out-of-fold(자기 폴드 제외), 비-dev(holdout/test)는 전체 dev 기준
- 미관측 KPN → 전역평균 폴백, EB 스무딩으로 저빈도 종모우 수축
- split_assignment.csv 의 기존 폴드 그대로 사용 (재생성 안 함)
"""
import pandas as pd, numpy as np

ROOT   = "/Users/chosun-nhn04/IdeaProjects/Hanwoo_Contest"
STEP11 = f"{ROOT}/data/processed/3_eda/step11_features.csv"
LINEAGE= f"{ROOT}/data/raw/hanwoo_lineage_0612.csv"
EBV    = f"{ROOT}/data/raw/KPN 유전능력 자료.xlsx"
SPLIT  = f"{ROOT}/data/processed/4_model/split_assignment.csv"
OUT    = f"{ROOT}/data/processed/4_model/lineage_ebv_features.csv"
M = 50  # EB 스무딩 prior (저빈도 종모우를 전역평균으로 수축)

print("[1/5] 로드 (가벼운 컬럼만)…")
base = pd.read_csv(STEP11, usecols=["CATTLE_NO", "grade_num"],
                   encoding="utf-8-sig", low_memory=False)
base["grade_num"] = pd.to_numeric(base["grade_num"], errors="coerce")
lin = pd.read_csv(LINEAGE, dtype=str,
                  usecols=["CATTLE_NO", "KPN_NO", "FATHER_CATTLE_NO", "F_GFATHER_CATTLE_NO"])
for c in lin.columns:
    lin[c] = lin[c].str.strip()
lin = lin.drop_duplicates("CATTLE_NO")
split = pd.read_csv(SPLIT, encoding="utf-8-sig")

df = (base.merge(lin, on="CATTLE_NO", how="left")
          .merge(split, on="CATTLE_NO", how="left"))
print(f"   모델링 개체: {len(df):,}  | split 분포: {df['split'].value_counts().to_dict()}")
dev_mask = (df["split"] == "dev").values
gm_grade = df.loc[dev_mask, "grade_num"].mean()
print(f"   dev 전역 평균등급(grade_num) = {gm_grade:.4f}")

print("[2/5] dev 기준 빈도(freq) 재계산…")
freq = df.loc[dev_mask, "KPN_NO"].value_counts()
df["KPN_NO_freq"] = df["KPN_NO"].map(freq).fillna(0).astype(int)

print("[3/5] EBV 조인 (dedup) + dev평균 대치 + has_ebv…")
g = pd.read_excel(EBV)
g["KPN명호"] = g["KPN명호"].astype(str).str.strip()
g = g[~g["KPN명호"].isin(["nan", ""])].drop_duplicates("KPN명호")
ebv_map = {"근내지방도 육종가": "ebv_marbling", "근내지방도 정확도": "ebv_marbling_acc",
           "도체중 육종가": "ebv_cwt", "등심단면적 육종가": "ebv_rea",
           "등지방두께 육종가": "ebv_backfat", "12개월체중 육종가": "ebv_wt12"}
g2 = g[["KPN명호"] + list(ebv_map)].rename(columns={"KPN명호": "KPN_NO", **ebv_map})
for c in ebv_map.values():
    g2[c] = pd.to_numeric(g2[c], errors="coerce")
df = df.merge(g2, on="KPN_NO", how="left")
df["has_ebv"] = df["ebv_marbling"].notna().astype(int)
for c in ebv_map.values():               # dev 평균으로만 대치 (누수방지)
    devmean = df.loc[dev_mask & df[c].notna(), c].mean()
    df[c] = df[c].fillna(devmean)
print(f"   EBV 커버리지(개체): {df['has_ebv'].mean()*100:.1f}%  | dev: {df.loc[dev_mask,'has_ebv'].mean()*100:.1f}%")


def eb(means, counts, prior, m):
    return (counts * means + m * prior) / (counts + m)


def oof_te(key, premium=False, m=M):
    """dev=out-of-fold, 비-dev=전체dev. 미관측→전역평균."""
    tgt = (df["grade_num"] >= 10).astype(float) if premium else df["grade_num"].astype(float)
    prior = tgt[dev_mask].mean()
    out = pd.Series(np.nan, index=df.index)
    tmp = pd.DataFrame({"k": df[key].values, "y": tgt.values, "fold": df["fold"].values}, index=df.index)
    dev_idx = df.index[dev_mask]
    for fk in sorted(df.loc[dev_mask, "fold"].dropna().unique()):
        tr = dev_idx[df.loc[dev_idx, "fold"] != fk]
        va = dev_idx[df.loc[dev_idx, "fold"] == fk]
        st = tmp.loc[tr].groupby("k")["y"].agg(["mean", "count"])
        enc = eb(st["mean"], st["count"], prior, m)
        out.loc[va] = tmp.loc[va, "k"].map(enc)
    nondev = df.index[~dev_mask]
    st_all = tmp.loc[dev_idx].groupby("k")["y"].agg(["mean", "count"])
    enc_all = eb(st_all["mean"], st_all["count"], prior, m)
    out.loc[nondev] = tmp.loc[nondev, "k"].map(enc_all)
    return out.fillna(prior)


print("[4/5] 누수안전 target-encoding (KPN/부/조부)…")
df["kpn_te_grade"]     = oof_te("KPN_NO")
df["kpn_te_highgrade"] = oof_te("KPN_NO", premium=True)
df["father_te_grade"]  = oof_te("FATHER_CATTLE_NO")
df["fgf_te_grade"]     = oof_te("F_GFATHER_CATTLE_NO")

# ── 검증 리포트 ──────────────────────────────────────────────
print("\n===== 검증 =====")
new_feats = ["KPN_NO_freq", "ebv_marbling", "ebv_marbling_acc", "ebv_cwt", "ebv_rea",
             "ebv_backfat", "ebv_wt12", "has_ebv",
             "kpn_te_grade", "kpn_te_highgrade", "father_te_grade", "fgf_te_grade"]
print("NaN 잔존:", {c: int(df[c].isnull().sum()) for c in new_feats if df[c].isnull().sum()} or "없음")
dv = df[dev_mask]
print("\n[dev에서 grade_num과의 상관 — TE는 OOF라 누수통제됨]")
for c in ["ebv_marbling", "ebv_cwt", "kpn_te_grade", "kpn_te_highgrade", "father_te_grade", "fgf_te_grade", "KPN_NO_freq"]:
    r = dv[[c, "grade_num"]].corr().iloc[0, 1]
    print(f"  {c:20s} r = {r:+.4f}")
print("  (기존 최강 daily_gain r≈0.52 대비 비교)")
# 폴드별 OOF 무결성: 각 폴드 TE가 그 폴드 타깃을 안 썼는지 간접확인(폴드별 평균이 전역과 다름)
print("\n[OOF 무결성 체크] kpn_te_grade 폴드별 평균(전역과 미세하게 다르면 정상):")
print(dv.groupby("fold")["kpn_te_grade"].mean().round(4).to_dict())

print("\n[5/5] 저장…")
df[["CATTLE_NO"] + new_feats].to_csv(OUT, index=False, encoding="utf-8-sig")
print(f"   저장 완료: {OUT}  ({len(df):,} 행 × {len(new_feats)+1} 열)")
