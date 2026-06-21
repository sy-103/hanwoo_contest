"""채택된 신규피처 weight_vs_peer(월령+성별 또래 대비 체격)를 step11에 추가.
CV에서 +0.0021로 유일하게 노이즈 넘은 피처. 타깃 미사용=누수0 (체중/월령/성별로만 계산)."""
import pandas as pd, time
ROOT="/Users/chosun-nhn04/IdeaProjects/Hanwoo_Contest"; t0=time.time()
P=f"{ROOT}/data/processed/3_eda/step11_features.csv"
df=pd.read_csv(P,encoding="utf-8-sig",low_memory=False)
print(f"[{time.time()-t0:.0f}s] 로드 {df.shape}",flush=True)
if "weight_vs_peer" in df.columns:
    df=df.drop(columns=["weight_vs_peer"])
# 월령+성별 또래 평균체중 대비 (feature-derived, 타깃 안 씀)
peer=df.groupby(["AGE","JUDGE_SEX"])["WEIGHT"].transform("mean")
df["weight_vs_peer"]=(df["WEIGHT"]/peer).astype("float32")
print(f"weight_vs_peer 결측 {df['weight_vs_peer'].isnull().mean()*100:.1f}% | grade_num 상관 {df[['weight_vs_peer','grade_num']].corr().iloc[0,1]:+.3f}",flush=True)
df.to_csv(P,index=False,encoding="utf-8-sig")
print(f"[{time.time()-t0:.0f}s] 저장 완료 {df.shape}",flush=True)
