"""21회차 잔차진단 — 실제 수치 + 그림 생성 (received 20_residuals.csv + 20_X_design.csv)."""
import pandas as pd, numpy as np
import statsmodels.api as sm
from statsmodels.stats.diagnostic import het_breuschpagan, het_white
from statsmodels.stats.stattools import jarque_bera, durbin_watson
from scipy import stats
import matplotlib, matplotlib.pyplot as plt, matplotlib.font_manager as fm
ROOT="/Users/chosun-nhn04/IdeaProjects/Hanwoo_Contest"; M="data/processed/4_model"

def set_korean_font():
    for nm in ["AppleGothic","Apple SD Gothic Neo","Nanum Gothic"]:
        if any(f.name==nm for f in fm.fontManager.ttflist):
            plt.rcParams["font.family"]=nm; break
    plt.rcParams["axes.unicode_minus"]=False

r=pd.read_csv(f"{ROOT}/{M}/20_residuals.csv", encoding="utf-8-sig")
X=pd.read_csv(f"{ROOT}/{M}/20_X_design.csv", encoding="utf-8-sig")
X.columns=[c.lstrip("﻿") for c in X.columns]
resid=r["resid"].to_numpy(); fitted=r["fitted"].to_numpy(); y=fitted+resid
exog=sm.add_constant(X, has_constant="add")
n=len(resid); print(f"n={n:,}, X변수={X.shape[1]} (+const), 잔차 평균={resid.mean():.3e}")

# ── 1. 잔차 분포·정규성 ──
sk=stats.skew(resid); ku=stats.kurtosis(resid, fisher=True)  # 초과첨도(정규=0)
JB,JBp,jbsk,jbku=jarque_bera(resid)
print(f"\n[정규성] 왜도={sk:.3f} 초과첨도={ku:.3f} | Jarque-Bera={JB:,.0f} p={JBp:.2e}")
print(f"  잔차 std={resid.std():.4f} | 5수치 {np.percentile(resid,[1,25,50,75,99]).round(3)}")

# ── 2. 등분산(이분산) — Breusch-Pagan 전체 ──
bp=het_breuschpagan(resid, exog)  # (LM, LMp, F, Fp)
print(f"\n[등분산] Breusch-Pagan LM={bp[0]:,.0f} p={bp[1]:.2e} | F p={bp[3]:.2e}")
print(f"  보조회귀 R²(=LM/n)={bp[0]/n:.4f}  → 잔차분산이 X로 설명되는 정도")
# White (샘플; 더미 교차항 특이행렬 위험 → try)
try:
    s=np.random.RandomState(42).choice(n,20000,replace=False)
    wt=het_white(resid[s], exog.iloc[s])
    print(f"  White(2만샘플) LM p={wt[1]:.2e}")
except Exception as e:
    print(f"  White 생략({type(e).__name__})")

# ── 3. 독립성 ──
dw=durbin_watson(resid)
print(f"\n[독립성] Durbin-Watson={dw:.3f} (2 근처=무자기상관; 횡단면이라 참고용)")

# ── 4. 영향점 — Cook's distance (샘플 재적합) ──
s=np.random.RandomState(0).choice(n,30000,replace=False)
mi=sm.OLS(y[s], exog.iloc[s].to_numpy()).fit()
infl=mi.get_influence(); cooks=infl.cooks_distance[0]; lev=infl.hat_matrix_diag
ns=len(s); thr=4/ns
print(f"\n[영향점] (3만 샘플) 쿡거리 max={cooks.max():.4f}, 4/n={thr:.2e} 초과 {int((cooks>thr).sum())}개({(cooks>thr).mean()*100:.1f}%)")
print(f"  레버리지 평균={lev.mean():.4f}(=k/n) max={lev.max():.4f}")

# ── 5. 그림 4종 ──
import seaborn as sns
plt.rcParams["figure.dpi"]=100; set_korean_font()
ps=np.random.RandomState(1).choice(n,25000,replace=False)
std_resid=resid/resid.std()
# (a) 잔차 vs 적합값
plt.figure(figsize=(6,4)); plt.scatter(fitted[ps],resid[ps],s=4,alpha=0.15)
plt.axhline(0,color="r",lw=1); plt.xlabel("적합값(예측 로그가격)"); plt.ylabel("잔차")
plt.title("잔차 vs 적합값 — 깔때기면 이분산"); plt.tight_layout()
plt.savefig(f"{ROOT}/figures/21_resid_vs_fitted.png",bbox_inches="tight"); plt.close()
# (b) Q-Q
sm.qqplot(resid[ps], line="s"); set_korean_font(); plt.title("Q-Q 플롯 — 직선에서 벗어나면 비정규")
plt.tight_layout(); plt.savefig(f"{ROOT}/figures/21_qq.png",bbox_inches="tight"); plt.close()
# (c) 히스토그램
plt.figure(figsize=(6,4)); plt.hist(resid,bins=120,density=True,alpha=0.6)
xx=np.linspace(resid.min(),resid.max(),200)
plt.plot(xx,stats.norm.pdf(xx,resid.mean(),resid.std()),"r",lw=1.5,label="정규분포")
plt.xlabel("잔차"); plt.ylabel("밀도"); plt.legend(); plt.title("잔차 분포 vs 정규분포"); plt.tight_layout()
plt.savefig(f"{ROOT}/figures/21_resid_hist.png",bbox_inches="tight"); plt.close()
# (d) Scale-Location
plt.figure(figsize=(6,4)); plt.scatter(fitted[ps],np.sqrt(np.abs(std_resid[ps])),s=4,alpha=0.15)
plt.xlabel("적합값"); plt.ylabel("√|표준화 잔차|"); plt.title("Scale-Location — 우상향이면 이분산"); plt.tight_layout()
plt.savefig(f"{ROOT}/figures/21_scale_location.png",bbox_inches="tight"); plt.close()
print("\n그림 4종 저장: figures/21_resid_vs_fitted/qq/resid_hist/scale_location.png")
