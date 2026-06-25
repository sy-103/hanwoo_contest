#!/bin/zsh
# lineage 0612 변경 영향 체인 재실행 (Step2_lineage는 이미 실행됨)
cd /Users/chosun-nhn04/IdeaProjects/Hanwoo_Contest
JL=./.venv/bin/jupyter
NBS=(
  notebooks/1_merge/Step3_area.ipynb
  notebooks/1_merge/Step4-4_weather.ipynb
  notebooks/1_merge/Verify.ipynb
  notebooks/2_preprocess/Step1_train.ipynb
  notebooks/2_preprocess/Step2_death.ipynb
  notebooks/2_preprocess/Step3_area.ipynb
  notebooks/2_preprocess/Step4_lineage.ipynb
  notebooks/2_preprocess/Step5_weather.ipynb
  notebooks/2_preprocess/Verify.ipynb
  notebooks/3_eda/13_상관관계.ipynb
  notebooks/3_eda/14_산점도_박스플롯.ipynb
  notebooks/3_eda/15_기상변수_탐색.ipynb
  notebooks/3_eda/16_시공간패턴.ipynb
  notebooks/3_eda/17_파생변수.ipynb
)
for nb in $NBS; do
  echo "[$(date +%H:%M:%S)] RUN: $nb"
  if $JL nbconvert --to notebook --execute --inplace "$nb" > /tmp/nbconv_last.log 2>&1; then
    echo "[$(date +%H:%M:%S)] OK:  $nb"
  else
    echo "[$(date +%H:%M:%S)] FAIL: $nb"; tail -15 /tmp/nbconv_last.log; exit 1
  fi
done
echo "[$(date +%H:%M:%S)] PIPELINE DONE"
ls -la data/processed/3_eda/step11_features.csv
