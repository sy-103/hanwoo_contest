// 한우 EDA 13~17회차 팀 공유 보고서 생성 스크립트
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  Header, Footer, AlignmentType, LevelFormat, TableOfContents, HeadingLevel,
  BorderStyle, WidthType, ShadingType, PageNumber, PageBreak,
} = require("docx");

const FIGDIR = path.join(__dirname, "figures");
const CW = 9026; // A4 content width (DXA)

// ---------- 공통 헬퍼 ----------
function runs(parts) {
  if (!Array.isArray(parts)) parts = [parts];
  return parts.map((p) =>
    typeof p === "string"
      ? new TextRun({ text: p })
      : new TextRun({
          text: p.t, bold: p.b, italics: p.i, color: p.color,
          size: p.sz, font: p.font,
        })
  );
}
const P = (parts, opts = {}) =>
  new Paragraph({ children: runs(parts), spacing: { after: 120 }, ...opts });
const H1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(t)] });
const H2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(t)] });
const H3 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun(t)] });
const BULLET = (parts) =>
  new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: runs(parts), spacing: { after: 60 } });
const NUM = (ref) => (parts) =>
  new Paragraph({ numbering: { reference: ref, level: 0 }, children: runs(parts), spacing: { after: 60 } });
const PAGEBREAK = () => new Paragraph({ children: [new PageBreak()] });

// 수식 단락 (들여쓰기 + 연한 배경)
const FORMULA = (t) =>
  new Paragraph({
    children: [new TextRun({ text: t, font: "Cambria Math", size: 22 })],
    shading: { fill: "EEF3FA", type: ShadingType.CLEAR },
    indent: { left: 360 },
    spacing: { before: 60, after: 120 },
  });

// 코드 블록
function codeBlock(code) {
  return code.split("\n").map(
    (l) =>
      new Paragraph({
        children: [new TextRun({ text: l.length ? l : " ", font: "Courier New", size: 16 })],
        shading: { fill: "F5F5F5", type: ShadingType.CLEAR },
        indent: { left: 240 },
        spacing: { after: 0, line: 240 },
      })
  );
}

// 강조 박스 (1셀 표)
function box(title, lines, fill = "E8F0E8") {
  const border = { style: BorderStyle.SINGLE, size: 4, color: "AAAAAA" };
  const children = [
    new Paragraph({ children: [new TextRun({ text: title, bold: true })], spacing: { after: 80 } }),
    ...lines.map((l) => new Paragraph({ children: runs(l), spacing: { after: 40 } })),
  ];
  return new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [CW],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: { top: border, bottom: border, left: border, right: border },
            width: { size: CW, type: WidthType.DXA },
            shading: { fill, type: ShadingType.CLEAR },
            margins: { top: 120, bottom: 120, left: 160, right: 160 },
            children,
          }),
        ],
      }),
    ],
  });
}

// 일반 표
function tbl(headers, rows, widths) {
  const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
  const borders = { top: border, bottom: border, left: border, right: border };
  if (!widths) widths = headers.map(() => Math.floor(CW / headers.length));
  const mkCell = (text, opts = {}) =>
    new TableCell({
      borders,
      width: { size: opts.w, type: WidthType.DXA },
      shading: { fill: opts.fill || "FFFFFF", type: ShadingType.CLEAR },
      margins: { top: 60, bottom: 60, left: 100, right: 100 },
      children: [
        new Paragraph({
          children: [
            new TextRun({
              text: String(text), bold: opts.bold, size: 18,
              color: opts.color || "000000",
            }),
          ],
          alignment: opts.center ? AlignmentType.CENTER : AlignmentType.LEFT,
        }),
      ],
    });
  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map((h, i) => mkCell(h, { w: widths[i], fill: "17407A", bold: true, color: "FFFFFF", center: true })),
  });
  const bodyRows = rows.map(
    (r, ri) =>
      new TableRow({
        children: r.map((c, i) => mkCell(c, { w: widths[i], fill: ri % 2 ? "F2F6FA" : "FFFFFF" })),
      })
  );
  return new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: widths, rows: [headerRow, ...bodyRows] });
}

// PNG 크기 읽어 비율 유지 삽입
function pngSize(file) {
  const b = fs.readFileSync(file);
  return { w: b.readUInt32BE(16), h: b.readUInt32BE(20) };
}
function figure(file, caption, maxW = 580, maxH = 460) {
  const fp = path.join(FIGDIR, file);
  const { w, h } = pngSize(fp);
  const scale = Math.min(maxW / w, maxH / h, 1);
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 80, after: 40 },
      children: [
        new ImageRun({
          type: "png",
          data: fs.readFileSync(fp),
          transformation: { width: Math.round(w * scale), height: Math.round(h * scale) },
          altText: { title: caption, description: caption, name: file },
        }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 160 },
      children: [new TextRun({ text: `[그림] ${caption}  (${file})`, size: 16, color: "666666" })],
    }),
  ];
}

// ---------- 문서 내용 ----------
const content = [];

// ===== 표지 =====
content.push(
  new Paragraph({ spacing: { before: 2400 } }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "2026 기상청 날씨 빅데이터 콘테스트", size: 28, color: "17407A", bold: true })],
    spacing: { after: 200 },
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "한우 EDA 13~17회차 종합 보고서", size: 52, bold: true })],
    spacing: { after: 120 },
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "이론 · 코드 · 변경 이력 · 결과 · 해석 · 다음 단계", size: 26, color: "444444" })],
    spacing: { after: 600 },
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "분석 데이터: 2,408,699마리 × 92변수 (2023~2025 도축 한우)", size: 22, color: "555555" })],
    spacing: { after: 80 },
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "상관관계 / 산점도·박스플롯 / 기상 탐색 / 시공간 패턴 / 파생변수", size: 22, color: "555555" })],
    spacing: { after: 80 },
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "작성일: 2026년 6월 11일", size: 22, color: "555555" })],
  }),
  PAGEBREAK()
);

// ===== 목차 =====
content.push(
  new Paragraph({ children: [new TextRun({ text: "목   차", size: 36, bold: true })], spacing: { after: 240 } }),
  new TableOfContents("목차", { hyperlink: true, headingStyleRange: "1-2" }),
  PAGEBREAK()
);

// ===== 1장 =====
content.push(
  H1("1. 이 보고서를 읽는 법"),
  P("이 보고서는 EDA(탐색적 데이터 분석) 13~17회차의 전 과정을 팀원 누구나 — 데이터 분석을 처음 접하는 학부생 3학년이라도 — 따라올 수 있도록 정리한 문서입니다. 각 회차 장(3~7장)은 모두 같은 순서로 구성됩니다."),
  tbl(
    ["순서", "내용", "답하는 질문"],
    [
      ["N.1 이론", "이 분석이 무엇이고 왜 하는지 + 공식", "“이게 뭐 하는 분석이지?”"],
      ["N.2 코드 설명", "노트북 코드가 단계별로 하는 일", "“코드가 뭘 하고 있지?”"],
      ["N.3 변경점", "기존 강의안 코드와 무엇이 다른지", "“어디를 왜 바꿨지?”"],
      ["N.4 결과", "실행해서 나온 실제 수치와 그림", "“그래서 뭐가 나왔지?”"],
      ["N.5 해석", "결과를 읽는 방법과 실제 해석", "“이 숫자가 무슨 뜻이지?”"],
      ["N.6 다음 단계 반영", "분류·회귀 분석에 어떻게 쓰이는지", "“그 다음은?”"],
    ],
    [1500, 4026, 3500]
  ),
  P([{ t: "" }]),
  P([
    "8장은 코드 변경 이력 전체, 9장은 AI 교차 검증 Q&A, 10장은 다음 분석 연결, 11장은 총정리입니다. 처음 읽는 분은 ",
    { t: "2장(큰 그림) → 11장(총정리)", b: true },
    "을 먼저 읽고 각 회차 장으로 들어가는 것을 추천합니다.",
  ]),
  PAGEBREAK()
);

// ===== 2장 =====
content.push(
  H1("2. 분석의 큰 그림"),
  H2("2.1 데이터와 분석 파이프라인"),
  P([
    "분석 대상은 ", { t: "2023~2025년 도축된 한우 2,408,699마리", b: true },
    "의 개체별 기록입니다. 전처리(1~12회차)가 끝난 step6_preprocess.csv(44개 변수)에서 출발해, 13~17회차를 거치며 변수가 추가되어 step11_features.csv(92개 변수)로 끝납니다.",
  ]),
  tbl(
    ["회차", "노트북", "하는 일", "출력 파일"],
    [
      ["13", "13_상관관계", "변수 간 상관 점검, 다중공선성 식별", "step7_corr.csv (45열)"],
      ["14", "14_산점도_박스플롯", "곡선 관계·그룹 차이를 눈으로 확인 + 검정", "step8_scatter.csv (46열)"],
      ["15", "15_기상변수_탐색", "더위 노출 비율 생성, 등급·가격과의 관계", "step9_weather.csv (51열)"],
      ["16", "16_시공간패턴", "연·월·지역 패턴, 인코딩 전략 결정", "step10_spatial.csv (56열)"],
      ["17", "17_파생변수", "파생변수 설계(효율·임계·상호작용·인코딩)", "step11_features.csv (92열)"],
    ],
    [800, 2400, 3826, 2000]
  ),
  P([{ t: "" }]),
  H2("2.2 두 개의 트랙 (Dual Track)"),
  P("이 프로젝트의 목표는 두 개입니다. 어떤 변수를 어디에 쓸 수 있는지가 트랙마다 다르므로, 이 구분이 모든 분석의 출발점입니다."),
  tbl(
    ["", "트랙1 (제출용 분류)", "트랙2 (기상 영향 분석)"],
    [
      ["목적", "소 한 마리의 등급(LAST_GRADE) 예측", "날씨가 한우 품질·수익성에 주는 영향 분석"],
      ["평가", "대회 제출 → 채점", "채점 없음 (보고서·인사이트)"],
      ["사용 가능 변수", "test에 있는 13개 컬럼 + 거기서 파생된 것만", "모든 변수 (육질·가격 포함)"],
      ["모델", "LightGBM 분류 (19회차)", "OLS/Lasso 회귀 (20회차)"],
    ],
    [1800, 3613, 3613]
  ),
  P([{ t: "" }]),
  H2("2.3 가장 중요한 개념 — 정답 누수 (Data Leakage)"),
  P([
    { t: "비유: ", b: true },
    "시험 문제지 뒤에 정답이 실수로 인쇄돼 있다고 합시다. 그걸 보고 100점을 맞으면 실력이 아니라 컨닝입니다. 실제 시험(test)에는 정답지가 없으니 점수는 폭락합니다. 데이터 분석에서 “실제 예측 시점에는 알 수 없는 정보를 학습에 몰래 쓰는 것”을 ",
    { t: "누수(leakage)", b: true }, "라고 합니다.",
  ]),
  P([
    "한우 등급은 도축 후 등심 단면을 보고 매깁니다. 그런데 근내지방도(INSFAT)·등심단면적(REA)·조직감(TISSUE) 같은 ",
    { t: "육질 변수도 똑같이 도축 후에 측정", b: true },
    "됩니다. 즉 육질 변수는 “정답의 일부”이며, 실제로 13회차에서 등급과의 상관이 r=0.93까지 나옵니다. 이걸 제출 모델에 넣으면 정답을 보고 정답을 맞히는 셈입니다.",
  ]),
  box("누수 판단 기준 (이번 검증의 핵심)", [
    [{ t: "기준은 “도축 후 측정이냐”가 아니라 “test 데이터에 있느냐”입니다.", b: true }],
    ["test 13개 컬럼: sido, sigungu, eupmyeondong, stn, ABATT_DATE, JUDGE_DATE, JUDGE_SEX, WEIGHT, AGE, BIRTH_YMD, CATTLE_NO, FARM_UNIQUE_NO, (LAST_GRADE)"],
    [{ t: "도체중 WEIGHT와 월령 AGE는 test에 있으므로 트랙1에서 사용 가능", b: true }, " — 도축 후 측정값이라도 주최 측이 test에 제공했다면 예측 시점에 쓸 수 있다는 뜻이므로 누수가 아닙니다."],
    ["반대로 육질 8개(BACKFAT, REA, WINDEX, INSFAT, YUKSAK, FATSAK, TISSUE, GROWTH)와 가격 COST_AMT는 test에 없으므로 트랙2 전용입니다."],
  ]),
  PAGEBREAK()
);

// ===== 3장: 13회차 =====
content.push(
  H1("3. 13회차 — 상관관계 분석"),
  H2("3.1 이론: 이 분석이 무엇이고 왜 하나"),
  P("상관계수는 두 변수가 “같이 움직이는 정도”를 -1 ~ +1 사이 숫자 하나로 요약합니다. +1에 가까우면 한쪽이 클 때 다른 쪽도 크고, -1에 가까우면 반대로 움직이며, 0이면 무관합니다. 본격적인 모델링 전에 변수들의 지도를 그리는 작업입니다."),
  H3("피어슨(Pearson) 상관계수 — 직선 관계"),
  FORMULA("r = Σ(xᵢ − x̄)(yᵢ − ȳ) / √[ Σ(xᵢ − x̄)² · Σ(yᵢ − ȳ)² ]"),
  P("분자는 “x가 평균보다 클 때 y도 평균보다 큰가”를 누적한 공분산이고, 분모는 각 변수의 흩어짐(표준편차)으로 나눠 -1~+1로 표준화한 것입니다. 직선 관계만 잡습니다."),
  H3("스피어만(Spearman) 상관계수 — 순위 관계"),
  FORMULA("ρ = 1 − 6Σdᵢ² / ( n(n² − 1) ),   dᵢ = xᵢ의 순위 − yᵢ의 순위"),
  P("값 자체가 아니라 순위로 바꿔 계산합니다. 곡선이라도 “꾸준히 증가/감소”(단조 관계)면 잡아내고, 극단값(이상치)에 강합니다. 피어슨과 스피어만이 많이 다르면 “직선이 아닌 관계”를 의심하고 14회차 산점도로 확인합니다."),
  H3("다중공선성 (Multicollinearity)"),
  P([
    "서로 거의 같은 변수(|r| ≥ 0.7)가 회귀모델에 함께 들어가면 계수가 불안정해지고 해석이 불가능해집니다. 회의에서 똑같은 말을 하는 사람이 3명 있으면 누구 말을 들어야 할지 헷갈리는 것과 같습니다. 정량 지표로는 VIF(분산팽창계수)를 쓰며, 회귀(20회차)에서 점검합니다.",
  ]),
  FORMULA("VIFⱼ = 1 / (1 − R²ⱼ)   — 변수 j를 나머지 변수로 회귀했을 때의 설명력 R²ⱼ. 10 이상이면 위험."),
  H2("3.2 코드 설명"),
  ...[
    "① step6 로드 후 등급 문자(1++A~등외)를 0~15 숫자(grade_num)로 변환 — 상관 계산은 숫자만 가능하므로.",
    "② 변수를 성격별 그룹(형질/농장/기상/타깃)으로 정의하고 25개 수치 변수의 피어슨 상관행렬 계산.",
    "③ 핵심 변수 7개 히트맵 + 그룹별 히트맵 4장 저장.",
    "④ |r| ≥ 0.7인 쌍을 자동 추출(상관행렬의 위 삼각형만 검사해 중복 제거).",
    "⑤ days_* 변수가 사육기간(AGE)과 얼마나 묶이는지 “일수 함정” 점검.",
    "⑥ 피어슨·스피어만 차이가 0.1 넘는 쌍 추출 → 곡선 관계 후보.",
    "⑦ grade_num·COST_AMT와의 상관을 절댓값 순으로 정렬 — 예측에 쓸 후보 선별.",
    "⑧ step7_corr.csv 저장.",
  ].map((t) => NUM("steps13")(t)),
  H2("3.3 강의안 대비 변경점"),
  tbl(
    ["어디를", "어떻게", "왜"],
    [
      ["그림 저장 경로 (4곳)", "“../figures” → “../../figures”", "노트북이 notebooks/3_eda/ 하위로 한 단계 깊어져, 강의안 경로 그대로면 그림이 notebooks/figures에 흩어짐. 프로젝트 루트 figures/로 통일"],
      ["공통 헤더 폰트 설정", "sns.set_style()을 set_korean_font()보다 먼저 호출", "seaborn의 set_style이 font.family를 sans-serif로 덮어써 그림의 모든 한글이 □로 깨지는 버그가 있었음 (8장 참고)"],
    ],
    [2200, 2800, 4026]
  ),
  P([{ t: "" }]),
  H2("3.4 결과"),
  P([{ t: "강한 상관 쌍: |r| ≥ 0.7이 총 28쌍", b: true }, " 발견. 대표 쌍은 다음과 같습니다."]),
  tbl(
    ["변수 쌍", "r", "이유"],
    [
      ["days_양호 ↔ days_total", "+0.981", "정의상 묶임 (총일수의 대부분이 양호일)"],
      ["C2025 ↔ C2024 ↔ C2023", "+0.958~0.974", "연도별 사육두수는 해마다 비슷"],
      ["grade_num ↔ INSFAT", "+0.930", "근내지방도는 등급 판정 기준 그 자체 (정답지)"],
      ["grade_num ↔ TISSUE", "−0.874", "조직감 점수도 판정 기준 (값이 작을수록 좋음)"],
      ["COST_AMT ↔ grade_num", "+0.867", "등급이 좋으면 비싸게 팔림"],
      ["AGE ↔ GROWTH", "+0.858", "월령과 성장 관련 측정값"],
      ["days_total ↔ AGE", "+0.800", "오래 산 소가 사육일수도 많음 (일수 함정)"],
    ],
    [2900, 1200, 4926]
  ),
  P([{ t: "" }]),
  P([{ t: "grade_num(등급)과의 상관 순위", b: true }, " — INSFAT +0.930, TISSUE −0.874, COST_AMT +0.867, REA +0.540, GROWTH −0.481, WEIGHT +0.475, AGE −0.465, days_양호 −0.389, … days_위험 −0.083, 기상 평균(최저기온 등) +0.07 수준."]),
  P([{ t: "COST_AMT(가격)와의 상관 순위", b: true }, " — grade_num +0.867, INSFAT +0.827, REA +0.619, WEIGHT +0.550, AGE −0.505, ta_min_mean(평균 최저기온) +0.215."]),
  P([{ t: "피어슨 vs 스피어만 차이 큰 쌍", b: true }, " — grade_num↔AGE (P=−0.46, S=−0.34, 차이 0.12) 한 쌍 → 월령-등급 관계는 직선이 아닐 가능성 → 14회차 LOWESS로 확인."]),
  ...figure("13_corr_key.png", "핵심 변수 7개 피어슨 상관 히트맵"),
  ...figure("13_grade_corr.png", "등급(grade_num)과의 상관 상위 20개 변수"),
  ...figure("13_corr_한우형질.png", "한우 형질 그룹 상관행렬"),
  ...figure("13_corr_기상.png", "기상 변수 그룹 상관행렬"),
  ...figure("13_corr_농장.png", "농장 변수 그룹 상관행렬"),
  H2("3.5 해석 (읽는 방법과 함께)"),
  box("히트맵·상관 읽는 방법", [
    ["빨강 진할수록 양의 상관, 파랑 진할수록 음의 상관, 흰색은 무관. 대각선 1.00은 자기 자신이라 무시."],
    ["|r| 0.2 미만 거의 무관 / 0.2~0.4 약함 / 0.4~0.7 중간 / 0.7~0.9 강함(다중공선성 위험) / 0.9 이상 사실상 같은 변수."],
    ["240만 행에서는 r=0.02도 p값이 유의하게 나오므로, p값이 아니라 상관의 크기(절댓값)로 판단."],
  ]),
  BULLET([{ t: "육질 변수(INSFAT 등)와 등급의 r>0.87은 “발견”이 아니라 “정의”", b: true }, " — 등급을 매길 때 쓰는 측정값이니 당연히 높습니다. 트랙1 제출 모델에서는 제외하고, 트랙2 분석에서만 씁니다."]),
  BULLET([{ t: "WEIGHT +0.475는 진짜 쓸모 있는 신호", b: true }, " — 무거운 소가 등급이 좋은 경향. test에 있는 변수라 트랙1의 핵심 피처가 됩니다."]),
  BULLET([{ t: "days_* 일수 변수의 음의 상관(−0.3대)은 “더위 효과”가 아닐 수 있음", b: true }, " — days_total↔AGE r=0.80, 즉 일수가 큰 소는 그냥 오래 산 소(노폐우 포함)입니다. 더위 좋은날 일수까지 등급과 음(−0.389)인 것이 증거. 그래서 15회차에서 비율로 바꿉니다."]),
  BULLET([{ t: "C2023~C2025는 셋 중 하나만 써도 충분", b: true }, " (r>0.95). 다중공선성 처리는 20회차 Lasso가 자동으로 하므로 지금은 제거하지 않고 기록만 합니다."]),
  H2("3.6 다음 분석 반영"),
  BULLET("트랙1 피처 후보 확정: WEIGHT, AGE (+나중에 파생변수). 육질·가격은 트랙2 전용으로 분리."),
  BULLET("등급↔AGE의 곡선 의심 → 17회차에서 Age_squared(제곱항) 생성의 근거."),
  BULLET("days 일수 함정 → 15회차 비율 변수(ratio_*) 설계의 근거."),
  BULLET("|r|≥0.7 쌍 목록 → 20회차 회귀에서 VIF 점검·Lasso 변수 선택의 입력."),
  PAGEBREAK()
);

// ===== 4장: 14회차 =====
content.push(
  H1("4. 14회차 — 산점도와 박스플롯"),
  H2("4.1 이론: 이 분석이 무엇이고 왜 하나"),
  P("상관계수는 숫자 하나라 곡선 관계, 숨은 그룹, 이상치를 놓칩니다. 그래서 그림으로 직접 봅니다. 산점도는 점 하나가 소 한 마리로 두 연속 변수의 관계 모양을, 박스플롯은 그룹별 분포(중앙값, 사분위, 이상치)를 보여줍니다."),
  H3("LOWESS (국소 가중 회귀)"),
  P("점 구름 사이로 부드러운 곡선을 그어주는 기법입니다. 각 지점 근처의 점들만으로 작은 회귀를 반복해 이어 붙이므로, 데이터가 실제로 직선인지 U자인지 눈으로 확인할 수 있습니다. 피어슨이 못 잡는 비선형을 잡는 도구입니다."),
  H3("Kruskal-Wallis 검정 — 그룹이 3개 이상일 때"),
  P("“그룹 간 차이가 우연인가”를 확인하는 검정입니다. 평균 비교인 ANOVA는 정규분포를 가정하지만, 우리 데이터(240만 건, 치우친 분포)는 그 가정이 안 맞으므로 순위 기반 비모수 검정을 씁니다. 모든 값을 통째로 순위 매긴 뒤, 그룹별 순위합이 고르게 섞였는지를 봅니다."),
  FORMULA("H = [ 12 / (N(N+1)) ] · Σⱼ (Rⱼ² / nⱼ) − 3(N+1)   (N=전체 표본, nⱼ·Rⱼ=그룹 j의 크기·순위합)"),
  P([
    "그룹들이 완전히 갈릴수록 H가 커지며, ", { t: "이론적 최대값은 N−1", b: true },
    "입니다(우리 데이터로는 약 240만). 성별처럼 그룹이 3개면 자유도 2의 카이제곱 분포와 비교해 p값을 얻습니다. 표본이 클수록 H 절대값도 커지는 것이 정상입니다.",
  ]),
  H2("4.2 코드 설명"),
  ...[
    "① step7 로드. 산점도용으로만 2만 행 무작위 샘플(240만 점을 다 찍으면 새까매서 안 보임). 박스플롯·검정은 전체 데이터로.",
    "② 가격(COST_AMT) vs 핵심 변수(WEIGHT/INSFAT/AGE/days_위험) 산점도 + 추세선. 상관계수는 전체 데이터로 계산해 제목에 표기.",
    "③ AGE×WEIGHT, AGE×가격에 LOWESS 곡선 — 13회차에서 의심된 곡선 관계 확인.",
    "④ 등급별 형질 박스플롯(WEIGHT/AGE/BACKFAT/REA/INSFAT), 도축 계절별·시도별 가격, 성별 3종(거세/암/수) 박스플롯.",
    "⑤ Kruskal-Wallis 검정 3종: 성별×도체중, 계절×가격, 등급×도체중.",
    "⑥ abatt_season(도축 계절) 컬럼 추가 후 step8_scatter.csv 저장.",
  ].map((t) => NUM("steps14")(t)),
  H2("4.3 강의안 대비 변경점"),
  P("그림 경로(7곳)와 폰트 호출 순서 — 3.3절과 동일한 공통 수정뿐, 분석 로직 변경 없음."),
  H2("4.4 결과"),
  tbl(
    ["검정", "그룹별 값", "H 통계량", "p값"],
    [
      ["성별 × 도체중", "거세 중앙값 473kg (N=1,222,465) / 수 443kg (N=11,820) / 암 372kg (N=1,174,414)", "1,088,062.9", "< 0.0001"],
      ["계절 × 가격", "가을·연말 높음, 봄·초여름 낮음", "80,818.7", "< 0.0001"],
      ["등급 × 도체중", "등급이 좋을수록 도체중 중앙값 상승", "697,821.7", "< 0.0001"],
    ],
    [2000, 4226, 1500, 1300]
  ),
  P([{ t: "" }]),
  ...figure("14_scatter_key.png", "가격 vs 핵심 변수 산점도 (2만 샘플, 빨간 선=추세선)"),
  ...figure("14_lowess.png", "LOWESS 곡선 — AGE×WEIGHT, AGE×가격"),
  ...figure("14_box_grade.png", "등급별 형질 박스플롯 (육질 변수는 계단형=정답지 증거)"),
  ...figure("14_box_sex.png", "성별(거세/암/수) × 도체중·근내지방·가격"),
  ...figure("14_box_season.png", "도축 계절별 가격 분포"),
  ...figure("14_box_sido.png", "시도별 가격 분포 (중앙값 낮은 순)"),
  H2("4.5 해석 (읽는 방법과 함께)"),
  box("박스플롯·검정 읽는 방법", [
    ["상자=가운데 50%(Q1~Q3), 상자 안 선=중앙값, 수염=정상 범위, 바깥 점=이상치."],
    ["그룹별 상자가 계단처럼 일정하게 오르내리면 그 변수가 그룹을 강하게 가른다는 뜻."],
    ["240만 행에서는 p값이 거의 다 유의하므로, p값보다 “상자가 실제로 얼마나 벌어졌나”(중앙값 차이)를 봐야 함."],
  ]),
  BULLET([{ t: "성별이 도체중을 결정적으로 가름", b: true }, " — 거세우(473kg)와 암소(372kg)는 중앙값이 100kg 차이. 성별은 분류·회귀 모두에서 반드시 넣어야 할 변수입니다."]),
  BULLET([{ t: "등급별 WEIGHT가 우상향 계단", b: true }, " — WEIGHT가 등급 예측에 실제로 유용함을 눈으로 재확인 (13회차 r=+0.475와 일치)."]),
  BULLET([{ t: "육질 변수의 박스가 등급별로 거의 완벽한 계단", b: true }, " — “답 보고 답 맞히기”라는 13회차 결론의 시각적 증거. 좋아할 일이 아니라 제외할 근거입니다."]),
  BULLET([{ t: "AGE×WEIGHT LOWESS가 휘어짐", b: true }, " — 월령이 늘수록 체중 증가가 둔화(성장 곡선). 직선 모델로는 부족 → 제곱항 필요."]),
  H2("4.6 다음 분석 반영"),
  BULLET("성별(JUDGE_SEX)·계절을 17회차 원-핫 인코딩 대상에 확정."),
  BULLET("LOWESS의 곡선 확인 → 17회차 Age_squared 생성 확정."),
  BULLET("“p값 대신 효과 크기” 원칙 → 이후 모든 회차의 해석 기준으로 적용."),
  PAGEBREAK()
);

// ===== 5장: 15회차 =====
content.push(
  H1("5. 15회차 — 기상 변수 탐색 (대회 핵심 주제)"),
  H2("5.1 이론: 이 분석이 무엇이고 왜 하나"),
  P("대회 주제인 “날씨가 한우에 영향을 주는가”를 처음으로 정면으로 다루는 회차입니다. 정량 회귀는 20회차에 하고, 여기서는 패턴이 보이는지 그림으로 먼저 확인합니다."),
  H3("THI (온습도지수) — 소의 더위 스트레스 지표"),
  FORMULA("THI = (1.8×T + 32) − (0.55 − 0.0055×RH) × (1.8×T − 26.8)   (T=일 최고기온°C, RH=일 평균상대습도%)"),
  P("기온과 습도를 합쳐 “소가 실제로 느끼는 더위”를 수치화한 지수입니다(NRC 1971). 같은 기온이라도 습하면 체감 더위가 심해지는 것을 반영합니다. 우리 데이터에는 각 소의 출생~도축 기간 동안 THI 등급별 일수가 집계돼 있습니다: 양호(THI<72) / 주의(72~78) / 경고(78~89) / 위험(89~98)."),
  H3("왜 일수가 아니라 비율인가"),
  FORMULA("ratio_고온 = (days_주의 + days_경고 + days_위험) / days_total     (THI 72 이상 노출 비율)"),
  FORMULA("ratio_강더위 = (days_경고 + days_위험) / days_total     ratio_위험 = days_위험 / days_total"),
  P("사육 기간이 소마다 다르므로(2년 vs 10년) 절대 일수로 비교하면 “오래 산 소”와 “더위 많이 겪은 소”가 섞입니다(13회차 일수 함정). 비율로 바꿔야 사육 기간과 무관하게 공정한 비교가 됩니다."),
  H3("Mann-Whitney U 검정 — 그룹이 2개일 때"),
  FORMULA("U = n₁n₂ + n₁(n₁+1)/2 − R₁   (R₁=그룹1의 순위합)"),
  P("두 그룹의 분포 차이를 보는 비모수 검정으로, Kruskal-Wallis의 2그룹 버전입니다. 더위 상위 25% vs 하위 25% 비교에 사용합니다."),
  H2("5.2 코드 설명"),
  ...[
    "① step8 로드, 더위 비율 3종(ratio_고온/강더위/위험) 생성.",
    "② 등급(16개)별 고온노출 비율 박스플롯 — 좋은 등급이 더위를 덜 겪었는지 확인.",
    "③ 고온노출 vs 가격 산점도(가격 있는 행만, 2만 샘플) + 상관계수.",
    "④ 출생 계절별 등급 박스플롯 + Kruskal-Wallis — 송아지 때 더위의 장기 영향.",
    "⑤ 고온노출을 4분위 구간(qcut)으로 나눠 구간별 평균 등급 + 표준오차(SE = s/√n) 오차막대. duplicates=“drop”으로 경계 중복 방어.",
    "⑥ 시도별 평균 더위×평균 등급 산점도 (점 크기=마릿수).",
    "⑦ 더위 상위 25% vs 하위 25% Mann-Whitney U 검정.",
    "⑧ step9_weather.csv 저장.",
  ].map((t) => NUM("steps15")(t)),
  H2("5.3 강의안 대비 변경점"),
  P("그림 경로(6곳)와 폰트 호출 순서 — 공통 수정뿐, 분석 로직 변경 없음."),
  H2("5.4 결과"),
  BULLET([{ t: "더위 노출 분포: ", b: true }, "ratio_고온 평균 0.415 (생애의 41.5%가 THI 72 이상), ratio_강더위 0.274, ratio_위험 0.032."]),
  BULLET([{ t: "등급별 고온노출이 반대 방향: ", b: true }, "1++A 0.422 > 1+A 0.417 > 1A 0.414 > 2A 0.409 > 3A 0.399 > 3C 0.395 — 좋은 등급일수록 더위를 “더” 겪었음."]),
  BULLET([{ t: "고온노출 × 가격 상관 r = +0.300", b: true }, " — 더위 많이 겪은 소가 비싸게 팔림(양의 상관!)."]),
  BULLET([{ t: "출생 계절: ", b: true }, "Kruskal-Wallis H=610.4, p≈5.7×10⁻¹³² — 통계적으로 유의하나 H가 성별(109만)의 0.06% 수준 = 실질 효과는 미미."]),
  BULLET([{ t: "구간별 평균 등급(U자형): ", b: true }, "낮음 9.31 → 중하 8.89 → 중상 9.54 → 높음 10.57. 중하에서 내려갔다가 다시 올라가는 U자."]),
  BULLET([{ t: "상·하위 25% 비교: ", b: true }, "하위 9.313 vs 상위 10.568, Mann-Whitney p≈0 — 차이는 확실히 존재하되 방향이 상식과 반대."]),
  ...figure("15_ratio_by_grade.png", "등급별 전 생애 고온노출 비율 박스플롯"),
  ...figure("15_bin_grade.png", "고온노출 4분위 구간별 평균 등급 (U자형, 오차막대=표준오차)"),
  ...figure("15_ratio_vs_cost.png", "고온노출 비율 vs 가격 산점도 (r=+0.300)"),
  ...figure("15_birth_season.png", "출생 계절별 등급 분포"),
  ...figure("15_sido_heat_grade.png", "시도별 더위 노출 × 평균 등급 (점 크기=마릿수)"),
  H2("5.5 해석 (읽는 방법과 함께)"),
  box("반직관적 결과를 해석하는 방법 — 상관 ≠ 인과", [
    ["“아이스크림 판매가 늘면 익사 사고도 는다” — 아이스크림 탓이 아니라 숨은 변수 “여름”이 둘 다 끌어올린 것. 이런 숨은 변수를 교란변수(confounder)라 부릅니다."],
    ["반직관적 상관이 나오면: ① 결론 내리지 말고 ② 교란 후보(지역, 농가 수준, 사육 방식)를 찾고 ③ 그 후보를 쪼개거나(층화) 회귀에서 통제한 뒤 다시 봅니다."],
  ]),
  P([
    { t: "“더위를 많이 겪은 소가 등급이 좋다”는 결과를 “더위가 한우에 좋다”로 읽으면 안 됩니다. ", b: true },
    "교란 후보 1순위는 지역입니다 — 더운 남부(전남·전북·경북)에 우량 한우 농가가 몰려 있다면, 그 지역 소는 더위도 많이 겪고 관리도 좋아서 등급이 높게 나옵니다. 진짜 원인은 더위가 아니라 농가 수준인 거죠.",
  ]),
  P([
    "그런데 시도 단위로 쪼개 보면 이 가설마저 단순하지 않습니다: ",
    { t: "광주(가장 더움 0.453)는 등급 9.18로 평균 이하, 인천(선선 0.377)은 11.13으로 최고", b: true },
    "입니다. 즉 “더운 곳=좋은 등급”이 시도 수준에서 깨집니다. 결론: 더위-등급 관계에는 지역·농가 교란이 강하게 끼어 있고, ",
    { t: "지역을 통제한 회귀(20회차) 후에야 순수한 날씨 효과를 말할 수 있습니다.", b: true },
  ]),
  P("출생 계절 효과는 “유의하지만 작음”이 정확한 표현입니다. H=610은 240만 표본에서는 사실상 잡음에 가까운 크기로, “여름 출생 송아지가 평생 불리하다” 같은 강한 주장은 데이터가 지지하지 않습니다."),
  H2("5.6 다음 분석 반영"),
  BULLET("ratio_* 3종을 17회차 파생변수 풀에 공식 편입 (트랙1·2 공용)."),
  BULLET("U자형·교란 확인 → 20회차 회귀에서 지역(시도) 변수를 반드시 함께 넣어 통제."),
  BULLET("heat_high(상위 25% 더미)의 임계값(0.449)이 여기서 결정됨."),
  PAGEBREAK()
);

// ===== 6장: 16회차 =====
content.push(
  H1("6. 16회차 — 시공간 패턴 분석"),
  H2("6.1 이론: 이 분석이 무엇이고 왜 하나"),
  P("데이터를 시간(연·월·계절)과 공간(시도·시군구) 축으로 쪼개 봅니다. 시간 추세를 무시하면 다른 변수가 그 추세를 흡수해 효과가 과장되고(가짜 효과), 지역 차이를 무시하면 더위 효과가 부풀려집니다. 또 종류가 수백 개인 지역 변수를 모델에 어떻게 넣을지(인코딩 전략)를 여기서 결정합니다."),
  H3("카디널리티와 인코딩 전략"),
  P("카디널리티 = 범주형 변수의 종류 수. 모델은 글자를 못 읽으므로 숫자로 바꿔야 하는데, 종류 수에 따라 방법이 다릅니다: 20개 이하면 원-핫 인코딩 가능, 수백 개면 컬럼 폭증으로 비효율 → Target Encoding(범주를 “그 범주의 평균 등급”으로 치환), 단 정답을 쓰므로 데이터 분할 후에만(18회차, 누수 방지)."),
  H2("6.2 코드 설명"),
  ...[
    "① step9 로드, 시간 파생 5종 생성: abatt_year/month/quarter, birth_year/month (+계절 변환).",
    "② 도축 연도별 등급·가격·마릿수 추세 (3년뿐이라 과대 해석 금지).",
    "③ 출생연도 분포 → 2015년 이전 출생(노폐우) 식별과 등급 비교.",
    "④ 도축 월별 가격·출하량 이중축 그래프 — 시즌성 확인.",
    "⑤ 시도별 요약(마릿수/등급/가격/더위) + 막대그래프 + 시도×등급 히트맵.",
    "⑥ sido/sigungu/eupmyeondong 카디널리티 점검 → 인코딩 결정.",
    "⑦ step10_spatial.csv 저장.",
  ].map((t) => NUM("steps16")(t)),
  H2("6.3 강의안 대비 변경점"),
  tbl(
    ["어디를", "어떻게", "왜"],
    [
      ["그림 경로(5곳)·폰트 순서", "공통 수정 (3.3절과 동일)", "동일"],
      ["⑥ 카디널리티 루프", "루프 시작에 if col not in df.columns: continue 가드 추가", "지역 컬럼이 없는 데이터로 실행해도 KeyError로 노트북 전체가 멈추지 않도록 방어"],
    ],
    [2600, 3200, 3226]
  ),
  P([{ t: "" }]),
  H2("6.4 결과"),
  tbl(
    ["도축연도", "마릿수", "평균 등급", "평균 가격(원/kg)"],
    [
      ["2023", "781,368", "9.35", "15,999.88"],
      ["2024", "835,792", "9.57", "15,964.38"],
      ["2025", "791,539", "9.81", "18,040.53 (+12.8%)"],
    ],
    [1800, 2200, 2200, 2826]
  ),
  P([{ t: "" }]),
  P([{ t: "월별 가격: ", b: true }, "9월 18,476원(최고·추석) → 10~12월 17,752~17,896원(연말) → 1~2월 16,142~16,144원(중간) → 5월 15,299원(최저). 설(1~2월) 급등은 데이터상 없음."]),
  P([{ t: "노폐우: ", b: true }, "2015년 이전 출생 75,972마리(3.2%), 평균 등급 3.73 vs 전체 9.58. 평균 도축 나이 10.3년 vs 정상 비육우 3.3년."]),
  P([{ t: "시도별: ", b: true }, "인천 11.13(최고, 더위 0.377) / 경북 9.82 / 전북 9.72 / 전남 9.59 / 광주·대구 9.18(더위 0.453·0.434) / 충남 9.16. 마릿수는 경북 60.6만 > 전남 40.9만 > 전북 29.6만."]),
  P([{ t: "카디널리티: ", b: true }, "sido 17종(원-핫 가능) / sigungu 184종(Target Encoding) / eupmyeondong 1,306종(Target Encoding 또는 제외, 희소 카테고리 146개)."]),
  ...figure("16_yearly.png", "도축 연도별 평균 등급·가격·마릿수"),
  ...figure("16_monthly.png", "도축 월별 가격·출하량 (9월 추석 피크)"),
  ...figure("16_sido_grade.png", "시도별 평균 등급"),
  ...figure("16_sido_grade_heatmap.png", "시도 × 등급 분포 히트맵 (행 기준 %)"),
  H2("6.5 해석 (읽는 방법과 함께)"),
  box("시계열·지역 패턴 읽는 방법", [
    ["연도 점 3개로 “추세”를 말하면 안 됨 — 비교 수준으로만. 2025 가격 +12.8%는 도매시장 전반의 거시 요인일 수 있어 “품질이 올라서”로 단정 금지."],
    ["월별 패턴은 수요(명절·연말)와 공급(출하 조절)이 겹친 결과 — 인과 분리는 불가능, 시즌성 변수로만 활용."],
    ["지역 차이는 “지역 자체”가 아니라 그 지역의 농가 구성·사육 방식 차이일 수 있음(교란)."],
  ]),
  BULLET([{ t: "추석(9월)·연말(10~12월) 시즌성은 뚜렷, 설 효과는 없음", b: true }, " — 1~2월 가격은 연평균 수준. 다만 설은 음력이라 1월 말~2월 중순을 오가므로 월 평균에 효과가 희석됐을 가능성은 남습니다(각주적 한계)."]),
  BULLET([{ t: "노폐우는 “다른 종류의 소”", b: true }, " — 평균 10.3년을 산 소는 번식을 마친 도태 암소로, 고기용 비육우(30개월 내외 도축)와 모집단 자체가 다릅니다. 등급 3.73은 날씨 탓이 아니라 집단 특성입니다. 2015년 컷은 “도축나이 8년 이상부터는 비육우로 설명 불가”라는 데이터 근거가 있습니다."]),
  BULLET([{ t: "인천 11.13의 함정", b: true }, " — 인천은 1.5만 마리뿐(경북의 2.5%)이라 소수 정예 농가 효과일 수 있음. 마릿수(점 크기)를 함께 봐야 합니다."]),
  H2("6.6 다음 분석 반영"),
  BULLET("시간 변수(연/월/분기/계절) → 트랙1 피처 + 회귀의 시즌 통제 변수."),
  BULLET("인코딩 확정: sido는 원-핫(회귀)/범주형(LightGBM), sigungu는 18회차 분할 후 Target Encoding, eupmyeondong은 Target Encoding 또는 제외."),
  BULLET("노폐우 → 18회차에서 별도 플래그 또는 분리 모델 검토 대상."),
  PAGEBREAK()
);

// ===== 7장: 17회차 =====
content.push(
  H1("7. 17회차 — 파생변수 설계"),
  H2("7.1 이론: 이 분석이 무엇이고 왜 하나"),
  P("원본에 없는 변수를 계산해 새로 만드는 단계입니다. 좋은 파생변수 하나가 하이퍼파라미터 튜닝보다 성능에 크게 기여하며, “왜 이 변수를 만들었나”가 공모전 창의성 점수의 핵심입니다."),
  H3("만든 변수와 공식"),
  FORMULA("fattening_days = 도축일 − 출생일 (사육 일수)"),
  FORMULA("daily_gain = WEIGHT / fattening_days   (일당 증체 = 하루에 몇 kg씩 컸나, 성장 효율)"),
  FORMULA("Age_squared = AGE²   (월령의 비선형 효과 — 14회차 LOWESS 곡선의 답)"),
  FORMULA("density = C2025 / AREA (사육밀도)     death_rate = death_count / C2025 (폐사율·대략치)"),
  FORMULA("density_x_heat = density × ratio_고온   (밀집×더위 상호작용 — 둘이 겹칠 때의 시너지 스트레스)"),
  FORMULA("age_optimal = 1 if 28 ≤ AGE ≤ 32 (등급 안정 구간 더미)     heat_high = 1 if ratio_고온 ≥ 상위25%"),
  H3("더미 변수 함정 (Dummy Variable Trap)"),
  P([
    "계절을 [봄,여름,가을,겨울] 4개 0/1 컬럼으로 펼치면(원-핫) 네 컬럼의 합이 항상 1이라, 3개만 알면 나머지가 자동 결정되는 완벽한 중복이 생깁니다. OLS 회귀는 이런 완전공선성에서 계수를 못 구합니다(특이행렬). 해결: ",
    { t: "drop_first=True", b: true },
    "로 기준 카테고리 하나를 떨어뜨립니다. 정보 손실은 없습니다(떨어뜨린 카테고리가 “기준점”이 될 뿐).",
  ]),
  H3("빈도 인코딩과 Target Encoding의 구분"),
  P("혈통 ID(해시)는 그대로 못 쓰므로 “이 핏줄이 데이터에 몇 번 등장하나”(빈도)로 치환합니다 — 정답을 안 쓰므로 누수 위험이 없어 지금 해도 됩니다. 반면 Target Encoding(범주→그 범주의 평균 등급)은 정답을 쓰므로 반드시 18회차 분할 후 OOF 방식으로 합니다."),
  H2("7.2 코드 설명"),
  ...[
    "① step10 로드 → 시간·효율 파생(fattening_days, daily_gain, Age_squared) 생성, 0으로 나눔·무한대 점검.",
    "② 더위 비율 확인(15회차에서 생성, 없으면 재생성).",
    "③ 임계값 더미(age_optimal, heat_high) 생성.",
    "④ 농장 파생(density, death_rate, farm_size 구간화) — AREA 결측 35%는 NaN 유지.",
    "⑤ 상호작용(density_x_heat).",
    "⑥ 혈통 3종 빈도 인코딩(KPN_NO, FATHER_CATTLE_NO, MOTHER_ANIMAL_NO).",
    "⑦ 시도·도축계절·출생계절·성별 원-핫(drop_first=True, .astype(int) — statsmodels OLS가 bool을 못 받는 문제 예방).",
    "⑧ 무한대만 NaN 처리, 결측은 채우지 않음(LightGBM이 NaN을 네이티브 처리; 회귀용 결측 처리는 20회차).",
    "⑨ 변수 분류표 출력 → step11_features.csv 저장.",
  ].map((t) => NUM("steps17")(t)),
  H2("7.3 강의안 대비 변경점 (이 회차가 가장 중요)"),
  tbl(
    ["어디를", "어떻게", "왜"],
    [
      ["셀21 변수 분류 categories", "“원본 형질(트랙2)”에 묶여 있던 WEIGHT·AGE를 “체격 원본(트랙1·2 공용)”으로 분리하고, 육질 8개만 “육질 형질(트랙2 전용)”으로 남김", "WEIGHT·AGE는 test 13개 컬럼에 포함 → 트랙1 사용 가능. 강의안 라벨대로면 18회차 피처 선택에서 핵심 변수 2개가 잘못 제외될 위험 (강의안 1장 트랙 표와도 모순이었음 — 작성자 확인 완료)"],
      ["셀21 daily_gain 주석", "“WEIGHT가 test에 있으므로 트랙1 사용 가능” 주석 명시", "daily_gain = WEIGHT/사육일수의 누수 우려가 제기됐으나, 구성 요소(WEIGHT·BIRTH_YMD·ABATT_DATE)가 전부 test에 있어 test에서도 계산 가능 → 누수 아님 확정"],
      ["셀17 더미 출력", "df[new_dummies[0]].dtype에 if new_dummies else ‘N/A’ 가드", "더미가 0개 생성될 경우 IndexError 방지 (방어적 코딩)"],
    ],
    [2200, 3300, 3526]
  ),
  P([{ t: "" }]),
  H2("7.4 결과"),
  BULLET([{ t: "daily_gain: ", b: true }, "평균 0.39kg/일 (최대 1.67) — 무한대 0건, fattening_days 0인 행 0건."]),
  BULLET([{ t: "age_optimal 39.7%, heat_high 25.0%", b: true }, " (설계대로 상위 4분위)."]),
  BULLET([{ t: "농장 변수 결측: ", b: true }, "density 91.1만 행, death_rate 58.9만 행 — NaN 유지. farm_size: 중규모 111.9만 > 소규모 71.0만 > 대규모 34.8만 > 초대규모 8.0만."]),
  BULLET([{ t: "혈통 빈도: ", b: true }, "KPN(씨수소) 중앙값 9,161회 — 소수 인기 종모우에 집중. 모계는 중앙값 1회(대부분 한 번만 등장)."]),
  BULLET([{ t: "원-핫 더미 24개", b: true }, " (시도 16 + 도축계절 3 + 출생계절 3 + 성별 2, 모두 기준 카테고리 1개씩 제외) — 정수형."]),
  BULLET([{ t: "최종: 2,408,699 × 92 변수", b: true }, " step11_features.csv 저장."]),
  P([{ t: "" }]),
  P([{ t: "신규 파생변수 검증 — grade_num과의 상관 (별도 계산):", b: true }]),
  tbl(
    ["파생변수", "r (vs 등급)", "읽기"],
    [
      ["daily_gain", "+0.522", "13~17회차 신규 변수 중 최강 — WEIGHT 원본(+0.475)보다도 강함"],
      ["fattening_days", "−0.465", "오래 키운 소(노폐우 포함)일수록 등급 낮음"],
      ["Age_squared", "−0.419", "월령 비선형 효과 포착"],
      ["age_optimal", "+0.285", "28~32개월 구간 더미도 단독으로 유효"],
      ["heat_high", "+0.151", "더위 상위 25% 더미 (방향은 교란 주의)"],
      ["ratio_고온", "+0.124", "비율 변수 (일수 함정 제거 후의 순수 신호)"],
    ],
    [2300, 1700, 5026]
  ),
  P([{ t: "" }]),
  H2("7.5 해석 (읽는 방법과 함께)"),
  box("파생변수의 가치를 판단하는 방법", [
    ["같은 정보라도 “조합”이 신호를 키울 수 있음: WEIGHT(+0.475)와 사육일수(−0.465)를 나눗셈 하나로 합친 daily_gain이 +0.522로 둘 다보다 강함."],
    ["파생변수도 누수 검사 필수: 구성 요소가 전부 test에서 계산 가능한지 확인 (daily_gain은 통과)."],
    ["결측을 억지로 채우지 않는 것도 설계 — LightGBM은 NaN을 자체 처리하며, 회귀용 대치는 20회차에서 따로."],
  ]),
  BULLET([{ t: "daily_gain이 새로운 1등 피처", b: true }, " — “하루에 살이 얼마나 잘 붙었나”는 성장 효율로, 사육 기술·사료·유전이 응축된 지표입니다. 트랙1 사용 가능이 확정되어 분류 모델의 핵심 무기가 됩니다."]),
  BULLET([{ t: "heat_high +0.151의 방향(더위↑=등급↑)은 15회차 교란 문제를 그대로 상속", b: true }, " — 단독 해석 금지, 회귀에서 지역과 함께 넣어야 합니다."]),
  BULLET([{ t: "density·death_rate는 “대략치” 한계 명시", b: true }, " — C2025와 도축 시점 불일치, death_count 기간 불명, AREA 결측 35%. 이 한계는 보고서·발표에서 정직하게 공개합니다."]),
  H2("7.6 다음 분석 반영"),
  BULLET("트랙1(LightGBM 분류) 피처 풀: 체격 2 + 기상 비율 3 + 시간 7 + 효율/임계 3 + 농장 7 + 상호작용 1 + 혈통 빈도 3 + 시군구 Target Encoding(18회차) — 육질·가격 제외."),
  BULLET("트랙2(OLS/Lasso 회귀): 원-핫 더미 24개 사용, drop_first 적용 완료. 다중공선성은 Lasso·VIF로."),
  BULLET("Target Encoding·StratifiedGroupKFold(농장 단위 분할)·클래스 불균형은 18회차에서."),
  PAGEBREAK()
);

// ===== 8장 =====
content.push(
  H1("8. 코드 변경 이력 총정리 (어디를 · 어떻게 · 왜)"),
  P("기존 13~17회차 강의안 코드와 최종 노트북의 차이 전체입니다. 분석 로직 변경은 17회차 변수 분류 1건뿐이고, 나머지는 경로·버그·방어 코드입니다."),
  tbl(
    ["#", "범위", "어디를", "어떻게", "왜"],
    [
      ["1", "13~16 공통", "그림 저장 경로 22곳", "../figures → ../../figures", "노트북 위치가 notebooks/3_eda/로 깊어져 그림이 notebooks/figures에 흩어짐 → 루트 figures/로 통일, 중복 폴더 삭제"],
      ["2", "13~16 공통", "공통 헤더 폰트 설정", "sns.set_style()을 set_korean_font()보다 먼저 호출하도록 순서 교체", "seaborn set_style이 font.family를 sans-serif로 덮어써 그림 한글이 전부 □로 깨짐. 강의안 순서 그대로면 모든 그림이 깨진 채 생성됨"],
      ["3", "16회차 셀19", "카디널리티 루프", "if col not in df.columns: continue 가드 추가", "컬럼 누락 시 KeyError로 노트북 중단 방지"],
      ["4", "17회차 셀17", "더미 dtype 출력", "if new_dummies else 'N/A' 가드 추가", "더미 0개일 때 IndexError 방지"],
      ["5", "17회차 셀21", "변수 분류 categories", "WEIGHT·AGE를 '체격(트랙1·2 공용)'으로 분리, 육질 8개만 트랙2 전용으로", "WEIGHT·AGE는 test에 있어 트랙1 사용 가능 — 잘못된 라벨이 18회차 피처 선택 오류로 이어질 위험 차단"],
    ],
    [500, 1300, 1800, 2400, 3026]
  ),
  P([{ t: "" }]),
  P([
    "1·2번의 영향: figures/ 안의 그림 20장 전체를 ",
    { t: "한글 정상 + 최신 코드 기준으로 재생성", b: true },
    "했고, step7~step11 중간 파일도 같은 실행에서 모두 다시 만들어 코드와 산출물이 1:1로 일치합니다.",
  ]),
  PAGEBREAK()
);

// ===== 9장 =====
content.push(
  H1("9. 교차 검증 Q&A — 지적과 판정"),
  P("분석 과정에서 제기된 의문(AI 교차 검토 포함)과 실제 데이터로 검증한 결과입니다."),
  tbl(
    ["#", "지적/의문", "검증 결과", "판정·조치"],
    [
      ["1", "daily_gain은 도축 후 측정값(WEIGHT) 기반이라 누수 아닌가?", "test 13개 컬럼에 WEIGHT·AGE·BIRTH_YMD·ABATT_DATE 모두 포함 → test에서도 계산 가능", "누수 아님 — 트랙1 사용 가능으로 확정, 코드 라벨 수정"],
      ["2", "“추석·설 명절 가격 급등” 표현이 맞나?", "9월 18,476원(최고), 1월 16,142·2월 16,144원(중간 수준)", "설 효과 없음 — “추석·연말”로 수정. 단 설은 음력(1월말~2월 중순 이동)이라 월 평균에 희석될 가능성 한 줄 명시"],
      ["3", "더위↑=등급↑ U자형을 어떻게 해석?", "광주(0.453, 가장 더움) 등급 9.18 / 인천(0.377, 선선) 11.13 — 시도 단위 반례 존재", "인과 아님 — 지역 교란. 회귀(20회차)에서 지역 통제 후 재해석"],
      ["4", "노폐우 2015년 컷이 자의적이지 않나?", "2015 이전 출생 평균 도축나이 10.3년 vs 이후 3.3년, 등급 3.73 vs 9.58", "근거 충분 — 비육우(30개월 도축)로 설명 불가능한 별도 집단"],
      ["5", "성별×도체중 H≈10.9억은 이론 상한(N−1≈240만) 초과 아닌가?", "재실행 결과 H=1,088,062.9 (약 109만) — 상한 이내. “10.9억”은 해설 문서의 전사 오류였음", "수치 정정 완료. H<N−1 검증 통과, 결론(차이 유의)은 불변"],
      ["6", "WEIGHT·AGE가 “트랙2 전용”으로 분류돼 있던 것", "test 컬럼 직접 확인으로 트랙1 가능 확정", "강의안 셀21의 라벨 오류 — 수정 완료 (8장 #5)"],
    ],
    [500, 2300, 3100, 3126]
  ),
  PAGEBREAK()
);

// ===== 10장 =====
content.push(
  H1("10. 다음 분석(18~21회차)에 어떻게 이어지나"),
  H2("10.1 트랙1 — 등급 분류 (제출)"),
  BULLET([{ t: "18회차 데이터 분할: ", b: true }, "StratifiedGroupKFold — 같은 농장(FARM_UNIQUE_NO)의 소가 train/검증에 동시에 들어가면 “농장 외우기” 누수가 생기므로 농장 단위로 묶어 나눕니다. 등급 비율도 유지(Stratified)."]),
  BULLET([{ t: "18회차 Target Encoding: ", b: true }, "sigungu(184종)·혈통을 분할 후 train에서만 OOF 방식으로 인코딩 (16회차 카디널리티 결정의 실행)."]),
  BULLET([{ t: "19회차 LightGBM: ", b: true }, "피처 풀 = 체격 2(WEIGHT·AGE) + daily_gain 등 파생 + 기상 비율 + 시간 + 농장 + 혈통 빈도. NaN은 모델이 자체 처리. 클래스 불균형은 class_weight로."]),
  H2("10.2 트랙2 — 기상 영향 회귀 (분석)"),
  BULLET([{ t: "20회차 OLS/Lasso: ", b: true }, "등급·가격을 종속변수로, ratio_* 더위 변수 + 시도 더미(원-핫, drop_first) + 시즌 더미 + 통제변수(성별·월령·농장)를 함께 투입. 시도 더미가 들어가야 15회차의 지역 교란이 통제되어 “순수한 더위 효과”의 부호와 크기를 읽을 수 있습니다."]),
  BULLET([{ t: "다중공선성 처리: ", b: true }, "13회차 |r|≥0.7 목록 기반으로 VIF 점검, Lasso의 자동 변수 선택 활용 (C2023~25는 하나만)."]),
  BULLET([{ t: "21회차 해석: ", b: true }, "상관≠인과 원칙으로 결론 작성. U자형·노폐우·인천 표본 수 같은 함정을 보고서에 명시."]),
  H2("10.3 EDA가 다음 단계에 넘기는 결정 사항 요약"),
  tbl(
    ["결정", "근거 회차"],
    [
      ["육질 8개+가격 → 트랙1 제외, 트랙2 전용", "13 (r>0.87 정답지)"],
      ["WEIGHT·AGE·daily_gain → 트랙1 핵심 피처", "13·17 (+test 컬럼 검증)"],
      ["성별·계절 → 필수 변수 (원-핫 완료)", "14 (H=109만)"],
      ["기상은 일수 대신 비율(ratio_*) 사용", "13(일수 함정)·15"],
      ["더위 효과는 지역 통제 후 재추정", "15·16 (U자형·교란)"],
      ["sigungu Target Encoding은 분할 후(18회차)", "16 (카디널리티 184)"],
      ["노폐우(2015 이전 출생) 별도 인지·플래그 검토", "16 (등급 3.73 vs 9.58)"],
      ["Age² 등 비선형 항 포함", "14 (LOWESS 곡선)"],
    ],
    [5526, 3500]
  ),
  PAGEBREAK()
);

// ===== 11장 =====
content.push(
  H1("11. 총정리"),
  box("13~17회차를 세 문장으로", [
    ["① 쓸 수 있는 재료를 골라냈다 — 육질·가격은 정답지(누수)라 제출 모델에서 빼고, WEIGHT·AGE·daily_gain은 test에 있음을 확인해 당당히 쓴다."],
    ["② 날씨-등급 관계는 겉보기와 다르다 — “더위 많은 소가 등급 좋음”(U자형)은 지역·농가 교란 탓일 가능성이 크고, 시도 반례(광주 9.18 vs 인천 11.13)까지 확인했으므로 회귀에서 지역을 통제한 뒤에만 결론 낸다."],
    ["③ 무기를 만들었다 — daily_gain(r=+0.522)을 비롯한 92개 변수의 step11_features.csv가 완성됐고, 인코딩·분할·불균형 처리 계획까지 18회차로 넘긴다."],
  ], "FFF3E0"),
  P([{ t: "" }]),
  H2("회차별 한 줄 요약"),
  tbl(
    ["회차", "한 줄 요약"],
    [
      ["13 상관관계", "변수 지도 작성 — 정답지(육질) 식별, 일수 함정 발견, 곡선 의심 1쌍(AGE)"],
      ["14 산점도·박스플롯", "눈으로 검증 — 성별 결정적(중앙값 100kg 차), WEIGHT 유용 확인, 성장 곡선 확인"],
      ["15 기상 탐색", "더위 비율 변수 완성 — U자형 반직관 결과와 지역 교란 발견 (인과 아님)"],
      ["16 시공간", "추석·연말 시즌성 확인(설 효과 없음), 노폐우 식별, 인코딩 전략 확정"],
      ["17 파생변수", "daily_gain 등 92변수 완성 — 트랙 분류 오류 수정, 더미 함정 회피"],
    ],
    [2400, 6626]
  ),
  P([{ t: "" }]),
  H2("우리가 지킨 원칙 5가지"),
  ...[
    "누수 검사 우선 — 모든 변수는 “test에서 계산 가능한가”부터 확인한다.",
    "상관 ≠ 인과 — 반직관적 결과는 교란 후보를 통제하기 전까지 결론 내리지 않는다.",
    "p값보다 효과 크기 — 240만 행에서는 모든 게 유의하므로 차이의 실제 크기를 본다.",
    "결측·한계의 정직한 공개 — density·death_rate의 대략치 한계, 인천의 표본 수까지 명시한다.",
    "코드와 산출물의 1:1 일치 — 모든 그림·CSV는 최종 코드의 재실행으로 생성한다.",
  ].map((t) => NUM("principles")(t)),
  PAGEBREAK()
);

// ===== 부록 =====
content.push(
  H1("부록 A. 공식 모음"),
  tbl(
    ["이름", "공식", "쓰임"],
    [
      ["피어슨 상관", "r = Σ(xᵢ−x̄)(yᵢ−ȳ) / √[Σ(xᵢ−x̄)²·Σ(yᵢ−ȳ)²]", "직선 관계 (13회차)"],
      ["스피어만 상관", "ρ = 1 − 6Σdᵢ²/(n(n²−1))", "순위·단조 관계 (13회차)"],
      ["Kruskal-Wallis", "H = [12/(N(N+1))]·Σ(Rⱼ²/nⱼ) − 3(N+1), 최대 N−1", "3그룹+ 차이 검정 (14·15회차)"],
      ["Mann-Whitney", "U = n₁n₂ + n₁(n₁+1)/2 − R₁", "2그룹 차이 검정 (15회차)"],
      ["THI (NRC 1971)", "(1.8T+32) − (0.55−0.0055RH)(1.8T−26.8)", "더위 스트레스 지수 (15회차)"],
      ["표준오차", "SE = s/√n", "구간 평균의 오차막대 (15회차)"],
      ["VIF", "VIFⱼ = 1/(1−R²ⱼ), 10 이상 위험", "다중공선성 정량화 (20회차 예정)"],
      ["더위 비율", "ratio_고온 = (주의+경고+위험일)/총일수", "일수 함정 제거 (15회차)"],
      ["일당 증체", "daily_gain = WEIGHT/fattening_days", "성장 효율 (17회차)"],
      ["사육밀도", "density = C2025/AREA", "농장 환경 (17회차)"],
    ],
    [2100, 4226, 2700]
  ),
  P([{ t: "" }]),
  H1("부록 B. 용어집"),
  tbl(
    ["용어", "쉬운 설명"],
    [
      ["누수 (Leakage)", "예측 시점에 알 수 없는 정보를 학습에 쓰는 것. 기준은 “test에 있느냐”"],
      ["교란변수 (Confounder)", "두 변수 모두에 영향을 줘서 가짜 상관을 만드는 숨은 제3의 변수"],
      ["다중공선성", "설명변수끼리 너무 닮아 회귀 계수가 불안정해지는 문제"],
      ["비모수 검정", "정규분포 가정 없이 순위로 하는 검정 (Kruskal-Wallis, Mann-Whitney)"],
      ["카디널리티", "범주형 변수의 종류 수 (sido 17, sigungu 184, 읍면동 1,306)"],
      ["원-핫 인코딩", "범주를 0/1 컬럼들로 펼치기. drop_first=True로 더미 함정 회피"],
      ["빈도 인코딩", "범주를 등장 횟수로 치환 — 정답을 안 쓰므로 분할 전에도 안전"],
      ["Target Encoding", "범주를 그 범주의 평균 정답으로 치환 — 정답을 쓰므로 분할 후 OOF로만"],
      ["OOF (Out-of-Fold)", "자기 fold를 뺀 나머지 데이터로 인코딩 값을 계산하는 누수 방지 기법"],
      ["LOWESS", "점 구름의 실제 모양을 보여주는 부드러운 곡선 적합"],
      ["U자형 관계", "내려갔다 올라가는 비선형 관계 — 제곱항이나 구간화로 모델링"],
      ["노폐우", "번식을 마치고 도태된 늙은 암소 — 비육우와 다른 모집단"],
      ["StratifiedGroupKFold", "등급 비율 유지 + 같은 농장은 같은 fold로 묶는 교차검증 분할"],
    ],
    [2700, 6326]
  ),
  P([{ t: "" }]),
  H1("부록 C. 산출물 목록"),
  tbl(
    ["종류", "파일", "비고"],
    [
      ["노트북", "notebooks/3_eda/13~17_*.ipynb (5개)", "최종 코드, 출력 포함 재실행 완료"],
      ["중간 데이터", "data/processed/3_eda/step7~step11 (5개 CSV)", "step11 = 2,408,699 × 92"],
      ["그림", "figures/*.png (20장)", "한글 정상, 최종 코드로 재생성"],
      ["해설 문서", "notebooks/3_eda/EDA_13-17_후배용_최종해설.md", "학부 3학년용 마크다운판"],
      ["본 보고서", "보고서_EDA_13-17회차_팀공유.docx", "팀 공유용"],
    ],
    [1800, 4226, 3000]
  )
);

// ---------- 문서 조립 ----------
const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 20 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial", color: "17407A" },
        paragraph: { spacing: { before: 280, after: 200 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: "2A5DA8" },
        paragraph: { spacing: { before: 220, after: 140 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 22, bold: true, font: "Arial", color: "444444" },
        paragraph: { spacing: { before: 160, after: 100 }, outlineLevel: 2 } },
    ],
  },
  numbering: {
    config: [
      { reference: "bullets",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 600, hanging: 300 } } } }] },
      ...["steps13", "steps14", "steps15", "steps16", "steps17", "principles"].map((ref) => ({
        reference: ref,
        levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 600, hanging: 300 } } } }],
      })),
    ],
  },
  sections: [
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 }, // A4
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
        },
      },
      headers: {
        default: new Header({
          children: [
            new Paragraph({
              alignment: AlignmentType.RIGHT,
              children: [new TextRun({ text: "한우 EDA 13~17회차 종합 보고서", size: 16, color: "888888" })],
              border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC", space: 2 } },
            }),
          ],
        }),
      },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({ text: "- ", size: 16, color: "888888" }),
                new TextRun({ children: [PageNumber.CURRENT], size: 16, color: "888888" }),
                new TextRun({ text: " -", size: 16, color: "888888" }),
              ],
            }),
          ],
        }),
      },
      children: content,
    },
  ],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync(path.join(__dirname, "보고서_EDA_13-17회차_팀공유.docx"), buffer);
  console.log("저장 완료: 보고서_EDA_13-17회차_팀공유.docx", buffer.length, "bytes");
});
