// 한우 데이터 병합·전처리 (1_merge step1~5 / 2_preprocess step1~6) 팀 공유 보고서
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  Header, Footer, AlignmentType, LevelFormat, TableOfContents, HeadingLevel,
  BorderStyle, WidthType, ShadingType, PageNumber, PageBreak,
} = require("docx");

const FIGDIR = path.join(__dirname, "figures");
const CW = 9026;

// ---------- 공통 헬퍼 (make_team_report.js와 동일) ----------
function runs(parts) {
  if (!Array.isArray(parts)) parts = [parts];
  return parts.map((p) =>
    typeof p === "string"
      ? new TextRun({ text: p })
      : new TextRun({ text: p.t, bold: p.b, italics: p.i, color: p.color, size: p.sz, font: p.font })
  );
}
const P = (parts, opts = {}) => new Paragraph({ children: runs(parts), spacing: { after: 120 }, ...opts });
const H1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(t)] });
const H2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(t)] });
const H3 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun(t)] });
const BULLET = (parts) =>
  new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: runs(parts), spacing: { after: 60 } });
const NUM = (ref) => (parts) =>
  new Paragraph({ numbering: { reference: ref, level: 0 }, children: runs(parts), spacing: { after: 60 } });
const PAGEBREAK = () => new Paragraph({ children: [new PageBreak()] });
const FORMULA = (t) =>
  new Paragraph({
    children: [new TextRun({ text: t, font: "Cambria Math", size: 22 })],
    shading: { fill: "EEF3FA", type: ShadingType.CLEAR },
    indent: { left: 360 },
    spacing: { before: 60, after: 120 },
  });
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
          children: [new TextRun({ text: String(text), bold: opts.bold, size: 18, color: opts.color || "000000" })],
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
      new TableRow({ children: r.map((c, i) => mkCell(c, { w: widths[i], fill: ri % 2 ? "F2F6FA" : "FFFFFF" })) })
  );
  return new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: widths, rows: [headerRow, ...bodyRows] });
}
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

// ---------- 내용 ----------
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
    children: [new TextRun({ text: "한우 데이터 병합·전처리 보고서", size: 52, bold: true })],
    spacing: { after: 120 },
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "원본 5종 → 분석용 단일 테이블 (step6_preprocess) 구축 전 과정", size: 26, color: "444444" })],
    spacing: { after: 600 },
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "최종 산출: 2,408,699마리 × 44변수", size: 22, color: "555555" })],
    spacing: { after: 80 },
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "병합 5단계(폐사·혈통·농장·기상·검증) + 전처리 6단계(컬럼 점검·정제·검증)", size: 22, color: "555555" })],
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
  P("EDA(13~17회차) 이전에 수행한 데이터 병합·전처리 전 과정을 정리한 문서입니다. 분석을 처음 접하는 학부생 3학년 기준으로, 각 단계가 “무엇을, 어떤 원리로, 왜” 했는지를 빠짐없이 설명합니다. 단계별 장(3~8장)은 이론 → 코드 설명 → 결과 → 해석 순서이고, 9장이 최종 결과, 10장이 해석 종합, 11장이 다음 분석 연결, 12장이 총정리입니다."),
  box("이 단계가 왜 중요한가", [
    ["모델 성능의 절반은 데이터 품질에서 나옵니다. 병합 키가 틀리면 모든 분석이 틀리고, 결측 코드(-99)를 숫자로 두면 평균·상관이 전부 오염됩니다."],
    ["여기서 만든 step6_preprocess.csv(240만 × 44)가 EDA 13~17회차와 이후 모든 모델링의 출발점입니다."],
  ]),
  PAGEBREAK()
);

// ===== 2장 =====
content.push(
  H1("2. 큰 그림 — 원본 5종과 파이프라인"),
  H2("2.1 원본 데이터 5종 + 추가 기상 3종"),
  tbl(
    ["파일", "크기", "내용", "병합 키"],
    [
      ["hanwoo_train.csv", "2,408,699 × 23", "소 개체별 도축 기록 (등급·형질·가격)", "기준 테이블"],
      ["hanwoo_death.csv", "321,389 × 4", "농장별 폐사 기록", "FARM_UNIQUE_NO"],
      ["hanwoo_lineage.csv", "1,809,455 × 8", "개체별 혈통 (부·모·조부모)", "CATTLE_NO"],
      ["hanwoo_area.csv", "91,896 × 5", "농장 면적·연도별 사육두수", "FARM_UNIQUE_NO"],
      ["hanwoo_weather.csv", "973,248 × 7", "관측소별 일별 기상", "stn + date"],
      ["OBS_ASOS_DD 3개 파일", "733,947행", "기상청 ASOS 추가 기상 (과거 기간 보강)", "stn + date"],
    ],
    [2400, 1700, 3226, 1700]
  ),
  P([{ t: "" }]),
  P([
    "키(key)란 두 테이블을 이어 붙일 때 “같은 대상”임을 알려주는 열쇠 컬럼입니다. 소 단위 정보(혈통)는 소 ID(CATTLE_NO)로, 농장 단위 정보(폐사·면적)는 농장 ID(FARM_UNIQUE_NO)로, 기상은 관측소+날짜(stn+date)로 잇습니다. ",
    { t: "어떤 단위로 합치느냐가 병합 설계의 전부입니다.", b: true },
  ]),
  H2("2.2 파이프라인 한눈에 보기"),
  tbl(
    ["단계", "스크립트", "하는 일", "산출"],
    [
      ["병합 1", "step1_death", "농장별 폐사 건수 집계 → train에 붙임", "step1_death.csv"],
      ["병합 2", "step2_lineage", "혈통 7개 컬럼 붙임", "step2_lineage.csv"],
      ["병합 3", "step3_area", "농장 면적·사육두수 합산 후 붙임", "step3_area.csv"],
      ["병합 4-1", "step4-1_weather", "기존+신규 기상 4개 파일 통합", "step4-1_weather.csv"],
      ["병합 4-2", "step4-2_weather", "기상 이상치 제거·결측 보간", "step4-2_weather.csv"],
      ["병합 4-3", "step4-3_weather", "THI 계산·등급 분류", "step4-3_weather.csv"],
      ["병합 4-4", "step4-4_weather", "소별 생애 THI 노출 일수 집계 → 붙임", "step4-4_weather.csv"],
      ["병합 5", "step5_merge", "통합 검증 + -99 일괄 처리", "step5_merge.csv"],
      ["전처리 1~5", "step1~5", "출처별 컬럼 정밀 점검·정제", "step1~5 (2_preprocess)"],
      ["전처리 6", "step6_preprocess", "최종 검증 (9개 섹션 PASS/FAIL)", "step6_preprocess.csv ★"],
    ],
    [1300, 2100, 3826, 1800]
  ),
  P([{ t: "" }]),
  H2("2.3 핵심 개념 — LEFT JOIN과 결측 코드"),
  P([
    { t: "LEFT JOIN(왼쪽 우선 병합): ", b: true },
    "기준 테이블(train)의 행은 하나도 잃지 않고, 상대 테이블에 짝이 있으면 붙이고 없으면 NaN으로 둡니다. 모든 병합 후 “행 수가 2,408,699 그대로인가”를 확인하는 이유가 이것입니다 — 행이 늘었다면 키 중복(1:N 병합 사고), 줄었다면 INNER JOIN 사고입니다.",
  ]),
  P([
    { t: "결측 코드 -99: ", b: true },
    "이 데이터는 “값 없음”을 -99라는 숫자로 기록해 뒀습니다. 이를 그대로 두면 평균·상관·모델이 전부 오염되므로(예: 가격 평균이 -99에 끌려 내려감) 반드시 NaN으로 바꿔야 합니다. 수치형 -99와 문자형 “-99”/“-99.0”이 섞여 있어 둘 다 처리합니다.",
  ]),
  PAGEBREAK()
);

// ===== 3장: 병합 1·2·3 =====
content.push(
  H1("3. 병합 1~3단계 — 폐사·혈통·농장"),
  H2("3.1 이론: groupby 집계와 병합 단위 맞추기"),
  P([
    "train은 “소 1마리 = 1행”인데 death는 “폐사 1건 = 1행”, area는 “허가 1건 = 1행”입니다. 단위가 다른 테이블을 그대로 붙이면 행이 불어나므로(1:N), 먼저 ",
    { t: "농장 단위로 집계(groupby)해 1:1로 만든 뒤", b: true },
    " 붙입니다. 폐사는 건수(count), 면적·두수는 합(sum — 한 농장이 허가를 여러 개 가질 수 있어 합산)을 씁니다.",
  ]),
  FORMULA("death_count(농장) = count(해당 농장의 폐사 기록)     C·AREA(농장) = Σ(해당 농장의 허가별 값)"),
  H2("3.2 코드 설명"),
  ...[
    "① [폐사] 날짜 컬럼(YYYYMMDD 정수)을 datetime으로 변환, DEAD_REASON의 문자 “-99”를 결측 처리. FARM_UNIQUE_NO별 폐사 건수 집계 후 LEFT JOIN.",
    "② [혈통] lineage의 CATTLE_NO 중복이 0임을 먼저 확인(중복 = 1:N 사고 위험) → 소 ID로 LEFT JOIN. KPN_NO(씨수소)·부모·조부모 7개 컬럼이 붙음.",
    "③ [농장] C2023~C2025·AREA의 -99를 결측으로 바꾼 뒤 농장별 합산 → LEFT JOIN. 각 병합마다 “병합 전후 행 수·열 수”를 출력해 행 보존을 확인.",
  ].map((t) => NUM("s3")(t)),
  H2("3.3 결과"),
  tbl(
    ["병합", "키", "추가 컬럼", "결측(붙지 않은 행)", "결측의 의미"],
    [
      ["폐사", "FARM_UNIQUE_NO", "death_count 1개", "518,737행 (21.5%)", "death 기록이 아예 없는 농장 소속 — “폐사 0건”과는 다름"],
      ["혈통", "CATTLE_NO", "KPN_NO 등 7개", "922,653행 (38.3%)", "혈통 미등록 개체"],
      ["농장", "FARM_UNIQUE_NO", "C2023~25, AREA 4개", "AREA 848,394행 (35.2%)", "area에 없는 농장 + 원본 -99"],
    ],
    [900, 1900, 1900, 1900, 2426]
  ),
  P([{ t: "" }]),
  P("세 병합 모두 행 수 2,408,699 유지 — 1:1 병합 성공."),
  H2("3.4 해석 (읽는 방법과 함께)"),
  box("병합 결측을 읽는 방법", [
    ["병합 후 결측은 “병합 실패”가 아니라 정보 그 자체입니다. 다만 출처를 구분해야 합니다."],
    ["death_count NaN = “폐사 기록 없음”(0건일 수도, 미신고일 수도) ≠ 0. 그래서 0으로 채우지 않고 NaN으로 둡니다."],
    ["혈통 NaN 38%는 미등록 개체 — 이후 빈도 인코딩에서 “결측도 하나의 그룹”으로 처리됩니다(17회차)."],
  ]),
  PAGEBREAK()
);

// ===== 4장: 병합 4 기상 =====
content.push(
  H1("4. 병합 4단계 — 기상 데이터 (이 파이프라인의 하이라이트)"),
  H2("4.1 (4-1) 기상 4개 파일 통합"),
  P("기존 hanwoo_weather(973,248행)는 기간이 부족해, 기상청 ASOS 일자료 3개 파일(733,947행)로 과거를 보강했습니다. 신규 파일은 ① train에 등장하는 관측소만 필터링하고 ② 컬럼명을 기존 형식(stn, date, ta_min, ta_max, rn_day, rhm_avg, ws_davg)으로 통일한 뒤 ③ 4개를 세로로 이어 붙였습니다(concat)."),
  BULLET([{ t: "결과: 1,645,237행, 관측소 444곳, 1996-01-01 ~ 2025-12-31 (30년), stn+date 중복 0", b: true }, " — 가장 오래된 소(1996년생)의 생애 전체를 덮는 기간 확보."]),
  H2("4.2 (4-2) 기상 품질 관리 — 이상치와 결측"),
  H3("이론 ① 이동 중앙값(rolling median) 이상치 탐지"),
  FORMULA("dev(t) = x(t) − median( x(t−7) … x(t+7) )   — 같은 관측소의 앞뒤 15일 중앙값과의 편차"),
  P([
    "센서 오류(예: 7월에 영하 10도)는 “그 시기 평소 값”과 비교해야 잡힙니다. 평균이 아니라 ",
    { t: "중앙값", b: true },
    "을 쓰는 이유: 평균은 이상치 자신에게 끌려가지만 중앙값은 끄떡없습니다(강건성). 편차 기준은 실제 사례를 눈으로 확인한 뒤 보수적으로 결정했습니다 — ta_min은 편차 −20 미만, ta_max는 편차 +18 초과만 이상치로 보고 NaN 처리. 1996년 2월의 +24도 같은 “진짜 이상고온”은 여러 관측소가 동시에 기록했으므로 살렸습니다.",
  ]),
  H3("이론 ② 선형 보간 (linear interpolation)"),
  FORMULA("x(t) = x(a) + ( x(b) − x(a) ) × (t − a) / (b − a)   — 결측 구간의 양끝 값을 직선으로 연결"),
  P([
    "하루이틀 빠진 기온은 앞뒤 날과 비슷할 것이므로 직선으로 메웁니다. 단 ",
    { t: "limit=3 (연속 3일까지만)", b: true },
    " — 일주일씩 빠진 구간을 직선으로 메우면 실제 한파·폭염을 지워버리기 때문입니다. 보간은 반드시 관측소별로(groupby) 합니다 — 안 그러면 서울의 마지막 날과 부산의 첫날이 이어져 버립니다.",
  ]),
  H3("처리 순서 (순서 자체가 설계)"),
  ...[
    "-99 → NaN (결측 코드 제거)",
    "이동 중앙값 편차로 센서 오류 제거 (ta_min 편차<−20, ta_max 편차>+18)",
    "강수량(rn_day) NaN → 0 (관측 없음 = 비 안 옴으로 간주)",
    "관측소별 선형 보간 (연속 3일 한도)",
    "보간 후 ta_min > ta_max 모순 행 → 둘 다 NaN (물리적으로 불가능한 값)",
    "남은 결측 → 관측소별+월별 중앙값 (그 동네 그 계절의 평소 값)",
    "그래도 남으면 → 관측소별 중앙값",
  ].map((t) => NUM("s42")(t)),
  P("뒤로 갈수록 거친 방법을 쓰는 “단계적 대치”입니다 — 정밀한 방법(보간)을 먼저, 거친 방법(중앙값)은 최후에 최소한으로."),
  H2("4.3 (4-3) THI 계산과 등급"),
  FORMULA("THI = (1.8 × ta_max + 32) − (0.55 − 0.0055 × rhm_avg) × (1.8 × ta_max − 26.8)"),
  P("일 최고기온과 일 평균상대습도로 “소가 느끼는 더위”를 수치화합니다(NRC 1971, 국립축산과학원 한우 기준). 같은 기온이라도 습할수록 THI가 높아집니다 — 소는 땀 대신 호흡으로 열을 식히는데 습하면 증발 냉각이 안 되기 때문입니다."),
  tbl(
    ["등급", "THI 구간", "전체 일수 중 비율", "의미"],
    [
      ["양호", "< 72", "61.9%", "쾌적"],
      ["주의", "72 ~ 78", "14.6%", "경증 더위 스트레스"],
      ["경고", "78 ~ 89", "21.4%", "중증 — 사료 섭취 감소"],
      ["위험", "89 ~ 98", "2.1%", "심각 — 생산성 직접 타격"],
      ["폐사", "≥ 98", "0% (최대 97.65)", "데이터 기간엔 미발생"],
    ],
    [1200, 1600, 2400, 3826]
  ),
  P([{ t: "" }]),
  ...figure("00_merge_thi_dist.png", "일별 THI 분포와 등급 경계 (444개 관측소 × 30년)"),
  H2("4.4 (4-4) 소별 생애 더위 노출 집계 — 누적합 알고리즘"),
  H3("문제: 240만 마리 × 평균 1,000일을 어떻게 빨리 세나"),
  P("소 한 마리마다 “출생일~도축일 사이에 양호/주의/경고/위험이 며칠씩이었나”를 세야 합니다. 소마다 날짜를 일일이 세면 240만 × 1,000일 = 수십억 번 연산이라 사실상 불가능합니다. 해결책이 누적합(prefix sum)입니다."),
  FORMULA("S(k) = x(1) + x(2) + … + x(k) 를 미리 만들어 두면,  구간합 Σ x(lo..hi) = S(hi) − S(lo)  — 단 한 번의 뺄셈"),
  P([
    "관측소별로 날짜순 누적합 테이블을 만들어 두고, 각 소의 출생일·도축일 위치(lo, hi)는 ",
    { t: "이진 탐색(searchsorted)", b: true },
    "으로 찾습니다. 그 결과 소 한 마리의 생애 노출 집계가 “뺄셈 한 번”이 되어, 240만 마리 전체가 몇 초 만에 끝납니다. THI 등급은 더미 변수(d_양호=0/1 등)로 바꿔 누적합하면 “등급별 일수”가, 강수량·풍속·최저기온은 누적합을 일수로 나누면 “생애 평균”이 됩니다.",
  ]),
  H3("결과와 자체 검증"),
  BULLET("추가된 컬럼 9개: days_total, days_양호/주의/경고/위험/폐사, rn_day_mean, ws_davg_mean, ta_min_mean."),
  BULLET([{ t: "검증: 등급별 일수의 합 = days_total이 240만 행 전부 성립", b: true }, " — 집계 논리에 구멍이 없다는 수학적 확인."]),
  BULLET("days_폐사는 전부 0 — THI 최대가 97.65로 폐사 기준(98) 미달이기 때문(상수 컬럼이지만 “정상”으로 문서화)."),
  H2("4.5 해석 (읽는 방법과 함께)"),
  box("기상 전처리를 평가하는 방법", [
    ["“얼마나 채웠나”가 아니라 “무엇을 지키며 채웠나”를 봅니다: 물리 법칙(ta_min ≤ ta_max), 지역성(관측소별), 계절성(월별 중앙값), 실제 사건 보존(진짜 한파·이상고온은 안 지움)."],
    ["보간 한도(3일)와 이상치 기준(−20/+18)은 임의값이 아니라 의심 사례를 눈으로 확인한 뒤 정한 보수적 기준입니다."],
  ]),
  BULLET("30년 × 444관측소의 빈틈없는 기상 테이블이 만들어졌고, 이것이 15회차 더위 비율(ratio_*)과 17회차 파생변수의 토대가 됩니다."),
  BULLET("THI 등급의 36.1%가 주의 이상 — 한우는 생애 3분의 1 이상을 더위 스트레스 속에서 보냅니다. “더위가 변수가 될 수 있다”는 대회 주제의 데이터적 근거."),
  PAGEBREAK()
);

// ===== 5장: 병합 5 검증 =====
content.push(
  H1("5. 병합 5단계 — 통합 검증 (step5_merge)"),
  H2("5.1 이론: 검증은 “재계산 대조”다"),
  P([
    "병합이 맞았는지 어떻게 알까요? ",
    { t: "원본에서 독립적으로 다시 계산한 값과 병합 결과를 대조", b: true },
    "합니다. 예: KPN_NO 결측 수는 “lineage에 없는 소의 수 + 원본 ‘-99’ 3,570개”와 정확히 같아야 합니다. 이렇게 결측 수를 ‘역산’으로 설명할 수 있으면 의도하지 않은 데이터 손실이 없다는 뜻입니다.",
  ]),
  H2("5.2 코드 설명 — check 함수와 7개 섹션"),
  ...codeBlock(`def check(name, condition):
    print(f"  [{'PASS' if condition else 'FAIL'}] {name}")
    return condition`),
  P("모든 검증을 PASS/FAIL 한 줄로 찍는 단순한 함수가 핵심 장치입니다. 사람 눈 대신 조건식이 판정하므로 누락도 착오도 없습니다. 섹션 구성:"),
  ...[
    "섹션 0 [-99 처리]: 수치형 -99 → NaN (육질 8개 + COST_AMT), 문자형 “-99” → NaN (KPN_NO 3,570개)",
    "섹션 1 [기본 구조]: 행 수 = 원본, 열 수 = 44, CATTLE_NO 중복·결측 0, ID 집합 원본과 동일",
    "섹션 2 [병합 정합성]: 출처별 결측 수 역산 대조 + 무작위 샘플 값이 원본과 일치하는지",
    "섹션 3 [날짜 무결성]: BIRTH_YMD ≤ ABATT_DATE, 1970-01-01(과거 DB 오류값) 부재, 도축연도 2023~25",
    "섹션 4 [타입]: 날짜 3개 datetime, 수치 25개 numeric, 범주 5개 문자",
    "섹션 5 [결측 현황]: 전체 NaN 컬럼 없음, 육질 결측은 전부 등외 등급에서만 발생",
    "섹션 6~7 [컬럼·타깃]: 상수 컬럼 없음(days_폐사 예외 문서화), 16개 등급 확인",
  ].map((t) => NUM("s5")(t)),
  H2("5.3 결과·해석"),
  BULLET([{ t: "육질 결측(약 5천 건, 0.2%)이 전부 ‘등외’ 등급", b: true }, " — 등외는 정식 육질 판정을 받지 않으므로 측정값이 없는 것이 자연스럽습니다. 결측 패턴이 도메인 지식과 일치하는지 확인하는 좋은 예."]),
  BULLET([{ t: "BIRTH_YMD ≤ ABATT_DATE 전 행 성립", b: true }, " — “도축보다 늦게 태어난 소” 같은 논리 모순이 없습니다. 이런 검증이 없으면 사육일수 계산(17회차 fattening_days)이 음수가 되는 사고로 이어집니다."]),
  PAGEBREAK()
);

// ===== 6장: 전처리 =====
content.push(
  H1("6. 전처리 1~5단계 — 출처별 컬럼 정밀 점검"),
  H2("6.1 이론: 점검의 4단계 루틴"),
  P("모든 컬럼을 같은 루틴으로 봅니다: ① 결측 수 ② 기술통계(describe) ③ 최솟값·최댓값 상위 10개 ④ 수상한 값은 해당 행 전체를 꺼내 맥락 확인. “이상해 보이는 값”을 기계적으로 지우지 않고, 원본까지 추적해 원인을 밝힌 뒤 처리하는 것이 원칙입니다."),
  H2("6.2 발견하고 처리한 것들 (이 단계의 실제 성과)"),
  tbl(
    ["#", "발견", "원인 추적", "처리"],
    [
      ["1", "WGRADE에 문자 “-99.0” 잔존", "원본이 문자형이라 수치 -99 치환에 안 걸림", "문자·수치 둘 다 치환 → A/B/C/D만 남김"],
      ["2", "COST_AMT에 0원 1건", "낙찰 무효 등 비정상 거래로 추정", "0 → NaN (가격 분석 오염 방지)"],
      ["3", "AREA·C2023~25에 0 다수 — 원본에는 0이 없음", "groupby.sum()이 “전부 NaN인 그룹”을 0으로 만드는 pandas 동작 (병합 3단계의 숨은 버그)", "0 → NaN 일괄 수정 + step6에서 ‘0 없음’ 검증 추가"],
      ["4", "혈통에 수상한 해시 1개가 비정상적으로 자주 등장", "원본 추적 결과 “부계 미상”을 뜻하는 더미(placeholder) 값으로 확인", "값은 유지하되 문서화 — 17회차 빈도 인코딩에서 ‘미상 그룹’으로 자연 처리"],
      ["5", "death_count 100건 이상 대형 농장 존재", "대규모 농장의 정상 기록으로 확인", "유지 (이상치 아님)"],
      ["6", "WEIGHT 3kg, AGE 1개월 등 극단값", "송아지 폐사 직전 도축 등 실제 사례", "유지 — 등외 등급과 일관, 실제 데이터"],
    ],
    [500, 2300, 3300, 2926]
  ),
  P([{ t: "" }]),
  H3("3번 버그의 교훈 (pandas 함정)"),
  ...codeBlock(`# pandas의 sum()은 그룹 전체가 NaN이어도 0을 돌려준다
pd.Series([np.nan, np.nan]).sum()   # → 0.0  (NaN이 아니라!)
# 그래서 -99를 NaN으로 바꾼 뒤 농장별 sum()을 하면
# "정보가 전혀 없는 농장"이 "면적 0인 농장"으로 둔갑한다`),
  P("“원본에 0이 있는가?”를 역추적해서 잡았습니다(원본 AREA에 0은 없었음 → 병합 산물). 집계 후에는 항상 “원본에 없던 값이 생기지 않았나”를 의심해야 한다는 교훈입니다."),
  PAGEBREAK()
);

// ===== 7장: step6 최종 검증 =====
content.push(
  H1("7. 전처리 6단계 — 최종 검증 (step6_preprocess)"),
  H2("7.1 무엇을 검증했나"),
  P("step5까지의 모든 처리가 끝난 데이터에 9개 섹션의 검증을 다시 돌립니다. 병합 5단계 검증과 겹치는 항목도 일부러 다시 합니다 — 전처리 과정에서 무언가 깨졌을 수 있기 때문입니다(회귀 검증)."),
  ...[
    "섹션 0: -99 완전 제거(수치·문자), WGRADE = A/B/C/D만, COST_AMT > 0, area 0 없음(버그 수정 확인)",
    "섹션 1: 행 2,408,699 · 열 44 · ID 무결성",
    "섹션 2: 전처리 안 한 컬럼(WEIGHT·AGE·sido·등급 등)이 원본과 완전 일치 — “건드리지 않은 것은 정말 안 건드렸나”",
    "섹션 3: 모든 결측 수가 ‘원본 -99 + 미등록 + 병합 부재’의 합으로 정확히 역산됨",
    "섹션 4: area 4개 컬럼 결측을 행 단위로 원인 분해 (농장 없음 vs 원본 -99)",
    "섹션 5~9: 타입·날짜 무결성·결측 현황·값 범위(WEIGHT>0 등)·등급 16종",
  ].map((t) => NUM("s7")(t)),
  H2("7.2 결과"),
  P([{ t: "전 항목 PASS — 2,408,699 × 44의 step6_preprocess.csv 확정.", b: true }, " 이 파일이 EDA(13~17회차)와 모델링의 공식 출발점입니다."]),
  PAGEBREAK()
);

// ===== 8장(통합): 최종 결과 =====
content.push(
  H1("8. 최종 데이터 결과"),
  H2("8.1 최종 테이블 구성 (44개 변수)"),
  tbl(
    ["출처", "컬럼", "개수"],
    [
      ["train 원본", "지역 3, 관측소, 날짜 3, 성별, 체격 2(WEIGHT·AGE), 육질 8, WGRADE, 가격, ID 2, 등급", "23"],
      ["death 병합", "death_count", "1"],
      ["lineage 병합", "KPN_NO, 부모 2, 조부모 4", "7"],
      ["area 병합", "C2023, C2024, C2025, AREA", "4"],
      ["weather 병합", "days_total, days_등급별 5, 생애 평균 3 (강수·풍속·최저기온)", "9"],
    ],
    [1800, 5726, 1500]
  ),
  P([{ t: "" }]),
  H2("8.2 결측 현황"),
  ...figure("00_merge_missing.png", "최종 데이터 결측 현황 (0.5% 이상 컬럼)"),
  tbl(
    ["컬럼군", "결측률", "출처"],
    [
      ["혈통 7개", "38.3~38.5%", "미등록 개체 922,653 (+KPN “-99” 3,570)"],
      ["COST_AMT", "36.7%", "원본 -99 (낙찰가 비공개) + 0원 1건"],
      ["AREA / C2023~25", "35.2% / 2.9~6.3%", "area에 없는 농장 + 원본 -99"],
      ["death_count", "21.5%", "폐사 기록 없는 농장"],
      ["육질 8개", "0.2%", "등외 등급 (판정 미실시)"],
    ],
    [2300, 2300, 4426]
  ),
  P([{ t: "" }]),
  H2("8.3 타깃(LAST_GRADE) 분포"),
  ...figure("00_merge_grade_dist.png", "16개 등급 분포 — 1++B 최다 13.3%, 등외 최소 0.23%"),
  BULLET("성별: 거세 1,222,465 / 암 1,174,414 / 수 11,820 — 수소는 0.5%뿐."),
  BULLET("WEIGHT 평균 423kg (범위 3~881), AGE 중앙값 33개월 (범위 1~344), 생애 사육일수 중앙값 993일."),
  BULLET([{ t: "클래스 불균형 58배", b: true }, " (1++B 319,588 vs 등외 5,540) — 분류 모델 설계의 핵심 제약."]),
  PAGEBREAK()
);

// ===== 9장: 해석 =====
content.push(
  H1("9. 결과 해석 (읽는 방법과 함께)"),
  box("결측을 해석하는 방법 — 출처가 3가지다", [
    ["① 원본 결측 코드(-99): 측정·공개를 안 한 값 (COST_AMT 37%, 육질 등외)"],
    ["② 미등록: 상대 테이블에 아예 없는 대상 (혈통 38%, death 21%)"],
    ["③ 가공 산물: 처리 과정에서 생긴 것 (area의 0 버그 — 발견 즉시 수정)"],
    ["같은 NaN이라도 출처가 다르면 의미와 대처가 다릅니다. ①②는 정보(“기록 없음” 자체가 특징)일 수 있어 채우지 않고 NaN 유지, ③은 버그이므로 반드시 수정."],
  ]),
  BULLET([{ t: "왜 결측을 채우지 않았나: ", b: true }, "주력 모델 LightGBM은 NaN을 자체 분기로 처리하며, 섣부른 평균 대치는 분포를 왜곡합니다(회귀용 대치는 20회차에서 별도). “모르는 것은 모른다고 두는 것”이 가장 정직한 전처리입니다."]),
  BULLET([{ t: "COST_AMT 37% 결측의 함의: ", b: true }, "가격 분석(트랙2)은 “가격이 공개된 소들만의 이야기”입니다. 결측이 무작위가 아닐 가능성(특정 유통 경로)을 보고서마다 한계로 명시합니다."]),
  BULLET([{ t: "클래스 불균형 58배의 함의: ", b: true }, "그냥 학습하면 모델이 다수 등급만 찍어도 정확도가 높게 나옵니다. 18회차에서 class_weight(소수 클래스 가중)·Stratified 분할로 대응하고, 평가는 정확도 대신 macro-F1 계열을 검토합니다."]),
  BULLET([{ t: "THI 위험일 2.1%의 함의: ", b: true }, "‘위험’ 단독으로는 신호가 희박해 15회차에서 주의+경고+위험을 합친 ratio_고온을 주력으로 삼은 근거가 됩니다."]),
  PAGEBREAK()
);

// ===== 10장: 다음 분석 반영 =====
content.push(
  H1("10. 다음 분석(EDA·분류·회귀)에 어떻게 이어지나"),
  tbl(
    ["이 단계의 산출", "어디서 어떻게 쓰이나"],
    [
      ["step6_preprocess.csv (240만 × 44)", "EDA 13~17회차의 입력 → step11_features.csv(92변수)로 확장"],
      ["days_* 생애 THI 일수 9종", "15회차 더위 비율(ratio_*), 17회차 heat_high·상호작용의 원료"],
      ["death_count, C2023~25, AREA", "17회차 density(사육밀도)·death_rate(폐사율)·farm_size의 원료"],
      ["혈통 7컬럼", "17회차 빈도 인코딩, 18회차 Target Encoding 후보"],
      ["FARM_UNIQUE_NO 무결성", "18회차 StratifiedGroupKFold(농장 단위 분할)의 그룹 키"],
      ["클래스 불균형 58배 확인", "18회차 class_weight, 19회차 LightGBM 평가지표 설계"],
      ["결측 출처 3분류 문서화", "20회차 회귀용 결측 대치 전략의 기준"],
      ["BIRTH_YMD ≤ ABATT_DATE 보장", "17회차 fattening_days·daily_gain이 음수 없이 계산되는 전제"],
    ],
    [3800, 5226]
  ),
  P([{ t: "" }]),
  P([
    { t: "한 가지 유의: ", b: true },
    "여기서 한 모든 병합·파생은 test 데이터에도 똑같이 적용해야 합니다(담당자 확인: test에 외부·제공 데이터 병합 허용). 같은 코드를 test에 재사용할 수 있도록 단계를 분리해 둔 것도 이 때문입니다.",
  ]),
  PAGEBREAK()
);

// ===== 11장: 총정리 =====
content.push(
  H1("11. 총정리"),
  box("병합·전처리를 세 문장으로", [
    ["① 흩어진 5종(+기상 3종) 원본을 키 단위를 맞춰 LEFT JOIN으로 합쳐, 행 하나 잃지 않고 240만 × 44 테이블을 만들었다."],
    ["② 기상은 30년 × 444관측소로 보강해 이상치 제거·보간·THI 계산을 거쳐, 누적합 알고리즘으로 소 한 마리당 생애 더위 노출을 뺄셈 한 번에 집계했다."],
    ["③ 모든 결측·이상값은 원본까지 역추적해 출처를 밝힌 뒤 처리했고(숨은 0 버그 발견·수정 포함), 9개 섹션 검증 전 항목 PASS로 확정했다."],
  ], "FFF3E0"),
  P([{ t: "" }]),
  H2("단계별 한 줄 요약"),
  tbl(
    ["단계", "한 줄 요약"],
    [
      ["병합 1~3 (폐사·혈통·농장)", "단위를 농장/개체로 맞춰 집계 후 LEFT JOIN — 행 수 보존 확인"],
      ["병합 4 (기상)", "4개 파일 통합 → 품질 관리 → THI → 누적합으로 생애 노출 집계"],
      ["병합 5 (검증)", "결측 수 역산 대조 + 샘플 원본 대조로 병합 정합성 증명"],
      ["전처리 1~5 (점검)", "컬럼별 4단계 루틴 점검 — 숨은 버그 3건·더미 해시 발견 처리"],
      ["전처리 6 (최종 검증)", "9개 섹션 전 항목 PASS → step6_preprocess.csv 확정"],
    ],
    [3000, 6026]
  ),
  P([{ t: "" }]),
  H2("이 단계가 남긴 원칙"),
  ...[
    "행 수 보존 — 모든 병합 후 2,408,699를 확인한다 (LEFT JOIN의 약속).",
    "결측은 출처별로 — 원본 코드(-99) / 미등록 / 가공 산물을 구분하고, 함부로 채우지 않는다.",
    "이상값은 원본까지 추적 — 지우기 전에 원인을 밝힌다 (진짜 이상고온은 살리고, 센서 오류만 지운다).",
    "검증은 역산으로 — 결측 수·샘플 값이 원본에서 재계산한 값과 일치해야 한다.",
    "모든 판정은 코드로 — check(조건) PASS/FAIL이 사람 눈을 대신한다.",
  ].map((t) => NUM("p11")(t)),
  PAGEBREAK()
);

// ===== 부록 =====
content.push(
  H1("부록 A. 공식·알고리즘 모음"),
  tbl(
    ["이름", "공식/원리", "쓰임"],
    [
      ["THI (NRC 1971)", "(1.8·ta_max+32) − (0.55−0.0055·rhm_avg)(1.8·ta_max−26.8)", "더위 스트레스 지수 (4-3)"],
      ["선형 보간", "x(t) = x(a) + (x(b)−x(a))·(t−a)/(b−a), limit=3", "기온·습도·풍속 결측 (4-2)"],
      ["이동 중앙값 편차", "dev = x(t) − median(앞뒤 15일)", "센서 오류 탐지 (4-2)"],
      ["누적합 구간합", "Σ x(lo..hi) = S(hi) − S(lo)", "생애 노출 집계 (4-4)"],
      ["이진 탐색", "searchsorted: 정렬 배열에서 위치를 log₂(n)번 비교로 탐색", "출생·도축일 위치 (4-4)"],
      ["그룹 집계", "count(폐사) / sum(면적·두수) by 농장", "병합 1·3"],
    ],
    [2100, 4426, 2500]
  ),
  P([{ t: "" }]),
  H1("부록 B. 용어집"),
  tbl(
    ["용어", "쉬운 설명"],
    [
      ["LEFT JOIN", "왼쪽(기준) 테이블 행은 전부 유지, 짝 없으면 NaN — 행 수가 변하면 사고"],
      ["키 (key)", "두 테이블에서 “같은 대상”을 잇는 컬럼 (소 ID, 농장 ID, 관측소+날짜)"],
      ["1:1 / 1:N", "키 하나에 상대 행이 하나/여럿 — 1:N을 그대로 붙이면 행이 불어남 → 집계 먼저"],
      ["결측 코드", "“값 없음”을 -99 같은 숫자로 기록한 것 — NaN으로 바꾸지 않으면 통계 오염"],
      ["보간 (interpolation)", "빈 값을 앞뒤 값으로 추정해 메우는 것 (여기선 직선, 3일 한도)"],
      ["강건성 (robustness)", "이상치에 흔들리지 않는 성질 — 중앙값이 평균보다 강건"],
      ["누적합 (prefix sum)", "처음부터 k번째까지의 합을 미리 계산해 두는 것 — 구간합이 뺄셈 한 번"],
      ["더미/플레이스홀더 값", "“미상”을 나타내려고 넣은 가짜 값 (혈통의 수상한 해시)"],
      ["회귀 검증", "고친 뒤에 “전에 통과한 검사”를 다시 돌려 깨진 게 없는지 확인"],
      ["클래스 불균형", "타깃 등급별 표본 수 격차 (1++B 32만 vs 등외 0.55만 = 58배)"],
    ],
    [2700, 6326]
  ),
  P([{ t: "" }]),
  H1("부록 C. 산출물 목록"),
  tbl(
    ["종류", "파일", "비고"],
    [
      ["병합 중간", "data/processed/1_merge/step1~step5 (8개 CSV)", "step4는 4-1~4-4 세분"],
      ["전처리 중간", "data/processed/2_preprocess/step1~step5 (5개 CSV)", "출처별 점검"],
      ["최종 ★", "data/processed/2_preprocess/step6_preprocess.csv", "2,408,699 × 44 — EDA·모델링 공식 입력"],
      ["결과 그림", "figures/00_merge_*.png (3장)", "등급 분포·결측·THI"],
      ["본 보고서", "보고서_병합전처리_팀공유.docx", "팀 공유용"],
    ],
    [1800, 4426, 2800]
  )
);

// ---------- 조립 ----------
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
      ...["s3", "s42", "s5", "s7", "p11"].map((ref) => ({
        reference: ref,
        levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 600, hanging: 300 } } } }],
      })),
    ],
  },
  sections: [
    {
      properties: {
        page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } },
      },
      headers: {
        default: new Header({
          children: [
            new Paragraph({
              alignment: AlignmentType.RIGHT,
              children: [new TextRun({ text: "한우 데이터 병합·전처리 보고서", size: 16, color: "888888" })],
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
  fs.writeFileSync(path.join(__dirname, "보고서_병합전처리_팀공유.docx"), buffer);
  console.log("저장 완료: 보고서_병합전처리_팀공유.docx", buffer.length, "bytes");
});
