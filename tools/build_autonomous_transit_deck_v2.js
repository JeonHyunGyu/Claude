// 스마트시티 자율주행 대중교통 정책의 국내외 사례 비교 및 발전방향 — v2 (절제된 학술 발표 스타일)
// 디자인 원칙: 흰 배경, 남색·진회색, 카드/아이콘/장식 요소 없음, 개조식 본문, 보고서형 3선 표, 그림 캡션, 하단 출처.
// 그래프로 제시한 수치는 표에 중복 기재하지 않음. 모든 수치·사실은 보고서(2026.11.05., 전현규) 원문 기준.
// 실행: node tools/build_autonomous_transit_deck_v2.js <out.pptx>
const path = require("path");
const pptxgen = require("pptxgenjs");
const { applyTheme } = require(process.env.APPLY_THEME ||
  "/root/.claude/skills/synced/2709c09c-c0dc-4f81-b385-fcb3527e2d18_99de5be8-b623-4241-b2e1-4194ab58e11a/pptx/scripts/apply_theme.js");

const OUT = process.argv[2] || "deck.pptx";
const FIG1 = path.join(__dirname, "fig1_seoul_a160.jpg");      // 보고서 [그림 1], 760×509
const FIG2 = path.join(__dirname, "fig2_punggol_shuttle.jpg"); // 보고서 [그림 2], 1037×496

const THEME = {
  name: "Academic Navy",
  headFontFace: "Malgun Gothic",
  bodyFontFace: "Malgun Gothic",
  colors: {
    dk1: "262626", lt1: "FFFFFF", dk2: "1F2A5C", lt2: "F2F2F2",
    accent1: "1F2A5C", accent2: "4F5D75", accent3: "A6A6A6",
    accent4: "8497B0", accent5: "7F7F7F", accent6: "D9D9D9",
    hlink: "1F4E9C", folHlink: "4F5D75",
  },
};
const HEX = THEME.colors;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "스마트시티 자율주행 대중교통 정책의 국내외 사례 비교 및 발전방향";
pres.author = "전현규";
const C = pres.SchemeColor;

const L = 0.6, CW = 12.13, R = L + CW;
const BODY_Y = 1.45;

// ---------- Layouts ----------
pres.defineSlideMaster({
  title: "COVER",
  background: { color: C.background1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: L, y: 2.0, w: 11.5, h: 1.5,
      fontSize: 34, bold: true, color: C.text2, valign: "bottom", align: "left", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: C.background1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: L, y: 0.5, w: CW, h: 0.6,
      fontSize: 26, bold: true, color: C.text2, valign: "middle", align: "left", margin: 0 }, text: "" } },
  ],
  slideNumber: { x: 12.33, y: 7.02, w: 0.4, h: 0.28, fontSize: 10, color: C.accent5, align: "right" },
});

// ---------- Helpers ----------
let uid = 0;
const nm = (s) => `${s}-${++uid}`;
const tx = (s, text, o) => { const { name, ...rest } = o; s.addText(text, { margin: 0, isTextBox: true, objectName: nm(name || "text"), ...rest }); };

function section(s, text) {
  tx(s, text, { x: L, y: 0.22, w: 8, h: 0.28, fontSize: 12, color: C.accent5, name: "section" });
}
function source(s, text) {
  tx(s, text, { x: L, y: 6.98, w: 11.4, h: 0.32, fontSize: 10, color: C.accent5, valign: "middle", name: "source" });
}
function caption(s, text, x, y, w, align = "left") {
  tx(s, text, { x, y, w, h: 0.32, fontSize: 13, bold: true, color: C.text1, align, valign: "middle", name: "caption" });
}
function note(s, text, x, y, w, h = 0.3, align = "left") {
  tx(s, text, { x, y, w, h, fontSize: 11, color: C.accent5, align, valign: "top", name: "note" });
}
// 개조식 본문: [level, text] 배열. level 0 = ■ 소제목, 1 = – 항목, 2 = · 보충(근거)
function outline(items, size = 16) {
  return items.map(([lv, t], i) => {
    const last = i === items.length - 1;
    const base = { breakLine: !last };
    if (lv === 0) return { text: t, options: { ...base, bold: true, color: C.text2, fontSize: size + 1,
      bullet: { characterCode: "25A0", indent: 20 }, paraSpaceBefore: i === 0 ? 0 : 12, paraSpaceAfter: 4 } };
    if (lv === 1) return { text: t, options: { ...base, color: C.text1, fontSize: size,
      bullet: { characterCode: "2013", indent: 18 }, indentLevel: 1, paraSpaceAfter: 4 } };
    return { text: t, options: { ...base, color: C.accent5, fontSize: size - 2,
      bullet: { characterCode: "00B7", indent: 16 }, indentLevel: 2, paraSpaceAfter: 4 } };
  });
}
function body(s, items, x, y, w, h, opt = {}) {
  tx(s, outline(items, opt.size || 16), { x, y, w, h, valign: "top", lineSpacingMultiple: opt.ls || 1.15, name: opt.name || "body" });
}
// 보고서형 3선 표: 상·하단 굵은 남색 선, 헤더 하단 선, 본문 행 사이 연회색 선, 세로선 없음
function academicTable(s, header, rowsIn, opt) {
  const none = { type: "none" };
  const heavy = { type: "solid", pt: 1.5, color: HEX.dk2 };
  const mid = { type: "solid", pt: 0.75, color: HEX.dk2 };
  const thin = { type: "solid", pt: 0.5, color: HEX.accent6 };
  const rows = [header.map((h, j) => ({ text: h, options: { bold: true, color: C.text2, fill: { color: C.background2 },
    align: (opt.headerAlign && opt.headerAlign[j]) || "center", border: [heavy, none, mid, none] } }))];
  rowsIn.forEach((r, i) => {
    const lastRow = i === rowsIn.length - 1;
    rows.push(r.map((c, j) => ({ text: c, options: { color: C.text1, bold: j === 0 && opt.boldFirstCol,
      align: (opt.align && opt.align[j]) || "left", border: [none, none, lastRow ? heavy : thin, none] } })));
  });
  s.addTable(rows, { x: opt.x, y: opt.y, w: opt.w, colW: opt.colW, rowH: opt.rowH, fontSize: opt.fontSize || 14, valign: "middle",
    margin: [0.03, 0.1, 0.03, 0.1], objectName: nm(opt.name || "table") });
}
// 단순 절차도: 테두리만 있는 사각형 + 순서 화살표
function flow(s, steps, x, y, w, h, gap = 0.45) {
  const bw = (w - gap * (steps.length - 1)) / steps.length;
  steps.forEach((t, i) => {
    const bx = x + i * (bw + gap);
    tx(s, t, { x: bx, y, w: bw, h, shape: pres.shapes.RECTANGLE, fill: { color: C.background1 }, line: { color: C.text2, width: 1 },
      fontSize: 14, bold: true, color: C.text2, align: "center", valign: "middle", margin: 0.05, name: "step" });
    if (i < steps.length - 1) {
      s.addShape(pres.shapes.LINE, { x: bx + bw + 0.08, y: y + h / 2, w: gap - 0.16, h: 0,
        line: { color: HEX.accent5, width: 1.25, endArrowType: "triangle" }, objectName: nm("flow-arrow") });
    }
  });
  return bw;
}

// ======================= 1. 표지 =======================
function slideCover() {
  const s = pres.addSlide({ masterName: "COVER" });
  s.addText("스마트시티 자율주행 대중교통 정책의\n국내외 사례 비교 및 발전방향", { placeholder: "title" });
  tx(s, "서울 새벽 동행 자율주행 버스와 싱가포르 Punggol 자율주행 셔틀을 중심으로", { x: L, y: 3.7, w: 11.5, h: 0.45, fontSize: 18, color: C.accent2, name: "subtitle" });
  tx(s, [{ text: "2026. 11. 5.(목)", options: { breakLine: true } }, { text: "전현규", options: { bold: true, color: C.text1 } }],
    { x: L, y: 5.5, w: 5, h: 0.8, fontSize: 16, color: C.accent2, paraSpaceAfter: 4, name: "date-author" });
  s.addNotes("스마트시티 자율주행 대중교통 정책의 국내외 사례 비교 및 발전방향에 대해 발표하겠습니다. 국내 사례로 서울 새벽 동행 자율주행 버스, 해외 사례로 싱가포르 Punggol 자율주행 셔틀을 다룹니다.");
}

// ======================= 2. 발표 순서 =======================
function slideAgenda() {
  const s = pres.addSlide({ masterName: "CONTENT" });
  s.addText("발표 순서", { placeholder: "title" });
  const items = [
    ["Ⅰ", "연구 배경 및 목적", "자율주행 대중교통 도입 배경, 사례 선정, 발표 목적"],
    ["Ⅱ", "국내 사례", "서울 새벽 동행 자율주행 버스(A160)"],
    ["Ⅲ", "해외 사례", "싱가포르 Punggol 자율주행 셔틀"],
    ["Ⅳ", "국내외 사례 비교 및 시사점", "도입 목적, 운영방식, 성과평가 비교"],
    ["Ⅴ", "향후 발전 방향 및 기말 발표 계획", "교통 공백 유형별 운영방식, 기말 연구주제"],
  ];
  items.forEach(([n, t, d], i) => {
    const y = 1.75 + i * 0.95;
    tx(s, n, { x: L + 0.1, y, w: 0.6, h: 0.45, fontSize: 20, bold: true, color: C.text2, name: "agenda-num" });
    tx(s, t, { x: L + 0.8, y, w: 10, h: 0.45, fontSize: 20, bold: true, color: C.text1, name: "agenda-title" });
    tx(s, d, { x: L + 0.8, y: y + 0.45, w: 10, h: 0.35, fontSize: 14, color: C.accent5, name: "agenda-desc" });
  });
  s.addNotes("발표는 연구 배경과 목적, 국내 사례, 해외 사례, 국내외 비교와 시사점, 향후 발전 방향과 기말 발표 계획 순서로 진행합니다.");
}

// ======================= 3. 연구 배경 및 목적 =======================
function slideBackground() {
  const s = pres.addSlide({ masterName: "CONTENT" });
  section(s, "Ⅰ. 연구 배경 및 목적");
  s.addText("연구 배경 및 목적", { placeholder: "title" });
  body(s, [
    [0, "자율주행 대중교통 도입 배경"],
    [1, "대중교통 공급이 부족한 시간대와 기존 교통망의 접근성이 낮은 지역의 이동 불편 해소 필요성 증대"],
    [1, "기존 대중교통체계의 한계를 보완하기 위한 자율주행 버스 및 셔틀 서비스 도입 추진"],
    [0, "국내외 정책 사례 선정"],
    [1, "국내: 서울 새벽 동행 자율주행 버스 — 새벽 첫차 혼잡 및 운전기사 수급 문제 대응 (서울특별시, 2024)"],
    [1, "해외: 싱가포르 Punggol 자율주행 셔틀 — 주거지역과 대중교통 거점 간 First/Last Mile 연결성 개선 (LTA, 2025)"],
    [2, "두 사례는 각각 시간적·공간적 교통 공백에 대응하여 운영방식의 차이를 비교할 수 있음"],
    [0, "발표 목적"],
    [1, "국내외 자율주행 대중교통 정책의 도입 배경, 활용 기술 및 운영방식, 성과평가 방법 비교"],
    [1, "사례 분석을 통한 타 도시 적용조건 및 향후 발전 방향 검토"],
  ], L, BODY_Y + 0.05, CW, 5.3);
  source(s, "출처: 보고서 p.1 (주제 선정 배경 및 목적)");
  s.addNotes("도시에서는 대중교통 공급이 부족한 시간대와 접근성이 낮은 지역에서 이동 불편이 발생하고, 이를 보완하기 위해 자율주행 버스와 셔틀 도입이 추진되고 있습니다. 국내는 새벽 첫차 혼잡과 운전기사 수급 문제에 대응한 서울 새벽 동행 자율주행 버스, 해외는 First/Last Mile 연결성 개선을 위한 싱가포르 Punggol 셔틀을 선정했습니다. 두 사례는 각각 시간적·공간적 교통 공백에 대응하므로 운영방식의 차이를 비교할 수 있습니다. 발표 목적은 두 사례의 도입 배경, 활용 기술 및 운영방식, 성과평가 방법을 비교하고 타 도시 적용조건과 발전 방향을 검토하는 것입니다.");
}

// ======================= 4. 서울 개요 =======================
function slideSeoulOverview() {
  const s = pres.addSlide({ masterName: "CONTENT" });
  section(s, "Ⅱ. 국내 사례");
  s.addText("서울 새벽 동행 자율주행 버스(A160) 개요", { placeholder: "title" });
  const lw = 7.1;
  body(s, [
    [0, "사업 개요"],
    [1, "최초 운행: 2024년 11월 26일"],
    [1, "운행구간: 도봉산역 광역환승센터 ~ 영등포역(편도 25.7km)"],
    [1, "운행 시간: 평일 오전 3시 30분 출발"],
    [2, "일반 시내버스 첫차보다 약 30분 조기 운행"],
    [1, "운영방식: 기존 160번 노선 기반, 주요 구간 87개 정류소 운행"],
    [0, "도입 배경"],
    [1, "환경미화원, 경비원 등 새벽 출근 근로자의 이동 불편 및 첫차 혼잡"],
    [2, "기존 160번 노선 새벽 첫차의 최대 재차 인원 50명 초과"],
    [1, "새벽 시간대 운전기사 수급의 어려움으로 인한 버스 증차 제약"],
    [1, "자율주행 기술로 교통서비스 공급 시간 확대 및 이용 편의 증진"],
  ], L, BODY_Y + 0.05, lw, 5.2, { ls: 1.2 });
  const fx = L + lw + 0.35, fw = R - fx, fh = fw * 509 / 760, fy = BODY_Y + 0.1;
  s.addImage({ path: FIG1, x: fx, y: fy, w: fw, h: fh, altText: "서울 새벽 동행 자율주행 버스 A160 (보고서 그림 1)", objectName: "fig1" });
  caption(s, "[그림 1] 서울 새벽 동행 자율주행 버스 A160", fx, fy + fh + 0.1, fw, "center");
  note(s, "자료: 서울특별시(2024)", fx, fy + fh + 0.42, fw, 0.28, "center");
  source(s, "출처: 보고서 p.1–2 (사업 개요, 도입 배경 및 목적, [그림 1])");
  s.addNotes("A160은 2024년 11월 26일 처음 운행했습니다. 도봉산역 광역환승센터에서 영등포역까지 편도 25.7km 구간을 평일 오전 3시 30분에 출발하며, 일반 시내버스 첫차보다 약 30분 이릅니다. 기존 160번 노선을 기반으로 주요 구간 87개 정류소를 운행합니다. 도입 배경은 새벽 출근 근로자의 이동 불편과 첫차 혼잡(기존 160번 첫차 최대 재차 인원 50명 초과), 그리고 새벽 시간대 운전기사 수급난으로 인한 증차 제약입니다.");
}

// ======================= 5. 서울 운영방식 및 성과 =======================
function slideSeoulResults() {
  const s = pres.addSlide({ masterName: "CONTENT" });
  section(s, "Ⅱ. 국내 사례");
  s.addText("A160 운영방식 및 운영성과", { placeholder: "title" });
  const lw = 6.4;
  body(s, [
    [0, "적용 기술 및 운영방식"],
    [1, "기존 시내버스 노선을 활용한 자율주행 버스 운행 및 단계적 상용화"],
    [2, "자율차 시범운행 지구 지정, 시험 운행 및 안전 검증 등 제도적·기술적 절차 추진"],
    [1, "승객 안전을 고려한 좌석제 운영(최대 22명 탑승, 입석 금지)"],
    [1, "BIT·모바일 서비스로 실시간 위치·도착·잔여 좌석 정보 제공"],
    [0, "운영성과 및 확대 현황"],
    [1, "통근 목적 이용자 비율이 높아 새벽 근로자 이동 지원이라는 도입 목적에 부합"],
    [1, "2026년 A741, A148, A504 노선 추가로 총 4개 노선으로 확대"],
  ], L, BODY_Y + 0.05, lw, 5.2, { ls: 1.2 });

  const tx0 = L + lw + 0.45, tw = R - tx0, ty = BODY_Y + 0.1;
  caption(s, "<표 1> 서울 새벽 동행 자율주행 버스 A160 주요 운영성과", tx0, ty, tw);
  academicTable(s, ["구분", "운영성과"], [
    ["누적 이용객", "약 27,600명"],
    ["자율주행 사고", "0건"],
    ["서비스 만족도", "평균 4.08 / 5점"],
    ["통근 목적 이용자 비율", "96.2%"],
    ["50대 이상 이용자 비율", "85.4%"],
  ], { x: tx0, y: ty + 0.4, w: tw, colW: [2.85, tw - 2.85], rowH: [0.44, 0.46, 0.46, 0.46, 0.46, 0.46], fontSize: 15,
    align: ["left", "center"], headerAlign: ["left", "center"], name: "a160-table" });
  note(s, "주: 운행 시작 이후 약 15개월간 실적(2026년 3월 발표 기준)", tx0, ty + 0.4 + 0.44 + 5 * 0.46 + 0.08, tw, 0.28);
  note(s, "자료: 서울특별시(2026.3.29.), 「새벽 동행 자율주행 버스 'A741' 첫 시동…새벽 출근길 더 빨라진다」를 바탕으로 작성",
    tx0, ty + 0.4 + 0.44 + 5 * 0.46 + 0.38, tw, 0.5);
  source(s, "출처: 보고서 p.2–3 (적용 기술 및 운영방식, 운영성과 및 확대 현황, <표 1>)");
  s.addNotes("A160은 기존 시내버스 노선을 활용해 단계적 상용화를 추진했고, 시범운행 지구 지정, 시험 운행, 안전 검증 등의 절차를 거쳤습니다. 최대 22명 좌석제로 운영하며 입석은 금지합니다. 운행 시작 이후 약 15개월간 누적 이용객 약 27,600명, 사고 0건, 만족도 평균 4.08점을 기록했습니다. 통근 목적 이용자가 96.2%로 도입 목적에 부합했고, 2026년에는 3개 노선이 추가되어 총 4개 노선으로 확대되었습니다.");
}

// ======================= 6. 싱가포르 개요 =======================
function slidePunggolOverview() {
  const s = pres.addSlide({ masterName: "CONTENT" });
  section(s, "Ⅲ. 해외 사례");
  s.addText("싱가포르 Punggol 자율주행 셔틀 개요", { placeholder: "title" });
  const lw = 6.6;
  body(s, [
    [0, "사업 개요"],
    [1, "사업명: Punggol Autonomous Shuttle Services"],
    [1, "시행기관: 싱가포르 육상교통청(LTA)"],
    [1, "운영기관: Grab(WeRide), ComfortDelGro(Pony.ai)"],
    [1, "운행 규모: 3개 고정노선, 노선별 4~7개 승하차 지점"],
    [2, "5~8인승 자율주행차량 도입"],
    [0, "운행 시간(주간 중심 운영)"],
    [1, "Route 1·3(Grab·WeRide): 평일 09:30~17:30"],
    [1, "Route 2(ComfortDelGro·Pony.ai): 화~일요일 09:30~17:00"],
  ], L, BODY_Y + 0.05, lw, 5.2, { ls: 1.2 });
  const fx = L + lw + 0.35, fw = R - fx, fh = fw * 496 / 1037, fy = BODY_Y + 0.1;
  s.addImage({ path: FIG2, x: fx, y: fy, w: fw, h: fh, altText: "싱가포르 Grab·WeRide의 자율주행 셔틀 차량 (보고서 그림 2)", objectName: "fig2" });
  caption(s, "[그림 2] 싱가포르 Grab·WeRide의 자율주행 셔틀 차량", fx, fy + fh + 0.1, fw, "center");
  note(s, "자료: Grab(2026.10.5. 갱신)", fx, fy + fh + 0.42, fw, 0.28, "center");
  source(s, "출처: 보고서 p.3 (사업 개요, [그림 2]), 운행 시간은 보고서 참고문헌의 ComfortDelGro 안내 포함");
  s.addNotes("Punggol 자율주행 셔틀은 싱가포르 육상교통청 LTA가 시행하고 Grab과 ComfortDelGro가 운영합니다. Grab은 WeRide, ComfortDelGro는 Pony.ai의 플랫폼을 씁니다. 3개 고정노선에 노선별 4~7개 승하차 지점이 있고 5~8인승 차량을 도입했습니다. 주간 중심 운영으로, Route 1·3은 평일 09:30~17:30, Route 2는 화~일요일 09:30~17:00 운행합니다.");
}

// ======================= 7. 싱가포르 도입 배경 및 기술 =======================
function slidePunggolBackground() {
  const s = pres.addSlide({ masterName: "CONTENT" });
  section(s, "Ⅲ. 해외 사례");
  s.addText("Punggol 셔틀 도입 배경 및 적용 기술", { placeholder: "title" });
  const cw = (CW - 0.5) / 2;
  body(s, [
    [0, "도입 배경 및 목적"],
    [1, "주거지역과 대중교통 거점·생활시설 간 First/Last Mile 연결성 개선"],
    [2, "의료시설, 시장, 쇼핑 시설, MRT역 등과의 연결성 강화"],
    [1, "기존 대중교통망 보완을 통한 지역 내 이동 편의성·접근성 향상"],
    [2, "기존 대중교통 대비 최대 15분 이동시간 단축을 목표로 설정"],
  ], L, BODY_Y + 0.05, cw, 2.9, { name: "body-left" });
  body(s, [
    [0, "적용 기술 및 운영방식"],
    [1, "자율주행 기술을 활용한 지역 내 고정노선 셔틀 서비스"],
    [1, "Grab은 WeRide, ComfortDelGro는 Pony.ai의 자율주행 플랫폼 적용"],
    [1, "정밀 노선 매핑, 센서 보정, AI 모델 검증, 현지 도로 환경 적응 수행"],
    [1, "이용자 의견과 운행 데이터를 수집하여 운영방식 개선 추진"],
  ], L + cw + 0.5, BODY_Y + 0.05, cw, 2.9, { name: "body-right" });

  const fy = 5.0;
  caption(s, "단계별 안전 검증 절차", L, fy - 0.45, CW);
  flow(s, ["폐쇄시험장(CETRAN) 검증", "실도로 무승객 시험", "주민 대상 무료 시승", "일반인 탑승"], L, fy, CW, 0.75, 0.5);
  note(s, "※ 차량 내 안전요원 탑승 및 실시간 운행 모니터링을 통한 안전관리 병행", L, fy + 0.95, CW, 0.3);
  source(s, "출처: 보고서 p.4 (도입 배경 및 목적, 적용 기술 및 운영방식)");
  s.addNotes("Punggol 셔틀은 주거지역과 대중교통 거점·생활시설 간 First/Last Mile 연결성 개선을 위해 도입되었고, 기존 대중교통 대비 최대 15분의 이동시간 단축을 목표로 설정했습니다. 이는 실적이 아닌 목표치입니다. 기술적으로는 WeRide와 Pony.ai 플랫폼을 적용해 정밀 노선 매핑, 센서 보정, AI 모델 검증을 수행했습니다. 안전 검증은 폐쇄시험장 CETRAN 검증, 실도로 무승객 시험, 주민 대상 무료 시승, 일반인 탑승 순으로 진행했고, 안전요원 탑승과 실시간 모니터링을 병행합니다.");
}

// ======================= 8. 싱가포르 운영성과 =======================
function slidePunggolResults() {
  const s = pres.addSlide({ masterName: "CONTENT" });
  section(s, "Ⅲ. 해외 사례");
  s.addText("Punggol 셔틀 운영성과 및 이용자 의견", { placeholder: "title" });

  // 좌: 운영 실적(그래프와 중복되지 않는 항목만) + 주요 결과
  const lw = 5.5, ty = BODY_Y + 0.05;
  caption(s, "<표 2> 주요 운영성과", L, ty, lw);
  academicTable(s, ["구분", "운영성과"], [
    ["고유 이용자 수", "11,500명 이상"],
    ["이용자 설문 응답", "약 900명"],
  ], { x: L, y: ty + 0.4, w: lw, colW: [2.9, lw - 2.9], rowH: [0.44, 0.48, 0.48], fontSize: 15,
    align: ["left", "center"], headerAlign: ["left", "center"], name: "punggol-table" });
  note(s, "주: 2026년 7월 12일 기준, 수요응답형 확대 이전 고정노선 단계의 실적\n자료: LTA(2026.7.29.)를 바탕으로 작성 (설문 비율 항목은 [그림 3])",
    L, ty + 0.4 + 0.44 + 2 * 0.48 + 0.08, lw, 0.55);
  body(s, [
    [0, "주요 결과"],
    [1, "설문 응답자 99%가 탑승 중 안전하다고 느끼고 추천 의향을 나타냄"],
    [1, "고정노선·고정 승하차 지점 운영에 대한 이용자 요구 확인"],
    [1, "이용자 의견을 반영하여 2026년 9월부터 수요응답형 서비스 병행"],
  ], L, 3.95, lw, 2.8, { name: "summary" });

  // 우: 설문 응답 비율 막대그래프
  const gx = L + lw + 0.6, gw = R - gx;
  caption(s, "[그림 3] 탑승 후 설문 응답 비율 (응답자 약 900명)", gx, ty, gw);
  s.addChart(pres.charts.BAR, [{
    name: "응답 비율",
    labels: ["직접적인 이동 경로 선호", "승하차 지점 선택 선호", "체감 안전성 및 추천 의향"],
    values: [40, 60, 99],
  }], {
    x: gx, y: ty + 0.45, w: gw, h: 4.1, barDir: "bar",
    chartColors: [HEX.dk2], barGapWidthPct: 80,
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: '0"%"', dataLabelFontSize: 14, dataLabelColor: HEX.dk1, dataLabelFontFace: "+mn-lt",
    valAxisMinVal: 0, valAxisMaxVal: 100, valAxisMajorUnit: 20, valAxisLabelFormatCode: '0"%"',
    valAxisLabelFontSize: 12, valAxisLabelColor: HEX.accent5, valAxisLabelFontFace: "+mn-lt",
    catAxisLabelFontSize: 14, catAxisLabelColor: HEX.dk1, catAxisLabelFontFace: "+mn-lt",
    valGridLine: { color: HEX.accent6, size: 0.5 }, catGridLine: { style: "none" },
    catAxisLineShow: true, valAxisLineShow: false, showLegend: false, objectName: "survey-chart",
  });
  note(s, "주: 원문 <표 2>의 비율 항목으로 발표자 작성. 60%, 40%는 원문의 '약' 수치이며 두 문항의 중복 응답 여부는 원문에 명시되지 않음",
    gx, ty + 4.65, gw, 0.5);
  source(s, "출처: 보고서 p.4–5 (운영성과 및 서비스 개선, <표 2>)");
  s.addNotes("2026년 7월 12일 기준 고유 이용자는 11,500명 이상이고, 탑승 후 설문에는 약 900명이 응답했습니다. 그래프와 같이 응답자의 99%가 안전하다고 느끼고 추천 의향을 나타냈습니다. 또한 약 60%는 승하차 지점 선택을, 약 40%는 목적지까지 직접적인 경로를 선호해 고정노선·고정 승하차 지점 운영에 대한 요구가 확인되었습니다. 이를 반영해 2026년 9월부터 수요응답형 서비스를 병행하고 있습니다. 이 실적은 수요응답형 확대 이전 고정노선 단계의 것입니다.");
}

// ======================= 9. 싱가포르 서비스 개선 =======================
function slidePunggolOnDemand() {
  const s = pres.addSlide({ masterName: "CONTENT" });
  section(s, "Ⅲ. 해외 사례");
  s.addText("서비스 개선: 고정노선과 수요응답형 병행", { placeholder: "title" });
  const lw = 5.2, ty = BODY_Y + 0.1;
  caption(s, "Punggol 자율주행 셔틀 추진 경과", L, ty, lw);
  academicTable(s, ["시점", "내용"], [
    ["2025. 9. 20.", "도입계획 발표"],
    ["2026. 4. 1.", "일반인 무료 탑승 개시\n(Grab Route 1·3 기준)"],
    ["2026. 9.", "수요응답형 서비스 병행"],
  ], { x: L, y: ty + 0.4, w: lw, colW: [1.7, lw - 1.7], rowH: [0.46, 0.65, 0.9, 0.65], fontSize: 16,
    align: ["center", "left"], headerAlign: ["center", "left"], name: "timeline-table" });
  note(s, "자료: 보고서 p.3–4 내용을 발표자 정리", L, ty + 0.4 + 0.46 + 0.65 + 0.9 + 0.65 + 0.1, lw, 0.28);

  body(s, [
    [0, "수요응답형(On-Demand) 서비스 도입"],
    [1, "이용자 의견을 반영하여 기존 고정노선과 병행 운영"],
    [1, "이용자가 지정된 승하차 지점 사이를 예약해 직접 이동하는 방식"],
    [1, "Grab의 기존 Route 1·3 승하차 지점 활용"],
    [1, "주민 수요와 이용자 의견을 반영하여 승하차 지점 단계적 확대 추진"],
  ], L + lw + 0.6, ty, R - (L + lw + 0.6), 4.5, { ls: 1.35, size: 17 });
  source(s, "출처: 보고서 p.3–4 (사업 개요, 운영성과 및 서비스 개선)");
  s.addNotes("Punggol 셔틀은 2025년 9월 20일 도입계획 발표 후 2026년 4월 1일 Grab Route 1·3 기준 일반인 무료 탑승을 시작했습니다. 이후 이용자 의견을 반영해 2026년 9월부터 기존 고정노선과 병행하는 수요응답형 서비스로 확대했습니다. 이용자가 지정된 승하차 지점 사이를 예약해 직접 이동하는 방식이며, Grab의 기존 Route 1·3 승하차 지점을 활용하고 승하차 지점을 단계적으로 확대할 계획입니다.");
}

// ======================= 10. 국내외 비교표 (축약) =======================
function slideCompareTable() {
  const s = pres.addSlide({ masterName: "CONTENT" });
  section(s, "Ⅳ. 국내외 사례 비교 및 시사점");
  s.addText("국내외 자율주행 대중교통 정책 사례 비교", { placeholder: "title" });
  body(s, [
    [1, "공통점: 기존 대중교통체계의 한계 보완   /   차이점: 지역별 교통 문제에 따른 도입 목적, 차량 형태, 운영방식"],
  ], L, BODY_Y, CW, 0.4, { name: "lead" });

  const ty = 1.95, rh = 0.45;
  caption(s, "<표 3> 국내외 자율주행 대중교통 정책 사례 비교", L, ty, CW);
  academicTable(s,
    ["구분", "서울 A160", "싱가포르 Punggol"],
    [
      ["대응 문제", "새벽 교통 공백, 첫차 혼잡", "First/Last Mile 연결성 부족"],
      ["도입 목적", "새벽 근로자 이동 지원, 공급 확대", "지역 내 이동 편의성·접근성 개선"],
      ["차량 형태", "대형 자율주행버스", "소형 자율주행 셔틀"],
      ["운영 방식", "기존 노선 기반 조기 운행", "고정노선 → 수요응답형 확대"],
      ["주요 이용자", "새벽 출근 근로자", "지역 주민, 생활시설·교통거점 이용자"],
      ["성과평가 지표", "이용객 수, 만족도, 이용목적, 사고 여부", "고유 이용자 수, 체감 안전성, 추천 의향, 이용자 의견"],
      ["주요 성과", "누적 약 27,600명, 만족도 4.08/5점", "고유 이용자 11,500명 이상, 응답자 99% 긍정 평가"],
      ["향후 방향", "새벽 혼잡노선 확대, 운영체계 개선", "수요응답형 운행, 승하차 지점 확대"],
    ],
    { x: L, y: ty + 0.4, w: CW, colW: [2.1, 4.6, CW - 6.7], rowH: [0.44, ...Array(8).fill(rh)], fontSize: 15,
      align: ["center", "left", "left"], boldFirstCol: true, name: "compare-table" });
  note(s, "주: 원문 <표 3>을 발표용으로 축약. 자료: 서울특별시(2026.3.29.), LTA(2025.9.20., 2026.7.29.), Grab(2026.10.5.)을 바탕으로 작성",
    L, 6.6, CW, 0.28);
  source(s, "출처: 보고서 p.5 (국내외 자율주행 대중교통 정책 비교, <표 3>)");
  s.addNotes("두 사례는 자율주행 기술로 기존 대중교통체계의 한계를 보완한다는 공통점이 있지만, 지역별 교통 문제에 따라 도입 목적, 차량 형태, 운영방식이 다릅니다. 서울은 대형 자율주행버스로 기존 노선을 기반으로 조기 운행하고, 싱가포르는 소형 셔틀로 고정노선을 도입한 뒤 수요응답형으로 확대했습니다. 성과평가도 서울은 이용객 수·만족도·이용목적·사고 여부를, 싱가포르는 고유 이용자 수·체감 안전성·추천 의향·이용자 의견을 활용했습니다.");
}

// ======================= 11. 시사점 =======================
function slideImplications() {
  const s = pres.addSlide({ masterName: "CONTENT" });
  section(s, "Ⅳ. 국내외 사례 비교 및 시사점");
  s.addText("국내외 사례 비교를 통한 시사점", { placeholder: "title" });
  body(s, [
    [0, "지역별 교통 문제에 따른 운영방식 차별화 필요"],
    [1, "서울: 새벽시간대 대중교통 공급 부족에 대응하여 기존 시내버스 노선 활용"],
    [1, "싱가포르: 거점–생활시설 연결성 개선을 위해 소형 셔틀 도입 후 수요응답형으로 확대"],
    [2, "지역별 교통수요와 서비스 목적을 고려한 운영방식 선정 필요"],
    [0, "이용자 중심의 서비스 개선 및 기존 대중교통과의 연계"],
    [1, "싱가포르: 이용자 의견을 반영하여 고정노선 중심의 운영방식 개선"],
    [2, "기존 교통수단과의 연계, 접근성 및 서비스 편의성을 종합적으로 고려할 필요"],
    [0, "운영성과와 이용자 수용성을 종합적으로 고려한 평가"],
    [1, "두 사례 모두 이용실적과 함께 만족도, 체감 안전성 등 이용자 평가 활용"],
    [2, "이동시간, 대기시간, 안전성, 이용자 수용성 등의 지표를 종합적으로 검토할 필요"],
  ], L, BODY_Y + 0.05, CW, 5.35, { ls: 1.1 });
  source(s, "출처: 보고서 p.5–6 (국내외 사례 비교를 통한 시사점)");
  s.addNotes("시사점은 세 가지입니다. 첫째, 서울과 싱가포르는 교통 문제에 따라 다른 운영방식을 택했으므로 지역별 교통수요와 서비스 목적을 고려한 운영방식 선정이 필요합니다. 둘째, 싱가포르처럼 이용자 의견을 반영한 개선과 기존 교통수단과의 연계, 접근성, 편의성을 함께 고려해야 합니다. 셋째, 두 사례 모두 이용실적과 함께 이용자 평가를 활용했으므로, 향후에는 이동시간, 대기시간, 안전성, 이용자 수용성 지표를 종합적으로 검토할 필요가 있습니다.");
}

// ======================= 12. 향후 발전 방향 =======================
function slideDirection() {
  const s = pres.addSlide({ masterName: "CONTENT" });
  section(s, "Ⅴ. 향후 발전 방향 및 기말 발표 계획");
  s.addText("향후 발전 방향: 교통 공백 특성에 따른 운영방식 선정", { placeholder: "title" });
  body(s, [
    [1, "교통 공백의 발생 특성에 따라 자율주행 대중교통의 운영방식이 달라질 수 있음을 확인"],
  ], L, BODY_Y, CW, 0.45, { name: "lead" });

  const ty = 2.2;
  caption(s, "교통 공백 유형별 자율주행 대중교통 운영방식", L, ty, CW);
  academicTable(s, ["교통 공백 유형", "정의", "사례", "운영방식"], [
    ["시간적 교통 공백", "새벽·심야 등 대중교통 공급이 부족한 시간대", "서울 A160", "새벽 시간대 교통수요 대응을 위한 기존 노선 기반 운행"],
    ["공간적 교통 공백", "대중교통 접근성이 낮거나 교통서비스 공급이 부족한 지역", "싱가포르 Punggol", "지역 내 접근성 개선을 위한 셔틀 운행 및 수요응답형 서비스 확대"],
  ], { x: L, y: ty + 0.4, w: CW, colW: [2.2, 3.7, 2.0, CW - 7.9], rowH: [0.44, 0.85, 0.85], fontSize: 15,
    align: ["center", "left", "center", "left"], headerAlign: ["center", "left", "center", "left"], boldFirstCol: true, name: "gap-table" });
  note(s, "자료: 보고서 p.6 (교통 공백 유형 정의, 사례별 운영방식)을 발표자 정리", L, ty + 0.4 + 0.44 + 1.7 + 0.08, CW, 0.28);

  body(s, [
    [0, "시사점"],
    [1, "지역별 교통수요 및 도로·교통환경을 고려한 적정 운영방식 선정 필요"],
  ], L, 5.4, CW, 1.1, { name: "conclusion" });
  source(s, "출처: 보고서 p.6 (교통 공백 특성에 따른 적정 운영방식 선정 필요)");
  s.addNotes("두 사례를 통해 교통 공백의 발생 특성에 따라 운영방식이 달라질 수 있음을 확인했습니다. 새벽·심야처럼 공급이 부족한 시간대의 시간적 공백에는 서울처럼 기존 노선 기반 운행이, 접근성이 낮은 지역의 공간적 공백에는 싱가포르처럼 셔틀 운행과 수요응답형 서비스가 활용되었습니다. 따라서 지역별 교통수요와 도로·교통환경을 고려한 적정 운영방식 선정이 필요합니다.");
}

// ======================= 13. 기말 연구계획 =======================
function slidePlan() {
  const s = pres.addSlide({ masterName: "CONTENT" });
  section(s, "Ⅴ. 향후 발전 방향 및 기말 발표 계획");
  s.addText("기말 발표 연구주제 및 수행 계획", { placeholder: "title" });
  body(s, [
    [0, "연구주제"],
    [1, "교통 공백 해소를 위한 자율주행 대중교통의 고정노선형·수요응답형 운영방식 비교 및 적용방안 연구"],
    [2, "운영방식별 특성 및 적용조건 비교, 교통수요 특성에 따른 적정 운영방안 제안"],
  ], L, BODY_Y + 0.05, CW, 1.5, { name: "topic" });

  const fy = 3.85, fh = 0.85;
  caption(s, "수행 절차", L, fy - 0.45, CW);
  const steps = ["① 선행연구 검토", "② 운영 특성 비교", "③ 적용조건 검토", "④ 적용방안 제시"];
  const gap = 0.45;
  const bw = flow(s, steps, L, fy, CW, fh, gap);
  const desc = [
    "교통 공백 유형(시간적·공간적) 및 자율주행 대중교통 운영방식 관련 선행연구",
    "고정노선형·수요응답형의 운영 특성 및 장단점\n(수요 집중·분산 정도, 대기시간, 이동 편의성, 운영 효율성)",
    "교통 공백 유형 및 수요 특성에 따른 운영방식별 적용조건",
    "적정 운영방식 선정 기준과 적용방안",
  ];
  desc.forEach((d, i) => tx(s, d, { x: L + i * (bw + gap), y: fy + fh + 0.22, w: bw, h: 1.9, fontSize: 15, lineSpacingMultiple: 1.15, color: C.text1, valign: "top", name: "step-desc" }));
  source(s, "출처: 보고서 p.6 (기말 발표 연구주제, 기말 발표 수행 계획)");
  s.addNotes("기말 발표에서는 교통 공백 해소를 위한 자율주행 대중교통의 고정노선형·수요응답형 운영방식 비교 및 적용방안을 연구합니다. 시간적·공간적 교통 공백과 운영방식에 관한 선행연구를 검토하고, 수요의 집중·분산 정도, 대기시간, 이동 편의성, 운영 효율성을 기준으로 두 운영방식을 비교합니다. 이어 공백 유형과 수요 특성에 따른 적용조건을 검토하고 적정 운영방식 선정 기준과 적용방안을 제시하겠습니다.");
}

// ======================= 14. 참고문헌 =======================
function slideReferences() {
  const s = pres.addSlide({ masterName: "CONTENT" });
  s.addText("참고문헌", { placeholder: "title" });
  const refs = [
    ["서울특별시(2024.11.25.), 「새벽 출근길 돕는 '자율주행버스' 나왔다! 무료 운행」, 내 손안에 서울", "https://mediahub.seoul.go.kr/archives/2012789"],
    ["서울특별시(2026.3.29.), 새벽동행 자율주행버스 A741 운행 및 A160 성과", "https://www.seoul.go.kr/news/news_report.do?nttNo=455157&srchCtgry=467"],
    ["LTA(2025.9.20.), New Autonomous Shuttle Services to be Progressively Deployed in Punggol", "https://www.lta.gov.sg/content/ltagov/en/newsroom/2025/9/news-releases/new-autonomous-shuttle-services-to-be-progressively-deployed-in-.html"],
    ["Grab(2026.4.1.), 「Grab in partnership with WeRide officially launch Singapore's first autonomous public ride service in Punggol」", "https://www.grab.com/sg/press/others/grab-in-partnership-with-weride-officially-launch-singapores-first-autonomous-public-ride-service-in-punggol/"],
    ["LTA(2026.7.29.), Service Enhancements to Autonomous Shuttle Services in Punggol", "https://www.lta.gov.sg/content/ltagov/en/newsroom/2026/7/news-releases/service-enhancements-autonomous-shuttle-Punggol.html"],
    ["Grab(2026.10.5. 갱신), Punggol 자율주행 셔틀 및 수요응답형 서비스 안내", "https://www.grab.com/inside-grab/stories/autonomous-shuttle-self-driving-punggol-route-app-booking/"],
    ["ComfortDelGro 운행시간 안내", "https://www.cdgtaxi.com.sg/zig_driverless/"],
  ];
  const runs = [];
  refs.forEach(([t, u], i) => {
    runs.push({ text: t, options: { fontSize: 15, color: C.text1, breakLine: true } });
    runs.push({ text: u, options: { fontSize: 12, color: C.accent2, hyperlink: { url: u }, breakLine: i < refs.length - 1, paraSpaceAfter: 16 } });
  });
  tx(s, runs, { x: L, y: BODY_Y + 0.05, w: CW, h: 5.3, valign: "top", paraSpaceAfter: 1, name: "references" });
  s.addNotes("보고서 p.7의 참고문헌 7건입니다. URL의 utm 추적 파라미터만 제거했습니다.");
}

(async () => {
  slideCover();
  slideAgenda();
  slideBackground();
  slideSeoulOverview();
  slideSeoulResults();
  slidePunggolOverview();
  slidePunggolBackground();
  slidePunggolResults();
  slidePunggolOnDemand();
  slideCompareTable();
  slideImplications();
  slideDirection();
  slidePlan();
  slideReferences();
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})().catch((e) => { console.error(e); process.exit(1); });
