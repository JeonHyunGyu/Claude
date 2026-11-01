// 스마트시티 자율주행 대중교통 정책의 국내외 사례 비교 및 발전방향 — 연구실 발표 덱
// 모든 수치·사실은 보고서(2026.11.05., 전현규) 원문에 근거함. 원문 페이지는 각 슬라이드 하단 출처에 표기.
// 실행: node tools/build_autonomous_transit_deck.js <out.pptx>
//   (pptxgenjs, react-icons, react, react-dom, sharp 필요. APPLY_THEME 환경변수로 apply_theme.js 경로 지정 가능)
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const { applyTheme } = require(process.env.APPLY_THEME ||
  "/root/.claude/skills/synced/2709c09c-c0dc-4f81-b385-fcb3527e2d18_99de5be8-b623-4241-b2e1-4194ab58e11a/pptx/scripts/apply_theme.js");

const OUT = process.argv[2] || "deck.pptx";
const FIG1 = path.join(__dirname, "fig1_seoul_a160.jpg");      // 보고서 [그림 1], 760×509
const FIG2 = path.join(__dirname, "fig2_punggol_shuttle.jpg"); // 보고서 [그림 2], 1037×496

// 기존 연구실 덱(교통용량분석_중간발표.pptx)과 같은 디자인 체계
const THEME = {
  name: "Autonomous Transit Policy",
  headFontFace: "Malgun Gothic",
  bodyFontFace: "Malgun Gothic",
  colors: {
    dk1: "1A1F36", lt1: "FFFFFF", dk2: "1F2A5C", lt2: "EEF3FA",
    accent1: "2E6FD8", accent2: "1F2A5C", accent3: "8FB3EA",
    accent4: "C26A00", accent5: "6B7280", accent6: "D6E2F5",
    hlink: "2E6FD8", folHlink: "5B5F97",
  },
};
const HEX = THEME.colors;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "스마트시티 자율주행 대중교통 정책의 국내외 사례 비교 및 발전방향";
pres.author = "전현규";
pres.subject = "연구실 발표";
const C = pres.SchemeColor;

const L = 0.6, CW = 12.13, R = L + CW;
const TOP = 1.95;
const FOOTER = "스마트시티 자율주행 대중교통 정책의 국내외 사례 비교 및 발전방향  |  2026.11.05.";

// ---------- Layouts ----------
pres.defineSlideMaster({
  title: "COVER",
  background: { color: C.background1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: L, y: 1.15, w: 11.0, h: 1.6,
      fontSize: 36, bold: true, color: C.text2, valign: "top", align: "left", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: C.background1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: L, y: 0.6, w: CW, h: 0.6,
      fontSize: 28, bold: true, color: C.text2, valign: "middle", align: "left", margin: 0 }, text: "" } },
    { text: { text: FOOTER, options: { x: L, y: 7.0, w: 9, h: 0.3, fontSize: 10, color: C.accent5, margin: 0 } } },
  ],
  slideNumber: { x: 12.23, y: 7.0, w: 0.5, h: 0.3, fontSize: 10, color: C.accent5, align: "right" },
});

// ---------- Helpers ----------
let uid = 0;
const nm = (s) => `${s}-${++uid}`;
const tx = (s, text, o) => s.addText(text, { margin: 0, isTextBox: true, objectName: nm(o.name || "text"), ...o, name: undefined });

function card(s, x, y, w, h, fill = C.background2, line) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: fill },
    line: line ? { color: line, width: 1.25 } : { color: fill, width: 0 }, objectName: nm("card") });
}
function kicker(s, text) {
  tx(s, text, { x: L, y: 0.3, w: 8, h: 0.3, fontSize: 13, bold: true, color: C.accent1, name: "kicker" });
}
function message(s, text) {
  tx(s, text, { x: L, y: 1.27, w: CW, h: 0.45, fontSize: 17, color: C.text1, valign: "middle", name: "key-message" });
}
function source(s, text, y = 6.58) {
  tx(s, text, { x: L, y, w: CW, h: 0.3, fontSize: 10.5, italic: true, color: C.accent5, valign: "middle", name: "source" });
}
function arrowRight(s, x, y, w = 0.36, h = 0.36, color = C.accent3) {
  s.addShape(pres.shapes.ISOSCELES_TRIANGLE, { x, y, w: h, h: w, rotate: 90, fill: { color }, line: { color, width: 0 }, objectName: nm("arrow") });
}
function arrowDown(s, x, y, w = 0.36, h = 0.3, color = C.accent3) {
  s.addShape(pres.shapes.ISOSCELES_TRIANGLE, { x, y, w, h, rotate: 180, fill: { color }, line: { color, width: 0 }, objectName: nm("arrow") });
}
function numCircle(s, x, y, n, d = 0.5, fill = C.text2) {
  tx(s, String(n), { x, y, w: d, h: d, shape: pres.shapes.OVAL, fill: { color: fill }, line: { color: fill, width: 0 },
    color: C.background1, fontSize: 16, bold: true, align: "center", valign: "middle", name: "num" });
}
const bl = (items, o = {}) => items.map((t, i) => ({ text: t, options: { bullet: { indent: 16 }, breakLine: i < items.length - 1, ...o } }));

async function icon(Comp, hex, px = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: `#${hex}`, size: px }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}
function iconBadge(s, data, x, y, d = 0.62, fill = C.accent6) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill, width: 0 }, objectName: nm("icon-bg") });
  const p = d * 0.25;
  s.addImage({ data, x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p, altText: "장식용 아이콘", objectName: nm("icon") });
}
const hdrCell = (t, o = {}) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text2 }, align: "center", valign: "middle", ...o } });

async function build() {
  const I = {
    bus: await icon(fa.FaBus, HEX.dk2),
    shuttle: await icon(fa.FaShuttleVan, HEX.accent1),
    clock: await icon(fa.FaRegClock, HEX.dk2),
    map: await icon(fa.FaMapMarkedAlt, HEX.accent1),
    check: await icon(fa.FaClipboardCheck, HEX.dk2),
    chair: await icon(fa.FaChair, HEX.dk2),
    mobile: await icon(fa.FaMobileAlt, HEX.dk2),
    home: await icon(fa.FaHome, HEX.dk2),
    subway: await icon(fa.FaSubway, HEX.dk2),
    hospital: await icon(fa.FaHospital, HEX.dk2),
    store: await icon(fa.FaStore, HEX.dk2),
    bag: await icon(fa.FaShoppingBag, HEX.dk2),
    route: await icon(fa.FaRoute, HEX.dk2),
    users: await icon(fa.FaUserCheck, HEX.dk2),
    chart: await icon(fa.FaChartLine, HEX.dk2),
    busW: await icon(fa.FaBus, HEX.lt1),
    shuttleW: await icon(fa.FaShuttleVan, HEX.lt1),
  };

  // ======================= S1 표지 =======================
  pres.addSection({ title: "표지" });
  {
    const s = pres.addSlide({ masterName: "COVER", sectionTitle: "표지" });
    tx(s, "연구실 발표", { x: L, y: 0.6, w: 6, h: 0.35, fontSize: 14, bold: true, color: C.accent1 });
    s.addText("스마트시티 자율주행 대중교통 정책의\n국내외 사례 비교 및 발전방향", { placeholder: "title" });
    tx(s, "서울 새벽 동행 자율주행 버스(A160)  ·  싱가포르 Punggol 자율주행 셔틀", { x: L, y: 2.95, w: 11, h: 0.4, fontSize: 18, color: C.text1 });

    const cy = 3.95, ch = 1.55, cw = (CW - 0.4) / 2;
    [[L, I.bus, "국내 사례", "서울 새벽 동행 자율주행 버스", "시간적 교통 공백 대응"],
     [L + cw + 0.4, I.shuttle, "해외 사례", "싱가포르 Punggol 자율주행 셔틀", "공간적 교통 공백 대응"]].forEach(([x, ic, tag, main, sub]) => {
      card(s, x, cy, cw, ch);
      iconBadge(s, ic, x + 0.3, cy + (ch - 0.75) / 2, 0.75, C.background1);
      tx(s, [
        { text: tag, options: { fontSize: 13, bold: true, color: C.accent1, breakLine: true } },
        { text: main, options: { fontSize: 18, bold: true, color: C.text2, breakLine: true } },
        { text: sub, options: { fontSize: 14, color: C.text1 } },
      ], { x: x + 1.3, y: cy + 0.15, w: cw - 1.5, h: ch - 0.3, valign: "middle", paraSpaceAfter: 3 });
    });

    tx(s, [{ text: "2026.11.05.(목)", options: { breakLine: true } }, { text: "전현규", options: { bold: true, color: C.text2 } }],
      { x: L, y: 5.95, w: 5, h: 0.75, fontSize: 16, color: C.text1, paraSpaceAfter: 4 });
    tx(s, "※ 사례 구분(시간적·공간적 교통 공백)은 보고서 p.1의 사례 선정 근거에 따름", { x: L, y: 6.95, w: 9, h: 0.3, fontSize: 10.5, italic: true, color: C.accent5 });
    s.addNotes("스마트시티 자율주행 대중교통 정책의 국내외 사례 비교 및 발전방향에 대해 발표하겠습니다. 국내 사례로 서울 새벽 동행 자율주행 버스, 해외 사례로 싱가포르 Punggol 자율주행 셔틀을 비교합니다.");
  }

  // ======================= S2 발표 순서 =======================
  pres.addSection({ title: "Ⅰ. 연구 배경 및 목적" });
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Ⅰ. 연구 배경 및 목적" });
    s.addText("발표 순서", { placeholder: "title" });
    const items = [
      ["연구 배경·목적", "교통 공백과 사례 선정"],
      ["국내 사례", "서울 새벽 동행\n자율주행 버스(A160)"],
      ["해외 사례", "싱가포르 Punggol\n자율주행 셔틀"],
      ["비교 및 시사점", "도입 목적·운영방식·\n성과평가 비교"],
      ["발전 방향", "운영방식 선정 및\n기말 발표 계획"],
    ];
    const gap = 0.3, w = (CW - gap * 4) / 5, y = 1.85, h = 3.5;
    items.forEach(([t, d], i) => {
      const x = L + i * (w + gap);
      const hot = i === 1 || i === 2;
      card(s, x, y, w, h, hot ? C.text2 : C.background2);
      tx(s, String(i + 1).padStart(2, "0"), { x: x + 0.25, y: y + 0.3, w: w - 0.5, h: 0.8, fontSize: 36, bold: true, color: hot ? C.accent3 : C.accent1 });
      tx(s, t, { x: x + 0.25, y: y + 1.25, w: w - 0.5, h: 0.5, fontSize: 19, bold: true, color: hot ? C.background1 : C.text2 });
      tx(s, d, { x: x + 0.25, y: y + 1.85, w: w - 0.5, h: 1.1, fontSize: 14, color: hot ? C.background1 : C.text1, valign: "top" });
    });
    tx(s, "핵심: 두 사례의 도입 배경, 활용 기술 및 운영방식, 성과평가 방법을 비교하고 타 도시 적용조건과 발전 방향을 검토", { x: L, y: 5.75, w: CW, h: 0.4, fontSize: 15, color: C.accent5 });
    s.addNotes("발표는 연구 배경과 목적, 국내 사례, 해외 사례, 두 사례의 비교와 시사점, 그리고 발전 방향과 기말 발표 계획 순서로 진행합니다.");
  }

  // ======================= S3 연구 배경 =======================
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Ⅰ. 연구 배경 및 목적" });
    kicker(s, "Ⅰ. 연구 배경 및 목적");
    s.addText("연구 배경: 기존 대중교통의 교통 공백", { placeholder: "title" });
    message(s, "자율주행 버스·셔틀은 기존 대중교통체계의 한계를 보완하기 위해 도입이 추진되고 있음");

    // 상단: 도입 배경 원문 요약
    card(s, L, TOP, CW, 1.05);
    tx(s, bl([
      "도시 내 대중교통 공급이 부족한 시간대와 기존 교통망의 접근성이 낮은 지역에서 발생하는 이동 불편을 해소할 필요성 증대",
      "기존 대중교통체계의 한계를 보완하기 위한 자율주행 버스 및 셔틀 서비스 도입 추진",
    ], { fontSize: 15, color: C.text1 }), { x: L + 0.3, y: TOP + 0.1, w: CW - 0.6, h: 0.85, valign: "middle", paraSpaceAfter: 4 });

    // 하단: 시간적 / 공간적 공백 → 사례
    const cy = 3.3, cw = (CW - 0.5) / 2;
    [[L, I.clock, "시간적 교통 공백", "새벽·심야 등 대중교통 공급이 부족한 시간대", "서울 새벽 동행 자율주행 버스", "새벽 시간대 첫차 혼잡 및 운전기사 수급 문제에 대응 (서울특별시, 2024)"],
     [L + cw + 0.5, I.map, "공간적 교통 공백", "대중교통 접근성이 낮거나 교통서비스 공급이 부족한 지역", "싱가포르 Punggol 자율주행 셔틀", "주거지역과 주요 대중교통 거점 간 First/Last Mile 연결성 개선 (LTA, 2025)"],
    ].forEach(([x, ic, head, def, caseName, caseDesc]) => {
      iconBadge(s, ic, x, cy, 0.62);
      tx(s, head, { x: x + 0.8, y: cy, w: cw - 0.8, h: 0.62, fontSize: 19, bold: true, color: C.text2, valign: "middle" });
      tx(s, def, { x: x + 0.8, y: cy + 0.68, w: cw - 0.8, h: 0.4, fontSize: 14, color: C.accent5, valign: "middle" });
      arrowDown(s, x + cw / 2 - 0.18, cy + 1.2, 0.36, 0.26);
      card(s, x, cy + 1.6, cw, 1.4, C.background1, C.accent3);
      tx(s, [
        { text: caseName, options: { fontSize: 17, bold: true, color: C.text2, breakLine: true } },
        { text: caseDesc, options: { fontSize: 14, color: C.text1 } },
      ], { x: x + 0.3, y: cy + 1.7, w: cw - 0.6, h: 1.2, valign: "middle", paraSpaceAfter: 4 });
    });
    source(s, "자료: 보고서 p.1(도입 배경·사례 선정), p.6(교통 공백 유형 정의)을 바탕으로 발표자 재구성");
    s.addNotes("도시에서는 대중교통 공급이 부족한 시간대와 접근성이 낮은 지역에서 이동 불편이 발생합니다. 이를 보완하기 위해 자율주행 버스와 셔틀 도입이 추진되고 있습니다. 서울은 새벽 시간대라는 시간적 공백에, 싱가포르는 주거지와 교통거점 간 연결이라는 공간적 공백에 대응하는 사례입니다.");
  }

  // ======================= S4 사례 선정 및 발표 목적 =======================
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Ⅰ. 연구 배경 및 목적" });
    kicker(s, "Ⅰ. 연구 배경 및 목적");
    s.addText("사례 선정 및 발표 목적", { placeholder: "title" });
    message(s, "시간적·공간적 교통 공백에 각각 대응하는 두 사례를 비교해 운영방식의 차이를 검토");

    const y = 2.2, h = 3.6;
    // 1단: 사례
    const c1w = 3.3;
    tx(s, "비교 대상", { x: L, y, w: c1w, h: 0.4, fontSize: 15, bold: true, color: C.accent1 });
    [["국내", "서울 새벽 동행\n자율주행 버스(A160)", C.text2], ["해외", "싱가포르 Punggol\n자율주행 셔틀", C.accent1]].forEach(([tag, t, f], i) => {
      const yy = y + 0.55 + i * 1.55;
      card(s, L, yy, c1w, 1.35, f);
      tx(s, [{ text: tag, options: { fontSize: 13, bold: true, color: C.accent6, breakLine: true } }, { text: t, options: { fontSize: 16, bold: true, color: C.background1 } }],
        { x: L + 0.25, y: yy + 0.1, w: c1w - 0.5, h: 1.15, valign: "middle", paraSpaceAfter: 3 });
    });
    arrowRight(s, L + c1w + 0.17, y + 1.95, 0.36, 0.36);

    // 2단: 비교 항목
    const c2x = L + c1w + 0.7, c2w = 4.0;
    tx(s, "비교 항목", { x: c2x, y, w: c2w, h: 0.4, fontSize: 15, bold: true, color: C.accent1 });
    ["도입 배경", "활용 기술 및 운영방식", "성과평가 방법"].forEach((t, i) => {
      const yy = y + 0.55 + i * 1.03;
      card(s, c2x, yy, c2w, 0.85);
      numCircle(s, c2x + 0.2, yy + 0.18, i + 1, 0.5);
      tx(s, t, { x: c2x + 0.9, y: yy, w: c2w - 1.0, h: 0.85, fontSize: 17, bold: true, color: C.text2, valign: "middle" });
    });
    arrowRight(s, c2x + c2w + 0.17, y + 1.95, 0.36, 0.36);

    // 3단: 목적
    const c3x = c2x + c2w + 0.7, c3w = R - c3x;
    tx(s, "발표 목적", { x: c3x, y, w: c3w, h: 0.4, fontSize: 15, bold: true, color: C.accent1 });
    card(s, c3x, y + 0.55, c3w, 2.9, C.background1, C.text2);
    tx(s, bl(["국내외 자율주행 대중교통 정책의 도입 배경, 활용 기술 및 운영방식, 성과평가 방법 비교", "사례 분석을 통한 타 도시 적용조건 및 향후 발전 방향 검토"], { fontSize: 15, color: C.text1 }),
      { x: c3x + 0.25, y: y + 0.7, w: c3w - 0.5, h: 2.6, valign: "middle", paraSpaceAfter: 10 });

    tx(s, "※ 두 사례는 각각 시간적·공간적 교통 공백에 대응하는 사례로, 도시문제에 따른 자율주행 대중교통 운영방식의 차이를 비교할 수 있음", { x: L, y: 6.0, w: CW, h: 0.45, fontSize: 13, color: C.accent5 });
    source(s, "자료: 보고서 p.1(국내외 정책 사례 선정, 발표 목적)을 바탕으로 발표자 재구성");
    s.addNotes("두 사례를 도입 배경, 활용 기술 및 운영방식, 성과평가 방법의 세 측면에서 비교하고, 이를 통해 타 도시 적용조건과 향후 발전 방향을 검토하는 것이 발표의 목적입니다.");
  }

  // ======================= S5 서울 사업 개요 =======================
  pres.addSection({ title: "Ⅱ. 국내 사례" });
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Ⅱ. 국내 사례" });
    kicker(s, "Ⅱ. 국내 사례 · 서울 새벽 동행 자율주행 버스");
    s.addText("사업 개요: 새벽 동행 자율주행 버스 A160", { placeholder: "title" });
    message(s, "기존 160번 노선을 기반으로 일반 시내버스 첫차보다 약 30분 일찍 운행");

    const iw = 6.1, ih = iw * 509 / 760;
    s.addImage({ path: FIG1, x: L, y: TOP, w: iw, h: ih, altText: "서울 새벽 동행 자율주행 버스 A160 (보고서 그림 1)", objectName: "fig1" });
    tx(s, "[그림 1] 서울 새벽 동행 자율주행 버스 A160 / 자료: 서울특별시(2024)", { x: L, y: TOP + ih + 0.08, w: iw, h: 0.3, fontSize: 11, color: C.accent5 });

    const rx = L + iw + 0.45, rw = R - rx;
    const tiles = [["25.7km", "편도 운행구간"], ["03:30", "평일 출발"], ["87개", "운행 정류소"]];
    const tw = (rw - 0.3) / 3;
    tiles.forEach(([v, l], i) => {
      const x = rx + i * (tw + 0.15);
      card(s, x, TOP, tw, 1.3);
      tx(s, v, { x, y: TOP + 0.15, w: tw, h: 0.65, fontSize: 26, bold: true, color: C.text2, align: "center", valign: "middle" });
      tx(s, l, { x, y: TOP + 0.8, w: tw, h: 0.35, fontSize: 13, color: C.accent5, align: "center", valign: "middle" });
    });
    const lab = (t) => ({ text: t, options: { bold: true, color: C.text2, fill: { color: C.background2 } } });
    const rows = [
      [lab("사업명"), "서울 새벽 동행 자율주행 버스(A160)"],
      [lab("최초 운행일"), "2024년 11월 26일"],
      [lab("운행구간"), "도봉산역 광역환승센터 ~ 영등포역\n(편도 25.7km)"],
      [lab("운행 시간"), "평일 오전 3시 30분 출발\n(일반 시내버스 첫차보다 약 30분 조기 운행)"],
      [lab("운영방식"), "기존 160번 시내버스 노선 기반,\n주요 구간의 87개 정류소 운행"],
    ];
    s.addTable(rows, { x: rx, y: TOP + 1.5, w: rw, colW: [1.45, rw - 1.45], rowH: [0.45, 0.45, 0.72, 0.72, 0.72], fontSize: 14, color: C.text1,
      border: { type: "solid", pt: 0.75, color: HEX.accent3 }, valign: "middle", margin: [0.03, 0.1, 0.03, 0.12], objectName: "a160-facts" });
    source(s, "자료: 보고서 p.1–2(사업 개요, [그림 1])");
    s.addNotes("서울 새벽 동행 자율주행 버스 A160은 2024년 11월 26일 처음 운행했습니다. 도봉산역 광역환승센터에서 영등포역까지 편도 25.7km 구간이며, 평일 오전 3시 30분에 출발해 일반 시내버스 첫차보다 약 30분 일찍 운행합니다. 기존 160번 노선을 기반으로 주요 구간의 87개 정류소를 운행합니다.");
  }

  // ======================= S6 서울 도입 배경 =======================
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Ⅱ. 국내 사례" });
    kicker(s, "Ⅱ. 국내 사례 · 서울 새벽 동행 자율주행 버스");
    s.addText("도입 배경 및 목적", { placeholder: "title" });
    message(s, "새벽 첫차 혼잡과 운전기사 수급난에 자율주행 기술로 대응");

    const lw = 7.0, y1 = TOP + 0.45, h1 = 1.95, h2 = 1.2;
    tx(s, "문제", { x: L, y: TOP, w: 3, h: 0.35, fontSize: 15, bold: true, color: C.accent1 });
    // 문제 ①
    card(s, L, y1, lw, h1);
    numCircle(s, L + 0.25, y1 + 0.25, 1, 0.45);
    tx(s, [
      { text: "새벽 출근 근로자의 이동 불편 및 첫차 혼잡", options: { fontSize: 17, bold: true, color: C.text2, breakLine: true } },
      { text: "환경미화원, 경비원 등 새벽 시간대 출근 근로자", options: { fontSize: 14, color: C.text1 } },
    ], { x: L + 0.9, y: y1 + 0.18, w: lw - 1.1, h: 0.85, valign: "top", paraSpaceAfter: 4 });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: L + 0.9, y: y1 + 1.08, w: lw - 1.15, h: 0.65, rectRadius: 0.06, fill: { color: C.background1 }, line: { color: C.accent3, width: 1 }, objectName: "evidence" });
    tx(s, [
      { text: "50명 초과  ", options: { fontSize: 20, bold: true, color: C.accent1 } },
      { text: "기존 160번 노선 새벽 첫차의 최대 재차 인원", options: { fontSize: 14, color: C.text1 } },
    ], { x: L + 1.1, y: y1 + 1.08, w: lw - 1.5, h: 0.65, valign: "middle" });
    // 문제 ②
    const y2 = y1 + h1 + 0.3;
    card(s, L, y2, lw, h2);
    numCircle(s, L + 0.25, y2 + 0.25, 2, 0.45);
    tx(s, [
      { text: "새벽 시간대 운전기사 수급의 어려움", options: { fontSize: 17, bold: true, color: C.text2, breakLine: true } },
      { text: "운전기사 수급 문제로 인한 버스 증차 제약", options: { fontSize: 14, color: C.text1 } },
    ], { x: L + 0.9, y: y2 + 0.18, w: lw - 1.1, h: 0.85, valign: "top", paraSpaceAfter: 4 });

    arrowRight(s, L + lw + 0.2, y1 + (h1 + 0.3 + h2) / 2 - 0.2, 0.4, 0.4, C.text2);

    // 대응
    const rx = L + lw + 0.85, rw = R - rx;
    tx(s, "대응", { x: rx, y: TOP, w: 3, h: 0.35, fontSize: 15, bold: true, color: C.accent1 });
    card(s, rx, y1, rw, h1 + 0.3 + h2, C.text2);
    iconBadge(s, I.busW, rx + 0.3, y1 + 0.3, 0.7, C.accent1);
    tx(s, [
      { text: "자율주행 버스 도입", options: { fontSize: 19, bold: true, color: C.background1, breakLine: true } },
      { text: " ", options: { fontSize: 8, breakLine: true } },
      ...bl(["교통서비스 공급 시간 확대", "대중교통 이용 편의 증진"], { fontSize: 15, color: C.background1 }),
    ], { x: rx + 0.3, y: y1 + 1.25, w: rw - 0.6, h: 1.9, valign: "top", paraSpaceAfter: 6 });
    source(s, "자료: 보고서 p.2(도입 배경 및 목적)를 바탕으로 발표자 재구성");
    s.addNotes("A160 도입 배경은 두 가지입니다. 첫째, 환경미화원·경비원 등 새벽 출근 근로자의 이동 불편과 첫차 혼잡으로, 기존 160번 노선의 새벽 첫차 최대 재차 인원이 50명을 초과했습니다. 둘째, 새벽 시간대 운전기사 수급이 어려워 증차에 제약이 있었습니다. 이에 자율주행 기술로 교통서비스 공급 시간을 확대하고 이용 편의를 높이고자 했습니다.");
  }

  // ======================= S7 서울 적용 기술 및 운영방식 =======================
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Ⅱ. 국내 사례" });
    kicker(s, "Ⅱ. 국내 사례 · 서울 새벽 동행 자율주행 버스");
    s.addText("적용 기술 및 운영방식", { placeholder: "title" });
    message(s, "기존 시내버스 노선을 활용한 자율주행 버스 운행 및 단계적 상용화 추진");

    const gap = 0.35, w = (CW - 2 * gap) / 3, y = TOP + 0.1, h = 4.3;
    const cols = [
      [I.check, "단계적 상용화 절차", "제도적·기술적 절차 추진", null, ["자율차 시범운행 지구 지정", "시험 운행", "안전 검증"]],
      [I.chair, "좌석제 운영", "승객 안전을 고려한 운영", ["22명", "최대 탑승 인원"], ["입석 금지"]],
      [I.mobile, "실시간 정보 제공", "버스정보안내단말기(BIT) 및 모바일 서비스", null, ["실시간 위치", "도착 정보", "잔여 좌석 정보"]],
    ];
    cols.forEach(([ic, head, sub, stat, items], i) => {
      const x = L + i * (w + gap);
      card(s, x, y, w, h);
      iconBadge(s, ic, x + 0.3, y + 0.3, 0.7, C.background1);
      tx(s, head, { x: x + 0.3, y: y + 1.15, w: w - 0.6, h: 0.45, fontSize: 19, bold: true, color: C.text2 });
      tx(s, sub, { x: x + 0.3, y: y + 1.62, w: w - 0.6, h: 0.65, fontSize: 14, color: C.accent5, valign: "top" });
      let yy = y + 2.35;
      if (stat) {
        tx(s, [{ text: stat[0], options: { fontSize: 36, bold: true, color: C.accent1 } }, { text: "  " + stat[1], options: { fontSize: 14, color: C.text1 } }],
          { x: x + 0.3, y: yy, w: w - 0.6, h: 0.8, valign: "middle" });
        yy += 0.95;
      }
      tx(s, bl(items, { fontSize: 15, color: C.text1 }), { x: x + 0.3, y: yy, w: w - 0.6, h: y + h - yy - 0.25, valign: "top", paraSpaceAfter: 6 });
    });
    source(s, "자료: 보고서 p.2(적용 기술 및 운영방식). 첫 번째 열의 항목은 원문의 나열을 옮긴 것으로 순서를 의미하지 않음");
    s.addNotes("A160은 기존 시내버스 노선을 활용해 단계적 상용화를 추진했습니다. 자율차 시범운행 지구 지정, 시험 운행, 안전 검증 등 제도적·기술적 절차를 거쳤습니다. 승객 안전을 위해 최대 22명 좌석제로 운영하고 입석은 금지합니다. 버스정보안내단말기와 모바일로 실시간 위치, 도착 정보, 잔여 좌석 정보를 제공합니다.");
  }

  // ======================= S8 서울 운영성과 =======================
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Ⅱ. 국내 사례" });
    kicker(s, "Ⅱ. 국내 사례 · 서울 새벽 동행 자율주행 버스");
    s.addText("운영성과 및 확대 현황", { placeholder: "title" });
    message(s, "약 15개월간 사고 0건, 이용자의 96.2%가 통근 목적으로 도입 취지에 부합");

    tx(s, "<표 1> 서울 새벽 동행 자율주행 버스 A160 주요 운영성과", { x: L, y: TOP - 0.05, w: 8, h: 0.35, fontSize: 13, bold: true, color: C.text2 });
    const kpis = [["약 27,600명", "누적 이용객"], ["0건", "자율주행 사고"], ["4.08 / 5점", "서비스 만족도(평균)"], ["96.2%", "통근 목적 이용자 비율"], ["85.4%", "50대 이상 이용자 비율"]];
    const gap = 0.25, w = (CW - 4 * gap) / 5, y = TOP + 0.4, h = 1.75;
    kpis.forEach(([v, l], i) => {
      const x = L + i * (w + gap);
      const hot = i === 3;
      card(s, x, y, w, h, hot ? C.text2 : C.background2);
      tx(s, v, { x: x + 0.1, y: y + 0.25, w: w - 0.2, h: 0.8, fontSize: 23, bold: true, color: hot ? C.background1 : C.text2, align: "center", valign: "middle" });
      tx(s, l, { x: x + 0.1, y: y + 1.1, w: w - 0.2, h: 0.45, fontSize: 14, color: hot ? C.accent6 : C.text1, align: "center", valign: "middle" });
    });
    tx(s, "기간: A160 운행 시작 이후 약 15개월(2026년 3월 발표 기준)", { x: L, y: y + h + 0.1, w: CW, h: 0.32, fontSize: 12, color: C.accent5 });

    // 하단 좌: 해석(원문)
    const by = 4.75, bh = 1.6, lw = 5.4;
    card(s, L, by, lw, bh, C.background1, C.accent3);
    tx(s, [
      { text: "도입 목적과의 부합성", options: { fontSize: 16, bold: true, color: C.text2, breakLine: true } },
      { text: "통근 목적 이용자 96.2% → 새벽 근로자의 이동 지원이라는 도입 목적에 부합", options: { fontSize: 14, color: C.text1 } },
    ], { x: L + 0.25, y: by + 0.12, w: lw - 0.5, h: bh - 0.24, valign: "middle", paraSpaceAfter: 6 });

    // 하단 우: 노선 확대
    const rx = L + lw + 0.35, rw = R - rx;
    card(s, rx, by, rw, bh);
    tx(s, "2026년 노선 확대: 총 4개 노선", { x: rx + 0.25, y: by + 0.12, w: rw - 0.5, h: 0.4, fontSize: 16, bold: true, color: C.text2 });
    const routes = [["A160", "기존"], ["A741", "추가"], ["A148", "추가"], ["A504", "추가"]];
    const pw = (rw - 0.5 - 3 * 0.2) / 4;
    routes.forEach(([r, t], i) => {
      const x = rx + 0.25 + i * (pw + 0.2);
      const base = i === 0;
      tx(s, [{ text: r, options: { fontSize: 17, bold: true, breakLine: true } }, { text: t, options: { fontSize: 12 } }], {
        x, y: by + 0.65, w: pw, h: 0.75, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.08,
        fill: { color: base ? C.text2 : C.background1 }, line: { color: C.text2, width: 1.25 },
        color: base ? C.background1 : C.text2, align: "center", valign: "middle", name: "route" });
    });
    source(s, "자료: 서울특별시(2026.3.29.), 「새벽 동행 자율주행 버스 'A741' 첫 시동…새벽 출근길 더 빨라진다」를 바탕으로 작성 (보고서 <표 1>, p.2–3)");
    s.addNotes("A160은 운행 시작 이후 약 15개월간 누적 이용객 약 27,600명을 기록했습니다(2026년 3월 발표 기준). 같은 기간 자율주행 사고는 0건, 서비스 만족도는 평균 4.08점입니다. 통근 목적 이용자가 96.2%로 새벽 근로자 이동 지원이라는 도입 목적에 부합했고, 50대 이상 이용자 비율은 85.4%입니다. 2026년에는 A741, A148, A504 노선이 추가되어 총 4개 노선으로 확대되었습니다.");
  }

  // ======================= S9 싱가포르 사업 개요 =======================
  pres.addSection({ title: "Ⅲ. 해외 사례" });
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Ⅲ. 해외 사례" });
    kicker(s, "Ⅲ. 해외 사례 · 싱가포르 Punggol 자율주행 셔틀");
    s.addText("사업 개요: Punggol 자율주행 셔틀", { placeholder: "title" });
    message(s, "LTA가 시행하고 Grab·ComfortDelGro가 운영하는 3개 고정노선의 소형 자율주행 셔틀");

    const iw = 5.6, ih = iw * 496 / 1037;
    s.addImage({ path: FIG2, x: L, y: TOP, w: iw, h: ih, altText: "싱가포르 Grab·WeRide의 자율주행 셔틀 차량 (보고서 그림 2)", objectName: "fig2" });
    tx(s, [{ text: "[그림 2] 싱가포르 Grab·WeRide의 자율주행 셔틀 차량", options: { breakLine: true } }, { text: "자료: Grab(2026.10.5. 갱신)" }],
      { x: L, y: TOP + ih + 0.06, w: iw, h: 0.5, fontSize: 11, color: C.accent5, valign: "top" });
    const ty = TOP + ih + 0.68, tw = (iw - 0.3) / 3, th = 1.05;
    [["3개", "고정노선"], ["4~7개", "노선별 승하차 지점"], ["5~8인승", "자율주행차량"]].forEach(([v, l], i) => {
      const x = L + i * (tw + 0.15);
      card(s, x, ty, tw, th);
      tx(s, v, { x, y: ty + 0.08, w: tw, h: 0.55, fontSize: 22, bold: true, color: C.text2, align: "center", valign: "middle" });
      tx(s, l, { x, y: ty + 0.62, w: tw, h: 0.35, fontSize: 13, color: C.accent5, align: "center", valign: "middle" });
    });

    const rx = L + iw + 0.45, rw = R - rx;
    const lab = (t) => ({ text: t, options: { bold: true, color: C.text2, fill: { color: C.background2 } } });
    s.addTable([
      [lab("사업명"), "Punggol Autonomous Shuttle Services"],
      [lab("시행기관"), "싱가포르 육상교통청\n(Land Transport Authority, LTA)"],
      [lab("운영기관"), "Grab(WeRide), ComfortDelGro(Pony.ai)"],
      [lab("사업지역"), "싱가포르 Punggol 지역"],
    ], { x: rx, y: TOP, w: rw, colW: [1.3, rw - 1.3], rowH: [0.5, 0.75, 0.5, 0.5], fontSize: 14, color: C.text1,
      border: { type: "solid", pt: 0.75, color: HEX.accent3 }, valign: "middle", margin: [0.03, 0.1, 0.03, 0.12], objectName: "punggol-facts" });

    tx(s, "운행 시간: 주간 중심 운영", { x: rx, y: 4.45, w: rw, h: 0.35, fontSize: 15, bold: true, color: C.text2 });
    s.addTable([
      [hdrCell("노선"), hdrCell("운영기관(플랫폼)"), hdrCell("운행 시간")],
      [{ text: "Route 1·3", options: { align: "center" } }, "Grab·WeRide", "평일 09:30~17:30"],
      [{ text: "Route 2", options: { align: "center" } }, "ComfortDelGro·Pony.ai", "화~일요일 09:30~17:00"],
    ], { x: rx, y: 4.88, w: rw, colW: [1.3, 2.55, rw - 3.85], rowH: [0.42, 0.5, 0.5], fontSize: 14, color: C.text1,
      border: { type: "solid", pt: 0.75, color: HEX.accent3 }, valign: "middle", margin: [0.03, 0.1, 0.03, 0.1], objectName: "punggol-hours" });
    source(s, "자료: 보고서 p.3(사업 개요, [그림 2]). 운행 시간은 보고서 참고문헌의 ComfortDelGro 운행시간 안내 포함");
    s.addNotes("싱가포르 Punggol 자율주행 셔틀은 육상교통청 LTA가 시행하고, Grab(WeRide 플랫폼)과 ComfortDelGro(Pony.ai 플랫폼)가 운영합니다. 3개 고정노선에 노선별 4~7개 승하차 지점이 있으며 5~8인승 차량을 도입했습니다. 주간 중심으로, Route 1·3은 평일 09:30~17:30, Route 2는 화~일요일 09:30~17:00 운행합니다.");
  }

  // ======================= S10 싱가포르 도입 배경 =======================
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Ⅲ. 해외 사례" });
    kicker(s, "Ⅲ. 해외 사례 · 싱가포르 Punggol 자율주행 셔틀");
    s.addText("도입 배경 및 목적", { placeholder: "title" });
    message(s, "주거지역과 대중교통 거점·생활시설 간 First/Last Mile 연결성 개선");

    const y = TOP + 0.15;
    // 주거지역
    const dh = 0.8, dg = 0.16;
    const hw = 2.3, hh = 1.5, hy = y + 0.32 + (4 * dh + 3 * dg) / 2 - hh / 2;
    card(s, L, hy, hw, hh, C.background2);
    iconBadge(s, I.home, L + (hw - 0.6) / 2, hy + 0.2, 0.6, C.background1);
    tx(s, "주거지역", { x: L, y: hy + 0.9, w: hw, h: 0.4, fontSize: 16, bold: true, color: C.text2, align: "center" });
    arrowRight(s, L + hw + 0.12, hy + hh / 2 - 0.18, 0.36, 0.36);
    // 셔틀
    const sx = L + hw + 0.6, sw = 2.5;
    card(s, sx, hy - 0.2, sw, hh + 0.4, C.text2);
    iconBadge(s, I.shuttleW, sx + (sw - 0.6) / 2, hy - 0.02, 0.6, C.accent1);
    tx(s, [{ text: "자율주행 셔틀", options: { fontSize: 16, bold: true, breakLine: true } }, { text: "First/Last Mile", options: { fontSize: 13 } }],
      { x: sx, y: hy + 0.7, w: sw, h: 0.8, color: C.background1, align: "center", valign: "middle" });
    arrowRight(s, sx + sw + 0.12, hy + hh / 2 - 0.18, 0.36, 0.36);
    // 목적지
    const dx = sx + sw + 0.6, dw = 2.5;
    tx(s, "대중교통 거점 · 생활시설", { x: dx, y: y - 0.1, w: dw + 0.4, h: 0.35, fontSize: 13, bold: true, color: C.accent1 });
    [[I.subway, "MRT역"], [I.hospital, "의료시설"], [I.store, "시장"], [I.bag, "쇼핑 시설"]].forEach(([ic, t], i) => {
      const yy = y + 0.32 + i * (dh + dg);
      card(s, dx, yy, dw, dh, C.background1, C.accent3);
      iconBadge(s, ic, dx + 0.14, yy + (dh - 0.52) / 2, 0.52, C.background2);
      tx(s, t, { x: dx + 0.75, y: yy, w: dw - 0.85, h: dh, fontSize: 15, bold: true, color: C.text2, valign: "middle" });
    });
    tx(s, "※ 기존 대중교통으로 접근하기 불편한 의료시설, 시장, 쇼핑 시설 및 MRT역 등과의 연결성 강화", { x: L, y: 5.6, w: dx - L - 0.3, h: 0.65, fontSize: 14, color: C.accent5, valign: "top" });

    // 우측: 목표
    const rx = dx + dw + 0.45, rw = R - rx;
    card(s, rx, y + 0.3, rw, 3.7);
    tx(s, "기존 대중교통망 보완", { x: rx + 0.25, y: y + 0.5, w: rw - 0.5, h: 0.4, fontSize: 15, bold: true, color: C.text2 });
    tx(s, "지역 내 이동 편의성 및 접근성 향상", { x: rx + 0.25, y: y + 0.9, w: rw - 0.5, h: 0.65, fontSize: 14, color: C.text1 });
    tx(s, "최대 15분", { x: rx + 0.25, y: y + 1.6, w: rw - 0.5, h: 0.8, fontSize: 36, bold: true, color: C.accent1, valign: "middle" });
    tx(s, "기존 대중교통 대비 이동시간 단축 목표", { x: rx + 0.25, y: y + 2.4, w: rw - 0.5, h: 0.6, fontSize: 14, color: C.text1, valign: "top" });
    tx(s, "목표치 · 달성 실적 아님", { x: rx + 0.25, y: y + 3.25, w: 2.3, h: 0.36, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.15,
      fill: { color: C.background1 }, line: { color: C.accent4, width: 1.25 }, color: C.accent4, fontSize: 12, bold: true, align: "center", valign: "middle", name: "target-chip" });
    source(s, "자료: 보고서 p.4(도입 배경 및 목적)를 바탕으로 발표자 재구성");
    s.addNotes("Punggol 셔틀은 주거지역과 주요 대중교통 거점 및 생활시설 사이의 First/Last Mile 연결성을 개선하기 위해 도입되었습니다. 기존 대중교통으로 접근하기 불편한 의료시설, 시장, 쇼핑 시설, MRT역과의 연결을 강화하는 것이 목적입니다. 기존 대중교통 대비 최대 15분의 이동시간 단축을 목표로 설정했는데, 이는 실적이 아닌 목표치입니다.");
  }

  // ======================= S11 싱가포르 적용 기술 및 안전 검증 =======================
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Ⅲ. 해외 사례" });
    kicker(s, "Ⅲ. 해외 사례 · 싱가포르 Punggol 자율주행 셔틀");
    s.addText("적용 기술 및 단계별 안전 검증", { placeholder: "title" });
    message(s, "4단계 안전 검증을 거쳐 일반인 탑승으로 확대");

    const steps = ["폐쇄시험장\n(CETRAN) 검증", "실도로\n무승객 시험", "주민 대상\n무료 시승", "일반인 탑승"];
    const gap = 0.55, w = (CW - 3 * gap) / 4, y = TOP + 0.05, h = 1.2;
    steps.forEach((t, i) => {
      const x = L + i * (w + gap);
      const last = i === 3;
      card(s, x, y, w, h, last ? C.text2 : C.background2);
      numCircle(s, x + 0.2, y + (h - 0.48) / 2, i + 1, 0.48, last ? C.accent1 : C.text2);
      tx(s, t, { x: x + 0.85, y, w: w - 0.95, h, fontSize: 16, bold: true, color: last ? C.background1 : C.text2, valign: "middle" });
      if (i < 3) arrowRight(s, x + w + 0.12, y + h / 2 - 0.16, 0.32, 0.32, C.accent3);
    });

    const cy = 3.65, ch = 2.75, cws = [5.0, (CW - 5.0 - 0.7) / 2, (CW - 5.0 - 0.7) / 2];
    const cards = [
      [I.route, "적용 기술", ["지역 내 고정노선 셔틀 서비스", "Grab은 WeRide, ComfortDelGro는 Pony.ai의 자율주행 플랫폼 적용", "정밀 노선 매핑, 센서 보정, AI 모델 검증, 현지 도로 환경 적응"]],
      [I.check, "안전관리", ["차량 내 안전요원 탑승", "실시간 운행 모니터링"]],
      [I.users, "서비스 개선", ["이용자 의견과 운행 데이터를 수집하여 서비스 운영방식 개선 추진"]],
    ];
    cards.forEach(([ic, head, items], i) => {
      const cw = cws[i], x = L + cws.slice(0, i).reduce((a, b) => a + b + 0.35, 0);
      card(s, x, cy, cw, ch, C.background1, C.accent3);
      iconBadge(s, ic, x + 0.25, cy + 0.22, 0.55);
      tx(s, head, { x: x + 0.95, y: cy + 0.22, w: cw - 1.1, h: 0.55, fontSize: 17, bold: true, color: C.text2, valign: "middle" });
      tx(s, bl(items, { fontSize: 14, color: C.text1 }), { x: x + 0.25, y: cy + 0.95, w: cw - 0.5, h: ch - 1.1, valign: "top", paraSpaceAfter: 5 });
    });
    source(s, "자료: 보고서 p.4(적용 기술 및 운영방식)를 바탕으로 발표자 재구성");
    s.addNotes("Punggol 셔틀은 지역 내 고정노선 서비스로, Grab은 WeRide, ComfortDelGro는 Pony.ai의 자율주행 플랫폼을 적용했습니다. 정밀 노선 매핑, 센서 보정, AI 모델 검증, 현지 도로 환경 적응 과정을 거쳤습니다. 안전 검증은 폐쇄시험장 CETRAN 검증, 실도로 무승객 시험, 주민 대상 무료 시승, 일반인 탑승의 4단계로 진행되었고, 차량 내 안전요원과 실시간 모니터링으로 안전을 관리합니다. 이용자 의견과 운행 데이터를 수집해 운영방식을 개선하고 있습니다.");
  }

  // ======================= S12 싱가포르 운영성과 =======================
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Ⅲ. 해외 사례" });
    kicker(s, "Ⅲ. 해외 사례 · 싱가포르 Punggol 자율주행 셔틀");
    s.addText("운영성과 및 이용자 의견", { placeholder: "title" });
    message(s, "설문 응답자 99%가 안전하다고 느꼈으며, 고정노선·고정 승하차 지점 운영에 대한 이용자 요구 확인");

    tx(s, "<표 2> 주요 운영성과 및 이용자 의견", { x: L, y: TOP - 0.05, w: 6, h: 0.35, fontSize: 13, bold: true, color: C.text2 });
    const lw = 4.4, th = 1.1, tg = 0.15;
    [["11,500명 이상", "고유 이용자 수"], ["약 900명", "이용자 설문 응답"], ["99%", "체감 안전성 및 추천 의향"]].forEach(([v, l], i) => {
      const y = TOP + 0.4 + i * (th + tg);
      const hot = i === 2;
      card(s, L, y, lw, th, hot ? C.text2 : C.background2);
      tx(s, v, { x: L + 0.3, y: y + 0.1, w: lw - 0.6, h: 0.55, fontSize: 24, bold: true, color: hot ? C.background1 : C.text2, valign: "middle" });
      tx(s, l, { x: L + 0.3, y: y + 0.65, w: lw - 0.6, h: 0.35, fontSize: 14, color: hot ? C.accent6 : C.text1, valign: "middle" });
    });

    const rx = L + lw + 0.4, rw = R - rx, ry = TOP + 0.4, rh = 3 * th + 2 * tg;
    card(s, rx, ry, rw, rh, C.background1, C.accent3);
    tx(s, "고정노선·고정 승하차 지점 운영에 대한 이용자 요구", { x: rx + 0.3, y: ry + 0.2, w: rw - 0.6, h: 0.4, fontSize: 16, bold: true, color: C.text2 });
    tx(s, "설문 응답자 중 각 항목을 선호한 비율", { x: rx + 0.3, y: ry + 0.6, w: rw - 0.6, h: 0.32, fontSize: 13, color: C.accent5 });
    const bx = rx + 0.3, bw = rw - 0.6, labW = 2.6, trackW = bw - labW - 1.2;
    [["승하차 지점 선택 선호", 60, "약 60%", HEX.dk2], ["직접적인 이동 경로 선호", 40, "약 40%", HEX.accent1]].forEach(([lab, v, vt, col], i) => {
      const yy = ry + 1.15 + i * 0.9;
      tx(s, lab, { x: bx, y: yy, w: labW, h: 0.55, fontSize: 15, color: C.text1, valign: "middle" });
      s.addShape(pres.shapes.RECTANGLE, { x: bx + labW, y: yy + 0.08, w: trackW, h: 0.4, fill: { color: C.background2 }, line: { color: C.background2, width: 0 }, objectName: nm("bar-track") });
      s.addShape(pres.shapes.RECTANGLE, { x: bx + labW, y: yy + 0.08, w: trackW * v / 100, h: 0.4, fill: { color: col }, line: { color: col, width: 0 }, objectName: nm("bar") });
      tx(s, vt, { x: bx + labW + trackW + 0.15, y: yy, w: 1.0, h: 0.55, fontSize: 17, bold: true, color: C.text2, valign: "middle" });
    });
    tx(s, "※ 막대 길이는 원문의 '약 60%', '약 40%'를 표시한 것이며, 두 응답의 배타성 여부는 원문에 명시되지 않음", { x: rx + 0.3, y: ry + rh - 0.72, w: rw - 0.6, h: 0.6, fontSize: 12, color: C.accent5, valign: "top" });

    tx(s, "※ 이용자 수는 2026년 7월 12일 기준 / 수요응답형 확대 이전 고정노선 셔틀 단계의 실적", { x: L, y: 6.12, w: CW, h: 0.32, fontSize: 12, color: C.accent5 });
    source(s, "자료: LTA(2026.7.29.), 「Service Enhancements to Autonomous Shuttle Services in Punggol」을 바탕으로 작성 (보고서 <표 2>, p.4–5)");
    s.addNotes("2026년 7월 12일 기준 누적 고유 이용자는 11,500명 이상입니다. 탑승 후 설문 응답자 약 900명 중 99%가 탑승 중 안전하다고 느끼고 타인에게 추천할 의향을 나타냈습니다. 한편 고정노선과 고정 승하차 지점 운영에 대한 요구도 확인되었는데, 응답자 약 60%가 승하차 지점 선택을, 약 40%가 목적지까지 직접적인 경로를 선호했습니다. 이 실적은 수요응답형 확대 이전 고정노선 단계의 것입니다.");
  }

  // ======================= S13 수요응답형 병행 =======================
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Ⅲ. 해외 사례" });
    kicker(s, "Ⅲ. 해외 사례 · 싱가포르 Punggol 자율주행 셔틀");
    s.addText("서비스 개선: 고정노선과 수요응답형 병행", { placeholder: "title" });
    message(s, "이용자 의견을 반영하여 수요응답형(On-Demand) 서비스로 확대");

    // 타임라인
    const ty = 2.55, events = [
      ["2025.9.20.", "도입계획 발표"],
      ["2026.4.1.", "일반인 무료 탑승 개시\n(Grab Route 1·3 기준)"],
      ["2026.7.12.", "누적 고유 이용자\n11,500명 이상"],
      ["2026.9.", "수요응답형 서비스\n병행 운영"],
    ];
    const seg = CW / 4;
    s.addShape(pres.shapes.LINE, { x: L + seg / 2, y: ty, w: seg * 3, h: 0, line: { color: HEX.accent3, width: 2.5 }, objectName: "timeline" });
    events.forEach(([d, t], i) => {
      const cx = L + seg * i + seg / 2, last = i === 3;
      s.addShape(pres.shapes.OVAL, { x: cx - 0.14, y: ty - 0.14, w: 0.28, h: 0.28, fill: { color: last ? C.accent1 : C.text2 }, line: { color: C.background1, width: 2 }, objectName: nm("dot") });
      tx(s, d, { x: cx - seg / 2, y: ty - 0.6, w: seg, h: 0.38, fontSize: 16, bold: true, color: last ? C.accent1 : C.text2, align: "center" });
      tx(s, t, { x: cx - seg / 2 + 0.1, y: ty + 0.25, w: seg - 0.2, h: 0.75, fontSize: 14, color: C.text1, align: "center", valign: "top" });
    });

    // 하단: 고정노선 + 수요응답형
    const by = 3.85, bh = 2.5, w1 = 4.3;
    card(s, L, by, w1, bh);
    tx(s, "기존: 고정노선", { x: L + 0.3, y: by + 0.2, w: w1 - 0.6, h: 0.45, fontSize: 18, bold: true, color: C.text2 });
    tx(s, bl(["3개 고정노선", "노선별 4~7개 승하차 지점", "계속 운영(병행)"], { fontSize: 15, color: C.text1 }), { x: L + 0.3, y: by + 0.8, w: w1 - 0.6, h: 1.5, valign: "top", paraSpaceAfter: 6 });
    tx(s, "+", { x: L + w1, y: by, w: 0.6, h: bh, fontSize: 36, bold: true, color: C.accent1, align: "center", valign: "middle" });
    const x2 = L + w1 + 0.6, w2 = R - x2;
    card(s, x2, by, w2, bh, C.text2);
    tx(s, "2026년 9월~: 수요응답형(On-Demand)", { x: x2 + 0.3, y: by + 0.2, w: w2 - 0.6, h: 0.45, fontSize: 18, bold: true, color: C.background1 });
    tx(s, bl([
      "이용자가 지정된 승하차 지점 사이를 예약해 직접 이동",
      "Grab의 기존 Route 1·3 승하차 지점 활용",
      "주민 수요와 이용자 의견을 반영하여 승하차 지점 단계적 확대 추진",
    ], { fontSize: 15, color: C.background1 }), { x: x2 + 0.3, y: by + 0.8, w: w2 - 0.6, h: 1.55, valign: "top", paraSpaceAfter: 6 });
    source(s, "자료: 보고서 p.3–4(사업 개요, 운영성과 및 서비스 개선)를 바탕으로 발표자 재구성. 타임라인은 원문에 명시된 날짜만 사용");
    s.addNotes("타임라인을 보면 2025년 9월 20일 도입계획 발표, 2026년 4월 1일 Grab Route 1·3 기준 일반인 무료 탑승 개시, 7월 12일 기준 누적 고유 이용자 11,500명 이상을 기록했습니다. 이후 이용자 의견을 반영해 2026년 9월부터 기존 고정노선과 병행하는 수요응답형 서비스로 확대했습니다. 이용자가 지정된 승하차 지점 사이를 예약해 직접 이동하는 방식이며, Grab의 기존 Route 1·3 승하차 지점을 활용하고 승하차 지점을 단계적으로 확대할 계획입니다.");
  }

  // ======================= S14 비교표 =======================
  pres.addSection({ title: "Ⅳ. 비교 및 시사점" });
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Ⅳ. 비교 및 시사점" });
    kicker(s, "Ⅳ. 국내외 사례 비교 및 시사점");
    s.addText("국내외 자율주행 대중교통 정책 사례 비교", { placeholder: "title" });
    message(s, "기존 대중교통체계의 한계를 보완한다는 점은 공통, 도입 목적·차량 형태·운영방식은 상이");

    const k = (t, hi) => ({ text: t, options: { bold: true, color: C.text2, fill: { color: hi ? C.accent6 : C.background2 }, align: "center" } });
    const v = (t, hi) => ({ text: t, options: { fill: { color: hi ? C.accent6 : C.background1 }, bold: !!hi } });
    const data = [
      ["해결할 도시 문제", "새벽 시간대 교통 공백 및 첫차 혼잡", "First/Last Mile 연결성 부족"],
      ["도입 목적", "새벽 근로자 이동 지원 및 교통서비스 공급 확대", "지역 내 이동 편의성 및 대중교통 접근성 개선"],
      ["차량 형태", "대형 자율주행버스", "소형 자율주행 셔틀", true],
      ["운영 방식", "기존 시내버스 노선 기반 조기 운행", "고정노선 도입 후 수요응답형 서비스 확대", true],
      ["주요 이용 대상", "새벽 출근 근로자 등", "지역 주민 및 주요 생활시설, 교통거점 이용자"],
      ["성과평가", "이용객 수, 만족도, 이용목적, 사고 발생 여부", "고유 이용자 수, 체감 안전성, 추천 의향, 이용자 의견"],
      ["주요 성과", "누적 탑승 인원 약 27,600명 이용, 만족도 4.08/5점", "고유 이용자 11,500명 이상 이용, 설문 응답자 99% 긍정 평가"],
      ["향후 발전 방향", "새벽 혼잡노선 확대 및 운영체계 개선", "수요응답형 운행 및 승하차 지점 확대"],
    ];
    const rows = [[hdrCell("구분"), hdrCell("서울 새벽 동행 자율주행버스(A160)"), hdrCell("싱가포르 Punggol 자율주행 셔틀", { fill: { color: C.accent1 } })],
      ...data.map(([a, b, c, hi]) => [k(a, hi), v(b, hi), v(c, hi)])];
    tx(s, "<표 3> 국내외 자율주행 대중교통 정책 사례 비교", { x: L, y: TOP - 0.1, w: 8, h: 0.33, fontSize: 13, bold: true, color: C.text2 });
    s.addTable(rows, { x: L, y: TOP + 0.28, w: CW, colW: [2.1, 5.0, CW - 7.1], rowH: [0.4, ...Array(8).fill(0.47)], fontSize: 14, color: C.text1,
      border: { type: "solid", pt: 0.75, color: HEX.accent3 }, valign: "middle", margin: [0.02, 0.12, 0.02, 0.12], objectName: "compare-table" });
    source(s, "자료: 서울특별시(2026.3.29.), LTA(2025.9.20., 2026.7.29.), Grab(2026.10.5.)을 바탕으로 작성 (보고서 <표 3>, p.5). 음영 강조는 발표자 표시");
    s.addNotes("두 사례 모두 자율주행 기술로 기존 대중교통체계의 한계를 보완한다는 공통점이 있지만, 지역별 교통 문제에 따라 도입 목적, 차량 형태, 운영방식에서 차이를 보입니다. 서울은 대형 자율주행버스로 기존 시내버스 노선을 기반으로 조기 운행하고, 싱가포르는 소형 셔틀로 고정노선을 도입한 뒤 수요응답형 서비스로 확대했습니다. 성과평가 역시 서울은 이용객 수·만족도·이용목적·사고 여부, 싱가포르는 고유 이용자 수·체감 안전성·추천 의향·이용자 의견을 활용했습니다.");
  }

  // ======================= S15 시사점 =======================
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Ⅳ. 비교 및 시사점" });
    kicker(s, "Ⅳ. 국내외 사례 비교 및 시사점");
    s.addText("국내외 사례 비교를 통한 시사점", { placeholder: "title" });
    message(s, "운영방식 선정, 이용자 중심 개선·연계, 종합적 성과평가의 세 측면");

    const gap = 0.3, w = (CW - 2 * gap) / 3, y = TOP, h = 4.45;
    const cols = [
      ["지역별 교통 문제에 따른\n운영방식 차별화",
        [["서울", "새벽시간대 대중교통 공급 부족 → 기존 시내버스 노선 활용"], ["싱가포르", "거점–생활시설 연결성 개선 → 소형 셔틀 도입, 수요응답형 확대"]],
        "지역별 교통수요와 서비스 목적을 고려한 운영방식 선정 필요"],
      ["이용자 중심 서비스 개선 및\n기존 대중교통과의 연계",
        [["싱가포르", "이용자 의견을 반영하여 고정노선 중심의 운영방식을 개선"]],
        "기존 교통수단과의 연계, 접근성 및 서비스 편의성을 종합적으로 고려"],
      ["운영성과와 이용자 수용성을\n종합적으로 고려한 평가",
        [["공통", "이용실적뿐 아니라 만족도, 체감 안전성 등 이용자 평가 활용"]],
        "이동시간, 대기시간, 안전성, 이용자 수용성 등의 지표를 종합 검토"],
    ];
    cols.forEach(([head, ev, imp], i) => {
      const x = L + i * (w + gap);
      card(s, x, y, w, h);
      numCircle(s, x + 0.25, y + 0.25, i + 1, 0.5);
      tx(s, head, { x: x + 0.9, y: y + 0.15, w: w - 1.05, h: 0.75, fontSize: 16, bold: true, color: C.text2, valign: "middle" });
      tx(s, "사례 근거", { x: x + 0.25, y: y + 1.05, w: w - 0.5, h: 0.3, fontSize: 12, bold: true, color: C.accent1 });
      const runs = [];
      ev.forEach(([tag, t], j) => {
        runs.push({ text: tag + "  ", options: { bold: true, color: C.text2 } });
        runs.push({ text: t, options: { color: C.text1, breakLine: j < ev.length - 1 } });
      });
      tx(s, runs, { x: x + 0.25, y: y + 1.38, w: w - 0.5, h: 1.6, fontSize: 14, valign: "top", paraSpaceAfter: 8 });
      const iy = y + h - 1.3;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.2, y: iy, w: w - 0.4, h: 1.1, rectRadius: 0.06, fill: { color: C.text2 }, line: { color: C.text2, width: 0 }, objectName: nm("implication") });
      tx(s, imp, { x: x + 0.4, y: iy, w: w - 0.8, h: 1.1, fontSize: 14, bold: true, color: C.background1, valign: "middle" });
    });
    source(s, "자료: 보고서 p.5–6(국내외 사례 비교를 통한 시사점)을 바탕으로 발표자 재구성. 하단 남색 상자 = 원문의 시사점");
    s.addNotes("시사점은 세 가지입니다. 첫째, 서울은 기존 시내버스 노선을, 싱가포르는 소형 셔틀과 수요응답형 서비스를 활용했듯이 지역별 교통수요와 서비스 목적에 맞는 운영방식 선정이 필요합니다. 둘째, 싱가포르처럼 이용자 의견을 반영한 개선과 기존 교통수단과의 연계, 접근성, 편의성을 함께 고려해야 합니다. 셋째, 두 사례 모두 이용실적과 함께 만족도·체감 안전성 같은 이용자 평가를 활용했으므로, 향후에는 이동시간, 대기시간, 안전성, 이용자 수용성 지표를 종합적으로 검토할 필요가 있습니다.");
  }

  // ======================= S16 향후 발전 방향 =======================
  pres.addSection({ title: "Ⅴ. 발전 방향 및 기말 계획" });
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Ⅴ. 발전 방향 및 기말 계획" });
    kicker(s, "Ⅴ. 향후 발전 방향 및 기말 발표 계획");
    s.addText("향후 발전 방향: 교통 공백 특성에 따른 운영방식", { placeholder: "title" });
    message(s, "교통 공백의 발생 특성에 따라 자율주행 대중교통의 운영방식이 달라질 수 있음을 확인");

    const rows = [
      [I.clock, "시간적 교통 공백", "서울", "새벽 시간대 교통수요 대응", "기존 노선 기반 운행"],
      [I.map, "공간적 교통 공백", "싱가포르", "지역 내 접근성 개선", "셔틀 운행 및 수요응답형 서비스 확대"],
    ];
    const rh = 1.15, rg = 0.25, c1 = 3.6, c2 = 3.6;
    rows.forEach(([ic, gapT, city, need, op], i) => {
      const y = TOP + 0.4 + i * (rh + rg);
      card(s, L, y, c1, rh);
      iconBadge(s, ic, L + 0.2, y + (rh - 0.6) / 2, 0.6, C.background1);
      tx(s, [{ text: gapT, options: { fontSize: 17, bold: true, color: C.text2, breakLine: true } }, { text: city + " 사례", options: { fontSize: 14, color: C.accent5 } }],
        { x: L + 1.0, y, w: c1 - 1.1, h: rh, valign: "middle", paraSpaceAfter: 2 });
      arrowRight(s, L + c1 + 0.12, y + rh / 2 - 0.16, 0.32, 0.32);
      const x2 = L + c1 + 0.55;
      card(s, x2, y, c2, rh, C.background1, C.accent3);
      tx(s, need, { x: x2 + 0.25, y, w: c2 - 0.5, h: rh, fontSize: 16, color: C.text1, valign: "middle" });
      arrowRight(s, x2 + c2 + 0.12, y + rh / 2 - 0.16, 0.32, 0.32);
      const x3 = x2 + c2 + 0.55, w3 = R - x3;
      card(s, x3, y, w3, rh, C.text2);
      tx(s, op, { x: x3 + 0.3, y, w: w3 - 0.6, h: rh, fontSize: 17, bold: true, color: C.background1, valign: "middle" });
    });
    // 헤더 라벨
    const hy = TOP;
    [["교통 공백 유형", L, c1], ["대응 목적", L + c1 + 0.55, c2], ["운영방식", L + c1 + 0.55 + c2 + 0.55, R - (L + c1 + 0.55 + c2 + 0.55)]].forEach(([t, x, w]) =>
      tx(s, t, { x, y: hy, w, h: 0.33, fontSize: 13, bold: true, color: C.accent1 }));

    const by = TOP + 0.4 + 2 * rh + rg + 0.45;
    arrowDown(s, L + CW / 2 - 0.2, by - 0.38, 0.4, 0.28, C.text2);
    card(s, L, by, CW, 0.95, C.background1, C.text2);
    tx(s, [{ text: "시사점  ", options: { bold: true, color: C.accent1 } }, { text: "지역별 교통수요 및 도로·교통환경을 고려한 적정 운영방식 선정 필요", options: { bold: true, color: C.text2 } }],
      { x: L + 0.35, y: by, w: CW - 0.7, h: 0.95, fontSize: 19, valign: "middle", align: "center" });
    source(s, "자료: 보고서 p.6(교통 공백 특성에 따른 적정 운영방식 선정 필요)을 바탕으로 발표자 재구성");
    s.addNotes("서울과 싱가포르 사례를 통해 교통 공백의 발생 특성에 따라 운영방식이 달라질 수 있음을 확인했습니다. 서울은 새벽 시간대 교통수요에 대응하기 위해 기존 노선 기반으로 운행했고, 싱가포르는 지역 내 접근성 개선을 위해 셔틀을 운행하고 수요응답형 서비스로 확대했습니다. 따라서 지역별 교통수요와 도로·교통환경을 고려한 적정 운영방식 선정이 필요합니다.");
  }

  // ======================= S17 기말 연구주제 및 수행 계획 =======================
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Ⅴ. 발전 방향 및 기말 계획" });
    kicker(s, "Ⅴ. 향후 발전 방향 및 기말 발표 계획");
    s.addText("기말 발표 연구주제 및 수행 계획", { placeholder: "title" });

    card(s, L, 1.35, CW, 1.6, C.text2);
    tx(s, "연구주제", { x: L + 0.35, y: 1.47, w: 3, h: 0.32, fontSize: 13, bold: true, color: C.accent3 });
    tx(s, "교통 공백 해소를 위한 자율주행 대중교통의 고정노선형·수요응답형 운영방식 비교 및 적용방안 연구", { x: L + 0.35, y: 1.8, w: CW - 0.7, h: 0.72, fontSize: 18, bold: true, color: C.background1, valign: "middle" });
    tx(s, "운영방식별 특성 및 적용조건 비교  ·  교통수요 특성에 따른 적정 운영방안 제안", { x: L + 0.35, y: 2.53, w: CW - 0.7, h: 0.32, fontSize: 13, color: C.accent6 });

    const steps = [
      ["선행연구 검토", "교통 공백 유형 및 자율주행 대중교통 운영방식", ["시간적 공백: 새벽·심야 등 대중교통 공급이 부족한 시간대", "공간적 공백: 대중교통 접근성이 낮거나 교통서비스 공급이 부족한 지역"]],
      ["운영 특성 비교", "고정노선형 vs 수요응답형 운영 특성 및 장단점", ["비교항목: 교통수요의 집중도·분산 정도, 이용자 대기시간, 이동 편의성, 운영 효율성"]],
      ["적용조건 검토", "교통 공백 유형 및 수요 특성에 따른 운영방식별 적용조건", []],
      ["적용방안 제시", "적정 운영방식 선정 기준과 적용방안", []],
    ];
    const gap = 0.3, w = (CW - 3 * gap) / 4, y = 3.15, h = 3.3;
    steps.forEach(([t, d, items], i) => {
      const x = L + i * (w + gap);
      card(s, x, y, w, h);
      numCircle(s, x + 0.22, y + 0.22, i + 1, 0.48, i === 3 ? C.accent1 : C.text2);
      tx(s, t, { x: x + 0.82, y: y + 0.22, w: w - 0.95, h: 0.48, fontSize: 16, bold: true, color: C.text2, valign: "middle" });
      tx(s, d, { x: x + 0.22, y: y + 0.82, w: w - 0.44, h: 0.8, fontSize: 14, bold: true, color: C.text1, valign: "top" });
      if (items.length) tx(s, bl(items, { fontSize: 13, color: C.text1 }), { x: x + 0.22, y: y + 1.62, w: w - 0.44, h: h - 1.72, valign: "top", paraSpaceAfter: 5 });
      if (i < 3) arrowRight(s, x + w + 0.04, y + 0.3, 0.22, 0.22, C.accent3);
    });
    source(s, "자료: 보고서 p.6(기말 발표 연구주제, 기말 발표 수행 계획)");
    s.addNotes("기말 발표에서는 교통 공백 해소를 위한 자율주행 대중교통의 고정노선형·수요응답형 운영방식 비교 및 적용방안을 연구합니다. 먼저 시간적·공간적 교통 공백 유형과 운영방식에 관한 선행연구를 검토하고, 수요의 집중도와 분산 정도, 대기시간, 이동 편의성, 운영 효율성을 기준으로 두 운영방식의 특성과 장단점을 비교합니다. 이어 공백 유형과 수요 특성에 따른 운영방식별 적용조건을 검토하고, 적정 운영방식 선정 기준과 적용방안을 제시하겠습니다.");
  }

  // ======================= S18 참고문헌 =======================
  pres.addSection({ title: "참고문헌" });
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "참고문헌" });
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
      runs.push({ text: u, options: { fontSize: 12, color: C.accent1, hyperlink: { url: u }, breakLine: i < refs.length - 1, paraSpaceAfter: 16 } });
    });
    tx(s, runs, { x: L, y: 1.45, w: CW, h: 5.35, valign: "top", paraSpaceAfter: 1 });
    s.addNotes("보고서 p.7의 참고문헌 7건입니다. URL의 utm 추적 파라미터만 제거했습니다.");
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
}

build().catch((e) => { console.error(e); process.exit(1); });
