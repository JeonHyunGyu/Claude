const pptxgen = require("pptxgenjs");
const { applyTheme } = require("/root/.claude/skills/synced/2709c09c-c0dc-4f81-b385-fcb3527e2d18_99de5be8-b623-4241-b2e1-4194ab58e11a/pptx/scripts/apply_theme.js");

const OUT = process.argv[2] || "deck.pptx";

const THEME = {
  name: "UOS Traffic Capacity",
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
pres.title = "고속도로 기본구간 및 신호교차로 분석 대상 선정과 조사 계획";
pres.subject = "교통 용량 분석 과제 중간발표";
pres.company = "서울시립대학교";
const C = pres.SchemeColor;

const FOOTER = "서울시립대학교  |  교통 용량 분석 과제 중간발표";

// ---------- Layouts ----------
pres.defineSlideMaster({
  title: "COVER",
  background: { color: C.background1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.95, w: 12.1, h: 1.45,
      fontSize: 34, bold: true, color: C.text2, valign: "top", align: "left", margin: 0 }, text: "" } },
    { text: { text: FOOTER, options: { x: 0.6, y: 7.0, w: 8, h: 0.3, fontSize: 10, color: C.accent5, margin: 0 } } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: C.background1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.4, w: 12.1, h: 0.7,
      fontSize: 30, bold: true, color: C.text2, valign: "middle", align: "left", margin: 0 }, text: "" } },
    { text: { text: FOOTER, options: { x: 0.6, y: 7.0, w: 8, h: 0.3, fontSize: 10, color: C.accent5, margin: 0 } } },
  ],
  slideNumber: { x: 12.23, y: 7.0, w: 0.5, h: 0.3, fontSize: 10, color: C.accent5, align: "right" },
});

// ---------- Helpers ----------
let uid = 0;
const nm = (s) => `${s}-${++uid}`;

function card(s, x, y, w, h, fill = C.background2) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: fill }, line: { color: fill, width: 0 }, objectName: nm("card") });
}
// kind: "ok" (확인/확보), "todo" (추가 확인), "info" (neutral)
function chip(s, x, y, text, kind = "ok", w) {
  const width = w || 0.3 + text.length * 0.14;
  const styles = {
    ok:   { fill: { color: C.accent1 }, line: { color: C.accent1, width: 1 }, color: C.background1 },
    todo: { fill: { color: C.background1 }, line: { color: C.accent4, width: 1.25 }, color: C.accent4 },
    info: { fill: { color: C.accent6 }, line: { color: C.accent6, width: 1 }, color: C.text2 },
  }[kind];
  s.addText(text, { x, y, w: width, h: 0.3, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.15,
    fill: styles.fill, line: styles.line, color: styles.color, fontSize: 11, bold: true, align: "center",
    valign: "middle", margin: 0, isTextBox: true, objectName: nm("chip") });
  return width;
}
function imageSlot(s, x, y, w, h, title, lines) {
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: C.background2 },
    line: { color: C.accent3, width: 1.5, dashType: "dash" }, objectName: nm("image-slot") });
  const runs = [{ text: title, options: { fontSize: 16, bold: true, color: C.accent1, breakLine: true } }];
  if (lines && lines.length) {
    runs.push({ text: " ", options: { fontSize: 8, breakLine: true } });
    lines.forEach((l, i) => runs.push({ text: l, options: { fontSize: 12, color: C.accent5, breakLine: i < lines.length - 1 } }));
  }
  s.addText(runs, { x: x + 0.2, y: y + 0.2, w: w - 0.4, h: h - 0.4, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: nm("image-slot-label") });
}
function notes(s, y, lines, h) {
  s.addText(lines.map((l, i) => ({ text: l, options: { breakLine: i < lines.length - 1 } })),
    { x: 0.6, y, w: 12.1, h: h || 0.3 * lines.length, fontSize: 12, color: C.accent5, margin: 0, valign: "top", isTextBox: true, paraSpaceAfter: 2, objectName: nm("notes") });
}
function source(s, text, y = 6.62) {
  s.addText(text, { x: 0.6, y, w: 12.1, h: 0.28, fontSize: 10, italic: true, color: C.accent5, margin: 0, isTextBox: true, objectName: nm("source") });
}
function sectionHead(s, x, y, w, text) {
  s.addText(text, { x, y, w, h: 0.4, fontSize: 16, bold: true, color: C.text2, margin: 0, valign: "middle", isTextBox: true, objectName: nm("head") });
}
const bullets = (items, opts = {}) => items.map((t, i) => ({ text: t, options: { bullet: { indent: 14 }, breakLine: i < items.length - 1, ...opts } }));

// ---------- v2 helpers ----------
// Amber run marking a value the team still has to supply
const need = (t, extra = {}) => ({ text: `[${t}]`, options: { color: C.accent4, bold: true, ...extra } });
const plain = (t, extra = {}) => ({ text: t, options: { color: C.text1, ...extra } });
const MAP_IMG = process.env.MAP_IMG || require("path").join(__dirname, "map_highway.png");

// ======================= Slide 1 =======================
pres.addSection({ title: "개요" });
{
  const s = pres.addSlide({ masterName: "COVER", sectionTitle: "개요" });
  s.addText("서울시립대학교  ·  교통 용량 분석 과제 중간발표", { x: 0.6, y: 0.45, w: 8, h: 0.35, fontSize: 14, bold: true, color: C.accent1, margin: 0, isTextBox: true });
  s.addText("고속도로 기본구간 및 신호교차로\n분석 대상 선정과 조사 계획", { placeholder: "title" });

  const ty = 2.7, th = 1.45, tw = 3.6;
  [[0.6, "고속도로 기본구간", "수도권제1순환고속도로", "성남TG → 송파IC 방향"],
   [0.6 + tw + 0.3, "신호교차로", "태릉입구역 교차로", "화랑로 × 동일로 사거리"]].forEach(([x, tag, main, sub]) => {
    card(s, x, ty, tw, th);
    s.addText([
      { text: tag, options: { fontSize: 12, bold: true, color: C.accent1, breakLine: true } },
      { text: main, options: { fontSize: 18, bold: true, color: C.text2, breakLine: true } },
      { text: sub, options: { fontSize: 14, color: C.text1 } },
    ], { x: x + 0.3, y: ty + 0.15, w: tw - 0.5, h: th - 0.3, margin: 0, valign: "middle", paraSpaceAfter: 4, isTextBox: true });
  });

  sectionHead(s, 0.6, 4.45, 4, "진행 순서");
  const steps = ["대상 선정", "자료 확보·\n예비 검토", "현장조사", "용량·서비스\n수준 분석"];
  const sw = 1.62, gap = 0.34, sy = 4.95, sh = 0.95;
  steps.forEach((t, i) => {
    const x = 0.6 + i * (sw + gap);
    const now = i < 2;
    s.addText([{ text: `${i + 1}  `, options: { bold: true } }, { text: t }], {
      x, y: sy, w: sw, h: sh, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.08,
      fill: { color: now ? C.text2 : C.background1 }, line: { color: now ? C.text2 : C.accent3, width: 1.5 },
      color: now ? C.background1 : C.text2, fontSize: 14, bold: true, align: "center", valign: "middle", margin: 2, isTextBox: true, objectName: nm("step") });
    if (i < steps.length - 1) {
      s.addShape(pres.shapes.ISOSCELES_TRIANGLE, { x: x + sw + 0.07, y: sy + sh / 2 - 0.1, w: 0.2, h: 0.2, rotate: 90,
        fill: { color: C.accent3 }, line: { color: C.accent3, width: 0 }, objectName: nm("arrow") });
    }
  });
  s.addText("중간발표: 대상 선정·자료 확보 현황·현장조사 계획  /  이후 진행: ③~④", { x: 0.6, y: 6.05, w: 7.5, h: 0.3, fontSize: 12, color: C.accent5, margin: 0, isTextBox: true });

  const rx = 8.4, rw = 4.33;
  sectionHead(s, rx, 2.65, rw, "팀원 및 업무분담");
  const hdr = (t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text2 }, align: "center" } });
  const rows = [[hdr("이름"), hdr("담당 업무")],
    [{ text: "전현규", options: { align: "center" } }, { text: "발표, 고속도로 자료 정리, 현장 신호 조사", options: {} }],
    [{ text: "송동석", options: { align: "center" } }, { text: "발표, 교차로 교통량 분석, 결과 취합", options: {} }]];
  s.addTable(rows, { x: rx, y: 3.15, w: rw, colW: [1.1, rw - 1.1], rowH: [0.48, 0.72, 0.72], fontSize: 12, color: C.text1,
    border: { type: "solid", pt: 0.75, color: HEX.accent3 }, valign: "middle", margin: [0.03, 0.08, 0.03, 0.12], objectName: "team-table" });
  s.addTable([[{ text: "발표일", options: { bold: true, color: C.text2, fill: { color: C.background2 }, align: "center" } }, { text: "2026.10.14", options: { align: "center" } }]],
    { x: rx, y: 5.42, w: rw, colW: [1.1, rw - 1.1], rowH: 0.48, fontSize: 13, color: C.text1,
      border: { type: "solid", pt: 0.75, color: HEX.accent3 }, valign: "middle", objectName: "date-table" });

  s.addNotes("고속도로 기본구간(수도권제1순환고속도로 성남TG→송파IC 방향)과 신호교차로(태릉입구역 교차로)를 대상으로 선정했습니다. 이번 중간발표에서는 대상 선정, 자료 확보 현황, 현장조사 계획을 다루며, 최종 용량·서비스수준 결과는 아직 없습니다.");
}

// ======================= Slide 2 =======================
pres.addSection({ title: "고속도로 기본구간" });
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "고속도로 기본구간" });
  s.addText("고속도로 기본구간 선정", { placeholder: "title" });

  // Map frame
  const fx = 0.6, fy = 1.35, fw = 7.0, fh = 4.6;
  s.addShape(pres.shapes.RECTANGLE, { x: fx, y: fy, w: fw, h: fh, fill: { color: C.background2 }, line: { color: C.accent6, width: 1 }, objectName: "map-frame" });
  const ih = fh, iw = ih * 412 / 1193, ix = 2.55, iy = fy; // native aspect, no stretch
  s.addImage({ path: MAP_IMG, x: ix, y: iy, w: iw, h: ih, altText: "성남TG→송파IC 방향 선택 구간 위성지도 (네이버지도)", objectName: "map-highway" });
  // Pin positions read off the image (fraction of width / height)
  const pin = (fxr, fyr) => [ix + iw * fxr, iy + ih * fyr];
  const [ex, ey] = pin(0.52, 0.16);  // 도착 (end)
  const [sx0, sy0] = pin(0.55, 0.826); // 출발 (start)
  const callout = (px, py, label, sub) => {
    s.addShape(pres.shapes.LINE, { x: px + 0.08, y: py, w: 4.6 - (px + 0.08), h: 0, line: { color: C.text2, width: 1 }, objectName: nm("leader") });
    s.addText([{ text: label, options: { bold: true, color: C.text2, breakLine: true } }, { text: sub, options: { color: C.text1 } }],
      { x: 4.6, y: py - 0.3, w: 2.85, h: 0.6, fontSize: 12, margin: [0.03, 0.08, 0.03, 0.08], valign: "middle",
        fill: { color: C.background1 }, line: { color: C.accent6, width: 0.75 }, isTextBox: true, objectName: nm("callout") });
  };
  callout(ex, ey, "종료점 (지도 ‘도착’)", "송파IC 쪽");
  callout(sx0, sy0, "시작점 (지도 ‘출발’)", "성남TG 쪽");
  // Direction arrow
  s.addShape(pres.shapes.LINE, { x: 4.85, y: 2.9, w: 0, h: 1.55, flipV: true, line: { color: C.accent1, width: 3, endArrowType: "triangle" }, objectName: "dir-arrow" });
  s.addText("진행방향\n성남TG → 송파IC", { x: 5.05, y: 3.35, w: 2.3, h: 0.6, fontSize: 12, bold: true, color: C.accent1, margin: 0, valign: "middle", isTextBox: true });
  // Left legend: items still to be drawn on the map
  s.addText([
    { text: "지도 추가 표기 예정", options: { bold: true, color: C.text2, breakLine: true } },
    { text: "VDS 00700 · 00800", options: { color: C.text1, breakLine: true } },
    { text: "(이정 기반 근사 위치)", options: { color: C.accent5, breakLine: true } },
    { text: "접속부 기준점·400m 경계", options: { color: C.text1 } },
  ], { x: 0.75, y: 1.5, w: 1.7, h: 1.4, fontSize: 11, margin: 0, valign: "top", isTextBox: true, objectName: "map-legend" });
  s.addText("지도: 네이버지도 위성사진", { x: 0.75, y: 5.6, w: 1.75, h: 0.25, fontSize: 10, italic: true, color: C.accent5, margin: 0, isTextBox: true });

  // Right cards
  const rx = 7.95, rw = 4.78, gap = 0.1;
  const cards = [
    { label: "대상 · 선정 이유", h: 0.86, body: [plain("성남TG → 송파IC 방향 선택 구간", { bold: true, breakLine: true }), plain("차로 수 일정 · 인접 VDS 자료 비교 가능")] },
    { label: "분석 경계 (잠정)", h: 0.86, body: [plain("접속부 기준점에서 각각 400m 이격", { breakLine: true }), need("기준점: 성남TG 측 ___ / 송파IC 측 ___")] },
    { label: "기본구간 판정 근거", h: 0.68, tag: ["입력 필요", "todo"], body: [need("수업 적용 기준 · TG·접속부 영향 검토 방법")] },
    { label: "차로 조건", h: 0.86, tag: ["확인", "ok"], body: [plain("선택 구간 내 편도 4차로 유지", { breakLine: true }), need("확인 자료")] },
    { label: "선택한 기본구간 길이", h: 0.68, tag: ["추가 확인", "todo"], body: [need("___ m")] },
  ];
  let y = 1.35;
  cards.forEach((c) => {
    card(s, rx, y, rw, c.h);
    s.addText(c.label, { x: rx + 0.25, y: y + 0.08, w: 2.8, h: 0.26, fontSize: 12, bold: true, color: C.accent1, margin: 0, isTextBox: true });
    if (c.tag) chip(s, rx + rw - 0.25 - (0.3 + c.tag[0].length * 0.14), y + 0.07, c.tag[0], c.tag[1]);
    s.addText(c.body, { x: rx + 0.25, y: y + 0.34, w: rw - 0.5, h: c.h - 0.4, fontSize: 13, margin: 0, valign: "top", isTextBox: true });
    y += c.h + gap;
  });

  notes(s, 6.1, [
    "※ 400m 이격만으로 접속부 영향이 완전히 제거됐다고 단정하지 않음",
    "※ VDS 지도 위치는 도로 이정 기반 근사 위치이며, 정확한 설치 위치는 추가 확인 예정",
  ]);
  s.addNotes("성남TG→송파IC 방향 중 편도 4차로가 유지되는 구간을 선택했습니다. 지도의 출발 지점이 시작점, 도착 지점이 종료점입니다. 접속부 기준점에서 각 400m 떨어진 지점을 잠정 경계로 두었으나, 이것만으로 접속부 영향이 완전히 제거됐다고 보지는 않습니다.");
}

// ======================= Slide 3 =======================
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "고속도로 기본구간" });
  s.addText("고속도로 입력자료 확보 현황", { placeholder: "title" });

  const H = (t, align = "left") => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text2 }, align } });
  const data = [
    ["대표 분석지점", "1000VDE00700, 이정 6.50km", "확보"],
    ["비교 지점", "1000VDE00800, 이정 7.20km", "확보"],
    ["교통량·속도", "2026년 6월 15·16·17·22·23일 5분 자료", "확보"],
    ["자료 일관성", "일 교통량 차이 약 1.9~3.7%\n시간대별 흐름 유사", "검토"],
    ["대표 지점 원시자료", "00700의 1~4차로 기록 확인\n5분 통계와 거의 일치", "확인"],
  ];
  const rows = [[H("항목"), H("현재 상태"), H("구분", "center")]];
  data.forEach(([a, b, c]) => rows.push([
    { text: a, options: { bold: true, color: C.text2 } }, { text: b, options: {} },
    { text: c, options: { bold: true, align: "center", color: C.accent1 } }]));
  s.addTable(rows, { x: 0.6, y: 1.35, w: 6.6, colW: [1.75, 3.95, 0.9], rowH: [0.42, 0.5, 0.5, 0.5, 0.7, 0.7], fontSize: 13, color: C.text1,
    border: { type: "solid", pt: 0.75, color: HEX.accent6 }, valign: "middle", margin: [0.03, 0.08, 0.03, 0.12], objectName: "hw-input-table" });

  // Evidence box under the table
  const ey = 4.8;
  card(s, 0.6, ey, 6.6, 1.15);
  s.addText([
    { text: "차이 계산 기준  ", options: { bold: true, color: C.accent1 } },
    need("분모 지점 · 동일 집계시간 · 차로 합계 범위", { breakLine: true }),
    { text: "근거 수치  ", options: { bold: true, color: C.accent1 } },
    need("시간대별 비교 그래프 또는 원시자료–5분 통계 차이(%)", { breakLine: true }),
    { text: "00800은 집계 차로 범위 미확인 → 비교용으로만 활용, 교통량 합산하지 않음", options: { color: C.text1 } },
  ], { x: 0.8, y: ey + 0.08, w: 6.25, h: 1.0, fontSize: 12, margin: 0, valign: "middle", paraSpaceAfter: 3, isTextBox: true, objectName: "evidence" });

  // Right: acquisition plan table
  const rx = 7.5, rw = 5.23;
  sectionHead(s, rx, 1.3, rw, "미확보 항목 보완 계획");
  const H2 = (t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.accent4 }, align: "center" } });
  const due = need("__.__");
  const plan = [
    ["차로폭 · 측방여유폭", "도로대장 · 설계자료"],
    ["종단경사", "도로대장 · 설계자료"],
    ["중차량 구성비", "대상 구간 직접 AVC 미확보\n→ 인근 AVC 지점 자료 또는 현장 표본계수로 대체(안), 적용 근거 명시"],
    ["VDS 설치 위치 ·\n00800 집계 차로", "제공기관(한국도로공사) 문의"],
  ];
  const prow = [[H2("항목"), H2("확보 경로"), H2("목표일")]];
  plan.forEach(([a, b]) => prow.push([{ text: a, options: { bold: true, color: C.text2 } }, { text: b, options: {} }, { text: due.text, options: { ...due.options, align: "center" } }]));
  s.addTable(prow, { x: rx, y: 1.8, w: rw, colW: [1.55, 2.78, 0.9], rowH: [0.42, 0.55, 0.55, 1.25, 0.75], fontSize: 12, color: C.text1,
    border: { type: "solid", pt: 0.75, color: HEX.accent6 }, valign: "middle", margin: [0.03, 0.06, 0.03, 0.1], objectName: "hw-plan-table" });

  source(s, "출처: 한국도로공사 VDS 5분 통계 및 차로별 원시자료", 6.15);
  s.addNotes("대표 분석지점은 00700, 비교 지점은 00800입니다. 두 지점의 5일간 하루 교통량 차이는 약 1.9~3.7%입니다. 00800은 집계 차로 범위가 확인되지 않아 비교용으로만 쓰고 교통량을 합산하지 않습니다. 차로폭·측방여유폭·종단경사·중차량 구성비는 오른쪽 경로로 보완할 계획입니다.");
}

// ======================= Slide 4 =======================
pres.addSection({ title: "신호교차로" });
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "신호교차로" });
  s.addText("태릉입구역 교차로 선정 및 차로 구성", { placeholder: "title" });

  // Approach schematic (replace with satellite map + arrows when available)
  const mx = 0.6, my = 1.35, mw = 4.0, mh = 4.6;
  s.addShape(pres.shapes.RECTANGLE, { x: mx, y: my, w: mw, h: mh, fill: { color: C.background2 }, line: { color: C.accent3, width: 1.5, dashType: "dash" }, objectName: "map-slot" });
  const cx = mx + mw / 2, cy = my + mh / 2, road = 0.9;
  s.addShape(pres.shapes.RECTANGLE, { x: cx - road / 2, y: my + 0.5, w: road, h: mh - 1.0, fill: { color: C.accent6 }, line: { color: C.accent6, width: 0 }, objectName: "road-ns" });
  s.addShape(pres.shapes.RECTANGLE, { x: mx + 0.15, y: cy - road / 2, w: mw - 0.3, h: road, fill: { color: C.accent6 }, line: { color: C.accent6, width: 0 }, objectName: "road-ew" });
  // Inbound arrows (right-hand traffic): from E on north half, from W on south half, from N on west half, from S on east half
  const arr = (x, y, w, h, flipH, flipV) => s.addShape(pres.shapes.LINE, { x, y, w, h, flipH, flipV, line: { color: C.text2, width: 2.5, endArrowType: "triangle" }, objectName: nm("inbound") });
  const q = road / 4;
  arr(cx + road / 2 + 0.15, cy - q, 1.1, 0, true, false);   // from east, heading west
  arr(cx - road / 2 - 1.25, cy + q, 1.1, 0, false, false);  // from west, heading east
  arr(cx - q, cy - road / 2 - 1.25, 0, 1.1, false, false);  // from north, heading south
  arr(cx + q, cy + road / 2 + 0.15, 0, 1.1, false, true);   // from south, heading north
  const lab = (t, x, y, w, align) => s.addText(t, { x, y, w, h: 0.42, fontSize: 11, color: C.text2, bold: true, align, margin: 0, valign: "middle", isTextBox: true });
  lab("북측 · 동일로 (공릉역 쪽)", mx + 0.1, my + 0.06, mw - 0.2, "center");
  lab("남측 · 동일로 (먹골역 쪽)", mx + 0.1, my + mh - 0.48, mw - 0.2, "center");
  lab("서측 · 화랑로\n(석계역 쪽)", mx + 0.15, cy + road / 2 + 0.1, 1.4, "left");
  lab("동측 · 화랑로", mx + mw - 1.55, cy - road / 2 - 0.5, 1.4, "right");

  // Location + reason (one line)
  const rx = 4.9, rw = 7.83;
  s.addText([{ text: "위치  ", options: { bold: true, color: C.accent1 } }, plain("화랑로 × 동일로 사거리      "),
    { text: "선정 이유  ", options: { bold: true, color: C.accent1 } }, plain("T-DATA 자료 · 현장 접근 · 회전별 분석 가능")],
    { x: rx, y: 1.35, w: rw, h: 0.4, fontSize: 13, margin: 0, valign: "middle", isTextBox: true });

  const H = (t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text2 }, align: "center" } });
  const lane = (t) => ({ text: t, options: { align: "center", color: C.text1, fill: { color: t.includes("좌회전") ? C.background2 : C.background1 } } });
  const rows = [[H("접근로 (진입방향)"), H("차로"), H("1 (중앙선)"), H("2"), H("3"), H("4"), H("5")]];
  [
    ["동측 · 화랑로", "동쪽에서 진입", 5, ["좌회전·유턴", "좌회전", "직진", "직진", "직진·우회전"]],
    ["서측 · 화랑로", "서쪽에서 진입", 5, ["좌회전·유턴", "직진", "직진", "직진", "우회전 전용"]],
    ["북측 · 동일로", "북쪽에서 진입", 4, ["좌회전·유턴", "직진", "직진", "직진·우회전", ""]],
    ["남측 · 동일로", "남쪽에서 진입", 5, ["좌회전·유턴", "좌회전", "직진", "직진", "직진·우회전"]],
  ].forEach(([a, d, n, ls]) => rows.push([
    { text: [{ text: a, options: { bold: true, color: C.text2, breakLine: true } }, { text: d, options: { color: C.accent5, fontSize: 11 } }] },
    { text: String(n), options: { align: "center", bold: true, color: C.text2 } },
    ...ls.map((t) => (t ? lane(t) : { text: "–", options: { align: "center", color: C.accent5 } })),
  ]));
  s.addTable(rows, { x: rx, y: 1.95, w: rw, colW: [1.78, 0.6, 1.09, 1.09, 1.09, 1.09, 1.09], rowH: [0.45, 0.78, 0.78, 0.78, 0.78],
    fontSize: 13, color: C.text1, border: { type: "solid", pt: 0.75, color: HEX.accent6 }, valign: "middle", margin: [0.03, 0.04, 0.03, 0.08], objectName: "lane-table" });
  s.addText("차로 구성은 중앙선 쪽부터 표시  ·  음영: 좌회전 포함 차로  ·  방향 표기는 5장 그래프와 동일", { x: rx, y: 5.6, w: rw, h: 0.25, fontSize: 11, color: C.accent5, margin: 0, isTextBox: true });

  notes(s, 6.0, [
    "※ 지도·거리뷰로 잠정 확인한 구성이며, 현장에서 대조 예정",
    "※ 버스전용차로(위치·운영시간·적용 범위)·유턴 조건은 현장 표지판으로 확인하고, 차로군 구성에 주는 영향까지 검토",
  ]);
  source(s, "출처: 네이버지도 위성사진·거리뷰", 6.68);
  s.addNotes("접근로는 교차로 기준 도로의 위치로 부르고, 괄호 안에 차량이 들어오는 방향을 적었습니다. 예를 들어 동측 접근로는 동쪽에서 진입해 서쪽으로 진행하는 차량입니다. 차로 구성은 지도·거리뷰로 잠정 확인했으며, 버스전용차로와 유턴 조건이 차로군 구성에 주는 영향은 현장에서 확인합니다.");
}

// ======================= Slide 5 =======================
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "신호교차로" });
  s.addText("교차로 교통량 예비 분석", { placeholder: "title" });

  card(s, 0.6, 1.3, 4.0, 1.05);
  s.addText([
    { text: "자료", options: { fontSize: 11, bold: true, color: C.accent1, breakLine: true } },
    { text: "서울시 T-DATA 스마트교차로 5분 교통량", options: { breakLine: true } },
    { text: "태릉입구역(ID 172) · 2026.9.29" },
  ], { x: 0.78, y: 1.36, w: 3.7, h: 0.93, fontSize: 12, color: C.text1, margin: 0, valign: "middle", isTextBox: true });
  card(s, 4.8, 1.3, 7.93, 1.05);
  s.addText([
    { text: "첨두 1시간 산정 방법", options: { fontSize: 11, bold: true, color: C.accent1, breakLine: true } },
    { text: "회전별(좌·직·우·유턴) 5분 VOL 합산 → 연속 12개(1시간) 합계가 최대인 구간 선정", options: { breakLine: true } },
    { text: "검토 시간범위 " }, need("오전 __~__ / 오후 __~__"), { text: "  ·  시각 표기 " }, need("원자료 시각 = 구간 시작/종료 확인"),
  ], { x: 4.98, y: 1.36, w: 7.6, h: 0.93, fontSize: 12, color: C.text1, margin: 0, valign: "middle", isTextBox: true });

  const sx = 0.6, sw = 4.0, sh = 1.3;
  [["오전 첨두 1시간", "07:25 ~ 08:25", "5,804", false, 2.55], ["오후 첨두 1시간", "17:20 ~ 18:20", "6,169", true, 4.0]].forEach(([k, t, v, hi, y]) => {
    card(s, sx, y, sw, sh, hi ? C.text2 : C.background2);
    s.addText([{ text: k, options: { fontSize: 13, bold: true, color: hi ? C.accent6 : C.accent1, breakLine: true } }, { text: t, options: { fontSize: 14, color: hi ? C.background1 : C.text1 } }],
      { x: sx + 0.25, y: y + 0.15, w: 1.9, h: sh - 0.3, margin: 0, valign: "middle", isTextBox: true });
    s.addText([{ text: v, options: { fontSize: 34, bold: true } }, { text: " 대", options: { fontSize: 14 } }],
      { x: sx + 2.0, y: y + 0.15, w: sw - 2.2, h: sh - 0.3, align: "right", margin: 0, valign: "middle", color: hi ? C.background1 : C.text2, isTextBox: true });
  });
  s.addText("전체 진입 교통량 기준", { x: sx, y: 5.35, w: sw, h: 0.25, fontSize: 11, color: C.accent5, margin: 0, isTextBox: true });

  s.addChart(pres.charts.BAR, [{ name: "진입 교통량", labels: ["서측 (서쪽에서 진입)", "동측 (동쪽에서 진입)", "남측 (남쪽에서 진입)", "북측 (북쪽에서 진입)"], values: [1992, 1755, 1295, 1127] }], {
    x: 4.95, y: 2.5, w: 7.78, h: 3.1, barDir: "col", barGapWidthPct: 70,
    chartColors: [HEX.dk2, HEX.accent3, HEX.accent3, HEX.accent3],
    showTitle: true, title: "오후 첨두 1시간 접근로별 진입 교통량 (대)", titleFontSize: 13, titleColor: HEX.dk2, titleFontFace: "+mn-lt",
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "#,##0", dataLabelFontSize: 12, dataLabelColor: HEX.dk1, dataLabelFontFace: "+mn-lt",
    catAxisLabelColor: HEX.dk1, catAxisLabelFontSize: 11, catAxisLabelFontFace: "+mn-lt",
    valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: 2400,
    valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false,
    catAxisLineColor: HEX.accent3, objectName: "dir-volume-chart",
  });

  card(s, 0.6, 5.7, 12.13, 0.92);
  s.addText([
    { text: "해석  ", options: { bold: true, color: C.accent1 } }, { text: "해당 날짜는 오후 첨두 교통량이 더 많고, 서측 접근로 진입 교통량이 가장 많음", options: { breakLine: true } },
    { text: "향후  ", options: { bold: true, color: C.accent1 } }, { text: "전체 진입량만으로 서비스수준을 판단하지 않음 → 접근로·회전별 교통량과 현장 신호시간을 연결해 차로군별 분석", options: { breakLine: true } },
    { text: "※ 한 날짜의 예비 결과로, 일반적인 첨두시간으로 확정하지 않음", options: { color: C.accent4, bold: true } },
  ], { x: 0.8, y: 5.72, w: 11.8, h: 0.88, fontSize: 12, color: C.text1, margin: 0, valign: "middle", isTextBox: true, objectName: "interpretation" });
  source(s, "출처: 서울시 T-DATA, VW_SC_TRF_VOL_5M_20260929.csv", 6.68);
  s.addNotes("회전별 5분 교통량을 합산한 뒤, 검토 시간범위 안에서 연속 12개 구간의 합계가 가장 큰 1시간을 첨두시간으로 골랐습니다. 오후 17:20~18:20이 6,169대로 오전보다 많았고, 접근로별로는 서측이 1,992대로 가장 많았습니다. 네 접근로 합계는 6,169대로 전체 진입 교통량과 일치합니다. 한 날짜 결과이므로 일반적인 첨두시간으로 확정하지는 않으며, 서비스수준은 회전별 교통량과 신호시간을 연결해 차로군별로 분석할 예정입니다.");
}

// ======================= Slide 6 =======================
pres.addSection({ title: "향후 계획" });
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "향후 계획" });
  s.addText("현장조사 및 향후 분석 계획", { placeholder: "title" });
  const top = 1.35;

  // Col 1: what to observe
  const c1x = 0.6, c1w = 4.55, c1h = 4.55;
  card(s, c1x, top, c1w, c1h);
  sectionHead(s, c1x + 0.25, top + 0.12, c1w - 0.5, "교차로 현장조사 항목");
  s.addText(bullets([
    "차로 수·회전 구성 대조, 버스전용차로·유턴 표지 → 차로군 구성 영향 확인",
    "신호주기·현시 순서, 현시별 녹색·황색·전적색시간",
    "보행신호 시간·현시",
    "3~5주기 예비 관측 → 변동이 있거나 관측이 불완전하면 추가 관측",
    "조사 날짜·시간 기록, 사진 촬영",
  ]), { x: c1x + 0.25, y: top + 0.62, w: c1w - 0.45, h: c1h - 0.75, fontSize: 13, color: C.text1, margin: 0, valign: "top", paraSpaceAfter: 7, isTextBox: true });

  // Col 2: execution plan table (replaces photo slot)
  const c2x = 5.4, c2w = 4.0;
  sectionHead(s, c2x, top, c2w, "조사 실행계획 (예정)");
  const K = (t) => ({ text: t, options: { bold: true, color: C.text2, fill: { color: C.background2 } } });
  const ex = [
    [K("조사일"), { text: [plain("10월 말 "), need("날짜")] }],
    [K("시간대"), { text: [plain("평일 오후 첨두(17:20~18:20) 포함 "), need("관측 시각")] }],
    [K("관측 위치"), { text: [plain("교차로 모서리 보도 "), need("지점")] }],
    [K("담당"), { text: [plain("신호시간: 전현규", { breakLine: true }), plain("차로·표지·사진: 송동석 "), need("확인")] }],
    [K("기록 방식"), { text: "현시별 기록지 + 스톱워치, 사진·영상" }],
  ];
  s.addTable(ex, { x: c2x, y: top + 0.5, w: c2w, colW: [1.0, 3.0], rowH: [0.62, 0.78, 0.62, 0.78, 0.62], fontSize: 12, color: C.text1,
    border: { type: "solid", pt: 0.75, color: HEX.accent6 }, valign: "middle", margin: [0.03, 0.06, 0.03, 0.1], objectName: "survey-plan" });
  s.addText([{ text: "고속도로  ", options: { bold: true, color: C.accent1 } }, plain("공개자료·기관 문의로 도로 조건·중차량 구성비·VDS 설치정보 보완 (3장)")],
    { x: c2x, y: top + 4.0, w: c2w, h: 0.55, fontSize: 12, margin: 0, valign: "top", isTextBox: true });

  // Col 3: next analysis steps
  const c3x = 9.65, c3w = 3.08;
  sectionHead(s, c3x, top, c3w, "향후 분석");
  const steps = [
    "현장조사 결과와 기존 자료 대조",
    "분석 교통량(오후 첨두)과 관측 신호의 시간대·운영조건 일치 확인",
    "용량·서비스수준 산정 필요 항목 확인",
    "입력자료 확보 후\n용량·서비스수준 분석",
  ];
  steps.forEach((t, i) => {
    const y = top + 0.55 + i * 1.0;
    s.addText(String(i + 1), { x: c3x, y: y + 0.15, w: 0.42, h: 0.42, shape: pres.shapes.OVAL, fill: { color: i === 3 ? C.text2 : C.accent1 }, line: { color: i === 3 ? C.text2 : C.accent1, width: 0 },
      color: C.background1, fontSize: 13, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: nm("step-num") });
    s.addText(t, { x: c3x + 0.55, y, w: c3w - 0.55, h: 0.75, fontSize: 13, color: C.text1, bold: i === 1 || i === 3, margin: 0, valign: "middle", isTextBox: true });
  });

  notes(s, 6.05, [
    "※ 다른 시간대·주말에 관측한 신호시간은 오후 첨두와 신호운영계획이 같은지 확인한 뒤에만 적용",
    "※ 최종 분석에 필요한 차종 구성·보행량 등은 추가 확보 계획에 포함",
  ]);
  s.addNotes("현장조사는 10월 말로 예정되어 있습니다. 5장에서 확인한 오후 첨두와 시간대를 맞추기 위해 평일 오후 첨두를 포함해 관측하는 것이 원칙이고, 다른 시간대에 관측한 신호시간은 운영계획이 같은지 확인한 뒤에만 적용합니다. 신호는 3~5주기 예비 관측 후 변동이나 누락이 있으면 추가 관측합니다.");
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
