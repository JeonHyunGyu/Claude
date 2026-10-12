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

// ======================= Slide 1 =======================
pres.addSection({ title: "개요" });
{
  const s = pres.addSlide({ masterName: "COVER", sectionTitle: "개요" });
  s.addText("서울시립대학교  ·  교통 용량 분석 과제 중간발표", { x: 0.6, y: 0.45, w: 8, h: 0.35, fontSize: 14, bold: true, color: C.accent1, margin: 0, isTextBox: true });
  s.addText("고속도로 기본구간 및 신호교차로\n분석 대상 선정과 조사 계획", { placeholder: "title" });

  // Target cards
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

  // Process flow
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
  s.addText("중간발표 범위: ①~②  /  이후 진행: ③~④", { x: 0.6, y: 6.05, w: 7.5, h: 0.3, fontSize: 12, color: C.accent5, margin: 0, isTextBox: true });

  // Team table (blank)
  const rx = 8.4, rw = 4.33;
  sectionHead(s, rx, 2.7 - 0.05, rw, "팀원 및 업무분담");
  const hdr = (t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text2 }, align: "center" } });
  const blank = () => ({ text: "", options: {} });
  const rows = [[hdr("이름"), hdr("담당 업무")]];
  for (let i = 0; i < 4; i++) rows.push([blank(), blank()]);
  s.addTable(rows, { x: rx, y: 3.15, w: rw, colW: [1.35, rw - 1.35], rowH: 0.48, fontSize: 13, color: C.text1,
    border: { type: "solid", pt: 0.75, color: HEX.accent3 }, valign: "middle", objectName: "team-table" });
  s.addTable([[{ text: "발표일", options: { bold: true, color: C.text2, fill: { color: C.background2 }, align: "center" } }, blank()]],
    { x: rx, y: 5.75, w: rw, colW: [1.35, rw - 1.35], rowH: 0.48, fontSize: 13, color: C.text1,
      border: { type: "solid", pt: 0.75, color: HEX.accent3 }, valign: "middle", objectName: "date-table" });

  s.addNotes("발표 개요: 고속도로 기본구간(수도권제1순환고속도로 성남TG→송파IC 방향)과 신호교차로(태릉입구역 교차로)를 대상으로 선정했고, 이번 중간발표는 대상 선정과 자료 확보·예비 검토 단계까지를 다룹니다. 최종 용량·서비스수준 결과는 아직 없습니다.");
}

// ======================= Slide 2 =======================
pres.addSection({ title: "고속도로 기본구간" });
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "고속도로 기본구간" });
  s.addText("고속도로 기본구간 선정", { placeholder: "title" });

  imageSlot(s, 0.6, 1.35, 7.0, 4.6, "지도 삽입 위치", [
    "표시 항목",
    "분석 시작점 · 종료점",
    "진행방향 (성남TG → 송파IC)",
    "VDS 00700 · 00800 (이정 기반 근사 위치)",
  ]);

  const rx = 7.95, rw = 4.78;
  const items = [
    ["대상", "성남TG → 송파IC 방향의 선택 구간", null],
    ["분석 경계 (잠정)", "지도상 접속부에서 각각 400m 이격하여 설정", null],
    ["차로 조건", "선택 구간 내 편도 4차로 유지", ["확인", "ok"]],
    ["선정 이유", "차로 수 일정 · 인접 VDS 자료 비교·검토 가능", null],
    ["전체 구간 길이", "미확정", ["추가 확인", "todo"]],
  ];
  let y = 1.35;
  const ih = 0.84, ig = 0.1;
  items.forEach(([label, body, tag]) => {
    card(s, rx, y, rw, ih);
    s.addText(label, { x: rx + 0.25, y: y + 0.1, w: 2.6, h: 0.28, fontSize: 12, bold: true, color: C.accent1, margin: 0, isTextBox: true });
    if (tag) chip(s, rx + rw - 0.25 - (0.3 + tag[0].length * 0.14), y + 0.1, tag[0], tag[1]);
    s.addText(body, { x: rx + 0.25, y: y + 0.38, w: rw - 0.5, h: 0.4, fontSize: 14, color: C.text1, bold: label === "대상", margin: 0, valign: "top", isTextBox: true });
    y += ih + ig;
  });

  notes(s, 6.1, [
    "※ 400m 이격만으로 접속부 영향이 완전히 제거됐다고 단정하지 않음",
    "※ VDS 지도 위치는 도로 이정 기반 근사 위치이며, 정확한 설치 위치는 추가 확인 예정",
  ]);
  s.addNotes("성남TG→송파IC 방향 중 편도 4차로가 유지되는 구간을 선택했습니다. 접속부에서 각 400m 떨어진 지점을 잠정 경계로 두었으나, 이것만으로 접속부 영향이 완전히 제거됐다고 보지는 않습니다. 전체 구간 길이는 아직 확정하지 않았습니다.");
}

// ======================= Slide 3 =======================
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "고속도로 기본구간" });
  s.addText("고속도로 입력자료 확보 현황", { placeholder: "title" });

  const H = (t, align = "left") => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text2 }, align } });
  const st = (t, todo) => ({ text: t, options: { bold: true, align: "center", color: todo ? C.accent4 : C.accent1 } });
  const L = (t) => ({ text: t, options: { bold: true, color: C.text2 } });
  const data = [
    ["대표 분석지점", "1000VDE00700, 이정 6.50km", "확보"],
    ["비교 지점", "1000VDE00800, 이정 7.20km", "확보"],
    ["교통량·속도", "2026년 6월 15·16·17·22·23일 5분 자료", "확보"],
    ["자료 일관성", "두 지점 5일간 일 교통량 차이 약 1.9~3.7%, 시간대별 흐름 유사", "검토"],
    ["대표 지점 원시자료", "00700의 1~4차로 기록 확인, 5분 통계와 거의 일치", "확인"],
    ["추가 확보 항목", "차로폭 · 측방여유폭 · 종단경사 · 중차량 구성비", "추가 확보"],
    ["추가 확인 항목", "정확한 VDS 설치 위치, 00800의 집계 차로 범위", "추가 확인"],
  ];
  const rows = [[H("항목"), H("현재 상태"), H("구분", "center")]];
  data.forEach(([a, b, c]) => {
    const todo = c.startsWith("추가");
    const fill = todo ? { fill: { color: C.background2 } } : {};
    rows.push([{ text: a, options: { ...L(a).options, ...fill } }, { text: b, options: { ...fill } }, { text: c, options: { ...st(c, todo).options, ...fill } }]);
  });
  s.addTable(rows, { x: 0.6, y: 1.35, w: 8.4, colW: [1.95, 5.25, 1.2], rowH: 0.56, fontSize: 13, color: C.text1,
    border: { type: "solid", pt: 0.75, color: HEX.accent6 }, valign: "middle", margin: [0.03, 0.08, 0.03, 0.12], objectName: "hw-input-table" });

  // Right panel
  const rx = 9.3, rw = 3.43;
  card(s, rx, 1.35, rw, 4.5);
  sectionHead(s, rx + 0.25, 1.5, rw - 0.5, "자료 활용·보완 계획");
  s.addText(bullets([
    "인접 VDS 교통량은 합산하지 않고, 대표 지점 분석과 비교 검토에 활용",
    "중차량 구성비: 대상 구간 직접 AVC 자료 미확보",
    "도로 조건: 도로대장·설계자료로 보완",
    "설치정보: 제공기관 문의로 보완",
  ]), { x: rx + 0.25, y: 2.0, w: rw - 0.45, h: 3.7, fontSize: 14, color: C.text1, valign: "top", margin: 0, paraSpaceAfter: 10, isTextBox: true });

  // Legend
  let lx = 0.6;
  lx += chip(s, lx, 5.98, "확보·확인", "ok") + 0.15;
  s.addText("자료 확보 또는 검토 완료", { x: lx, y: 5.98, w: 2.4, h: 0.3, fontSize: 11, color: C.accent5, margin: 0, valign: "middle", isTextBox: true });
  lx += 2.5;
  lx += chip(s, lx, 5.98, "추가 확인", "todo") + 0.15;
  s.addText("미확보·미확인, 추후 보완", { x: lx, y: 5.98, w: 2.6, h: 0.3, fontSize: 11, color: C.accent5, margin: 0, valign: "middle", isTextBox: true });

  source(s, "출처: 한국도로공사 VDS 5분 통계 및 차로별 원시자료");
  s.addNotes("대표 분석지점은 00700, 비교 지점은 00800입니다. 두 지점의 5일간 하루 교통량 차이는 약 1.9~3.7%로 흐름이 유사했습니다. 차로폭·측방여유폭·종단경사·중차량 구성비와 정확한 VDS 설치 위치는 아직 확보·확인되지 않았습니다.");
}

// ======================= Slide 4 =======================
pres.addSection({ title: "신호교차로" });
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "신호교차로" });
  s.addText("태릉입구역 교차로 선정 및 차로 구성", { placeholder: "title" });

  // Approach schematic / map slot
  const mx = 0.6, my = 1.35, mw = 4.0, mh = 4.6;
  s.addShape(pres.shapes.RECTANGLE, { x: mx, y: my, w: mw, h: mh, fill: { color: C.background2 }, line: { color: C.accent3, width: 1.5, dashType: "dash" }, objectName: "map-slot" });
  const cx = mx + mw / 2, cy = my + mh / 2, road = 0.75;
  s.addShape(pres.shapes.RECTANGLE, { x: cx - road / 2, y: my + 0.55, w: road, h: mh - 1.1, fill: { color: C.accent6 }, line: { color: C.accent6, width: 0 }, objectName: "road-ns" });
  s.addShape(pres.shapes.RECTANGLE, { x: mx + 0.15, y: cy - road / 2, w: mw - 0.3, h: road, fill: { color: C.accent6 }, line: { color: C.accent6, width: 0 }, objectName: "road-ew" });
  const lab = (t, x, y, w, align) => s.addText(t, { x, y, w, h: 0.45, fontSize: 11, color: C.text2, bold: true, align, margin: 0, valign: "middle", isTextBox: true });
  lab("북 · 동일로 (공릉역 쪽)", mx + 0.1, my + 0.08, mw - 0.2, "center");
  lab("남 · 동일로 (먹골역 쪽)", mx + 0.1, my + mh - 0.53, mw - 0.2, "center");
  lab("서 · 화랑로\n(석계역 쪽)", mx + 0.15, cy - road / 2 - 0.55, 1.4, "left");
  lab("동 · 화랑로", mx + mw - 1.55, cy - road / 2 - 0.55, 1.4, "right");
  s.addText("지도 / 접근로 도식 삽입 위치", { x: mx + 0.15, y: cy + road / 2 + 0.15, w: mw - 0.3, h: 0.35, fontSize: 12, bold: true, color: C.accent1, align: "center", margin: 0, isTextBox: true });

  // Location + reasons
  const rx = 4.9, rw = 7.83;
  s.addText([{ text: "위치  ", options: { bold: true, color: C.accent1 } }, { text: "화랑로와 동일로가 만나는 태릉입구역 사거리", options: { color: C.text1 } }],
    { x: rx, y: 1.35, w: rw, h: 0.35, fontSize: 14, margin: 0, valign: "middle", isTextBox: true });
  s.addText("선정 이유", { x: rx, y: 1.85, w: 1.0, h: 0.3, fontSize: 14, bold: true, color: C.accent1, margin: 0, valign: "middle", isTextBox: true });
  let cxp = rx + 1.05;
  ["스마트교차로 자료 확보 가능", "현장 접근 가능", "회전별·접근로별 분석 가능"].forEach((t) => { cxp += chip(s, cxp, 1.85, t, "info") + 0.12; });

  // Lane table
  const H = (t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text2 }, align: "center" } });
  const lane = (t) => ({ text: t, options: { align: "center", color: t ? C.text1 : C.accent5, fill: { color: t.includes("좌회전") ? C.background2 : C.background1 } } });
  const rows = [[H("진입 접근로"), H("차로 수"), H("1 (중앙선)"), H("2"), H("3"), H("4"), H("5")]];
  [
    ["화랑로 동쪽", 5, ["좌회전·유턴", "좌회전", "직진", "직진", "직진·우회전"]],
    ["화랑로 서쪽\n(석계역 쪽)", 5, ["좌회전·유턴", "직진", "직진", "직진", "우회전 전용"]],
    ["동일로 북쪽\n(공릉역 쪽)", 4, ["좌회전·유턴", "직진", "직진", "직진·우회전", ""]],
    ["동일로 남쪽\n(먹골역 쪽)", 5, ["좌회전·유턴", "좌회전", "직진", "직진", "직진·우회전"]],
  ].forEach(([a, n, ls]) => rows.push([
    { text: a, options: { bold: true, color: C.text2 } },
    { text: String(n), options: { align: "center", bold: true, color: C.text2 } },
    ...ls.map((t) => (t ? lane(t) : { text: "–", options: { align: "center", color: C.accent5 } })),
  ]));
  s.addTable(rows, { x: rx, y: 2.4, w: rw, colW: [1.63, 0.8, 1.08, 1.08, 1.08, 1.08, 1.08], rowH: [0.45, 0.72, 0.72, 0.72, 0.72],
    fontSize: 12, color: C.text1, border: { type: "solid", pt: 0.75, color: HEX.accent6 }, valign: "middle", margin: [0.03, 0.04, 0.03, 0.08], objectName: "lane-table" });
  s.addText("차로 구성은 중앙선 쪽부터 표시  ·  음영: 좌회전 포함 차로", { x: rx, y: 5.73, w: rw, h: 0.25, fontSize: 11, color: C.accent5, margin: 0, isTextBox: true });

  notes(s, 6.1, [
    "※ 지도·거리뷰로 잠정 확인한 구성이며, 현장에서 대조 예정",
    "※ 버스전용차로 위치·운영시간·적용 범위와 유턴 조건은 현장 표지판으로 확인 예정",
  ]);
  source(s, "출처: 네이버지도 위성사진·거리뷰", 6.72);
  s.addNotes("태릉입구역 교차로는 스마트교차로 교통량 자료를 얻을 수 있고 현장 접근이 가능해 선정했습니다. 차로 구성은 지도·거리뷰로 잠정 확인한 것으로 현장조사에서 대조할 예정입니다.");
}

// ======================= Slide 5 =======================
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "신호교차로" });
  s.addText("교차로 교통량 예비 분석", { placeholder: "title" });

  // Meta strip
  const meta = [
    ["자료", "서울시 T-DATA\n스마트교차로 5분 교통량", 2.75],
    ["날짜", "2026년 9월 29일", 2.0],
    ["대상", "태릉입구역 교차로\n(ID 172)", 2.15],
    ["방법", "회전별 5분 VOL 연속 12개 합산 → 첨두 1시간\n(좌회전·직진·우회전·유턴 모두 포함)", 4.6],
  ];
  let mx = 0.6;
  meta.forEach(([k, v, w]) => {
    card(s, mx, 1.3, w, 1.0);
    s.addText([{ text: k, options: { fontSize: 11, bold: true, color: C.accent1, breakLine: true } }, { text: v, options: { fontSize: 12, color: C.text1 } }],
      { x: mx + 0.18, y: 1.36, w: w - 0.3, h: 0.88, margin: 0, valign: "middle", isTextBox: true });
    mx += w + 0.2;
  });

  // Stat callouts
  const sx = 0.6, sw = 4.0, sh = 1.38;
  [["오전 첨두 1시간", "07:25 ~ 08:25", "5,804", false, 2.5], ["오후 첨두 1시간", "17:20 ~ 18:20", "6,169", true, 4.03]].forEach(([k, t, v, hi, y]) => {
    card(s, sx, y, sw, sh, hi ? C.text2 : C.background2);
    s.addText([{ text: k, options: { fontSize: 13, bold: true, color: hi ? C.accent6 : C.accent1, breakLine: true } }, { text: t, options: { fontSize: 14, color: hi ? C.background1 : C.text1 } }],
      { x: sx + 0.25, y: y + 0.2, w: 1.9, h: sh - 0.4, margin: 0, valign: "middle", isTextBox: true });
    s.addText([{ text: v, options: { fontSize: 34, bold: true } }, { text: " 대", options: { fontSize: 14 } }],
      { x: sx + 2.0, y: y + 0.2, w: sw - 2.2, h: sh - 0.4, align: "right", margin: 0, valign: "middle", color: hi ? C.background1 : C.text2, isTextBox: true });
  });
  s.addText("전체 진입 교통량 기준", { x: sx, y: 5.5, w: sw, h: 0.25, fontSize: 11, color: C.accent5, margin: 0, isTextBox: true });

  // Bar chart
  s.addChart(pres.charts.BAR, [{ name: "진입 교통량", labels: ["서쪽에서 진입", "동쪽에서 진입", "남쪽에서 진입", "북쪽에서 진입"], values: [1992, 1755, 1295, 1127] }], {
    x: 4.95, y: 2.45, w: 7.78, h: 3.3, barDir: "col", barGapWidthPct: 70,
    chartColors: [HEX.dk2, HEX.accent3, HEX.accent3, HEX.accent3],
    showTitle: true, title: "오후 첨두 1시간 진입방향별 교통량 (대)", titleFontSize: 13, titleColor: HEX.dk2, titleFontFace: "+mn-lt",
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "#,##0", dataLabelFontSize: 12, dataLabelColor: HEX.dk1, dataLabelFontFace: "+mn-lt",
    catAxisLabelColor: HEX.dk1, catAxisLabelFontSize: 12, catAxisLabelFontFace: "+mn-lt",
    valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: 2400,
    valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false,
    catAxisLineColor: HEX.accent3, objectName: "dir-volume-chart",
  });

  // Interpretation
  card(s, 0.6, 5.95, 12.13, 0.62);
  s.addText([
    { text: "해석  ", options: { bold: true, color: C.accent1 } },
    { text: "해당 날짜는 오후 첨두 교통량이 더 많고, 서쪽 접근로 진입 교통량이 가장 많음", options: { color: C.text1, breakLine: true } },
    { text: "※ 한 날짜의 예비 결과로, 일반적인 첨두시간으로 확정하지 않음", options: { color: C.accent4, bold: true, fontSize: 12 } },
  ], { x: 0.8, y: 5.95, w: 11.8, h: 0.62, fontSize: 13, paraSpaceAfter: 0, margin: 0, valign: "middle", isTextBox: true });
  source(s, "출처: 서울시 T-DATA, VW_SC_TRF_VOL_5M_20260929.csv", 6.65);
  s.addNotes("2026년 9월 29일 하루 자료로, 회전별 5분 교통량을 연속 12개씩 합산해 첨두 1시간을 찾았습니다. 오후 17:20~18:20이 6,169대로 오전보다 많았고, 방향별로는 서쪽 진입이 1,992대로 가장 많았습니다. 한 날짜 결과이므로 일반적인 첨두시간으로 확정하지는 않습니다.");
}

// ======================= Slide 6 =======================
pres.addSection({ title: "향후 계획" });
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "향후 계획" });
  s.addText("현장조사 및 향후 분석 계획", { placeholder: "title" });
  const top = 1.35, ch = 4.5;

  // Col 1: intersection survey
  const c1x = 0.6, c1w = 5.1;
  card(s, c1x, top, c1w, ch);
  sectionHead(s, c1x + 0.25, top + 0.15, 3, "교차로 현장조사");
  chip(s, c1x + c1w - 0.25 - 1.4, top + 0.2, "조사일·시간 미정", "todo", 1.4);
  s.addText([{ text: "일정  ", options: { bold: true, color: C.accent1 } }, { text: "평일 아침 또는 주말 중 가능한 시간", options: { color: C.text1 } }],
    { x: c1x + 0.25, y: top + 0.65, w: c1w - 0.5, h: 0.35, fontSize: 14, margin: 0, valign: "middle", isTextBox: true });
  s.addText(bullets([
    "차로 수·회전 구성·버스전용차로·유턴 표지 확인",
    "신호주기·현시 순서·현시별 녹색 및 황색시간 기록",
    "신호 3~5주기 반복 관측 → 변동 여부 확인",
    "조사 날짜·시간과 사진 기록",
    "보도 등 안전한 관측 위치에서 조사",
  ]), { x: c1x + 0.25, y: top + 1.15, w: c1w - 0.45, h: ch - 1.3, fontSize: 14, color: C.text1, margin: 0, valign: "top", paraSpaceAfter: 9, isTextBox: true });

  // Col 2: highway + photo slot
  const c2x = 6.0, c2w = 3.25;
  card(s, c2x, top, c2w, 2.0);
  sectionHead(s, c2x + 0.25, top + 0.15, c2w - 0.5, "고속도로 추가 조사");
  s.addText([
    { text: "공개자료와 기관 문의 중심으로 보완", options: { color: C.text1, breakLine: true } },
    { text: "도로 조건 · 중차량 구성비 · VDS 설치정보", options: { color: C.text2, bold: true } },
  ], { x: c2x + 0.25, y: top + 0.65, w: c2w - 0.45, h: 1.2, fontSize: 14, margin: 0, valign: "top", paraSpaceAfter: 8, isTextBox: true });
  imageSlot(s, c2x, top + 2.25, c2w, ch - 2.25, "현장사진 삽입 위치", ["현장조사 후 추가"]);

  // Col 3: next analysis steps
  const c3x = 9.55, c3w = 3.18;
  sectionHead(s, c3x, top + 0.0, c3w, "향후 분석");
  const steps = [
    "현장조사 결과와 기존 자료 대조",
    "최종 분석용 교통량과 신호 운영의 시간대 일치",
    "용량·서비스수준 산정에 필요한 추가 항목 확인",
    "입력자료 확보 후 용량·서비스수준 분석",
  ];
  const sy0 = top + 0.5, sstep = 1.0;
  steps.forEach((t, i) => {
    const y = sy0 + i * sstep;
    s.addText(String(i + 1), { x: c3x, y: y + 0.12, w: 0.45, h: 0.45, shape: pres.shapes.OVAL, fill: { color: i === 3 ? C.text2 : C.accent1 }, line: { color: i === 3 ? C.text2 : C.accent1, width: 0 },
      color: C.background1, fontSize: 14, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: nm("step-num") });
    s.addText(t, { x: c3x + 0.6, y, w: c3w - 0.6, h: 0.7, fontSize: 14, color: C.text1, bold: i === 3, margin: 0, valign: "middle", isTextBox: true });
  });

  notes(s, 6.05, [
    "※ 주말에 관측한 신호시간을 평일 첨두 교통량에 그대로 적용하지 않음",
    "※ 최종 분석에 필요한 차종 구성·보행량 등은 추가 확보 계획에 포함",
  ]);
  s.addNotes("교차로 현장조사는 평일 아침 또는 주말 중 가능한 시간에 진행할 예정이며 구체적 일정은 미정입니다. 신호는 3~5주기 반복 관측합니다. 고속도로는 공개자료와 기관 문의로 도로 조건, 중차량 구성비, VDS 설치정보를 보완합니다. 이후 자료 대조와 시간대 일치를 거쳐 용량·서비스수준 분석을 진행합니다.");
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
