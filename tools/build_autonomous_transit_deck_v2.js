// 스마트시티 자율주행 대중교통 정책의 국내외 사례 비교 및 발전방향 — v2 (절제된 학술 발표 스타일)
// 디자인 원칙: 흰 배경, 남색·진회색, 카드/아이콘/장식 화살표 없음, 보고서형 표(3선)·그림 캡션, 하단 출처.
// 모든 수치·사실은 보고서(2026.11.05., 전현규) 원문 기준.
// 실행: node tools/build_autonomous_transit_deck_v2.js <out.pptx> [sample]
const path = require("path");
const pptxgen = require("pptxgenjs");
const { applyTheme } = require(process.env.APPLY_THEME ||
  "/root/.claude/skills/synced/2709c09c-c0dc-4f81-b385-fcb3527e2d18_99de5be8-b623-4241-b2e1-4194ab58e11a/pptx/scripts/apply_theme.js");

const OUT = process.argv[2] || "deck.pptx";
const FIG1 = path.join(__dirname, "fig1_seoul_a160.jpg");      // 보고서 [그림 1], 760×509

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

// ---------- Layout ----------
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
// 개조식 본문: [level, text] 배열. level 0 = ■ 소제목, 1 = – 항목, 2 = 보충(괄호·근거)
function outline(items) {
  return items.map(([lv, t], i) => {
    const last = i === items.length - 1;
    const base = { breakLine: !last };
    if (lv === 0) return { text: t, options: { ...base, bold: true, color: C.text2, fontSize: 17,
      bullet: { characterCode: "25A0", indent: 20 }, paraSpaceBefore: i === 0 ? 0 : 12, paraSpaceAfter: 4 } };
    if (lv === 1) return { text: t, options: { ...base, color: C.text1, fontSize: 16,
      bullet: { characterCode: "2013", indent: 18 }, indentLevel: 1, paraSpaceAfter: 4 } };
    return { text: t, options: { ...base, color: C.accent5, fontSize: 14,
      bullet: { characterCode: "00B7", indent: 16 }, indentLevel: 2, paraSpaceAfter: 4 } };
  });
}
// 보고서형 3선 표: 상·하단 굵은 남색 선, 헤더 하단 선, 본문 행 사이 연회색 선, 세로선 없음
function academicTable(s, header, body, opt) {
  const none = { type: "none" };
  const heavy = { type: "solid", pt: 1.5, color: HEX.dk2 };
  const mid = { type: "solid", pt: 0.75, color: HEX.dk2 };
  const thin = { type: "solid", pt: 0.5, color: HEX.accent6 };
  const fs = opt.fontSize || 14;
  const rows = [header.map((h, j) => ({ text: h, options: { bold: true, color: C.text2, fill: { color: C.background2 },
    align: (opt.headerAlign && opt.headerAlign[j]) || "center", border: [heavy, none, mid, none] } }))];
  body.forEach((r, i) => {
    const lastRow = i === body.length - 1;
    rows.push(r.map((c, j) => ({ text: c, options: { color: C.text1, bold: j === 0 && opt.boldFirstCol,
      align: (opt.align && opt.align[j]) || "left", border: [none, none, lastRow ? heavy : thin, none] } })));
  });
  s.addTable(rows, { x: opt.x, y: opt.y, w: opt.w, colW: opt.colW, rowH: opt.rowH, fontSize: fs, valign: "middle",
    margin: [0.03, 0.1, 0.03, 0.1], objectName: nm(opt.name || "table") });
}

// ======================= 샘플 A: 일반 설명 슬라이드 =======================
function slideSeoulOverview() {
  const s = pres.addSlide({ masterName: "CONTENT" });
  section(s, "Ⅱ. 국내 사례");
  s.addText("서울 새벽 동행 자율주행 버스(A160) 개요", { placeholder: "title" });

  const lw = 7.1;
  tx(s, outline([
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
  ]), { x: L, y: BODY_Y + 0.05, w: lw, h: 5.2, valign: "top", lineSpacingMultiple: 1.2, name: "body" });

  const fx = L + lw + 0.35, fw = R - fx, fh = fw * 509 / 760, fy = BODY_Y + 0.1;
  s.addImage({ path: FIG1, x: fx, y: fy, w: fw, h: fh, altText: "서울 새벽 동행 자율주행 버스 A160 (보고서 그림 1)", objectName: "fig1" });
  caption(s, "[그림 1] 서울 새벽 동행 자율주행 버스 A160", fx, fy + fh + 0.1, fw, "center");
  tx(s, "자료: 서울특별시(2024)", { x: fx, y: fy + fh + 0.42, w: fw, h: 0.28, fontSize: 11, color: C.accent5, align: "center", name: "fig-source" });

  source(s, "출처: 보고서 p.1–2 (사업 개요, 도입 배경 및 목적, [그림 1])");
  s.addNotes("A160은 2024년 11월 26일 처음 운행했습니다. 도봉산역 광역환승센터에서 영등포역까지 편도 25.7km 구간을 평일 오전 3시 30분에 출발하며, 일반 시내버스 첫차보다 약 30분 이릅니다. 기존 160번 노선을 기반으로 주요 구간 87개 정류소를 운행합니다. 도입 배경은 새벽 출근 근로자의 이동 불편과 첫차 혼잡(기존 160번 첫차 최대 재차 인원 50명 초과), 그리고 새벽 시간대 운전기사 수급난으로 인한 증차 제약입니다.");
}

// ======================= 샘플 B: 표 중심 슬라이드 =======================
function slideCompareTable() {
  const s = pres.addSlide({ masterName: "CONTENT" });
  section(s, "Ⅳ. 국내외 사례 비교 및 시사점");
  s.addText("국내외 자율주행 대중교통 정책 사례 비교", { placeholder: "title" });

  tx(s, outline([
    [1, "공통점: 자율주행 기술을 활용하여 기존 대중교통체계의 한계를 보완"],
    [1, "차이점: 지역별 교통 문제에 따라 도입 목적, 차량 형태, 서비스 운영방식이 다름"],
  ]), { x: L, y: BODY_Y, w: CW, h: 0.75, valign: "top", name: "lead" });

  const ty = 2.3;
  caption(s, "<표 3> 국내외 자율주행 대중교통 정책 사례 비교", L, ty, CW);
  academicTable(s,
    ["구분", "서울 새벽 동행 자율주행버스(A160)", "싱가포르 Punggol 자율주행 셔틀"],
    [
      ["해결할 도시 문제", "새벽 시간대 교통 공백 및 첫차 혼잡", "First/Last Mile 연결성 부족"],
      ["도입 목적", "새벽 근로자 이동 지원 및 교통서비스 공급 확대", "지역 내 이동 편의성 및 대중교통 접근성 개선"],
      ["차량 형태", "대형 자율주행버스", "소형 자율주행 셔틀"],
      ["운영 방식", "기존 시내버스 노선 기반 조기 운행", "고정노선 도입 후 수요응답형 서비스 확대"],
      ["주요 이용 대상", "새벽 출근 근로자 등", "지역 주민 및 주요 생활시설, 교통거점 이용자"],
      ["성과평가", "이용객 수, 만족도, 이용목적, 사고 발생 여부", "고유 이용자 수, 체감 안전성, 추천 의향, 이용자 의견"],
      ["주요 성과", "누적 탑승 인원 약 27,600명 이용, 만족도 4.08/5점", "고유 이용자 11,500명 이상 이용, 설문 응답자 99% 긍정 평가"],
      ["향후 발전 방향", "새벽 혼잡노선 확대 및 운영체계 개선", "수요응답형 운행 및 승하차 지점 확대"],
    ],
    { x: L, y: ty + 0.38, w: CW, colW: [2.0, 4.9, CW - 6.9], rowH: [0.42, ...Array(8).fill(0.42)], fontSize: 14,
      align: ["center", "left", "left"], boldFirstCol: true, name: "compare-table" });
  tx(s, "자료: 서울특별시(2026.3.29.), LTA(2025.9.20., 2026.7.29.), Grab(2026.10.5.)을 바탕으로 작성", {
    x: L, y: ty + 0.38 + 0.42 + 8 * 0.42 + 0.12, w: CW, h: 0.28, fontSize: 11, color: C.accent5, name: "table-source" });

  source(s, "출처: 보고서 p.5 (국내외 자율주행 대중교통 정책 비교, <표 3>)");
  s.addNotes("두 사례 모두 자율주행 기술로 기존 대중교통체계의 한계를 보완한다는 공통점이 있지만, 지역별 교통 문제에 따라 도입 목적, 차량 형태, 운영방식에서 차이를 보입니다. 서울은 대형 자율주행버스로 기존 시내버스 노선을 기반으로 조기 운행하고, 싱가포르는 소형 셔틀로 고정노선을 도입한 뒤 수요응답형 서비스로 확대했습니다.");
}

// ======================= 샘플 C: 분석 결과 슬라이드 =======================
function slidePunggolResults() {
  const s = pres.addSlide({ masterName: "CONTENT" });
  section(s, "Ⅲ. 해외 사례");
  s.addText("Punggol 자율주행 셔틀 운영성과 및 이용자 의견", { placeholder: "title" });

  // 좌: 원문 <표 2>
  const lw = 5.6, ty = BODY_Y + 0.05;
  caption(s, "<표 2> 주요 운영성과 및 이용자 의견", L, ty, lw);
  academicTable(s, ["구분", "운영성과"], [
    ["고유 이용자 수", "11,500명 이상"],
    ["이용자 설문", "약 900명"],
    ["체감 안전성 및 추천 의향", "99%"],
    ["승하차 지점 선택 선호", "약 60%"],
    ["직접적인 이동 경로 선호", "약 40%"],
  ], { x: L, y: ty + 0.38, w: lw, colW: [3.0, lw - 3.0], rowH: [0.42, 0.44, 0.44, 0.44, 0.44, 0.44], fontSize: 14,
    align: ["left", "center"], headerAlign: ["left", "center"], name: "punggol-table" });
  tx(s, [
    { text: "자료: LTA(2026.7.29.)를 바탕으로 작성", options: { breakLine: true } },
    { text: "이용자 수는 2026년 7월 12일 기준, 수요응답형 확대 이전 고정노선 단계의 실적" },
  ], { x: L, y: ty + 0.38 + 0.42 + 5 * 0.44 + 0.06, w: lw, h: 0.55, fontSize: 11, color: C.accent5, valign: "top", name: "table-source" });

  // 우: 설문 응답 비율 막대그래프 (표 2의 비율 항목, 원문 수치)
  const gx = L + lw + 0.55, gw = R - gx;
  caption(s, "[그림 3] 탑승 후 설문 응답 비율 (응답자 약 900명)", gx, ty, gw);
  s.addChart(pres.charts.BAR, [{
    name: "응답 비율",
    labels: ["직접적인 이동 경로 선호", "승하차 지점 선택 선호", "체감 안전성 및 추천 의향"],
    values: [40, 60, 99],
  }], {
    x: gx, y: ty + 0.4, w: gw, h: 2.75, barDir: "bar",
    chartColors: [HEX.dk2], barGapWidthPct: 70,
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: '0"%"', dataLabelFontSize: 13, dataLabelColor: HEX.dk1, dataLabelFontFace: "+mn-lt",
    valAxisMinVal: 0, valAxisMaxVal: 100, valAxisMajorUnit: 20, valAxisLabelFormatCode: '0"%"',
    valAxisLabelFontSize: 11, valAxisLabelColor: HEX.accent5, valAxisLabelFontFace: "+mn-lt",
    catAxisLabelFontSize: 13, catAxisLabelColor: HEX.dk1, catAxisLabelFontFace: "+mn-lt",
    valGridLine: { color: HEX.accent6, size: 0.5 }, catGridLine: { style: "none" },
    catAxisLineShow: true, valAxisLineShow: false, showLegend: false, objectName: "survey-chart",
  });
  tx(s, "주: <표 2>의 비율 항목으로 발표자 작성. 60%, 40%는 원문의 '약' 수치이며 두 문항의 중복 응답 여부는 원문에 명시되지 않음", {
    x: gx, y: ty + 3.2, w: gw, h: 0.5, fontSize: 11, color: C.accent5, valign: "top", name: "chart-note" });

  // 하단: 결과 요약(원문)
  const by = 5.6;
  tx(s, outline([
    [0, "주요 결과"],
    [1, "탑승 후 설문 응답자 99%가 탑승 중 안전하다고 느끼고 타인에게 추천할 의향을 나타냄"],
    [1, "고정노선·고정 승하차 지점 운영에 대한 이용자 요구 확인 → 2026년 9월부터 수요응답형 서비스 병행"],
  ]), { x: L, y: by, w: CW, h: 1.25, valign: "top", name: "summary" });

  source(s, "출처: 보고서 p.4–5 (운영성과 및 서비스 개선, <표 2>)");
  s.addNotes("2026년 7월 12일 기준 고유 이용자는 11,500명 이상입니다. 탑승 후 설문 응답자 약 900명 중 99%가 안전하다고 느끼고 추천 의향을 나타냈습니다. 한편 응답자 약 60%가 승하차 지점 선택을, 약 40%가 직접적인 경로를 선호해 고정노선·고정 승하차 지점 운영에 대한 요구가 확인되었고, 이를 반영해 2026년 9월부터 수요응답형 서비스를 병행하고 있습니다.");
}

(async () => {
  slideSeoulOverview();
  slideCompareTable();
  slidePunggolResults();
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})().catch((e) => { console.error(e); process.exit(1); });
