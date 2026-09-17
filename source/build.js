const pptxgen = require("pptxgenjs");

// ============================================================
// DESIGN SYSTEM
// ============================================================
const FONT = "Noto Sans KR";

const C = {
  bgLight: "FAF7F1",
  bgDark: "26221E",
  ink: "2A2622",
  inkSoft: "5B564D",
  inkFaint: "8A8276",
  line: "DDD5C7",
  lineDark: "463F37",
  cream: "F3EEE6",
  creamMute: "B8AFA0",
  gold: "A6813F",
  goldOnDark: "C9A968",
  sand: "BFA06E",
  paleBlue: "8FA9B3",
  seaGlass: "77998C",
  mutedPink: "BE9494",
  placeholderFill: "F1EBE0",
  placeholderFillDark: "332D26",
};

const W = 13.333, H = 7.5;
const MX = 0.9;              // left/right margin
const CW = W - MX * 2;       // content width
const HEADER_Y = 0.58;
const TITLE_Y = 0.92;
const FOOTER_Y = 7.12;

function newSlide(pres, dark) {
  const s = pres.addSlide();
  s.background = { color: dark ? C.bgDark : C.bgLight };
  return s;
}

function footer(s, pageNum, dark) {
  const fg = dark ? C.creamMute : C.inkFaint;
  s.addText("WOULD YOU, TOKKI?", {
    x: MX, y: FOOTER_Y, w: 4, h: 0.25, fontFace: FONT, fontSize: 8,
    color: fg, charSpacing: 2, isTextBox: true, margin: 0,
  });
  s.addText(pageNum, {
    x: W - MX - 1.5, y: FOOTER_Y, w: 1.5, h: 0.25, fontFace: FONT, fontSize: 8,
    color: fg, align: "right", charSpacing: 1, isTextBox: true, margin: 0,
  });
}

// Section tag: small dot + tracked caps label
function sectionTag(s, label, dark) {
  const dotColor = dark ? C.goldOnDark : C.gold;
  const textColor = dark ? C.goldOnDark : C.gold;
  s.addShape("ellipse", {
    x: MX, y: HEADER_Y + 0.065, w: 0.09, h: 0.09, fill: { color: dotColor }, line: { type: "none" },
  });
  s.addText(label.toUpperCase(), {
    x: MX + 0.2, y: HEADER_Y, w: 6, h: 0.22, fontFace: FONT, fontSize: 10.5, bold: true,
    color: textColor, charSpacing: 2.5, isTextBox: true, margin: 0,
  });
}

function title(s, text, opts = {}) {
  const dark = opts.dark || false;
  s.addText(text, {
    x: MX, y: opts.y || TITLE_Y, w: opts.w || CW, h: opts.h || 1.0,
    fontFace: FONT, fontSize: opts.size || 28, bold: true,
    color: dark ? C.cream : C.ink, lineSpacing: opts.lineSpacing || (opts.size ? opts.size * 1.22 : 34),
    isTextBox: true, margin: 0, align: opts.align || "left",
  });
}

function body(s, text, x, y, w, h, opts = {}) {
  s.addText(text, {
    x, y, w, h, fontFace: FONT, fontSize: opts.size || 11.5,
    color: opts.color || (opts.dark ? C.creamMute : C.inkSoft),
    lineSpacing: opts.lineSpacing || (opts.size ? opts.size * 1.35 : 15.5),
    isTextBox: true, margin: 0, align: opts.align || "left", valign: opts.valign || "top",
  });
}

// Labeled placeholder box for a missing image
function placeholder(s, x, y, w, h, label, dark) {
  s.addShape("rect", {
    x, y, w, h,
    fill: { color: dark ? C.placeholderFillDark : C.placeholderFill },
    line: { color: dark ? C.lineDark : C.line, width: 1, dashType: "dash" },
  });
  s.addText("IMAGE", {
    x: x + 0.12, y: y + 0.1, w: w - 0.24, h: 0.2, fontFace: FONT, fontSize: 7.5, bold: true,
    color: dark ? C.creamMute : C.inkFaint, charSpacing: 2, isTextBox: true, margin: 0,
  });
  s.addText(label, {
    x: x + 0.2, y: y + h / 2 - 0.4, w: w - 0.4, h: 0.8, fontFace: FONT, fontSize: 9.5,
    color: dark ? C.creamMute : C.inkFaint, align: "center", valign: "middle",
    lineSpacing: 13, isTextBox: true, margin: 0,
  });
}

function chip(s, text, x, y, w, h, opts = {}) {
  const dark = opts.dark || false;
  s.addShape("rect", {
    x, y, w, h, fill: { type: "none" }, line: { color: dark ? C.lineDark : C.line, width: 1 },
  });
  s.addText(text, {
    x, y, w, h, fontFace: FONT, fontSize: opts.size || 10, bold: opts.bold || false,
    color: dark ? C.cream : C.ink, align: "center", valign: "middle", isTextBox: true, margin: 0,
    charSpacing: opts.charSpacing || 0.5,
  });
}

function bullets(s, items, x, y, w, h, opts = {}) {
  const dark = opts.dark || false;
  const arr = items.map((t, i) => ({
    text: t,
    options: {
      bullet: { code: "2013", indent: 14 },
      breakLine: i < items.length - 1,
      color: dark ? C.creamMute : C.inkSoft,
      fontSize: opts.size || 10.5,
      paraSpaceAfter: opts.spaceAfter !== undefined ? opts.spaceAfter : 5,
    },
  }));
  s.addText(arr, {
    x, y, w, h, fontFace: FONT, lineSpacing: (opts.size || 10.5) * 1.3,
    isTextBox: true, margin: 0, valign: "top",
  });
}

// ============================================================
const pres = new pptxgen();
pres.defineLayout({ name: "TOKKI", width: W, height: H });
pres.layout = "TOKKI";

// ------------------------------------------------------------
// P1 COVER
// ------------------------------------------------------------
{
  const s = newSlide(pres, true);

  s.addText("CHARACTER ART TOY IP", {
    x: MX, y: 0.75, w: 4.6, h: 0.25, fontFace: FONT, fontSize: 10.5, bold: true,
    color: C.goldOnDark, charSpacing: 3, isTextBox: true, margin: 0,
  });

  s.addText("Would you, Tokki?", {
    x: MX, y: 1.28, w: 4.7, h: 0.58, fontFace: FONT, fontSize: 29, bold: true,
    color: C.cream, isTextBox: true, margin: 0, lineSpacing: 34,
  });
  s.addText("우주…토끼?", {
    x: MX, y: 1.93, w: 4.7, h: 0.55, fontFace: FONT, fontSize: 25, bold: true,
    color: C.cream, isTextBox: true, margin: 0,
  });

  s.addText("Material-based Character Art Toy IP", {
    x: MX, y: 2.68, w: 4.6, h: 0.3, fontFace: FONT, fontSize: 13.5,
    color: C.creamMute, isTextBox: true, margin: 0,
  });
  s.addText("Born from different worlds.", {
    x: MX, y: 3.02, w: 4.6, h: 0.32, fontFace: FONT, fontSize: 13.5, italic: true,
    color: C.goldOnDark, isTextBox: true, margin: 0,
  });

  s.addText(
    "하나의 기본 토끼 조형에 서로 다른 세계의 물성과 환경을 적용해\n확장하는 수집형 캐릭터 아트토이 IP.",
    {
      x: MX, y: 3.55, w: 4.55, h: 1.0, fontFace: FONT, fontSize: 11,
      color: C.creamMute, lineSpacing: 17, isTextBox: true, margin: 0,
    }
  );

  s.addText("예비창업 단계 · IP 개발 및 제품화 준비 중", {
    x: MX, y: 6.35, w: 4.6, h: 0.3, fontFace: FONT, fontSize: 9.5,
    color: C.creamMute, charSpacing: 0.5, isTextBox: true, margin: 0,
  });

  placeholder(s, 5.85, 0.65, 6.58, 6.15,
    "Would you, Tokki?\nFull Character Lineup — Hero Image\n(원본 비율 유지 / 실제 이미지 교체 예정)", true);

  footer(s, "01 / 12", true);
}

// ------------------------------------------------------------
// P2 PROBLEM
// ------------------------------------------------------------
{
  const s = newSlide(pres, false);
  sectionTag(s, "Problem", false);
  title(s, "캐릭터는 많지만,\n오래 소유하고 싶은 캐릭터는 적다.", { size: 25, y: 0.92, h: 1.05, lineSpacing: 31 });

  body(
    s,
    "캐릭터 시장에는 다양한 그래픽과 굿즈가 존재하지만, 성인 소비자가 자신의 생활공간에 오래 두고 싶은 조형적 오브젝트형 캐릭터는 상대적으로 선택지가 제한적입니다.\n\n반면 컬렉터 중심의 아트토이는 작품성과 수집가치는 높지만 가격과 접근성에서 진입장벽이 존재합니다.",
    MX, 2.35, 5.15, 3.4, { size: 11.5 }
  );

  // Venn diagram: Character / Art Toy / Lifestyle Object (fixed, verified coordinates)
  const r = 1.15;
  const positions = [
    { cx: 8.15, cy: 3.55, color: C.paleBlue, label: "CHARACTER", lx: 8.15, ly: 2.78 },
    { cx: 10.05, cy: 3.55, color: C.sand, label: "ART TOY", lx: 10.05, ly: 2.78 },
    { cx: 9.1, cy: 4.95, color: C.seaGlass, label: "LIFESTYLE\nOBJECT", lx: 9.1, ly: 5.55 },
  ];
  // bounding box: x 7.0–11.2, y 2.4–6.1 (all within margins, clear of title above and closing line below)
  positions.forEach((p) => {
    s.addShape("ellipse", {
      x: p.cx - r, y: p.cy - r, w: r * 2, h: r * 2,
      fill: { color: p.color, transparency: 88 },
      line: { color: p.color, width: 1.25 },
    });
  });
  positions.forEach((p) => {
    s.addText(p.label, {
      x: p.lx - 1.05, y: p.ly - 0.25, w: 2.1, h: 0.5, fontFace: FONT, fontSize: 10, bold: true,
      color: C.ink, align: "center", valign: "top", isTextBox: true, margin: 0, lineSpacing: 12, charSpacing: 0.5,
    });
  });

  s.addText("Would you, Tokki?는 이 세 영역의 접점을 지향합니다.", {
    x: 6.35, y: 6.35, w: 5.85, h: 0.4, fontFace: FONT, fontSize: 12, bold: true, italic: true,
    color: C.ink, align: "center", isTextBox: true, margin: 0,
  });

  footer(s, "02 / 12", false);
}

// ------------------------------------------------------------
// P3 SOLUTION
// ------------------------------------------------------------
{
  const s = newSlide(pres, false);
  sectionTag(s, "Solution", false);
  title(s, "같은 토끼, 다른 세계.", { size: 27, y: 0.92, h: 0.55 });
  s.addText("같은 형태, 다른 물성.", {
    x: MX, y: 1.5, w: CW, h: 0.35, fontFace: FONT, fontSize: 15, color: C.inkSoft, italic: true,
    isTextBox: true, margin: 0,
  });

  body(
    s,
    "Would you, Tokki?는 하나의 기본 토끼 실루엣을 유지하면서 각기 다른 행성, 재료, 액체, 음식, 자연환경의 물성을 캐릭터 신체 자체에 적용하는 방식으로 확장됩니다.",
    MX, 2.0, CW, 0.65, { size: 11 }
  );

  const materials = [
    ["MOON", "크레이터"], ["ICE", "얼음 결정"], ["ROCK", "갈라진 암석"], ["SAND", "퇴적층"],
    ["SEA", "바다와 해양생물"], ["BUBBLE", "탄산과 기포"], ["PUDDING", "흔들리는 젤"], ["BREAD", "구워진 반죽"],
    ["BUTTERCREAM", "크리미한 질감"], ["ANGORA", "풍성한 긴 털"],
  ];
  const n = materials.length;
  const gap = 0.14;
  const cellW = (CW - gap * (n - 1)) / n;
  const rowY = 2.9, boxH = cellW;
  materials.forEach((m, i) => {
    const x = MX + i * (cellW + gap);
    placeholder(s, x, rowY, cellW, boxH, "", false);
    s.addText(m[0], {
      x: x - 0.05, y: rowY + boxH + 0.08, w: cellW + 0.1, h: 0.2, fontFace: FONT, fontSize: 8, bold: true,
      color: C.ink, align: "center", isTextBox: true, margin: 0, charSpacing: 0.5,
    });
    s.addText(m[1], {
      x: x - 0.05, y: rowY + boxH + 0.29, w: cellW + 0.1, h: 0.35, fontFace: FONT, fontSize: 7.5,
      color: C.inkFaint, align: "center", isTextBox: true, margin: 0, lineSpacing: 9,
    });
  });

  s.addText(
    "새로운 캐릭터를 매번 처음부터 만드는 것이 아니라, 하나의 조형 시스템을 지속적으로 변주합니다.",
    {
      x: MX, y: rowY + boxH + 0.85, w: CW, h: 0.6, fontFace: FONT, fontSize: 13, bold: true,
      color: C.ink, align: "center", isTextBox: true, margin: 0, lineSpacing: 18,
    }
  );

  footer(s, "03 / 12", false);
}

// ------------------------------------------------------------
// P4 PRODUCT
// ------------------------------------------------------------
{
  const s = newSlide(pres, false);
  sectionTag(s, "Product", false);
  title(s, "COLLECTION 01 — ORIGIN", { size: 27, y: 0.92, h: 0.5 });
  s.addText("첫 번째 제품은 네 가지 기본 물성을 가진 아트토이로 시작합니다.", {
    x: MX, y: 1.46, w: CW, h: 0.32, fontFace: FONT, fontSize: 12.5, color: C.inkSoft, isTextBox: true, margin: 0,
  });

  const four = [["MOON", "크레이터"], ["ICE", "얼음 결정"], ["ROCK", "갈라진 암석"], ["SAND", "퇴적층"]];
  const gap = 0.28;
  const cw4 = (CW - gap * 3) / 4;
  const y4 = 2.0, h4 = 2.15;
  four.forEach((m, i) => {
    const x = MX + i * (cw4 + gap);
    placeholder(s, x, y4, cw4, h4, "", false);
    s.addText(m[0], {
      x, y: y4 + h4 + 0.1, w: cw4, h: 0.22, fontFace: FONT, fontSize: 11, bold: true,
      color: C.ink, align: "center", isTextBox: true, margin: 0, charSpacing: 1,
    });
    s.addText(m[1], {
      x, y: y4 + h4 + 0.33, w: cw4, h: 0.2, fontFace: FONT, fontSize: 9, color: C.inkFaint,
      align: "center", isTextBox: true, margin: 0,
    });
  });

  const listY = y4 + h4 + 0.68;
  s.addText("COMMON BASE", { x: MX, y: listY, w: 3.6, h: 0.22, fontFace: FONT, fontSize: 10, bold: true, color: C.gold, charSpacing: 1.5, isTextBox: true, margin: 0 });
  bullets(s, ["둥근 구형 바디", "짧은 팔과 다리", "토끼 귀", "작은 얼굴 요소", "단순하고 기억하기 쉬운 실루엣"], MX, listY + 0.28, 3.6, 1.3, { size: 9.5, spaceAfter: 3 });

  s.addText("VARIATION", { x: MX + 3.95, y: listY, w: 3.6, h: 0.22, fontFace: FONT, fontSize: 10, bold: true, color: C.gold, charSpacing: 1.5, isTextBox: true, margin: 0 });
  bullets(s, ["표면 재질", "투명도", "균열", "기포", "퇴적층", "주변 파편", "궤도 오브젝트"], MX + 3.95, listY + 0.28, 3.6, 1.3, { size: 9.5, spaceAfter: 3 });

  // Size system
  const sizeX = MX + 8.0, sizeW = CW - 7.1;
  s.addText("SIZE SYSTEM", { x: sizeX, y: listY, w: sizeW, h: 0.22, fontFace: FONT, fontSize: 10, bold: true, color: C.gold, charSpacing: 1.5, isTextBox: true, margin: 0 });
  s.addText([
    { text: "MINI  7–9cm", options: { fontSize: 9, color: C.inkSoft, breakLine: true } },
    { text: "STANDARD  12–15cm", options: { fontSize: 9, color: C.ink, bold: true, breakLine: true } },
    { text: "ART EDITION  25–40cm", options: { fontSize: 9, color: C.inkSoft } },
  ], { x: sizeX, y: listY + 0.28, w: sizeW, h: 0.75, fontFace: FONT, lineSpacing: 15, isTextBox: true, margin: 0 });
  s.addText("초기에는 STANDARD 사이즈를 중심으로 프로토타입 제작과 소비자 반응 테스트를 진행합니다.", {
    x: sizeX, y: listY + 1.08, w: sizeW, h: 0.5, fontFace: FONT, fontSize: 8.5, color: C.inkFaint,
    lineSpacing: 12, isTextBox: true, margin: 0,
  });

  footer(s, "04 / 12", false);
}

// ------------------------------------------------------------
// P5 MATERIAL -> QUESTION
// ------------------------------------------------------------
{
  const s = newSlide(pres, false);
  sectionTag(s, "Brand Experience", false);
  title(s, "A rabbit.\nA material.\nA question.", { size: 24, y: 0.92, h: 1.55, lineSpacing: 29 });

  body(
    s,
    "Would you, Tokki?에서 물질은 캐릭터의 외형을 만드는 데서 끝나지 않습니다. 각 물질의 실제 성질에서 출발한 짧은 질문을 통해 캐릭터와 사용자의 경험을 연결합니다.",
    MX, 2.62, 4.0, 1.55, { size: 10.5 }
  );

  s.addShape("rect", { x: MX, y: 4.35, w: 4.0, h: 1.0, fill: { color: C.placeholderFill }, line: { type: "none" } });
  s.addText("이것은 심리검사나 성격유형 테스트가 아닙니다. 정답을 제시하지 않고, 사용자가 물질을 매개로 자신의 경험을 자유롭게 해석하게 합니다.", {
    x: MX + 0.18, y: 4.5, w: 3.64, h: 0.72, fontFace: FONT, fontSize: 8.7, italic: true, color: C.inkSoft,
    lineSpacing: 12, isTextBox: true, margin: 0,
  });

  s.addText("Art Toy → Question Card → Art Book → Exhibition → Web Experience → Question Archive", {
    x: MX, y: 5.55, w: 4.0, h: 0.6, fontFace: FONT, fontSize: 8, color: C.inkFaint, lineSpacing: 11.5,
    isTextBox: true, margin: 0,
  });

  // Flow diagram
  const stages = ["MATERIAL", "PHYSICAL\nPROPERTY", "QUESTION", "PERSONAL\nINTERPRETATION"];
  const fx = 5.35, fy = 1.9, fw = CW - (fx - MX);
  const fGap = 0.28;
  const fCellW = (fw - fGap * (stages.length - 1) - 0.5) / stages.length;
  stages.forEach((st, i) => {
    const x = fx + i * (fCellW + fGap + 0.125);
    s.addShape("rect", { x, y: fy, w: fCellW, h: 0.68, fill: { color: C.placeholderFill }, line: { color: C.line, width: 1 } });
    s.addText(st, { x, y: fy, w: fCellW, h: 0.68, fontFace: FONT, fontSize: 8.3, bold: true, color: C.ink, align: "center", valign: "middle", lineSpacing: 10.5, isTextBox: true, margin: 0, charSpacing: 0.3 });
    if (i < stages.length - 1) {
      s.addText("→", { x: x + fCellW, y: fy, w: 0.25, h: 0.68, fontFace: FONT, fontSize: 14, color: C.inkFaint, align: "center", valign: "middle", isTextBox: true, margin: 0 });
    }
  });

  const examples = [
    { m: "ROCK", tone: C.inkSoft, props: "단단함 · 무게 · 오래 지속됨", q: "당신을 오래 버티게 한 것은 무엇인가요?" },
    { m: "CLOUD", tone: C.paleBlue, props: "가벼움 · 이동 · 머무르지 않음", q: "잠시 멈춰도 된다고 생각하나요?" },
    { m: "BUBBLE", tone: C.seaGlass, props: "투명함 · 가벼움 · 짧은 지속성", q: "짧았지만 오래 기억되는 순간이 있나요?" },
  ];
  let ey = fy + 1.05;
  const rowH = 1.18;
  examples.forEach((ex) => {
    s.addShape("rect", { x: fx, y: ey, w: 0.62, h: rowH - 0.16, fill: { color: ex.tone, transparency: 82 }, line: { color: ex.tone, width: 1 } });
    s.addText(ex.m, { x: fx, y: ey, w: 0.62, h: rowH - 0.16, fontFace: FONT, fontSize: 8, bold: true, color: C.ink, align: "center", valign: "middle", isTextBox: true, margin: 0, lineSpacing: 9 });
    s.addText(ex.props, { x: fx + 0.78, y: ey, w: 2.0, h: rowH - 0.16, fontFace: FONT, fontSize: 8.3, color: C.inkFaint, valign: "middle", lineSpacing: 11, isTextBox: true, margin: 0 });
    s.addText(`“${ex.q}”`, { x: fx + 2.95, y: ey, w: fw - 2.95, h: rowH - 0.16, fontFace: FONT, fontSize: 13, bold: true, italic: true, color: C.ink, valign: "middle", lineSpacing: 16, isTextBox: true, margin: 0 });
    ey += rowH;
  });

  footer(s, "05 / 12", false);
}

// ------------------------------------------------------------
// P6 POSITIONING
// ------------------------------------------------------------
{
  const s = newSlide(pres, false);
  sectionTag(s, "Positioning", false);
  title(s, "아이에게는 캐릭터, 어른에게는 오브제.", { size: 24, y: 0.92, h: 0.55 });

  body(
    s,
    "Would you, Tokki?는 어린이 캐릭터와 고가 컬렉터 아트토이 사이의 중간 지점을 지향합니다. 캐릭터를 좋아하지만 지나치게 아동적이지 않고, 자신의 책상과 선반, 생활공간에 자연스럽게 놓을 수 있는 조형적인 제품을 원하는 소비자를 초기 고객으로 설정합니다.",
    MX, 1.62, 4.35, 1.9, { size: 10.8 }
  );

  s.addText("Character  ×  Art Toy  ×  Lifestyle Object", {
    x: MX, y: 3.55, w: 4.35, h: 0.4, fontFace: FONT, fontSize: 13, bold: true, color: C.ink,
    isTextBox: true, margin: 0,
  });

  const kws = ["CUTE", "OBJECT", "COLLECTIBLE", "MATERIAL", "LIFESTYLE"];
  s.addText(kws.join("   ·   "), {
    x: MX, y: 4.1, w: 4.35, h: 0.55, fontFace: FONT, fontSize: 9.5, color: C.inkFaint,
    charSpacing: 1, lineSpacing: 15, isTextBox: true, margin: 0,
  });

  placeholder(s, 5.75, 1.62, CW - (5.75 - MX), 4.85,
    "Would you, Tokki? — Desk / Shelf Lifestyle Mockup\n(생활공간에 놓인 오브제 이미지, 원본 비율 유지)", false);

  footer(s, "06 / 12", false);
}

// ------------------------------------------------------------
// P7 TARGET
// ------------------------------------------------------------
{
  const s = newSlide(pres, false);
  sectionTag(s, "Target", false);
  title(s, "누가 Would you, Tokki?를 사는가", { size: 26, y: 0.92, h: 0.55 });

  s.addText("PRIMARY TARGET", { x: MX, y: 1.72, w: 5.2, h: 0.24, fontFace: FONT, fontSize: 10.5, bold: true, color: C.gold, charSpacing: 1.5, isTextBox: true, margin: 0 });
  s.addText("20–30대 디자인·캐릭터 소비자", { x: MX, y: 2.0, w: 5.2, h: 0.45, fontFace: FONT, fontSize: 17, bold: true, color: C.ink, isTextBox: true, margin: 0 });
  body(s, "캐릭터를 좋아하지만 지나치게 아동적인 제품보다 조형성과 디자인 완성도가 높은 제품을 선호하는 초기 예상 소비자층.", MX, 2.55, 5.2, 1.0, { size: 10.5 });

  s.addText("INTEREST", { x: MX, y: 3.7, w: 5.2, h: 0.22, fontFace: FONT, fontSize: 9.5, bold: true, color: C.inkFaint, charSpacing: 1.5, isTextBox: true, margin: 0 });
  s.addText("Character IP · Art Toy · Design Object · Stationery · Interior · Independent Brand · Exhibition · Pop-up", {
    x: MX, y: 3.95, w: 5.2, h: 0.65, fontFace: FONT, fontSize: 9.5, color: C.inkSoft, lineSpacing: 14, isTextBox: true, margin: 0,
  });

  s.addText("SECONDARY", { x: MX, y: 4.85, w: 5.2, h: 0.22, fontFace: FONT, fontSize: 9.5, bold: true, color: C.inkFaint, charSpacing: 1.5, isTextBox: true, margin: 0 });
  bullets(s, ["선물용 소비자", "아동과 함께 캐릭터 제품을 소비하는 가족층", "캐릭터 컬래버레이션이 필요한 브랜드·기업"], MX, 5.1, 5.2, 0.9, { size: 10, spaceAfter: 3 });

  s.addShape("rect", { x: MX, y: 6.2, w: 5.2, h: 0.62, fill: { color: C.placeholderFill }, line: { type: "none" } });
  s.addText("실제 고객조사를 완료한 결과가 아닌, 초기 가설로서의 예상 핵심 고객입니다.", {
    x: MX + 0.15, y: 6.2, w: 4.9, h: 0.62, fontFace: FONT, fontSize: 9, italic: true, color: C.inkSoft,
    valign: "middle", lineSpacing: 12, isTextBox: true, margin: 0,
  });

  // Concentric target diagram: primary inner ring / secondary outer ring
  const cx = 9.7, cy = 4.1;
  const rings = [
    { r: 2.05, color: C.line, label: "" },
    { r: 1.5, color: C.paleBlue, label: "" },
    { r: 0.9, color: C.gold, label: "" },
  ];
  rings.forEach((rg) => {
    s.addShape("ellipse", {
      x: cx - rg.r, y: cy - rg.r, w: rg.r * 2, h: rg.r * 2,
      fill: { type: "none" }, line: { color: rg.color, width: 1.25 },
    });
  });
  s.addText("PRIMARY", { x: cx - 0.9, y: cy - 0.18, w: 1.8, h: 0.24, fontFace: FONT, fontSize: 9, bold: true, color: C.ink, align: "center", isTextBox: true, margin: 0, charSpacing: 1 });
  s.addText("SECONDARY", { x: cx - 1.4, y: cy + 1.55, w: 2.8, h: 0.22, fontFace: FONT, fontSize: 8.5, color: C.inkFaint, align: "center", isTextBox: true, margin: 0, charSpacing: 1 });

  footer(s, "07 / 12", false);
}

// ------------------------------------------------------------
// P8 DIFFERENTIATION
// ------------------------------------------------------------
{
  const s = newSlide(pres, false);
  sectionTag(s, "Differentiation", false);
  title(s, "Would you, Tokki?가 확장되는 방식", { size: 26, y: 0.92, h: 0.55 });

  // Expansion flow strip
  const flow = ["Base\nBunny", "Material", "Character", "Question", "Art Toy", "Object /\nSculpture"];
  const fy = 1.75, fh = 0.72;
  const fGap = 0.22;
  const fCellW = (CW - fGap * (flow.length - 1) - 0.22 * (flow.length - 1)) / flow.length;
  flow.forEach((st, i) => {
    const x = MX + i * (fCellW + fGap + 0.22);
    s.addShape("rect", { x, y: fy, w: fCellW, h: fh, fill: { color: C.placeholderFill }, line: { color: C.line, width: 1 } });
    s.addText(st, { x, y: fy, w: fCellW, h: fh, fontFace: FONT, fontSize: 9.5, bold: true, color: C.ink, align: "center", valign: "middle", lineSpacing: 12, isTextBox: true, margin: 0 });
    if (i < flow.length - 1) {
      s.addText("→", { x: x + fCellW, y: fy, w: 0.22, h: fh, fontFace: FONT, fontSize: 13, color: C.inkFaint, align: "center", valign: "middle", isTextBox: true, margin: 0 });
    }
  });

  const items = [
    ["01", "MATERIAL AS CHARACTER", "색이나 의상을 바꾸는 것이 아니라 물질 자체가 캐릭터 디자인이 됩니다."],
    ["02", "ONE BASE, MANY WORLDS", "하나의 기본 조형을 유지하면서 서로 다른 세계로 확장할 수 있습니다."],
    ["03", "3D FIRST", "처음부터 입체 조형과 실제 제품화를 고려해 디자인합니다."],
    ["04", "MATERIAL → QUESTION", "물질의 실제 성질을 사용자의 경험을 떠올리게 하는 열린 질문으로 확장합니다."],
    ["05", "ARTWORK → PRODUCT", "소형 아트토이뿐 아니라 대형 조각, 전시, 공간 설치까지 동일한 IP로 확장할 수 있습니다."],
  ];
  const gridY = 2.85;
  const colW = (CW - 0.5) / 2;
  const rowH = 0.98;
  items.forEach((it, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = MX + col * (colW + 0.5);
    const y = gridY + row * rowH;
    s.addText(it[0], { x, y, w: 0.55, h: 0.9, fontFace: FONT, fontSize: 19, bold: true, color: C.gold, isTextBox: true, margin: 0 });
    s.addText(it[1], { x: x + 0.55, y, w: colW - 0.55, h: 0.3, fontFace: FONT, fontSize: 11.5, bold: true, color: C.ink, isTextBox: true, margin: 0 });
    s.addText(it[2], { x: x + 0.55, y: y + 0.32, w: colW - 0.55, h: 0.6, fontFace: FONT, fontSize: 9.3, color: C.inkSoft, lineSpacing: 12.5, isTextBox: true, margin: 0 });
  });

  footer(s, "08 / 12", false);
}

// ------------------------------------------------------------
// P9 BUSINESS MODEL
// ------------------------------------------------------------
{
  const s = newSlide(pres, false);
  sectionTag(s, "Business Model", false);
  title(s, "IP를 중심으로 확장되는 구조", { size: 26, y: 0.92, h: 0.5 });

  // Center emblem + 5 branch columns (fixed, non-trigonometric layout for predictable bounds)
  const cr = 0.42, cx = W / 2, cy = 1.98;
  s.addShape("ellipse", { x: cx - cr, y: cy - cr, w: cr * 2, h: cr * 2, fill: { color: C.placeholderFill }, line: { color: C.line, width: 1.25 } });
  s.addText("Would you,\nTokki?", { x: cx - cr, y: cy - cr, w: cr * 2, h: cr * 2, fontFace: FONT, fontSize: 8.5, bold: true, color: C.ink, align: "center", valign: "middle", lineSpacing: 10.5, isTextBox: true, margin: 0 });

  const branches = [
    { label: "D2C", items: "Art Toy · Design Object\nLimited Edition · Lifestyle Goods" },
    { label: "B2B", items: "Brand Collaboration · Corporate Goods\nExhibition · Space Collaboration" },
    { label: "LICENSING", items: "Character License\nProduct License · Content Collab." },
    { label: "ART", items: "Limited Art Toy\nLarge Sculpture · Exhibition Work" },
    { label: "CONTENT /\nEXPERIENCE", items: "Question Card · Art Book\nWeb Experience · Exhibition Archive" },
  ];
  const brY = 2.7, brGap = 0.2, brW = (CW - brGap * 4) / 5;
  branches.forEach((b, i) => {
    const x = MX + i * (brW + brGap);
    s.addText(b.label, {
      x, y: brY, w: brW, h: 0.44, fontFace: FONT, fontSize: 10.5, bold: true,
      color: C.ink, align: "center", valign: "top", isTextBox: true, margin: 0, charSpacing: 0.3, lineSpacing: 13,
    });
    s.addText(b.items, {
      x, y: brY + 0.56, w: brW, h: 0.85, fontFace: FONT, fontSize: 7.8,
      color: C.inkFaint, align: "center", lineSpacing: 11, isTextBox: true, margin: 0,
    });
  });

  // Production model
  const py = 4.55;
  s.addText("생산은 전문 제작 파트너와 함께합니다", { x: MX, y: py, w: CW, h: 0.26, fontFace: FONT, fontSize: 10.5, bold: true, color: C.gold, charSpacing: 0.5, isTextBox: true, margin: 0 });
  body(s, "브랜드는 초기부터 자체 제조시설을 구축하지 않습니다.", MX, py + 0.28, CW, 0.25, { size: 9.5 });

  s.addText("핵심 역량", { x: MX, y: py + 0.6, w: 5.6, h: 0.2, fontFace: FONT, fontSize: 9, bold: true, color: C.inkFaint, charSpacing: 1, isTextBox: true, margin: 0 });
  s.addText("IP 기획 · 캐릭터 디자인 · 아트디렉션 · 제품기획 · 브랜드 경험 설계", { x: MX, y: py + 0.82, w: 5.6, h: 0.4, fontFace: FONT, fontSize: 9.5, color: C.inkSoft, lineSpacing: 13, isTextBox: true, margin: 0 });

  s.addText("제작 파트너 협업", { x: MX + 5.9, y: py + 0.6, w: 5.6, h: 0.2, fontFace: FONT, fontSize: 9, bold: true, color: C.inkFaint, charSpacing: 1, isTextBox: true, margin: 0 });
  s.addText("3D Modeling · 3D Printing · Mold · Production · Painting · Printing", { x: MX + 5.9, y: py + 0.82, w: 5.6, h: 0.4, fontFace: FONT, fontSize: 9.5, color: C.inkSoft, lineSpacing: 13, isTextBox: true, margin: 0 });

  footer(s, "09 / 12", false);
}

// ------------------------------------------------------------
// P10 GO-TO-MARKET
// ------------------------------------------------------------
{
  const s = newSlide(pres, false);
  sectionTag(s, "Go-to-Market", false);
  title(s, "많이 만들기보다, 먼저 검증합니다.", { size: 26, y: 0.92, h: 0.55 });

  const steps = [
    ["01", "IP FOUNDATION", "기본 캐릭터 형태와\n디자인 가이드 확정"],
    ["02", "PROTOTYPE", "ORIGIN 4종 아트토이\n시제품 제작"],
    ["03", "CONTENT TEST", "SNS·숏폼·Question\nContent, 캐릭터 반응 확인"],
    ["04", "SMALL BATCH", "소량 제품 제작,\n판매 테스트"],
    ["05", "OFFLINE", "팝업, 디자인마켓,\n전시"],
    ["06", "SCALE-UP", "검증된 캐릭터 중심 생산\n확대, 협업·라이선싱"],
  ];
  const ty = 2.35;
  const gap = 0.2;
  const cellW = (CW - gap * (steps.length - 1)) / steps.length;
  s.addShape("line", { x: MX + cellW / 2, y: ty + 0.22, w: CW - cellW, h: 0, line: { color: C.line, width: 1 } });
  steps.forEach((st, i) => {
    const x = MX + i * (cellW + gap);
    s.addShape("ellipse", { x: x + cellW / 2 - 0.22, y: ty, w: 0.44, h: 0.44, fill: { color: C.bgLight }, line: { color: C.gold, width: 1.5 } });
    s.addText(st[0], { x: x + cellW / 2 - 0.22, y: ty, w: 0.44, h: 0.44, fontFace: FONT, fontSize: 11, bold: true, color: C.gold, align: "center", valign: "middle", isTextBox: true, margin: 0 });
    s.addText(st[1], { x, y: ty + 0.62, w: cellW, h: 0.4, fontFace: FONT, fontSize: 9.5, bold: true, color: C.ink, align: "center", isTextBox: true, margin: 0, charSpacing: 0.3 });
    s.addText(st[2], { x, y: ty + 1.0, w: cellW, h: 0.85, fontFace: FONT, fontSize: 8.3, color: C.inkSoft, align: "center", lineSpacing: 11.5, isTextBox: true, margin: 0 });
  });

  placeholder(s, MX, 4.65, CW, 2.0, "Prototype / Content Test 진행 예시 이미지 (검증 예정)", false);

  footer(s, "10 / 12", false);
}

// ------------------------------------------------------------
// P11 CURRENT STATUS + ROADMAP
// ------------------------------------------------------------
{
  const s = newSlide(pres, false);
  sectionTag(s, "Current Status / Roadmap", false);
  title(s, "현재는 IP 개발과 제품화 준비 단계입니다.", { size: 22, y: 0.92, h: 0.5 });

  const curX = MX, valX = MX + 5.9, topY = 1.65, colW = 5.55;
  s.addText("CURRENT — 현재 진행된 부분", { x: curX, y: topY, w: colW, h: 0.22, fontFace: FONT, fontSize: 9.5, bold: true, color: C.gold, charSpacing: 0.5, isTextBox: true, margin: 0 });
  bullets(s, [
    "브랜드 콘셉트 정립", "Would you, Tokki? 네이밍 확정", "메인 조형 언어 개발", "기본 캐릭터 시스템 기획",
    "ORIGIN 4종 기획", "Sea / Bubble / Food / Angora 등 확장 캐릭터 개발",
    "Material → Question 브랜드 경험 구조 기획", "2D·3D 비주얼 테스트", "아트토이 제품 방향 설정", "브랜드·사업 제안서 초안 작성",
  ], curX, topY + 0.26, colW, 2.0, { size: 8.7, spaceAfter: 2 });

  s.addText("TO VALIDATE — 아직 검증이 필요한 부분", { x: valX, y: topY, w: colW, h: 0.22, fontFace: FONT, fontSize: 9.5, bold: true, color: C.gold, charSpacing: 0.5, isTextBox: true, margin: 0 });
  bullets(s, [
    "실제 제작단가", "적정 판매가", "소재 및 생산방식", "첫 출시 캐릭터 수",
    "실제 소비자 반응", "판매채널", "손익구조", "생산 파트너",
  ], valX, topY + 0.26, colW, 2.0, { size: 8.7, spaceAfter: 2 });

  // Roadmap
  const ry = 4.1;
  s.addText("12 MONTH ROADMAP", { x: MX, y: ry, w: CW, h: 0.22, fontFace: FONT, fontSize: 9.5, bold: true, color: C.inkFaint, charSpacing: 1.5, isTextBox: true, margin: 0 });
  const phases = [
    ["PHASE 1", "FOUNDATION", "1–3개월", ["디자인 시스템 확정", "브랜드 가이드 정리", "상표·저작권 검토", "3D 모델링"]],
    ["PHASE 2", "PROTOTYPE", "3–6개월", ["ORIGIN 4종 시제품 제작", "재질 테스트", "크기 테스트", "패키지 개발"]],
    ["PHASE 3", "VALIDATION", "6–9개월", ["SNS 콘텐츠 테스트", "Question Content 테스트", "디자인마켓 · 팝업", "소량 판매"]],
    ["PHASE 4", "LAUNCH", "9–12개월", ["대표 캐릭터 선정", "정식 제품 생산", "온라인 판매채널 구축", "브랜드 협업·라이선싱"]],
  ];
  const pGap = 0.28;
  const pW = (CW - pGap * 3) / 4;
  const py2 = ry + 0.3;
  s.addShape("line", { x: MX, y: py2, w: CW, h: 0, line: { color: C.line, width: 0.75 } });
  phases.forEach((p, i) => {
    const x = MX + i * (pW + pGap);
    s.addText([
      { text: p[0] + "  ", options: { bold: true, color: C.gold, fontSize: 9.5, charSpacing: 0.5 } },
      { text: p[1], options: { bold: true, color: C.ink, fontSize: 9.5, charSpacing: 0.5 } },
    ], { x, y: py2 + 0.1, w: pW, h: 0.24, fontFace: FONT, isTextBox: true, margin: 0 });
    s.addText(p[2], { x, y: py2 + 0.34, w: pW, h: 0.2, fontFace: FONT, fontSize: 8, color: C.inkFaint, isTextBox: true, margin: 0 });
    bullets(s, p[3], x, py2 + 0.6, pW, 1.2, { size: 8, spaceAfter: 2 });
  });

  footer(s, "11 / 12", false);
}

// ------------------------------------------------------------
// P12 ASK
// ------------------------------------------------------------
{
  const s = newSlide(pres, true);
  sectionTag(s, "Ask", true);
  title(s, "현재 가장 필요한 것은\n더 많은 캐릭터가 아니라 검증입니다.", { size: 22, y: 0.92, h: 0.95, dark: true, lineSpacing: 27 });

  body(s, "Would you, Tokki?는 현재 디자인 아이디어에서 실제 사업으로 넘어가기 위한 전환점에 있습니다. 이번 컨설팅에서 확인하고 싶은 내용:", MX, 2.0, CW, 0.55, { size: 10.5, dark: true });

  const questions = [
    "현재 사업모델에서 가장 부족한 요소는 무엇인가",
    "지원사업 및 투자 관점에서 설득력을 갖추려면 어떤 데이터와 검증이 필요한가",
    "첫 제품을 ORIGIN 4종으로 시작하는 전략이 적절한가",
    "1인 IP 브랜드가 제조를 외주화하는 구조가 적절한가",
    "시장검증 전 사업자등록이 필요한가",
    "초기 고객 반응을 어떤 방식으로 측정하는 것이 적절한가",
    "Material → Question 구조가 브랜드 차별화와 고객 경험으로 실질적인 가치가 있는가",
    "현재 IR Deck에서 사업적으로 가장 약한 부분은 무엇인가",
  ];
  const qY = 2.68, colW = (CW - 0.5) / 2, rowH = 0.95;
  questions.forEach((q, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = MX + col * (colW + 0.5);
    const y = qY + row * rowH;
    s.addText(String(i + 1).padStart(2, "0"), { x, y, w: 0.4, h: 0.8, fontFace: FONT, fontSize: 13, bold: true, color: C.goldOnDark, isTextBox: true, margin: 0 });
    s.addText(q, { x: x + 0.42, y, w: colW - 0.42, h: 0.8, fontFace: FONT, fontSize: 9.5, color: C.cream, lineSpacing: 13, isTextBox: true, margin: 0 });
  });

  s.addText("Would you, Tokki?는 귀여운 캐릭터를 많이 만드는 브랜드가 아니라, 하나의 조형 시스템과 서로 다른 물질, 그리고 질문을 통해 각자의 세계를 수집하게 만드는 캐릭터 IP를 목표로 합니다.", {
    x: MX, y: 6.28, w: CW, h: 0.5, fontFace: FONT, fontSize: 9, color: C.creamMute, lineSpacing: 12.5, isTextBox: true, margin: 0,
  });

  s.addText("A rabbit. A material. A question.   —   Born from different worlds.", {
    x: MX, y: 6.82, w: CW, h: 0.25, fontFace: FONT, fontSize: 9.5, italic: true, color: C.goldOnDark,
    isTextBox: true, margin: 0,
  });

  footer(s, "12 / 12", true);
}

pres.writeFile({ fileName: "/tmp/claude-0/-home-user-hoho/af0e2641-eed0-5969-9532-d0b21417fcdb/scratchpad/tokki_ir/WOULD_YOU_TOKKI_IR_DECK_12P.pptx" }).then(() => {
  console.log("DONE");
});
