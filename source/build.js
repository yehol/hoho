const path = require("path");
const pptxgen = require("pptxgenjs");

// ============================================================
// DESIGN SYSTEM
// ============================================================
const FONT = "Noto Sans KR";
const A = (f) => path.join(__dirname, "..", "assets", f);

const C = {
  bgLight: "FAF7F1",
  bgDark: "1C1915",
  ink: "2A2622",
  inkSoft: "5B564D",
  inkFaint: "8A8276",
  line: "DDD5C7",
  lineDark: "463F37",
  cream: "F3EEE6",
  creamMute: "B8AFA0",
  gold: "A6813F",
  goldOnDark: "D0AE6B",
  sand: "BFA06E",
  paleBlue: "8FA9B3",
  seaGlass: "77998C",
  card: "F1EBE0",
};

// native pixel ratios (w / h) of every asset — used to place images without distortion
const R = {
  cover: 1.7768,
  strip: 4.8810,
  hero: 1.2202,
  thumb: 1.0400,
  grid: 1.7162,
  shelf: 1.5000,
  band: 5.6889,
  turn: 3.2263,
  detail: 5.9146,
  size: 1.9535,
};

const W = 13.333, H = 7.5;
const MX = 0.9;
const CW = W - MX * 2;
const HEADER_Y = 0.55;
const FOOTER_Y = 7.12;

function newSlide(pres, dark) {
  const s = pres.addSlide();
  s.background = { color: dark ? C.bgDark : C.bgLight };
  return s;
}

// place an image by width (h derived) or by height (w derived) so the aspect never distorts
function imgW(s, file, x, y, w, ratio) {
  s.addImage({ path: A(file), x, y, w, h: w / ratio });
  return w / ratio;
}
function imgH(s, file, x, y, h, ratio) {
  s.addImage({ path: A(file), x, y, w: h * ratio, h });
  return h * ratio;
}

function footer(s, pageNum, dark, brand = true) {
  const fg = dark ? C.creamMute : C.inkFaint;
  if (brand) {
    s.addText("WOULD YOU, TOKKI?", {
      x: MX, y: FOOTER_Y, w: 4, h: 0.25, fontFace: FONT, fontSize: 8,
      color: fg, charSpacing: 2, isTextBox: true, margin: 0,
    });
  }
  s.addText(pageNum, {
    x: W - MX - 1.5, y: FOOTER_Y, w: 1.5, h: 0.25, fontFace: FONT, fontSize: 8,
    color: fg, align: "right", charSpacing: 1, isTextBox: true, margin: 0,
  });
}

function sectionTag(s, label, dark) {
  const col = dark ? C.goldOnDark : C.gold;
  s.addShape("ellipse", {
    x: MX, y: HEADER_Y + 0.065, w: 0.09, h: 0.09, fill: { color: col }, line: { type: "none" },
  });
  s.addText(label.toUpperCase(), {
    x: MX + 0.2, y: HEADER_Y, w: 6, h: 0.22, fontFace: FONT, fontSize: 10.5, bold: true,
    color: col, charSpacing: 2.5, isTextBox: true, margin: 0,
  });
}

function title(s, text, opts = {}) {
  const dark = opts.dark || false;
  const size = opts.size || 29;
  s.addText(text, {
    x: opts.x !== undefined ? opts.x : MX, y: opts.y || 0.84,
    w: opts.w || CW, h: opts.h || 0.58,
    fontFace: FONT, fontSize: size, bold: true,
    color: dark ? C.cream : C.ink,
    lineSpacing: opts.lineSpacing || size * 1.24,
    isTextBox: true, margin: 0, align: opts.align || "left",
  });
}

function body(s, text, x, y, w, h, opts = {}) {
  s.addText(text, {
    x, y, w, h, fontFace: FONT, fontSize: opts.size || 11,
    color: opts.color || (opts.dark ? C.creamMute : C.inkSoft),
    lineSpacing: opts.lineSpacing || (opts.size || 11) * 1.45,
    italic: opts.italic || false, bold: opts.bold || false,
    isTextBox: true, margin: 0, align: opts.align || "left", valign: opts.valign || "top",
  });
}

function label(s, text, x, y, w, opts = {}) {
  s.addText(text, {
    x, y, w, h: opts.h || 0.22, fontFace: FONT, fontSize: opts.size || 10, bold: true,
    color: opts.color || C.gold, charSpacing: opts.charSpacing !== undefined ? opts.charSpacing : 1.5,
    align: opts.align || "left", isTextBox: true, margin: 0,
  });
}

function rule(s, x, y, w, dark) {
  s.addShape("line", { x, y, w, h: 0, line: { color: dark ? C.lineDark : C.line, width: 0.75 } });
}

function bullets(s, items, x, y, w, h, opts = {}) {
  const dark = opts.dark || false;
  const arr = items.map((t, i) => ({
    text: t,
    options: {
      bullet: { code: "2013", indent: 13 },
      breakLine: i < items.length - 1,
      paraSpaceAfter: opts.spaceAfter !== undefined ? opts.spaceAfter : 3,
    },
  }));
  s.addText(arr, {
    x, y, w, h, fontFace: FONT, fontSize: opts.size || 9.5,
    color: opts.color || (dark ? C.creamMute : C.inkSoft),
    isTextBox: true, margin: 0, valign: "top",
  });
}

// ============================================================
const pres = new pptxgen();
pres.defineLayout({ name: "TOKKI", width: W, height: H });
pres.layout = "TOKKI";

// ------------------------------------------------------------
// P1 COVER — full-bleed cosmos + ORIGIN specimen card
// ------------------------------------------------------------
{
  const s = pres.addSlide();
  s.background = { path: A("cover_space.jpg") };

  s.addText("CHARACTER ART TOY IP", {
    x: MX, y: 3.42, w: 6.0, h: 0.26, fontFace: FONT, fontSize: 11, bold: true,
    color: C.goldOnDark, charSpacing: 3.5, isTextBox: true, margin: 0,
  });
  s.addText("Would you, Tokki?", {
    x: MX, y: 3.80, w: 6.4, h: 0.8, fontFace: FONT, fontSize: 42, bold: true,
    color: C.cream, isTextBox: true, margin: 0, lineSpacing: 48,
  });
  s.addText("우주…토끼?", {
    x: MX, y: 4.66, w: 6.4, h: 0.6, fontFace: FONT, fontSize: 30, bold: true,
    color: C.cream, isTextBox: true, margin: 0, lineSpacing: 36,
  });
  s.addText("Born from different worlds.", {
    x: MX, y: 5.38, w: 6.0, h: 0.3, fontFace: FONT, fontSize: 14, italic: true,
    color: C.goldOnDark, isTextBox: true, margin: 0,
  });
  s.addText(
    "하나의 기본 토끼 조형에 서로 다른 세계의 물성과 환경을 적용해\n확장하는 수집형 캐릭터 아트토이 IP.",
    {
      x: MX, y: 5.78, w: 6.2, h: 0.72, fontFace: FONT, fontSize: 11,
      color: C.creamMute, lineSpacing: 17, isTextBox: true, margin: 0, valign: "top",
    }
  );
  s.addText("예비창업 단계 · IP 개발 및 제품화 준비 중", {
    x: MX, y: 6.62, w: 6.0, h: 0.28, fontFace: FONT, fontSize: 9.5,
    color: C.creamMute, charSpacing: 0.5, isTextBox: true, margin: 0,
  });

  // specimen card
  const cx = 7.55, cy = 3.45, cw = 4.88, ch = 2.95;
  s.addShape("rect", { x: cx, y: cy, w: cw, h: ch, fill: { color: C.card }, line: { type: "none" } });
  label(s, "COLLECTION 01 — ORIGIN", cx + 0.25, cy + 0.25, cw - 0.5, { size: 8.5, charSpacing: 2 });
  const stripW = cw - 0.5;
  imgW(s, "heroes_strip.jpg", cx + 0.25, cy + 0.60, stripW, R.strip);
  const cellW = stripW / 4;
  [["MOON", "달토끼"], ["ICE", "얼음토끼"], ["ROCK", "암석토끼"], ["SAND", "모래토끼"]].forEach((m, i) => {
    const lx = cx + 0.25 + i * cellW;
    s.addText(m[0], {
      x: lx, y: cy + 1.62, w: cellW, h: 0.2, fontFace: FONT, fontSize: 8.5, bold: true,
      color: C.ink, align: "center", isTextBox: true, margin: 0, charSpacing: 0.8,
    });
    s.addText(m[1], {
      x: lx, y: cy + 1.83, w: cellW, h: 0.2, fontFace: FONT, fontSize: 7.5,
      color: C.inkFaint, align: "center", isTextBox: true, margin: 0,
    });
  });
  s.addText("네 가지 기본 물성으로 시작하는 첫 번째 컬렉션.", {
    x: cx + 0.25, y: cy + 2.16, w: stripW, h: 0.28, fontFace: FONT, fontSize: 8.5, italic: true,
    color: C.inkSoft, align: "center", isTextBox: true, margin: 0,
  });

  footer(s, "01 / 12", true);
}

// ------------------------------------------------------------
// P2 PROBLEM
// ------------------------------------------------------------
{
  const s = newSlide(pres, false);
  sectionTag(s, "Problem", false);
  title(s, "캐릭터는 많지만,\n오래 소유하고 싶은 캐릭터는 적다.", { size: 29, h: 1.2, lineSpacing: 38 });

  const LW = 5.3;
  label(s, "CHARACTER GOODS", MX, 2.32, LW, { size: 9.5 });
  s.addText("캐릭터 굿즈", {
    x: MX, y: 2.56, w: LW, h: 0.3, fontFace: FONT, fontSize: 14, bold: true, color: C.ink,
    isTextBox: true, margin: 0,
  });
  body(s, "다양한 그래픽과 굿즈는 많지만, 성인 소비자가 생활공간에 오래 두고 싶은 조형적 오브젝트형 캐릭터는 선택지가 제한적입니다.", MX, 2.92, LW, 0.85, { size: 10.5 });

  rule(s, MX, 3.92, LW, false);

  label(s, "COLLECTOR ART TOY", MX, 4.08, LW, { size: 9.5 });
  s.addText("컬렉터 아트토이", {
    x: MX, y: 4.32, w: LW, h: 0.3, fontFace: FONT, fontSize: 14, bold: true, color: C.ink,
    isTextBox: true, margin: 0,
  });
  body(s, "작품성과 수집가치는 높지만, 가격과 접근성에서 진입장벽이 존재합니다.", MX, 4.68, LW, 0.6, { size: 10.5 });

  rule(s, MX, 5.42, LW, false);
  s.addText("Would you, Tokki?는\n이 세 영역의 접점을 지향합니다.", {
    x: MX, y: 5.62, w: LW, h: 0.9, fontFace: FONT, fontSize: 15, bold: true, color: C.ink,
    lineSpacing: 21, isTextBox: true, margin: 0,
  });

  // Venn — fixed geometry, bounds x 7.15–11.45 / y 2.30–6.10
  const r = 1.2;
  const circles = [
    { cx: 8.35, cy: 3.50, color: C.paleBlue, text: "CHARACTER", ly: 2.72 },
    { cx: 10.25, cy: 3.50, color: C.sand, text: "ART TOY", ly: 2.72 },
    { cx: 9.30, cy: 4.90, color: C.seaGlass, text: "LIFESTYLE\nOBJECT", ly: 5.42 },
  ];
  circles.forEach((p) => {
    s.addShape("ellipse", {
      x: p.cx - r, y: p.cy - r, w: r * 2, h: r * 2,
      fill: { color: p.color, transparency: 87 },
      line: { color: p.color, width: 1.25 },
    });
  });
  circles.forEach((p) => {
    s.addText(p.text, {
      x: p.cx - 1.05, y: p.ly, w: 2.1, h: 0.52, fontFace: FONT, fontSize: 10, bold: true,
      color: C.ink, align: "center", valign: "top", isTextBox: true, margin: 0,
      lineSpacing: 12.5, charSpacing: 0.5,
    });
  });

  footer(s, "02 / 12", false);
}

// ------------------------------------------------------------
// P3 SOLUTION — the full lineup, front and centre
// ------------------------------------------------------------
{
  const s = newSlide(pres, false);
  sectionTag(s, "Solution", false);
  title(s, "같은 토끼, 다른 세계.", { size: 30, w: 6.0, h: 0.56 });
  s.addText("같은 형태, 다른 물성.\n하나의 조형 시스템을 계속 변주해 세계를 넓혀갑니다.", {
    x: 6.9, y: 0.88, w: 5.53, h: 0.6, fontFace: FONT, fontSize: 10.5, color: C.inkSoft,
    align: "right", lineSpacing: 15.5, isTextBox: true, margin: 0,
  });

  imgW(s, "grid18.jpg", 2.12, 1.58, 9.10, R.grid); // h = 5.30 → bottom 6.88

  footer(s, "03 / 12", false);
}

// ------------------------------------------------------------
// P4 PRODUCT
// ------------------------------------------------------------
{
  const s = newSlide(pres, false);
  sectionTag(s, "Product", false);
  title(s, "COLLECTION 01 — ORIGIN", { size: 29, h: 0.52 });
  s.addText("첫 번째 제품은 네 가지 기본 물성을 가진 아트토이로 시작합니다.", {
    x: MX, y: 1.40, w: CW, h: 0.3, fontFace: FONT, fontSize: 12.5, color: C.inkSoft,
    isTextBox: true, margin: 0,
  });

  const hw = 2.68, hgap = 0.271;
  const four = [
    ["hero_moon.jpg", "MOON", "달토끼 · 크레이터"],
    ["hero_ice.jpg", "ICE", "얼음토끼 · 얼음 결정"],
    ["hero_rock.jpg", "ROCK", "암석토끼 · 갈라진 암석"],
    ["hero_sand.jpg", "SAND", "모래토끼 · 퇴적층"],
  ];
  four.forEach((m, i) => {
    const x = MX + i * (hw + hgap);
    imgW(s, m[0], x, 1.85, hw, R.hero); // h = 2.196 → bottom 4.046
    s.addText(m[1], {
      x, y: 4.13, w: hw, h: 0.26, fontFace: FONT, fontSize: 12.5, bold: true, color: C.ink,
      align: "center", isTextBox: true, margin: 0, charSpacing: 1.2,
    });
    s.addText(m[2], {
      x, y: 4.40, w: hw, h: 0.24, fontFace: FONT, fontSize: 9.5, color: C.inkFaint,
      align: "center", isTextBox: true, margin: 0,
    });
  });

  rule(s, MX, 4.85, CW, false);

  label(s, "COMMON BASE", MX, 5.05, 3.3);
  bullets(s, ["둥근 구형 바디", "짧은 팔과 다리", "토끼 귀", "작은 얼굴 요소", "단순하고 기억하기 쉬운 실루엣"],
    MX, 5.33, 3.3, 1.1, { size: 9.5 });

  label(s, "VARIATION", 4.5, 5.05, 3.3);
  bullets(s, ["표면 재질", "투명도", "균열", "기포", "퇴적층", "주변 파편", "궤도 오브젝트"],
    4.5, 5.33, 3.3, 1.4, { size: 9.5 });

  label(s, "SIZE SYSTEM", 8.1, 5.05, 4.33);
  imgW(s, "size_variation.jpg", 8.77, 5.33, 3.0, R.size); // h = 1.536 → bottom 6.87

  footer(s, "04 / 12", false);
}

// ------------------------------------------------------------
// P5 MATERIAL → QUESTION
// ------------------------------------------------------------
{
  const s = newSlide(pres, false);
  sectionTag(s, "Brand Experience", false);
  title(s, "A rabbit.\nA material.\nA question.", { size: 30, w: 4.3, h: 1.95, lineSpacing: 38 });

  body(s, "물질은 캐릭터의 외형을 만드는 데서 끝나지 않습니다. 각 물질의 실제 성질에서 출발한 짧은 질문이 캐릭터와 사용자의 경험을 연결합니다.",
    MX, 2.95, 4.0, 1.2, { size: 10.5 });

  s.addShape("rect", { x: MX, y: 4.30, w: 4.0, h: 1.15, fill: { color: C.card }, line: { type: "none" } });
  s.addText("이것은 심리검사나 성격유형 테스트가 아닙니다. 정답을 제시하지 않고, 사용자가 물질을 매개로 자신의 경험을 자유롭게 해석하게 합니다.", {
    x: MX + 0.2, y: 4.44, w: 3.6, h: 0.9, fontFace: FONT, fontSize: 9, italic: true, color: C.inkSoft,
    lineSpacing: 13, isTextBox: true, margin: 0,
  });

  label(s, "EXPANDS INTO", MX, 5.68, 4.0, { size: 8.5, color: C.inkFaint });
  s.addText("Art Toy → Question Card → Art Book → Exhibition → Web Experience → Question Archive", {
    x: MX, y: 5.92, w: 4.0, h: 0.62, fontFace: FONT, fontSize: 8.3, color: C.inkSoft,
    lineSpacing: 12, isTextBox: true, margin: 0, valign: "top",
  });

  // flow
  const FX = 5.35, FW = 7.083;
  const stages = ["MATERIAL", "PHYSICAL\nPROPERTY", "QUESTION", "PERSONAL\nINTERPRETATION"];
  const sw = (FW - 3 * 0.30) / 4;
  stages.forEach((st, i) => {
    const x = FX + i * (sw + 0.30);
    s.addShape("rect", { x, y: 1.10, w: sw, h: 0.72, fill: { color: C.card }, line: { color: C.line, width: 1 } });
    s.addText(st, {
      x, y: 1.10, w: sw, h: 0.72, fontFace: FONT, fontSize: 8.3, bold: true, color: C.ink,
      align: "center", valign: "middle", lineSpacing: 10.5, isTextBox: true, margin: 0, charSpacing: 0.3,
    });
    if (i < 3) {
      s.addText("→", {
        x: x + sw, y: 1.10, w: 0.30, h: 0.72, fontFace: FONT, fontSize: 13, color: C.inkFaint,
        align: "center", valign: "middle", isTextBox: true, margin: 0,
      });
    }
  });

  const examples = [
    ["r_rock.jpg", "ROCK", "단단함 · 무게 · 오래 지속됨", "당신을 오래 버티게 한 것은 무엇인가요?"],
    ["r_cloud.jpg", "CLOUD", "가벼움 · 이동 · 머무르지 않음", "잠시 멈춰도 된다고 생각하나요?"],
    ["r_bubble.jpg", "BUBBLE", "투명함 · 가벼움 · 짧은 지속성", "짧았지만 오래 기억되는 순간이 있나요?"],
  ];
  examples.forEach((ex, i) => {
    const y = 2.30 + i * 1.42;
    imgW(s, ex[0], FX, y, 0.95, R.thumb); // h = 0.913
    s.addText(ex[1], {
      x: 6.45, y: y + 0.06, w: 1.9, h: 0.24, fontFace: FONT, fontSize: 10, bold: true, color: C.ink,
      isTextBox: true, margin: 0, charSpacing: 0.8,
    });
    s.addText(ex[2], {
      x: 6.45, y: y + 0.34, w: 1.9, h: 0.55, fontFace: FONT, fontSize: 8.5, color: C.inkFaint,
      lineSpacing: 11.5, isTextBox: true, margin: 0, valign: "top",
    });
    s.addText(`“${ex[3]}”`, {
      x: 8.50, y: y, w: 3.93, h: 0.92, fontFace: FONT, fontSize: 13.5, bold: true, italic: true,
      color: C.ink, valign: "middle", lineSpacing: 18, isTextBox: true, margin: 0,
    });
  });

  footer(s, "05 / 12", false);
}

// ------------------------------------------------------------
// P6 POSITIONING
// ------------------------------------------------------------
{
  const s = newSlide(pres, false);
  sectionTag(s, "Positioning", false);
  title(s, "아이에게는 캐릭터, 어른에게는 오브제.", { size: 28, h: 0.56 });

  const LW = 4.3;
  body(s, "어린이 캐릭터와 고가 컬렉터 아트토이 사이의 중간 지점을 지향합니다. 캐릭터를 좋아하지만 지나치게 아동적이지 않고, 자신의 책상과 선반에 자연스럽게 놓을 수 있는 조형적인 제품을 원하는 소비자를 초기 고객으로 설정합니다.",
    MX, 1.62, LW, 2.4, { size: 11 });

  rule(s, MX, 3.45, LW, false);
  s.addText("Character  ×  Art Toy  ×  Lifestyle Object", {
    x: MX, y: 3.68, w: LW, h: 0.36, fontFace: FONT, fontSize: 14, bold: true, color: C.ink,
    isTextBox: true, margin: 0,
  });
  s.addText("CUTE\nOBJECT\nCOLLECTIBLE\nMATERIAL\nLIFESTYLE", {
    x: MX, y: 4.24, w: LW, h: 1.3, fontFace: FONT, fontSize: 10.5, color: C.inkFaint,
    charSpacing: 1.5, lineSpacing: 18, isTextBox: true, margin: 0, valign: "top",
  });
  s.addText("생활공간에 자연스럽게 놓이는 조형적 오브제.", {
    x: MX, y: 5.80, w: LW, h: 0.34, fontFace: FONT, fontSize: 11.5, italic: true, color: C.inkSoft,
    isTextBox: true, margin: 0,
  });

  imgW(s, "shelf.jpg", 5.55, 1.62, 6.88, R.shelf); // h = 4.587 → bottom 6.21
  s.addText("생활공간에 놓인 오브제로서의 제품 이미지", {
    x: 5.55, y: 6.32, w: 6.88, h: 0.26, fontFace: FONT, fontSize: 8.5, italic: true,
    color: C.inkFaint, align: "right", isTextBox: true, margin: 0,
  });

  footer(s, "06 / 12", false);
}

// ------------------------------------------------------------
// P7 TARGET
// ------------------------------------------------------------
{
  const s = newSlide(pres, false);
  sectionTag(s, "Target", false);
  title(s, "누가 Would you, Tokki?를 사는가", { size: 28, h: 0.56 });

  const LW = 5.2;
  label(s, "PRIMARY TARGET", MX, 1.72, LW, { size: 10 });
  s.addText("20–30대 디자인·캐릭터 소비자", {
    x: MX, y: 2.00, w: LW, h: 0.46, fontFace: FONT, fontSize: 18, bold: true, color: C.ink,
    isTextBox: true, margin: 0,
  });
  body(s, "캐릭터를 좋아하지만 지나치게 아동적인 제품보다 조형성과 디자인 완성도가 높은 제품을 선호하는 초기 예상 소비자층.",
    MX, 2.58, LW, 0.85, { size: 10.5 });

  label(s, "INTEREST", MX, 3.62, LW, { size: 9.5, color: C.inkFaint });
  s.addText("Character IP · Art Toy · Design Object · Stationery · Interior · Independent Brand · Exhibition · Pop-up", {
    x: MX, y: 3.88, w: LW, h: 0.7, fontFace: FONT, fontSize: 9.5, color: C.inkSoft,
    lineSpacing: 14, isTextBox: true, margin: 0, valign: "top",
  });

  label(s, "SECONDARY", MX, 4.78, LW, { size: 9.5, color: C.inkFaint });
  bullets(s, ["선물용 소비자", "아동과 함께 캐릭터 제품을 소비하는 가족층", "캐릭터 컬래버레이션이 필요한 브랜드·기업"],
    MX, 5.04, LW, 0.95, { size: 10 });

  s.addShape("rect", { x: MX, y: 6.16, w: LW, h: 0.66, fill: { color: C.card }, line: { type: "none" } });
  s.addText("실제 고객조사를 완료한 결과가 아닌, 초기 가설로서의 예상 핵심 고객입니다.", {
    x: MX + 0.18, y: 6.16, w: LW - 0.36, h: 0.66, fontFace: FONT, fontSize: 9, italic: true,
    color: C.inkSoft, valign: "middle", lineSpacing: 12, isTextBox: true, margin: 0,
  });

  // concentric rings, centre 9.9 / 3.9 — bounds x 7.85–11.95, y 1.85–5.95
  const cx = 9.9, cy = 3.9;
  [[2.05, C.line], [1.5, C.paleBlue], [0.9, C.gold]].forEach(([rr, col]) => {
    s.addShape("ellipse", {
      x: cx - rr, y: cy - rr, w: rr * 2, h: rr * 2,
      fill: { type: "none" }, line: { color: col, width: 1.25 },
    });
  });
  s.addImage({ path: A("r_mochi_round.jpg"), x: cx - 0.44, y: cy - 0.60, w: 0.88, h: 0.88 });
  s.addText("PRIMARY", {
    x: cx - 1.0, y: cy + 0.30, w: 2.0, h: 0.24, fontFace: FONT, fontSize: 9, bold: true,
    color: C.ink, align: "center", isTextBox: true, margin: 0, charSpacing: 1,
  });
  s.addText("SECONDARY", {
    x: cx - 1.4, y: cy + 1.52, w: 2.8, h: 0.22, fontFace: FONT, fontSize: 8.5,
    color: C.inkFaint, align: "center", isTextBox: true, margin: 0, charSpacing: 1,
  });

  footer(s, "07 / 12", false);
}

// ------------------------------------------------------------
// P8 DIFFERENTIATION
// ------------------------------------------------------------
{
  const s = newSlide(pres, false);
  sectionTag(s, "Differentiation", false);
  title(s, "Would you, Tokki?가 확장되는 방식", { size: 28, h: 0.56 });

  const flow = ["Base\nBunny", "Material", "Character", "Question", "Art Toy", "Object /\nSculpture"];
  const fw = (CW - 5 * 0.34) / 6;
  flow.forEach((st, i) => {
    const x = MX + i * (fw + 0.34);
    s.addShape("rect", { x, y: 1.62, w: fw, h: 0.70, fill: { color: C.card }, line: { color: C.line, width: 1 } });
    s.addText(st, {
      x, y: 1.62, w: fw, h: 0.70, fontFace: FONT, fontSize: 9.5, bold: true, color: C.ink,
      align: "center", valign: "middle", lineSpacing: 12, isTextBox: true, margin: 0,
    });
    if (i < 5) {
      s.addText("→", {
        x: x + fw, y: 1.62, w: 0.34, h: 0.70, fontFace: FONT, fontSize: 13, color: C.inkFaint,
        align: "center", valign: "middle", isTextBox: true, margin: 0,
      });
    }
  });

  const items = [
    ["01", "MATERIAL AS CHARACTER", "색이나 의상을 바꾸는 것이 아니라 물질 자체가 캐릭터 디자인이 됩니다."],
    ["02", "ONE BASE, MANY WORLDS", "하나의 기본 조형을 유지하면서 서로 다른 세계로 확장할 수 있습니다."],
    ["03", "3D FIRST", "처음부터 입체 조형과 실제 제품화를 고려해 디자인합니다."],
    ["04", "MATERIAL → QUESTION", "물질의 실제 성질을 사용자의 경험을 떠올리게 하는 열린 질문으로 확장합니다."],
    ["05", "ARTWORK → PRODUCT", "소형 아트토이뿐 아니라 대형 조각, 전시, 공간 설치까지 확장할 수 있습니다."],
  ];
  items.forEach((it, i) => {
    const y = 2.72 + i * 0.83;
    s.addText(it[0], {
      x: MX, y, w: 0.52, h: 0.7, fontFace: FONT, fontSize: 18, bold: true, color: C.gold,
      isTextBox: true, margin: 0,
    });
    s.addText(it[1], {
      x: 1.42, y, w: 4.88, h: 0.28, fontFace: FONT, fontSize: 11.5, bold: true, color: C.ink,
      isTextBox: true, margin: 0,
    });
    s.addText(it[2], {
      x: 1.42, y: y + 0.29, w: 4.88, h: 0.5, fontFace: FONT, fontSize: 9.3, color: C.inkSoft,
      lineSpacing: 12.5, isTextBox: true, margin: 0, valign: "top",
    });
  });

  const RX = 7.55, RW = 4.88;
  label(s, "PRODUCT SYSTEM", RX, 2.72, RW, { size: 9.5 });
  imgW(s, "turnaround.jpg", RX, 3.02, RW, R.turn); // h = 1.513 → bottom 4.533
  s.addText("FRONT · SIDE · BACK · TOP — 360° 조형 검토", {
    x: RX, y: 4.62, w: RW, h: 0.24, fontFace: FONT, fontSize: 8.3, color: C.inkFaint,
    align: "center", isTextBox: true, margin: 0,
  });
  imgW(s, "detail_macro.jpg", RX, 5.00, RW, R.detail); // h = 0.825 → bottom 5.825
  s.addText("표면 질감 · 크레이터 · 귀 단면 · 파편 디테일", {
    x: RX, y: 5.90, w: RW, h: 0.24, fontFace: FONT, fontSize: 8.3, color: C.inkFaint,
    align: "center", isTextBox: true, margin: 0,
  });
  s.addText("물성은 렌더링 효과가 아니라 실제 조형과 표면 설계로 구현됩니다.", {
    x: RX, y: 6.36, w: RW, h: 0.5, fontFace: FONT, fontSize: 9.3, italic: true, color: C.inkSoft,
    align: "center", lineSpacing: 13, isTextBox: true, margin: 0,
  });

  footer(s, "08 / 12", false);
}

// ------------------------------------------------------------
// P9 BUSINESS MODEL
// ------------------------------------------------------------
{
  const s = newSlide(pres, false);
  sectionTag(s, "Business Model", false);
  title(s, "IP를 중심으로 확장되는 구조", { size: 28, h: 0.52 });

  const MID = W / 2;
  s.addText("Would you, Tokki?", {
    x: MID - 1.5, y: 1.44, w: 3.0, h: 0.24, fontFace: FONT, fontSize: 10, bold: true,
    color: C.gold, align: "center", isTextBox: true, margin: 0, charSpacing: 0.5,
  });
  s.addImage({ path: A("r_moon_round.jpg"), x: MID - 0.46, y: 1.70, w: 0.92, h: 0.92 }); // bottom 2.62

  const brW = (CW - 4 * 0.2) / 5;
  s.addShape("line", { x: MID, y: 2.60, w: 0, h: 0.14, line: { color: C.line, width: 0.75 } });
  s.addShape("line", {
    x: MX + brW / 2, y: 2.74, w: CW - brW, h: 0, line: { color: C.line, width: 0.75 },
  });

  const branches = [
    ["D2C", "Art Toy · Design Object\nLimited Edition · Lifestyle Goods"],
    ["B2B", "Brand Collaboration · Corporate Goods\nExhibition · Space Collaboration"],
    ["LICENSING", "Character License\nProduct License · Content Collab."],
    ["ART", "Limited Art Toy\nLarge Sculpture · Exhibition Work"],
    ["CONTENT /\nEXPERIENCE", "Question Card · Art Book\nWeb Experience · Exhibition Archive"],
  ];
  branches.forEach((b, i) => {
    const x = MX + i * (brW + 0.2);
    s.addShape("line", { x: x + brW / 2, y: 2.74, w: 0, h: 0.14, line: { color: C.line, width: 0.75 } });
    s.addText(b[0], {
      x, y: 2.98, w: brW, h: 0.46, fontFace: FONT, fontSize: 11, bold: true, color: C.ink,
      align: "center", valign: "top", isTextBox: true, margin: 0, charSpacing: 0.3, lineSpacing: 13.5,
    });
    s.addText(b[1], {
      x, y: 3.54, w: brW, h: 0.64, fontFace: FONT, fontSize: 8, color: C.inkFaint,
      align: "center", valign: "top", lineSpacing: 11.5, isTextBox: true, margin: 0,
    });
  });

  rule(s, MX, 4.66, CW, false);

  label(s, "생산은 전문 제작 파트너와 함께합니다", MX, 4.86, CW, { size: 11, charSpacing: 0.5 });
  body(s, "브랜드는 초기부터 자체 제조시설을 구축하지 않습니다.", MX, 5.14, CW, 0.26, { size: 9.5 });

  label(s, "핵심 역량", MX, 5.55, 5.4, { size: 9, color: C.inkFaint, charSpacing: 1 });
  s.addText("IP 기획 · 캐릭터 디자인 · 아트디렉션 · 제품기획 · 브랜드 경험 설계", {
    x: MX, y: 5.80, w: 5.4, h: 0.5, fontFace: FONT, fontSize: 9.5, color: C.inkSoft,
    lineSpacing: 13.5, isTextBox: true, margin: 0, valign: "top",
  });
  label(s, "제작 파트너 협업", 6.6, 5.55, 5.4, { size: 9, color: C.inkFaint, charSpacing: 1 });
  s.addText("3D Modeling · 3D Printing · Mold · Production · Painting · Printing", {
    x: 6.6, y: 5.80, w: 5.4, h: 0.5, fontFace: FONT, fontSize: 9.5, color: C.inkSoft,
    lineSpacing: 13.5, isTextBox: true, margin: 0, valign: "top",
  });

  footer(s, "09 / 12", false);
}

// ------------------------------------------------------------
// P10 GO-TO-MARKET
// ------------------------------------------------------------
{
  const s = newSlide(pres, false);
  sectionTag(s, "Go-to-Market", false);
  title(s, "많이 만들기보다, 먼저 검증합니다.", { size: 28, h: 0.56 });

  const steps = [
    ["01", "IP FOUNDATION", "기본 캐릭터 형태와\n디자인 가이드 확정"],
    ["02", "PROTOTYPE", "ORIGIN 4종 아트토이\n시제품 제작"],
    ["03", "CONTENT TEST", "SNS·숏폼·Question\nContent 반응 확인"],
    ["04", "SMALL BATCH", "소량 제품 제작,\n판매 테스트"],
    ["05", "OFFLINE", "팝업, 디자인마켓,\n전시"],
    ["06", "SCALE-UP", "검증된 캐릭터 중심\n생산 확대, 라이선싱"],
  ];
  const cw6 = (CW - 5 * 0.2) / 6;
  s.addShape("line", {
    x: MX + cw6 / 2, y: 2.33, w: CW - cw6, h: 0, line: { color: C.line, width: 1 },
  });
  steps.forEach((st, i) => {
    const x = MX + i * (cw6 + 0.2);
    s.addShape("ellipse", {
      x: x + cw6 / 2 - 0.23, y: 2.10, w: 0.46, h: 0.46,
      fill: { color: C.bgLight }, line: { color: C.gold, width: 1.5 },
    });
    s.addText(st[0], {
      x: x + cw6 / 2 - 0.23, y: 2.10, w: 0.46, h: 0.46, fontFace: FONT, fontSize: 11, bold: true,
      color: C.gold, align: "center", valign: "middle", isTextBox: true, margin: 0,
    });
    s.addText(st[1], {
      x, y: 2.74, w: cw6, h: 0.3, fontFace: FONT, fontSize: 9.5, bold: true, color: C.ink,
      align: "center", isTextBox: true, margin: 0, charSpacing: 0.3,
    });
    s.addText(st[2], {
      x, y: 3.08, w: cw6, h: 0.8, fontFace: FONT, fontSize: 8.5, color: C.inkSoft,
      align: "center", valign: "top", lineSpacing: 12, isTextBox: true, margin: 0,
    });
  });

  imgW(s, "shelf_band.jpg", MX, 4.35, CW, R.band); // h = 2.027 → bottom 6.377
  s.addText("검증 단계에서는 완성된 전체 라인업이 아니라, 반응이 확인된 캐릭터부터 순차적으로 생산합니다.", {
    x: MX, y: 6.62, w: CW, h: 0.28, fontFace: FONT, fontSize: 9, italic: true, color: C.inkFaint,
    align: "center", isTextBox: true, margin: 0,
  });

  footer(s, "10 / 12", false);
}

// ------------------------------------------------------------
// P11 CURRENT STATUS + ROADMAP
// ------------------------------------------------------------
{
  const s = newSlide(pres, false);
  sectionTag(s, "Current Status / Roadmap", false);
  title(s, "현재는 IP 개발과 제품화 준비 단계입니다.", { size: 26, h: 0.52 });

  const colW = 5.4;
  label(s, "CURRENT — 현재 진행된 부분", MX, 1.55, colW, { size: 10, charSpacing: 0.5 });
  bullets(s, [
    "브랜드 콘셉트 정립", "Would you, Tokki? 네이밍 확정", "메인 조형 언어 개발",
    "기본 캐릭터 시스템 기획", "ORIGIN 4종 기획",
    "Sea / Bubble / Food / Angora 등 확장 캐릭터 개발",
    "Material → Question 브랜드 경험 구조 기획", "2D·3D 비주얼 테스트",
    "아트토이 제품 방향 설정", "브랜드·사업 제안서 초안 작성",
  ], MX, 1.83, colW, 2.6, { size: 9.8, spaceAfter: 5 });

  label(s, "TO VALIDATE — 아직 검증이 필요한 부분", 6.6, 1.55, colW, { size: 10, charSpacing: 0.5 });
  bullets(s, [
    "실제 제작단가", "적정 판매가", "소재 및 생산방식", "첫 출시 캐릭터 수",
    "실제 소비자 반응", "판매채널", "손익구조", "생산 파트너",
  ], 6.6, 1.83, colW, 2.6, { size: 9.8, spaceAfter: 5 });

  label(s, "12 MONTH ROADMAP", MX, 4.55, CW, { size: 10, color: C.inkFaint, charSpacing: 2 });
  rule(s, MX, 4.85, CW, false);

  const phases = [
    ["PHASE 1", "FOUNDATION", "1–3개월", ["디자인 시스템 확정", "브랜드 가이드 정리", "상표·저작권 검토", "3D 모델링"]],
    ["PHASE 2", "PROTOTYPE", "3–6개월", ["ORIGIN 4종 시제품 제작", "재질 테스트", "크기 테스트", "패키지 개발"]],
    ["PHASE 3", "VALIDATION", "6–9개월", ["SNS 콘텐츠 테스트", "Question Content 테스트", "디자인마켓 · 팝업", "소량 판매"]],
    ["PHASE 4", "LAUNCH", "9–12개월", ["대표 캐릭터 선정", "정식 제품 생산", "온라인 판매채널 구축", "브랜드 협업·라이선싱"]],
  ];
  const pW = (CW - 3 * 0.28) / 4;
  phases.forEach((p, i) => {
    const x = MX + i * (pW + 0.28);
    s.addText([
      { text: p[0] + "  ", options: { bold: true, color: C.gold, fontSize: 9.5, charSpacing: 0.5 } },
      { text: p[1], options: { bold: true, color: C.ink, fontSize: 9.5, charSpacing: 0.5 } },
    ], { x, y: 5.00, w: pW, h: 0.24, fontFace: FONT, isTextBox: true, margin: 0 });
    s.addText(p[2], {
      x, y: 5.26, w: pW, h: 0.22, fontFace: FONT, fontSize: 8.3, color: C.inkFaint,
      isTextBox: true, margin: 0,
    });
    bullets(s, p[3], x, 5.54, pW, 0.95, { size: 9, spaceAfter: 4 });
  });

  s.addText("본 로드맵은 현재 계획이며, 각 단계의 검증 결과에 따라 조정될 수 있습니다.", {
    x: MX, y: 6.62, w: CW, h: 0.28, fontFace: FONT, fontSize: 9.5, italic: true, color: C.inkFaint,
    align: "center", isTextBox: true, margin: 0,
  });

  footer(s, "11 / 12", false);
}

// ------------------------------------------------------------
// P12 ASK
// ------------------------------------------------------------
{
  const s = pres.addSlide();
  s.background = { path: A("cover_space.jpg") };
  s.addShape("rect", {
    x: 0, y: 0, w: W, h: H, fill: { color: C.bgDark, transparency: 28 }, line: { type: "none" },
  });

  sectionTag(s, "Ask", true);
  title(s, "현재 가장 필요한 것은\n더 많은 캐릭터가 아니라 검증입니다.", {
    size: 26, h: 1.0, dark: true, lineSpacing: 33,
  });

  body(s, "Would you, Tokki?는 디자인 아이디어에서 실제 사업으로 넘어가기 위한 전환점에 있습니다. 이번 컨설팅에서 확인하고 싶은 내용:",
    MX, 2.02, CW, 0.4, { size: 11, dark: true });

  const questions = [
    "현재 사업모델에서 가장 부족한 요소는 무엇인가",
    "지원사업·투자 관점에서 어떤 데이터와 검증이 필요한가",
    "첫 제품을 ORIGIN 4종으로 시작하는 전략이 적절한가",
    "1인 IP 브랜드가 제조를 외주화하는 구조가 적절한가",
    "시장검증 전 사업자등록이 필요한가",
    "초기 고객 반응을 어떤 방식으로 측정하는 것이 적절한가",
    "Material → Question 구조가 실질적인 차별화 가치가 있는가",
    "현재 IR Deck에서 사업적으로 가장 약한 부분은 무엇인가",
  ];
  const qcw = (CW - 0.5) / 2;
  questions.forEach((q, i) => {
    const x = MX + (i % 2) * (qcw + 0.5);
    const y = 2.70 + Math.floor(i / 2) * 0.92;
    s.addText(String(i + 1).padStart(2, "0"), {
      x, y, w: 0.42, h: 0.7, fontFace: FONT, fontSize: 13, bold: true, color: C.goldOnDark,
      isTextBox: true, margin: 0, valign: "top",
    });
    s.addText(q, {
      x: x + 0.46, y, w: qcw - 0.46, h: 0.78, fontFace: FONT, fontSize: 10, color: C.cream,
      lineSpacing: 14, isTextBox: true, margin: 0, valign: "top",
    });
  });

  s.addText("귀여운 캐릭터를 많이 만드는 브랜드가 아니라, 하나의 조형 시스템과 서로 다른 물질, 그리고 질문을 통해 각자의 세계를 수집하게 만드는 캐릭터 IP를 목표로 합니다.", {
    x: MX, y: 6.38, w: CW, h: 0.42, fontFace: FONT, fontSize: 9.2, color: C.creamMute,
    lineSpacing: 12.5, isTextBox: true, margin: 0, valign: "top",
  });
  s.addText("A rabbit. A material. A question.   —   Born from different worlds.", {
    x: MX, y: 6.84, w: CW, h: 0.24, fontFace: FONT, fontSize: 9.5, italic: true,
    color: C.goldOnDark, isTextBox: true, margin: 0,
  });

  footer(s, "12 / 12", true, false);
}

pres.writeFile({ fileName: path.join(__dirname, "..", "WOULD_YOU_TOKKI_IR_DECK_12P.pptx") })
  .then((f) => console.log("DONE:", f));
