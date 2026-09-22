/**
 * wuzu tokki / 우주토끼 — IP 소개서 (확정본)
 *
 * build-deck-warm.js 와 내용은 같고 디자인 언어만 다르다.
 *  - 무채색: 색은 캐릭터 이미지에서만 나온다
 *  - 작은 제목, 넓은 여백
 *  - 비대칭 / 화면 가장자리에 붙는 이미지
 *  - 소문자 라틴 레이블, 얇은 헤어라인
 */
const PptxGenJS = require("pptxgenjs");
const path = require("path");

const A = (f) => path.join(__dirname, "assets", f);

// 원본 픽셀 비율. 이미지는 항상 이 비율로만 배치한다(왜곡 방지).
const IMG = {
  key: { file: A("keyvisual.png"), ar: 1586 / 992 },
  lineup: { file: A("lineup.png"), ar: 2000 / 713 },
  turn: { file: A("turnaround_angora.png"), ar: 1774 / 887 },
};

const C = {
  bg: "FFFFFF",
  bgAlt: "F2F2F1",
  ink: "1A1A1A",
  inkSoft: "5A5A5A",
  muted: "9A9A9A",
  rule: "DCDCDC",
  ruleStrong: "1A1A1A",
};

const F = "Noto Sans KR";

const W = 13.333;
const H = 7.5;
const M = 0.9;          // 좌측 기준선
const CW = W - M * 2;   // 본문 폭

const pptx = new PptxGenJS();
pptx.defineLayout({ name: "W16x9", width: W, height: H });
pptx.layout = "W16x9";
pptx.author = "wuzu tokki";
pptx.title = "wuzu tokki 우주토끼 IP 소개서 (29CM)";

/* ---------- helpers ---------- */

function slide(alt) {
  const s = pptx.addSlide();
  s.background = { color: alt ? C.bgAlt : C.bg };
  return s;
}

function imgH(s, def, x, y, h) {
  s.addImage({ path: def.file, x, y, w: h * def.ar, h });
  return h * def.ar;
}
function imgW(s, def, x, y, w) {
  s.addImage({ path: def.file, x, y, w, h: w / def.ar });
  return w / def.ar;
}

// 소문자 라틴 + 넓은 자간 — 29CM 계열의 섹션 레이블
function label(s, n, txt, x = M, y = 0.54) {
  s.addText([
    { text: n, options: { color: C.ink } },
    { text: "   " + txt, options: { color: C.muted } },
  ], {
    x, y, w: 6, h: 0.24,
    fontFace: F, fontSize: 7.5, charSpacing: 2.6,
  });
}

function title(s, txt, x = M, y = 0.88, w = CW, size = 18) {
  s.addText(txt, {
    x, y, w, h: 0.5,
    fontFace: F, fontSize: size, color: C.ink, lineSpacing: size * 1.4,
  });
}

function lines(arr, opt = {}) {
  return arr.map((t, i) => ({
    text: t,
    options: { breakLine: i < arr.length - 1, ...opt },
  }));
}

function body(s, arr, x, y, w, size = 10.5, color = C.inkSoft) {
  s.addText(lines(arr), {
    x, y, w, h: arr.length * (size * 2.0) / 72,
    fontFace: F, fontSize: size, color, lineSpacing: size * 2.0,
  });
}

function rule(s, x, y, w, color = C.rule, h = 0.006) {
  s.addShape(pptx.ShapeType.rect, { x, y, w, h, fill: { color }, line: { width: 0 } });
}

function footer(s, n) {
  s.addText("wuzu tokki", {
    x: M, y: 7.06, w: 3, h: 0.24,
    fontFace: F, fontSize: 7.5, color: C.muted, charSpacing: 1.6,
  });
  s.addText(String(n).padStart(2, "0"), {
    x: W - M - 1, y: 7.06, w: 1, h: 0.24,
    fontFace: F, fontSize: 7.5, color: C.muted, align: "right",
  });
}

/* ---------- 캐릭터 데이터 (기존 자료 근거) ---------- */

const CHARS = [
  {
    ko: "앙고라토끼", en: "angora",
    mat: "길게 자란 털이 몸의 윤곽을 부드럽게 덮습니다.",
    q: "나는 서로 다른 사람들 사이를 잇는 자리에 서 있을까요?",
    row: { 물질: "긴 털", 표면: "결을 따라 흐르는 장모", 형태: "털이 감싼 둥근 몸",
           색감: "아이보리 · 연한 분홍", 인상: "폭신하고 따뜻함", 질문: "나는 어디에 서 있을까요" },
  },
  {
    ko: "달토끼", en: "moon",
    mat: "회백색 표면 위에 크고 작은 분화구가 남아 있습니다.",
    q: "나는 지나온 시간의 흔적을 오래 간직하는 편일까요?",
    row: { 물질: "달 표면", 표면: "크고 작은 분화구", 형태: "구체에 가까운 몸",
           색감: "회백색", 인상: "서늘하고 고요함", 질문: "흔적을 오래 간직할까요" },
  },
  {
    ko: "구름토끼", en: "cloud",
    mat: "희고 푸른 기가 도는 몸이 덩어리처럼 뭉쳐 있습니다.",
    q: "나는 흩어졌다가도 다시 모이는 편일까요?",
    row: { 물질: "구름", 표면: "매끄러운 무광", 형태: "뭉친 덩어리 실루엣",
           색감: "백색 · 청백색", 인상: "부드럽고 가벼움", 질문: "흩어졌다 다시 모일까요" },
  },
  {
    ko: "모래토끼", en: "sand",
    mat: "가는 알갱이가 층층이 쌓여 결을 만듭니다.",
    q: "나는 눌린 자리에서 새로운 모양이 되는 편일까요?",
    row: { 물질: "모래", 표면: "층층이 쌓인 결", 형태: "결이 감도는 둥근 몸",
           색감: "웜베이지", 인상: "서걱이고 건조함", 질문: "눌린 자리에서 달라질까요" },
  },
  {
    ko: "암석토끼", en: "rock",
    mat: "단단하게 굳은 표면 안쪽으로 가는 결이 지나갑니다.",
    q: "나는 단단해 보이는 안쪽에 어떤 결을 품고 있을까요?",
    row: { 물질: "암석", 표면: "각지고 단단한 면", 형태: "면으로 깎인 둥근 몸",
           색감: "짙은 회색 · 금빛 결", 인상: "묵직하고 단단함", 질문: "안쪽에 어떤 결이 있을까요" },
  },
];

/* ================= 01 표지 — 이미지가 우하단으로 빠져나간다 ================= */
{
  const s = slide();
  const h = 5.9;
  imgH(s, IMG.key, W - h * IMG.key.ar, H - h, h); // 우측·하단 블리드

  s.addText("wuzu tokki", {
    x: M, y: 2.55, w: 2.9, h: 0.46,
    fontFace: F, fontSize: 24, color: C.ink, charSpacing: 0.4,
  });
  s.addText("우주토끼", {
    x: M, y: 3.06, w: 2.9, h: 0.28,
    fontFace: F, fontSize: 10.5, color: C.inkSoft, charSpacing: 2.8,
  });
  rule(s, M, 3.62, 0.5, C.ruleStrong);
  s.addText("온 우주의 물질로\n이루어진 토끼들", {
    x: M, y: 3.84, w: 2.9, h: 0.6,
    fontFace: F, fontSize: 10, color: C.ink, lineSpacing: 19,
  });
  s.addText("character ip", {
    x: M, y: 6.9, w: 2.9, h: 0.24,
    fontFace: F, fontSize: 7.5, color: C.muted, charSpacing: 2.6,
  });
}

/* ================= 02 우주토끼의 시작 — 좌측 블리드 ================= */
{
  const s = slide();
  label(s, "01", "the beginning");
  title(s, "우주의 모든 물질이 토끼가 된다면?");
  body(s, [
    "우주토끼는 달, 구름, 모래, 암석처럼",
    "서로 다른 물질로 이루어진 토끼들의 세계입니다.",
    "같은 토끼의 형태를 가지고 있지만",
    "각자가 가진 촉감과 무게, 흔적과 변화는 모두 다릅니다.",
  ], M, 1.72, 8.6);

  imgW(s, IMG.lineup, 0, 3.5, 9.8); // h = 3.49 → bottom 6.99
  footer(s, 2);
}

/* ================= 03 물질과 사람 — 타이포만, 여백 최대 ================= */
{
  const s = slide(true);
  label(s, "02", "material & person");
  title(s, "물질의 성질은 사람의 모습과 닮아 있습니다");

  body(s, [
    "어떤 물질은 쉽게 모양이 달라지고,",
    "어떤 물질은 오랜 시간 흔적을 품습니다.",
    "어떤 것은 흩어졌다가 다시 모이고,",
    "어떤 것은 단단하지만 작은 균열을 가지고 있습니다.",
  ], M, 2.35, 6.4, 11.5, C.ink);

  body(s, [
    "우주토끼는 물질의 성질을 통해",
    "우리가 가진 서로 다른 모습을 바라봅니다.",
  ], M, 3.95, 6.4, 10, C.inkSoft);

  rule(s, M, 5.5, CW, C.rule);
  s.addText("당신은 어떤 토끼인가요?", {
    x: M, y: 5.78, w: CW, h: 0.42,
    fontFace: F, fontSize: 16, color: C.ink,
  });
  footer(s, 3);
}

/* ================= 04 형태 — 이미지 우측 블리드 ================= */
{
  const s = slide();
  label(s, "03", "one form");
  title(s, "다르지만, 같은 세계에서 태어난 형태", M, 0.88, 5.2);

  const h = 4.3;
  imgH(s, IMG.turn, W - h * IMG.turn.ar, 2.05, h); // w = 8.60 → x = 4.73

  const items = [
    "둥글지만 완전한 구형은 아닌 몸",
    "안정적으로 바닥에 닿는 배와 발",
    "짧고 단순한 앞발과 뒷발",
    "반드시 존재하는 귀와 꼬리",
    "캐릭터마다 통일된 얼굴 비례",
    "물성에 따라 달라지는 표면과 실루엣",
    "장식보다 재료 자체의 특징을 우선",
  ];
  let y = 2.12;
  items.forEach((t) => {
    rule(s, M, y, 3.4, C.rule);
    s.addText(t, {
      x: M, y: y + 0.1, w: 3.4, h: 0.34,
      fontFace: F, fontSize: 9, color: C.ink,
    });
    y += 0.58;
  });
  footer(s, 4);
}

/* ================= 05 첫 번째 물질들 — 이미지 좌측 블리드 ================= */
{
  const s = slide();
  label(s, "04", "first materials");
  title(s, "우주토끼를 이루는 첫 번째 물질들", M, 0.88, 6.2);

  const h = 4.95;
  imgH(s, IMG.key, 0, 1.75, h); // w = 7.91 → bottom 6.70

  let y = 1.72;
  CHARS.forEach((c) => {
    s.addText([
      { text: c.ko, options: { fontSize: 10.5, color: C.ink } },
      { text: "   " + c.en, options: { fontSize: 7, color: C.muted, charSpacing: 1.8 } },
    ], { x: 8.42, y, w: 4.0, h: 0.24, fontFace: F });
    s.addText(c.mat, {
      x: 8.42, y: y + 0.26, w: 4.0, h: 0.38,
      fontFace: F, fontSize: 8.5, color: C.inkSoft, lineSpacing: 12.5,
    });
    s.addText(c.q, {
      x: 8.42, y: y + 0.62, w: 4.0, h: 0.38,
      fontFace: F, fontSize: 8.5, color: C.ink, lineSpacing: 12.5,
    });
    y += 1.02;
  });
  footer(s, 5);
}

/* ================= 06 비교 — 헤어라인 표 ================= */
{
  const s = slide(true);
  label(s, "05", "comparison");
  title(s, "같은 형태, 서로 다른 감각");

  const labels = ["물질", "표면", "형태", "색감", "인상", "질문"];
  const labW = 1.1, colW = 1.92, step = 2.07;
  const cx = (i) => M + labW + 0.15 + i * step; // 마지막 열 끝 = 12.43

  CHARS.forEach((c, i) => {
    s.addText(c.ko, {
      x: cx(i), y: 1.72, w: colW, h: 0.24,
      fontFace: F, fontSize: 10, color: C.ink,
    });
    s.addText(c.en, {
      x: cx(i), y: 1.96, w: colW, h: 0.2,
      fontFace: F, fontSize: 6.5, color: C.muted, charSpacing: 1.6,
    });
  });
  rule(s, M, 2.28, CW, C.ruleStrong, 0.007);

  let y = 2.44;
  labels.forEach((lab) => {
    s.addText(lab, {
      x: M, y: y + 0.03, w: labW, h: 0.22,
      fontFace: F, fontSize: 8.5, color: C.muted,
    });
    CHARS.forEach((c, i) => {
      s.addText(c.row[lab], {
        x: cx(i), y, w: colW, h: 0.46,
        fontFace: F, fontSize: 8.5,
        color: lab === "질문" ? C.inkSoft : C.ink,
        lineSpacing: 12.5,
      });
    });
    y += 0.67;
    rule(s, M, y - 0.11, CW, C.rule);
  });

  s.addText("같은 형태를 공유하지만, 차이는 색이 아니라 표면과 실루엣에서 먼저 드러납니다.", {
    x: M, y: 6.56, w: CW, h: 0.28,
    fontFace: F, fontSize: 9, color: C.inkSoft,
  });
  footer(s, 6);
}

/* ================= 07 시각 언어 — 전폭 블리드 ================= */
{
  const s = slide();
  label(s, "06", "visual language");
  title(s, "부드럽고 조용하지만, 모두 다른 존재들");

  const traits = [
    "아이보리 기가 남아 있는 부드러운 색감",
    "파스텔과 저채도 중심의 색 구성",
    "광택보다 재료의 표면감을 우선",
    "단순하고 안정적인 실루엣",
    "최소한으로 절제한 하이라이트",
    "귀여움과 오브제성의 균형",
    "표면이 흔들리지 않는 정돈된 마감",
    "사실적이지 않으면서 물성은 남는 표현",
  ];
  const cw = 2.78, cs = 2.95;
  traits.forEach((t, i) => {
    const x = M + (i % 4) * cs;
    const y = 1.62 + Math.floor(i / 4) * 0.44;
    rule(s, x, y, cw, C.rule);
    s.addText(t, {
      x, y: y + 0.08, w: cw, h: 0.28,
      fontFace: F, fontSize: 8.5, color: C.ink,
    });
  });

  // 좌·우·하단 전폭 블리드
  imgW(s, IMG.lineup, 0, H - W / IMG.lineup.ar, W); // h = 4.75 → y = 2.75
}

/* ================= 08 확장되는 물질 ================= */
{
  const s = slide(true);
  label(s, "07", "expanding materials");
  title(s, "우주토끼는 계속 다른 물질을 발견합니다");
  body(s, [
    "캐릭터의 수를 늘리는 방식이 아니라, 새로운 물질을 만날 때마다",
    "세계가 한 겹씩 넓어지는 구조입니다.",
  ], M, 1.7, 9.4, 10);

  const mats = ["자개", "푸딩", "모찌", "오로라", "바다", "아이스크림", "커피", "비눗방울", "태양", "크림"];
  const colW = 2.2, step = 2.32;
  mats.forEach((m, i) => {
    const x = M + (i % 5) * step;
    const y = 2.95 + Math.floor(i / 5) * 1.34;
    rule(s, x, y, colW, C.ruleStrong);
    s.addText(m, {
      x, y: y + 0.2, w: colW, h: 0.36,
      fontFace: F, fontSize: 12, color: C.ink,
    });
  });

  rule(s, M, 6.16, CW, C.rule);
  s.addText("확정된 제품 계획이 아니라, 앞으로 탐구할 수 있는 물질의 방향입니다.", {
    x: M, y: 6.38, w: CW, h: 0.28,
    fontFace: F, fontSize: 9, color: C.inkSoft,
  });
  footer(s, 8);
}

/* ================= 09 경험으로의 확장 ================= */
{
  const s = slide();
  label(s, "08", "experience");
  title(s, "보고, 만지고, 모으고, 공간에서 만나는 토끼");

  const groups = [
    { h: "오브제로", items: ["수집형 아트토이", "디자인 오브제", "미니 피규어", "대표 캐릭터 봉제 인형"] },
    { h: "곁에 두는 것으로", items: ["키링과 백참", "패키지와 질문 카드"] },
    { h: "공간에서", items: ["팝업 전시", "공간 연출"] },
    { h: "이미지로", items: ["브랜드 협업", "이미지와 짧은 콘텐츠"] },
  ];
  groups.forEach((g, i) => {
    const x = M + i * 2.95;
    rule(s, x, 1.72, 2.78, C.ruleStrong);
    s.addText(g.h, {
      x, y: 1.9, w: 2.78, h: 0.26,
      fontFace: F, fontSize: 10, color: C.ink,
    });
    s.addText(lines(g.items), {
      x, y: 2.24, w: 2.78, h: 1.1,
      fontFace: F, fontSize: 8.5, color: C.inkSoft, lineSpacing: 16,
    });
  });

  // 이미지 미확보 구간 — 플레이스홀더
  ["아트토이 · 오브제", "패키지 · 질문 카드", "팝업 · 공간"].forEach((t, i) => {
    const x = M + i * 3.93;
    s.addShape(pptx.ShapeType.rect, {
      x, y: 3.85, w: 3.65, h: 2.6,
      fill: { color: C.bgAlt }, line: { color: C.rule, width: 0.75 },
    });
    s.addText(t, {
      x, y: 4.94, w: 3.65, h: 0.26,
      fontFace: F, fontSize: 9, color: C.inkSoft, align: "center",
    });
    s.addText("이미지 준비 중", {
      x, y: 5.2, w: 3.65, h: 0.24,
      fontFace: F, fontSize: 7.5, color: C.muted, align: "center", charSpacing: 1.2,
    });
  });
  footer(s, 9);
}

/* ================= 10 마무리 — 좌측 블리드 ================= */
{
  const s = slide();
  const h = 5.6;
  imgH(s, IMG.key, 0, 1.0, h); // w = 8.95 → bottom 6.60

  s.addText("Would you, Tokki?", {
    x: 9.5, y: 1.55, w: 2.95, h: 0.42,
    fontFace: F, fontSize: 19, color: C.ink,
  });
  rule(s, 9.5, 2.18, 0.5, C.ruleStrong);

  s.addText(lines([
    "서로 다른 물질로 이루어진",
    "우주토끼들은 우리에게",
    "하나의 질문을 건넵니다.",
  ]), {
    x: 9.5, y: 2.5, w: 2.95, h: 0.9,
    fontFace: F, fontSize: 9.5, color: C.inkSoft, lineSpacing: 17,
  });

  s.addText(lines([
    "나는 어떤 성질을 가지고 있는지,",
    "어떤 흔적을 품고 있는지,",
    "그리고 어떤 모습으로",
    "변해가고 있는지.",
  ]), {
    x: 9.5, y: 3.68, w: 2.95, h: 1.2,
    fontFace: F, fontSize: 9.5, color: C.ink, lineSpacing: 17,
  });

  rule(s, 9.5, 5.2, 2.95, C.rule);
  s.addText("당신은 어떤 토끼인가요?", {
    x: 9.5, y: 5.46, w: 2.95, h: 0.38,
    fontFace: F, fontSize: 13.5, color: C.ink,
  });
  s.addText("wuzu tokki · 우주토끼", {
    x: 9.5, y: 6.32, w: 2.95, h: 0.24,
    fontFace: F, fontSize: 7.5, color: C.muted, charSpacing: 1.6,
  });
}

const OUT = path.join(__dirname, "Wuzu_Tokki_IP_소개서_20260922.pptx");
pptx.writeFile({ fileName: OUT }).then(() => console.log("OK ->", OUT));
