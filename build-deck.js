/**
 * wuzu tokki / 우주토끼 — IP 소개서
 * 편집용 소스. `node build-deck.js` 로 pptx 재생성.
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
  bg: "F7F4EE",
  bgAlt: "FBF9F5",
  ink: "33312E",
  inkSoft: "6B665F",
  muted: "938C83",
  bronze: "A8875C",
  sky: "CBD9E0",
  warm: "C9C2B7",
  rule: "DED7CB",
};

const F = "Noto Sans KR";
const FL = "Pretendard"; // 설치 시 우선 적용되도록 라틴 제목에 병기

const W = 13.333;
const H = 7.5;

const pptx = new PptxGenJS();
pptx.defineLayout({ name: "W16x9", width: W, height: H });
pptx.layout = "W16x9";
pptx.author = "wuzu tokki";
pptx.title = "wuzu tokki 우주토끼 IP 소개서";

/* ---------- helpers ---------- */

function slide(alt) {
  const s = pptx.addSlide();
  s.background = { color: alt ? C.bgAlt : C.bg };
  return s;
}

// 높이 기준으로 폭을 계산해 원본 비율 유지
function imgH(s, def, x, y, h) {
  s.addImage({ path: def.file, x, y, w: h * def.ar, h });
  return h * def.ar;
}
// 폭 기준
function imgW(s, def, x, y, w) {
  s.addImage({ path: def.file, x, y, w, h: w / def.ar });
  return w / def.ar;
}

function eyebrow(s, txt, x = 0.85, y = 0.52) {
  s.addText(txt, {
    x, y, w: 6, h: 0.26,
    fontFace: F, fontSize: 9, color: C.bronze, charSpacing: 2.2,
  });
}

function title(s, txt, x = 0.85, y = 0.88, w = 11.6, size = 25) {
  s.addText(txt, {
    x, y, w, h: 0.62,
    fontFace: F, fontSize: size, color: C.ink, lineSpacing: size * 1.35,
  });
}

function lines(arr, opt = {}) {
  return arr.map((t, i) => ({
    text: t,
    options: { breakLine: i < arr.length - 1, ...opt },
  }));
}

function body(s, arr, x, y, w, size = 12, color = C.inkSoft) {
  s.addText(lines(arr), {
    x, y, w, h: arr.length * (size * 1.85) / 72,
    fontFace: F, fontSize: size, color, lineSpacing: size * 1.85,
  });
}

function rule(s, x, y, w, color = C.rule, h = 0.012) {
  s.addShape(pptx.ShapeType.rect, { x, y, w, h, fill: { color }, line: { width: 0 } });
}

// 본문은 y=6.95 이전에 끝내고, 푸터는 그 아래에만 둔다.
function footer(s, n) {
  s.addText("wuzu tokki", {
    x: 0.85, y: 7.05, w: 3, h: 0.26,
    fontFace: F, fontSize: 8, color: C.muted, charSpacing: 1.4,
  });
  s.addText(String(n).padStart(2, "0"), {
    x: 11.5, y: 7.05, w: 0.98, h: 0.26,
    fontFace: F, fontSize: 8, color: C.muted, align: "right",
  });
}

/* ---------- 캐릭터 데이터 (기존 자료 근거) ---------- */

const CHARS = [
  {
    ko: "앙고라토끼",
    en: "ANGORA",
    mat: "길게 자란 털이 몸의 윤곽을 부드럽게 덮습니다.",
    q: "나는 서로 다른 사람들 사이를 잇는 자리에 서 있을까요?",
    row: {
      물질: "긴 털",
      표면: "결을 따라 흐르는 장모",
      형태: "털이 감싼 둥근 몸",
      색감: "아이보리 · 연한 분홍",
      인상: "폭신하고 따뜻함",
      질문: "나는 어디에 서 있을까요",
    },
  },
  {
    ko: "달토끼",
    en: "MOON",
    mat: "회백색 표면 위에 크고 작은 분화구가 남아 있습니다.",
    q: "나는 지나온 시간의 흔적을 오래 간직하는 편일까요?",
    row: {
      물질: "달 표면",
      표면: "크고 작은 분화구",
      형태: "구체에 가까운 몸",
      색감: "회백색",
      인상: "서늘하고 고요함",
      질문: "흔적을 오래 간직할까요",
    },
  },
  {
    ko: "구름토끼",
    en: "CLOUD",
    mat: "희고 푸른 기가 도는 몸이 덩어리처럼 뭉쳐 있습니다.",
    q: "나는 흩어졌다가도 다시 모이는 편일까요?",
    row: {
      물질: "구름",
      표면: "매끄러운 무광",
      형태: "뭉친 덩어리 실루엣",
      색감: "백색 · 청백색",
      인상: "부드럽고 가벼움",
      질문: "흩어졌다 다시 모일까요",
    },
  },
  {
    ko: "모래토끼",
    en: "SAND",
    mat: "가는 알갱이가 층층이 쌓여 결을 만듭니다.",
    q: "나는 눌린 자리에서 새로운 모양이 되는 편일까요?",
    row: {
      물질: "모래",
      표면: "층층이 쌓인 결",
      형태: "결이 감도는 둥근 몸",
      색감: "웜베이지",
      인상: "서걱이고 건조함",
      질문: "눌린 자리에서 달라질까요",
    },
  },
  {
    ko: "암석토끼",
    en: "ROCK",
    mat: "단단하게 굳은 표면 안쪽으로 가는 결이 지나갑니다.",
    q: "나는 단단해 보이는 안쪽에 어떤 결을 품고 있을까요?",
    row: {
      물질: "암석",
      표면: "각지고 단단한 면",
      형태: "면으로 깎인 둥근 몸",
      색감: "짙은 회색 · 금빛 결",
      인상: "묵직하고 단단함",
      질문: "안쪽에 어떤 결이 있을까요",
    },
  },
];

/* ================= 01 표지 ================= */
{
  const s = slide();
  s.addText("wuzu tokki", {
    x: 0, y: 0.45, w: W, h: 0.66,
    fontFace: FL, fontSize: 36, color: C.ink, align: "center", charSpacing: 1.2,
  });
  s.addText("우주토끼", {
    x: 0, y: 1.1, w: W, h: 0.34,
    fontFace: F, fontSize: 15, color: C.inkSoft, align: "center", charSpacing: 2.4,
  });
  rule(s, (W - 0.78) / 2, 1.56, 0.78, C.bronze);
  s.addText("온 우주의 물질로 이루어진 토끼들", {
    x: 0, y: 1.7, w: W, h: 0.32,
    fontFace: F, fontSize: 12.5, color: C.inkSoft, align: "center",
  });

  const h = 5.05, w = h * IMG.key.ar;
  imgH(s, IMG.key, (W - w) / 2, 2.15, h);
}

/* ================= 02 우주토끼의 시작 ================= */
{
  const s = slide();
  eyebrow(s, "THE BEGINNING");
  title(s, "우주의 모든 물질이 토끼가 된다면?");
  body(s, [
    "우주토끼는 달, 구름, 모래, 암석처럼",
    "서로 다른 물질로 이루어진 토끼들의 세계입니다.",
    "같은 토끼의 형태를 가지고 있지만",
    "각자가 가진 촉감과 무게, 흔적과 변화는 모두 다릅니다.",
  ], 0.85, 1.70, 9.4, 12.5);

  // P7이 같은 라인업을 전폭으로 쓰므로, 여기서는 작게 놓아 도입부로 읽히게 한다.
  const w = 9.6; // h = 3.42 → bottom 6.62
  imgW(s, IMG.lineup, (W - w) / 2, 3.2, w);
  footer(s, 2);
}

/* ================= 03 물질과 사람을 연결하는 질문 ================= */
{
  const s = slide(true);
  eyebrow(s, "MATERIAL & PERSON");
  title(s, "물질의 성질은 사람의 모습과 닮아 있습니다");
  rule(s, 0.85, 1.92, 0.62, C.bronze);

  body(s, [
    "어떤 물질은 쉽게 모양이 달라지고,",
    "어떤 물질은 오랜 시간 흔적을 품습니다.",
    "어떤 것은 흩어졌다가 다시 모이고,",
    "어떤 것은 단단하지만 작은 균열을 가지고 있습니다.",
  ], 0.85, 2.42, 10.4, 13.5, C.ink);

  body(s, [
    "우주토끼는 물질의 성질을 통해",
    "우리가 가진 서로 다른 모습을 바라봅니다.",
  ], 0.85, 4.08, 10.4, 12, C.inkSoft);

  rule(s, 0.85, 5.32, 11.63, C.rule);
  s.addText("당신은 어떤 토끼인가요?", {
    x: 0.85, y: 5.6, w: 11.63, h: 0.5,
    fontFace: F, fontSize: 21, color: C.bronze, charSpacing: 0.6,
  });
  footer(s, 3);
}

/* ================= 04 우주토끼의 형태 ================= */
{
  const s = slide();
  eyebrow(s, "ONE FORM");
  title(s, "다르지만, 같은 세계에서 태어난 형태");
  s.addText("물성이 달라도 하나의 패밀리로 읽히는 이유는 형태의 기준을 공유하기 때문입니다.", {
    x: 0.85, y: 1.62, w: 9.2, h: 0.3,
    fontFace: F, fontSize: 11, color: C.inkSoft,
  });

  const h = 4.55;
  imgH(s, IMG.turn, 0.6, 2.2, h); // w = 9.10

  const items = [
    "둥글지만 완전한 구형은 아닌 몸",
    "안정적으로 바닥에 닿는 배와 발",
    "짧고 단순한 앞발과 뒷발",
    "반드시 존재하는 귀와 꼬리",
    "캐릭터마다 통일된 얼굴 비례",
    "물성에 따라 달라지는 표면과 실루엣",
    "장식보다 재료 자체의 특징을 우선",
  ];
  let y = 2.28;
  items.forEach((t) => {
    rule(s, 10.05, y, 2.55, C.rule, 0.008);
    s.addText(t, {
      x: 10.05, y: y + 0.1, w: 2.55, h: 0.44,
      fontFace: F, fontSize: 9.5, color: C.ink, lineSpacing: 15,
    });
    y += 0.63;
  });
  footer(s, 4);
}

/* ================= 05 대표 캐릭터와 주요 캐릭터 ================= */
{
  const s = slide();
  eyebrow(s, "FIRST MATERIALS", 0.5);
  title(s, "우주토끼를 이루는 첫 번째 물질들", 0.5, 0.86, 8.6, 24);

  const h = 5.15;
  imgH(s, IMG.key, 0.5, 1.78, h); // w = 8.23 → bottom 6.93

  let y = 1.68;
  CHARS.forEach((c) => {
    s.addText([
      { text: c.ko, options: { fontSize: 11.5, color: C.ink } },
      { text: "  " + c.en, options: { fontSize: 7.5, color: C.bronze, charSpacing: 1.4 } },
    ], { x: 9.22, y, w: 3.6, h: 0.26, fontFace: F });
    s.addText(c.mat, {
      x: 9.22, y: y + 0.27, w: 3.6, h: 0.42,
      fontFace: F, fontSize: 8.8, color: C.inkSoft, lineSpacing: 13,
    });
    s.addText(c.q, {
      x: 9.22, y: y + 0.66, w: 3.6, h: 0.42,
      fontFace: F, fontSize: 8.8, color: C.bronze, lineSpacing: 13,
    });
    y += 1.04;
  });
  footer(s, 5);
}

/* ================= 06 물성별 캐릭터 비교 ================= */
{
  const s = slide(true);
  eyebrow(s, "COMPARISON");
  title(s, "같은 형태, 서로 다른 감각");

  // 우측 여백(x=12.48)을 넘지 않도록 5열을 계산
  const labels = ["물질", "표면", "형태", "색감", "인상", "질문"];
  const x0 = 0.85, labW = 1.15, colW = 1.92, step = 2.06, tblW = 11.63;
  const cx = (i) => x0 + labW + 0.18 + i * step; // 마지막 열 끝 = 12.34

  // 헤더
  CHARS.forEach((c, i) => {
    s.addText(c.ko, {
      x: cx(i), y: 1.78, w: colW, h: 0.26,
      fontFace: F, fontSize: 11.5, color: C.ink,
    });
    s.addText(c.en, {
      x: cx(i), y: 2.04, w: colW, h: 0.2,
      fontFace: F, fontSize: 7, color: C.bronze, charSpacing: 1.3,
    });
  });
  rule(s, x0, 2.36, tblW, C.ink, 0.009);

  let y = 2.52;
  labels.forEach((lab) => {
    s.addText(lab, {
      x: x0, y: y + 0.04, w: labW, h: 0.24,
      fontFace: F, fontSize: 9, color: C.muted, charSpacing: 0.8,
    });
    CHARS.forEach((c, i) => {
      s.addText(c.row[lab], {
        x: cx(i), y, w: colW, h: 0.5,
        fontFace: F, fontSize: 8.8,
        color: lab === "질문" ? C.bronze : C.ink,
        lineSpacing: 13,
      });
    });
    y += 0.68;
    rule(s, x0, y - 0.1, tblW, C.rule, 0.006);
  });

  s.addText("같은 형태를 공유하지만, 차이는 색이 아니라 표면과 실루엣에서 먼저 드러납니다.", {
    x: x0, y: 6.62, w: tblW, h: 0.3,
    fontFace: F, fontSize: 9.5, color: C.inkSoft,
  });
  footer(s, 6);
}

/* ================= 07 시각 언어 ================= */
{
  const s = slide();
  eyebrow(s, "VISUAL LANGUAGE");
  title(s, "부드럽고 조용하지만, 모두 다른 존재들");

  // 서로 다른 물성이 하나의 색조 안에서 읽히는 장면 — 라인업을 전폭으로
  const iw = 11.63;
  imgW(s, IMG.lineup, 0.85, 1.7, iw); // h = 4.15 → bottom 5.85

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
    const x = 0.85 + (i % 4) * cs;
    const y = 6.02 + Math.floor(i / 4) * 0.46;
    rule(s, x, y, cw, C.rule, 0.007);
    s.addText(t, {
      x, y: y + 0.09, w: cw, h: 0.3,
      fontFace: F, fontSize: 9, color: C.ink,
    });
  });
  footer(s, 7);
}

/* ================= 08 확장되는 물질 ================= */
{
  const s = slide(true);
  eyebrow(s, "EXPANDING MATERIALS");
  title(s, "우주토끼는 계속 다른 물질을 발견합니다");
  body(s, [
    "캐릭터의 수를 늘리는 방식이 아니라, 새로운 물질을 만날 때마다",
    "세계가 한 겹씩 넓어지는 구조입니다.",
  ], 0.85, 1.72, 10.4, 12);

  const mats = ["자개", "푸딩", "모찌", "오로라", "바다", "아이스크림", "커피", "비눗방울", "태양", "크림"];
  const x0 = 0.85, colW = 2.24, step = 2.42;
  mats.forEach((m, i) => {
    const col = i % 5, row = Math.floor(i / 5);
    const x = x0 + col * step, y = 2.92 + row * 1.42;
    rule(s, x, y, colW, C.bronze, 0.009);
    s.addText(m, {
      x, y: y + 0.22, w: colW, h: 0.4,
      fontFace: F, fontSize: 14, color: C.ink,
    });
  });

  rule(s, x0, 6.1, 12.0, C.rule);
  s.addText("확정된 제품 계획이 아니라, 앞으로 탐구할 수 있는 물질의 방향입니다.", {
    x: x0, y: 6.32, w: 12.0, h: 0.3,
    fontFace: F, fontSize: 9.5, color: C.inkSoft,
  });
  footer(s, 8);
}

/* ================= 09 경험으로의 확장 ================= */
{
  const s = slide();
  eyebrow(s, "EXPERIENCE");
  title(s, "보고, 만지고, 모으고, 공간에서 만나는 토끼");

  const groups = [
    { h: "오브제로", items: ["수집형 아트토이", "디자인 오브제", "미니 피규어", "대표 캐릭터 봉제 인형"] },
    { h: "곁에 두는 것으로", items: ["키링과 백참", "패키지와 질문 카드"] },
    { h: "공간에서", items: ["팝업 전시", "공간 연출"] },
    { h: "이미지로", items: ["브랜드 협업", "이미지와 짧은 콘텐츠"] },
  ];
  const x0 = 0.85, step = 3.05;
  groups.forEach((g, i) => {
    const x = x0 + i * step;
    rule(s, x, 1.78, 2.8, C.bronze, 0.009);
    s.addText(g.h, {
      x, y: 1.98, w: 2.8, h: 0.28,
      fontFace: F, fontSize: 11, color: C.ink,
    });
    s.addText(lines(g.items), {
      x, y: 2.34, w: 2.8, h: 1.2,
      fontFace: F, fontSize: 9.3, color: C.inkSoft, lineSpacing: 17,
    });
  });

  // 이미지 미확보 구간 — 플레이스홀더
  const ph = ["아트토이 · 오브제", "패키지 · 질문 카드", "팝업 · 공간"];
  ph.forEach((t, i) => {
    const x = 0.85 + i * 4.1;
    s.addShape(pptx.ShapeType.rect, {
      x, y: 3.95, w: 3.8, h: 2.55,
      fill: { color: C.bgAlt }, line: { color: C.rule, width: 0.75 },
    });
    s.addText(t, {
      x, y: 5.0, w: 3.8, h: 0.28,
      fontFace: F, fontSize: 9.5, color: C.muted, align: "center",
    });
    s.addText("이미지 준비 중", {
      x, y: 5.26, w: 3.8, h: 0.26,
      fontFace: F, fontSize: 8, color: C.warm, align: "center", charSpacing: 1,
    });
  });
  footer(s, 9);
}

/* ================= 10 마무리 ================= */
{
  const s = slide();
  const h = 5.3;
  imgH(s, IMG.key, 0.35, 1.3, h); // w = 8.47

  s.addText("Would you, Tokki?", {
    x: 9.15, y: 1.62, w: 3.7, h: 0.5,
    fontFace: FL, fontSize: 23, color: C.ink,
  });
  rule(s, 9.15, 2.3, 0.62, C.bronze);

  s.addText(lines([
    "서로 다른 물질로 이루어진 우주토끼들은",
    "우리에게 하나의 질문을 건넵니다.",
  ]), {
    x: 9.15, y: 2.66, w: 3.7, h: 0.72,
    fontFace: F, fontSize: 10.5, color: C.inkSoft, lineSpacing: 19,
  });

  s.addText(lines([
    "나는 어떤 성질을 가지고 있는지,",
    "어떤 흔적을 품고 있는지,",
    "그리고 어떤 모습으로 변해가고 있는지.",
  ]), {
    x: 9.15, y: 3.62, w: 3.7, h: 1.1,
    fontFace: F, fontSize: 10.5, color: C.ink, lineSpacing: 19,
  });

  rule(s, 9.15, 5.12, 3.7, C.rule);
  s.addText("당신은 어떤 토끼인가요?", {
    x: 9.15, y: 5.4, w: 3.7, h: 0.44,
    fontFace: F, fontSize: 16, color: C.bronze,
  });
  s.addText("wuzu tokki · 우주토끼", {
    x: 9.15, y: 6.28, w: 3.7, h: 0.26,
    fontFace: F, fontSize: 8.5, color: C.muted, charSpacing: 1.4,
  });
}

const OUT = path.join(__dirname, "Wuzu_Tokki_IP_소개서_20260922.pptx");
pptx.writeFile({ fileName: OUT }).then(() => console.log("OK ->", OUT));
