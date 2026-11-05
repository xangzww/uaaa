const pptxgen = require("pptxgenjs");
const sharp = require("sharp");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const gi = require("react-icons/gi");
const { applyTheme } = require("./apply_theme.js");

const OUT = process.argv[2] || "Yaroslav_Mudry.pptx";

const THEME = {
  name: "Byzantine Porphyry",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1F1A1C", lt1: "FFFFFF", dk2: "3A1626", lt2: "F3EEEA",
    accent1: "B8892E", accent2: "8C2A2A", accent3: "D9C08A", accent4: "6E5A63",
    accent5: "5A2338", accent6: "8A7F84", hlink: "8C2A2A", folHlink: "5A2338",
  },
};
const H = THEME.colors;

// ---------- images ----------
async function iconPng(Comp, color) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: "#" + color, size: 256 }));
  const buf = await sharp(Buffer.from(svg)).resize(256, 256).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

function cathedralSvg(stroke, fill, sw) {
  return `
  <g fill="${fill}" stroke="${stroke}" stroke-width="${sw}" stroke-linejoin="round" stroke-linecap="round">
    <!-- outer galleries -->
    <rect x="60" y="400" width="480" height="160"/>
    <path d="M60 400 A40 40 0 0 1 140 400 A40 40 0 0 1 220 400 M380 400 A40 40 0 0 1 460 400 A40 40 0 0 1 540 400"/>
    <!-- main body -->
    <rect x="130" y="320" width="340" height="240"/>
    <path d="M130 320 A56.7 56.7 0 0 1 243.3 320 A56.7 56.7 0 0 1 356.7 320 A56.7 56.7 0 0 1 470 320"/>
    <!-- side drums -->
    <rect x="160" y="230" width="56" height="70"/>
    <path d="M154 230 Q188 160 222 230 Z"/>
    <rect x="384" y="230" width="56" height="70"/>
    <path d="M378 230 Q412 160 446 230 Z"/>
    <!-- central drum -->
    <rect x="258" y="190" width="84" height="110"/>
    <path d="M248 190 Q300 60 352 190 Z"/>
    <!-- small far domes -->
    <rect x="80" y="352" width="34" height="48"/>
    <path d="M76 352 Q97 312 118 352 Z"/>
    <rect x="486" y="352" width="34" height="48"/>
    <path d="M482 352 Q503 312 524 352 Z"/>
  </g>
  <g fill="none" stroke="${stroke}" stroke-width="${sw}" stroke-linecap="round">
    <path d="M300 98 V40 M284 58 H316"/>
    <path d="M188 172 V140 M178 152 H198"/>
    <path d="M412 172 V140 M402 152 H422"/>
    <!-- door and windows -->
    <path d="M270 560 V480 A30 30 0 0 1 330 480 V560"/>
    <path d="M170 450 V410 A13 13 0 0 1 196 410 V450 M404 450 V410 A13 13 0 0 1 430 410 V450"/>
    <path d="M287 280 V240 A13 13 0 0 1 313 240 V280"/>
    <path d="M178 290 V262 A10 10 0 0 1 198 262 V290 M402 290 V262 A10 10 0 0 1 422 290"/>
    <path d="M30 560 H570"/>
  </g>`;
}

function patternDefs() {
  return `
  <defs>
    <radialGradient id="g" cx="30%" cy="40%" r="85%">
      <stop offset="0%" stop-color="#4E1D33"/>
      <stop offset="100%" stop-color="#220C17"/>
    </radialGradient>
    <pattern id="p" width="72" height="72" patternUnits="userSpaceOnUse">
      <path d="M36 6 L66 36 L36 66 L6 36 Z" fill="none" stroke="#B8892E" stroke-width="1.4"/>
      <path d="M36 26 V46 M26 36 H46" stroke="#B8892E" stroke-width="1.4"/>
      <circle cx="0" cy="0" r="2.5" fill="#B8892E"/><circle cx="72" cy="0" r="2.5" fill="#B8892E"/>
      <circle cx="0" cy="72" r="2.5" fill="#B8892E"/><circle cx="72" cy="72" r="2.5" fill="#B8892E"/>
    </pattern>
  </defs>
  <rect width="1920" height="1080" fill="url(#g)"/>
  <rect width="1920" height="1080" fill="url(#p)" opacity="0.07"/>`;
}

async function bgTitle() {
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="1920" height="1080" viewBox="0 0 1920 1080">
    ${patternDefs()}
    <circle cx="1480" cy="560" r="400" fill="#B8892E" opacity="0.07"/>
    <circle cx="1480" cy="560" r="400" fill="none" stroke="#B8892E" stroke-width="2" opacity="0.45"/>
    <circle cx="1480" cy="560" r="430" fill="none" stroke="#B8892E" stroke-width="1" opacity="0.25"/>
    <g transform="translate(1180 270) scale(1.0)">${cathedralSvg("#D9C08A", "none", 3.2)}</g>
  </svg>`;
  return sharp(Buffer.from(svg)).png().toFile("bg_title.png");
}

async function bgEnd() {
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="1920" height="1080" viewBox="0 0 1920 1080">
    ${patternDefs()}
  </svg>`;
  return sharp(Buffer.from(svg)).png().toFile("bg_end.png");
}

async function cathedralCard() {
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="900" height="760" viewBox="-150 0 900 760">
    <circle cx="300" cy="360" r="320" fill="#E9DFD6"/>
    ${cathedralSvg("#3A1626", "#FFFFFF", 4.5)}
  </svg>`;
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// ---------- deck ----------
(async () => {
  await bgTitle();
  await bgEnd();

  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9"; // 10 x 5.625
  pres.title = "Ярослав Мудрый";
  pres.subject = "История России: Киевская Русь";
  pres.author = "Презентация по истории";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  const C = pres.SchemeColor;

  // Layouts
  pres.defineSlideMaster({
    title: "TITLE_DARK",
    background: { path: "bg_title.png" },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 1.65, w: 5.6, h: 1.35, fontSize: 46, bold: true, color: C.background1, align: "left", valign: "bottom", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "body", type: "body", x: 0.6, y: 3.1, w: 5.3, h: 0.9, fontSize: 17, color: C.accent3, align: "left", valign: "top", margin: 0 }, text: "" } },
    ],
  });
  pres.defineSlideMaster({
    title: "CONTENT",
    background: { color: C.background1 },
    margin: [0.5, 0.5, 0.5, 0.5],
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.3, w: 9, h: 0.75, fontSize: 30, bold: true, color: C.text2, align: "left", valign: "middle", margin: 0 }, text: "" } },
      { text: { text: "Ярослав Мудрый · Киевская Русь", options: { x: 0.5, y: 5.2, w: 5, h: 0.3, fontSize: 10, color: C.accent6, margin: 0 } } },
    ],
    slideNumber: { x: 9.0, y: 5.2, w: 0.5, h: 0.3, fontSize: 10, color: C.accent6, align: "right" },
  });
  pres.defineSlideMaster({
    title: "END_DARK",
    background: { path: "bg_end.png" },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.45, w: 9, h: 0.8, fontSize: 34, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0 }, text: "" } },
    ],
    slideNumber: { x: 9.0, y: 5.2, w: 0.5, h: 0.3, fontSize: 10, color: C.accent3, align: "right" },
  });

  // Icons
  const I = {};
  const gold = H.accent1, porph = H.dk2;
  for (const [k, comp, col] of [
    ["king", gi.GiKing, gold], ["castle", gi.GiCastle, gold], ["ring", gi.GiDiamondRing, gold],
    ["swords", gi.GiCrossedSwords, gold], ["shield", gi.GiShield, gold], ["ship", gi.GiShipBow, gold],
    ["fort", gi.GiCapitol, gold], ["crown", gi.GiCrown, gold], ["tree", gi.GiFamilyTree, gold],
    ["scroll", gi.GiScrollUnfurled, gold], ["church", gi.GiChurch, gold], ["book", gi.GiOpenBook, gold],
    ["books", gi.GiBookCover, gold], ["quill", gi.GiQuillInk, gold],
    ["dShield", gi.GiShield, porph], ["dScales", gi.GiScales, porph], ["dChurch", gi.GiChurch, porph],
  ]) I[k] = await iconPng(comp, col);
  const cathedral = await cathedralCard();

  // helpers
  const iconCircle = (slide, img, x, y, d, fill, name) => {
    slide.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill }, objectName: name + " circle" });
    const p = d * 0.22;
    slide.addImage({ data: img, x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p, objectName: name + " icon", altText: name });
  };
  const shadow = () => ({ type: "outer", color: "3A1626", blur: 8, offset: 2, angle: 90, opacity: 0.15 });
  const card = (slide, x, y, w, h, fill, name) =>
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color: fill }, line: { color: fill }, shadow: shadow(), objectName: name });
  const tb = (slide, text, opts) => slide.addText(text, Object.assign({ isTextBox: true, margin: 0 }, opts));

  // ===== 1. Title =====
  pres.addSection({ title: "Введение" });
  let s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "Введение" });
  tb(s, "ИСТОРИЯ РОССИИ · КИЕВСКАЯ РУСЬ", { x: 0.6, y: 1.2, w: 5.4, h: 0.35, fontSize: 12, bold: true, color: C.accent1, charSpacing: 3, objectName: "Kicker" });
  s.addText("Ярослав Мудрый", { placeholder: "title" });
  s.addText("Великий князь Киевский и расцвет Древнерусского государства", { placeholder: "body" });
  tb(s, "978 — 1054", { x: 0.6, y: 4.2, w: 3, h: 0.5, fontSize: 22, bold: true, color: C.accent1, fontFace: "Cambria", objectName: "Years" });
  s.addNotes("Ярослав Владимирович, прозванный Мудрым, — один из самых известных правителей Киевской Руси. Его правление считается временем наивысшего расцвета Древнерусского государства.");

  // ===== 2. Biography =====
  pres.addSection({ title: "Биография" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Биография" });
  s.addText("Князь и его происхождение", { placeholder: "title" });
  const rows2 = [
    ["king", "Сын Крестителя Руси", "Родился ок. 978 г. в семье князя Владимира Святославича и полоцкой княжны Рогнеды"],
    ["castle", "Путь через северные столы", "С 988 г. княжил в Ростове, с 1010 г. — в Новгороде, главном городе Севера Руси"],
    ["ring", "Союз со Швецией", "Супруга — Ингигерда (Ирина), дочь шведского короля Олафа Шётконунга"],
  ];
  rows2.forEach(([ic, h, d], i) => {
    const y = 1.35 + i * 1.22;
    iconCircle(s, I[ic], 0.5, y, 0.62, C.text2, h);
    tb(s, h, { x: 1.35, y: y - 0.02, w: 4.5, h: 0.36, fontSize: 17, bold: true, color: C.text2, fontFace: "Cambria", objectName: h + " heading" });
    tb(s, d, { x: 1.35, y: y + 0.36, w: 4.5, h: 0.62, fontSize: 14, color: C.text1, valign: "top", objectName: h + " text" });
  });
  const stats2 = [["35 лет", "на киевском престоле (1019–1054)"], ["1036", "год, когда стал единовластным правителем Руси"], ["3", "дочери Ярослава стали королевами Европы"]];
  stats2.forEach(([n, l], i) => {
    const y = 1.3 + i * 1.27;
    card(s, 6.25, y, 3.25, 1.1, C.text2, "Stat card " + (i + 1));
    tb(s, n, { x: 6.4, y: y + 0.12, w: 1.6, h: 0.86, fontSize: 30, bold: true, color: C.accent1, fontFace: "Cambria", valign: "middle", objectName: "Stat " + n });
    tb(s, l, { x: 8.05, y: y + 0.12, w: 1.35, h: 0.86, fontSize: 12, color: C.background1, valign: "middle", objectName: "Stat label " + (i + 1) });
  });
  s.addNotes("Ярослав был одним из многочисленных сыновей Владимира Святого. Ещё при жизни отца он получил в управление Ростов, а затем Новгород — второй по значению город Руси. Брак с шведской принцессой Ингигердой укрепил связи Руси со Скандинавией.");

  // ===== 3. Timeline =====
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Биография" });
  s.addText("Борьба за киевский престол", { placeholder: "title" });
  const tl = [
    ["1015", "Смерть Владимира. Святополк убивает братьев — Бориса, Глеба и Святослава"],
    ["1016", "Ярослав разбивает Святополка под Любечем и занимает Киев"],
    ["1018", "Святополк и польский король Болеслав Храбрый захватывают Киев"],
    ["1019", "Битва на реке Альте — окончательная победа Ярослава"],
    ["1024", "Поражение при Листвене; в 1026\u00A0г. Русь поделена с Мстиславом по Днепру"],
    ["1036", "Смерть Мстислава — Ярослав единовластный правитель Руси"],
  ];
  s.addShape(pres.shapes.LINE, { x: 0.75, y: 2.55, w: 8.5, h: 0, line: { color: C.accent1, width: 2 }, objectName: "Timeline line" });
  tl.forEach(([yr, d], i) => {
    const cx = 1.25 + i * 1.5;
    const key = yr === "1019";
    const dd = key ? 0.4 : 0.3;
    s.addShape(pres.shapes.OVAL, { x: cx - dd / 2, y: 2.55 - dd / 2, w: dd, h: dd, fill: { color: key ? C.accent2 : C.text2 }, line: { color: C.background1, width: 2 }, objectName: "Node " + yr });
    tb(s, yr, { x: cx - 0.7, y: 1.7, w: 1.4, h: 0.5, fontSize: 22, bold: true, color: key ? C.accent2 : C.text2, fontFace: "Cambria", align: "center", valign: "bottom", objectName: "Year " + yr });
    card(s, cx - 0.69, 3.0, 1.38, 1.7, key ? C.text2 : C.background2, "Event card " + yr);
    tb(s, d, { x: cx - 0.6, y: 3.12, w: 1.2, h: 1.5, fontSize: 12, color: key ? C.background1 : C.text1, valign: "top", objectName: "Event " + yr });
  });
  s.addNotes("После смерти Владимира в 1015 году между его сыновьями началась кровавая усобица. Святополк, прозванный Окаянным, расправился с братьями Борисом и Глебом — позже они стали первыми русскими святыми. Ярослав окончательно победил Святополка в 1019 году на реке Альте, но ещё долго делил власть с братом Мстиславом. Лишь после смерти Мстислава в 1036 году Ярослав стал единовластным правителем.");

  // ===== 4. Foreign policy =====
  pres.addSection({ title: "Правление" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Правление" });
  s.addText("Внешняя политика: меч и щит Руси", { placeholder: "title" });
  const fp = [
    ["swords", "Разгром печенегов, 1036", "Победа у стен Киева навсегда избавила Русь от набегов печенегов"],
    ["shield", "На западе, 1030–1031", "Возвращены Червенские города, в Прибалтике основан город Юрьев (Тарту)"],
    ["ship", "Поход на Византию, 1043", "Последний поход Руси на Константинополь; флот погиб в бурю, мир заключён в 1046\u00A0г."],
    ["fort", "Защита границ", "Основан Ярославль на Волге, на юге по реке Рось построена линия крепостей"],
  ];
  fp.forEach(([ic, h, d], i) => {
    const x = i % 2 === 0 ? 0.5 : 5.15, y = i < 2 ? 1.3 : 3.2;
    card(s, x, y, 4.35, 1.7, C.background2, h + " card");
    iconCircle(s, I[ic], x + 0.25, y + 0.22, 0.56, C.text2, h);
    tb(s, h, { x: x + 0.98, y: y + 0.22, w: 3.2, h: 0.56, fontSize: 16, bold: true, color: C.text2, fontFace: "Cambria", valign: "middle", objectName: h + " heading" });
    tb(s, d, { x: x + 0.25, y: y + 0.9, w: 3.85, h: 0.65, fontSize: 14, color: C.text1, valign: "top", objectName: h + " text" });
  });
  s.addNotes("Ярослав вёл активную внешнюю политику. В 1036 году он нанёс сокрушительное поражение печенегам — по преданию, на месте битвы был заложен Софийский собор. На западе Русь вернула Червенские города, на северо-западе был основан Юрьев. Поход на Византию в 1043 году закончился неудачей, но вскоре отношения были восстановлены.");

  // ===== 5. Dynastic marriages =====
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Правление" });
  s.addText("Ярослав — «тесть Европы»", { placeholder: "title" });
  const dau = [
    ["Анна", "Королева Франции", "Жена Генриха I, мать короля Филиппа I"],
    ["Елизавета", "Королева Норвегии", "Жена Харальда Сурового, короля Норвегии"],
    ["Анастасия", "Королева Венгрии", "Жена Андраша I, короля Венгрии"],
  ];
  dau.forEach(([n, t, d], i) => {
    const x = 0.5 + i * 3.1;
    card(s, x, 1.3, 2.8, 2.5, C.background2, n + " card");
    iconCircle(s, I.crown, x + 1.05, 1.5, 0.7, C.text2, n);
    tb(s, n, { x: x + 0.15, y: 2.32, w: 2.5, h: 0.45, fontSize: 21, bold: true, color: C.text2, fontFace: "Cambria", align: "center", valign: "middle", objectName: n + " name" });
    tb(s, t, { x: x + 0.15, y: 2.78, w: 2.5, h: 0.32, fontSize: 14, bold: true, color: C.accent2, align: "center", valign: "middle", objectName: n + " title" });
    tb(s, d, { x: x + 0.2, y: 3.12, w: 2.4, h: 0.55, fontSize: 13, color: C.text1, align: "center", valign: "top", objectName: n + " text" });
  });
  card(s, 0.5, 4.05, 9, 0.95, C.text2, "Sons banner");
  s.addImage({ data: I.tree, x: 0.75, y: 4.22, w: 0.6, h: 0.6, objectName: "Family tree icon", altText: "Родословное древо" });
  tb(s, [
    { text: "Сыновья ", options: { bold: true, color: C.accent1 } },
    { text: "женились на принцессах Польши, Византии и Германии — Русь стала равноправной частью европейского мира", options: { color: C.background1 } },
  ], { x: 1.6, y: 4.12, w: 7.7, h: 0.8, fontSize: 15, valign: "middle", objectName: "Sons text" });
  s.addNotes("Ярослава называют «тестем Европы». Его дочери стали королевами: Анна — Франции, Елизавета — Норвегии, Анастасия — Венгрии. Сыновья были женаты на польской, византийской и немецкой принцессах. Так Киевская Русь вошла в круг крупнейших европейских держав.");

  // ===== 6. Russkaya Pravda =====
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Правление" });
  s.addText("Свод законов «Русская Правда»", { placeholder: "title" });
  card(s, 0.5, 1.3, 3.4, 3.7, C.text2, "Pravda stat card");
  s.addImage({ data: I.scroll, x: 1.75, y: 1.55, w: 0.9, h: 0.9, objectName: "Scroll icon", altText: "Свиток" });
  tb(s, "40 гривен", { x: 0.7, y: 2.6, w: 3.0, h: 0.75, fontSize: 36, bold: true, color: C.accent1, fontFace: "Cambria", align: "center", valign: "middle", objectName: "Stat 40" });
  tb(s, "вира за убийство свободного человека", { x: 0.8, y: 3.35, w: 2.8, h: 0.65, fontSize: 15, color: C.background1, align: "center", valign: "top", objectName: "Stat 40 label" });
  tb(s, "Вира — денежный штраф в пользу князя", { x: 0.8, y: 4.25, w: 2.8, h: 0.5, fontSize: 12, italic: true, color: C.accent3, align: "center", valign: "middle", objectName: "Vira note" });
  const pr = [
    ["Ок. 1016 г. — ", "«Правда Ярослава», древнейшая часть первого письменного свода законов Руси"],
    ["Кровная месть ", "ограничена кругом ближайших родственников"],
    ["Если мстить некому, ", "убийца платит виру"],
    ["Размер штрафа ", "зависел от положения человека в обществе"],
    ["При сыновьях ", "дополнена «Правдой Ярославичей» (ок.\u00A01072\u00A0г.)"],
  ];
  pr.forEach(([b, t], i) => {
    const y = 1.4 + i * 0.72;
    s.addShape(pres.shapes.DIAMOND, { x: 4.3, y: y + 0.08, w: 0.16, h: 0.16, fill: { color: C.accent1 }, line: { color: C.accent1 }, objectName: "Marker " + (i + 1) });
    tb(s, [
      { text: b, options: { bold: true, color: C.text2 } },
      { text: t, options: { color: C.text1 } },
    ], { x: 4.6, y, w: 4.9, h: 0.62, fontSize: 15, valign: "top", objectName: "Pravda item " + (i + 1) });
  });
  s.addImage({ data: I.dScales, x: 8.95, y: 0.4, w: 0.55, h: 0.55, objectName: "Scales icon", altText: "Весы правосудия" });
  s.addNotes("«Русская Правда» — первый письменный свод законов Древней Руси. Его древнейшая часть, «Правда Ярослава», была дана новгородцам около 1016 года. Закон ограничивал кровную месть и заменял её денежным штрафом — вирой. Размер штрафа зависел от того, кем был пострадавший. Позже сыновья Ярослава дополнили закон «Правдой Ярославичей».");

  // ===== 7. Culture =====
  pres.addSection({ title: "Культура" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Культура" });
  s.addText("Культура и просвещение", { placeholder: "title" });
  card(s, 0.5, 1.3, 3.6, 3.7, C.background2, "Cathedral card");
  s.addImage({ data: cathedral, x: 0.65, y: 1.4, w: 3.3, h: 2.79, objectName: "Cathedral drawing", altText: "Рисунок Софийского собора в Киеве" });
  tb(s, "Софийский собор в Киеве, заложен в 1037 г.", { x: 0.7, y: 4.3, w: 3.2, h: 0.5, fontSize: 12, italic: true, color: C.text2, align: "center", valign: "middle", objectName: "Cathedral caption" });
  const cu = [
    ["church", "Софийский собор", "Главный храм Руси с мозаиками и фресками; рядом — Золотые ворота"],
    ["book", "Первые школы", "В Новгороде открыто училище, где обучались 300 детей"],
    ["books", "Книжность", "При Софийском соборе собрана библиотека, переводились греческие книги"],
    ["quill", "Митрополит Иларион, 1051", "Первый митрополит из русских, автор «Слова о законе и благодати»"],
  ];
  cu.forEach(([ic, h, d], i) => {
    const y = 1.3 + i * 0.95;
    iconCircle(s, I[ic], 4.5, y + 0.05, 0.52, C.text2, h);
    tb(s, h, { x: 5.25, y, w: 4.25, h: 0.32, fontSize: 15, bold: true, color: C.text2, fontFace: "Cambria", valign: "middle", objectName: h + " heading" });
    tb(s, d, { x: 5.25, y: y + 0.32, w: 4.25, h: 0.52, fontSize: 13, color: C.text1, valign: "top", objectName: h + " text" });
  });
  s.addNotes("При Ярославе Киев превратился в один из красивейших городов Европы. Был построен Софийский собор, украшенный мозаиками и фресками, и Золотые ворота. Князь заботился о просвещении: открывал школы, собирал книги, поощрял переводы. В 1051 году митрополитом впервые стал русский — Иларион, автор знаменитого «Слова о законе и благодати».");

  // ===== 8. Testament =====
  pres.addSection({ title: "Наследие" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Наследие" });
  s.addText("Завещание Ярослава", { placeholder: "title" });
  const ccx = 2.6, ccy = 3.15;
  const pills = [
    ["Святослав", "Чернигов", 0.5, 1.3], ["Всеволод", "Переяславль", 2.9, 1.3],
    ["Вячеслав", "Смоленск", 0.5, 4.25], ["Игорь", "Владимир-Волынский", 2.9, 4.25],
  ];
  pills.forEach(([n, c, x, y]) => {
    const px = x + 0.9, py = y + 0.375;
    const lx = Math.min(px, ccx), ly = Math.min(py, ccy);
    const flipH = (px > ccx) !== (py > ccy);
    s.addShape(pres.shapes.LINE, { x: lx, y: ly, w: Math.abs(px - ccx), h: Math.abs(py - ccy), flipH, line: { color: C.accent1, width: 1.5, dashType: "dash" }, objectName: n + " link" });
  });
  pills.forEach(([n, c, x, y]) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 1.8, h: 0.75, rectRadius: 0.37, fill: { color: C.background2 }, line: { color: C.accent1, width: 1 }, objectName: n + " pill" });
    tb(s, [
      { text: n, options: { bold: true, color: C.text2, fontSize: 14, breakLine: true } },
      { text: c, options: { color: C.text1, fontSize: 11 } },
    ], { x: x + 0.05, y: y + 0.05, w: 1.7, h: 0.65, align: "center", valign: "middle", objectName: n + " label" });
  });
  s.addShape(pres.shapes.OVAL, { x: ccx - 0.72, y: ccy - 0.72, w: 1.44, h: 1.44, fill: { color: C.text2 }, line: { color: C.accent1, width: 2 }, shadow: shadow(), objectName: "Izyaslav circle" });
  tb(s, [
    { text: "Изяслав", options: { bold: true, color: C.accent1, fontSize: 16, fontFace: "Cambria", breakLine: true } },
    { text: "Киев, Новгород", options: { color: C.background1, fontSize: 11 } },
  ], { x: ccx - 0.68, y: ccy - 0.4, w: 1.36, h: 0.8, align: "center", valign: "middle", objectName: "Izyaslav label" });

  card(s, 5.2, 1.3, 4.3, 2.4, C.background2, "Quote card");
  tb(s, "«", { x: 5.45, y: 1.38, w: 0.6, h: 0.55, fontSize: 40, bold: true, color: C.accent1, fontFace: "Cambria", valign: "top", objectName: "Quote mark" });
  tb(s, "Имейте любовь между собой, потому что вы братья, от одного отца и от одной матери", { x: 5.5, y: 1.98, w: 3.8, h: 1.1, fontSize: 16, italic: true, color: C.text2, fontFace: "Cambria", valign: "top", objectName: "Quote" });
  tb(s, "Завещание сыновьям, «Повесть временных лет»", { x: 5.5, y: 3.15, w: 3.8, h: 0.4, fontSize: 11, color: C.accent6, valign: "middle", objectName: "Quote source" });
  card(s, 5.2, 3.9, 4.3, 1.1, C.text2, "Order card");
  tb(s, [
    { text: "Лествичный порядок: ", options: { bold: true, color: C.accent1 } },
    { text: "власть переходит к старшему в роду. Но вскоре споры братьев приведут к усобицам", options: { color: C.background1 } },
  ], { x: 5.4, y: 3.98, w: 3.9, h: 0.94, fontSize: 13, valign: "middle", objectName: "Order text" });
  s.addNotes("Перед смертью в 1054 году Ярослав разделил земли между сыновьями. Старший, Изяслав, получил Киев и Новгород, остальные — другие крупные города. Князь завещал сыновьям жить в мире и слушаться старшего брата. Установился лествичный порядок наследования, однако со временем он стал причиной усобиц и привёл Русь к раздробленности.");

  // ===== 9. Conclusion =====
  s = pres.addSlide({ masterName: "END_DARK", sectionTitle: "Наследие" });
  s.addText("Итоги правления", { placeholder: "title" });
  const res = [
    ["dShield", "Единство и сила", "Русь объединена под властью Киева, границы надёжно защищены"],
    ["dScales", "Закон и порядок", "«Русская Правда» заложила основы древнерусского права"],
    ["dChurch", "Культура", "Храмы, школы и книги сделали Киев соперником Константинополя"],
  ];
  res.forEach(([ic, h, d], i) => {
    const x = 0.6 + i * 3.0;
    iconCircle(s, I[ic], x + 0.95, 1.55, 0.85, C.accent1, h);
    tb(s, h, { x, y: 2.55, w: 2.8, h: 0.45, fontSize: 19, bold: true, color: C.background1, fontFace: "Cambria", align: "center", valign: "middle", objectName: h + " heading" });
    tb(s, d, { x: x + 0.15, y: 3.02, w: 2.5, h: 0.85, fontSize: 14, color: C.accent3, align: "center", valign: "top", objectName: h + " text" });
  });
  tb(s, "При Ярославе Киевская Русь достигла вершины могущества", { x: 0.8, y: 4.3, w: 8.4, h: 0.55, fontSize: 18, italic: true, color: C.accent1, fontFace: "Cambria", align: "center", valign: "middle", objectName: "Tagline" });
  s.addNotes("Подводя итог: Ярослав укрепил единство Руси и защитил её границы, дал стране первый письменный свод законов, превратил Киев в крупнейший культурный центр. Не случайно в исторической традиции за ним закрепилось прозвище «Мудрый». Спасибо за внимание!");

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
