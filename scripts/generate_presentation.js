const fs = require("fs");
const path = require("path");
const JSZip = require("jszip");
const PptxGenJS = require("@bapunhansdah/pptxgenjs");

const ROOT = path.resolve(__dirname, "..");
const ASSET_DIR = path.join(ROOT, "assets", "generated");
const OUTPUT_DIR = path.join(ROOT, "output");
const PPTX_PATH = path.join(
  OUTPUT_DIR,
  "schule-wohlbefinden-keynote-style.pptx",
);
const NOTES_PATH = path.join(OUTPUT_DIR, "moderationsnotizen.md");

const COLORS = {
  bg: "FFFFFF",
  text: "424242",
  muted: "6B7280",
  blue: "CFEFFF",
  pink: "FFD9E8",
  yellow: "FFF7C7",
  green: "DDF4D7",
  line: "E8EAF0",
  darkBlue: "7BB9D6",
  darkPink: "E69BB9",
  darkYellow: "E6C96A",
  accentGreen: "86B97B",
  accentRed: "D98697",
};

const FONT_HEAD = "Avenir Next";
const FONT_BODY = "Avenir Next";
const SHADOW = undefined;

function ensureDir(dirPath) {
  fs.mkdirSync(dirPath, { recursive: true });
}

function svgWrap(inner) {
  return `<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="1600" height="900" viewBox="0 0 1600 900" fill="none">
  <rect width="1600" height="900" rx="56" fill="#FFFFFF"/>
  ${inner}
</svg>`;
}

function writeSvg(fileName, contents) {
  const fullPath = path.join(ASSET_DIR, fileName);
  fs.writeFileSync(fullPath, contents, "utf8");
  return fullPath;
}

function buildIllustrations() {
  ensureDir(ASSET_DIR);

  const hero = writeSvg(
    "slide1-school-hero.svg",
    svgWrap(`
      <ellipse cx="1260" cy="165" rx="190" ry="95" fill="#CFEFFF"/>
      <ellipse cx="1425" cy="690" rx="180" ry="110" fill="#FFD9E8" fill-opacity="0.72"/>
      <ellipse cx="220" cy="760" rx="220" ry="120" fill="#FFF7C7" fill-opacity="0.8"/>
      <rect x="520" y="250" width="560" height="330" rx="26" fill="#F8FBFD" stroke="#D9E7EE" stroke-width="3"/>
      <rect x="675" y="160" width="250" height="110" rx="22" fill="#FFFFFF" stroke="#D9E7EE" stroke-width="3"/>
      <path d="M645 255 L800 160 L955 255" fill="#FFD9E8" stroke="#EAB7CC" stroke-width="8" stroke-linejoin="round"/>
      <rect x="770" y="415" width="70" height="165" rx="18" fill="#CFEFFF"/>
      <rect x="595" y="315" width="90" height="72" rx="16" fill="#CFEFFF"/>
      <rect x="713" y="315" width="90" height="72" rx="16" fill="#FFF7C7"/>
      <rect x="835" y="315" width="90" height="72" rx="16" fill="#FFD9E8"/>
      <rect x="957" y="315" width="90" height="72" rx="16" fill="#CFEFFF"/>
      <rect x="595" y="410" width="90" height="72" rx="16" fill="#FFF7C7"/>
      <rect x="957" y="410" width="90" height="72" rx="16" fill="#FFF7C7"/>
      <rect x="210" y="620" width="1180" height="58" rx="29" fill="#EAF7EC"/>
      <circle cx="290" cy="560" r="42" fill="#FFD9E8"/>
      <rect x="250" y="602" width="80" height="96" rx="30" fill="#7BB9D6"/>
      <path d="M278 560 Q290 574 304 560" stroke="#424242" stroke-width="5" stroke-linecap="round"/>
      <circle cx="382" cy="576" r="18" fill="#FFF7C7"/>
      <circle cx="470" cy="560" r="42" fill="#FFF7C7"/>
      <rect x="430" y="602" width="80" height="96" rx="30" fill="#E69BB9"/>
      <path d="M455 548 Q470 535 485 548" stroke="#424242" stroke-width="5" stroke-linecap="round"/>
      <circle cx="1178" cy="560" r="42" fill="#FFD9E8"/>
      <rect x="1138" y="602" width="80" height="96" rx="30" fill="#7BB9D6"/>
      <path d="M1156 576 Q1178 556 1200 576" stroke="#424242" stroke-width="5" stroke-linecap="round"/>
      <line x1="1158" y1="546" x2="1168" y2="552" stroke="#424242" stroke-width="4" stroke-linecap="round"/>
      <line x1="1198" y1="552" x2="1208" y2="546" stroke="#424242" stroke-width="4" stroke-linecap="round"/>
      <circle cx="1290" cy="560" r="42" fill="#FFF7C7"/>
      <rect x="1250" y="602" width="80" height="96" rx="30" fill="#E69BB9"/>
      <path d="M1268 576 Q1290 556 1312 576" stroke="#424242" stroke-width="5" stroke-linecap="round"/>
      <line x1="1268" y1="546" x2="1278" y2="552" stroke="#424242" stroke-width="4" stroke-linecap="round"/>
      <line x1="1302" y1="552" x2="1312" y2="546" stroke="#424242" stroke-width="4" stroke-linecap="round"/>
      <circle cx="1340" cy="420" r="54" fill="#FFD9E8"/>
      <circle cx="1412" cy="410" r="54" fill="#CFEFFF"/>
      <circle cx="1470" cy="440" r="46" fill="#FFF7C7"/>
      <circle cx="1252" cy="215" r="30" fill="#FFF7C7"/>
      <line x1="1252" y1="160" x2="1252" y2="132" stroke="#E6C96A" stroke-width="7" stroke-linecap="round"/>
      <line x1="1282" y1="190" x2="1305" y2="167" stroke="#E6C96A" stroke-width="7" stroke-linecap="round"/>
      <line x1="1222" y1="190" x2="1199" y2="167" stroke="#E6C96A" stroke-width="7" stroke-linecap="round"/>
    `),
  );

  const balance = writeSvg(
    "slide2-balance.svg",
    svgWrap(`
      <ellipse cx="1280" cy="720" rx="200" ry="110" fill="#CFEFFF" fill-opacity="0.55"/>
      <ellipse cx="240" cy="170" rx="180" ry="95" fill="#FFD9E8" fill-opacity="0.72"/>
      <rect x="764" y="250" width="72" height="340" rx="28" fill="#424242"/>
      <rect x="640" y="570" width="320" height="34" rx="17" fill="#424242"/>
      <circle cx="800" cy="220" r="26" fill="#FFD9E8"/>
      <line x1="800" y1="246" x2="800" y2="305" stroke="#424242" stroke-width="10" stroke-linecap="round"/>
      <line x1="800" y1="305" x2="470" y2="390" stroke="#424242" stroke-width="10" stroke-linecap="round"/>
      <line x1="800" y1="305" x2="1130" y2="380" stroke="#424242" stroke-width="10" stroke-linecap="round"/>
      <line x1="470" y1="390" x2="470" y2="460" stroke="#424242" stroke-width="6" stroke-linecap="round"/>
      <line x1="1130" y1="380" x2="1130" y2="450" stroke="#424242" stroke-width="6" stroke-linecap="round"/>
      <path d="M340 460 H600 L560 550 H380 Z" fill="#FFF7C7" stroke="#D8C781" stroke-width="4"/>
      <path d="M1000 450 H1260 L1220 540 H1040 Z" fill="#CFEFFF" stroke="#9DCFE2" stroke-width="4"/>
      <rect x="390" y="408" width="74" height="92" rx="14" fill="#7BB9D6"/>
      <rect x="447" y="392" width="78" height="108" rx="14" fill="#FFD9E8"/>
      <rect x="510" y="418" width="52" height="82" rx="12" fill="#FFF7C7"/>
      <path d="M1092 416 C1068 382 1012 398 1012 452 C1012 504 1083 540 1130 582 C1175 540 1248 503 1248 452 C1248 398 1193 382 1168 416 C1151 438 1110 438 1092 416Z" fill="#E69BB9"/>
      <circle cx="1112" cy="480" r="7" fill="#FFFFFF"/>
      <circle cx="1150" cy="480" r="7" fill="#FFFFFF"/>
      <path d="M1098 512 Q1130 540 1162 512" stroke="#FFFFFF" stroke-width="8" stroke-linecap="round"/>
      <circle cx="300" cy="705" r="22" fill="#FFD9E8"/>
      <circle cx="350" cy="705" r="22" fill="#CFEFFF"/>
      <circle cx="400" cy="705" r="22" fill="#FFF7C7"/>
    `),
  );

  const stress = writeSvg(
    "slide5-stress.svg",
    svgWrap(`
      <ellipse cx="1230" cy="165" rx="190" ry="90" fill="#CFEFFF"/>
      <ellipse cx="310" cy="760" rx="220" ry="115" fill="#FFF7C7" fill-opacity="0.9"/>
      <rect x="380" y="590" width="860" height="54" rx="27" fill="#EAF7EC"/>
      <rect x="480" y="430" width="420" height="170" rx="28" fill="#FFFFFF" stroke="#E6EAF0" stroke-width="4"/>
      <rect x="960" y="300" width="180" height="300" rx="24" fill="#F8FBFD" stroke="#E6EAF0" stroke-width="4"/>
      <circle cx="680" cy="300" r="78" fill="#FFD9E8"/>
      <path d="M630 300 Q680 230 730 300" stroke="#424242" stroke-width="8" stroke-linecap="round"/>
      <path d="M640 333 Q680 308 720 333" stroke="#424242" stroke-width="7" stroke-linecap="round"/>
      <circle cx="650" cy="286" r="8" fill="#424242"/>
      <circle cx="710" cy="286" r="8" fill="#424242"/>
      <rect x="602" y="362" width="156" height="184" rx="44" fill="#7BB9D6"/>
      <path d="M550 500 L660 455" stroke="#424242" stroke-width="10" stroke-linecap="round"/>
      <path d="M800 500 L720 455" stroke="#424242" stroke-width="10" stroke-linecap="round"/>
      <rect x="965" y="345" width="170" height="90" rx="18" fill="#CFEFFF"/>
      <rect x="965" y="450" width="170" height="90" rx="18" fill="#FFD9E8"/>
      <rect x="520" y="470" width="270" height="26" rx="13" fill="#FFF7C7"/>
      <circle cx="1160" cy="180" r="58" fill="#FFF7C7"/>
      <circle cx="1160" cy="180" r="44" stroke="#424242" stroke-width="6"/>
      <line x1="1160" y1="180" x2="1160" y2="150" stroke="#424242" stroke-width="6" stroke-linecap="round"/>
      <line x1="1160" y1="180" x2="1188" y2="198" stroke="#424242" stroke-width="6" stroke-linecap="round"/>
      <line x1="1265" y1="290" x2="1320" y2="230" stroke="#D98697" stroke-width="10" stroke-linecap="round"/>
      <line x1="1320" y1="290" x2="1265" y2="230" stroke="#D98697" stroke-width="10" stroke-linecap="round"/>
      <rect x="1210" y="420" width="160" height="92" rx="22" fill="#FFF7C7"/>
      <path d="M1242 482 L1268 448 L1294 482 L1322 430" stroke="#424242" stroke-width="8" stroke-linecap="round" stroke-linejoin="round"/>
    `),
  );

  const dialogue = writeSvg(
    "slide6-dialogue.svg",
    svgWrap(`
      <ellipse cx="1320" cy="720" rx="200" ry="120" fill="#CFEFFF" fill-opacity="0.55"/>
      <ellipse cx="260" cy="160" rx="190" ry="100" fill="#FFD9E8" fill-opacity="0.72"/>
      <rect x="210" y="610" width="1180" height="58" rx="29" fill="#EAF7EC"/>
      <circle cx="720" cy="255" r="70" fill="#FFF7C7"/>
      <rect x="652" y="322" width="136" height="198" rx="42" fill="#7BB9D6"/>
      <path d="M628 430 L565 515" stroke="#424242" stroke-width="10" stroke-linecap="round"/>
      <path d="M814 430 L875 505" stroke="#424242" stroke-width="10" stroke-linecap="round"/>
      <circle cx="486" cy="330" r="58" fill="#FFD9E8"/>
      <rect x="432" y="388" width="108" height="164" rx="36" fill="#E69BB9"/>
      <circle cx="1060" cy="340" r="58" fill="#CFEFFF"/>
      <rect x="1006" y="398" width="108" height="164" rx="36" fill="#FFF7C7"/>
      <path d="M571 486 L642 450" stroke="#424242" stroke-width="9" stroke-linecap="round"/>
      <path d="M951 492 L868 452" stroke="#424242" stroke-width="9" stroke-linecap="round"/>
      <path d="M420 210 H555 C580 210 600 230 600 255 V296 C600 322 580 340 555 340 H490 L454 374 L462 340 H420 C395 340 375 322 375 296 V255 C375 230 395 210 420 210Z" fill="#FFFFFF" stroke="#D8E5ED" stroke-width="4"/>
      <path d="M980 190 H1190 C1216 190 1236 210 1236 236 V286 C1236 312 1216 332 1190 332 H1118 L1082 365 L1088 332 H980 C954 332 934 312 934 286 V236 C934 210 954 190 980 190Z" fill="#FFFFFF" stroke="#D8E5ED" stroke-width="4"/>
      <circle cx="452" cy="273" r="10" fill="#7BB9D6"/>
      <circle cx="490" cy="273" r="10" fill="#E69BB9"/>
      <circle cx="528" cy="273" r="10" fill="#E6C96A"/>
      <path d="M1010 274 Q1060 228 1110 274" stroke="#7BB9D6" stroke-width="7" stroke-linecap="round"/>
      <path d="M1010 274 Q1060 320 1110 274" stroke="#E69BB9" stroke-width="7" stroke-linecap="round"/>
    `),
  );

  const finale = writeSvg(
    "slide11-finale.svg",
    svgWrap(`
      <linearGradient id="sky" x1="0" y1="0" x2="0" y2="1">
        <stop offset="0%" stop-color="#FFFFFF"/>
        <stop offset="100%" stop-color="#F7FBFD"/>
      </linearGradient>
      <rect width="1600" height="900" rx="56" fill="url(#sky)"/>
      <ellipse cx="1350" cy="170" rx="200" ry="95" fill="#FFF7C7"/>
      <ellipse cx="1240" cy="755" rx="270" ry="125" fill="#CFEFFF" fill-opacity="0.6"/>
      <ellipse cx="290" cy="735" rx="280" ry="135" fill="#FFD9E8" fill-opacity="0.65"/>
      <rect x="190" y="640" width="1220" height="62" rx="31" fill="#EAF7EC"/>
      <rect x="575" y="330" width="450" height="250" rx="28" fill="#F9FCFE" stroke="#DAE7EE" stroke-width="3"/>
      <path d="M540 340 L800 215 L1060 340" fill="#CFEFFF" stroke="#A8D4E6" stroke-width="8" stroke-linejoin="round"/>
      <rect x="782" y="435" width="56" height="145" rx="16" fill="#FFF7C7"/>
      <rect x="635" y="385" width="92" height="72" rx="16" fill="#FFD9E8"/>
      <rect x="747" y="385" width="92" height="72" rx="16" fill="#CFEFFF"/>
      <rect x="859" y="385" width="92" height="72" rx="16" fill="#FFD9E8"/>
      <circle cx="320" cy="600" r="42" fill="#86B97B"/>
      <rect x="306" y="600" width="28" height="70" rx="14" fill="#9B7B59"/>
      <circle cx="1260" cy="590" r="44" fill="#86B97B"/>
      <rect x="1246" y="590" width="28" height="76" rx="14" fill="#9B7B59"/>
      <circle cx="1275" cy="230" r="28" fill="#FFF7C7"/>
      <line x1="1275" y1="180" x2="1275" y2="154" stroke="#E6C96A" stroke-width="7" stroke-linecap="round"/>
      <line x1="1302" y1="206" x2="1324" y2="186" stroke="#E6C96A" stroke-width="7" stroke-linecap="round"/>
      <line x1="1248" y1="206" x2="1226" y2="186" stroke="#E6C96A" stroke-width="7" stroke-linecap="round"/>
    `),
  );

  return { hero, balance, stress, dialogue, finale };
}

function anim(order, extra = {}) {
  return {
    type: extra.type || "fadein",
    duration: extra.duration || 450,
    trigger: order === 0 ? "withPrevious" : "afterPrevious",
    delay: order === 0 ? 0 : extra.delay || 120,
    ...extra,
  };
}

function addBlob(slide, x, y, w, h, color, transparency = 0, rotate = 0) {
  slide.addShape(slide._pptx.ShapeType.ellipse, {
    x,
    y,
    w,
    h,
    rotate,
    line: { color, transparency: 100 },
    fill: { color, transparency },
  });
}

function decorateSlide(slide, variant) {
  const variants = [
    [
      { x: -0.2, y: -0.25, w: 2.9, h: 1.4, color: COLORS.blue, transparency: 5 },
      { x: 11.1, y: 5.8, w: 2.5, h: 1.1, color: COLORS.pink, transparency: 25 },
    ],
    [
      { x: 10.8, y: -0.2, w: 2.6, h: 1.3, color: COLORS.yellow, transparency: 10 },
      { x: -0.4, y: 5.75, w: 2.9, h: 1.25, color: COLORS.blue, transparency: 25 },
    ],
    [
      { x: -0.35, y: -0.25, w: 2.4, h: 1.1, color: COLORS.pink, transparency: 10 },
      { x: 10.9, y: 6.0, w: 2.2, h: 0.95, color: COLORS.yellow, transparency: 20 },
    ],
  ];
  slide.background = { color: COLORS.bg };
  variants[variant % variants.length].forEach((shape) =>
    addBlob(
      slide,
      shape.x,
      shape.y,
      shape.w,
      shape.h,
      shape.color,
      shape.transparency,
    ),
  );

  slide.addShape(slide._pptx.ShapeType.line, {
    x: 0.55,
    y: 6.95,
    w: 12.15,
    h: 0,
    line: { color: COLORS.line, width: 1 },
  });
}

function addPageNumber(slide, number) {
  slide.addText(String(number).padStart(2, "0"), {
    x: 12.25,
    y: 7.0,
    w: 0.45,
    h: 0.22,
    fontFace: FONT_BODY,
    fontSize: 10,
    color: COLORS.muted,
    align: "right",
    margin: 0,
  });
}

function addTitleBlock(slide, title, subtitle, options = {}) {
  const x = options.x ?? 0.75;
  const y = options.y ?? 0.55;
  const w = options.w ?? 6.2;
  if (options.eyebrow) {
    slide.addText(options.eyebrow, {
      x,
      y,
      w: 2.4,
      h: 0.28,
      fontFace: FONT_BODY,
      fontSize: 11,
      color: COLORS.muted,
      bold: true,
      charSpacing: 0.6,
      animation: anim(0, { type: "wipe", direction: "left" }),
    });
  }
  slide.addText(title, {
    x,
    y: y + (options.eyebrow ? 0.28 : 0),
    w,
    h: options.titleH ?? 1.2,
    fontFace: FONT_HEAD,
    fontSize: options.titleSize ?? 24,
    bold: true,
    color: COLORS.text,
    margin: 0,
    fit: "shrink",
    animation: anim(1, { type: "fadein" }),
  });
  if (subtitle) {
    slide.addText(subtitle, {
      x,
      y: y + (options.eyebrow ? 1.1 : 0.82),
      w: options.subtitleW ?? w,
      h: options.subtitleH ?? 0.65,
      fontFace: FONT_BODY,
      fontSize: options.subtitleSize ?? 12.5,
      color: COLORS.muted,
      margin: 0,
      fit: "shrink",
      animation: anim(2, { type: "fadein" }),
    });
  }
}

function addCard(slide, opts) {
  const shapeOptions = {
    x: opts.x,
    y: opts.y,
    w: opts.w,
    h: opts.h,
    rectRadius: 0.12,
    fill: { color: opts.fill || "FFFFFF" },
    line: { color: opts.line || "FFFFFF", width: 1 },
    animation: opts.animation,
  };
  if (opts.shadow !== false && SHADOW) {
    shapeOptions.shadow = SHADOW;
  }
  slide.addShape(slide._pptx.ShapeType.roundRect, shapeOptions);
}

function addChip(slide, label, x, y, w, fill, order, textColor = COLORS.text) {
  const textOptions = {
    x,
    y,
    w,
    h: 0.34,
    shape: slide._pptx.ShapeType.roundRect,
    rectRadius: 0.12,
    fill: { color: fill },
    line: { color: fill, width: 1 },
    fontFace: FONT_BODY,
    fontSize: 10.5,
    color: textColor,
    bold: true,
    align: "center",
    valign: "middle",
    margin: 0,
    animation: anim(order, { type: "zoom", direction: "objectCenter" }),
  };
  if (SHADOW) {
    textOptions.shadow = SHADOW;
  }
  slide.addText(label, textOptions);
}

function addBulletLine(slide, text, x, y, w, order, color = COLORS.text) {
  slide.addText(`• ${text}`, {
    x,
    y,
    w,
    h: 0.32,
    fontFace: FONT_BODY,
    fontSize: 15,
    color,
    margin: 0,
    animation: anim(order, { type: "fadein" }),
  });
}

function addBodyText(slide, text, x, y, w, h, order, opts = {}) {
  slide.addText(text, {
    x,
    y,
    w,
    h,
    fontFace: opts.fontFace || FONT_BODY,
    fontSize: opts.fontSize || 12,
    color: opts.color || COLORS.text,
    bold: opts.bold || false,
    align: opts.align || "left",
    valign: opts.valign || "top",
    margin: opts.margin ?? 0,
    fit: opts.fit || "shrink",
    animation: opts.animation || anim(order, { type: "fadein" }),
  });
}

function buildSpeakerNotes() {
  return [
    {
      slide: 1,
      title: "Titel",
      text:
        "Ich stelle heute meine Facharbeit mit dem Titel \"Wie Schule wirkt\" vor. Im Mittelpunkt steht die Frage, wie strukturelle Unterschiede im deutschen Schulsystem das Wohlbefinden von Schülerinnen und Schülern beeinflussen. Dabei geht es also nicht nur um Noten oder Abschlüsse, sondern auch um Stress, Motivation, Selbstwertgefühl und die Frage, wie Schule im Alltag erlebt wird. Meine Arbeit zeigt, dass Schule zugleich Bildungsort und Lebensraum ist. Genau deshalb lohnt es sich, Schulstrukturen nicht nur aus der Perspektive von Leistung zu betrachten, sondern auch mit Blick auf psychische und soziale Folgen. Im Vortrag gehe ich deshalb zuerst auf die Relevanz des Themas ein, dann auf zentrale Begriffe, anschließend auf das deutsche Schulsystem, Belastungsfaktoren, internationale Vergleiche und mögliche Reformansätze. So lässt sich am Ende die Leitfrage klar beantworten. Damit beginne ich jetzt mit der Frage, warum dieses Thema überhaupt so wichtig ist.",
    },
    {
      slide: 2,
      title: "Warum ist das Thema wichtig?",
      text:
        "Das Thema ist wichtig, weil Schule einen großen Teil des Alltags von Kindern und Jugendlichen bestimmt. In meiner Facharbeit wird deutlich, dass Schule nicht nur Wissen vermittelt, sondern auch Bildungs- und Berufswege beeinflusst. Gleichzeitig erleben viele Jugendliche schulische Anforderungen als Belastung. Die Arbeit verweist dabei auf Leistungsdruck, Prüfungsstress und soziale Vergleichsprozesse als wichtige Stressfaktoren. Genau deshalb spielt Wohlbefinden in der Schule eine zentrale Rolle: Wer sich dauerhaft unter Druck, unsicher oder wenig unterstützt fühlt, lernt unter anderen Bedingungen als jemand, der Schule als sicheren und motivierenden Ort erlebt. Das Wohlbefinden ist also kein Nebenaspekt, sondern eng mit Motivation, Lernprozessen und Entwicklung verbunden. Aus dieser Relevanz ergibt sich die Leitfrage meiner Facharbeit: Inwiefern beeinflussen strukturelle Unterschiede im deutschen Schulsystem das Wohlbefinden von Schülerinnen und Schülern? Um diese Frage beantworten zu können, müssen zuerst einige grundlegende Begriffe geklärt werden.",
    },
    {
      slide: 3,
      title: "Grundlagen",
      text:
        "Für die Facharbeit waren drei Grundbegriffe besonders wichtig. Erstens die Strukturmerkmale von Schule: Dazu gehören zum Beispiel die verschiedenen Schulformen, das Notensystem, Versetzungsregelungen und Abschlüsse. Diese Rahmenbedingungen prägen den schulischen Alltag und beeinflussen Bildungswege. Zweitens das Wohlbefinden. In der Arbeit wird es als psychische und soziale Größe verstanden, also etwa als Lebenszufriedenheit, emotionale Sicherheit, soziale Einbindung und psychische Gesundheit. Es geht damit deutlich über reine Leistung hinaus. Drittens die Bildungsgerechtigkeit. Damit ist gemeint, dass Bildungszugänge und Chancen fair verteilt sein sollten und nicht stark von der sozialen Herkunft abhängen. Genau hier zeigt sich ein wichtiger Zusammenhang: Wenn Strukturen sozial selektiv wirken, beeinflussen sie nicht nur Erfolgschancen, sondern oft auch das Selbstbild und das Wohlbefinden. Mit diesen Begriffen im Hintergrund kann man nun das deutsche Schulsystem gezielter betrachten.",
    },
    {
      slide: 4,
      title: "Das deutsche Schulsystem",
      text:
        "Das deutsche Schulsystem ist traditionell mehrgliedrig. Nach der Grundschule werden Schülerinnen und Schüler in verschiedene Bildungswege aufgeteilt. Das Gymnasium ist stärker akademisch ausgerichtet und führt zum Abitur, bringt aber oft auch hohe Leistungsanforderungen mit sich. Die Realschule verbindet allgemeinbildende Inhalte mit mehr Praxisorientierung und eröffnet unterschiedliche Anschlussmöglichkeiten. Die Hauptschule war besonders auf praktische und berufliche Wege ausgerichtet, wird heute aber häufig kritisch gesehen, unter anderem wegen möglicher Stigmatisierung. Die Gesamtschule setzt auf längeres gemeinsames Lernen und versucht, Unterschiede stärker innerhalb einer Schule auszugleichen. Dazu kommen berufliche Schulen, die im dualen System eng mit Betrieben zusammenarbeiten. In meiner Facharbeit wird deutlich, dass diese Struktur sowohl Chancen als auch Probleme mit sich bringt. Besonders wichtig ist nun die Frage, wie sich diese Organisation auf Druck und Wohlbefinden im Schulalltag auswirkt.",
    },
    {
      slide: 5,
      title: "Leistungsdruck und Wohlbefinden",
      text:
        "Ein zentrales Ergebnis meiner Facharbeit ist, dass Leistungsdruck das Wohlbefinden deutlich beeinflussen kann. Auslöser sind vor allem Noten, Prüfungen, Konkurrenzsituationen und hohe Erwartungen von außen. Schulische Leistungen sind eng mit Bildungs- und Berufschancen verbunden, deshalb werden Anforderungen oft als besonders folgenreich erlebt. Gleichzeitig zeigt die Arbeit, dass die Folgen sehr unterschiedlich ausfallen können. Häufig genannt werden Stress, Schlafprobleme, Unsicherheit und ein geringeres Selbstwertgefühl. Besonders problematisch wird es, wenn schulischer Erfolg stark mit persönlichem Wert verknüpft wird oder Misserfolge sich wiederholen. Wichtig ist aber auch: Leistungsanforderungen werden nicht nur negativ erlebt. Manche Schülerinnen und Schüler empfinden sie durchaus als motivierend. Genau deshalb lautet der Merksatz auf dieser Folie, dass Leistungsdruck Motivation fördern kann, aber eben auch das Wohlbefinden beeinträchtigt. Neben Druck spielt jedoch noch ein zweiter großer Faktor eine Rolle: das Schulklima.",
    },
    {
      slide: 6,
      title: "Schulklima",
      text:
        "Neben Leistungsanforderungen ist das soziale Umfeld der Schule entscheidend. In der Facharbeit wird das Schulklima als Qualität der Beziehungen, als Atmosphäre und als Gefühl von Sicherheit und Zugehörigkeit beschrieben. Besonders wichtig ist das Lehrer-Schüler-Verhältnis. Wenn Lehrkräfte respektvoll, unterstützend und verlässlich handeln, kann das Stress reduzieren und Lernmotivation fördern. Ein weiterer Punkt ist die Mitbestimmung. Beteiligungsmöglichkeiten wie Schülervertretungen, Klassenräte oder Projekte stärken das Gefühl, selbst etwas bewirken zu können. Das ist für die Selbstwirksamkeit von Schülerinnen und Schülern sehr bedeutsam. Hinzu kommt die Feedbackkultur. Die Arbeit zeigt, dass nicht nur Noten wichtig sind, sondern auch qualitative Rückmeldungen, die Lernfortschritte sichtbar machen und individuelle Entwicklung stärker berücksichtigen. Ein positives Schulklima kann das Wohlbefinden also aktiv fördern. Nachdem wir nun Druck und Beziehungen betrachtet haben, möchte ich euch kurz direkt einbeziehen.",
    },
    {
      slide: 7,
      title: "Kurze Umfrage",
      text:
        "An dieser Stelle würde ich gern kurz die Runde einbeziehen. Die Frage lautet: Was beeinflusst euer Wohlbefinden in der Schule am stärksten? Ist es eher der Notendruck, sind es die Lehrkräfte, die Mitschülerinnen und Mitschüler oder der allgemeine Zeitstress? Hier würde ich kurz um ein Handzeichen bitten. Diese kleine Umfrage passt gut zur Facharbeit, weil dort deutlich wird, dass sich Wohlbefinden nie nur aus einem einzigen Faktor ergibt. Vielmehr wirken mehrere Aspekte zusammen. Leistungsdruck, Schulklima, soziale Beziehungen und schulische Anforderungen greifen ineinander. Genau deshalb ist die Frage nach Schulstruktur so spannend: Strukturen formen nicht nur Unterricht, sondern auch Erfahrungen. Nach dieser kurzen Aktivierung schaue ich im nächsten Schritt auf andere Länder. Der internationale Vergleich hilft dabei, das deutsche System besser einzuordnen und mögliche Alternativen sichtbar zu machen.",
    },
    {
      slide: 8,
      title: "Internationaler Vergleich",
      text:
        "Der internationale Vergleich zeigt, dass es auch andere Wege gibt, Schule zu organisieren. Deutschland ist durch eine frühe Aufteilung und eine starke Rolle von Noten und Prüfungen geprägt. Finnland setzt dagegen auf längeres gemeinsames Lernen und eine spätere Leistungsdifferenzierung. In der Facharbeit wird hervorgehoben, dass dadurch soziale Unterschiede teilweise geringer ausfallen können. Angelsächsische Systeme arbeiten häufig mit gemeinsamen Schulen, mehr Wahlmöglichkeiten und stärkerer Praxisorientierung. Dabei können Schülerinnen und Schüler zum Teil früher eigene Schwerpunkte setzen. Gleichzeitig macht die Arbeit aber auch klar, dass kein System perfekt ist. Jedes Bildungssystem muss unterschiedliche gesellschaftliche Anforderungen erfüllen. Entscheidend ist also nicht, ein Land einfach zu kopieren, sondern zu verstehen, welche strukturellen Entscheidungen das Wohlbefinden eher fördern und welche eher Druck verstärken. Mit diesem Vergleich im Blick lassen sich die Stärken und Schwächen des deutschen Systems nun genauer bewerten.",
    },
    {
      slide: 9,
      title: "Stärken und Schwächen",
      text:
        "In der kritischen Bewertung meiner Facharbeit zeigt sich ein gemischtes Bild. Zu den Stärken des deutschen Bildungssystems gehört, dass es verschiedene Bildungswege gibt. Dadurch können unterschiedliche Fähigkeiten und Interessen grundsätzlich berücksichtigt werden. Besonders positiv wird außerdem das duale Ausbildungssystem bewertet, weil es schulische Bildung mit praktischer Ausbildung verbindet und den Übergang in den Arbeitsmarkt erleichtert. Auch die hohen akademischen Standards, vor allem im Gymnasium, werden als Stärke genannt. Dem stehen jedoch deutliche Schwächen gegenüber. Kritisch ist vor allem die frühe Selektion, weil Bildungsentscheidungen in einem Alter getroffen werden, in dem soziale Einflüsse sehr stark wirken. Hinzu kommen Leistungsdruck und soziale Ungleichheiten. Die Arbeit zeigt also: Das System ist nicht einfach gut oder schlecht, sondern widersprüchlich. Genau daraus leitet sich die Frage ab, welche Reformen denkbar wären, um Bildungsqualität und Wohlbefinden besser zusammenzubringen.",
    },
    {
      slide: 10,
      title: "Reformansätze",
      text:
        "Die Facharbeit diskutiert drei zentrale Reformansätze. Erstens mehr Lebenskompetenzen im Curriculum. Gemeint sind zum Beispiel finanzielle Grundbildung, Medienkompetenz und praktische Alltagsfähigkeiten. Solche Inhalte könnten Schule für viele Jugendliche lebensnäher machen. Zweitens eine spätere Leistungsdifferenzierung. Wenn Schülerinnen und Schüler länger gemeinsam lernen, hätten sie mehr Zeit, Fähigkeiten zu entwickeln und fundiertere Bildungsentscheidungen zu treffen. Drittens alternative Bewertungsformen. Neben klassischen Noten könnten ausführlichere Lernfeedbacks, Portfolioarbeit oder projektbasierte Bewertungen stärker genutzt werden. Der gemeinsame Gedanke hinter allen drei Ansätzen ist, Lernen weniger eng auf Prüfungsergebnisse zu verengen und individuelle Entwicklung stärker mitzudenken. Reformen sollen also nicht Leistung abschaffen, sondern Schule ausgewogener gestalten. Damit komme ich zum Schluss und zur eigentlichen Antwort auf die Leitfrage meiner Facharbeit.",
    },
    {
      slide: 11,
      title: "Fazit",
      text:
        "Zusammenfassend lässt sich sagen: Schule beeinflusst deutlich mehr als nur Noten. Erstens wirken schulische Strukturen unmittelbar auf das Wohlbefinden, weil sie Lernbedingungen, Vergleichssituationen und Bildungswege mitbestimmen. Zweitens spielen Leistungsdruck und Schulklima eine besonders zentrale Rolle. Sie entscheiden oft darüber, ob Schule als motivierender oder belastender Raum erlebt wird. Drittens zeigen die diskutierten Reformansätze, dass Bildung und Wohlbefinden stärker zusammengedacht werden könnten. Die Leitfrage meiner Facharbeit lässt sich also so beantworten: Strukturelle Unterschiede im deutschen Schulsystem beeinflussen das Wohlbefinden von Schülerinnen und Schülern in erheblichem Maß, sowohl durch frühe Selektion und Leistungsbewertung als auch durch Schulklima und Beteiligungsmöglichkeiten. Damit bin ich am Ende meiner Präsentation. Vielen Dank für eure Aufmerksamkeit.",
    },
  ];
}

function createPresentation(assets) {
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "Cursor";
  pptx.company = "Cursor Cloud";
  pptx.subject =
    "Wie Schule wirkt: Strukturelle Unterschiede und ihr Effekt auf das Wohlbefinden von Schülerinnen und Schülern";
  pptx.title =
    "Wie Schule wirkt: Strukturelle Unterschiede und ihr Effekt auf das Wohlbefinden von Schülerinnen und Schülern";
  pptx.theme = {
    headFontFace: FONT_HEAD,
    bodyFontFace: FONT_BODY,
  };

  const notes = buildSpeakerNotes();

  // Helper to expose shape enums inside decorators.
  pptx.ShapeType = pptx.ShapeType;

  // Slide 1
  let slide = pptx.addSlide();
  slide._pptx = pptx;
  decorateSlide(slide, 0);
  addPageNumber(slide, 1);
  addChip(slide, "Seminarfachkurs Glück", 0.78, 0.62, 1.9, COLORS.blue, 0);
  slide.addText("Wie Schule wirkt:", {
    x: 0.8,
    y: 1.2,
    w: 5.8,
    h: 0.7,
    fontFace: FONT_HEAD,
    fontSize: 27,
    bold: true,
    color: COLORS.text,
    margin: 0,
    animation: anim(1, { type: "fadein" }),
  });
  slide.addText(
    "Strukturelle Unterschiede und ihr Effekt auf das Wohlbefinden von Schülerinnen und Schülern",
    {
      x: 0.8,
      y: 1.9,
      w: 5.9,
      h: 1.8,
      fontFace: FONT_HEAD,
      fontSize: 20,
      bold: true,
      color: COLORS.text,
      margin: 0,
      fit: "shrink",
      animation: anim(2, { type: "fadein" }),
    },
  );
  slide.addText("Präsentation auf Grundlage der Facharbeit", {
    x: 0.8,
    y: 3.75,
    w: 4.7,
    h: 0.36,
    fontFace: FONT_BODY,
    fontSize: 12.5,
    color: COLORS.muted,
    margin: 0,
    animation: anim(3, { type: "fadein" }),
  });
  slide.addImage(withOptionalShadow({
    path: assets.hero,
    x: 6.65,
    y: 0.78,
    w: 5.85,
    h: 5.95,
    animation: anim(4, { type: "zoom", direction: "objectCenter", duration: 700 }),
  }));
  slide.addNotes(notes[0].text);

  // Slide 2
  slide = pptx.addSlide();
  slide._pptx = pptx;
  decorateSlide(slide, 1);
  addPageNumber(slide, 2);
  addTitleBlock(slide, "Warum ist das Thema wichtig?", "", {
    eyebrow: "Relevanz",
    x: 0.8,
    y: 0.58,
    w: 5.6,
  });
  addCard(slide, {
    x: 0.8,
    y: 1.55,
    w: 5.25,
    h: 3.18,
    fill: "FFFFFF",
    line: "EEF0F5",
    animation: anim(3, { type: "fadein" }),
  });
  const whyBullets = [
    "Schule prägt den Alltag junger Menschen.",
    "Bildung beeinflusst Bildungs- und Berufswege.",
    "Leistungsdruck und Anforderungen werden oft als Belastung erlebt.",
    "Wohlbefinden ist eng mit Motivation, Lernen und Entwicklung verbunden.",
  ];
  whyBullets.forEach((bullet, index) =>
    addBulletLine(slide, bullet, 1.08, 1.95 + index * 0.58, 4.6, 4 + index),
  );
  slide.addImage(withOptionalShadow({
    path: assets.balance,
    x: 7.05,
    y: 1.2,
    w: 5.1,
    h: 3.7,
    animation: anim(8, { type: "zoom", direction: "objectCenter" }),
  }));
  addCard(slide, {
    x: 0.95,
    y: 5.25,
    w: 11.45,
    h: 1.18,
    fill: COLORS.blue,
    line: COLORS.blue,
    animation: anim(9, { type: "wipe", direction: "left" }),
  });
  addBodyText(
    slide,
    "Inwiefern beeinflussen strukturelle Unterschiede im deutschen Schulsystem das Wohlbefinden von Schülerinnen und Schülern?",
    1.25,
    5.52,
    10.9,
    0.55,
    10,
    { fontSize: 17, bold: true, align: "center" },
  );
  slide.addNotes(notes[1].text);

  // Slide 3
  slide = pptx.addSlide();
  slide._pptx = pptx;
  decorateSlide(slide, 2);
  addPageNumber(slide, 3);
  addTitleBlock(slide, "Grundlagen", "Die Facharbeit arbeitet mit drei zentralen Begriffen.", {
    eyebrow: "Begriffsrahmen",
    x: 0.8,
    y: 0.58,
    w: 5,
  });
  slide.addShape(pptx.ShapeType.ellipse, withOptionalShadow({
    x: 5.2,
    y: 2.45,
    w: 2.95,
    h: 1.32,
    fill: { color: "FFFFFF" },
    line: { color: COLORS.line, width: 1.5 },
    animation: anim(3, { type: "zoom", direction: "objectCenter" }),
  }));
  addBodyText(slide, "Wie Schule wirkt", 5.6, 2.83, 2.1, 0.28, 4, {
    fontSize: 18,
    bold: true,
    align: "center",
  });
  const conceptCards = [
    {
      title: "Strukturmerkmale",
      body: "Schulformen\nNotensystem\nVersetzungen\nAbschlüsse",
      x: 1.05,
      y: 2.1,
      fill: COLORS.blue,
    },
    {
      title: "Wohlbefinden",
      body: "psychische Gesundheit\nsoziale Beziehungen\nLebenszufriedenheit",
      x: 8.65,
      y: 1.9,
      fill: COLORS.yellow,
    },
    {
      title: "Bildungsgerechtigkeit",
      body: "faire Chancen\nunabhängig von sozialer Herkunft",
      x: 4.2,
      y: 4.35,
      fill: COLORS.pink,
    },
  ];
  conceptCards.forEach((card, index) => {
    addCard(slide, {
      x: card.x,
      y: card.y,
      w: 3.05,
      h: index === 2 ? 1.7 : 2.15,
      fill: "FFFFFF",
      line: "EEF0F5",
      animation: anim(5 + index, { type: "fadein" }),
    });
    slide.addShape(pptx.ShapeType.roundRect, {
      x: card.x + 0.18,
      y: card.y + 0.18,
      w: 1.42,
      h: 0.28,
      rectRadius: 0.12,
      fill: { color: card.fill },
      line: { color: card.fill, width: 1 },
      animation: anim(6 + index, { type: "fadein" }),
    });
    addBodyText(slide, card.title, card.x + 0.22, card.y + 0.55, 2.45, 0.35, 7 + index, {
      fontSize: 15.5,
      bold: true,
    });
    addBodyText(slide, card.body, card.x + 0.22, card.y + 0.95, 2.45, 1, 8 + index, {
      fontSize: 12.2,
      color: COLORS.muted,
    });
  });
  slide.addShape(pptx.ShapeType.line, {
    x: 4.1,
    y: 3.0,
    w: 1.08,
    h: 0,
    line: { color: COLORS.line, width: 1.5 },
  });
  slide.addShape(pptx.ShapeType.line, {
    x: 8.15,
    y: 2.95,
    w: 0.48,
    h: -0.02,
    line: { color: COLORS.line, width: 1.5 },
  });
  slide.addShape(pptx.ShapeType.line, {
    x: 6.62,
    y: 3.78,
    w: -0.8,
    h: 0.56,
    line: { color: COLORS.line, width: 1.5 },
  });
  slide.addNotes(notes[2].text);

  // Slide 4
  slide = pptx.addSlide();
  slide._pptx = pptx;
  decorateSlide(slide, 0);
  addPageNumber(slide, 4);
  addTitleBlock(slide, "Das deutsche Schulsystem", "Fünf zentrale Schulformen im Überblick.", {
    eyebrow: "Mehrgliedrigkeit",
    x: 0.8,
    y: 0.58,
    w: 5.6,
  });
  addChip(slide, "Aufteilung nach der Grundschule", 0.8, 1.52, 2.5, COLORS.yellow, 3);
  const schoolCards = [
    {
      icon: "🎓",
      title: "Gymnasium",
      body: "hohe akademische Anforderungen\nAbitur\nhäufig hoher Prüfungsdruck",
      x: 0.9,
      y: 2.05,
      fill: COLORS.blue,
    },
    {
      icon: "🧭",
      title: "Realschule",
      body: "mittlerer Abschluss\nmehr Praxisbezug\nverschiedene Anschlüsse",
      x: 4.1,
      y: 2.05,
      fill: COLORS.yellow,
    },
    {
      icon: "🛠",
      title: "Hauptschule",
      body: "stärker praktisch ausgerichtet\ndirekter Bezug zur Ausbildung\nteils stigmatisiert",
      x: 7.3,
      y: 2.05,
      fill: COLORS.pink,
    },
    {
      icon: "🔀",
      title: "Gesamtschule",
      body: "längeres gemeinsames Lernen\nmehr Durchlässigkeit\nverschiedene Abschlüsse",
      x: 2.25,
      y: 4.55,
      fill: COLORS.pink,
    },
    {
      icon: "💼",
      title: "Berufliche Schule",
      body: "duales System\nSchule und Betrieb\nÜbergang in den Arbeitsmarkt",
      x: 6.35,
      y: 4.55,
      fill: COLORS.blue,
    },
  ];
  schoolCards.forEach((card, index) => {
    addCard(slide, {
      x: card.x,
      y: card.y,
      w: 2.7,
      h: 1.95,
      fill: "FFFFFF",
      line: "EEF0F5",
      animation: anim(4 + index, { type: "fadein" }),
    });
    addChip(slide, card.icon, card.x + 0.18, card.y + 0.18, 0.42, card.fill, 5 + index);
    addBodyText(slide, card.title, card.x + 0.18, card.y + 0.68, 2.12, 0.34, 6 + index, {
      fontSize: 14.5,
      bold: true,
    });
    addBodyText(slide, card.body, card.x + 0.18, card.y + 1.05, 2.25, 0.65, 7 + index, {
      fontSize: 11.6,
      color: COLORS.muted,
    });
  });
  slide.addNotes(notes[3].text);

  // Slide 5
  slide = pptx.addSlide();
  slide._pptx = pptx;
  decorateSlide(slide, 1);
  addPageNumber(slide, 5);
  addTitleBlock(slide, "Leistungsdruck und Wohlbefinden", "", {
    eyebrow: "Belastungsfaktor",
    x: 0.8,
    y: 0.58,
    w: 6.3,
  });
  slide.addImage(withOptionalShadow({
    path: assets.stress,
    x: 7.9,
    y: 1.15,
    w: 4.2,
    h: 4.05,
    animation: anim(3, { type: "zoom", direction: "objectCenter" }),
  }));
  addCard(slide, {
    x: 0.9,
    y: 1.72,
    w: 3.2,
    h: 3.18,
    fill: "FFFFFF",
    line: "EEF0F5",
    animation: anim(4, { type: "fadein" }),
  });
  addCard(slide, {
    x: 4.45,
    y: 1.72,
    w: 3.2,
    h: 3.18,
    fill: "FFFFFF",
    line: "EEF0F5",
    animation: anim(5, { type: "fadein" }),
  });
  addBodyText(slide, "Leistungsdruck", 1.18, 2.0, 2.5, 0.3, 6, {
    fontSize: 18,
    bold: true,
  });
  addBodyText(slide, "Folgen", 4.73, 2.0, 2.2, 0.3, 7, {
    fontSize: 18,
    bold: true,
  });
  ["Noten", "Prüfungen", "Konkurrenz", "Erwartungen"].forEach((label, index) =>
    addChip(slide, label, 1.18, 2.42 + index * 0.5, 1.78, COLORS.blue, 8 + index),
  );
  ["Stress", "Schlafprobleme", "Unsicherheit", "geringeres Selbstwertgefühl"].forEach(
    (label, index) =>
      addChip(slide, label, 4.73, 2.42 + index * 0.5, 2.2, COLORS.pink, 12 + index),
  );
  addCard(slide, {
    x: 1.05,
    y: 5.42,
    w: 10.9,
    h: 0.76,
    fill: COLORS.yellow,
    line: COLORS.yellow,
    animation: anim(17, { type: "wipe", direction: "left" }),
  });
  addBodyText(
    slide,
    "Leistungsdruck kann Motivation fördern, aber auch das Wohlbefinden beeinträchtigen.",
    1.35,
    5.65,
    10.35,
    0.26,
    18,
    { fontSize: 15, bold: true, align: "center" },
  );
  slide.addNotes(notes[4].text);

  // Slide 6
  slide = pptx.addSlide();
  slide._pptx = pptx;
  decorateSlide(slide, 2);
  addPageNumber(slide, 6);
  addTitleBlock(slide, "Schulklima", "Beziehungen und Beteiligung prägen Schule ebenso stark wie Leistung.", {
    eyebrow: "Soziales Umfeld",
    x: 0.8,
    y: 0.58,
    w: 6.2,
  });
  slide.addImage(withOptionalShadow({
    path: assets.dialogue,
    x: 8.0,
    y: 0.95,
    w: 4.1,
    h: 2.45,
    animation: anim(3, { type: "zoom", direction: "objectCenter" }),
  }));
  const climateCards = [
    {
      title: "Lehrer-Schüler-Verhältnis",
      fill: COLORS.blue,
      body: ["Unterstützung und Respekt", "ernst genommen werden", "kann Stress reduzieren"],
      x: 0.95,
      y: 2.2,
    },
    {
      title: "Mitbestimmung",
      fill: COLORS.yellow,
      body: ["Schülervertretung", "Klassenrat und Projekte", "stärkt Selbstwirksamkeit"],
      x: 4.28,
      y: 2.2,
    },
    {
      title: "Feedbackkultur",
      fill: COLORS.pink,
      body: ["nicht nur Noten", "qualitative Rückmeldungen", "individuelle Entwicklung im Blick"],
      x: 7.61,
      y: 2.2,
    },
  ];
  climateCards.forEach((card, index) => {
    addCard(slide, {
      x: card.x,
      y: card.y,
      w: 3.0,
      h: 3.15,
      fill: "FFFFFF",
      line: "EEF0F5",
      animation: anim(4 + index, { type: "fadein" }),
    });
    addChip(slide, card.title, card.x + 0.2, card.y + 0.2, 2.15, card.fill, 6 + index);
    card.body.forEach((bullet, bulletIndex) =>
      addBulletLine(
        slide,
        bullet,
        card.x + 0.22,
        card.y + 0.85 + bulletIndex * 0.56,
        2.5,
        9 + index + bulletIndex,
      ),
    );
  });
  slide.addNotes(notes[5].text);

  // Slide 7
  slide = pptx.addSlide();
  slide._pptx = pptx;
  decorateSlide(slide, 0);
  addPageNumber(slide, 7);
  addTitleBlock(slide, "Kurze Umfrage", "Publikum einbinden", {
    eyebrow: "Interaktion",
    x: 0.8,
    y: 0.58,
    w: 4.2,
  });
  addCard(slide, {
    x: 1.35,
    y: 1.55,
    w: 10.6,
    h: 1.55,
    fill: "FFFFFF",
    line: "EEF0F5",
    animation: anim(3, { type: "fadein" }),
  });
  addBodyText(
    slide,
    "Was beeinflusst euer Wohlbefinden in der Schule am stärksten?",
    1.85,
    2.05,
    9.6,
    0.48,
    4,
    { fontSize: 21, bold: true, align: "center" },
  );
  const buttons = [
    { label: "📚  Notendruck", fill: COLORS.blue, x: 2.1, y: 3.55 },
    { label: "👩‍🏫  Lehrkräfte", fill: COLORS.pink, x: 6.75, y: 3.55 },
    { label: "👥  Mitschüler", fill: COLORS.yellow, x: 2.1, y: 4.65 },
    { label: "🕒  Zeitstress", fill: COLORS.green, x: 6.75, y: 4.65 },
  ];
  buttons.forEach((button, index) =>
    slide.addText(button.label, withOptionalShadow({
      x: button.x,
      y: button.y,
      w: 3.2,
      h: 0.62,
      shape: pptx.ShapeType.roundRect,
      rectRadius: 0.12,
      fill: { color: button.fill },
      line: { color: button.fill, width: 1 },
      fontFace: FONT_BODY,
      fontSize: 16,
      color: COLORS.text,
      bold: true,
      align: "center",
      valign: "middle",
      margin: 0,
      animation: anim(5 + index, { type: "zoom", direction: "objectCenter" }),
    })),
  );
  addBodyText(slide, "Kurzes Handzeichen", 5.05, 5.62, 3.3, 0.25, 9, {
    fontSize: 12.5,
    color: COLORS.muted,
    align: "center",
  });
  addCard(slide, {
    x: 2.2,
    y: 6.05,
    w: 8.95,
    h: 0.58,
    fill: COLORS.blue,
    line: COLORS.blue,
    animation: {
      type: "fadein",
      duration: 500,
      trigger: "afterPrevious",
      delay: 450,
    },
  });
  addBodyText(
    slide,
    "Alle Faktoren spielen laut Facharbeit eine wichtige Rolle.",
    2.55,
    6.23,
    8.25,
    0.22,
    11,
    { fontSize: 14.5, bold: true, align: "center" },
  );
  slide.addNotes(notes[6].text);

  // Slide 8
  slide = pptx.addSlide();
  slide._pptx = pptx;
  decorateSlide(slide, 1);
  addPageNumber(slide, 8);
  addTitleBlock(slide, "Internationaler Vergleich", "Deutschland, Finnland und angelsächsische Systeme im Kontrast.", {
    eyebrow: "Einordnung",
    x: 0.8,
    y: 0.58,
    w: 6.6,
  });
  const compareCards = [
    {
      title: "Deutschland",
      body: ["frühe Aufteilung", "Noten und Prüfungen zentral"],
      fill: COLORS.blue,
      x: 0.95,
    },
    {
      title: "Finnland",
      body: ["längeres gemeinsames Lernen", "spätere Leistungsdifferenzierung"],
      fill: COLORS.yellow,
      x: 4.35,
    },
    {
      title: "Angelsächsische Systeme",
      body: ["mehr Praxisorientierung", "mehr Wahlmöglichkeiten"],
      fill: COLORS.pink,
      x: 7.75,
    },
  ];
  compareCards.forEach((card, index) => {
    addCard(slide, {
      x: card.x,
      y: 2.1,
      w: 2.95,
      h: 3.45,
      fill: "FFFFFF",
      line: "EEF0F5",
      animation: anim(3 + index, { type: "fadein" }),
    });
    addChip(slide, card.title, card.x + 0.2, card.y + 0.2, 1.85, card.fill, 5 + index);
    card.body.forEach((bullet, bulletIndex) =>
      addBulletLine(
        slide,
        bullet,
        card.x + 0.22,
        card.y + 1.0 + bulletIndex * 0.85,
        2.34,
        7 + index + bulletIndex,
      ),
    );
  });
  addCard(slide, {
    x: 1.65,
    y: 6.0,
    w: 10.0,
    h: 0.62,
    fill: COLORS.yellow,
    line: COLORS.yellow,
    animation: anim(10, { type: "wipe", direction: "left" }),
  });
  addBodyText(slide, "Kein System ist perfekt.", 1.95, 6.2, 9.4, 0.2, 11, {
    fontSize: 16,
    bold: true,
    align: "center",
  });
  slide.addNotes(notes[7].text);

  // Slide 9
  slide = pptx.addSlide();
  slide._pptx = pptx;
  decorateSlide(slide, 2);
  addPageNumber(slide, 9);
  addTitleBlock(slide, "Stärken und Schwächen", "Eine ausgewogene Bewertung des deutschen Bildungssystems.", {
    eyebrow: "Kritische Einordnung",
    x: 0.8,
    y: 0.58,
    w: 6.4,
  });
  addCard(slide, {
    x: 0.95,
    y: 2.0,
    w: 5.3,
    h: 3.8,
    fill: COLORS.green,
    line: COLORS.green,
    animation: anim(3, { type: "fadein" }),
  });
  addCard(slide, {
    x: 7.05,
    y: 2.0,
    w: 5.3,
    h: 3.8,
    fill: COLORS.pink,
    line: COLORS.pink,
    animation: anim(4, { type: "fadein" }),
  });
  addBodyText(slide, "Stärken", 1.3, 2.35, 1.5, 0.3, 5, {
    fontSize: 21,
    bold: true,
  });
  addBodyText(slide, "Schwächen", 7.4, 2.35, 1.8, 0.3, 6, {
    fontSize: 21,
    bold: true,
  });
  [
    "verschiedene Bildungswege",
    "duales Ausbildungssystem",
    "hohe akademische Standards",
  ].forEach((bullet, index) =>
    addBulletLine(slide, bullet, 1.32, 3.02 + index * 0.72, 4.1, 7 + index),
  );
  ["frühe Selektion", "Leistungsdruck", "soziale Ungleichheiten"].forEach(
    (bullet, index) =>
      addBulletLine(slide, bullet, 7.42, 3.02 + index * 0.72, 4.1, 10 + index),
  );
  slide.addNotes(notes[8].text);

  // Slide 10
  slide = pptx.addSlide();
  slide._pptx = pptx;
  decorateSlide(slide, 0);
  addPageNumber(slide, 10);
  addTitleBlock(slide, "Reformansätze", "Drei Vorschläge aus der Facharbeit für eine ausgewogenere Schule.", {
    eyebrow: "Blick nach vorn",
    x: 0.8,
    y: 0.58,
    w: 6.6,
  });
  slide.addShape(pptx.ShapeType.chevron, {
    x: 1.1,
    y: 5.55,
    w: 10.7,
    h: 0.6,
    fill: { color: "F6F8FB" },
    line: { color: "F6F8FB", width: 1 },
    rotate: 0,
    animation: anim(3, { type: "wipe", direction: "left" }),
  });
  const reforms = [
    {
      title: "Mehr Lebenskompetenzen",
      body: "finanzielle Grundbildung,\nMedienkompetenz,\nAlltagswissen",
      fill: COLORS.blue,
      x: 1.25,
    },
    {
      title: "Spätere Leistungsdifferenzierung",
      body: "mehr Zeit für Entwicklung\nund fundiertere\nBildungsentscheidungen",
      fill: COLORS.yellow,
      x: 4.65,
    },
    {
      title: "Alternative Bewertungsformen",
      body: "Lernfeedback,\nPortfolioarbeit,\nprojektbasierte Bewertung",
      fill: COLORS.pink,
      x: 8.05,
    },
  ];
  reforms.forEach((reform, index) => {
    slide.addShape(pptx.ShapeType.ellipse, withOptionalShadow({
      x: reform.x,
      y: 2.2,
      w: 2.25,
      h: 2.25,
      fill: { color: reform.fill },
      line: { color: reform.fill, width: 1 },
      animation: anim(4 + index, { type: "zoom", direction: "objectCenter" }),
    }));
    addBodyText(slide, reform.title, reform.x + 0.2, 2.78, 1.85, 0.75, 7 + index, {
      fontSize: 14.5,
      bold: true,
      align: "center",
      valign: "middle",
    });
    addBodyText(slide, reform.body, reform.x - 0.05, 4.72, 2.35, 0.95, 9 + index, {
      fontSize: 12.2,
      color: COLORS.muted,
      align: "center",
    });
  });
  slide.addNotes(notes[9].text);

  // Slide 11
  slide = pptx.addSlide();
  slide._pptx = pptx;
  decorateSlide(slide, 1);
  addPageNumber(slide, 11);
  addTitleBlock(
    slide,
    "Schule beeinflusst mehr als nur Noten.",
    "",
    {
      eyebrow: "Fazit",
      x: 0.8,
      y: 0.58,
      w: 6.8,
    },
  );
  slide.addImage(withOptionalShadow({
    path: assets.finale,
    x: 7.8,
    y: 1.05,
    w: 4.4,
    h: 4.0,
    animation: anim(3, { type: "zoom", direction: "objectCenter" }),
  }));
  const conclusions = [
    "✓ Strukturen wirken auf das Wohlbefinden.",
    "✓ Leistungsdruck und Schulklima spielen eine zentrale Rolle.",
    "✓ Reformen könnten Bildung und Wohlbefinden stärker verbinden.",
  ];
  conclusions.forEach((line, index) =>
    addBulletLine(slide, line.replace(/^✓ /, "✓ "), 1.0, 2.05 + index * 0.78, 6.25, 4 + index),
  );
  addCard(slide, {
    x: 1.0,
    y: 5.55,
    w: 5.8,
    h: 0.78,
    fill: COLORS.blue,
    line: COLORS.blue,
    animation: anim(8, { type: "wipe", direction: "left" }),
  });
  addBodyText(
    slide,
    "Danke für eure Aufmerksamkeit!",
    1.28,
    5.82,
    5.2,
    0.25,
    9,
    { fontSize: 18, bold: true, align: "center" },
  );
  slide.addNotes(notes[10].text);

  return { pptx, notes };
}

async function patchTransitions(filePath) {
  const zip = await JSZip.loadAsync(fs.readFileSync(filePath));
  const transitionMap = {
    1: `<p:transition spd="slow"><p:fade/></p:transition>`,
    2: `<p:transition spd="med"><p:zoom dir="in"/></p:transition>`,
    3: `<p:transition spd="med"><p:cover dir="r"/></p:transition>`,
    4: `<mc:AlternateContent><mc:Choice Requires="p15"><p:transition spd="med"><p15:morph p15:option="byObject"/></p:transition></mc:Choice><mc:Fallback><p:transition spd="med"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>`,
    5: `<p:transition spd="med"><p:zoom dir="out"/></p:transition>`,
    6: `<p:transition spd="med"><p:push dir="l"/></p:transition>`,
    7: `<p:transition spd="med"><p:dissolve/></p:transition>`,
    8: `<p:transition spd="fast"><p:zoom dir="in"/></p:transition>`,
    9: `<p:transition spd="med"><p:wipe dir="r"/></p:transition>`,
    10: `<p:transition spd="med"><p:wheel spokes="1"/></p:transition>`,
    11: `<p:transition spd="slow"><p:fade/></p:transition>`,
  };

  for (const [index, transitionXml] of Object.entries(transitionMap)) {
    const slidePath = `ppt/slides/slide${index}.xml`;
    const file = zip.file(slidePath);
    if (!file) continue;
    let xml = await file.async("string");
    if (Number(index) === 4 && !xml.includes("xmlns:p15")) {
      xml = xml.replace(
        "<p:sld ",
        '<p:sld xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2015/09/main" mc:Ignorable="p15" ',
      );
    }
    xml = xml.replace("</p:cSld>", `</p:cSld>${transitionXml}`);
    zip.file(slidePath, xml);
  }

  const buffer = await zip.generateAsync({ type: "nodebuffer" });
  fs.writeFileSync(filePath, buffer);
}

function writeNotesMarkdown(notes) {
  const lines = [
    "# Moderationsnotizen",
    "",
    "Diese Notizen wurden parallel zur Präsentation erstellt und zusätzlich in die Sprecheransicht der PPTX-Datei eingebettet.",
    "",
  ];
  notes.forEach((note) => {
    lines.push(`## Folie ${note.slide}: ${note.title}`, "");
    lines.push(note.text, "");
  });
  fs.writeFileSync(NOTES_PATH, lines.join("\n"), "utf8");
}

function withOptionalShadow(baseOptions) {
  if (!SHADOW) return baseOptions;
  return { ...baseOptions, shadow: SHADOW };
}

async function main() {
  ensureDir(OUTPUT_DIR);
  const assets = buildIllustrations();
  const { pptx, notes } = createPresentation(assets);
  await pptx.writeFile({ fileName: PPTX_PATH, compression: true });
  await patchTransitions(PPTX_PATH);
  writeNotesMarkdown(notes);
  console.log(`Wrote ${PPTX_PATH}`);
  console.log(`Wrote ${NOTES_PATH}`);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
