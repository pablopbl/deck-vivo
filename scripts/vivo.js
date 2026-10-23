// deck-vivo — biblioteca para montar apresentações na identidade visual Vivo com pptxgenjs.
//
//   const vivo = require("<skill>/scripts/vivo.js");
//   const deck = vivo.createDeck({ title: "Título do arquivo" });
//   deck.cover({ kicker: "Vivo · Projeto", title: "Título", subtitle: "Subtítulo", meta: "Autor", date: "Outubro / 2026" });
//   const s = deck.content({ section: "Síntese", kicker: "01 · Síntese", title: "Mensagem do slide" });
//   await deck.statCards(s, [{ value: "64%", label: "...", detail: "...", icon: "FaTrophy" }]);
//   await deck.save("saida.pptx");
//
// Canvas 13,333" x 7,5" (LAYOUT_WIDE). Área útil dos slides de conteúdo: x 0,6" a 12,733", y 2,2" a 6,6".
const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");

const ASSETS = path.join(__dirname, "..", "assets");
const A = (f) => path.join(ASSETS, f);

// ---------- Identidade ----------
const HEX = {
  ink: "1A1A1A", white: "FFFFFF", roxo: "660099", lilasBg: "F3ECF8", roxoMedio: "8533AD",
  lilas: "C9A7E0", magenta: "CD1E5A", magentaBg: "FCE9EF", verde: "32A04B", cinza: "666666", grade: "E6E0EB",
};
const THEME = {
  name: "Vivo",
  headFontFace: "Arial",
  bodyFontFace: "Arial",
  colors: {
    dk1: HEX.ink, lt1: HEX.white, dk2: HEX.roxo, lt2: HEX.lilasBg,
    accent1: HEX.roxo, accent2: HEX.roxoMedio, accent3: HEX.lilas,
    accent4: HEX.magenta, accent5: HEX.verde, accent6: HEX.cinza,
    hlink: HEX.roxo, folHlink: HEX.roxoMedio,
  },
};
const W = 13.333, M = 0.6, CW = W - 2 * M;
const GAP = 0.3, TOP = 2.2, BOTTOM = 6.6;
const LOGO_R = 105 / 36, VIV_R = 25 / 34;

// ---------- Tema (cores no theme1.xml) ----------
function loadJSZip() {
  return require(require.resolve("jszip", { paths: [require.resolve("pptxgenjs")] }));
}
async function writeTheme(file) {
  const zip = await loadJSZip().loadAsync(fs.readFileSync(file));
  const part = "ppt/theme/theme1.xml";
  const slots = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
  const scheme = `<a:clrScheme name="${THEME.name}">` + slots.map((k) => `<a:${k}><a:srgbClr val="${THEME.colors[k]}"/></a:${k}>`).join("") + "</a:clrScheme>";
  const xml = (await zip.file(part).async("string"))
    .replace(/<a:clrScheme\b[\s\S]*?<\/a:clrScheme>/, () => scheme)
    .replace(/(<a:(?:theme|fontScheme)\b[^>]*?\bname=")[^"]*"/g, (_, h) => `${h}${THEME.name}"`);
  zip.file(part, xml);
  // Rótulos de dados numa linha só: o PowerPoint quebra "62,0%" em duas linhas quando a coluna é estreita.
  for (const name of Object.keys(zip.files).filter((f) => /^ppt\/charts\/chart\d+\.xml$/.test(f))) {
    const cx = (await zip.file(name).async("string")).replace(/<c:dLbls>[\s\S]*?<\/c:dLbls>/g, (blk) =>
      blk.replace(/<a:bodyPr(\s[^>]*)?\/>/g, (m, attrs = "") => /wrap=/.test(attrs) ? m : `<a:bodyPr${attrs} wrap="none"/>`)
         .replace(/<a:bodyPr(\s[^>]*)?>/g, (m, attrs = "") => /wrap=/.test(attrs) || m.endsWith("/>") ? m : `<a:bodyPr${attrs} wrap="none">`));
    zip.file(name, cx);
  }
  for (const name of Object.keys(zip.files)) {
    if (!name.endsWith(".xml")) continue;
    const bad = (await zip.file(name).async("string")).match(/<a:srgbClr val="((?![0-9A-Fa-f]{6}")[^"]*)"/);
    if (bad) throw new Error(`${name}: cor "${bad[1]}" não é hex de 6 dígitos. Use deck.HEX em opções que só aceitam hex (chartColors, valGridLine, shadow).`);
  }
  fs.writeFileSync(file, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
}

// ---------- Ícones (react-icons/fa6 -> PNG) ----------
const iconCache = {};
async function iconData(name, color = "#FFFFFF") {
  const key = name + color;
  if (iconCache[key]) return iconCache[key];
  const React = require("react");
  const RDS = require("react-dom/server");
  const sharp = require("sharp");
  const fa = require("react-icons/fa6");
  if (!fa[name]) throw new Error(`Ícone "${name}" não existe em react-icons/fa6 (ex.: FaTrophy, FaChartPie, FaHandshake).`);
  const svg = RDS.renderToStaticMarkup(React.createElement(fa[name], { color, size: "256" }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return (iconCache[key] = "image/png;base64," + buf.toString("base64"));
}

// ---------- Deck ----------
function createDeck(opts = {}) {
  const partnerLogo = opts.partnerLogo === undefined ? A("logo-parceiro.png") : opts.partnerLogo; // null = sem logo do parceiro
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.title = opts.title || "Apresentação Vivo";
  if (opts.author) pres.author = opts.author;
  if (opts.company) pres.company = String(opts.company).replace(/&/g, "&amp;");
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  const C = pres.SchemeColor;
  const col = {
    roxo: C.accent1, roxoMedio: C.accent2, lilas: C.accent3, magenta: C.accent4, verde: C.accent5, cinza: C.accent6,
    texto: C.text1, branco: C.background1, fundo: C.background2, magentaBg: HEX.magentaBg,
  };

  // Layouts
  pres.defineSlideMaster({
    title: "VIVO_CONTEUDO",
    background: { color: col.branco },
    objects: [
      { image: { x: M, y: 6.92, w: 0.26 * LOGO_R, h: 0.26, path: A("logo-vivo-purple.png") } },
      ...(partnerLogo ? [{ image: { x: 12.33, y: 6.83, w: 0.4, h: 0.4, path: partnerLogo } }] : []),
      { placeholder: { options: { name: "kicker", type: "body", x: M, y: 0.42, w: 10, h: 0.32, fontSize: 12, bold: true, color: col.roxoMedio, charSpacing: 1, margin: 0, valign: "middle" }, text: "" } },
      { placeholder: { options: { name: "title", type: "title", x: M, y: 0.78, w: CW, h: 1.05, fontSize: 28, bold: true, color: col.roxo, margin: 0, valign: "top", align: "left" }, text: "" } },
    ],
    slideNumber: { x: 11.55, y: 6.88, w: 0.6, h: 0.3, fontSize: 10, color: col.cinza, align: "right" },
  });
  pres.defineSlideMaster({
    title: "VIVO_DESTAQUE",
    background: { color: col.roxo },
    objects: [
      { image: { x: 8.75, y: 0.55, w: 6.4 * VIV_R, h: 6.4, path: A("vivinho-white.png"), transparency: 88 } },
      { image: { x: M, y: 0.55, w: 0.42 * LOGO_R, h: 0.42, path: A("logo-vivo-white.png") } },
      ...(partnerLogo ? [
        { rect: { x: 12.13, y: 6.62, w: 0.6, h: 0.6, fill: { color: col.branco } } },
        { image: { x: 12.21, y: 6.7, w: 0.44, h: 0.44, path: partnerLogo } },
      ] : []),
    ],
  });

  let n = 0;
  const nm = (s) => `${s}-${++n}`;
  let currentSection = null;
  const section = (title) => {
    if (title && title !== currentSection) { pres.addSection({ title }); currentSection = title; }
    return currentSection || undefined;
  };

  // ---- Primitivas ----
  const text = (slide, t, o) => slide.addText(t, { isTextBox: true, margin: 0, objectName: nm("texto"), ...o });
  const card = (slide, x, y, w, h, fill = col.fundo) =>
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { type: "none" }, rectRadius: 0.16, objectName: nm("card") });
  const pill = (slide, x, y, w, h, fill) =>
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { type: "none" }, rectRadius: Math.min(0.12, h / 2), objectName: nm("pill") });
  async function iconDot(slide, x, y, d, name, fill = col.roxo) {
    slide.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { type: "none" }, objectName: nm("icone-fundo") });
    const p = d * 0.25;
    slide.addImage({ data: await iconData(name), x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p, objectName: nm("icone") });
  }
  const chip = (slide, x, y, color) =>
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 0.2, h: 0.2, fill: { color }, line: { type: "none" }, rectRadius: 0.05, objectName: nm("legenda") });
  // items: [{ label, color, width }]
  function legend(slide, x, y, items) {
    let cx = x;
    for (const it of items) {
      chip(slide, cx, y + 0.04, it.color);
      text(slide, it.label, { x: cx + 0.3, y, w: it.width || 1.6, h: 0.28, fontSize: 12, color: col.texto, valign: "middle" });
      cx += 0.3 + (it.width || 1.6) + 0.3;
    }
  }
  const note = (slide, t, o = {}) => text(slide, t, { x: M, y: 6.25, w: CW, h: 0.3, fontSize: 12, color: col.cinza, italic: true, ...o });

  // ---- Slides ----
  function cover({ kicker, title, subtitle, meta, date, notes, section: sec = "Abertura" }) {
    const s = pres.addSlide({ masterName: "VIVO_DESTAQUE", sectionTitle: section(sec) });
    if (kicker) text(s, kicker, { x: M, y: 2.2, w: 8, h: 0.4, fontSize: 16, bold: true, color: col.lilas });
    text(s, title, { x: M, y: 2.65, w: 10.5, h: 1.1, fontSize: 52, bold: true, color: col.branco });
    if (subtitle) text(s, subtitle, { x: M, y: 3.8, w: 8, h: 0.55, fontSize: 26, color: col.branco });
    if (meta) text(s, meta, { x: M, y: 5.45, w: 8, h: 0.35, fontSize: 14, color: col.lilas });
    if (date) {
      pill(s, M, 5.95, 1.9, 0.42, col.branco);
      text(s, date, { x: M, y: 5.95, w: 1.9, h: 0.42, fontSize: 13, bold: true, color: col.roxo, align: "center", valign: "middle" });
    }
    if (notes) s.addNotes(notes);
    return s;
  }
  function content({ section: sec, kicker, title, notes }) {
    const s = pres.addSlide({ masterName: "VIVO_CONTEUDO", sectionTitle: section(sec) });
    if (kicker) s.addText(kicker.toUpperCase(), { placeholder: "kicker" });
    s.addText(title, { placeholder: "title" });
    if (notes) s.addNotes(notes);
    return s;
  }
  // runs: string ou [{ text, highlight: true }]
  function statement({ section: sec, kicker, runs, body, notes }) {
    const s = pres.addSlide({ masterName: "VIVO_DESTAQUE", sectionTitle: section(sec) });
    if (kicker) text(s, kicker.toUpperCase(), { x: M, y: 1.65, w: 8, h: 0.35, fontSize: 14, bold: true, color: col.lilas, charSpacing: 1 });
    const r = typeof runs === "string" ? [{ text: runs }] : runs;
    text(s, r.map((x) => ({ text: x.text, options: x.highlight ? { color: col.lilas } : {} })), { x: M, y: 2.15, w: 8.6, h: 2.6, fontSize: 34, bold: true, color: col.branco, valign: "top" });
    if (body) text(s, body, { x: M, y: 4.95, w: 8.2, h: 1.1, fontSize: 17, color: col.branco, valign: "top" });
    if (notes) s.addNotes(notes);
    return s;
  }
  function closing({ section: sec, title = "Obrigado.", subtitle, notes }) {
    const s = pres.addSlide({ masterName: "VIVO_DESTAQUE", sectionTitle: section(sec) });
    text(s, title, { x: M, y: 2.9, w: 8, h: 1.2, fontSize: 60, bold: true, color: col.branco });
    if (subtitle) text(s, subtitle, { x: M, y: 4.15, w: 8, h: 0.45, fontSize: 18, color: col.lilas });
    if (notes) s.addNotes(notes);
    return s;
  }

  // ---- Composições ----
  // items: [{ value, label, detail, icon }]; o card heroIndex (padrão 0) sai em roxo sólido.
  async function statCards(slide, items, { y = TOP, h = 4.3, heroIndex = 0, valueSize } = {}) {
    const k = items.length, cw = (CW - (k - 1) * GAP) / k;
    const vs = valueSize || (k <= 3 ? 56 : 44);
    for (let i = 0; i < k; i++) {
      const it = items[i], x = M + i * (cw + GAP), hero = i === heroIndex;
      card(slide, x, y, cw, h, hero ? col.roxo : col.fundo);
      let ty = y + 0.35;
      if (it.icon) { await iconDot(slide, x + 0.4, ty, 0.7, it.icon, hero ? col.roxoMedio : col.roxo); ty += 0.95; }
      text(slide, it.value, { x: x + 0.4, y: ty, w: cw - 0.8, h: 1.0, fontSize: vs, bold: true, color: hero ? col.branco : col.roxo });
      text(slide, it.label || "", { x: x + 0.4, y: ty + 1.05, w: cw - 0.8, h: 0.8, fontSize: 18, bold: true, color: hero ? col.branco : col.texto, valign: "top" });
      if (it.detail) text(slide, it.detail, { x: x + 0.4, y: ty + 1.9, w: cw - 0.8, h: y + h - (ty + 1.9) - 0.2, fontSize: 14, color: hero ? col.lilas : col.cinza, valign: "top" });
    }
  }
  // Comparação em barras horizontais com valor grande. rows: [{ label, value (número), display, sub, color }]
  function compareBars(slide, rows, { y = TOP, h = 2.55, badge, footnote } = {}) {
    card(slide, M, y, CW, h);
    const max = Math.max(...rows.map((r) => r.value)), barX = M + 2.6, barMax = badge ? 6.4 : 7.6;
    rows.forEach((r, i) => {
      const ry = y + 0.4 + i * 1.0, bw = barMax * r.value / max;
      text(slide, r.label, { x: M + 0.4, y: ry, w: 2.1, h: 0.75, fontSize: 16, bold: true, color: col.texto, valign: "middle" });
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: barX, y: ry + 0.05, w: bw, h: 0.65, fill: { color: r.color || (i ? col.roxo : col.lilas) }, line: { type: "none" }, rectRadius: 0.12, objectName: nm("barra") });
      text(slide, [{ text: r.display, options: { fontSize: 32, bold: true, color: col.roxo } }, { text: r.sub ? "   " + r.sub : "", options: { fontSize: 13, color: col.cinza } }], { x: barX + bw + 0.2, y: ry, w: 4.2, h: 0.75, valign: "middle" });
    });
    if (badge) {
      slide.addShape(pres.shapes.OVAL, { x: M + CW - 1.65, y: y + 0.45, w: 1.3, h: 1.3, fill: { color: col.magenta }, line: { type: "none" }, objectName: nm("selo") });
      text(slide, [{ text: badge.value, options: { fontSize: 26, bold: true, breakLine: true } }, { text: badge.label || "", options: { fontSize: 11 } }], { x: M + CW - 1.65, y: y + 0.45, w: 1.3, h: 1.3, color: col.branco, align: "center", valign: "middle" });
    }
    if (footnote) text(slide, footnote, { x: M + 0.4, y: y + h - 0.45, w: CW - 0.8, h: 0.3, fontSize: 11, color: col.cinza, italic: true });
  }
  // Cards com ícone, título e texto em grade. items: [{ title, text, icon, bullets: [..] }]
  async function iconCards(slide, items, { cols = 2, y = TOP, h = BOTTOM - TOP, heroIndex = -1 } = {}) {
    const rows = Math.ceil(items.length / cols), cw = (CW - (cols - 1) * GAP) / cols, ch = (h - (rows - 1) * GAP) / rows;
    for (let i = 0; i < items.length; i++) {
      const it = items[i], x = M + (i % cols) * (cw + GAP), cy = y + Math.floor(i / cols) * (ch + GAP), hero = i === heroIndex;
      card(slide, x, cy, cw, ch, hero ? col.roxo : col.fundo);
      const d = 0.7;
      const compact = ch < 2.2 && !it.bullets;
      if (it.icon) await iconDot(slide, x + 0.4, cy + (compact ? 0.3 : 0.35), compact ? 0.6 : d, it.icon, hero ? col.roxoMedio : col.roxo);
      const tx = it.icon ? x + 1.35 : x + 0.4;
      if (!(ch < 2.2 && !it.bullets)) text(slide, it.title, { x: tx, y: cy + 0.35, w: x + cw - 0.4 - tx, h: d, fontSize: 18, bold: true, color: hero ? col.branco : col.roxo, valign: "middle" });
      const body = it.bullets
        ? it.bullets.map((b, j) => ({ text: b, options: { bullet: true, breakLine: j < it.bullets.length - 1 } }))
        : it.text || "";
      if (ch < 2.2 && !it.bullets) { // compacto: título e texto ao lado do ícone
        text(slide, it.title, { x: tx, y: cy + 0.3, w: x + cw - 0.4 - tx, h: 0.4, fontSize: 16, bold: true, color: hero ? col.branco : col.roxo });
        text(slide, body, { x: tx, y: cy + 0.72, w: x + cw - 0.4 - tx, h: ch - 0.9, fontSize: 14, color: hero ? col.branco : col.texto, valign: "top" });
        continue;
      }
      const by = cy + 0.35 + d + 0.2;
      text(slide, body, { x: x + 0.4, y: by, w: cw - 0.8, h: cy + ch - by - 0.25, fontSize: 14, color: hero ? col.branco : col.texto, valign: "top", paraSpaceAfter: it.bullets ? 8 : 0 });
    }
  }
  // Faixa de destaque. tone: "alerta" (magenta) | "info" (roxo)
  async function callout(slide, t, { y = 5.75, h = 0.8, tone = "alerta", icon = "FaTriangleExclamation", label } = {}) {
    const fill = tone === "alerta" ? col.magentaBg : col.fundo, dotc = tone === "alerta" ? col.magenta : col.roxo;
    card(slide, M, y, CW, h, fill);
    await iconDot(slide, M + 0.25, y + (h - 0.5) / 2, 0.5, icon, dotc);
    const runs = label ? [{ text: label + " ", options: { bold: true, color: dotc } }, { text: t }] : t;
    text(slide, runs, { x: M + 0.95, y, w: CW - 1.2, h, fontSize: 15, color: col.texto, valign: "middle" });
  }
  // Tabela no padrão: cabeçalho roxo, zebra lilás. header: [..], rows: [[..]], emphasisCol destaca uma coluna em roxo negrito.
  function table(slide, header, rows, { x = M, y = TOP, w = CW, colW, rowH = 0.55, emphasisCol = -1, positiveCol = -1, fontSize = 14 } = {}) {
    const al = (j) => (j === 0 ? "left" : "center");
    slide.addTable([
      header.map((t, j) => ({ text: t, options: { bold: true, color: col.branco, fill: { color: col.roxo }, align: al(j) } })),
      ...rows.map((r, i) => r.map((t, j) => ({ text: String(t), options: {
        fill: { color: i % 2 ? col.fundo : col.branco }, align: al(j),
        bold: j === 0 || j === emphasisCol || j === positiveCol,
        color: j === positiveCol ? col.verde : j === emphasisCol ? col.roxo : col.texto,
      } }))),
    ], { x, y, w, colW, rowH, fontSize, valign: "middle", margin: [0, 0.15, 0, 0.15], border: { type: "solid", pt: 0.75, color: HEX.grade }, objectName: nm("tabela") });
  }
  // Painel roxo com número grande (ou frase) ao lado de outro conteúdo.
  async function heroPanel(slide, { x = M, y = TOP, w = 3.5, h = BOTTOM - TOP, value, label, detail, icon, statementText }) {
    card(slide, x, y, w, h, col.roxo);
    let ty = y + 0.4;
    if (icon) { await iconDot(slide, x + 0.4, ty, 0.7, icon, col.roxoMedio); ty += 0.95; }
    if (value) { text(slide, value, { x: x + 0.4, y: ty, w: w - 0.8, h: 1.1, fontSize: 64, bold: true, color: col.branco }); ty += 1.15; }
    if (label) { text(slide, label, { x: x + 0.4, y: ty, w: w - 0.8, h: 0.75, fontSize: 16, bold: true, color: col.branco, valign: "top" }); ty += 0.8; }
    if (detail) text(slide, detail, { x: x + 0.4, y: ty, w: w - 0.8, h: y + h - ty - 0.3, fontSize: 14, color: col.lilas, valign: "top" });
    if (statementText) text(slide, statementText, { x: x + 0.4, y: ty, w: w - 0.8, h: y + h - ty - 0.3, fontSize: 17, bold: true, color: col.branco, valign: "top" });
  }
  // Gráfico de colunas/barras no padrão. series: [{ name, labels, values }]
  function barChart(slide, series, o = {}) {
    const pct = o.percent;
    slide.addChart(pres.charts.BAR, series, {
      x: M, y: 2.4, w: CW, h: 4.1,
      catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", legendFontFace: "+mn-lt",
      catAxisLabelColor: HEX.ink, valAxisLabelColor: HEX.cinza, catAxisLabelFontSize: 12, valAxisLabelFontSize: 11,
      dataLabelColor: HEX.ink, dataLabelFontSize: 13, dataLabelFontBold: true,
      valGridLine: { color: HEX.grade, size: 0.75 }, catGridLine: { style: "none" }, catAxisLineShow: false, valAxisLineShow: false, showLegend: false,
      barDir: "col", barGapWidthPct: 55,
      chartColors: series.length === 1 ? [HEX.roxo] : series.length === 2 ? [HEX.lilas, HEX.roxo] : [HEX.lilas, HEX.roxo, HEX.magenta, HEX.roxoMedio, HEX.verde],
      showValue: o.barGrouping !== "stacked",
      dataLabelPosition: o.barGrouping === "stacked" ? undefined : "outEnd",
      dataLabelFormatCode: pct ? "0.0%" : "#,##0",
      valAxisLabelFormatCode: pct ? "0%" : "#,##0",
      objectName: nm("grafico"),
      ...o,
    });
  }

  // Cards com gráfico no lugar do número em texto (2–3 por slide). Cada item:
  //   { type: "columns", title, subtitle, icon, categories: [a, b], values: [va, vb], format: "pct" | "brl" | "int",
  //     badge: { value: "2,5×", label: "mais" } }          -> duas colunas; o selo fica sobre a coluna mais baixa
  //   { type: "donut", title, subtitle, icon, value: 0.55, display: "55,0%", centerLabel: "grupo B",
  //     compare: [{ label: "jan", value: 0.15, display: "15,0%" }, ...] } -> anel + barrinhas de comparação
  // Lilás = grupo de referência (primeira categoria), roxo = grupo em destaque. O card heroIndex sai em roxo sólido.
  async function chartCards(slide, items, { y = TOP, h = BOTTOM - TOP, heroIndex = 0 } = {}) {
    const k = items.length, cw = (CW - (k - 1) * GAP) / k;
    const FMT = { pct: "0.0%", brl: '"R$ "#,##0', int: "#,##0" };
    for (let i = 0; i < k; i++) {
      const it = items[i], x = M + i * (cw + GAP), hero = i === heroIndex;
      const fg = hero ? col.branco : col.texto, sub = hero ? col.lilas : col.cinza, accent = hero ? col.branco : col.roxo;
      const accentHex = hero ? HEX.white : HEX.roxo;
      card(slide, x, y, cw, h, hero ? col.roxo : col.fundo);
      if (it.icon) await iconDot(slide, x + 0.4, y + 0.4, 0.7, it.icon, hero ? col.roxoMedio : col.roxo);
      const hx = it.icon ? x + 1.3 : x + 0.4;
      text(slide, [{ text: it.title, options: { fontSize: 18, bold: true, color: fg, breakLine: !!it.subtitle } }, ...(it.subtitle ? [{ text: it.subtitle, options: { fontSize: 12, color: sub } }] : [])],
        { x: hx, y: y + 0.4, w: x + cw - 0.3 - hx, h: 0.7, valign: "middle" });

      if (it.type === "donut") {
        const d = Math.min(2.2, h - 2.25), dx = x + (cw - d) / 2, dy = y + 1.25;
        slide.addChart(pres.charts.DOUGHNUT, [{ name: it.title, labels: ["destaque", "restante"], values: [it.value, 1 - it.value] }], {
          x: dx, y: dy, w: d, h: d, holeSize: 68, chartColors: [accentHex, hero ? HEX.roxoMedio : HEX.lilas],
          showLegend: false, showValue: false, showPercent: false, showLabel: false, showTitle: false, dataBorder: { pt: 0, color: hero ? HEX.roxo : HEX.lilasBg },
          objectName: nm("grafico-anel"),
        });
        text(slide, [{ text: it.display, options: { fontSize: 24, bold: true, color: accent, breakLine: !!it.centerLabel } }, ...(it.centerLabel ? [{ text: it.centerLabel, options: { fontSize: 11, color: sub } }] : [])],
          { x: dx, y: dy + d / 2 - 0.45, w: d, h: 0.9, align: "center", valign: "middle" });
        const cmp = it.compare || [], max = Math.max(it.value, ...cmp.map((c) => c.value)), barMax = cw - 1.8;
        cmp.forEach((c, j) => {
          const ry = y + h - 0.3 - (cmp.length - j) * 0.36, bw = Math.max(0.08, barMax * c.value / max), last = j === cmp.length - 1;
          text(slide, c.label, { x: x + 0.4, y: ry, w: 0.45, h: 0.26, fontSize: 12, bold: true, color: fg, valign: "middle" });
          slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.86, y: ry + 0.03, w: bw, h: 0.2, fill: { color: last ? accent : (hero ? col.roxoMedio : col.lilas) }, line: { type: "none" }, rectRadius: 0.05, objectName: nm("barra-comparacao") });
          text(slide, c.display, { x: x + 0.96 + bw, y: ry, w: 0.9, h: 0.26, fontSize: 12, bold: true, color: accent, valign: "middle" });
        });
        continue;
      }

      // colunas
      const fmt = FMT[it.format || "int"], cx = x + 0.25, cyy = y + 1.3, chw = cw - 0.5, chh = h - 1.55;
      slide.addChart(pres.charts.BAR, [{ name: it.title, labels: it.categories, values: it.values }], {
        x: cx, y: cyy, w: chw, h: chh, barDir: "col", barGapWidthPct: 30,
        chartColors: [hero ? HEX.lilas : HEX.lilas, accentHex],
        showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: fmt, dataLabelFontSize: 18, dataLabelFontBold: true, dataLabelColor: accentHex,
        dataLabelFontFace: "+mn-lt", catAxisLabelFontFace: "+mn-lt",
        catAxisLabelColor: hero ? HEX.white : HEX.ink, catAxisLabelFontSize: 12, catAxisLabelFontBold: true, catAxisLineShow: false,
        valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: Math.max(...it.values) * 1.25, valGridLine: { style: "none" }, catGridLine: { style: "none" },
        showLegend: false, objectName: nm("grafico-colunas"),
      });
      if (it.badge) {
        const low = it.values[0] <= it.values[1] ? 0 : 1, colC = low === 0 ? cx + 0.9 * chw / 3.35 : cx + chw - 0.9 * chw / 3.35, bd = 0.9;
        slide.addShape(pres.shapes.OVAL, { x: colC - bd / 2, y: cyy - 0.1, w: bd, h: bd, fill: { color: col.magenta }, line: { type: "none" }, objectName: nm("selo") });
        text(slide, [{ text: it.badge.value, options: { fontSize: 16, bold: true, breakLine: !!it.badge.label } }, ...(it.badge.label ? [{ text: it.badge.label, options: { fontSize: 10 } }] : [])],
          { x: colC - bd / 2, y: cyy - 0.1, w: bd, h: bd, color: col.branco, align: "center", valign: "middle", wrap: false });
      }
    }
  }

  async function save(file) {
    await pres.writeFile({ fileName: file });
    await writeTheme(file);
    return file;
  }

  return {
    pres, HEX, col, M, CW, W, GAP, TOP, BOTTOM,
    cover, content, statement, closing,
    text, card, pill, chip, legend, note, iconDot, iconData,
    statCards, chartCards, compareBars, iconCards, callout, table, heroPanel, barChart,
    save,
  };
}

module.exports = { createDeck, HEX, THEME, ASSETS };
