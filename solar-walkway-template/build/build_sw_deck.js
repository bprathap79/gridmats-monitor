// GridMats Solar Walkway proposal template: deck generator (pptxgenjs, structured).
// Usage: node build_sw_deck.js <client.json> <outputs.json> <chartdata.json> <out.pptx>
const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const Fa = require("react-icons/fa");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");

const [CFGF, OUTF, CDF, OUT] = process.argv.slice(2);
const CFG = JSON.parse(fs.readFileSync(CFGF));
const O = JSON.parse(fs.readFileSync(OUTF));
const D = JSON.parse(fs.readFileSync(CDF));
const ASSETS = path.join(__dirname, "..", "assets");
const asset = (f) => path.join(ASSETS, f);

const THEME = {
  name: "GridMats Solar",
  headFontFace: "Arial",
  bodyFontFace: "Arial",
  colors: {
    dk1: "1A1A1A", lt1: "FFFFFF", dk2: "111A22", lt2: "F6F6F4",
    accent1: "C22935", accent2: "2E6B3A", accent3: "B87503", accent4: "6F6F6F",
    accent5: "FBEAEC", accent6: "0E2E4F", hlink: "C22935", folHlink: "A11F2A",
  },
};
const HX = THEME.colors;
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.theme = { headFontFace: "Arial", bodyFontFace: "Arial" };
pres.author = "GridMats (GroundReality Enterprises Pvt Ltd)";
pres.company = "GridMats";
pres.title = `GridMats Solar Walkway: pilot proposal for ${CFG.client_short}`;
pres.subject = "GridMats Solar Walkway pilot and rollout proposal";
const C = pres.SchemeColor;
const MX = 0.6, W = 13.333, CW = W - 2 * MX;

// ---------------------------------------------------------------- formatting
const inr = (n, d = 0) => Number(n).toLocaleString("en-IN", { minimumFractionDigits: d, maximumFractionDigits: d });
const lakh = (r, d = 2) => `₹${inr(r / 1e5, d)}\u00A0L`;
const pct = (f, d = 0) => `${(f * 100).toFixed(d)}%`;
const yrs = (n) => (typeof n === "number" ? `${n.toFixed(1)} years` : String(n));
const sqft = (perM2) => perM2 / 10.7639;
const client = CFG.client_short;

// ---------------------------------------------------------------- icons
const cache = {};
async function icon(name, color = "FFFFFF") {
  const k = name + color;
  if (cache[k]) return cache[k];
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Fa[name], { color: "#" + color, size: "256" }));
  return (cache[k] = "image/png;base64," + (await sharp(Buffer.from(svg)).png().toBuffer()).toString("base64"));
}

// ---------------------------------------------------------------- layouts
const logoBox = (x, y) => [
  { rect: { x, y, w: 2.05, h: 0.62, fill: { color: "FFFFFF" }, line: { color: "FFFFFF", width: 0 } } },
  { image: { path: asset("logo_gridmats_solar.png"), x: x + 0.1, y: y + 0.08, w: 1.85, h: 0.445 } },
];
pres.defineSlideMaster({
  title: "TITLE",
  background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: MX, y: 1.95, w: 5.6, h: 1.7, fontSize: 40, bold: true, color: C.background1, valign: "bottom", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: MX, y: 3.85, w: 5.5, h: 1.3, fontSize: 18, color: C.background2, valign: "top", align: "left", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: C.background2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: MX, y: 0.32, w: 10.3, h: 0.72, fontSize: 32, bold: true, color: C.text1, valign: "middle", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: MX, y: 1.04, w: 10.3, h: 0.45, fontSize: 15, color: C.accent4, valign: "top", align: "left", margin: 0 }, text: "" } },
    ...logoBox(11.05, 0.28),
    { text: { text: "GRIDMATS®  |  GroundReality Enterprises Pvt. Ltd.", options: { x: MX, y: 7.05, w: 6, h: 0.28, fontSize: 10, color: C.accent4, margin: 0 } } },
    { text: { text: "www.gridmats.com", options: { x: 9.6, y: 7.05, w: 2.6, h: 0.28, fontSize: 10, color: C.accent4, align: "right", margin: 0 } } },
  ],
  slideNumber: { x: 12.3, y: 7.05, w: 0.45, h: 0.28, fontSize: 10, color: C.accent4, align: "right" },
});
pres.defineSlideMaster({
  title: "CLOSING",
  background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: MX, y: 0.45, w: 10.0, h: 0.8, fontSize: 34, bold: true, color: C.background1, valign: "middle", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: MX, y: 1.25, w: 10.0, h: 0.45, fontSize: 15, color: C.background2, valign: "top", align: "left", margin: 0 }, text: "" } },
    ...logoBox(11.05, 0.28),
  ],
  slideNumber: { x: 12.3, y: 7.05, w: 0.45, h: 0.28, fontSize: 10, color: C.background2, align: "right" },
});

// ---------------------------------------------------------------- helpers
function content(section, title, lead) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
  s.addText(title, { placeholder: "title" });
  if (lead) s.addText(lead, { placeholder: "body" });
  return s;
}
function text(s, t, o) {
  s.addText(t, Object.assign({ isTextBox: true, margin: 0, fontSize: 14, color: C.text1, valign: "top" }, o));
}
function card(s, x, y, w, h, fill = C.background1, name = "card") {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: fill }, line: { color: fill === C.background1 ? "E3E3E3" : fill, width: 0.75 }, objectName: name });
}
function badge(s, label, x, y, d, fill) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill, width: 0 }, objectName: "badge" });
  text(s, label, { x, y, w: d, h: d, fontSize: d > 0.6 ? 16 : 13, bold: true, color: C.background1, align: "center", valign: "middle" });
}
async function iconBadge(s, name, x, y, d, fill) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill, width: 0 }, objectName: "icon-badge" });
  const ic = d * 0.5;
  s.addImage({ data: await icon(name), x: x + (d - ic) / 2, y: y + (d - ic) / 2, w: ic, h: ic, altText: name.replace(/^Fa/, "") });
}
function stat(s, x, y, w, value, label, color = C.accent1, size = 32) {
  text(s, value, { x, y, w, h: size / 58, fontSize: size, bold: true, color, valign: "bottom", objectName: "stat-value" });
  text(s, label, { x, y: y + size / 58 + 0.06, w, h: 0.6, fontSize: 13, color: C.text2, objectName: "stat-label" });
}
function source(s, t, y = 6.6) {
  text(s, t, { x: MX, y, w: CW, h: 0.32, fontSize: 10, italic: true, color: C.accent4, objectName: "source" });
}
const chartBase = () => ({
  catAxisLabelColor: "3F3F3F", valAxisLabelColor: "6F6F6F", catAxisLabelFontSize: 12, valAxisLabelFontSize: 11,
  catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", legendFontFace: "+mn-lt", titleFontFace: "+mn-lt",
  valGridLine: { color: "E3E3E3", size: 0.5 }, catGridLine: { style: "none" }, dataLabelFontSize: 11, dataLabelColor: "1A1A1A",
  legendFontSize: 11, legendColor: "3F3F3F", catAxisLineShow: true, valAxisLineShow: false, catAxisLineColor: "9A9A9A",
  showTitle: true, titleFontSize: 14, titleColor: "1A1A1A", plotArea: { fill: { color: "FFFFFF" } },
});
const tblOpts = (o) => Object.assign({ fontFace: "Arial", fontSize: 12, color: "1A1A1A", border: { type: "solid", pt: 0.5, color: "E3E3E3" }, fill: { color: "FFFFFF" }, valign: "middle" }, o);
const th = (t, fill = "111A22", align = "left") => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: fill }, align } });

async function main() {
  const area = O.PilotArea, kwp = O.PilotKWp, kwh = O.PilotKWh1;

  // ============================================================ 1. TITLE
  pres.addSection({ title: "Proposal" });
  let s = pres.addSlide({ masterName: "TITLE", sectionTitle: "Proposal" });
  s.addImage({ path: asset("hero_walkway.jpg"), x: 6.3, y: 0, w: 7.033, h: 7.5, sizing: { type: "cover", w: 7.033, h: 7.5 }, altText: "Concept visualization of a GridMats Solar Walkway through a campus" });
  s.addShape(pres.shapes.RECTANGLE, { x: 11.05, y: 0.28, w: 2.05, h: 0.62, fill: { color: "FFFFFF" }, line: { color: "FFFFFF", width: 0 }, objectName: "logo-box" });
  s.addImage({ path: asset("logo_gridmats_solar.png"), x: 11.15, y: 0.36, w: 1.85, h: 0.445, altText: "GridMats Solar logo" });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: MX, y: 1.2, w: 3.6, h: 0.42, rectRadius: 0.08, fill: { color: C.accent1 }, line: { color: C.accent1, width: 0 }, objectName: "pill" });
  text(s, `PILOT PROPOSAL  ·  ${client.toUpperCase()}`, { x: MX + 0.15, y: 1.2, w: 3.4, h: 0.42, fontSize: 12, bold: true, color: C.background1, valign: "middle" });
  s.addText("GridMats® Solar Walkway", { placeholder: "title" });
  s.addText(CFG.deck_tagline, { placeholder: "body" });
  text(s, `Prepared for ${CFG.client_entity}\n${CFG.site_name}, ${CFG.site_city}  ·  ${CFG.proposal_date}`, { x: MX, y: 5.35, w: 5.5, h: 0.75, fontSize: 13, color: C.background2 });
  text(s, "GroundReality Enterprises Pvt. Ltd., Bengaluru  ·  Indian Patent 398992", { x: MX, y: 6.75, w: 5.5, h: 0.3, fontSize: 11, color: C.accent4 });
  text(s, "Concept visualization", { x: 11.0, y: 7.1, w: 2.1, h: 0.25, fontSize: 10, italic: true, color: C.background1, align: "right" });
  s.addNotes(`Open with the idea: a ${inr(area)} m² walkway on the ${client} campus that people walk on and that generates power. It is a pilot: GridMats Solar Walkway is pre-commercial, and the pilot exists to measure output, heat, grip and upkeep on ${client}'s own site before any rollout.`);

  // ============================================================ 2. ONE PAGE
  s = content("Proposal", "The proposal on one page", `A ${inr(area)} m² generating walkway on the ${client} campus, measured for ${CFG.evaluation_months} months before any rollout decision.`);
  const pts = [
    ["What", `${inr(area)} m² GridMats Solar Walkway: ${inr(O.PilotModules)} walkable 40 W modules on a recycled-plastic GridMats base, ${kwp.toFixed(1)} kWp in all.`],
    ["Output", `About ${inr(Math.round(kwh / 10) * 10)} kWh a year (${O.PilotKWhDay.toFixed(1)} kWh a day on average), metered and sensed from day one.`],
    ["Cost", `${client} pays ${lakh(O.PilotPays)} ex-GST (₹${inr(O.PilotPaysSqft)}/sq ft). GridMats contributes ${lakh(O.PilotContrib)}: design, half the PV layer, supervision and the performance study.`],
    ["Decision", `${CFG.evaluation_months}-month evaluation against an agreed scorecard, monitoring for ${CFG.monitoring_months} months. At campus scale (${inr(O.RO_Area)} m²) the model shows a ${O.RO_Payback.toFixed(1)}-year payback.`],
  ];
  for (let i = 0; i < pts.length; i++) {
    const y = 1.8 + i * 1.2;
    badge(s, String(i + 1), MX, y, 0.62, [C.accent6, C.accent1, C.accent2, C.text2][i]);
    text(s, pts[i][0], { x: MX + 0.85, y: y - 0.02, w: 6.6, h: 0.34, fontSize: 16, bold: true });
    text(s, pts[i][1], { x: MX + 0.85, y: y + 0.34, w: 6.6, h: 0.8, fontSize: 14, color: C.text2 });
  }
  card(s, 8.35, 1.8, 4.38, 4.6, C.text2, "glance-card");
  text(s, "THE PILOT AT A GLANCE", { x: 8.65, y: 2.0, w: 3.9, h: 0.35, fontSize: 13, bold: true, color: "F2B33D" });
  const gl = [[`${kwp.toFixed(1)} kWp`, `on ${inr(area)} m² (${inr(O.PilotModules)} modules)`], [`${inr(Math.round(kwh / 10) * 10)} kWh`, "a year, estimated"], [lakh(O.PilotPays), `${client} pays, ex-GST`], [yrs(O.PE_Payback), `payback on ${client}'s outlay`]];
  gl.forEach((g, i) => {
    const x = 8.65 + (i % 2) * 2.0, y = 2.55 + Math.floor(i / 2) * 1.85;
    text(s, g[0], { x, y, w: 1.9, h: 0.55, fontSize: 24, bold: true, color: C.background1, valign: "bottom" });
    text(s, g[1], { x, y: y + 0.6, w: 1.85, h: 0.75, fontSize: 12, color: C.background2 });
  });
  source(s, "All figures from the companion model (base case). Money ex-GST. The GridMats Solar Walkway is pre-commercial; yield, life and upkeep are estimates the pilot measures.");
  s.addNotes(`Four lines and one ask. The ask: a ${inr(area)} m² pilot for which ${client} pays about ${lakh(O.PilotPays)} ex-GST, with GridMats carrying design, half the PV layer, supervision and a ${CFG.monitoring_months}-month performance study.`);

  // ============================================================ 3. WHY
  s = content("Proposal", `Why a walkway at ${client}`, "Every paved surface can do more: use it, build it from recycled plastic, and make power on it.");
  const icons = ["FaBuilding", "FaEye", "FaMicrochip", "FaShoePrints"];
  for (let i = 0; i < CFG.why_points.length && i < 4; i++) {
    const x = MX + (i % 2) * 6.15, y = 1.8 + Math.floor(i / 2) * 1.75;
    card(s, x, y, 5.95, 1.55, C.background1, "why-card");
    await iconBadge(s, icons[i] in Fa ? icons[i] : "FaCheckCircle", x + 0.25, y + 0.3, 0.7, [C.accent6, C.accent1, C.accent2, C.text2][i]);
    text(s, CFG.why_points[i][0], { x: x + 1.15, y: y + 0.22, w: 4.6, h: 0.36, fontSize: 16, bold: true });
    text(s, CFG.why_points[i][1], { x: x + 1.15, y: y + 0.62, w: 4.6, h: 0.85, fontSize: 13, color: C.text2 });
  }
  const chev = [["USE EXISTING SURFACES", "E9EEF4", C.accent6], ["RECYCLE INTO STRUCTURE", C.accent5, C.accent1], ["GENERATE CLEAN ENERGY", "E7F0E8", C.accent2]];
  chev.forEach(([t, f, tc], i) => {
    const x = MX + i * 4.1;
    s.addShape(pres.shapes.CHEVRON, { x, y: 5.5, w: 3.9, h: 0.75, fill: { color: f }, line: { color: f, width: 0 }, objectName: "chevron" });
    text(s, t, { x: x + 0.4, y: 5.5, w: 3.1, h: 0.75, fontSize: 13, bold: true, color: tc, align: "center", valign: "middle" });
  });
  s.addNotes("Why this site: the roof is already committed, walkways are open sky that is paved anyway, and a walkway is where people see it. The three chevrons are the GridMats Solar Walkway proposition.");

  // ============================================================ 4. WHAT IT IS
  pres.addSection({ title: "Technology" });
  s = content("Technology", "What it is", "A modular surface that carries people and produces electricity.");
  s.addImage({ path: asset("exploded_layers.jpg"), x: MX, y: 1.75, w: 6.3, h: 3.55, sizing: { type: "cover", w: 6.3, h: 3.55 }, altText: "Exploded view: PV module, MLP lumber, GridMats base on aggregate" });
  text(s, "Concept visualization", { x: MX, y: 5.35, w: 6.3, h: 0.25, fontSize: 10, italic: true, color: C.accent4, align: "right" });
  const layers = [["PV", C.accent6, "1. Walkable photovoltaic layer", "12 mm toughened glass over a 40 W module; modular, so one module can be replaced without lifting the walkway."],
    ["DC", "B87503", "2. MLP lumber electrical layer", "40 × 40 mm recycled MLP lumber carries the wiring, with about 50 mm open gaps for airflow and drainage."],
    ["GM", C.accent1, "3. Patented GridMats structural base", "100% recycled polypropylene grid (500 × 500 mm, 4 per m²) gives confinement, stability and a level bed."]];
  for (let i = 0; i < layers.length; i++) {
    const y = 1.75 + i * 1.22;
    card(s, 7.2, y, 5.53, 1.08, C.background1, "layer-card");
    badge(s, layers[i][0], 7.4, y + 0.24, 0.6, layers[i][1]);
    text(s, layers[i][2], { x: 8.2, y: y + 0.12, w: 4.4, h: 0.32, fontSize: 15, bold: true });
    text(s, layers[i][3], { x: 8.2, y: y + 0.45, w: 4.4, h: 0.6, fontSize: 12, color: C.text2 });
  }
  card(s, MX, 5.75, CW, 0.7, "E7F0E8", "basis-card");
  text(s, `Design basis:  40 W per 550 × 400 mm module  ·  4 modules per m²  ·  about ${inr(O.PowerDens)} W per m² of walkway  ·  15-16% surface efficiency  ·  no extra land or mounting`,
    { x: MX + 0.25, y: 5.75, w: CW - 0.5, h: 0.7, fontSize: 14, bold: true, color: C.text1, valign: "middle" });
  source(s, "GridMats Solar Walkway flier (6 Oct 2026) and Client/Partner Deck v2. Status: pre-commercial; pilot validation under way.");
  s.addNotes("Three layers. Walkable PV on top, MLP lumber that carries the wiring and lets water and air through, and the patented GridMats base. Design basis about 160 W per square metre of walkway.");

  // ============================================================ 5. HOW IT WORKS + SAFETY
  s = content("Technology", "How it works, and how it stays safe", "A proven GridMats platform extended into distributed energy; safety is designed in and verified in the pilot.");
  const steps = [["Capture", "Walkable PV modules harvest sunlight on a surface already in use."], ["Connect", "Low-voltage DC wiring runs inside the MLP lumber, out of sight and out of the way."],
    ["Convert", "An inverter feeds the campus low-voltage board, behind the meter."], ["Use and measure", "Every kWh is metered; sensors log module heat, moisture and uptime."]];
  steps.forEach((st, i) => {
    const y = 1.8 + i * 1.12;
    badge(s, `0${i + 1}`, MX, y, 0.6, C.accent1);
    text(s, st[0], { x: MX + 0.85, y: y - 0.02, w: 5.2, h: 0.32, fontSize: 16, bold: true });
    text(s, st[1], { x: MX + 0.85, y: y + 0.32, w: 5.2, h: 0.6, fontSize: 13, color: C.text2 });
  });
  card(s, 6.95, 1.8, 5.78, 4.55, C.text2, "safety-card");
  await iconBadge(s, "FaShieldAlt", 7.2, 2.0, 0.6, C.accent1);
  text(s, "Safety by design", { x: 7.95, y: 2.05, w: 4.5, h: 0.5, fontSize: 18, bold: true, color: C.background1, valign: "middle" });
  text(s, ["Textured toughened glass: target R11 or wet PTV ≥ 36", "Load target 4 kN/m² and 4.5 kN point load (IS 875 Part 2)", "Extra-low-voltage DC strings; IP68 junction boxes at grade", "BIS-registered modules (IS 14286, IS/IEC 61730)", "50 mm gaps drain and cool; edge ramps; one-module repairs"].map((t, j, a) => ({ text: t, options: { bullet: true, breakLine: j < a.length - 1 } })),
    { x: 7.2, y: 2.8, w: 5.3, h: 2.45, fontSize: 13, color: C.background1, paraSpaceAfter: 6 });
  text(s, "Tested at handover and through the pilot; no result is claimed until it is measured.", { x: 7.2, y: 5.2, w: 5.3, h: 0.8, fontSize: 13, italic: true, color: "F2B33D" });
  s.addNotes("Walk through the four steps, then safety. Be clear that walkability, load and heat are pilot measurements; GridMats does not claim test results it does not yet have.");


  // ============================================================ 5b. TOPCON
  s = content("Technology", "Built on the latest TOPCon cells", "n-type TOPCon runs cooler and ages slower than PERC, which matters on a hot, flat walkway.");
  const tstats = [["25.6-26%", "TOPCon cell efficiency in mass production (2026)", C.accent6], ["26.66%", "record industrial TOPCon cell, M10 wafer (Feb 2026)", C.accent1],
    ["−0.29%/°C", "power loss per degree, against −0.35%/°C for PERC", "B87503"], ["1% + 0.4%/yr", "degradation: first year, then each year", C.accent2]];
  tstats.forEach((t2, i) => {
    const x = MX + (i % 2) * 3.0, y = 1.75 + Math.floor(i / 2) * 1.6;
    card(s, x, y, 2.85, 1.45, C.background1, "tech-card");
    stat(s, x + 0.2, y + 0.15, 2.5, t2[0], t2[1], t2[2], 24);
  });
  const tt = [[th("Same 4 kWp pilot"), th("Heat loss", "111A22", "center"), th("20-yr MWh", "111A22", "center"), th("W/m² (derived)", "111A22", "center")]];
  ["Mono PERC", "TOPCon", "HJT"].forEach((nm, i) => tt.push([{ text: nm, options: { bold: true, color: i === 1 ? "C22935" : "1A1A1A" } },
    { text: pct(O[`TC_TL_${i}`], 1), options: { align: "center" } }, { text: (O[`TC_Y20_${i}`] / 1000).toFixed(1), options: { align: "center", bold: i === 1 } }, { text: inr(O[`TC_WM2_${i}`]), options: { align: "center" } }]));
  s.addTable(tt, tblOpts({ x: MX, y: 5.0, w: 5.85, colW: [1.75, 1.3, 1.3, 1.5], rowH: 0.34, fontSize: 11 }));
  const losses = [["Heat", O.LossTemp], ["Dust", O.LossSoil], ["Footfall shade", O.LossFoot], ["Reflection", O.LossRefl], ["Wiring, mismatch", O.LossMism], ["Inverter", O.LossInv], ["Availability", O.LossAvail]];
  s.addChart(pres.charts.BAR, [{ name: "Loss", labels: losses.map((l) => l[0]), values: losses.map((l) => l[1] * 100) }], Object.assign(chartBase(), {
    x: 6.85, y: 1.75, w: 5.88, h: 4.0, barDir: "bar", chartColors: [HX.accent1], showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0\"%\"",
    valAxisHidden: true, valGridLine: { style: "none" }, showLegend: false, catAxisOrientation: "maxMin", valAxisMinVal: 0, valAxisMaxVal: 12,
    title: `Where the energy goes: base-case losses (performance ratio ${pct(O.PR)})`,
  }));
  text(s, `TOPCon generates ${pct(O.TC_Y20_1 / O.TC_Y20_0 - 1, 1)} more than PERC over 20 years on this walkway. Rated on the 40 W design basis; a TOPCon-derived rating of about ${inr(O.TC_WM2_1)} W/m² is the upgrade path, to be validated.`,
    { x: 6.85, y: 5.85, w: 5.88, h: 0.75, fontSize: 12, color: C.text2 });
  source(s, "pv-magazine (25 Feb 2026); taiyangnews; Canadian Solar TOPHiKu6 and Jinko Tiger Neo datasheets; JMK Research (Apr 2026). Losses are model assumptions the pilot measures.", 6.68);
  s.addNotes("Cell technology matters more on a walkway than on a roof: the modules lie flat, run hot and collect dust. TOPCon's lower temperature coefficient and slower degradation are worth a few percent of lifetime energy over PERC; HJT, which Reliance makes at Jamnagar, is slightly better again and is a candidate for a co-developed module.");

  // ============================================================ 6. PILOT LAYOUT
  pres.addSection({ title: "Pilot" });
  s = content("Pilot", `The pilot on the ${client} campus`, `${CFG.site_name}: ${inr(CFG.pilot_length_m, 1)} m × ${inr(CFG.pilot_width_m, 1)} m of walkway, ${inr(O.PilotModules)} modules, behind the meter.`);
  const L = CFG.pilot_length_m, Wd = CFG.pilot_width_m;
  const scale = Math.min(7.2 / L, 2.0 / Wd);
  const wx = MX + 0.2, wy = 2.4, ww = L * scale, wh = Wd * scale;
  s.addShape(pres.shapes.RECTANGLE, { x: wx - 0.08, y: wy - 0.08, w: ww + 0.16, h: wh + 0.16, fill: { color: "D9D4C7" }, line: { color: "B8B1A3", width: 0.75 }, objectName: "walkway-edge" });
  const cols = Math.max(1, Math.round(L / 0.5)), rows = Math.max(1, Math.round(Wd / 0.5));
  const cw = ww / cols, chh = wh / rows;
  for (let i = 0; i < cols; i++) for (let j = 0; j < rows; j++)
    s.addShape(pres.shapes.RECTANGLE, { x: wx + i * cw + cw * 0.06, y: wy + j * chh + chh * 0.06, w: cw * 0.88, h: chh * 0.88, fill: { color: "1F3354" }, line: { color: "6E7FA0", width: 0.5 }, objectName: "module" });
  text(s, `${inr(L, 1)} m`, { x: wx, y: wy + wh + 0.15, w: ww, h: 0.3, fontSize: 12, color: C.text2, align: "center" });
  text(s, `${inr(Wd, 1)} m`, { x: wx + ww + 0.12, y: wy + wh / 2 - 0.15, w: 0.7, h: 0.3, fontSize: 12, color: C.text2 });
  // sensors
  [[0.15, 0.3], [0.45, 0.7], [0.75, 0.3], [0.95, 0.7]].forEach(([fx, fy]) => s.addShape(pres.shapes.OVAL, { x: wx + ww * fx - 0.08, y: wy + wh * fy - 0.08, w: 0.16, h: 0.16, fill: { color: C.accent1 }, line: { color: "FFFFFF", width: 1 }, objectName: "sensor" }));
  // inverter and board
  const ix = wx + ww - 1.9, iy = wy + wh + 0.75;
  s.addShape(pres.shapes.LINE, { x: wx + ww, y: wy + wh, w: 0.0, h: 0.95, line: { color: "B87503", width: 2, dashType: "dash" }, objectName: "cable" });
  s.addShape(pres.shapes.LINE, { x: ix + 1.45, y: wy + wh + 0.95, w: wx + ww - ix - 1.45, h: 0, line: { color: "B87503", width: 2, dashType: "dash" }, objectName: "cable" });
  card(s, ix, iy, 1.45, 0.62, C.accent6, "inverter");
  text(s, "Inverter +\nenergy meter", { x: ix, y: iy, w: 1.45, h: 0.62, fontSize: 11, bold: true, color: C.background1, align: "center", valign: "middle" });
  card(s, ix - 1.9, iy, 1.55, 0.62, C.text2, "board");
  text(s, "Campus LV board\n(≤ 30 m)", { x: ix - 1.9, y: iy, w: 1.55, h: 0.62, fontSize: 11, bold: true, color: C.background1, align: "center", valign: "middle" });
  s.addShape(pres.shapes.LINE, { x: ix - 0.35, y: iy + 0.31, w: 0.35, h: 0, line: { color: "B87503", width: 2, dashType: "dash" }, objectName: "cable" });
  // legend
  const ly = 5.25;
  s.addShape(pres.shapes.RECTANGLE, { x: MX + 0.2, y: ly, w: 0.25, h: 0.25, fill: { color: "1F3354" }, line: { color: "6E7FA0", width: 0.5 }, objectName: "legend" });
  text(s, "40 W module on a 500 × 500 mm GridMats grid", { x: MX + 0.55, y: ly - 0.02, w: 4.0, h: 0.3, fontSize: 12, color: C.text2 });
  s.addShape(pres.shapes.OVAL, { x: MX + 0.24, y: ly + 0.45, w: 0.18, h: 0.18, fill: { color: C.accent1 }, line: { color: "FFFFFF", width: 1 }, objectName: "legend" });
  text(s, "Sensors: module temperature, moisture, reference cell", { x: MX + 0.55, y: ly + 0.4, w: 4.6, h: 0.3, fontSize: 12, color: C.text2 });
  s.addShape(pres.shapes.LINE, { x: MX + 0.15, y: ly + 0.97, w: 0.35, h: 0, line: { color: "B87503", width: 2, dashType: "dash" }, objectName: "legend" });
  text(s, "Low-voltage cable route", { x: MX + 0.55, y: ly + 0.82, w: 4.0, h: 0.3, fontSize: 12, color: C.text2 });
  card(s, 8.6, 1.8, 4.13, 4.6, C.background1, "spec-card");
  text(s, "Pilot specification", { x: 8.85, y: 1.95, w: 3.7, h: 0.35, fontSize: 16, bold: true });
  const spec = [["Area", `${inr(area, 1)} m² (${inr(area * 10.7639)} sq ft)`], ["Modules", `${inr(O.PilotModules)} × ${inr(O.ModW)} W`], ["Cells", O.TechSel], ["Rated power", `${kwp.toFixed(1)} kWp`], ["Build", CFG.build_type], ["Connection", "Behind the meter, campus LV board"], ["Monitoring", `${CFG.monitoring_months} months, live dashboard`], ["Install time", "2-3 days on site"]];
  spec.forEach(([k, v], i) => {
    text(s, k, { x: 8.85, y: 2.4 + i * 0.48, w: 1.4, h: 0.42, fontSize: 13, bold: true, color: C.text2, valign: "middle" });
    text(s, v, { x: 10.25, y: 2.4 + i * 0.48, w: 2.4, h: 0.42, fontSize: 13, valign: "middle" });
  });
  source(s, "Indicative layout; final layout follows the site assessment (shade, footfall, drainage, connection point).");
  s.addNotes("This is the pilot as drawn: four modules across, one per 500 mm grid, the cable run to an inverter and meter, then to the nearest low-voltage board. Sensor positions are indicative.");

  // ============================================================ 7. ENERGY
  s = content("Pilot", "What it will generate", `${CFG.irradiance_city} sunshine, ${O.TechSel} modules, ${pct(O.PR)} performance ratio after losses, ${pct(CFG.shading_loss)} site shading.`);
  s.addChart(pres.charts.BAR, [{ name: "kWh", labels: D.months, values: D.monthly_kwh }], Object.assign(chartBase(), {
    x: MX, y: 1.75, w: 7.6, h: 4.65, barDir: "col", chartColors: [HX.accent3], showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "#,##0",
    valAxisLabelFormatCode: "#,##0", valAxisMinVal: 0, showLegend: false, barGapWidthPct: 45, title: `Pilot generation by month (kWh), ${kwp.toFixed(1)} kWp`,
  }));
  const en = [[`${inr(Math.round(kwh / 10) * 10)} kWh`, "a year, estimated", C.accent3], [`${O.PilotKWhDay.toFixed(1)} kWh`, "a day on average", C.accent6], [`${inr(O.BestMonth)} / ${inr(O.WorstMonth)}`, "kWh in the best and the monsoon month", C.text2], [`${(O.PilotCO2kg / 1000).toFixed(1)} t`, "CO2 avoided a year (CEA grid factor)", C.accent2]];
  en.forEach((e, i) => stat(s, 8.75, 1.85 + i * 1.15, 4.0, e[0], e[1], e[2], 26));
  source(s, `Sunshine: typical-year ${CFG.irradiance_city} profile, ${inr(O.SunYear)} kWh/m²/yr (replace with site data after the site walk). Specific yield ${inr(O.SpecYield)} kWh/kWp/yr.`);
  s.addNotes("Generation follows the sun: strongest in March to May, weakest in the monsoon. The pilot will measure the real monthly figures against this curve; that comparison is the first line of the scorecard.");

  // ============================================================ 8. COST & OFFER
  s = content("Pilot", "What the pilot costs, and who pays", `List value of the pilot scope against what ${client} pays. Ex-GST; GST extra at the applicable rate.`);
  const b = D.boq;
  const grp = (idx) => idx.reduce((acc, i) => ({ list: acc.list + b[i].list, contrib: acc.contrib + b[i].contrib, pays: acc.pays + b[i].pays }), { list: 0, contrib: 0, pays: 0 });
  const lines = [["Site assessment and engineering design", grp([0])], ["Base works and GridMats OC-LW 40 base", grp([1, 2])], [`Walkable PV layer (${inr(O.PilotModules)} × 40 W, MLP lumber, wiring)`, grp([3, 4])],
    ["Electrical BOS, grid integration and metering", grp([5, 7])], ["Logistics and installation", grp([6, 8])], ["GridMats supervision, training and travel", grp([9, 10])], ["IoT sensors, dashboard and performance study", grp([11])]];
  const rowsT = [[th("Scope"), th("List", "111A22", "right"), th("GridMats covers", "2E6B3A", "right"), th(`${client} pays`, "C22935", "right")]];
  lines.forEach(([l, v]) => rowsT.push([{ text: l }, { text: lakh(v.list), options: { align: "right" } }, { text: v.contrib ? lakh(v.contrib) : "-", options: { align: "right", color: "2E6B3A" } }, { text: v.pays ? lakh(v.pays) : "Included", options: { align: "right", bold: !v.pays, color: v.pays ? "1A1A1A" : "2E6B3A" } }]));
  rowsT.push([{ text: "Total, ex-GST", options: { bold: true } }, { text: lakh(O.PilotList), options: { align: "right", bold: true } }, { text: lakh(O.PilotContrib), options: { align: "right", bold: true, color: "2E6B3A" } }, { text: lakh(O.PilotPays), options: { align: "right", bold: true, color: "C22935" } }]);
  s.addTable(rowsT, tblOpts({ x: MX, y: 1.75, w: 8.0, colW: [4.1, 1.25, 1.35, 1.3], rowH: 0.47 }));
  card(s, 8.95, 1.75, 3.78, 4.65, C.background1, "pay-card");
  stat(s, 9.2, 1.95, 3.3, lakh(O.PilotPays), `${client} pays, ex-GST: ₹${inr(O.PilotPaysSqft)}/sq ft (${lakh(O.PilotPaysGST)} with GST)`, C.accent1, 30);
  stat(s, 9.2, 3.45, 3.3, lakh(O.PilotContrib), `GridMats contributes (${pct(O.PilotContribPct)} of list): design, half the PV layer, supervision, IoT study`, C.accent2, 30);
  text(s, `Defect liability: 12 months from handover, with joint inspections. PV layer price ₹${inr(O.PVPilotSqft)}/sq ft (₹9,500/m²).`, { x: 9.2, y: 5.15, w: 3.3, h: 1.1, fontSize: 12, color: C.text2 });
  source(s, `Co-funding of the PV layer: ${pct(CFG.cofund)} (pilot offer, needs MD approval). Full physical installed cost before co-funding: ${lakh(O.PilotPhys)} (₹${inr(O.PilotPhysSqft)}/sq ft).`);
  s.addNotes(`The offer: ${client} pays for the base, half the PV layer, the electrical work and installation; GridMats carries engineering, the other half of the PV layer, supervision and the monitoring study. Prices shown per sq ft first.`);

  // ============================================================ 9. ECONOMICS
  pres.addSection({ title: "Economics" });
  s = content("Economics", "Pilot economics, and what rollout looks like", `The pilot is for evidence; the business case is the campus rollout at ${inr(O.RO_Area)} m².`);
  card(s, MX, 1.75, 3.85, 4.65, C.background1, "pilot-econ");
  text(s, "PILOT", { x: MX + 0.25, y: 1.9, w: 3.4, h: 0.3, fontSize: 13, bold: true, color: C.accent1 });
  stat(s, MX + 0.25, 2.3, 3.4, yrs(O.PE_Payback), `payback on ${client}'s ${lakh(O.PilotPays)} outlay`, C.accent1, 28);
  stat(s, MX + 0.25, 3.65, 3.4, yrs(O.PF_Payback), `payback on the full ${lakh(O.PilotPhys)} installed cost, without co-funding`, C.text2, 28);
  text(s, `Year-1 bill saving about ₹${inr(Math.round(O.PE_Save1 / 100) * 100)} at ₹${CFG.tariff.toFixed(2)}/kWh.`, { x: MX + 0.25, y: 5.2, w: 3.4, h: 0.9, fontSize: 12, color: C.text2 });
  s.addChart(pres.charts.LINE, [{ name: "Cumulative net cash (₹ lakh)", labels: D.rollout_cum.map((_, i) => String(i)), values: D.rollout_cum.map((v) => v / 1e5) }], Object.assign(chartBase(), {
    x: 4.65, y: 1.75, w: 5.25, h: 4.65, chartColors: [HX.accent2], lineSize: 3, lineDataSymbol: "circle", lineDataSymbolSize: 5, showLegend: false,
    valAxisLabelFormatCode: "#,##0", catAxisLabelFontSize: 10, title: `Campus rollout: cumulative net cash (₹ lakh), ${inr(O.RO_Area)} m²`,
  }));
  const ro = [[`₹${inr(O.RO_CostSqft)}/sq ft`, `installed at rollout (₹${inr(O.RO_CostM2)}/m²)`, C.text1], [yrs(O.RO_Payback), "payback, net of the walkway it replaces", C.accent2], [pct(O.RO_IRR), "IRR over 20 years, pre-tax", C.accent2], [`₹${O.RO_LCOE.toFixed(2)}`, `per kWh, against grid power at ₹${CFG.tariff.toFixed(2)}`, C.accent6]];
  ro.forEach((r, i) => stat(s, 10.15, 1.8 + i * 1.15, 2.6, r[0], r[1], r[2], 22));
  source(s, `Rollout uses the rollout PV-layer price (₹8,500/m²), the GridMats price ladder and a ₹1,200/m² credit for the paver walkway it replaces. ${pct(CFG.tariff_escalation)} tariff escalation, ${pct(CFG.discount_rate)} discount rate.`);
  s.addNotes("Separate the two cases. The pilot is an evidence purchase, made cheap by co-funding. The business case is rollout: at campus scale, with rollout prices and a credit for the walkway that would be built anyway, payback is a little under five years.");

  // ============================================================ 10. COMPARISON
  s = content("Economics", "How it compares", "Rooftop solar is cheaper per kWp. The walkway wins where roofs are full and a walkway is being built anyway.");
  const cmp = D.comparison;
  s.addChart(pres.charts.BAR, [{ name: "₹ per kWp", labels: ["C&I rooftop", "Solar carport", "GridMats Solar Walkway"], values: [cmp["Capex per kWp (₹)"][1], cmp["Capex per kWp (₹)"][2], cmp["Capex per kWp (₹)"][0]] }],
    Object.assign(chartBase(), { x: MX, y: 1.75, w: 5.7, h: 4.65, barDir: "bar", chartColors: [HX.accent6], showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "₹#,##0",
      valAxisHidden: true, valGridLine: { style: "none" }, showLegend: false, catAxisOrientation: "maxMin", valAxisMinVal: 0, valAxisMaxVal: 95000, title: "Installed cost per kWp" }));
  const ct = [[th("On the same footprint"), th("Walkway", "C22935", "center"), th("Rooftop", "111A22", "center"), th("Carport", "111A22", "center"), th("Pavers + rooftop", "111A22", "center")]];
  const crow = (lab, key, f) => [{ text: lab, options: { bold: true } }].concat(cmp[key].map((v) => ({ text: f(v), options: { align: "center" } })));
  ct.push(crow("Payback (years)", "Simple payback on year-1 figures (years)", (v) => Number(v).toFixed(1)));
  ct.push(crow("Walking surface", "Provides a walking surface", (v) => String(v)));
  ct.push(crow("Roof or structure", "Needs roof or extra structure", (v) => String(v)));
  ct.push(crow("Recycled plastic (t)", "Recycled plastic in the structure (t)", (v) => (v ? Number(v).toFixed(1) : "-")));
  s.addTable(ct, tblOpts({ x: 6.6, y: 1.75, w: 6.13, colW: [1.6, 1.1, 1.05, 1.18, 1.2], rowH: 0.62, fontSize: 11 }));
  text(s, "Positioning: on walkways that are being built or relaid, on campuses where the roof is spoken for, and where visibility matters.", { x: 6.6, y: 5.2, w: 6.13, h: 0.9, fontSize: 14, bold: true, color: C.accent1 });
  source(s, `Footprint ${inr(O.RO_Area)} m². Rooftop ₹40/Wp and carport ₹61/Wp all-in (earlier GridMats ROI model benchmarks). Payback on year-1 figures, before any walkway credit.`);
  s.addNotes("Be honest here: per kWp, a roof is cheaper and pays back faster. The walkway is not a replacement for rooftop solar; it adds generation where the roof is full and turns a walkway budget into a generating asset.");

  // ============================================================ 11. SENSITIVITY
  s = content("Economics", "What moves the payback", `Campus-rollout payback in years: power price against performance ratio. Base case: ₹${CFG.tariff.toFixed(2)}/kWh, ${pct(O.PR)}.`);
  const sp = D.sens_pr;
  const cellFill = (v) => (typeof v !== "number" ? "FFFFFF" : v <= 6 ? "E3F1E6" : v <= 9 ? "FFF4D6" : "FBEAEC");
  const st = [[th("₹ / kWh  ↓   ·   PR  →")].concat(sp.cols.map((c) => th(pct(c), "111A22", "center")))];
  sp.rows.forEach((r) => {
    const isBaseRow = Math.abs(r[0] - CFG.tariff) < 0.01;
    st.push([{ text: `₹${Number(r[0]).toFixed(0)}`, options: { bold: true, align: "center" } }].concat(r.slice(1).map((v, j) => ({ text: typeof v === "number" ? v.toFixed(1) : String(v),
      options: { align: "center", fill: { color: cellFill(v) }, bold: isBaseRow && Math.abs(sp.cols[j] - O.PR) < 0.001 } }))));
  });
  s.addTable(st, tblOpts({ x: MX, y: 1.75, w: 7.4, colW: [2.0, 1.08, 1.08, 1.08, 1.08, 1.08], rowH: 0.52, fontSize: 13 }));
  card(s, 8.4, 1.75, 4.33, 4.65, C.background1, "sens-card");
  text(s, "Reading the table", { x: 8.65, y: 1.9, w: 3.9, h: 0.35, fontSize: 16, bold: true });
  text(s, [{ text: "Green: ", options: { bold: true, color: "2E6B3A" } }, { text: "6 years or less", options: { breakLine: true } }, { text: "Amber: ", options: { bold: true, color: "B87503" } }, { text: "6 to 9 years", options: { breakLine: true } }, { text: "Red: ", options: { bold: true, color: "C22935" } }, { text: "above 9 years" }],
    { x: 8.65, y: 2.35, w: 3.9, h: 1.0, fontSize: 13 });
  const cf = D.sens_cofund;
  text(s, `Pilot payback on ${client}'s outlay vs GridMats co-funding of the PV layer`, { x: 8.65, y: 3.55, w: 3.9, h: 0.6, fontSize: 13, bold: true });
  s.addTable([[th("Co-funding")].concat(cf.cols.map((c) => th(pct(c), "111A22", "center"))), [{ text: "Payback (yrs)", options: { bold: true } }].concat(cf.vals.map((v) => ({ text: Number(v).toFixed(1), options: { align: "center", fill: { color: cellFill(v) } } })))],
    tblOpts({ x: 8.65, y: 4.25, w: 3.85, colW: [1.25, 0.65, 0.65, 0.65, 0.65], rowH: 0.45, fontSize: 11 }));
  text(s, "The pilot measures the performance ratio; the electricity bill sets the power price.", { x: 8.65, y: 5.35, w: 3.9, h: 0.9, fontSize: 12, italic: true, color: C.text2 });
  source(s, "Model sheet 'Sensitivity'. Simple payback on year-1 figures, net of the walkway it replaces.");
  s.addNotes("Two numbers drive everything: the landed power price at the campus and the real performance ratio of flat, walk-on modules. The bill gives the first; the pilot gives the second.");

  // ============================================================ 12. MEASURED
  pres.addSection({ title: "Plan" });
  s = content("Plan", "Measured, not assumed", `A scorecard agreed with ${client} before installation and reviewed at ${CFG.evaluation_months} months.`);
  const kt = [[th("KPI"), th("Target"), th("Weight", "111A22", "center")]];
  D.kpis.forEach((k) => kt.push([{ text: k.kpi, options: { bold: true } }, { text: k.target }, { text: pct(k.weight), options: { align: "center" } }]));
  s.addTable(kt, tblOpts({ x: MX, y: 1.72, w: 7.6, colW: [3.4, 3.3, 0.9], rowH: 0.42, fontSize: 11 }));
  card(s, 8.55, 1.72, 4.18, 4.7, C.text2, "dash-card");
  await iconBadge(s, "FaMicrochip", 8.8, 1.92, 0.6, C.accent1);
  text(s, "Live pilot dashboard", { x: 9.55, y: 1.97, w: 3.0, h: 0.5, fontSize: 17, bold: true, color: C.background1, valign: "middle" });
  text(s, ["Energy: kWh per day and W per m²", "Module temperature and ventilation", "Moisture, rainfall and drainage", "Uptime and fault alerts", "Data exportable to the campus IoT stack"].map((t, j, a) => ({ text: t, options: { bullet: true, breakLine: j < a.length - 1 } })),
    { x: 8.8, y: 2.75, w: 3.75, h: 2.4, fontSize: 13, color: C.background1, paraSpaceAfter: 6 });
  text(s, "GridMats Pavement Monitor, extended for the Solar Walkway.", { x: 8.8, y: 5.4, w: 3.75, h: 0.7, fontSize: 12, italic: true, color: "F2B33D" });
  source(s, "KPIs and weights from the earlier GridMats Solar Walkway ROI model (pilot scorecard). Targets are finalised jointly before installation.");
  s.addNotes("The scorecard is the decision tool: ten measures, weighted, agreed before a single module goes down. The dashboard makes the data visible to the client's team throughout.");

  // ============================================================ 13. TIMELINE
  s = content("Plan", "From site walk to switch-on", `About six weeks to switch-on, a ${CFG.evaluation_months}-month evaluation, then monitoring through the monsoon.`);
  const M = Math.max(12, CFG.monitoring_months + 2);
  const gx = MX + 3.9, gw = CW - 3.9, mw = gw / M;
  for (let m = 0; m < M; m++) text(s, `M${m + 1}`, { x: gx + m * mw, y: 1.75, w: mw, h: 0.3, fontSize: 11, color: C.accent4, align: "center" });
  const tasks = [["Site walk, shading and generation estimate", 0, 0.4, C.accent6], ["Engineering, electrical and safety approvals", 0.3, 1.0, C.accent6], ["Order; supply date fixed in the PO", 0.9, 1.4, C.accent4],
    ["Install (2-3 days) and commission", 1.4, 1.6, C.accent1], [`Evaluation against the scorecard (${CFG.evaluation_months} months)`, 1.6, 1.6 + CFG.evaluation_months, C.accent2],
    ["Decision gate: rollout or refine", 1.6 + CFG.evaluation_months, 1.8 + CFG.evaluation_months, C.accent1], [`Monitoring, live dashboard (${CFG.monitoring_months} months)`, 1.6, Math.min(M, 1.6 + CFG.monitoring_months), "B87503"]];
  tasks.forEach(([t, a, bb, col], i) => {
    const y = 2.15 + i * 0.52;
    text(s, t, { x: MX, y, w: 3.8, h: 0.4, fontSize: 13, valign: "middle" });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: gx + a * mw, y: y + 0.05, w: Math.max((bb - a) * mw, 0.15), h: 0.3, rectRadius: 0.06, fill: { color: col }, line: { color: col, width: 0 }, objectName: "gantt" });
  });
  card(s, MX, 6.05, CW, 0.5, C.background1, "promise-card");
  text(s, "Built by teams GridMats trains, with a GridMats engineer on site for the whole installation.", { x: MX + 0.25, y: 6.05, w: CW - 0.5, h: 0.5, fontSize: 13, bold: true, valign: "middle" });
  s.addNotes("Indicative timing. The supply date is fixed in the purchase order once the site is ready; GridMats does not promise a fixed number of days before that.");

  // ============================================================ 14. SCALE
  s = content("Plan", "From pilot to campus", "Year-1 figures at each step; later phases use rollout prices for their area.");
  const sc = D.scale;
  const ph = [["Pilot", 0, C.accent1], ["Phase 2", 1, C.accent6], ["Campus", 2, C.accent2], ["Portfolio", 3, C.text2]];
  ph.forEach(([nm, j, col], i) => {
    const x = MX + i * 3.08, w = 2.9;
    card(s, x, 1.75, w, 4.0, C.background1, "phase-card");
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.75, w, h: 0.75, rectRadius: 0.08, fill: { color: col }, line: { color: col, width: 0 }, objectName: "phase-head" });
    text(s, nm, { x: x + 0.2, y: 1.75, w: w - 0.4, h: 0.45, fontSize: 15, bold: true, color: C.background1, valign: "middle" });
    text(s, `${inr(sc["Area (m²)"][j])} m²`, { x: x + 0.2, y: 2.12, w: w - 0.4, h: 0.32, fontSize: 12, color: C.background1 });
    const kv = [[`${inr(sc["Rated power (kWp)"][j], sc["Rated power (kWp)"][j] < 10 ? 1 : 0)} kWp`, "rated power"], [`${inr(sc["Generation, year 1 (MWh)"][j], sc["Generation, year 1 (MWh)"][j] < 100 ? 1 : 0)} MWh`, "a year"], [`₹${inr(sc["Electricity saved, year 1 (₹ lakh)"][j], 1)} L`, "electricity saved a year"], [`${inr(sc["CO2 avoided per year (t)"][j], 1)} t`, "CO2 avoided a year"], [yrs(sc["Simple payback on full cost (years)"][j]), j === 0 ? "payback at full pilot cost" : "payback, net of walkway credit"]];
    kv.forEach(([v, l], k) => {
      text(s, v, { x: x + 0.2, y: 2.65 + k * 0.6, w: w - 0.4, h: 0.32, fontSize: 16, bold: true, color: col });
      text(s, l, { x: x + 0.2, y: 2.95 + k * 0.6, w: w - 0.4, h: 0.26, fontSize: 11, color: C.text2 });
    });
  });
  source(s, `Portfolio: ${CFG.portfolio_label}. Phase areas are editable on the model's 'Client' sheet.`, 5.95);
  s.addNotes("The rollout path. Each phase is a decision on its own, taken on measured data from the one before.");

  // ============================================================ 15. CREDENTIALS
  s = content("Plan", "Built on the GridMats platform", "The structural platform is proven in roads and surfaces; the Solar Walkway is the next product on it.");
  const cr = [["32 sites, 65,081 m²", "built across 12 states"], ["1 million+ sq ft", "of GridMats supplied"], ["Indian Patent 398992", "patented 3D recycled-polymer grid"], ["CII GreenPro", "certified green product"], ["IISc · IIT Hyderabad", "R&D and pavement testing partners"], ["100% recycled PP", "post-consumer polypropylene base"]];
  cr.forEach((c, i) => {
    const x = MX + (i % 3) * 2.45, y = 1.75 + Math.floor(i / 3) * 1.45;
    card(s, x, y, 2.3, 1.3, C.background1, "cred-card");
    text(s, c[0], { x: x + 0.15, y: y + 0.18, w: 2.0, h: 0.5, fontSize: 15, bold: true, color: C.accent1 });
    text(s, c[1], { x: x + 0.15, y: y + 0.7, w: 2.0, h: 0.5, fontSize: 12, color: C.text2 });
  });
  s.addImage({ path: asset("client_logos.jpg"), x: MX, y: 4.75, w: 7.2, h: 1.81, altText: "Existing GridMats customers" });
  card(s, 8.15, 1.75, 4.58, 4.81, C.background1, "founder-card");
  s.addImage({ path: asset("founder.jpg"), x: 8.4, y: 1.95, w: 1.6, h: 2.2, sizing: { type: "cover", w: 1.6, h: 2.2 }, altText: "Dr. Prathap Bhimasena Rao" });
  text(s, "Dr. Prathap Bhimasena Rao", { x: 10.15, y: 2.0, w: 2.45, h: 0.6, fontSize: 15, bold: true });
  text(s, "Founder & Managing Director", { x: 10.15, y: 2.6, w: 2.45, h: 0.5, fontSize: 12, bold: true, color: C.accent1 });
  text(s, ["26 years in leadership and innovation", "Ex-Indian Air Force; Global VP, Hewlett-Packard (2010-11)", "Two PhDs in management science"].map((t, j, a) => ({ text: t, options: { bullet: true, breakLine: j < a.length - 1 } })),
    { x: 8.4, y: 4.35, w: 4.1, h: 1.8, fontSize: 12, color: C.text2, paraSpaceAfter: 4 });
  source(s, "Figures and customers relate to existing (non-solar) GridMats road and surface applications. The Solar Walkway is pre-commercial.", 6.65);
  s.addNotes("Credentials with the approved figures. Make the distinction explicit: customers and metrics are for the existing GridMats platform; the Solar Walkway is entering its first pilots.");

  // ============================================================ 16. NEXT STEPS
  pres.addSection({ title: "Close" });
  s = pres.addSlide({ masterName: "CLOSING", sectionTitle: "Close" });
  s.addText("Let's walk the site", { placeholder: "title" });
  s.addText(`What GridMats needs from ${client}, and the four steps to a measured result.`, { placeholder: "body" });
  text(s, `FROM ${client.toUpperCase()}`, { x: MX, y: 2.05, w: 5.5, h: 0.35, fontSize: 13, bold: true, color: "F2B33D" });
  text(s, CFG.site_requirements.map((t, j, a) => ({ text: t, options: { bullet: true, breakLine: j < a.length - 1 } })), { x: MX, y: 2.5, w: 5.6, h: 3.4, fontSize: 14, color: C.background1, paraSpaceAfter: 8 });
  const ns = [["Site walk and opportunity assessment", "Shade, footfall, drainage and the electrical connection."], ["Concept layout and generation estimate", "Module layout, monthly kWh, final price."], ["Pilot engineering and order", "Electrical, safety and monitoring details; PO with the supply date."], ["Install, measure, decide", `2-3 days on site; ${CFG.evaluation_months}-month scorecard review.`]];
  ns.forEach(([h, d2], i) => {
    const y = 2.05 + i * 1.05;
    badge(s, String(i + 1), 6.9, y, 0.55, C.accent1);
    text(s, h, { x: 7.65, y: y - 0.02, w: 5.1, h: 0.32, fontSize: 15, bold: true, color: C.background1 });
    text(s, d2, { x: 7.65, y: y + 0.32, w: 5.1, h: 0.5, fontSize: 13, color: C.background2 });
  });
  text(s, "Dr. Prathap Bhimasena Rao  ·  prathap@gridmats.com  ·  sales@gridmats.com  ·  +91 63626 46287  ·  www.gridmats.com", { x: MX, y: 6.55, w: 11.5, h: 0.35, fontSize: 12, color: C.background1 });
  s.addNotes("Close on the site walk: name the walkway, the connection point and the facilities contact, and the pilot can be engineered within weeks.");

  // ============================================================ 17. APPENDIX
  pres.addSection({ title: "Appendix" });
  s = content("Appendix", "Key assumptions and sources", "Every value is an input in the companion model; change it there and rebuild the deck.");
  const ap = [[th("Assumption"), th("Value", "C22935"), th("Source / basis")]];
  [["Design basis", `40 W module · 4 per m² · ${inr(O.PowerDens)} W/m²`, "Flier (6 Oct 2026); Client/Partner Deck v2"],
    ["Sunshine", `${CFG.irradiance_city}: ${inr(O.SunYear)} kWh/m²/yr`, "Typical year; replace with site data"],
    ["Cell technology", `${O.TechSel}: ${(O.TechTempCoef * 100).toFixed(2)}%/°C; ${pct(O.Deg1)} then ${pct(O.Degr, 1)}/yr`, "2026 datasheets; walk-on wear +0.2%/yr"],
    ["Performance ratio; shading", `${pct(O.PR)} (seven losses); ${pct(CFG.shading_loss)}`, "Loss stack on 'Product & Costs'; site assessment"],
    ["Power price; escalation", `₹${CFG.tariff.toFixed(2)}/kWh; ${pct(CFG.tariff_escalation)}/yr`, CFG.tariff_source || "Site electricity bill"],
    ["PV layer price", `₹${inr(O.PVPilotSqft)}/sq ft pilot; ₹${inr(8500 / 10.7639)}/sq ft rollout`, "₹9,500/m² quote; ₹8,000-9,000/m² range"],
    ["Base and electrical", "OC-LW 40 ₹64/sq ft; BOS ₹8/Wp", "Price ladder (Sep 2026); earlier ROI model"],
    ["O&M; life", "1.5% of capex/yr; 20 years", "Earlier ROI model"],
    ["Rooftop; carport", "₹40/Wp; ₹61/Wp all-in", "Earlier ROI model benchmarks"],
    ["GST", "8.9% composite (5% goods, 18% services)", "Notification 9/2025, from 22 Sep 2025"]].forEach((r) => ap.push(r.map((t, j) => ({ text: t, options: { bold: j === 0 } }))));
  s.addTable(ap, tblOpts({ x: MX, y: 1.68, w: CW, colW: [3.0, 4.4, 4.73], rowH: 0.43, fontSize: 12 }));
  source(s, `Changes from earlier calculations: walkway power 126-110 → 160 W/m²; PV layer ₹2,500 (input error) → ₹9,500/m²; base HD ₹1,200 → OC-LW ₹690/m²; flat 78% PR → ${pct(O.PR)} loss stack; GST 13.8% → 8.9%.`, 6.55);
  s.addNotes("Reference slide. Every value here is an editable input in the model.");

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
}
main().catch((e) => { console.error(e); process.exit(1); });
