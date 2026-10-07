// GridMats Solar x Reliance: cost-benefit and pilot proposal deck (pptxgenjs, structured).
const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const Fa = require("react-icons/fa");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");

const OUT = process.argv[2];
const O = JSON.parse(fs.readFileSync(path.join(__dirname, "outputs.json")));
const D = JSON.parse(fs.readFileSync(path.join(__dirname, "chartdata.json")));

// ------------------------------------------------------------------ theme
const THEME = {
  name: "GridMats",
  headFontFace: "Arial",
  bodyFontFace: "Arial",
  colors: {
    dk1: "1A1A1A", lt1: "FFFFFF", dk2: "3F3F3F", lt2: "F6F6F4",
    accent1: "C22935", accent2: "1F7A3A", accent3: "B87503", accent4: "6F6F6F",
    accent5: "FBEAEC", accent6: "A11F2A", hlink: "C22935", folHlink: "A11F2A",
  },
};
const HX = THEME.colors; // hex values for hex-only options (charts, gridlines, shadows, icons)
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.author = "GridMats (GroundReality Enterprises Pvt Ltd)";
pres.company = "GridMats";
pres.title = "GridMats Solar x Reliance: cost-benefit case and pilot proposal";
pres.subject = "Solar-park roads, solar walkways and a pilot offer for Reliance New Energy";
const C = pres.SchemeColor;
const W = 13.333;
const MX = 0.6; // side margin
const CW = W - 2 * MX; // content width

// ------------------------------------------------------------------ formatting
const inr = (n, d = 0) => Number(n).toLocaleString("en-IN", { minimumFractionDigits: d, maximumFractionDigits: d });
const cr = (n, d = 1) => `₹${inr(n, d)} Cr`;
const lakh = (rupees, d = 1) => `₹${inr(rupees / 1e5, d)} L`;
const pct = (f, d = 0) => `${(f * 100).toFixed(d)}%`;
const yrs = (n) => `${Number(n).toFixed(1)} yrs`;

// ------------------------------------------------------------------ icons
const iconCache = {};
async function icon(name, color = "FFFFFF", size = 256) {
  const key = `${name}-${color}`;
  if (iconCache[key]) return iconCache[key];
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Fa[name], { color: "#" + color, size: String(size) }));
  const png = await sharp(Buffer.from(svg)).png().toBuffer();
  return (iconCache[key] = "image/png;base64," + png.toString("base64"));
}

// ------------------------------------------------------------------ layouts
const FOOTER = "GridMats Solar × Reliance  |  Proposal for discussion  |  October 2026";
pres.defineSlideMaster({
  title: "TITLE_DARK",
  background: { color: C.text1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: MX, y: 2.2, w: 9.2, h: 2.1, fontSize: 44, bold: true, color: C.background1, valign: "bottom", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: MX, y: 4.5, w: 9.2, h: 1.2, fontSize: 18, color: C.background2, valign: "top", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: C.background1 },
  margin: [0.4, 0.6, 0.7, 0.6],
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: MX, y: 0.35, w: CW, h: 0.75, fontSize: 34, bold: true, color: C.text1, valign: "middle", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: MX, y: 1.1, w: CW, h: 0.5, fontSize: 16, color: C.accent4, valign: "top", margin: 0 }, text: "" } },
    { text: { text: FOOTER, options: { x: MX, y: 7.02, w: 8, h: 0.3, fontSize: 10, color: C.accent4, margin: 0 } } },
  ],
  slideNumber: { x: 12.23, y: 7.02, w: 0.5, h: 0.3, fontSize: 10, color: C.accent4, align: "right" },
});
pres.defineSlideMaster({
  title: "CLOSING_DARK",
  background: { color: C.text1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: MX, y: 0.5, w: CW, h: 0.9, fontSize: 38, bold: true, color: C.background1, valign: "middle", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: MX, y: 1.4, w: CW, h: 0.5, fontSize: 16, color: C.background2, valign: "top", margin: 0 }, text: "" } },
  ],
  slideNumber: { x: 12.23, y: 7.02, w: 0.5, h: 0.3, fontSize: 10, color: C.background2, align: "right" },
});

// ------------------------------------------------------------------ helpers
function content(section, title, lead) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
  s.addText(title, { placeholder: "title" });
  if (lead) s.addText(lead, { placeholder: "body" });
  return s;
}
const shadow = () => ({ type: "outer", color: "000000", blur: 6, offset: 1.5, angle: 90, opacity: 0.12 });
function card(s, x, y, w, h, fill = C.background2, name = "card") {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: fill }, line: { color: fill, width: 0 }, objectName: name });
}
async function hexIcon(s, iconName, x, y, d = 0.62, fill = C.accent1, iconColor = "FFFFFF") {
  s.addShape(pres.shapes.HEXAGON, { x, y, w: d, h: d * 0.88, fill: { color: fill }, line: { color: fill, width: 0 }, objectName: `hex-${iconName}` });
  const ic = d * 0.46;
  s.addImage({ data: await icon(iconName, iconColor), x: x + (d - ic) / 2, y: y + (d * 0.88 - ic) / 2, w: ic, h: ic, altText: iconName.replace(/^Fa/, "") });
}
function text(s, t, o) {
  s.addText(t, Object.assign({ isTextBox: true, margin: 0, fontSize: 14, color: C.text1, valign: "top" }, o));
}
function stat(s, x, y, w, value, label, color = C.accent1, vSize = 36) {
  text(s, value, { x, y, w, h: vSize / 60, fontSize: vSize, bold: true, color, valign: "bottom", objectName: "stat-value" });
  text(s, label, { x, y: y + vSize / 60 + 0.05, w, h: 0.62, fontSize: 13, color: C.text2, objectName: "stat-label" });
}
function source(s, t, y = 6.62) {
  text(s, t, { x: MX, y, w: CW, h: 0.3, fontSize: 10, color: C.accent4, italic: true, objectName: "source" });
}
const chartBase = () => ({
  catAxisLabelColor: "3F3F3F", valAxisLabelColor: "6F6F6F", catAxisLabelFontSize: 12, valAxisLabelFontSize: 11,
  catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", legendFontFace: "+mn-lt", titleFontFace: "+mn-lt",
  valGridLine: { color: "E3E3E3", size: 0.5 }, catGridLine: { style: "none" },
  dataLabelFontSize: 11, dataLabelColor: "1A1A1A", legendFontSize: 11, legendColor: "3F3F3F",
  catAxisLineShow: true, valAxisLineShow: false, catAxisLineColor: "9A9A9A",
});

async function main() {
  // ================================================================ 1. TITLE
  pres.addSection({ title: "Opening" });
  let s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "Opening" });
  s.addText("Roads that keep panels clean. Walkways that make power.", { placeholder: "title" });
  s.addText("GridMats Solar × Reliance New Energy: cost-benefit case and pilot proposal\nPrepared for Reliance Industries Ltd  ·  October 2026", { placeholder: "body" });
  // hexagon motif cluster
  const hexes = [[10.3, 1.3, C.accent1], [11.35, 1.9, C.accent6], [10.3, 2.5, C.text2], [11.35, 3.1, C.accent1], [10.3, 3.7, C.accent6], [11.35, 4.3, C.text2]];
  for (const [x, y, f] of hexes) s.addShape(pres.shapes.HEXAGON, { x, y, w: 1.15, h: 1.0, fill: { color: f }, line: { color: C.text1, width: 2 }, objectName: "motif-hex" });
  s.addImage({ data: await icon("FaSolarPanel", "FFFFFF"), x: 11.6, y: 3.3, w: 0.62, h: 0.62, altText: "solar panel" });
  s.addImage({ data: await icon("FaRoad", "FFFFFF"), x: 10.55, y: 2.7, w: 0.62, h: 0.62, altText: "road" });
  s.addImage({ data: await icon("FaRecycle", "FFFFFF"), x: 10.55, y: 3.9, w: 0.62, h: 0.62, altText: "recycle" });
  text(s, "GridMats Engineered Surface Systems  ·  www.gridmats.com", { x: MX, y: 6.7, w: 8, h: 0.35, fontSize: 12, color: C.background2 });
  s.addNotes("Open with the one-line idea: the same GridMats grid that Reliance already used at its Kutch solar block can (1) cut the life-cycle cost of solar-park roads, (2) keep dust off modules, and (3) as a GridMats Solar Walkway, turn walkways into generators. Every number in this deck comes from the companion Excel model; assumptions are labelled and the pilot is designed to measure them.");

  // ================================================================ 2. EXEC SUMMARY
  s = content("Opening", "The case in one page", "GridMats on Reliance's solar parks pays back in under five years; a pilot this dry season proves it on Reliance's own data.");
  const rows = [
    ["FaRoad", "Solar-park roads", `S1 on main roads, S2 on O&M tracks: NPV ${cr(O.Park_NPV)} per 100 MW block over 25 years, ${pct(O.Park_IRR)} IRR, ${yrs(O.Park_Payback)} payback; premium only ${pct(O.Park_PremPct, 1)} of plant capex.`],
    ["FaSolarPanel", "Cleaner modules", `Dust-confining tracks cut traffic dust on nearby rows: about ${inr(Math.round(D.mwh_y1 / 10) * 10)} MWh more a year per 100 MW (assumption the pilot measures).`],
    ["FaWalking", "Walkways that generate", `GridMats Solar Walkway, about 160 W/m²: payback about ${O.Wk_SP_Mum.toFixed(1)} years in Navi Mumbai and ${O.Wk_SP_Jam.toFixed(1)} at Jamnagar.`],
    ["FaRecycle", "Circular and low-carbon", `Per 100 MW: ${inr(O.Sus_PP)} t of recycled PP put to use and about ${inr(Math.round(O.Sus_CO2 / 100) * 100)} t CO2e avoided over 25 years.`],
  ];
  for (let i = 0; i < rows.length; i++) {
    const y = 1.9 + i * 1.17;
    await hexIcon(s, rows[i][0], MX, y, 0.7);
    text(s, rows[i][1], { x: MX + 0.95, y: y - 0.02, w: 6.85, h: 0.35, fontSize: 16, bold: true });
    text(s, rows[i][2], { x: MX + 0.95, y: y + 0.36, w: 6.85, h: 0.72, fontSize: 14, color: C.text2 });
  }
  card(s, 8.55, 1.95, 4.18, 4.55, C.accent5, "ask-card");
  text(s, "The ask", { x: 8.85, y: 2.15, w: 3.6, h: 0.4, fontSize: 18, bold: true, color: C.accent1 });
  text(s, [
    { text: "Pilot this dry season", options: { bold: true, breakLine: true } },
    { text: "1 km of solar-park road (S1 + S2) beside control sections, a 50 m² GridMats Solar Walkway, and 12 months of joint measurement.", options: { breakLine: true } },
    { text: " ", options: { fontSize: 8, breakLine: true } },
    { text: `Reliance pays ${lakh(O.Pilot_Offer)}`, options: { bold: true, breakLine: true } },
    { text: `for ${lakh(O.Pilot_List)} of scope (ex-GST). Grid at ₹900/m², ${pct(O.Pilot_vsTrial)} below the 2025 trial price.`, options: { breakLine: true } },
    { text: " ", options: { fontSize: 8, breakLine: true } },
    { text: "Decision at 12 months", options: { bold: true, breakLine: true } },
    { text: "If the measures are met, ₹900/m² is held for a park-wide 12-month rate contract.", options: {} },
  ], { x: 8.85, y: 2.65, w: 3.6, h: 3.6, fontSize: 14, color: C.text1, paraSpaceAfter: 4 });
  source(s, "Base case of the companion model (GridMats_Reliance_Solar_CostBenefit_Model.xlsx). Money ex-GST, 2026 prices, 10% discount rate.");
  s.addNotes("Four messages, one ask. Lead with the roads: that is where the money is. The soiling benefit and the walkway economics are flagged as assumptions to be measured. The ask is small relative to Reliance's build programme: about one crore all-in including civil works by Reliance's own EPC.");

  // ================================================================ 3. CONTEXT
  pres.addSection({ title: "Why now" });
  s = content("Why now", "Kutch is being built now, on roads", "Every block needs its roads before a single module goes up, and the choice of road shows up in O&M for 25 years.");
  const facts = [
    ["5.5 lakh", "acres at Kutch for Reliance's renewable giga-project"],
    [">40 bn", "kWh a year planned output, about 3% of India's demand"],
    ["55 MWp", "of modules a day at peak installation"],
    ["8-12 km", "of internal road for every 100 MW of solar"],
  ];
  facts.forEach((f, i) => {
    const x = MX + (i % 2) * 3.2, y = 1.95 + Math.floor(i / 2) * 2.05;
    card(s, x, y, 2.95, 1.8, C.background2, "fact-card");
    stat(s, x + 0.25, y + 0.2, 2.5, f[0], f[1], C.accent1, 34);
  });
  text(s, "What conventional roads cost a solar park", { x: 7.25, y: 1.95, w: 5.5, h: 0.4, fontSize: 18, bold: true });
  const pains = [
    ["FaWind", "WBM / gravel", "ruts, washes out and raises dust; regraded every monsoon"],
    ["FaClock", "Concrete", "14-28 days of curing on the critical path"],
    ["FaTools", "Bitumen", "hot mix on site; renewal every 4-6 years"],
    ["FaSun", "Dust on modules", "West-India soiling 0.3-0.5% a day in the dry season"],
  ];
  for (let i = 0; i < pains.length; i++) {
    const y = 2.55 + i * 0.98;
    await hexIcon(s, pains[i][0], 7.25, y, 0.6, C.text2);
    text(s, [{ text: pains[i][1] + ": ", options: { bold: true } }, { text: pains[i][2] }], { x: 8.05, y: y + 0.02, w: 4.7, h: 0.75, fontSize: 14 });
  }
  source(s, "Reliance AGM 2026 as reported (BusinessToday, 19 Jun 2026); GridMats solar and wind parks flier; IRC:82-2015; soiling studies (Gandhinagar 0.37%/day).");
  s.addNotes("Reliance's Kutch build-out is ramping from this post-monsoon season. Roads come first. Each conventional choice has a known cost: gravel needs regrading and throws dust on modules, concrete puts curing on the critical path, bitumen needs periodic renewal. GridMats is designed to avoid all three.");

  // ================================================================ 4. PROOF
  pres.addSection({ title: "Proof" });
  s = content("Proof", "Already proven on Reliance ground", "Reliance New Energy, Z2 block, Kutch: GridMats OC-HD 40 access road on saline, low-CBR ground (ordered Sep 2025).");
  const proof = [
    ["FaLayerGroup", "2,000 m²", "OC-HD 40 grid on 200 gsm geotextile"],
    ["FaFlask", "IITPAVE", "designed for CBR 3-5% saline subgrade"],
    ["FaHardHat", "4 days", "to train Reliance's own civil crew"],
    ["FaTruck", "Same day", "opened to traffic the day it was filled"],
  ];
  for (let i = 0; i < proof.length; i++) {
    const x = MX + i * 1.62;
    await hexIcon(s, proof[i][0], x + 0.42, 1.95, 0.66);
    text(s, proof[i][1], { x, y: 2.7, w: 1.5, h: 0.45, fontSize: 20, bold: true, color: C.accent1, align: "center" });
    text(s, proof[i][2], { x, y: 3.15, w: 1.5, h: 0.9, fontSize: 12, color: C.text2, align: "center" });
  }
  card(s, MX, 4.35, 6.35, 1.75, C.background2, "learn-card");
  text(s, [
    { text: "What the trial proved, and what it did not measure", options: { bold: true, breakLine: true } },
    { text: "It proved the build: Reliance's crew laid the grid and opened the road at once. It did not record rutting, dust or upkeep against a conventional section. The pilot adds exactly that, so the next decision rests on Reliance's own data." },
  ], { x: MX + 0.25, y: 4.5, w: 5.85, h: 1.5, fontSize: 14, paraSpaceAfter: 4 });
  // cross-section of the 2025 design
  text(s, "2025 design, CBR 5% section (GM/OC/HD-5/GR)", { x: 7.35, y: 1.95, w: 5.4, h: 0.35, fontSize: 14, bold: true });
  const sec25 = [
    ["12 mm-down aggregate in 40 mm OC-HD grid", "70 mm", "C22935", "FFFFFF", 0.42],
    ["Upper sub-base, 20-40 mm blended gravel", "300 mm", "B8B1A3", "1A1A1A", 0.95],
    ["GSB Grade II", "200 mm", "CFC8B8", "1A1A1A", 0.7],
    ["200 gsm geofabric separator", "", "1A1A1A", "FFFFFF", 0.12],
    ["Subgrade, CBR 5%", "", "8C7B6B", "FFFFFF", 0.55],
  ];
  let yy = 2.45;
  for (const [lab, th, fill, tc, h] of sec25) {
    s.addShape(pres.shapes.RECTANGLE, { x: 7.35, y: yy, w: 3.6, h, fill: { color: fill }, line: { color: "FFFFFF", width: 0.75 }, objectName: "layer" });
    if (h > 0.2) text(s, lab, { x: 7.45, y: yy, w: 3.4, h, fontSize: 11, color: tc, valign: "middle" });
    text(s, th || (h <= 0.2 ? lab : ""), { x: 11.1, y: yy + h / 2 - 0.16, w: 1.7, h: 0.32, fontSize: 12, color: C.text2, valign: "middle" });
    yy += h;
  }
  source(s, "Reliance purchase and work orders (Sep 2025); GridMats design note 'Design of gravel roads for low CBR% soils', IITPAVE (Aug 2025). Not to scale.");
  s.addNotes("Reliance has already bought and built GridMats once. Be candid: it was a trial without measurement. The pilot proposed here is designed to produce the performance data that the 2025 trial did not, against control sections built the conventional way.");

  // ================================================================ 5. GRIDMATS
  s = content("Proof", "GridMats: surfaces made from waste plastic", "Six systems, one GridMats platform: interlocking grids of 100% post-consumer recycled polypropylene.");
  const creds = [
    ["FaRecycle", "100% recycled PP", "6.6 kg of post-consumer plastic in every m² of OC-HD 40"],
    ["FaCubes", "407 t/m²", "lab test on the empty OC-HD 40 panel (Bureau Veritas, 2021)"],
    ["FaChartLine", "2.0× base modulus", "at CBR 5% with the grid at the GSB/WMM interface (IIT Hyderabad, interim report)"],
    ["FaAward", "Patent and GreenPro", "Indian Patent 398992; CII GreenPro certified"],
    ["FaMapMarkerAlt", "32 sites, 65,081 m²", "built across 12 states; over a million sq ft supplied"],
    ["FaIndustry", "Trusted by industry", "Reliance, Toyo Engineering, Tata Advanced Systems, GMR Hyderabad Airport, Godrej, Aditya Birla"],
  ];
  for (let i = 0; i < creds.length; i++) {
    const col = i % 3, row = Math.floor(i / 3);
    const x = MX + col * 4.1, y = 1.95 + row * 1.6;
    await hexIcon(s, creds[i][0], x, y, 0.62);
    text(s, creds[i][1], { x: x + 0.8, y: y - 0.02, w: 3.15, h: 0.38, fontSize: 16, bold: true });
    text(s, creds[i][2], { x: x + 0.8, y: y + 0.38, w: 3.15, h: 1.0, fontSize: 13, color: C.text2 });
  }
  card(s, MX, 5.25, CW, 1.2, C.accent5, "godrej-card");
  text(s, [
    { text: "Recent heavy-duty result: ", options: { bold: true, color: C.accent1 } },
    { text: "Godrej, Vikhroli (May 2026). 400 m² of OC-HD 40 built in 3 days, carrying about 50-80 trucks a day with no cracks since June, at 21% less than a conventional concrete road." },
  ], { x: MX + 0.3, y: 5.42, w: CW - 0.6, h: 0.9, fontSize: 14, valign: "middle" });
  source(s, "GridMats master data Rev 1.3; Bureau Veritas BVIPL:SPL:BL/2545/7/2021/A; IIT Hyderabad interim report (Oct 2025); Godrej site observation (approximate truck count).");
  s.addNotes("Credentials, with the approved wording. 407 t/m² is a lab test on the empty panel, not a pavement capacity. IIT Hyderabad work is an interim report presented at TGS2026. The Godrej Vikhroli road is the most recent heavy-truck proof.");

  // ================================================================ 6. TOOLKIT
  pres.addSection({ title: "Solution" });
  s = content("Solution", "GridMats Solar: four tools for Reliance", "Each part of a solar park gets the surface it needs; all of them share one grid, one method statement and one supplier.");
  const tools = [
    ["FaTruck", "S1  Cement-treated gravel", "Main roads, inverter and transformer yards", ["Vehicles up to 49 t gross (GVW)", "Light traffic at 3 days, trucks at 7", "No saw-cut joints, no hot mix"]],
    ["FaRoad", "S2  Permeable stone chip", "O&M tracks, laydown and substation yards", ["Opens the same day it is filled", "Rain goes through; no surface drains", "Chips held in cells, so less dust"]],
    ["FaLayerGroup", "S6  Top-covered mats", "Crane pads and temporary access", ["Laid or lifted in a day", "Moves with the installation crew", "Reused block after block"]],
    ["FaSolarPanel", "Solar Walkway + IoT", "Walkways, EV hubs, campus paths (pilot stage)", ["About 160 W of PV per m²", "No extra land, roof or mounting", "Sensors report load, moisture, heat"]],
  ];
  for (let i = 0; i < tools.length; i++) {
    const x = MX + i * 3.08, w = 2.86;
    card(s, x, 1.95, w, 4.5, i === 3 ? C.accent5 : C.background2, "tool-card");
    await hexIcon(s, tools[i][0], x + 0.25, 2.15, 0.66, i === 3 ? C.accent3 : C.accent1);
    text(s, tools[i][1], { x: x + 0.25, y: 2.95, w: w - 0.5, h: 0.7, fontSize: 16, bold: true });
    text(s, tools[i][2], { x: x + 0.25, y: 3.65, w: w - 0.5, h: 0.7, fontSize: 13, color: C.accent1, bold: true });
    text(s, tools[i][3].map((t, j) => ({ text: t, options: { bullet: true, breakLine: j < tools[i][3].length - 1 } })),
      { x: x + 0.25, y: 4.4, w: w - 0.4, h: 1.9, fontSize: 13, color: C.text2, paraSpaceAfter: 6 });
  }
  source(s, "GridMats solar and wind parks flier (Oct 2026); GridMats Solar Walkway flier (6 Oct 2026). Solar Walkway and IoT are at pilot stage.");
  s.addNotes("S1 for heavy roads and yards, S2 for O&M tracks (what Reliance used at Kutch), S6 mats for temporary crane and access pads, and the Solar Walkway with IoT sensing for campuses. One supplier, one method statement for every EPC contractor, hold-point QC by GridMats engineers.");

  // ================================================================ 7. PAVEMENT DESIGN
  s = content("Solution", "Pavement design for Kutch's saline ground", "GridMats base tables for CBR 2-5%; final thickness from IITPAVE on the pilot block's measured CBR.");
  function drawSection(x, y, title, sub, layers, cost) {
    text(s, title, { x, y, w: 5.6, h: 0.38, fontSize: 18, bold: true });
    text(s, sub, { x, y: y + 0.38, w: 5.6, h: 0.32, fontSize: 13, color: C.accent1, bold: true });
    let yy2 = y + 0.85;
    for (const [lab, th, fill, tc, h] of layers) {
      s.addShape(pres.shapes.RECTANGLE, { x, y: yy2, w: 3.3, h, fill: { color: fill }, line: { color: "FFFFFF", width: 0.75 }, objectName: "layer" });
      if (h > 0.2) text(s, lab, { x: x + 0.1, y: yy2, w: 3.1, h, fontSize: 11, color: tc, valign: "middle" });
      text(s, h > 0.2 ? th : `${lab}${th ? "  " + th : ""}`, { x: x + 3.45, y: yy2 + h / 2 - 0.16, w: 2.1, h: 0.32, fontSize: 12, color: C.text2, valign: "middle" });
      yy2 += h;
    }
    text(s, cost, { x, y: yy2 + 0.15, w: 5.6, h: 0.35, fontSize: 13, bold: true, color: C.text1 });
  }
  drawSection(MX, 1.85, "S1 main road and yards", "Load class D: vehicles up to 49 t GVW", [
    ["OC-HD 40 + cement-treated gravel", "40 mm", "C22935", "FFFFFF", 0.38],
    ["M-sand bedding", "25 mm", "E6E1D6", "1A1A1A", 0.24],
    ["WMM", "150 mm", "B8B1A3", "1A1A1A", 0.9],
    ["GSB", "200 mm", "CFC8B8", "1A1A1A", 1.1],
    ["Geotextile 200 gsm", "", "1A1A1A", "FFFFFF", 0.1],
    ["Subgrade, CBR 2-5%, 98% MDD", "", "8C7B6B", "FFFFFF", 0.5],
  ], `Installed about ₹${inr(O.S1_Capex)}/m²  ·  bitumen ₹2,000  ·  concrete ₹3,200`);
  drawSection(6.75, 1.85, "S2 O&M track", "Load class C: trucks and water tankers", [
    ["OC-HD 40 + 8-12 mm stone chips", "40 mm", "C22935", "FFFFFF", 0.38],
    ["Bedding", "20 mm", "E6E1D6", "1A1A1A", 0.22],
    ["WMM", "150 mm", "B8B1A3", "1A1A1A", 0.9],
    ["GSB", "100 mm", "CFC8B8", "1A1A1A", 0.6],
    ["Geotextile 200 gsm", "", "1A1A1A", "FFFFFF", 0.1],
    ["Subgrade, CBR 2-5%, 98% MDD", "", "8C7B6B", "FFFFFF", 0.5],
  ], `Installed about ₹${inr(O.S2_Capex)}/m²  ·  WBM ₹600`);
  source(s, "GridMats Systems Manual Rev 1.3 base tables; unit rates from Design Engine 73 BoQ with grid at ₹900/m². U-nails 1 per 2-3 m², edge restraint throughout. Not to scale.");
  s.addNotes("Both sections sit on geotextile over Kutch's saline, low-CBR formation, as at Z2. Final layer thicknesses are confirmed with IITPAVE once CBR and axle loads are measured at the pilot block. Costs are the model's bottom-up build-ups, ex-GST, with Reliance's EPC doing the civil works.");

  // ================================================================ 8. SOLAR WALKWAY DESIGN
  s = content("Solution", "Solar Walkway: paving that also powers", "Walkable PV on wiring lumber over the GridMats base. Pilot stage: the pilot measures yield, heat and grip.");
  // exploded layers
  const lx = MX + 0.2, lw = 5.2;
  const exploded = [
    ["Walkable PV layer: 40 W module on each grid", "2B3A55", 2.05],
    ["MLP lumber, 40 × 40 mm, carries the wiring; 50 mm gaps drain", "8C8C8C", 3.25],
    ["GridMats structural base, 100% recycled PP, on open-graded GSB", "C22935", 4.45],
  ];
  for (const [lab, fill, y] of exploded) {
    s.addShape(pres.shapes.PARALLELOGRAM, { x: lx, y, w: lw, h: 0.75, fill: { color: fill }, line: { color: "FFFFFF", width: 1 }, objectName: "walk-layer" });
  }
  // grid lines on PV layer suggest cells
  for (let k = 1; k < 6; k++) s.addShape(pres.shapes.LINE, { x: lx + 0.35 + k * 0.8, y: 2.08, w: 0.0, h: 0.69, line: { color: "6E7FA0", width: 1 }, objectName: "cell-line" });
  for (let k = 0; k < 7; k++) s.addShape(pres.shapes.RECTANGLE, { x: lx + 0.45 + k * 0.68, y: 3.42, w: 0.36, h: 0.42, fill: { color: "6F6F6F" }, line: { color: "6F6F6F", width: 0 }, objectName: "lumber" });
  for (let k = 0; k < 8; k++) s.addShape(pres.shapes.HEXAGON, { x: lx + 0.42 + k * 0.58, y: 4.6, w: 0.48, h: 0.42, fill: { color: "A11F2A" }, line: { color: "FFFFFF", width: 1 }, objectName: "base-hex" });
  exploded.forEach(([lab, , y], i) => text(s, lab, { x: lx, y: y + 0.78, w: lw, h: 0.4, fontSize: 12, color: C.text2 }));
  const ws_ = [["~160 W/m²", "rated power (4 grids × 40 W per m²)"], ["0", "extra land, roof or mounting"], ["15-16%", "surface efficiency"]];
  ws_.forEach((st, i) => stat(s, 7.0, 1.95 + i * 1.12, 2.6, st[0], st[1], i === 0 ? C.accent3 : C.accent1, 30));
  text(s, "Pilot method", { x: 9.85, y: 1.95, w: 2.9, h: 0.38, fontSize: 16, bold: true });
  text(s, ["Site check: shade, footfall, drainage, grid connection", "Engineering: module and lumber layout, cabling, inverter, safety", "Field pilot: measured W/m², kWh a day, module heat, grip, uptime", "Scale-up with lighting, sensors and batteries"].map((t, j, a) => ({ text: t, options: { bullet: { type: "number" }, breakLine: j < a.length - 1 } })),
    { x: 9.85, y: 2.4, w: 2.9, h: 3.4, fontSize: 13, color: C.text2, paraSpaceAfter: 6 });
  card(s, 7.0, 5.45, 5.73, 0.95, C.accent5, "hjt-card");
  text(s, [{ text: "Co-development option: ", options: { bold: true, color: C.accent1 } }, { text: "build the walkable module around Reliance's own Jamnagar HJT cells." }],
    { x: 7.2, y: 5.55, w: 5.35, h: 0.75, fontSize: 14, valign: "middle" });
  source(s, "GridMats Solar Walkway flier (6 Oct 2026). Status: pre-commercial; pilot validation under way; no commercial deployment claimed. Concept diagram, not to scale.");
  s.addNotes("Be clear this is pilot-stage. The design basis is about 160 W per m² of walkway: one 40 W module on each of four grids per m². What the pilot measures: real W/m², kWh per m² per day, module temperature, slip resistance and uptime. A natural extension for Reliance: co-develop the walkable module using Reliance HJT cells from Jamnagar.");

  // ================================================================ 9. DAY ONE VS 25 YEARS
  pres.addSection({ title: "Economics" });
  s = content("Economics", "Day one vs 25 years", "GridMats costs more than gravel to build and about the same as bitumen, and less than both once upkeep is counted.");
  const lccOpts = (title, cats, build, life) => ({
    data: [{ name: "Build cost", labels: cats, values: build }, { name: "25-year life-cycle cost (PV)", labels: cats, values: life }],
    opts: Object.assign(chartBase(), {
      barDir: "col", barGrouping: "clustered", barGapWidthPct: 60, chartColors: ["9A9A9A", HX.accent1],
      showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "#,##0",
      showLegend: true, legendPos: "b", showTitle: true, title, titleFontSize: 14, titleColor: "1A1A1A",
      valAxisLabelFormatCode: "#,##0", valAxisMinVal: 0, valAxisMaxVal: 4000, valAxisMajorUnit: 1000,
    }),
  });
  let ch = lccOpts("O&M tracks (₹ per m²)", ["WBM / gravel", "GridMats S2"], [O.LCC_WBM - O.Upk_WBM, O.S2_Capex], [O.LCC_WBM, O.LCC_S2]);
  s.addChart(pres.charts.BAR, ch.data, Object.assign({ x: MX, y: 1.85, w: 4.6, h: 4.0 }, ch.opts));
  ch = lccOpts("Main roads (₹ per m²)", ["Bituminous", "GridMats S1", "Concrete"], [O.LCC_BT - O.Upk_BT, O.S1_Capex, O.LCC_PQC - O.Upk_PQC], [O.LCC_BT, O.LCC_S1, O.LCC_PQC]);
  s.addChart(pres.charts.BAR, ch.data, Object.assign({ x: 5.4, y: 1.85, w: 4.6, h: 4.0 }, ch.opts));
  stat(s, 10.4, 2.0, 2.4, pct(O.SavePct_S2_WBM), "lower 25-year cost than WBM on O&M tracks", C.accent2, 34);
  stat(s, 10.4, 3.45, 2.4, pct(O.SavePct_S1_BT), "lower 25-year cost than bitumen on main roads", C.accent2, 34);
  stat(s, 10.4, 4.9, 2.4, pct(O.SavePct_S1_PQC), "lower 25-year cost than concrete", C.accent2, 34);
  source(s, "Model sheet 'Road LCC'. Upkeep: WBM regrading ₹40/m²/yr and regravelling every 4 yrs; bitumen BC renewal every 5 yrs; GridMats 0.5%/yr plus top-ups. 10% discount rate, 5% cost escalation.", 6.35);
  s.addNotes("The honest picture. On O&M tracks GridMats S2 costs about 2.8 times WBM to build, but WBM is regraded every monsoon and regravelled every few years; over 25 years S2 comes out cheaper. On main roads S1 costs about the same as bitumen to build and avoids renewals every five years. Against concrete, S1 is cheaper on day one and over the life.");

  // ================================================================ 10. PER 100 MW
  s = content("Economics", `${cr(O.Park_NPV)} NPV for every 100 MW block`, "One block: 10 km of internal road, 30% main roads (S1 instead of bitumen) and 70% O&M tracks (S2 instead of WBM).");
  s.addChart(pres.charts.BAR, [{ name: "₹ Cr (present value)", labels: ["Up-front GridMats premium", "Road upkeep avoided", "More energy from cleaner modules", "Carbon value"], values: [-O.Park_Premium, O.PV_Upkeep, O.PV_Energy, O.PV_Carbon] }],
    Object.assign(chartBase(), {
      x: MX, y: 1.85, w: 7.0, h: 4.5, barDir: "bar", chartColors: [HX.accent2], invertedColors: [HX.accent1],
      showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "#,##0.00;-#,##0.00", showLegend: false,
      valAxisLabelFormatCode: "#,##0", valAxisMinVal: -5, valAxisMaxVal: 7, valAxisMajorUnit: 1, catAxisOrientation: "maxMin", catAxisLabelPos: "low",
      showTitle: true, title: "Present value over 25 years, ₹ Cr per 100 MW", titleFontSize: 14, titleColor: "1A1A1A",
    }));
  const kp = [[cr(O.Park_NPV), "net present value over 25 years"], [pct(O.Park_IRR), "IRR on the GridMats premium"], [yrs(O.Park_Payback), "simple payback"], [pct(O.Park_PremPct, 1), "premium as a share of plant capex (₹4 Cr/MW)"]];
  kp.forEach((k, i) => {
    const x = 8.0 + (i % 2) * 2.4, y = 1.95 + Math.floor(i / 2) * 1.85;
    card(s, x, y, 2.25, 1.65, C.background2, "kpi-card");
    stat(s, x + 0.2, y + 0.18, 1.9, k[0], k[1], i === 0 ? C.accent2 : C.accent1, 28);
  });
  text(s, `Benefits are ${O.Park_BCR.toFixed(1)}× the premium. Road upkeep alone covers it.`, { x: 8.0, y: 5.75, w: 4.7, h: 0.6, fontSize: 14, bold: true, color: C.text1 });
  source(s, "Model sheet 'Solar Park 100MW', base case. Energy valued at ₹2.70/kWh (solar PPA level); carbon at ₹700/t CO2e.", 6.45);
  s.addNotes("The bridge from premium to value for a 100 MW block. About 3.5 crore more up front, under 1% of a 400 crore plant. Road upkeep avoided alone is worth about 5.7 crore in present value; cleaner modules add about 1.4 crore and carbon about 0.5 crore. Net: about 4.1 crore per 100 MW, paying back in under five years.");

  // ================================================================ 11. STRESS TEST
  s = content("Economics", "Positive NPV in every stress test", "Each bar moves one driver to a pessimistic and an optimistic value; everything else stays at base.");
  const tor = [...D.tornado].sort((a, b) => b.swing - a.swing);
  const lab = (t) => t.driver.replace(" (± 10%)", " ±10%").replace(" (− / + 50%)", " ±50%").replace(" (+ / − 50%)", " ±50%");
  s.addChart(pres.charts.BAR, [
    { name: "Pessimistic", labels: tor.map(lab), values: tor.map((t) => Math.min(t.npv_lo, t.npv_hi)) },
    { name: "Optimistic", labels: tor.map(lab), values: tor.map((t) => Math.max(t.npv_lo, t.npv_hi)) },
  ], Object.assign(chartBase(), {
    x: MX, y: 1.85, w: 7.6, h: 4.6, barDir: "bar", barGrouping: "clustered", barGapWidthPct: 50, chartColors: ["9A9A9A", HX.accent2],
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "#,##0.0", catAxisOrientation: "maxMin",
    valAxisMinVal: 0, valAxisMaxVal: 8, valAxisMajorUnit: 2, valAxisLabelFormatCode: "#,##0", showLegend: true, legendPos: "b",
    catAxisLabelFontSize: 11, showTitle: true, title: "NPV per 100 MW block, ₹ Cr", titleFontSize: 14, titleColor: "1A1A1A",
  }));
  const scn = [["Conservative", O.Scn_Cons, C.text2], ["Base", O.Scn_Base, C.accent1], ["Optimistic", O.Scn_Opt, C.accent2], ["Roads only, no soiling benefit", O.Scn_RoadsOnly, C.text2]];
  scn.forEach((k, i) => {
    const y = 1.95 + i * 1.12;
    card(s, 8.6, y, 4.13, 0.98, C.background2, "scn-card");
    text(s, cr(k[1]), { x: 8.8, y: y + 0.1, w: 1.8, h: 0.78, fontSize: 26, bold: true, color: k[2], valign: "middle" });
    text(s, k[0], { x: 10.6, y: y + 0.1, w: 2.0, h: 0.78, fontSize: 13, color: C.text1, valign: "middle" });
  });
  source(s, "Model sheet 'Sensitivity'. The lowest result, ₹1.7 Cr, assumes WBM upkeep at half the base estimate.", 6.5);
  s.addNotes("Robustness. The biggest swing is how much WBM really costs to keep up; even at half the base estimate, NPV stays positive at about 1.7 crore per block. With no soiling benefit at all, roads alone still return about 2.3 crore per block. Discount rate 8 to 12% moves NPV between about 2.9 and 5.7 crore.");

  // ================================================================ 12. DUST & SOILING
  s = content("Economics", "Cleaner tracks, cleaner modules", "Traffic on unpaved tracks throws dust on the nearest rows; chips held in GridMats cells throw far less.");
  function dustBand(y, label, roadColor, dots, labelColor) {
    text(s, label, { x: MX, y, w: 6.3, h: 0.35, fontSize: 14, bold: true, color: labelColor });
    s.addShape(pres.shapes.RECTANGLE, { x: MX, y: y + 0.45, w: 0.9, h: 1.55, fill: { color: roadColor }, line: { color: roadColor, width: 0 }, objectName: "track" });
    [1.8, 3.3, 4.8].forEach((px, k) => {
      s.addShape(pres.shapes.PARALLELOGRAM, { x: MX + px, y: y + 0.75, w: 1.15, h: 0.95, fill: { color: "2B3A55" }, line: { color: "FFFFFF", width: 1 }, objectName: "module-row" });
      text(s, ["10 m", "40 m", "100 m"][k], { x: MX + px, y: y + 1.75, w: 1.15, h: 0.28, fontSize: 11, color: C.accent4, align: "center" });
    });
    dots.forEach(([dx, dy, d]) => s.addShape(pres.shapes.OVAL, { x: MX + dx, y: y + dy, w: d, h: d, fill: { color: "B8A68A", transparency: 25 }, line: { color: "B8A68A", width: 0 }, objectName: "dust" }));
  }
  dustBand(1.85, "WBM / gravel track", "B8B1A3", [[0.9, 0.7, 0.32], [1.2, 1.05, 0.26], [1.5, 0.6, 0.22], [1.95, 0.85, 0.24], [2.3, 1.2, 0.2], [2.7, 0.7, 0.18], [3.4, 0.95, 0.15], [4.1, 0.8, 0.12], [1.05, 1.5, 0.22], [1.65, 1.35, 0.18]], C.text1);
  dustBand(4.15, "GridMats S2 track", "C22935", [[0.95, 0.85, 0.16], [1.35, 1.2, 0.12], [2.0, 0.95, 0.1]], C.accent1);
  const ds = [[`${inr(D.mwh_y1)} MWh`, "more energy in year 1 per 100 MW (base assumption)", C.accent3], [`₹${inr(D.energy_y1_lakh, 1)} L`, "a year at ₹2.70/kWh", C.accent2], [`${inr(O.Sus_Water)} kL`, "of dust-suppression water avoided per 100 MW", C.accent1]];
  ds.forEach((k, i) => stat(s, 7.45, 1.95 + i * 1.18, 2.6, k[0], k[1], k[2], 26));
  card(s, 10.25, 1.95, 2.48, 3.35, C.background2, "measure-card");
  text(s, [
    { text: "Base assumption", options: { bold: true, breakLine: true } },
    { text: "3% yearly soiling loss; 15% of it from track dust; GridMats cuts that dust by 70%.", options: { breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "How the pilot measures it", options: { bold: true, breakLine: true } },
    { text: "Dust gauges and soiling reference cells on rows 10-100 m from the GridMats track and from a WBM control track." },
  ], { x: 10.4, y: 2.1, w: 2.2, h: 3.1, fontSize: 13, color: C.text1, paraSpaceAfter: 2 });
  source(s, "Assumption informed by soiling studies (Gandhinagar 0.37%/day) and Univ. of Johannesburg (2025): soiling highest next to a gravel road, falling with distance.", 6.5);
  s.addNotes("This is the benefit unique to solar parks, and the one we are most careful about: it is an assumption until measured. Base case: of a 3% yearly soiling loss, 15% comes from track dust, and GridMats cuts that dust by 70%, so about 0.3% more energy, roughly 590 MWh a year per 100 MW. The pilot design puts reference cells and dust gauges at matching distances from a GridMats track and a WBM control track.");

  // ================================================================ 13. WALKWAY ECONOMICS
  s = content("Economics", "Walkways that pay back in 5-8 years", "A generating walkway costs more per kWp than a roof, but it replaces a walkway you would build anyway.");
  const wk = [[`${O.Wk_SP_Mum.toFixed(1)} yrs`, "payback in Navi Mumbai (grid power about ₹12/kWh)", C.accent2], [`${O.Wk_SP_Jam.toFixed(1)} yrs`, "payback at Jamnagar (grid power about ₹7.5/kWh)", C.accent1], [pct(O.Wk_IRR), "pre-tax IRR over 20 years, Navi Mumbai", C.accent2], [`₹${inr(O.Wk_Save1)}`, "electricity bill saved per m² in year 1, Navi Mumbai", C.accent3]];
  wk.forEach((k, i) => {
    const x = MX + (i % 2) * 3.15, y = 1.95 + Math.floor(i / 2) * 2.0;
    card(s, x, y, 2.95, 1.8, C.background2, "wk-card");
    stat(s, x + 0.22, y + 0.18, 2.55, k[0], k[1], k[2], 30);
  });
  s.addChart(pres.charts.BAR, [{ name: "₹ per kWp", labels: ["C&I rooftop", "Solar carport", "GridMats Solar Walkway"], values: [40000, 50000, O.WkPerKWp] }],
    Object.assign(chartBase(), {
      x: 7.2, y: 1.85, w: 5.53, h: 3.3, barDir: "bar", chartColors: [HX.accent3], showValue: true, dataLabelPosition: "outEnd",
      dataLabelFormatCode: "₹#,##0", valAxisHidden: true, valGridLine: { style: "none" }, showLegend: false, catAxisOrientation: "maxMin",
      valAxisMinVal: 0, valAxisMaxVal: 90000, showTitle: true, title: "Capex per kWp, solar part only", titleFontSize: 14, titleColor: "1A1A1A",
    }));
  text(s, [
    { text: "Why a walkway: ", options: { bold: true } },
    { text: `no roof, land or mounting; a 50 m² pilot is about ${inr(O.PW_kWp)} kWp and ${inr(O.PW_kWh)} kWh a year. Across Reliance campuses and Jio-bp EV hubs (${inr(O.WkScale_Area)} m², illustrative), about ${inr(O.WkScale_MWp, 1)} MWp and ₹${inr(O.WkScale_Cr, 1)} Cr a year.` },
  ], { x: 7.2, y: 5.3, w: 5.53, h: 1.1, fontSize: 13 });
  source(s, "Model sheet 'Solar Walkway'. PV layer ₹9,500/m² (GridMats pilot quote); yield 1,150-1,300 kWh/kWp for flat walk-on modules (assumption); 20-year life; paver walkway ₹1,100/m² avoided.", 6.5);
  s.addNotes("Walkway economics depend on the price of the power it replaces. At commercial tariffs in Mumbai, payback is under six years; at Jamnagar's lower industrial tariff, about eight. The per-kWp cost is higher than rooftop, so position it where roofs are full or shaded and the walkway is needed anyway, and for visibility at headquarters, Jio World Centre, data-centre campuses and EV hubs.");

  // ================================================================ 14. SUSTAINABILITY
  s = content("Economics", "Circular, low-carbon, water-wise", "Per 100 MW block over 25 years, base case.");
  const sus = [
    ["FaRecycle", `${inr(O.Sus_PP)} t`, "of post-consumer polypropylene turned into road"],
    ["FaLeaf", `${inr(Math.round(O.Sus_CO2))} t`, "CO2e avoided: plastic reuse, bitumen avoided, cleaner energy"],
    ["FaBolt", `${inr(Math.round(O.Park_MWh))} MWh`, "of extra clean energy from cleaner modules"],
    ["FaTint", `${inr(O.Sus_Water)} kL`, "of dust-suppression water avoided in the build years"],
    ["FaCloudRain", `${inr(O.Sus_Rain)} m³`, "of rain returned to the ground every year on S2 tracks"],
    ["FaIndustry", `${inr(Math.round(O.Sus_Cement))} t`, "of cement avoided if main roads would otherwise be concrete"],
  ];
  for (let i = 0; i < sus.length; i++) {
    const col = i % 3, row = Math.floor(i / 3);
    const x = MX + col * 4.1, y = 1.95 + row * 1.75;
    await hexIcon(s, sus[i][0], x, y + 0.1, 0.72, i < 3 ? C.accent2 : C.accent1);
    text(s, sus[i][1], { x: x + 0.9, y, w: 3.0, h: 0.55, fontSize: 28, bold: true, color: C.text1, valign: "bottom" });
    text(s, sus[i][2], { x: x + 0.9, y: y + 0.6, w: 3.0, h: 0.8, fontSize: 13, color: C.text2 });
  }
  card(s, MX, 5.55, CW, 0.9, C.accent5, "circ-card");
  text(s, [{ text: "Fits Reliance's circularity story: ", options: { bold: true, color: C.accent1 } }, { text: "GridMats complements ReRoute. ReRoute puts plastic into bituminous roads; GridMats goes where a road must be permeable, cement-free or open the same day." }],
    { x: MX + 0.3, y: 5.65, w: CW - 0.6, h: 0.7, fontSize: 14, valign: "middle" });
  source(s, "Model sheet 'Sustainability'. Factors: recycled PP 1.2 kg CO2e/kg; cement 0.65 t/t; bitumen 0.55 t/t; CEA grid factor 0.710 t/MWh; Kutch rainfall 400 mm.", 6.55);
  s.addNotes("Sustainability numbers that can go straight into BRSR-style reporting, with the method shown in the model. Most of the CO2e comes from the extra clean energy, so it depends on the soiling assumption. Position GridMats alongside Reliance's own ReRoute, not against it.");

  // ================================================================ 15. SCALE
  s = content("Economics", "At Kutch scale, it adds up", "Base-case results scaled from one 100 MW block. Kutch's planned output implies roughly 18-20 GW of solar.");
  const sc = D.scale;
  const hdr = ["", "100 MW", "1 GW", "5 GW", "20 GW"].map((t, i) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: i === 0 ? "3F3F3F" : "C22935" }, align: i ? "right" : "left" } }));
  const scRows = [
    ["Internal road (km)", sc["Internal road (km)"], (v) => inr(v)],
    ["GridMats area (lakh m²)", sc["GridMats area (m²)"], (v) => inr(v / 1e5, 1)],
    ["Up-front premium (₹ Cr)", sc["Up-front GridMats premium (₹ Cr)"], (v) => inr(v, 1)],
    ["NPV to Reliance, 25 years (₹ Cr)", sc["NPV to Reliance over 25 years (₹ Cr)"], (v) => inr(v, 1)],
    ["Road upkeep avoided, undiscounted (₹ Cr)", sc["Road upkeep avoided, undiscounted (₹ Cr)"], (v) => inr(v, 0)],
    ["Recycled PP put to use (t)", sc["Recycled PP put to use (t)"], (v) => inr(v)],
    ["CO2e avoided, 25 years (t)", sc["CO2e avoided over 25 years (t)"], (v) => inr(Math.round(v / 100) * 100)],
  ];
  const tbl = [hdr].concat(scRows.map(([l, vals, f], ri) => [{ text: l, options: { bold: ri === 3 } }].concat(vals.map((v) => ({ text: f(v), options: { align: "right", bold: ri === 3, color: ri === 3 ? "1F7A3A" : "1A1A1A" } })))));
  s.addTable(tbl, { x: MX, y: 1.9, w: 8.2, colW: [3.4, 1.2, 1.2, 1.2, 1.2], fontSize: 14, fontFace: "Arial", color: "1A1A1A", rowH: 0.5, border: { type: "solid", pt: 0.5, color: "E3E3E3" }, fill: { color: "FFFFFF" }, valign: "middle" });
  card(s, 9.15, 1.9, 3.58, 3.5, C.background2, "scale-card");
  stat(s, 9.4, 2.1, 3.1, cr(O.Scale1GW_NPV, 0), "NPV for every GW built with GridMats roads", C.accent2, 34);
  stat(s, 9.4, 3.55, 3.1, `${inr(O.Scale1GW_PP)} t`, "of recycled PP per GW: a real outlet for post-consumer polypropylene", C.accent1, 30);
  source(s, "Model sheet 'Scale-up'. Indicative: linear scaling of one block; supply would run on a 12-month park rate contract.", 6.5);
  s.addNotes("Scale view. Per GW, about 100 km of internal road, about 41 crore of NPV and over 3,000 tonnes of recycled PP. The 20 GW column is indicative only; supply at that scale would run under a park-wide rate contract with phased deliveries.");

  // ================================================================ 16. PILOT SCOPE
  pres.addSection({ title: "Pilot" });
  s = content("Pilot", "The pilot: 1 km of road + 50 m² of solar", "Built this dry season in a Kutch block, measured through the 2027 monsoon, decided on Reliance's own data.");
  const comp = [
    ["FaRoad", "A  Solar-park road, 1 km", `0.4 km × 6 m main road in S1 and 0.6 km × 4.5 m O&M track in S2 (${inr(O.Pilot_RoadArea)} m²), beside WBM and bituminous control sections built by Reliance's EPC.`],
    ["FaSolarPanel", "B  Solar Walkway, 50 m²", `About ${inr(O.PW_kWp)} kWp at Reliance Corporate Park, Navi Mumbai (or the Jamnagar Giga Complex), metered.`],
    ["FaMicrochip", "C  Measurement, 12 months", "IoT sensors and live dashboard, dust gauges, soiling reference cells, joint reports at 6 and 12 months."],
  ];
  for (let i = 0; i < comp.length; i++) {
    const y = 1.9 + i * 1.5;
    await hexIcon(s, comp[i][0], MX, y, 0.62, i === 1 ? C.accent3 : C.accent1);
    text(s, comp[i][1], { x: MX + 0.82, y: y - 0.03, w: 4.9, h: 0.34, fontSize: 15, bold: true });
    text(s, comp[i][2], { x: MX + 0.82, y: y + 0.34, w: 4.9, h: 1.05, fontSize: 13, color: C.text2 });
  }
  // Gantt
  const gx = 6.55, gw = 6.18, months = ["Oct", "Nov", "Dec", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const mw = gw / months.length;
  text(s, "2026", { x: gx, y: 1.85, w: mw * 3, h: 0.28, fontSize: 11, bold: true, color: C.accent4 });
  text(s, "2027", { x: gx + mw * 3, y: 1.85, w: mw * 4, h: 0.28, fontSize: 11, bold: true, color: C.accent4 });
  months.forEach((m, i) => text(s, m, { x: gx + i * mw, y: 2.12, w: mw, h: 0.28, fontSize: 10, color: C.accent4, align: "center" }));
  const tasks = [
    ["Site walk, CBR and axle loads, IITPAVE design", 0, 1.2, C.text2],
    ["QAP, test reports, factory visit; PO", 0.6, 1.4, C.text2],
    ["Supply; dispatch date fixed in the PO", 1.4, 2.4, C.accent4],
    [`Build road (about ${O.Pilot_Days} crew-days) + sensors`, 2.1, 3.0, C.accent1],
    ["Solar Walkway built and metered", 2.6, 3.6, C.accent3],
    ["Monitoring through the 2027 monsoon", 3.0, 15, C.accent2],
    ["6-month joint report", 8.7, 9.2, C.accent1],
    ["12-month report and rate-contract decision", 14.2, 15, C.accent1],
  ];
  tasks.forEach(([t, a, b, col], i) => {
    const y = 2.5 + i * 0.47;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: gx + a * mw, y, w: Math.max((b - a) * mw, 0.14), h: 0.3, rectRadius: 0.06, fill: { color: col }, line: { color: col, width: 0 }, objectName: "gantt-bar" });
    const inside = (b - a) * mw > 3.2;
    const labelRight = b * mw < gw - 2.8;
    if (inside) text(s, t, { x: gx + a * mw + 0.12, y: y - 0.02, w: (b - a) * mw - 0.2, h: 0.34, fontSize: 11, color: C.background1, bold: true, valign: "middle" });
    else text(s, t, { x: labelRight ? gx + b * mw + 0.08 : gx + a * mw - 3.6, y: y - 0.02, w: 3.5, h: 0.34, fontSize: 11, color: C.text1, valign: "middle", align: labelRight ? "left" : "right" });
  });
  source(s, "Indicative dates; supply date is confirmed in the PO once the site is ready. Control sections are built by Reliance's EPC as it would anyway.", 6.5);
  s.addNotes("Three parts. A: one kilometre of road in a block being built this season, half main road and half O&M track, beside conventional control sections. B: a 50 m² Solar Walkway where leadership walks past it every day. C: a year of joint measurement, through one monsoon, with a rate-contract decision at the end.");

  // ================================================================ 17. OFFER
  s = content("Pilot", "An offer built to say yes to", "List value of the pilot scope against what Reliance pays, ex-GST. Civil works are by Reliance's EPC at its own rates.");
  const it = D.pilot_items;
  const short = (t) => t.replace("GridMats OC-HD 40, main road S1 (0.4 km × 6 m)", "OC-HD 40 grid, S1 main road").replace("GridMats OC-HD 40, O&M track S2 (0.6 km × 4.5 m)", "OC-HD 40 grid, S2 O&M track")
    .replace("Site design (IITPAVE on measured CBR), BoQ, method statement, hold-point QC", "Site design, BoQ, method statement, QC").replace("On-site GridMats supervisor for the whole build", "Supervisor for the whole build")
    .replace("GridMats base, OC-LW 40 on open-graded GSB", "Walkway base, OC-LW 40").replace("Walkable PV layer + inverter, cabling, metering", "Walkway PV layer + inverter")
    .replace("IoT sensors, live dashboard, soiling study, 6- and 12-month joint reports", "IoT, soiling study, joint reports").replace("Geotextile 200 gsm", "Geotextile, U-nails, freight");
  const merged = [];
  for (const r of it) {
    if (/^Total/.test(r.item)) continue;
    if (/U-nails|Freight to site/.test(r.item)) { const g = merged.find((m) => /Geotextile/.test(m.item)); g.list += r.list; g.offer += r.offer; continue; }
    merged.push(Object.assign({}, r));
  }
  const oh = ["Scope", "List", "Reliance pays"].map((t, i) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: i ? "C22935" : "3F3F3F" }, align: i ? "right" : "left" } }));
  const orows = merged.map((r) => [{ text: short(r.item) }, { text: lakh(r.list, 2), options: { align: "right" } },
    { text: r.offer === 0 ? "Included" : lakh(r.offer, 2), options: { align: "right", bold: r.offer === 0, color: r.offer === 0 ? "1F7A3A" : "1A1A1A" } }]);
  orows.push([{ text: "Total", options: { bold: true } }, { text: lakh(O.Pilot_List, 2), options: { align: "right", bold: true } }, { text: lakh(O.Pilot_Offer, 2), options: { align: "right", bold: true, color: "C22935" } }]);
  s.addTable([oh].concat(orows), { x: MX, y: 1.85, w: 6.4, colW: [3.5, 1.3, 1.6], fontSize: 12, fontFace: "Arial", color: "1A1A1A", rowH: 0.36, border: { type: "solid", pt: 0.5, color: "E3E3E3" }, valign: "middle" });
  const perks = [
    ["FaRupeeSign", `₹900/m² grid: ${pct(O.Pilot_vsTrial)} below the 2025 trial`, "the annual-contract rate, at pilot size"],
    ["FaHandshake", `${lakh(O.Pilot_Contrib)} of scope included (${pct(O.Pilot_ContribPct)})`, "design, full-build supervision and travel, monitoring, half the walkway PV"],
    ["FaCalendarCheck", "₹900/m² held for 12 months", "on a park rate contract; 3% year-end rebate above 50,000 m²"],
    ["FaCheckCircle", `${lakh(O.Pilot_Credit)} credited back`, "10% of pilot grid value off the first order of 25,000 m² or more"],
    ["FaShieldAlt", `Pilot road saves ${lakh(O.Pilot_RoadNPV)} (PV)`, "over 25 years against building the same 1 km conventionally"],
  ];
  for (let i = 0; i < perks.length; i++) {
    const y = 1.85 + i * 0.93;
    await hexIcon(s, perks[i][0], 7.35, y, 0.56, i === 4 ? C.accent2 : C.accent1);
    text(s, perks[i][1], { x: 8.1, y: y - 0.02, w: 4.63, h: 0.34, fontSize: 14, bold: true });
    text(s, perks[i][2], { x: 8.1, y: y + 0.32, w: 4.63, h: 0.5, fontSize: 12, color: C.text2 });
  }
  source(s, `Civil works by Reliance's EPC, indicative ${lakh(O.Pilot_Civil)}; total pilot budget about ${lakh(O.Pilot_Budget)}. 12-month defect liability with joint inspections at 6 and 12 months. GST extra.`, 6.5);
  s.addNotes("The offer. Reliance pays about 55 lakh for 68.5 lakh of scope. The grid is at the annual-contract rate, 23% below what Reliance paid in 2025. Design, full-build supervision with travel, a year of monitoring and half the walkway PV are on GridMats. If the pilot meets its measures, the rate is held for a 12-month park contract and 10% of the pilot grid value comes back as a credit.");

  // ================================================================ 18. SUCCESS
  s = content("Pilot", "How we will judge success", "Agreed up front, measured jointly, reported at 6 and 12 months against the control sections.");
  const kh = ["Measure", "Success looks like"].map((t, i) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: i ? "C22935" : "3F3F3F" } } }));
  const krows = [
    ["Rut depth", "≤ 10 mm at 12 months"], ["Exposed cell tops", "≤ 5% of area"], ["Drainage", "No standing water 1 hour after rain"],
    ["Cracking", "None crossing more than 3 cells"], ["Dust and soiling", "Lower dust deposition and soiling on rows near GridMats than near WBM"],
    ["Upkeep", "Fewer maintenance events and lower cost than control sections"], ["Solar Walkway", "Measured W/m², kWh/m²/day, module heat, grip, uptime"],
  ].map(([a, b]) => [{ text: a, options: { bold: true } }, { text: b }]);
  s.addTable([kh].concat(krows), { x: MX, y: 1.85, w: 7.2, colW: [2.2, 5.0], fontSize: 13, fontFace: "Arial", color: "1A1A1A", rowH: 0.52, border: { type: "solid", pt: 0.5, color: "E3E3E3" }, valign: "middle" });
  card(s, 8.2, 1.85, 4.53, 4.05, C.text1, "dash-card");
  await hexIcon(s, "FaMicrochip", 8.45, 2.05, 0.62, C.accent1);
  text(s, "Live pilot dashboard", { x: 9.25, y: 2.1, w: 3.3, h: 0.5, fontSize: 18, bold: true, color: C.background1, valign: "middle" });
  text(s, ["Temperature, humidity and subgrade moisture", "Vibration from passing trucks", "Rainfall and flood alerts", "Site map with every sensor", "Readings Reliance can export at any time"].map((t, j, a) => ({ text: t, options: { bullet: true, breakLine: j < a.length - 1 } })),
    { x: 8.45, y: 2.9, w: 4.1, h: 2.4, fontSize: 14, color: C.background1, paraSpaceAfter: 6 });
  text(s, "GridMats Pavement Monitor, already running on GridMats sites.", { x: 8.45, y: 5.1, w: 4.1, h: 0.6, fontSize: 12, italic: true, color: C.background2 });
  source(s, "Thresholds from the GridMats Systems Manual maintenance limits. Dashboard: GridMats Pavement Monitor (temperature, humidity, moisture, CO2, vibration, rainfall, light, flood).", 6.5);
  s.addNotes("Success is defined before we start, using GridMats' own maintenance thresholds, and judged against Reliance's control sections. The live dashboard shows the sensor feed from the road; Reliance can export the data at any time.");

  // ================================================================ 19. LESSONS
  s = content("Pilot", "Lessons from 2025, built in", "What we changed so the pilot runs to plan.");
  const lessons = [
    ["FaCalendarCheck", "Supply on time", "QAP, test reports and a factory-inspection slot are cleared before the PO, and the dispatch date is fixed in the PO once the site is ready."],
    ["FaCubes", "Right aggregate", "8-12 mm chip gradation agreed and sample-tested before dispatch; low fines so chips lock in and stay dust-free."],
    ["FaFlask", "Design on measured soil", "Sections from IITPAVE on the pilot block's measured CBR; WMM added where CBR is below 5%."],
    ["FaHardHat", "Supervision end to end", "A GridMats engineer on site for the whole build with hold-point QC at every layer, not just the first days."],
    ["FaEye", "Data, not anecdotes", "Control sections, sensors and joint inspections at 6 and 12 months, so the result is Reliance's own evidence."],
  ];
  for (let i = 0; i < lessons.length; i++) {
    const col = i < 3 ? 0 : 1, row = i < 3 ? i : i - 3;
    const x = MX + col * 6.2, y = 1.95 + row * 1.5;
    await hexIcon(s, lessons[i][0], x, y, 0.66);
    text(s, lessons[i][1], { x: x + 0.88, y: y - 0.02, w: 5.0, h: 0.36, fontSize: 16, bold: true });
    text(s, lessons[i][2], { x: x + 0.88, y: y + 0.36, w: 5.0, h: 0.95, fontSize: 13, color: C.text2 });
  }
  card(s, MX + 6.2, 4.95, 5.93, 1.35, C.accent5, "dlp-card");
  text(s, [{ text: "Assurance: ", options: { bold: true, color: C.accent1 } }, { text: "12-month defect liability from handover with joint inspections at 6 and 12 months; repairs are cell by cell, with no traffic diversion." }],
    { x: MX + 6.45, y: 5.05, w: 5.45, h: 1.15, fontSize: 14, valign: "middle" });
  s.addNotes("Name the 2025 learnings without dwelling on them: supply timing, aggregate quality, design for very weak soil, and supervision length. Each now has a specific fix in the pilot plan.");

  // ================================================================ 20. NEXT STEPS
  pres.addSection({ title: "Close" });
  s = pres.addSlide({ masterName: "CLOSING_DARK", sectionTitle: "Close" });
  s.addText("Next steps", { placeholder: "title" });
  s.addText("Five steps from this conversation to a measured result.", { placeholder: "body" });
  const steps = [
    ["1", "This month", "Name the pilot block and walkway site, and a Reliance sponsor"],
    ["2", "Oct-Nov 2026", "Site walk, CBR and axle-load data; IITPAVE design and BoQ"],
    ["3", "Nov 2026", "QAP and tests cleared; pilot PO with the dispatch date fixed"],
    ["4", "Dec 2026-Jan 2027", "Build 1 km of road and the 50 m² Solar Walkway; sensors live"],
    ["5", "Jun and Dec 2027", "Joint reports; decide the park-wide rate contract"],
  ];
  for (let i = 0; i < steps.length; i++) {
    const x = MX + i * 2.45;
    s.addShape(pres.shapes.HEXAGON, { x, y: 2.35, w: 0.9, h: 0.79, fill: { color: i === 4 ? C.accent2 : C.accent1 }, line: { color: C.text1, width: 0 }, objectName: "step-hex" });
    text(s, steps[i][0], { x, y: 2.35, w: 0.9, h: 0.79, fontSize: 24, bold: true, color: C.background1, align: "center", valign: "middle" });
    text(s, steps[i][1], { x, y: 3.35, w: 2.2, h: 0.4, fontSize: 15, bold: true, color: C.background1 });
    text(s, steps[i][2], { x, y: 3.8, w: 2.2, h: 1.5, fontSize: 14, color: C.background2 });
  }
  text(s, [
    { text: "GridMats  ·  GroundReality Enterprises Pvt Ltd, Bengaluru", options: { bold: true, breakLine: true } },
    { text: "sales@gridmats.com  ·  +91 63626 46287  ·  www.gridmats.com" },
  ], { x: MX, y: 5.7, w: 9, h: 0.8, fontSize: 14, color: C.background1 });
  s.addNotes("Close on a date: name the block and the sponsor this month so that the road is built in this dry season and measured through the 2027 monsoon.");

  // ================================================================ 21. APPENDIX
  pres.addSection({ title: "Appendix" });
  s = content("Appendix", "Key assumptions and sources", "All inputs are editable in the companion model (sheet 'Inputs'). Money ex-GST, 2026 prices.");
  const ah = ["Assumption", "Base value", "Source / basis"].map((t, i) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: i === 1 ? "C22935" : "3F3F3F" } } }));
  const arows = [
    ["Discount rate; cost escalation", "10%; 5% a year", "Assumption; 8-12% tested"],
    ["Road per 100 MW; main share", "10 km; 30%", "GridMats parks flier (8-12 km)"],
    ["WBM; bitumen; concrete build", "₹600; ₹2,000; ₹3,200 /m²", "TN G.O.; CPWD DSR 2023; GridMats pitch"],
    ["GridMats S1; S2 installed", `₹${inr(O.S1_Capex)}; ₹${inr(O.S2_Capex)} /m²`, "Bottom-up, Design Engine 73 rates, grid ₹900"],
    ["WBM upkeep", "₹40/m²/yr + ₹250 every 4 yrs", "Assumption; TRL/World Bank"],
    ["Bitumen upkeep", "₹20/m²/yr + ₹450 every 5 yrs", "IRC:82-2015"],
    ["Yield; energy value", "1,870 kWh/kWp; ₹2.70/kWh", "Gujarat plant data; SECI/NTPC PPAs"],
    ["Soiling; track-dust share; cut", "3%; 15%; 70%", "Assumption; measured in the pilot"],
    ["Carbon value; grid factor", "₹700/t; 0.710 t/MWh", "India CCTS range; CEA v21.0"],
    ["Walkway power; PV price; yield", "160 W/m²; ₹9,500/m²; 1,150-1,300", "GridMats flier and quote; assumption"],
  ].map((r) => r.map((t, j) => ({ text: t, options: { bold: j === 0 } })));
  s.addTable([ah].concat(arows), { x: MX, y: 1.8, w: CW, colW: [4.0, 3.7, 4.43], fontSize: 12, fontFace: "Arial", color: "1A1A1A", rowH: 0.41, border: { type: "solid", pt: 0.5, color: "E3E3E3" }, valign: "middle" });
  source(s, "Full source list on the model's 'Sources' sheet: GridMats master data and price ladder (Sep 2026), Systems Manual, Design Engine 73, Reliance PO/WO (2025), RIL AGM 2026 coverage, CEA, GERC/MERC, SECI, IRC.", 6.45);
  s.addNotes("Reference slide for questions on assumptions. Every value is an input in the model and can be changed live.");

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
}
main().catch((e) => { console.error(e); process.exit(1); });
