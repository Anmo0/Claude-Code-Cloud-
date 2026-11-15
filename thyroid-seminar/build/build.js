// Thyroid Tumors seminar — built to match the "Acute Scrotum" template style.
const pptxgen = require("pptxgenjs");
const path = require("path");
const SKILL = "/root/.claude/skills/synced/730feb66-4d48-4bb5-acac-3abb72bcc0cb_b7742437-68e3-449d-8f17-9a7d66d0088d/pptx/scripts/apply_theme.js";
const { applyTheme } = require(SKILL);

// ---------- palette (from the template) ----------
const NAVY = "0C2A45", TEAL = "13736B", RED = "C84B41", MINT = "7ECFC4", BG = "F4F6FA",
  WHITE = "FFFFFF", BORDER = "D9DEE6", BODY = "2E3B4E", MUTED = "5A6678", YELLOW = "F3C04F",
  PINK = "F2B5AE", TEAL_T = "E6F2F0", RED_T = "FBECEA", ORANGE = "B65A12", ROWALT = "EEF1F6";
const HEAD = "Georgia", BODYF = "Calibri";

const THEME = {
  name: "Thyroid Seminar", headFontFace: HEAD, bodyFontFace: BODYF,
  colors: { dk1: NAVY, lt1: WHITE, dk2: BODY, lt2: BG, accent1: TEAL, accent2: RED, accent3: MINT,
    accent4: YELLOW, accent5: ORANGE, accent6: MUTED, hlink: TEAL, folHlink: MUTED },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.theme = { headFontFace: HEAD, bodyFontFace: BODYF };
pres.title = "Thyroid Tumors";
pres.subject = "Surgery seminar";

const W = 13.333, X0 = 0.6, CW = 12.133;

// ---------- layouts ----------
const circles = () => [
  { ellipse: { x: 9.2, y: 0.7, w: 4.6, h: 4.6, fill: { type: "none" }, line: { color: "1F6A6B", width: 1.75 } } },
  { ellipse: { x: 8.3, y: 3.5, w: 3.0, h: 3.0, fill: { type: "none" }, line: { color: "1F6A6B", width: 1.75 } } },
  { ellipse: { x: 10.4, y: 2.6, w: 3.4, h: 3.4, fill: { type: "none" }, line: { color: "B5473F", width: 1.75 } } },
];
pres.defineSlideMaster({
  title: "DARK", background: { color: NAVY },
  slideNumber: { x: 12.3, y: 7.0, w: 0.6, h: 0.3, fontSize: 9, color: "8FA3B8", align: "right" },
});
pres.defineSlideMaster({
  title: "CONTENT", background: { color: BG },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: X0, y: 0.55, w: CW, h: 0.85, fontFace: HEAD, fontSize: 28, bold: true, color: NAVY, margin: 0, valign: "middle", align: "left" }, text: "" } },
  ],
  slideNumber: { x: 12.3, y: 7.0, w: 0.6, h: 0.3, fontSize: 9, color: MUTED, align: "right" },
});

// ---------- helpers ----------
function addCircles(s) {
  s.addShape(pres.shapes.OVAL, { x: 9.2, y: 0.7, w: 4.6, h: 4.6, fill: { type: "none" }, line: { color: "1F6A6B", width: 1.75 }, objectName: "circle1" });
  s.addShape(pres.shapes.OVAL, { x: 8.3, y: 3.5, w: 3.0, h: 3.0, fill: { type: "none" }, line: { color: "1F6A6B", width: 1.75 }, objectName: "circle2" });
  s.addShape(pres.shapes.OVAL, { x: 10.4, y: 2.6, w: 3.4, h: 3.4, fill: { type: "none" }, line: { color: "B5473F", width: 1.75 }, objectName: "circle3" });
}
const SECTIONS = ["Introduction and Overview", "Anatomy and Physiology", "Classification of Thyroid Tumors",
  "Pathophysiology and Natural History", "Clinical Presentation", "Diagnostic Workup", "Management Principles",
  "Prognosis and Outcomes", "Case Discussions"];
function divider(secIdx) {
  const sec = SECTIONS[secIdx];
  pres.addSection({ title: sec });
  const s = pres.addSlide({ masterName: "DARK", sectionTitle: sec });
  addCircles(s);
  s.addText(String(secIdx + 1).padStart(2, "0"), { x: 0.85, y: 1.75, w: 3, h: 1.3, fontSize: 80, bold: true, color: MINT, margin: 0, valign: "bottom", isTextBox: true });
  s.addText(sec, { x: 0.88, y: 3.2, w: 7.0, h: 1.75, fontFace: HEAD, fontSize: 44, bold: true, color: WHITE, margin: 0, valign: "top", isTextBox: true });
  for (let i = 0; i < 9; i++) {
    s.addShape(pres.shapes.OVAL, { x: 0.9 + i * 0.32, y: 6.32, w: 0.17, h: 0.17, fill: { color: i === secIdx ? MINT : NAVY }, line: { color: i === secIdx ? MINT : "6F8499", width: 1 }, objectName: `dot${i}` });
  }
  s.addNotes(`Section ${secIdx + 1}: ${sec}.`);
}
function content(secIdx, title, sources, notes) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: SECTIONS[secIdx] });
  s.addText(title, { placeholder: "title" });
  if (sources) s.addText(`Sources: ${sources}`, { x: X0, y: 6.98, w: 9, h: 0.3, fontSize: 10, color: MUTED, margin: 0, isTextBox: true });
  if (notes) s.addNotes(notes);
  return s;
}
function box(s, x, y, w, h, fill = WHITE, line = BORDER) {
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { color: line, width: 0.75 } });
}
function topbar(s, x, y, w, color) { s.addShape(pres.shapes.RECTANGLE, { x, y, w, h: 0.1, fill: { color }, line: { type: "none" } }); }
function label(s, text, x, y, w, color, size = 10) {
  s.addText(text.toUpperCase(), { x, y, w, h: 0.26, fontSize: size, bold: true, color, charSpacing: 2.5, margin: 0, valign: "middle", isTextBox: true });
}
// Template card: top bar, kicker, title, rows of LABEL + text separated by hairlines.
function card(s, o) {
  const { x, y, w, h, color = TEAL, kicker, title, rows = [], titleSize = 18, textSize = 14, badge } = o;
  box(s, x, y, w, h); topbar(s, x, y, w, color);
  const px = x + 0.28, pw = w - 0.56;
  let cy = y + 0.3;
  if (kicker) { label(s, kicker, px, cy, pw, color, 10.5); cy += 0.32; }
  if (badge) {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + w - 1.95, y: y + 0.27, w: 1.7, h: 0.34, fill: { color: YELLOW }, line: { color: "E2AE36", width: 0.75 }, rectRadius: 0.17 });
    s.addText("★  " + badge, { x: x + w - 1.95, y: y + 0.27, w: 1.7, h: 0.34, fontSize: 9.5, bold: true, color: NAVY, align: "center", valign: "middle", charSpacing: 1.5, margin: 0, isTextBox: true });
  }
  if (title) { s.addText(title, { x: px, y: cy, w: pw, h: 0.42, fontSize: titleSize, bold: true, color: NAVY, margin: 0, valign: "middle", isTextBox: true }); cy += 0.55; } else cy += 0.12;
  if (!rows.length) return;
  const rh = (y + h - 0.12 - cy) / rows.length;
  rows.forEach((r, i) => {
    const ry = cy + i * rh;
    s.addShape(pres.shapes.LINE, { x: px, y: ry, w: pw, h: 0, line: { color: BORDER, width: 0.75 } });
    label(s, r.l, px, ry + 0.1, pw, color);
    s.addText(r.t, { x: px, y: ry + 0.38, w: pw, h: rh - 0.42, fontSize: r.size || textSize, bold: !!r.b, color: r.b ? NAVY : BODY, margin: 0, valign: "top", isTextBox: true, paraSpaceAfter: 2 });
  });
}
function remember(s, text, y = 6.1, lead = "Remember: ") {
  box(s, X0, y, CW, 0.72, NAVY, NAVY);
  s.addShape(pres.shapes.RECTANGLE, { x: X0, y, w: 0.1, h: 0.72, fill: { color: RED }, line: { type: "none" } });
  s.addText([{ text: lead, options: { bold: true, color: PINK } }, { text, options: { color: WHITE } }],
    { x: X0 + 0.35, y, w: CW - 0.6, h: 0.72, fontSize: 15, valign: "middle", margin: 0, isTextBox: true });
}
function bullets(items, size = 14, color = BODY) {
  return items.map((t, i) => {
    const runs = Array.isArray(t) ? t : [{ text: t }];
    return runs.map((r, j) => ({ text: r.text, options: { bullet: j === 0 ? { indent: 14 } : undefined, bold: r.b, color: r.b ? NAVY : color, fontSize: size, breakLine: j === runs.length - 1 && i < items.length - 1, paraSpaceAfter: 6 } }));
  }).flat();
}

// =====================================================================
// 1. Title
// =====================================================================
{
  pres.addSection({ title: "Opening" });
  const s = pres.addSlide({ masterName: "DARK", sectionTitle: "Opening" }); addCircles(s);
  s.addText("SURGERY  ·  SEMINAR  ·  GROUP A6", { x: 0.9, y: 1.95, w: 7.5, h: 0.35, fontSize: 12, bold: true, color: MINT, charSpacing: 5, margin: 0, isTextBox: true });
  s.addText("Thyroid Tumors", { x: 0.85, y: 2.4, w: 8.2, h: 1.25, fontFace: HEAD, fontSize: 60, bold: true, color: WHITE, margin: 0, isTextBox: true });
  s.addText("From the thyroid nodule to long-term outcome: classification, workup, management and prognosis",
    { x: 0.9, y: 3.85, w: 7.3, h: 0.8, fontSize: 18, color: "C9D6E3", margin: 0, isTextBox: true });
  s.addNotes("Title slide. Group A6 members introduce themselves and their sections.\nThe seminar follows the nine course objectives in order; each objective has its own slide.");
}

// =====================================================================
// 2. Objectives
// =====================================================================
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Opening" });
  s.addText("Our objectives", { placeholder: "title" });
  const cw = 3.9, ch = 1.5, gx = 0.215, gy = 0.22;
  SECTIONS.forEach((t, i) => {
    const c = i % 3, r = Math.floor(i / 3);
    const x = X0 + c * (cw + gx), y = 1.7 + r * (ch + gy);
    box(s, x, y, cw, ch); topbar(s, x, y, cw, TEAL);
    s.addText(String(i + 1).padStart(2, "0"), { x: x + 0.28, y: y + 0.28, w: 1.1, h: 0.95, fontSize: 40, bold: true, color: TEAL, margin: 0, valign: "middle", isTextBox: true });
    s.addText(t, { x: x + 1.35, y: y + 0.25, w: cw - 1.55, h: 1.05, fontSize: 16, bold: true, color: NAVY, margin: 0, valign: "middle", isTextBox: true });
  });
  s.addNotes("Nine objectives from the course outline. Each one is covered on its own slide(s), in this order.");
}

divider(0);
// =====================================================================
// 3. Introduction (definition + epidemiology + risk factors)
// =====================================================================
{
  const s = content(0, "Thyroid tumors: common nodules, few cancers", "[8] ATA 2015; [3] GLOBOCAN 2022; [4] Saudi Cancer Registry",
    "Definition: a thyroid tumour is a neoplasm of follicular cells (most) or parafollicular C cells; lymphoma and metastases are rare.\n" +
    "Nodules are found on ultrasound in 19–68% of adults; only 7–15% are malignant (ATA 2015).\n" +
    "GLOBOCAN 2022: about 821,000 new cases worldwide, 5th most common cancer in women, roughly 3 times more frequent in women. In Saudi Arabia it is the 2nd most common cancer in women.\n" +
    "Much of the rise in incidence reflects small papillary cancers found on imaging.\n" +
    "Risk factors: childhood neck irradiation (strongest), inherited syndromes (MEN2 → medullary; FAP, Cowden → differentiated), iodine deficiency (follicular), Hashimoto thyroiditis (lymphoma).");
  const st = [["19–68%", "adults with a nodule on US", TEAL], ["7–15%", "of nodules are malignant", RED], ["821,000", "new cases worldwide, 2022", TEAL], ["2nd", "commonest cancer in Saudi women", RED]];
  st.forEach(([n, l, c], i) => {
    const x = X0 + i * 3.078, w = 2.9;
    box(s, x, 1.65, w, 1.05);
    s.addText(n, { x: x + 0.22, y: 1.7, w: w - 0.3, h: 0.55, fontFace: HEAD, fontSize: 26, bold: true, color: c, margin: 0, valign: "middle", isTextBox: true });
    s.addText(l, { x: x + 0.22, y: 2.25, w: w - 0.3, h: 0.36, fontSize: 13, color: BODY, margin: 0, isTextBox: true });
  });
  card(s, { x: X0, y: 2.92, w: 5.95, h: 3.0, color: TEAL, kicker: "Definition · benign vs malignant", title: null, rows: [
    { l: "Benign", t: "Follicular adenoma, cysts: no invasion, no spread" },
    { l: "Malignant", t: "Invades capsule or vessels; spreads by lymph or blood" },
    { l: "Key", t: "FNA cannot separate follicular adenoma from carcinoma", b: true } ] });
  card(s, { x: X0 + 6.18, y: 2.92, w: 5.95, h: 3.0, color: RED, kicker: "Risk factors · who is at risk?", title: null, rows: [
    { l: "Radiation", t: "Childhood neck irradiation: strongest risk factor", b: true },
    { l: "Genetic", t: "MEN2 (RET) → medullary; FAP, Cowden → differentiated" },
    { l: "Background", t: "Iodine deficiency → follicular; Hashimoto → lymphoma" } ] });
  remember(s, "the rising incidence is driven mainly by small papillary cancers found on imaging", 6.12, "Note: ");
}

divider(1);
// =====================================================================
// 5. Anatomy
// =====================================================================
{
  const s = content(1, "Anatomy: blood supply and lymphatic drainage", "[1] Bailey & Love",
    "The gland lies at C5–T1, wrapped in pretracheal fascia, so it moves up on swallowing.\n" +
    "Arteries: superior thyroid (first branch of the external carotid; runs with the external branch of the superior laryngeal nerve) and inferior thyroid (thyrocervical trunk; closely related to the recurrent laryngeal nerve). A thyroidea ima artery is occasionally present.\n" +
    "Veins: superior and middle drain to the internal jugular; inferior veins drain to the brachiocephalic veins.\n" +
    "Lymphatics: first to the central compartment (level VI: prelaryngeal, pretracheal, paratracheal), then to the lateral neck (levels II–V) and the superior mediastinum (level VII). This explains where papillary and medullary cancers spread.");
  box(s, X0, 1.65, 7.35, 3.56);
  s.addImage({ path: path.join(__dirname, "anatomy.png"), x: X0 + 0.02, y: 1.67, w: 7.31, h: 3.502, altText: "Labelled anterior view of the thyroid with arteries, veins, recurrent laryngeal nerve and parathyroids" });
  s.addText("Schematic anterior view. Parathyroids and the recurrent laryngeal nerve lie posteriorly.", { x: X0, y: 5.3, w: 7.35, h: 0.3, fontSize: 11, italic: true, color: MUTED, margin: 0, isTextBox: true });
  const lv = [["Level VI", "central: pre- and paratracheal, prelaryngeal", TEAL], ["Levels II–V", "lateral neck", ORANGE], ["Level VII", "superior mediastinum", RED]];
  label(s, "Lymphatic drainage, in order", X0, 5.72, 5, NAVY, 10);
  lv.forEach(([t, d, c], i) => {
    const x = X0 + i * 2.5, w = 2.3;
    s.addShape(pres.shapes.RECTANGLE, { x, y: 6.05, w, h: 0.75, fill: { color: WHITE }, line: { color: c, width: 1.25 } });
    s.addText([{ text: t, options: { bold: true, color: c, fontSize: 14, breakLine: true } }, { text: d, options: { color: BODY, fontSize: 11 } }], { x: x + 0.12, y: 6.05, w: w - 0.2, h: 0.75, valign: "middle", margin: 0, isTextBox: true });
    if (i < 2) s.addShape(pres.shapes.LINE, { x: x + w + 0.02, y: 6.42, w: 0.16, h: 0, line: { color: NAVY, width: 1.5, endArrowType: "triangle" } });
  });
  card(s, { x: 8.2, y: 1.65, w: 4.53, h: 5.15, color: TEAL, kicker: "Surgical anatomy", title: "Supply and drainage", rows: [
    { l: "Arterial", t: "Superior thyroid a. (ECA) and inferior thyroid a. (thyrocervical trunk)" },
    { l: "Venous", t: "Superior and middle → IJV; inferior → brachiocephalic v." },
    { l: "Lymphatic", t: "Central level VI first → lateral II–V and mediastinum VII", b: true },
    { l: "At risk", t: "Recurrent and external laryngeal nerves; parathyroids" } ] });
}

// =====================================================================
// 6. Physiology
// =====================================================================
{
  const s = content(1, "Thyroid physiology and hormone regulation", "[1] Bailey & Love",
    "Hypothalamus releases TRH → anterior pituitary releases TSH → TSH acts on follicular cells to make T4 (most of the output) and some T3. T4 is converted to the active T3 in peripheral tissues. Free T4 and T3 feed back negatively on the hypothalamus and pituitary.\n" +
    "Synthesis: iodide trapping by the sodium–iodide symporter (NIS) → oxidation and organification by thyroid peroxidase onto thyroglobulin → coupling of MIT and DIT → storage in colloid → release.\n" +
    "Why this matters for tumours: NIS is why radioiodine works on follicular-cell cancers only; TSH stimulates differentiated cancer growth, hence TSH suppression; thyroglobulin is a tumour marker after total thyroidectomy; C cells secrete calcitonin, the marker for medullary cancer.");
  // axis flow
  const fx = X0, fw = 4.3;
  const steps = [["Hypothalamus", "TRH"], ["Anterior pituitary", "TSH"], ["Thyroid follicular cells", "T4 (mainly) + T3"], ["Peripheral tissues", "T4 → active T3"]];
  steps.forEach(([t, h], i) => {
    const y = 1.7 + i * 1.22;
    s.addShape(pres.shapes.RECTANGLE, { x: fx, y, w: fw, h: 0.78, fill: { color: i === 2 ? TEAL : WHITE }, line: { color: i === 2 ? TEAL : BORDER, width: 0.75 } });
    s.addText(t, { x: fx + 0.2, y, w: fw - 0.4, h: 0.78, fontSize: 16, bold: true, color: i === 2 ? WHITE : NAVY, margin: 0, valign: "middle", isTextBox: true });
    if (i < 3) {
      s.addShape(pres.shapes.LINE, { x: fx + 1.2, y: y + 0.8, w: 0, h: 0.4, line: { color: NAVY, width: 1.75, endArrowType: "triangle" } });
      s.addText(steps[i][1], { x: fx + 1.35, y: y + 0.8, w: 2.8, h: 0.4, fontSize: 13, bold: true, color: TEAL, margin: 0, valign: "middle", isTextBox: true });
    }
  });
  // negative feedback loop
  s.addShape(pres.shapes.LINE, { x: fx + fw + 0.25, y: 2.09, w: 0, h: 3.66, line: { color: RED, width: 1.75, dashType: "dash" } });
  s.addShape(pres.shapes.LINE, { x: fx + fw, y: 5.75, w: 0.25, h: 0, line: { color: RED, width: 1.75, dashType: "dash" } });
  s.addShape(pres.shapes.LINE, { x: fx + fw, y: 2.09, w: 0.25, h: 0, line: { color: RED, width: 1.75, dashType: "dash", beginArrowType: "triangle" } });
  s.addShape(pres.shapes.LINE, { x: fx + fw, y: 3.31, w: 0.25, h: 0, line: { color: RED, width: 1.75, dashType: "dash", beginArrowType: "triangle" } });
  s.addText("Negative feedback (−)", { x: fx + fw + 0.38, y: 3.8, w: 0.95, h: 0.9, fontSize: 12, bold: true, color: RED, margin: 0, valign: "middle", isTextBox: true });
  card(s, { x: 6.3, y: 1.7, w: 6.43, h: 4.15, color: TEAL, kicker: "Hormone synthesis", title: "From iodide to T4", rows: [
    { l: "1 · Trapping", t: "Iodide taken up by the sodium–iodide symporter (NIS)" },
    { l: "2 · Organification", t: "Thyroid peroxidase binds iodine to thyroglobulin; MIT + DIT couple" },
    { l: "3 · Storage and release", t: "Stored in colloid; T4 and T3 released into blood" },
    { l: "C cells", t: "Parafollicular cells secrete calcitonin" } ] });
  remember(s, "NIS → radioiodine; TSH → TSH suppression; thyroglobulin and calcitonin → tumour markers", 6.12, "Link to tumours: ");
}

divider(2);
// =====================================================================
// 7. Classification (benign + malignant tree)
// =====================================================================
{
  const s = content(2, "Classification by behaviour and cell of origin", "[5] WHO 2022; [1] Bailey & Love; [10] ATA MTC 2015",
    "Benign: follicular adenoma (including oncocytic/Hürthle adenoma and the 'hot' toxic adenoma) and cysts. WHO 2022 lists NIFTP as a low-risk neoplasm, treated by lobectomy alone. Colloid nodules and multinodular goitre are hyperplasia, not neoplasms.\n" +
    "Malignant, follicular-cell derived (about 95%): papillary (~80%), follicular (~10%), oncocytic carcinoma, poorly differentiated and anaplastic (~1–2%).\n" +
    "C-cell: medullary (~5%; about 25% hereditary, RET).\n" +
    "Other: primary lymphoma (in Hashimoto thyroiditis) and metastases (kidney, lung, breast, melanoma). Percentages are approximate.");
  s.addShape(pres.shapes.OVAL, { x: 0.45, y: 3.05, w: 2.35, h: 2.35, fill: { color: WHITE }, line: { color: "7FB3AE", width: 1.5 } });
  s.addShape(pres.shapes.OVAL, { x: 0.6, y: 3.2, w: 2.05, h: 2.05, fill: { color: NAVY }, line: { type: "none" } });
  s.addText("Thyroid\ntumours", { x: 0.6, y: 3.2, w: 2.05, h: 2.05, fontFace: HEAD, fontSize: 19, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0, isTextBox: true });
  const groups = [
    { name: "Benign", sub: "no invasion", color: TEAL, tint: TEAL_T, items: [["Follicular adenoma", "incl. oncocytic, toxic"], ["Cysts", "pure cyst: very low risk"], ["NIFTP", "low-risk neoplasm"]] },
    { name: "Follicular-cell", sub: "malignant, ~95%", color: RED, tint: RED_T, items: [["Papillary carcinoma", "~80%", true], ["Follicular carcinoma", "~10%"], ["Oncocytic carcinoma", ""], ["Poorly diff. · Anaplastic", "anaplastic ~1–2%"]] },
    { name: "C-cell", sub: "malignant", color: ORANGE, tint: "F8EDE3", items: [["Medullary carcinoma", "~5%; 25% hereditary"]] },
    { name: "Other", sub: "malignant", color: NAVY, tint: "E8ECF2", items: [["Lymphoma · Metastases", "Hashimoto; kidney, lung"]] },
  ];
  const rowH = 0.46, rg = 0.05, gg = 0.2;
  let gy = 1.62;
  const centers = [];
  groups.forEach((g) => {
    const n = g.items.length, gh = n * rowH + (n - 1) * rg, cy = gy + gh / 2;
    centers.push(cy);
    s.addShape(pres.shapes.RECTANGLE, { x: 3.3, y: gy - 0.06, w: 9.43, h: gh + 0.12, fill: { color: g.tint }, line: { color: BORDER, width: 0.5 } });
    s.addShape(pres.shapes.LINE, { x: 3.05, y: cy, w: 0.4, h: 0, line: { color: "8A97AA", width: 1 } });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 3.45, y: cy - 0.3, w: 2.3, h: 0.6, fill: { color: g.color }, line: { type: "none" }, rectRadius: 0.06 });
    s.addText([{ text: g.name, options: { bold: true, fontSize: 14, breakLine: true } }, { text: g.sub, options: { fontSize: 10.5 } }],
      { x: 3.45, y: cy - 0.3, w: 2.3, h: 0.6, color: WHITE, align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addShape(pres.shapes.LINE, { x: 5.75, y: cy, w: 0.2, h: 0, line: { color: "8A97AA", width: 1 } });
    if (n > 1) s.addShape(pres.shapes.LINE, { x: 5.95, y: gy + rowH / 2, w: 0, h: (n - 1) * (rowH + rg), line: { color: "8A97AA", width: 1 } });
    g.items.forEach(([t, p, star], i) => {
      const iy = gy + i * (rowH + rg);
      s.addShape(pres.shapes.LINE, { x: 5.95, y: iy + rowH / 2, w: 0.2, h: 0, line: { color: "8A97AA", width: 1 } });
      s.addShape(pres.shapes.RECTANGLE, { x: 6.15, y: iy, w: 6.45, h: rowH, fill: { color: WHITE }, line: { color: BORDER, width: 0.75 } });
      s.addText(t, { x: 6.32, y: iy, w: 3.4, h: rowH, fontSize: 14, bold: true, color: NAVY, valign: "middle", margin: 0, isTextBox: true });
      if (star) {
        s.addText(p, { x: 9.8, y: iy, w: 0.8, h: rowH, fontSize: 13, color: BODY, valign: "middle", margin: 0, isTextBox: true });
        s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 10.75, y: iy + 0.07, w: 1.7, h: 0.32, fill: { color: YELLOW }, line: { color: "E2AE36", width: 0.75 }, rectRadius: 0.16 });
        s.addText("★  MOST COMMON", { x: 10.75, y: iy + 0.07, w: 1.7, h: 0.32, fontSize: 9.5, bold: true, color: NAVY, align: "center", valign: "middle", charSpacing: 1.5, margin: 0, isTextBox: true });
      } else if (p) {
        s.addText(p, { x: 9.8, y: iy, w: 2.7, h: rowH, fontSize: 12.5, color: BODY, valign: "middle", margin: 0, isTextBox: true });
      }
    });
    gy += gh + gg;
  });
  s.addShape(pres.shapes.LINE, { x: 3.05, y: centers[0], w: 0, h: centers[3] - centers[0], line: { color: "8A97AA", width: 1 } });
  s.addShape(pres.shapes.LINE, { x: 2.8, y: 4.22, w: 0.25, h: 0, line: { color: "8A97AA", width: 1 } });
}

divider(3);
// =====================================================================
// 9. Patterns of spread
// =====================================================================
{
  const s = content(3, "Patterns of spread differ by tumor type", "[8] ATA 2015; [10] ATA MTC 2015; [11] ATA ATC 2021",
    "Papillary: lymphatic spread to the central (level VI) then lateral neck nodes; often multifocal; nodal metastases are common at diagnosis but have a modest effect on survival in young patients.\n" +
    "Follicular: blood-borne spread to lung and bone; nodal metastases are uncommon. Diagnosis rests on capsular and vascular invasion.\n" +
    "Oncocytic carcinoma behaves like follicular cancer but is often less iodine-avid.\n" +
    "Medullary: early nodal spread (central and lateral) even from small primaries, then liver, lung and bone.\n" +
    "Anaplastic: rapid local invasion of trachea, oesophagus, recurrent laryngeal nerve and great vessels; distant (often lung) metastases are frequently present at diagnosis.\n" +
    "Lymphoma: rapid local growth, may involve cervical nodes and become systemic.");
  const rows = [
    ["Tumour", "Main route", "Typical sites", "Key point", true],
    ["Papillary", "Lymphatic", "Central (VI) → lateral neck nodes; lung", "Often multifocal; nodes common at diagnosis", RED],
    ["Follicular", "Blood-borne", "Lung, bone", "Nodes uncommon; invasion defines it", TEAL],
    ["Oncocytic", "Blood-borne ± lymphatic", "Nodes, lung, bone", "Often less radioiodine-avid", TEAL],
    ["Medullary", "Lymphatic, then blood", "Central and lateral nodes; liver, lung, bone", "Nodes early, even with small primaries", TEAL],
    ["Anaplastic", "Direct invasion + blood", "Trachea, oesophagus, RLN; lung", "Airway threat; distant spread often at diagnosis", RED],
    ["Lymphoma", "Local, nodal", "Cervical nodes; systemic", "Rapid growth in Hashimoto thyroiditis", TEAL],
  ];
  const colX = [0.8, 2.85, 5.6, 9.1], colW = [2.0, 2.7, 3.45, 3.6];
  rows.forEach((r, i) => {
    const y = 1.7 + i * 0.6;
    const header = r[4] === true;
    s.addShape(pres.shapes.RECTANGLE, { x: X0, y, w: CW, h: 0.6, fill: { color: header ? NAVY : (i % 2 ? WHITE : ROWALT) }, line: { color: header ? NAVY : BORDER, width: 0.5 } });
    if (!header) s.addShape(pres.shapes.RECTANGLE, { x: X0, y, w: 0.1, h: 0.6, fill: { color: r[4] }, line: { type: "none" } });
    r.slice(0, 4).forEach((t, j) => s.addText(t, { x: colX[j], y, w: colW[j], h: 0.6, fontSize: header ? 12 : 14, bold: header || j === 0, color: header ? WHITE : (j === 0 ? NAVY : BODY), charSpacing: header ? 1.5 : 0, valign: "middle", margin: 0, isTextBox: true }));
  });
  remember(s, "papillary and medullary → nodes; follicular → blood (lung, bone); anaplastic → local invasion", 6.12);
}

divider(4);
// =====================================================================
// 10. Clinical presentation (signs + suspicious vs benign)
// =====================================================================
{
  const s = content(4, "Presentation: when is a nodule suspicious?", "[1] Bailey & Love; [6] ACR TI-RADS",
    "Most thyroid cancers present as a painless nodule, often found incidentally on imaging; patients are usually euthyroid.\n" +
    "Red flags of local invasion: hoarseness (recurrent laryngeal nerve), dysphagia, stridor, a hard fixed mass, cervical nodes.\n" +
    "Type clues: medullary (diarrhoea, flushing, MEN2 features), anaplastic (elderly, mass growing over weeks), lymphoma (rapid growth in Hashimoto).\n" +
    "Suspicious: childhood radiation, family history/MEN2, rapid growth, hard or fixed nodule, nodes. US (TI-RADS): solid, very hypoechoic, taller-than-wide, irregular margins, microcalcifications, extrathyroidal extension.\n" +
    "Reassuring: purely cystic or spongiform nodule, smooth margins, comet-tail artefact, 'hot' nodule on scan. No single feature is diagnostic.");
  card(s, { x: X0, y: 1.65, w: 5.95, h: 4.2, color: TEAL, kicker: "Signs and symptoms", title: "How it presents", rows: [
    { l: "Common", t: "Painless nodule, often incidental; usually euthyroid" },
    { l: "Local invasion", t: "Hoarseness, dysphagia, stridor; hard fixed mass; nodes", b: true },
    { l: "Type clues", t: "Medullary: diarrhoea, MEN2 · Anaplastic: grows over weeks · Lymphoma: Hashimoto" } ] });
  [[1.65, RED, RED_T, "!", "Raises suspicion", "Radiation, family history, rapid growth, hard or fixed; US: solid, hypoechoic, taller-than-wide, irregular, microcalcifications"],
   [3.82, TEAL, TEAL_T, "✓", "Reassuring", "Long-standing, soft, mobile, 'hot' on scan; US: purely cystic, spongiform, smooth margins, comet-tail"]].forEach(([y, c, tint, ic, h, t]) => {
    const x = X0 + 6.18, w = 5.95, hh = 2.03;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w, h: hh, fill: { color: tint }, line: { color: BORDER, width: 0.75 } });
    s.addShape(pres.shapes.OVAL, { x: x + 0.25, y: y + 0.22, w: 0.44, h: 0.44, fill: { color: c }, line: { type: "none" } });
    s.addText(ic, { x: x + 0.25, y: y + 0.22, w: 0.44, h: 0.44, fontSize: 15, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addText(h, { x: x + 0.85, y: y + 0.22, w: 4.5, h: 0.44, fontSize: 18, bold: true, color: NAVY, valign: "middle", margin: 0, isTextBox: true });
    s.addText(t, { x: x + 0.25, y: y + 0.82, w: w - 0.5, h: 1.1, fontSize: 14, color: BODY, valign: "top", margin: 0, isTextBox: true });
  });
  remember(s, "no single feature is diagnostic; suspicious features decide who needs FNA", 6.12);
}

divider(5);
// =====================================================================
// 12. Examination + laboratory tests
// =====================================================================
{
  const s = content(5, "Workup: examination and laboratory tests", "[2] Browse; [8] ATA 2015; [10] ATA MTC 2015",
    "Examination: a thyroid swelling moves on swallowing (pretracheal fascia); a thyroglossal cyst also moves on tongue protrusion. Palpate from behind: size, surface, consistency, mobility, lower border (retrosternal extension). Examine all cervical node levels. Check trachea, Pemberton sign, and document vocal cords by laryngoscopy before surgery.\n" +
    "Laboratory: TSH first. If TSH is low, isotope scan: a 'hot' nodule is rarely malignant. Thyroglobulin is not diagnostic; it is used after total thyroidectomy. Calcitonin and CEA are the medullary markers. If medullary: germline RET, plasma metanephrines (exclude phaeochromocytoma before surgery), serum calcium.");
  card(s, { x: X0, y: 1.65, w: 5.95, h: 4.2, color: TEAL, kicker: "Clinical examination", title: "Neck palpation", rows: [
    { l: "Inspect", t: "Moves on swallowing; thyroglossal cyst moves with tongue" },
    { l: "Palpate", t: "From behind: size, consistency, mobility, lower border" },
    { l: "Nodes", t: "All cervical levels: hard or matted nodes", b: true },
    { l: "Special", t: "Trachea, Pemberton sign; laryngoscopy before surgery" } ] });
  card(s, { x: X0 + 6.18, y: 1.65, w: 5.95, h: 4.2, color: RED, kicker: "Laboratory tests", title: "Function first, then markers", rows: [
    { l: "TSH", t: "First test; low TSH → isotope scan ('hot' = rarely cancer)", b: true },
    { l: "Thyroglobulin", t: "Not diagnostic; for follow-up after thyroidectomy" },
    { l: "Calcitonin · CEA", t: "Medullary carcinoma markers" },
    { l: "If medullary", t: "RET testing; metanephrines; calcium" } ] });
  remember(s, "a hard, fixed nodule with nodes or a hoarse voice is malignant until proven otherwise", 6.12);
}

// =====================================================================
// 14. Imaging + FNAC
// =====================================================================
{
  const s = content(5, "Workup: ultrasound and FNAC decide the next step", "[6] ACR TI-RADS; [7] Bethesda 2023; [8] ATA 2015",
    "Ultrasound is first-line for every nodule and the neck nodes. ACR TI-RADS: TR3 FNA ≥2.5 cm (follow ≥1.5), TR4 FNA ≥1.5 cm (follow ≥1.0), TR5 FNA ≥1.0 cm (follow ≥0.5); TR1–2 no FNA.\n" +
    "CT/MRI for retrosternal, invasive or bulky nodal disease (iodinated contrast can delay radioiodine). Isotope scan only when TSH is low. FDG-PET for anaplastic staging and iodine-refractory disease.\n" +
    "Bethesda 2023 risk of malignancy and usual action: I nondiagnostic 13% (repeat FNA); II benign 4% (follow-up); III AUS 22% (repeat FNA, molecular test or lobectomy); IV follicular neoplasm 30% (molecular test or lobectomy); V suspicious 74% (surgery); VI malignant 97% (surgery).\n" +
    "Core biopsy: repeated nondiagnostic FNA, or suspected anaplastic carcinoma or lymphoma (needs architecture and immunohistochemistry).");
  const tr = [["TR1", "Benign", "No FNA", "5E8F6B"], ["TR2", "Not suspicious", "No FNA", "5E8F6B"], ["TR3", "Mildly suspicious", "FNA ≥ 2.5 cm", "C9A227"], ["TR4", "Moderately susp.", "FNA ≥ 1.5 cm", ORANGE], ["TR5", "Highly suspicious", "FNA ≥ 1.0 cm", RED]];
  const tx = X0, tw = 5.95;
  box(s, tx, 1.65, tw, 4.2); topbar(s, tx, 1.65, tw, TEAL);
  label(s, "Ultrasound · ACR TI-RADS", tx + 0.28, 1.93, 5, TEAL, 10.5);
  tr.forEach(([k, n, a, c], i) => {
    const y = 2.4 + i * 0.66;
    s.addShape(pres.shapes.RECTANGLE, { x: tx + 0.28, y, w: tw - 0.56, h: 0.56, fill: { color: i % 2 ? WHITE : ROWALT }, line: { type: "none" } });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: tx + 0.36, y: y + 0.09, w: 0.72, h: 0.38, fill: { color: c }, line: { type: "none" }, rectRadius: 0.08 });
    s.addText(k, { x: tx + 0.36, y: y + 0.09, w: 0.72, h: 0.38, fontSize: 13, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addText(n, { x: tx + 1.22, y, w: 2.6, h: 0.56, fontSize: 14, color: NAVY, bold: true, valign: "middle", margin: 0, isTextBox: true });
    s.addText(a, { x: tx + 3.85, y, w: 1.9, h: 0.56, fontSize: 14, bold: i > 1, color: i > 1 ? NAVY : BODY, valign: "middle", margin: 0, isTextBox: true });
  });
  const cats = [["I", "Nondiagnostic", 13], ["II", "Benign", 4], ["III", "AUS", 22], ["IV", "Follicular neoplasm", 30], ["V", "Suspicious", 74], ["VI", "Malignant", 97]];
  const bx0 = X0 + 6.18;
  box(s, bx0, 1.65, tw, 4.2); topbar(s, bx0, 1.65, tw, RED);
  label(s, "FNAC · Bethesda 2023 risk of malignancy", bx0 + 0.28, 1.93, 5.4, RED, 10.5);
  cats.forEach(([k, n, m], i) => {
    const y = 2.35 + i * 0.57;
    const c = m >= 70 ? RED : m >= 20 ? ORANGE : TEAL;
    s.addText(k, { x: bx0 + 0.28, y, w: 0.45, h: 0.5, fontSize: 15, bold: true, color: c, valign: "middle", margin: 0, isTextBox: true });
    s.addText(n, { x: bx0 + 0.75, y, w: 2.2, h: 0.5, fontSize: 14, bold: true, color: NAVY, valign: "middle", margin: 0, isTextBox: true });
    const bx = bx0 + 3.0, bw = 1.9;
    s.addShape(pres.shapes.RECTANGLE, { x: bx, y: y + 0.15, w: bw, h: 0.2, fill: { color: "E3E7EE" }, line: { type: "none" } });
    s.addShape(pres.shapes.RECTANGLE, { x: bx, y: y + 0.15, w: Math.max(0.04, bw * m / 100), h: 0.2, fill: { color: c }, line: { type: "none" } });
    s.addText(`${m}%`, { x: bx + bw + 0.1, y, w: 0.65, h: 0.5, fontSize: 14, bold: true, color: NAVY, valign: "middle", margin: 0, isTextBox: true });
  });
  remember(s, "CT/MRI if invasive or retrosternal · isotope scan only if TSH is low · core biopsy if anaplastic or lymphoma suspected", 6.12, "Also: ");
}

divider(6);
// =====================================================================
// 16. Management principles by tumour type
// =====================================================================
{
  const s = content(6, "Management principles depend on the tumor type", "[8] ATA 2015; [10] ATA MTC 2015; [11] ATA ATC 2021",
    "Benign nodules: most need only clinical and ultrasound follow-up. Surgery for compressive symptoms, growth, cosmetic concern, or indeterminate cytology. Thermal ablation (e.g. radiofrequency) is an option in selected benign nodules.\n" +
    "Differentiated cancer (papillary, follicular, oncocytic): surgery, then radioiodine and TSH suppression chosen by risk of recurrence (risk-adapted, increasingly de-escalated).\n" +
    "Medullary: total thyroidectomy with central neck dissection (lateral dissection if nodes involved). Radioiodine and TSH suppression have no role because C cells do not take up iodine and are not TSH-driven. Exclude phaeochromocytoma first.\n" +
    "Anaplastic: urgent multidisciplinary care; resection only if achievable; test for BRAF V600E (dabrafenib + trametinib), radiotherapy, and early palliative care; protect the airway.\n" +
    "Lymphoma: chemotherapy with or without radiotherapy; surgery is not the treatment.");
  const rows = [
    ["Tumour", "Core treatment", "Add-on", true],
    ["Benign nodule", "Observe with US follow-up; surgery if compressive, growing or indeterminate", "Thermal ablation in selected cases", TEAL],
    ["Differentiated", "Lobectomy or total thyroidectomy", "Radioiodine and TSH suppression by risk", TEAL],
    ["Medullary", "Total thyroidectomy + central neck dissection", "No radioiodine; RET-targeted drugs if advanced", ORANGE],
    ["Anaplastic", "Urgent MDT; resect only if feasible", "BRAF-targeted therapy, radiotherapy, palliation", RED],
    ["Lymphoma", "Chemotherapy ± radiotherapy", "Surgery only for diagnosis or airway", TEAL],
  ];
  const colX = [0.8, 3.3, 8.2], colW = [2.4, 4.8, 4.4];
  rows.forEach((r, i) => {
    const y = 1.7 + i * (i === 0 ? 0 : 0.7) + (i > 0 ? -0.15 : 0);
    const h = i === 0 ? 0.55 : 0.7;
    const yy = i === 0 ? 1.7 : 2.25 + (i - 1) * 0.72;
    s.addShape(pres.shapes.RECTANGLE, { x: X0, y: yy, w: CW, h: i === 0 ? 0.55 : 0.72, fill: { color: i === 0 ? NAVY : (i % 2 ? WHITE : ROWALT) }, line: { color: i === 0 ? NAVY : BORDER, width: 0.5 } });
    if (i > 0) s.addShape(pres.shapes.RECTANGLE, { x: X0, y: yy, w: 0.1, h: 0.72, fill: { color: r[3] }, line: { type: "none" } });
    r.slice(0, 3).forEach((t, j) => s.addText(t, { x: colX[j], y: yy, w: colW[j], h: i === 0 ? 0.55 : 0.72, fontSize: i === 0 ? 12 : 14, bold: i === 0 || j === 0, color: i === 0 ? WHITE : (j === 0 ? NAVY : BODY), charSpacing: i === 0 ? 1.5 : 0, valign: "middle", margin: 0, isTextBox: true }));
  });
  remember(s, "radioiodine and TSH suppression work only for follicular-cell (differentiated) cancers", 6.12);
}

// =====================================================================
// 17. Surgery + non-surgical
// =====================================================================
{
  const s = content(6, "Treatment options: surgical and non-surgical", "[8] ATA 2015; [11] ATA ATC 2021; [18] LIBRETTO-001",
    "Lobectomy: low-risk differentiated cancer < 1 cm; acceptable for low-risk 1–4 cm tumours; diagnostic for indeterminate nodules.\n" +
    "Total thyroidectomy: tumour > 4 cm, gross extrathyroidal extension, clinical nodal or distant metastases, medullary cancer, or when radioiodine is planned.\n" +
    "Neck dissection: therapeutic central (VI) for involved nodes and routinely in medullary; lateral (II–V) only for proven lateral nodes.\n" +
    "Complications: recurrent laryngeal nerve palsy, hypocalcaemia, haematoma (airway emergency), external laryngeal nerve injury.\n" +
    "Non-surgical: radioiodine (I-131) after total thyroidectomy for differentiated cancer only, selective in low risk; TSH suppression by risk; external beam radiotherapy for anaplastic or gross residual disease; systemic therapy: lenvatinib/sorafenib (iodine-refractory), selpercatinib (RET, medullary), dabrafenib + trametinib (BRAF anaplastic).");
  card(s, { x: X0, y: 1.65, w: 5.95, h: 5.15, color: TEAL, kicker: "Surgical", title: "Operation by extent", rows: [
    { l: "Lobectomy", t: "Low-risk DTC < 1 cm; option for 1–4 cm; diagnostic" },
    { l: "Total thyroidectomy", t: "> 4 cm, gross ETE, nodes or metastases, medullary", b: true },
    { l: "Neck dissection", t: "Central (VI) if nodes and in medullary; lateral only if proven" },
    { l: "Complications", t: "Nerve palsy, hypocalcaemia, haematoma (airway emergency)" } ] });
  card(s, { x: X0 + 6.18, y: 1.65, w: 5.95, h: 5.15, color: RED, kicker: "Non-surgical", title: "Adjuvant and advanced disease", rows: [
    { l: "Radioiodine", t: "After total thyroidectomy; differentiated cancer only", b: true },
    { l: "TSH suppression", t: "Levothyroxine; target set by risk and response" },
    { l: "External beam RT", t: "Anaplastic; gross residual or unresectable disease" },
    { l: "Systemic therapy", t: "Lenvatinib (iodine-refractory); RET and BRAF inhibitors" } ] });
}

divider(7);
// =====================================================================
// 19. MY PART — Follow-up and surveillance
// =====================================================================
{
  const s = content(7, "Follow-up: response to therapy guides surveillance", "[8] ATA 2015; [9] ATA 2025; [10] ATA MTC 2015",
`SCRIPT (about 55 seconds)
"Now to prognosis. After treatment of differentiated cancer we follow two things: thyroglobulin, always with its antibodies, and neck ultrasound. Six to twelve months after treatment, each patient falls into one of four response groups. Excellent response means no evidence of disease, and recurrence of only one to four percent. Indeterminate patients mostly stay stable. Biochemical incomplete means a raised thyroglobulin with normal imaging; about a third settle on their own. Structural incomplete means visible disease and the highest risk. This response sets the TSH target and how often we follow up. The 2025 ATA guideline adds de-escalation: low-risk patients with an excellent response can stop ultrasound after five to eight years. Medullary cancer is followed with calcitonin and CEA instead."

BACKUP Q&A (not read aloud)
Q: Why measure anti-thyroglobulin antibodies with thyroglobulin?
A: The antibodies interfere with the assay and can make thyroglobulin falsely low. A falling antibody level is reassuring; a rising level is a surrogate for persistent disease.

Q: What are the thyroglobulin cut-offs for an excellent response?
A: After total thyroidectomy and radioiodine: suppressed Tg < 0.2 ng/mL or stimulated Tg < 1 ng/mL, with negative antibodies and negative imaging (ATA 2015). Biochemical incomplete: suppressed Tg ≥ 1 or stimulated Tg ≥ 10 ng/mL, or rising antibodies.

Q: What TSH target do you use?
A: ATA 2015: structural incomplete or high risk: below 0.1 mU/L. Biochemical incomplete: 0.1–0.5. Excellent response in low-risk patients: 0.5–2.0. Over-suppression causes atrial fibrillation and osteoporosis, especially in older patients.

Q: Can you follow a patient who had only a lobectomy?
A: Yes, mainly with neck ultrasound. Thyroglobulin is less useful because the remaining lobe makes it; the trend matters more than a single value.

Q: Why no radioiodine scan in medullary cancer follow-up?
A: C cells do not take up iodine. Follow-up uses calcitonin and CEA; a doubling time under 6 months indicates poor prognosis, over 2 years a good one.`);
  // response to therapy 2x2
  const resp = [
    ["Excellent", "No biochemical or structural disease", "Recurrence ~1–4%", TEAL],
    ["Indeterminate", "Non-specific imaging or low Tg", "Most stay stable", "C9A227"],
    ["Biochemical incomplete", "Raised Tg or rising TgAb; imaging negative", "~30% settle spontaneously", ORANGE],
    ["Structural incomplete", "Disease visible on imaging", "Highest risk; most persist", RED],
  ];
  label(s, "Response to therapy · 6–12 months after treatment", X0, 1.62, 7, NAVY, 10.5);
  resp.forEach(([t, d, o, c], i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = X0 + col * 3.6, y = 1.98 + row * 1.98, w = 3.45, h = 1.85;
    box(s, x, y, w, h); topbar(s, x, y, w, c);
    s.addText(t, { x: x + 0.22, y: y + 0.2, w: w - 0.4, h: 0.42, fontSize: 16, bold: true, color: NAVY, valign: "middle", margin: 0, isTextBox: true });
    s.addText(d, { x: x + 0.22, y: y + 0.66, w: w - 0.4, h: 0.62, fontSize: 13.5, color: BODY, valign: "top", margin: 0, isTextBox: true });
    s.addText(o, { x: x + 0.22, y: y + 1.32, w: w - 0.4, h: 0.38, fontSize: 13.5, bold: true, color: c === "C9A227" ? "8A6D00" : c, valign: "middle", margin: 0, isTextBox: true });
  });
  card(s, { x: 7.95, y: 1.62, w: 4.78, h: 4.21, color: TEAL, kicker: "Surveillance tools", title: "What we measure", rows: [
    { l: "Differentiated", t: "Thyroglobulin + anti-Tg antibodies; neck ultrasound", b: true },
    { l: "Medullary", t: "Calcitonin and CEA from 3 months; doubling time" },
    { l: "Long term", t: "Levothyroxine; TSH target set by response" } ] });
  remember(s, "low risk + excellent response: stop ultrasound after 5–8 years and thyroglobulin after 10–15 years", 6.12, "ATA 2025 de-escalation: ");
}

// =====================================================================
// 20. MY PART — Prognostic factors
// =====================================================================
{
  const s = content(7, "Prognostic factors: patient, tumor and treatment", "[12] AJCC 8th; [13] Hay 1993 (MACIS); [9] ATA 2025",
`SCRIPT (about 55 seconds)
"What predicts outcome? Age is the strongest patient factor: AJCC uses 55 years, and below 55 a patient is stage one unless there are distant metastases. Tumour factors are histology, size, gross extrathyroidal extension, vascular invasion, nodes, and above all distant metastases; BRAF together with TERT mutations adds risk. Treatment matters too: a complete resection and iodine-avid disease are favourable. The MACIS score combines these: below six predicts ninety-nine percent survival at twenty years, eight or more only twenty-four percent. One key point: AJCC staging predicts death, while the ATA system predicts recurrence."

BACKUP Q&A (not read aloud)
Q: How is MACIS calculated?
A: 3.1 if age < 40 years, or 0.08 × age if ≥ 40; plus 0.3 × tumour size in cm; plus 1 if resection incomplete; plus 1 if locally invasive; plus 3 if distant metastases. 20-year cause-specific survival: < 6 = 99%, 6–6.99 = 89%, 7–7.99 = 56%, ≥ 8 = 24% (Hay 1993).

Q: Other scoring systems?
A: AGES (Age, Grade, Extent, Size) and AMES (Age, Metastases, Extent, Size). MACIS replaced grade with completeness of resection because grade is less reproducible.

Q: Stage a 60-year-old with a 3 cm intrathyroidal PTC and central nodes.
A: AJCC 8th, age ≥ 55: T2 N1a M0 = stage II. The same tumour at age 40 = stage I.

Q: Why is age so important?
A: Older patients more often have aggressive histology, less iodine-avid disease, and higher mortality; young patients have excellent survival even with nodal disease.

Q: Which single factor is worst?
A: Distant metastases (especially non-iodine-avid ones), and anaplastic histology.`);
  card(s, { x: X0, y: 1.7, w: 3.15, h: 4.15, color: TEAL, kicker: "Patient · treatment", title: "Who and how", rows: [
    { l: "Age", t: "AJCC cut-off 55 years", b: true },
    { l: "Under 55", t: "Stage I (II only if M1)" },
    { l: "Treatment", t: "Complete resection; iodine-avid disease" } ] });
  card(s, { x: X0 + 3.33, y: 1.7, w: 3.95, h: 4.15, color: RED, kicker: "Tumour", title: "What", rows: [
    { l: "Histology", t: "Type and aggressive subtypes" },
    { l: "Local", t: "Size; gross ETE; vascular invasion" },
    { l: "Spread", t: "Nodes; distant metastases (worst)", b: true },
    { l: "Molecular", t: "BRAF V600E + TERT together" } ] });
  card(s, { x: X0 + 7.46, y: 1.7, w: 4.67, h: 4.15, color: NAVY, kicker: "Score · MACIS", title: "20-year survival", rows: [] });
  const mac = [["< 6", "99%", TEAL], ["6–6.99", "89%", "C9A227"], ["7–7.99", "56%", ORANGE], ["≥ 8", "24%", RED]];
  s.addText("Metastasis · Age · Completeness · Invasion · Size", { x: X0 + 7.74, y: 2.82, w: 4.2, h: 0.3, fontSize: 12, color: MUTED, margin: 0, isTextBox: true });
  mac.forEach(([k, v, c], i) => {
    const y = 3.22 + i * 0.62;
    s.addShape(pres.shapes.RECTANGLE, { x: X0 + 7.74, y, w: 4.11, h: 0.54, fill: { color: i % 2 ? WHITE : ROWALT }, line: { type: "none" } });
    s.addText("MACIS " + k, { x: X0 + 7.9, y, w: 1.9, h: 0.54, fontSize: 14, bold: true, color: NAVY, valign: "middle", margin: 0, isTextBox: true });
    s.addShape(pres.shapes.RECTANGLE, { x: X0 + 9.8, y: y + 0.17, w: 1.2 * parseInt(v) / 100, h: 0.2, fill: { color: c }, line: { type: "none" } });
    s.addText(v, { x: X0 + 11.05, y, w: 0.75, h: 0.54, fontSize: 14, bold: true, color: NAVY, valign: "middle", align: "right", margin: 0, isTextBox: true });
  });
  remember(s, "AJCC staging predicts the risk of death; ATA risk predicts the risk of recurrence", 6.12, "Key: ");
}

// =====================================================================
// 21. MY PART — Survival and recurrence
// =====================================================================
{
  const s = content(7, "Survival is excellent; recurrence is common", "[14] SEER; [16] Roman 2006; [11] ATA ATC 2021; [9] ATA 2025",
`SCRIPT (about 50 seconds)
"Survival is excellent overall. In SEER data, five-year survival is almost one hundred percent for localised and regional disease, but about fifty percent with distant metastases. Histology matters: medullary cancer has ten-year survival of ninety-six percent when localised but forty percent with distant spread, while anaplastic cancer has a median survival of about five months. Recurrence, however, is common. The 2025 ATA guideline grades it in four tiers, from below ten percent to above thirty percent. Most recurrences are in the neck nodes and appear within the first decade, which is why follow-up is long."

BACKUP Q&A (not read aloud)
Q: Where does differentiated cancer recur?
A: Mostly in cervical lymph nodes (central and lateral neck), then the thyroid bed; distant recurrence is mainly lung and bone.

Q: What is the long-term recurrence rate of papillary cancer?
A: In a 30-year cohort, recurrence was about 30%, with about two-thirds in the first decade (Mazzaferri & Kloos 2001).

Q: How is recurrence treated?
A: Biopsy the lesion (FNA with Tg washout for nodes), then compartment-oriented neck dissection for structural nodal disease, radioiodine if iodine-avid, active surveillance for small stable nodes, and systemic therapy for progressive iodine-refractory disease.

Q: Why does anaplastic cancer have such poor survival?
A: It is undifferentiated, grows within weeks, invades the airway early, is often metastatic at diagnosis, and does not take up iodine. All cases are stage IV by definition.

Q: What changed between the 2015 and 2025 ATA risk systems?
A: 2015 had three tiers (low, intermediate, high). 2025 splits intermediate into low-intermediate and intermediate-high, and reassesses risk dynamically at every visit.`);
  // chart
  box(s, X0, 1.62, 5.3, 4.21); topbar(s, X0, 1.62, 5.3, TEAL);
  label(s, "5-year relative survival by stage (SEER)", X0 + 0.28, 1.9, 4.8, TEAL, 10.5);
  s.addChart(pres.charts.BAR, [{ name: "5-year survival", labels: ["Localized", "Regional", "Distant"], values: [99.9, 98.1, 50.2] }], {
    x: X0 + 0.15, y: 2.25, w: 5.0, h: 3.5, barDir: "col", chartColors: [TEAL, TEAL, RED], barGapWidthPct: 60,
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 14, dataLabelFontBold: true, dataLabelColor: NAVY, dataLabelFormatCode: '0.0"%"',
    valAxisHidden: true, valAxisMaxVal: 115, valAxisMinVal: 0, valGridLine: { style: "none" }, catGridLine: { style: "none" },
    catAxisLabelColor: NAVY, catAxisLabelFontSize: 13, catAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", catAxisLineShow: true,
    showLegend: false, showTitle: false,
  });
  card(s, { x: X0 + 5.5, y: 1.62, w: 6.63, h: 2.55, color: RED, kicker: "By histology", title: null, rows: [
    { l: "Medullary · 10-year survival", t: "Localized 96% · Regional 76% · Distant 40%", b: true },
    { l: "Anaplastic", t: "Median survival ≈ 5 months; 1-year ≈ 20%", b: true } ] });
  // recurrence tiers
  box(s, X0 + 5.5, 4.32, 6.63, 1.51); topbar(s, X0 + 5.5, 4.32, 6.63, ORANGE);
  label(s, "Recurrence risk · ATA 2025 (4 tiers)", X0 + 5.78, 4.55, 5, ORANGE, 10.5);
  const tiers = [["Low", "< 10%", TEAL], ["Low-int.", "10–15%", "B08A1A"], ["Int.-high", "16–30%", ORANGE], ["High", "> 30%", RED]];
  tiers.forEach(([t, v, c], i) => {
    const w = 1.5, x = X0 + 5.78 + i * (w + 0.03);
    s.addShape(pres.shapes.RECTANGLE, { x, y: 4.92, w, h: 0.78, fill: { color: c }, line: { type: "none" } });
    s.addText([{ text: v, options: { bold: true, fontSize: 16, breakLine: true } }, { text: t, options: { fontSize: 12 } }], { x, y: 4.92, w, h: 0.78, color: WHITE, align: "center", valign: "middle", margin: 0, isTextBox: true });
  });
  remember(s, "nodes rarely change survival, but they drive recurrence; most recurrences are in the neck", 6.12);
}

divider(8);
// =====================================================================
// 22. MY PART — Case 1
// =====================================================================
{
  const s = content(8, "Case 1: a thyroid nodule in a young woman", "[6] ACR TI-RADS; [7] Bethesda 2023; [8] ATA 2015; [12] AJCC 8th",
`SCRIPT (about 55 seconds)
"Case one, an illustrative case. A thirty-four-year-old woman has a two-point-two centimetre nodule; TSH is normal. Ultrasound shows TI-RADS five features and FNA is Bethesda six: papillary carcinoma. Being under fifty-five, she is AJCC stage one and low risk. Lobectomy or total thyroidectomy are both acceptable; she has a total thyroidectomy showing only micrometastases, so no routine radioiodine. At one year, her thyroglobulin is zero-point-one with negative antibodies and a clear ultrasound: an excellent response. We relax TSH suppression and space out follow-up. Had thyroglobulin been five with a clear scan, that would be a biochemical incomplete response, and we would watch the trend."

BACKUP Q&A (not read aloud)
Q: Why not give radioiodine routinely?
A: In ATA low-risk disease, radioiodine has not been shown to improve survival or meaningfully reduce recurrence; the 2025 guideline is even more selective.

Q: What would push her to total thyroidectomy as the only option?
A: Tumour > 4 cm, gross extrathyroidal extension, clinical nodal (cN1) or distant metastases.

Q: What would make her high risk?
A: Gross extrathyroidal extension, incomplete resection, distant metastases, or bulky nodal disease (a node ≥ 3 cm).

Q: Pregnancy plans?
A: Usually advised to wait 6–12 months after radioiodine if given; levothyroxine dose often needs to increase in pregnancy.`);
  const cols = [
    ["Scenario", WHITE, NAVY, BODY, "34-year-old woman; painless 2.2 cm right nodule; TSH normal. US: solid, hypoechoic, irregular, microcalcifications (TR5). FNA: Bethesda VI."],
    ["Think", TEAL_T, TEAL, BODY, "Papillary carcinoma, cT2 N0 M0. Under 55: AJCC stage I. Low risk of recurrence."],
    ["Decide", TEAL_T, TEAL, BODY, "Lobectomy or total thyroidectomy (1–4 cm, low risk). Total done: 2 central micrometastases < 2 mm. No routine radioiodine."],
    ["Outcome at 12 months", NAVY, MINT, WHITE, "Suppressed Tg 0.1 ng/mL, antibodies negative, US clear: excellent response. TSH 0.5–2; spaced follow-up."],
  ];
  cols.forEach(([h, fill, hc, tc, t], i) => {
    const x = X0 + i * 3.093, w = 2.92, y = 1.7, hh = 4.15;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w, h: hh, fill: { color: fill }, line: { color: fill === WHITE ? BORDER : fill === TEAL_T ? "BFDDD8" : NAVY, width: 0.75 } });
    label(s, h, x + 0.25, y + 0.25, w - 0.4, hc, 10.5);
    s.addText(t, { x: x + 0.25, y: y + 0.65, w: w - 0.45, h: hh - 0.85, fontSize: 16.5, color: tc, bold: i === 3, valign: "top", margin: 0, isTextBox: true });
    if (i < 3) s.addShape(pres.shapes.LINE, { x: x + w + 0.005, y: y + hh / 2, w: 0.165, h: 0, line: { color: NAVY, width: 1.5, endArrowType: "triangle" } });
  });
  remember(s, "if Tg were 5 ng/mL with a clear US: biochemical incomplete response → watch the trend", 6.12, "What if? ");
}

// =====================================================================
// 23. MY PART — Cases 2 and 3
// =====================================================================
{
  const s = content(8, "Cases 2 and 3: what would you do?", "[10] ATA MTC 2015; [11] ATA ATC 2021",
`SCRIPT (about 60 seconds)
"Two contrasting cases. A forty-two-year-old man has a nodule, a lateral neck node, diarrhoea, and a father with thyroid cancer and severe hypertension. This is medullary cancer, probably MEN2A. Before thyroid surgery we must exclude a phaeochromocytoma, check calcium, and send RET testing. Treatment is total thyroidectomy with central and lateral neck dissection, then follow-up with calcitonin and CEA; relatives are offered RET testing.
A seventy-four-year-old woman has a hard mass growing over three weeks, with hoarseness and breathlessness. Think anaplastic cancer, but exclude lymphoma with a core biopsy, and test for BRAF. She needs airway assessment, urgent multidisciplinary care, targeted therapy if BRAF-positive, and early palliative care. Thank you."

BACKUP Q&A (not read aloud)
Q: Why exclude phaeochromocytoma before thyroidectomy in MTC?
A: In MEN2A about half of patients develop phaeochromocytoma; anaesthesia and surgery can trigger a catecholamine crisis. Adrenalectomy with alpha-blockade comes first.

Q: When is prophylactic thyroidectomy done in RET carriers?
A: Timing depends on the ATA risk class of the RET mutation: highest risk (M918T, MEN2B) within the first year of life; high risk (codon 634) at or before age 5; moderate risk guided by calcitonin.

Q: How do you differentiate anaplastic carcinoma from lymphoma?
A: Both are rapidly enlarging masses in older patients; lymphoma often arises in Hashimoto thyroiditis. Core biopsy with immunohistochemistry (and flow cytometry) is needed; FNA alone may be insufficient.

Q: Is tracheostomy routine in anaplastic cancer?
A: No. It is avoided if possible and decided by the multidisciplinary team according to goals of care; airway is assessed early.`);
  const cases = [
    { x: X0, color: RED, k: "Case 2", t: "42-year-old man", sc: "2 cm firm nodule + level III node; diarrhoea; father: thyroid cancer and severe hypertension episodes. High calcitonin and CEA.",
      th: "Medullary carcinoma, likely MEN2A: RET testing; exclude phaeochromocytoma; check calcium",
      de: "Adrenal first if needed → total thyroidectomy + central and lateral neck dissection. Follow calcitonin/CEA" },
    { x: X0 + 6.18, color: TEAL, k: "Case 3", t: "74-year-old woman", sc: "Hard, fixed neck mass growing over 3 weeks; hoarseness and breathlessness.",
      th: "Anaplastic carcinoma vs lymphoma: core biopsy, BRAF V600E; stage IV by definition",
      de: "Secure airway; urgent MDT; BRAF-targeted therapy ± RT; early palliative care" },
  ];
  cases.forEach((c) => {
    const w = 5.95;
    box(s, c.x, 1.62, w, 5.2); topbar(s, c.x, 1.62, w, c.color);
    label(s, c.k, c.x + 0.28, 1.85, 3, c.color, 10.5);
    s.addText(c.t, { x: c.x + 0.28, y: 2.12, w: w - 0.5, h: 0.42, fontSize: 18, bold: true, color: NAVY, valign: "middle", margin: 0, isTextBox: true });
    const ix = c.x + 0.28, iw = w - 0.56;
    // scenario
    s.addShape(pres.shapes.RECTANGLE, { x: ix, y: 2.65, w: iw, h: 1.4, fill: { color: WHITE }, line: { color: BORDER, width: 0.75 } });
    label(s, "Scenario", ix + 0.15, 2.72, 3, TEAL, 10);
    s.addText(c.sc, { x: ix + 0.15, y: 2.98, w: iw - 0.3, h: 1.02, fontSize: 14, color: BODY, valign: "top", margin: 0, isTextBox: true });
    s.addShape(pres.shapes.RECTANGLE, { x: ix, y: 4.12, w: iw, h: 1.25, fill: { color: TEAL_T }, line: { color: "BFDDD8", width: 0.75 } });
    label(s, "Think", ix + 0.15, 4.19, 3, TEAL, 10);
    s.addText(c.th, { x: ix + 0.15, y: 4.45, w: iw - 0.3, h: 0.88, fontSize: 14, color: BODY, valign: "top", margin: 0, isTextBox: true });
    s.addShape(pres.shapes.RECTANGLE, { x: ix, y: 5.44, w: iw, h: 1.25, fill: { color: NAVY }, line: { color: NAVY, width: 0.75 } });
    label(s, "Decide", ix + 0.15, 5.51, 3, MINT, 10);
    s.addText(c.de, { x: ix + 0.15, y: 5.77, w: iw - 0.3, h: 0.88, fontSize: 14, bold: true, color: WHITE, valign: "top", margin: 0, isTextBox: true });
  });
}

// =====================================================================
// 24. References
// =====================================================================
{
  pres.addSection({ title: "Closing" });
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Closing" });
  s.addText("References", { placeholder: "title" });
  const refs = [
    "Williams NS, O'Connell PR, McCaskie AW, eds. Bailey & Love's Short Practice of Surgery. 28th ed. CRC Press; 2023.",
    "Burnand KG, Black J, Corbett SA, Thomas WEG. Browse's Introduction to the Symptoms & Signs of Surgical Disease. 6th ed. CRC Press; 2021.",
    "Bray F, Laversanne M, Sung H, et al. Global cancer statistics 2022: GLOBOCAN estimates. CA Cancer J Clin. 2024;74(3):229–263.",
    "Saudi Health Council, National Cancer Center. Cancer Incidence Report, Saudi Arabia 2020. Saudi Cancer Registry.",
    "Baloch ZW, Asa SL, Barletta JA, et al. Overview of the 2022 WHO classification of thyroid neoplasms. Endocr Pathol. 2022;33(1):27–63.",
    "Tessler FN, Middleton WD, Grant EG, et al. ACR Thyroid Imaging, Reporting and Data System (TI-RADS). J Am Coll Radiol. 2017;14(5):587–595.",
    "Ali SZ, Baloch ZW, Cochand-Priollet B, et al. The 2023 Bethesda System for Reporting Thyroid Cytopathology. Thyroid. 2023;33(9):1039–1044.",
    "Haugen BR, Alexander EK, Bible KC, et al. 2015 ATA management guidelines for adult patients with thyroid nodules and differentiated thyroid cancer. Thyroid. 2016;26(1):1–133.",
    "Ringel MD, Sosa JA, Baloch Z, et al. 2025 ATA management guidelines for adult patients with differentiated thyroid cancer. Thyroid. 2025;35(8):841–985.",
    "Wells SA Jr, Asa SL, Dralle H, et al. Revised ATA guidelines for the management of medullary thyroid carcinoma. Thyroid. 2015;25(6):567–610.",
    "Bible KC, Kebebew E, Brierley J, et al. 2021 ATA guidelines for management of patients with anaplastic thyroid cancer. Thyroid. 2021;31(3):337–386.",
    "Amin MB, Edge SB, Greene FL, et al., eds. AJCC Cancer Staging Manual. 8th ed. Springer; 2017.",
    "Hay ID, Bergstralh EJ, Goellner JR, et al. Predicting outcome in papillary thyroid carcinoma (MACIS). Surgery. 1993;114(6):1050–1058.",
    "National Cancer Institute. SEER Cancer Stat Facts: Thyroid Cancer. seer.cancer.gov/statfacts/html/thyro.html",
    "Mazzaferri EL, Kloos RT. Current approaches to primary therapy for papillary and follicular thyroid cancer. J Clin Endocrinol Metab. 2001;86(4):1447–1463.",
    "Roman S, Lin R, Sosa JA. Prognosis of medullary thyroid carcinoma: predictors of survival in 1252 patients. Cancer. 2006;107(9):2134–2142.",
    "Barbet J, Campion L, Kraeber-Bodéré F, et al. Prognostic impact of calcitonin and CEA doubling-times in medullary thyroid carcinoma. J Clin Endocrinol Metab. 2005;90(11):6077–6084.",
    "Wirth LJ, Sherman E, Robinson B, et al. Efficacy of selpercatinib in RET-altered thyroid cancers (LIBRETTO-001). N Engl J Med. 2020;383(9):825–835.",
  ];
  const half = 9;
  [refs.slice(0, half), refs.slice(half)].forEach((col, ci) => {
    const runs = [];
    col.forEach((r, i) => {
      const n = ci * half + i + 1;
      runs.push({ text: `${n}  `, options: { bold: true, color: TEAL } });
      runs.push({ text: r, options: { color: BODY, breakLine: i < col.length - 1, paraSpaceAfter: 7 } });
    });
    s.addText(runs, { x: X0 + ci * 6.18, y: 1.6, w: 5.95, h: 5.3, fontSize: 10.5, valign: "top", margin: 0, isTextBox: true });
  });
  s.addNotes("One reference per fact: where several sources say the same thing, only one is cited on the slide.\nTextbooks [1] and [2] are the course essential references; guideline and registry sources were used for current numbers.");
}

// =====================================================================
// 25. Thank you
// =====================================================================
{
  const s = pres.addSlide({ masterName: "DARK", sectionTitle: "Closing" }); addCircles(s);
  s.addText("Thank you", { x: 0.85, y: 2.65, w: 8, h: 1.3, fontFace: HEAD, fontSize: 66, bold: true, color: WHITE, margin: 0, isTextBox: true });
  s.addText("Questions and discussion", { x: 0.9, y: 4.05, w: 7, h: 0.5, fontSize: 18, color: "A9C3D6", margin: 0, isTextBox: true });
  s.addNotes("Invite questions. Backup answers for the prognosis and case sections are in the notes of slides 19–23.");
}

(async () => {
  const out = path.join(__dirname, "..", "Thyroid_Tumors_Seminar.pptx");
  await pres.writeFile({ fileName: out });
  await applyTheme(out, THEME);
  console.log("written", out);
})();
