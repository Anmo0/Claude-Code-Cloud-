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
  // per-slide source lines removed at user request; sources stay in speaker notes
  if (notes || sources) s.addNotes((notes || "") + (sources ? `\n\nSources: ${sources}` : ""));
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
    "Surgical landmarks: the recurrent laryngeal nerve is most at risk where it passes close to the ligament of Berry, and the tubercle of Zuckerkandl usually points to it; the external branch of the superior laryngeal nerve runs with the superior pole vessels, so they are ligated close to the gland. Superior parathyroids are usually posterior at the level of the cricoid; inferior ones are more variable.\n" +
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
    { l: "Surgical landmarks", t: "RLN at Berry's ligament and under Zuckerkandl's tubercle; EBSLN at the upper pole", b: true } ] });
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
    "Malignant, follicular-cell derived (about 95%): papillary (~80%), follicular (~10%), oncocytic carcinoma, high-grade carcinoma (WHO 2022: poorly differentiated and differentiated high-grade) and anaplastic (~1–2%). WHO 2022 also renamed Hürthle cell tumours as oncocytic.\n" +
    "C-cell: medullary (~5%; about 25% hereditary, RET).\n" +
    "Other: primary lymphoma (in Hashimoto thyroiditis) and metastases (kidney, lung, breast, melanoma). Percentages are approximate.");
  s.addShape(pres.shapes.OVAL, { x: 0.45, y: 3.05, w: 2.35, h: 2.35, fill: { color: WHITE }, line: { color: "7FB3AE", width: 1.5 } });
  s.addShape(pres.shapes.OVAL, { x: 0.6, y: 3.2, w: 2.05, h: 2.05, fill: { color: NAVY }, line: { type: "none" } });
  s.addText("Thyroid\ntumours", { x: 0.6, y: 3.2, w: 2.05, h: 2.05, fontFace: HEAD, fontSize: 19, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0, isTextBox: true });
  const groups = [
    { name: "Benign", sub: "no invasion", color: TEAL, tint: TEAL_T, items: [["Follicular adenoma", "incl. oncocytic, toxic"], ["Cysts", "pure cyst: very low risk"], ["NIFTP", "low-risk neoplasm"]] },
    { name: "Follicular-cell", sub: "malignant, ~95%", color: RED, tint: RED_T, items: [["Papillary carcinoma", "~80%", true], ["Follicular carcinoma", "~10%"], ["Oncocytic carcinoma", ""], ["High-grade · Anaplastic", "anaplastic ~1–2%"]] },
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
// 12. Clinical examination
// =====================================================================
{
  const s = content(5, "Clinical examination and neck palpation", "[2] Browse",
    "Inspect from the front: a thyroid swelling moves up on swallowing (pretracheal fascia). A thyroglossal cyst also moves on tongue protrusion.\n" +
    "Palpate from behind with the neck slightly flexed: size, surface, consistency, tenderness, mobility, and whether you can get below the lower border (if not, suspect retrosternal extension).\n" +
    "Examine all cervical node levels systematically.\n" +
    "Special tests: tracheal position; Pemberton sign for retrosternal goitre (facial plethora on raising both arms); percussion over the manubrium; auscultation for a bruit.\n" +
    "Document vocal cord function (laryngoscopy) before any thyroid surgery.");
  const steps = [
    ["Step 1", "Inspect", TEAL, "Front view; ask the patient to swallow", "Thyroid rises on swallowing; thyroglossal cyst rises on tongue protrusion"],
    ["Step 2", "Palpate", TEAL, "From behind, neck slightly flexed", "Size, surface, consistency, mobility; can you get below it?"],
    ["Step 3", "Nodes", RED, "All cervical levels on both sides", "Hard or matted nodes suggest metastasis"],
    ["Step 4", "Special tests", RED, "Trachea, Pemberton sign, percussion, bruit", "Retrosternal extension; laryngoscopy before surgery"],
  ];
  steps.forEach(([k, t, c, how, look], i) => card(s, { x: X0 + i * 3.093, y: 1.7, w: 2.92, h: 4.15, color: c, kicker: k, title: t, rows: [
    { l: "How", t: how }, { l: "Look for", t: look, b: true } ] }));
  remember(s, "a hard, fixed nodule with ipsilateral nodes or a hoarse voice is malignant until proven otherwise", 6.12);
}

// =====================================================================
// 13. Laboratory tests
// =====================================================================
{
  const s = content(5, "Laboratory tests: function first, then markers", "[8] ATA 2015; [10] ATA MTC 2015",
    "Serum TSH is the first test for every nodule. If TSH is low, do radionuclide scintigraphy: a hyperfunctioning ('hot') nodule is rarely malignant and does not need FNA.\n" +
    "Thyroglobulin is not a diagnostic test (it is raised in benign disease); it is used for surveillance after total thyroidectomy, interpreted with anti-thyroglobulin antibodies.\n" +
    "Calcitonin: marker of medullary carcinoma. Routine screening of all nodules is debated; measure when MTC is suspected. CEA supports staging and follow-up.\n" +
    "If medullary cancer is diagnosed: germline RET testing, plasma free metanephrines (exclude phaeochromocytoma before surgery), and serum calcium (hyperparathyroidism).\n" +
    "Molecular testing of indeterminate FNA (for example BRAF, RAS, gene classifiers) can refine risk.");
  card(s, { x: X0, y: 1.7, w: 3.88, h: 4.15, color: TEAL, kicker: "First test", title: "Thyroid function", rows: [
    { l: "TSH", t: "First for every nodule", b: true },
    { l: "Low TSH", t: "Isotope scan: 'hot' nodule rarely malignant" },
    { l: "Free T4 / T3", t: "Most cancers are euthyroid" } ] });
  card(s, { x: X0 + 4.125, y: 1.7, w: 3.88, h: 4.15, color: RED, kicker: "Tumour markers", title: "Use them correctly", rows: [
    { l: "Thyroglobulin", t: "Not diagnostic; for follow-up after total thyroidectomy", b: true },
    { l: "Calcitonin", t: "Medullary carcinoma marker" },
    { l: "CEA", t: "Medullary staging and follow-up" } ] });
  card(s, { x: X0 + 8.25, y: 1.7, w: 3.88, h: 4.15, color: TEAL, kicker: "If medullary", title: "Before surgery", rows: [
    { l: "Genetics", t: "Germline RET testing" },
    { l: "Adrenal", t: "Metanephrines: exclude phaeochromocytoma", b: true },
    { l: "Calcium", t: "Look for hyperparathyroidism (MEN2A)" } ] });
  remember(s, "thyroglobulin does not diagnose cancer; it monitors patients after total thyroidectomy", 6.12);
}

// =====================================================================
// 14. Imaging
// =====================================================================
{
  const s = content(5, "Imaging: ultrasound first, the rest by question", "[6] ACR TI-RADS; [8] ATA 2015",
    "Ultrasound is the first-line test for every nodule and for the cervical nodes. Before surgery for any confirmed or suspected cancer, a dedicated neck ultrasound maps the central and lateral nodes: it changes the operation (ATA 2015). ACR TI-RADS adds points for composition, echogenicity, shape, margin and echogenic foci; the level and the nodule size decide FNA or follow-up.\n" +
    "TR3: FNA if ≥2.5 cm, follow if ≥1.5 cm. TR4: FNA if ≥1.5 cm, follow if ≥1.0 cm. TR5: FNA if ≥1.0 cm, follow if ≥0.5 cm. TR1 and TR2: no FNA.\n" +
    "CT or MRI: retrosternal extension, invasion of trachea or oesophagus, bulky or mediastinal nodes. Iodinated contrast can delay radioiodine therapy, so plan it.\n" +
    "Radionuclide scan (Tc-99m or I-123): only when TSH is low, to find a hot nodule.\n" +
    "FDG-PET: staging anaplastic cancer and radioiodine-refractory disease.");
  const tr = [["TR1", "Benign", "0", "No FNA", "5E8F6B"], ["TR2", "Not suspicious", "2", "No FNA", "5E8F6B"], ["TR3", "Mildly suspicious", "3", "FNA ≥ 2.5 cm", "C9A227"], ["TR4", "Moderately suspicious", "4–6", "FNA ≥ 1.5 cm", ORANGE], ["TR5", "Highly suspicious", "≥ 7", "FNA ≥ 1.0 cm", RED]];
  const tx = X0, tw = 6.9;
  box(s, tx, 1.7, tw, 4.15); topbar(s, tx, 1.7, tw, TEAL);
  label(s, "Ultrasound · ACR TI-RADS", tx + 0.28, 1.98, 5, TEAL, 10.5);
  ["Level", "Points", "Action"].forEach((h, j) => s.addText(h.toUpperCase(), { x: [tx + 0.28, tx + 3.85, tx + 4.85][j], y: 2.32, w: 1.8, h: 0.3, fontSize: 10, bold: true, color: MUTED, charSpacing: 2, margin: 0, isTextBox: true }));
  tr.forEach(([k, n, p, a, c], i) => {
    const y = 2.68 + i * 0.6;
    s.addShape(pres.shapes.RECTANGLE, { x: tx + 0.28, y, w: tw - 0.56, h: 0.52, fill: { color: i % 2 ? WHITE : ROWALT }, line: { type: "none" } });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: tx + 0.36, y: y + 0.07, w: 0.72, h: 0.38, fill: { color: c }, line: { type: "none" }, rectRadius: 0.08 });
    s.addText(k, { x: tx + 0.36, y: y + 0.07, w: 0.72, h: 0.38, fontSize: 13, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addText(n, { x: tx + 1.2, y, w: 2.6, h: 0.52, fontSize: 14, color: NAVY, bold: true, valign: "middle", margin: 0, isTextBox: true });
    s.addText(p, { x: tx + 3.85, y, w: 0.9, h: 0.52, fontSize: 14, color: BODY, valign: "middle", margin: 0, isTextBox: true });
    s.addText(a, { x: tx + 4.85, y, w: 1.8, h: 0.52, fontSize: 14, bold: i > 1, color: i > 1 ? NAVY : BODY, valign: "middle", margin: 0, isTextBox: true });
  });
  card(s, { x: 7.75, y: 1.7, w: 4.98, h: 4.15, color: TEAL, kicker: "Beyond ultrasound", title: "When to add", rows: [
    { l: "Pre-op neck US", t: "Map central and lateral nodes in every cancer", b: true },
    { l: "CT / MRI", t: "Retrosternal, invasive or bulky nodes" },
    { l: "Isotope scan", t: "Only if TSH is low: hot vs cold" },
    { l: "FDG-PET", t: "Anaplastic; iodine-refractory disease" } ] });
  remember(s, "iodinated CT contrast can delay radioiodine therapy; request it only when it changes management", 6.12, "Caution: ");
}

// =====================================================================
// 15. FNAC and core biopsy (Bethesda 2023)
// =====================================================================
{
  const s = content(5, "FNAC and core biopsy: Bethesda 2023 guides action", "[7] Bethesda 2023",
    "Ultrasound-guided FNA is the key diagnostic test. Results are reported in the six Bethesda categories, each with an implied risk of malignancy (2023, 3rd edition; mean and range):\n" +
    "I Nondiagnostic 13% (5–20): repeat US-guided FNA. II Benign 4% (2–7): follow-up. III AUS 22% (13–30): repeat FNA, molecular testing, or lobectomy. IV Follicular neoplasm 30% (23–34): molecular testing or lobectomy. V Suspicious for malignancy 74% (67–83): molecular testing or surgery. VI Malignant 97% (97–100): surgery.\n" +
    "Limitation: FNA cannot tell follicular adenoma from follicular carcinoma (needs capsule and vessels on histology).\n" +
    "Core needle biopsy keeps tissue architecture: useful after repeated nondiagnostic FNA and when anaplastic carcinoma or lymphoma is suspected (immunohistochemistry, flow studies).");
  const cats = [["I", "Nondiagnostic", 13, "5–20", "Repeat US-guided FNA"], ["II", "Benign", 4, "2–7", "Clinical and US follow-up"], ["III", "AUS", 22, "13–30", "Repeat FNA / molecular / lobectomy"], ["IV", "Follicular neoplasm", 30, "23–34", "Molecular testing or lobectomy"], ["V", "Suspicious for malignancy", 74, "67–83", "Molecular testing or surgery"], ["VI", "Malignant", 97, "97–100", "Surgery"]];
  const tx = X0, tw = 8.0;
  box(s, tx, 1.7, tw, 4.15); topbar(s, tx, 1.7, tw, RED);
  label(s, "Risk of malignancy, mean (range)", tx + 0.28, 1.98, 5, RED, 10.5);
  cats.forEach(([k, n, m, r, a], i) => {
    const y = 2.38 + i * 0.57;
    const c = m >= 70 ? RED : m >= 20 ? ORANGE : TEAL;
    s.addText(k, { x: tx + 0.28, y, w: 0.45, h: 0.5, fontSize: 15, bold: true, color: c, valign: "middle", margin: 0, isTextBox: true });
    s.addText(n, { x: tx + 0.75, y, w: 2.3, h: 0.5, fontSize: 13.5, bold: true, color: NAVY, valign: "middle", margin: 0, isTextBox: true });
    const bx = tx + 3.1, bw = 1.6;
    s.addShape(pres.shapes.RECTANGLE, { x: bx, y: y + 0.15, w: bw, h: 0.2, fill: { color: "E3E7EE" }, line: { type: "none" } });
    s.addShape(pres.shapes.RECTANGLE, { x: bx, y: y + 0.15, w: Math.max(0.04, bw * m / 100), h: 0.2, fill: { color: c }, line: { type: "none" } });
    s.addText(`${m}% (${r})`, { x: bx + bw + 0.1, y, w: 1.2, h: 0.5, fontSize: 12.5, bold: true, color: NAVY, valign: "middle", margin: 0, isTextBox: true });
    s.addText(a, { x: tx + 6.0, y, w: 1.85, h: 0.5, fontSize: 11.5, color: BODY, valign: "middle", margin: 0, isTextBox: true });
  });
  card(s, { x: 8.85, y: 1.7, w: 3.88, h: 4.15, color: TEAL, kicker: "Tissue sampling", title: "FNA vs core", rows: [
    { l: "FNAC", t: "First-line, US-guided; reported as Bethesda", b: true },
    { l: "Limit", t: "Cannot separate follicular adenoma from carcinoma" },
    { l: "Core biopsy", t: "Repeat nondiagnostic; suspected anaplastic or lymphoma" } ] });
  remember(s, "Bethesda III–IV are 'indeterminate': molecular testing or diagnostic lobectomy settles them", 6.12, "Key: ");
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
    ["Benign nodule", "Observe with US follow-up; surgery if compressive, growing or indeterminate", "Thermal ablation; no levothyroxine suppression", TEAL],
    ["Differentiated", "Lobectomy or total; active surveillance for low-risk microcarcinoma", "Radioiodine and TSH suppression by risk", TEAL],
    ["Medullary", "Total thyroidectomy + central neck dissection", "Prophylactic surgery in RET carriers; no radioiodine", ORANGE],
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
// 17. Surgical management
// =====================================================================
{
  const s = content(6, "Surgery: lobectomy, thyroidectomy, neck dissection", "[8] ATA 2015; [1] Bailey & Love",
    "Papillary microcarcinoma (< 1 cm) without extrathyroidal extension or nodes: lobectomy, or active surveillance in selected patients (ATA 2015). Lobectomy is acceptable for 1–4 cm low-risk tumours; diagnostic lobectomy for indeterminate nodules.\n" +
    "Completion thyroidectomy: offered when the final histology after lobectomy shows higher-risk features (e.g. gross extrathyroidal extension, extensive vascular invasion, nodal disease) or when radioiodine becomes indicated.\n" +
    "Technique: identify and preserve the recurrent laryngeal nerve and parathyroids; intraoperative nerve monitoring is a useful adjunct.\n" +
    "Total thyroidectomy: tumour > 4 cm, gross extrathyroidal extension, clinical nodal or distant metastases, medullary carcinoma, and when radioiodine is planned.\n" +
    "Neck dissection: therapeutic central (level VI) dissection for clinically involved nodes; prophylactic central dissection may be considered for advanced primaries (T3–T4) or clinical lateral nodes; lateral compartmental dissection (II–V) for biopsy-proven lateral nodes. Routine central dissection in medullary carcinoma.\n" +
    "Complications to consent for: recurrent laryngeal nerve palsy, hypocalcaemia from hypoparathyroidism, neck haematoma (an airway emergency: open the wound at the bedside), and external branch of the superior laryngeal nerve injury (loss of high-pitched voice). Bilateral nerve palsy presents with stridor after extubation. Post-operative PTH or calcium guides calcium and vitamin D replacement.");
  card(s, { x: X0, y: 1.7, w: 3.88, h: 4.15, color: TEAL, kicker: "Option 1", title: "Lobectomy", rows: [
    { l: "< 1 cm, low risk", t: "Lobectomy, or active surveillance", b: true },
    { l: "1–4 cm, low risk", t: "Lobectomy or total thyroidectomy" },
    { l: "Completion", t: "If histology upstages after lobectomy" } ] });
  card(s, { x: X0 + 4.125, y: 1.7, w: 3.88, h: 4.15, color: RED, kicker: "Option 2", title: "Total thyroidectomy", rows: [
    { l: "Tumour", t: "> 4 cm or gross extrathyroidal extension", b: true },
    { l: "Spread", t: "Clinical nodes or distant metastases" },
    { l: "Other", t: "Medullary cancer; RAI planned; bilateral disease" } ] });
  card(s, { x: X0 + 8.25, y: 1.7, w: 3.88, h: 4.15, color: RED, kicker: "Nodes", title: "Neck dissection", rows: [
    { l: "Central (VI)", t: "Therapeutic for involved nodes; routine in medullary" },
    { l: "Prophylactic", t: "Consider in T3–T4 or lateral node disease" },
    { l: "Lateral (II–V)", t: "Only for proven lateral nodes", b: true } ] });
  remember(s, "RLN palsy (bilateral = stridor) · hypocalcaemia (check PTH) · haematoma: open the wound at the bedside", 6.12, "Complications: ");
}

// =====================================================================
// 18. Non-surgical options
// =====================================================================
{
  const s = content(6, "Non-surgical: radioiodine, radiotherapy, drugs", "[9] ATA 2025; [11] ATA ATC 2021; [18] LIBRETTO-001",
    "Radioactive iodine (I-131): used after total thyroidectomy for differentiated cancer only, as remnant ablation, adjuvant therapy, or treatment of known disease. Needs TSH stimulation (withdrawal or recombinant TSH) and a low-iodine diet. ATA 2025 is more selective for low-risk patients.\n" +
    "TSH suppression with levothyroxine: degree set by risk and response; over-suppression risks atrial fibrillation and bone loss.\n" +
    "External beam radiotherapy: anaplastic cancer, unresectable or gross residual disease, non-iodine-avid disease, palliation of bone metastases.\n" +
    "Systemic therapy: multikinase inhibitors (lenvatinib, sorafenib) for progressive radioiodine-refractory differentiated cancer; for medullary: vandetanib, cabozantinib, or the selective RET inhibitor selpercatinib; for BRAF V600E anaplastic cancer: dabrafenib + trametinib; lymphoma: chemotherapy.");
  card(s, { x: X0, y: 1.7, w: 2.92, h: 4.15, color: TEAL, kicker: "Option 1", title: "Radioiodine", rows: [
    { l: "Use", t: "I-131 after total thyroidectomy" }, { l: "Needs", t: "TSH stimulation; low-iodine diet" }, { l: "Only for", t: "Differentiated cancer; selective in low risk", b: true } ] });
  card(s, { x: X0 + 3.093, y: 1.7, w: 2.92, h: 4.15, color: TEAL, kicker: "Option 2", title: "TSH suppression", rows: [
    { l: "How", t: "Levothyroxine" }, { l: "Target", t: "Set by risk and response", b: true }, { l: "Risk", t: "Atrial fibrillation, bone loss" } ] });
  card(s, { x: X0 + 6.186, y: 1.7, w: 2.92, h: 4.15, color: ORANGE, kicker: "Option 3", title: "External beam RT", rows: [
    { l: "Anaplastic", t: "Local control, with drugs" }, { l: "Residual", t: "Gross residual or unresectable disease", b: true }, { l: "Palliation", t: "Painful bone metastases" } ] });
  card(s, { x: X0 + 9.279, y: 1.7, w: 2.92, h: 4.15, color: RED, kicker: "Option 4", title: "Systemic therapy", rows: [
    { l: "Iodine-refractory", t: "Lenvatinib, sorafenib" }, { l: "Medullary", t: "RET inhibitor (selpercatinib)" }, { l: "BRAF anaplastic", t: "Dabrafenib + trametinib", b: true } ] });
  remember(s, "medullary and anaplastic cancers do not take up iodine: radioiodine has no role", 6.12);
}

divider(7);
// =====================================================================
// 21. MY PART — Prognosis and outcomes (single slide)
// =====================================================================
{
  const s = content(7, "Prognosis: follow-up, risk factors and survival", "[8] ATA 2015; [9] ATA 2025; [12] AJCC 8th; [14] SEER; [16] Roman 2006; [11] ATA ATC 2021",
`SCRIPT (about 60 seconds)
"Prognosis in one slide. After treatment of differentiated cancer we follow thyroglobulin with its antibodies, plus neck ultrasound. At six to twelve months each patient is classed by response: an excellent response carries a recurrence risk of only one to four percent, while a structural incomplete response carries the highest risk. Medullary cancer is followed with calcitonin and CEA.
The main prognostic factors are age, with fifty-five years as the AJCC cut-off, distant metastases, gross extrathyroidal extension, aggressive histology, and incomplete resection.
Survival is excellent: in SEER data it is almost one hundred percent at five years for localised and regional disease, but about fifty percent with distant metastases. Anaplastic cancer is the exception, with a median survival of about five months. Recurrence is more common than death, so follow-up is long."

BACKUP Q&A (not read aloud)
Q: Thyroglobulin cut-offs for an excellent response?
A: After total thyroidectomy and radioiodine: suppressed Tg < 0.2 ng/mL or stimulated Tg < 1 ng/mL, negative antibodies and imaging (ATA 2015). Biochemical incomplete: suppressed Tg ≥ 1 or stimulated ≥ 10, or rising antibodies.

Q: Why measure anti-Tg antibodies?
A: They make thyroglobulin falsely low; a rising antibody level suggests persistent disease.

Q: TSH targets?
A: High risk or structural disease < 0.1 mU/L; biochemical incomplete 0.1–0.5; low risk with excellent response 0.5–2.0 (ATA 2015).

Q: What is MACIS?
A: Metastasis, Age, Completeness of resection, Invasion, Size. 20-year survival: < 6 = 99%, 6–6.99 = 89%, 7–7.99 = 56%, ≥ 8 = 24% (Hay 1993).

Q: AJCC vs ATA?
A: AJCC 8th staging predicts death (age cut-off 55: under 55 is stage I, or II if M1). ATA risk predicts recurrence; ATA 2025 uses four tiers: < 10%, 10–15%, 16–30%, > 30%.

Q: Medullary and anaplastic survival?
A: Medullary 10-year: localized 96%, regional 76%, distant 40% (Roman 2006); calcitonin/CEA doubling time < 6 months is poor. Anaplastic: median about 5 months, 1-year about 20%.

Q: ATA 2025 de-escalation?
A: Low-risk patients with an excellent response can stop ultrasound after 5–8 years and thyroglobulin after 10–15 years.`);
  // Column 1: follow-up
  const resp = [["Excellent", "Recurrence ~1–4%", TEAL], ["Indeterminate", "Most stay stable", "B08A1A"], ["Biochemical incomplete", "Raised Tg, imaging negative", ORANGE], ["Structural incomplete", "Visible disease: highest risk", RED]];
  box(s, X0, 1.65, 3.95, 4.2); topbar(s, X0, 1.65, 3.95, TEAL);
  label(s, "Follow-up · response to therapy", X0 + 0.25, 1.92, 3.5, TEAL, 10);
  s.addText("Tg + anti-Tg antibodies and neck US at 6–12 months", { x: X0 + 0.25, y: 2.22, w: 3.5, h: 0.5, fontSize: 12.5, color: BODY, margin: 0, valign: "top", isTextBox: true });
  resp.forEach(([t, d, c], i) => {
    const y = 2.8 + i * 0.75;
    s.addShape(pres.shapes.RECTANGLE, { x: X0 + 0.25, y, w: 0.09, h: 0.62, fill: { color: c }, line: { type: "none" } });
    s.addText([{ text: t, options: { bold: true, color: NAVY, fontSize: 14, breakLine: true } }, { text: d, options: { color: BODY, fontSize: 12.5 } }], { x: X0 + 0.45, y, w: 3.3, h: 0.62, valign: "middle", margin: 0, isTextBox: true });
  });
  // Column 2: prognostic factors
  card(s, { x: X0 + 4.1, y: 1.65, w: 3.95, h: 4.2, color: RED, kicker: "Prognostic factors", title: null, rows: [
    { l: "Age", t: "AJCC cut-off 55 years", b: true },
    { l: "Spread", t: "Distant metastases (worst)", b: true },
    { l: "Tumour", t: "Gross ETE; aggressive histology" },
    { l: "Treatment", t: "Incomplete resection" } ] });
  // Column 3: survival chart
  const cx = X0 + 8.2, cw = 3.93;
  box(s, cx, 1.65, cw, 4.2); topbar(s, cx, 1.65, cw, NAVY);
  label(s, "5-year survival by stage (SEER)", cx + 0.25, 1.92, 3.5, NAVY, 10);
  s.addChart(pres.charts.BAR, [{ name: "5-year survival", labels: ["Localized", "Regional", "Distant"], values: [99.9, 98.1, 50.2] }], {
    x: cx + 0.1, y: 2.2, w: cw - 0.2, h: 2.75, barDir: "col", chartColors: [TEAL, TEAL, RED], barGapWidthPct: 55,
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 13, dataLabelFontBold: true, dataLabelColor: NAVY, dataLabelFormatCode: '0.0"%"',
    valAxisHidden: true, valAxisMaxVal: 118, valAxisMinVal: 0, valGridLine: { style: "none" }, catGridLine: { style: "none" },
    catAxisLabelColor: NAVY, catAxisLabelFontSize: 12, catAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", showLegend: false, showTitle: false,
  });
  s.addText([{ text: "Anaplastic: ", options: { bold: true, color: RED } }, { text: "median survival ≈ 5 months", options: { color: BODY } }], { x: cx + 0.25, y: 5.08, w: cw - 0.45, h: 0.6, fontSize: 13, valign: "middle", margin: 0, isTextBox: true });
  remember(s, "recurrence is more common than death: most recur in neck nodes within the first decade", 6.12, "Key: ");
}

divider(8);
// =====================================================================
// 23. MY PART — Case discussion (single slide)
// =====================================================================
{
  const s = content(8, "Case discussion: three nodules, three outcomes", "[7] Bethesda 2023; [8] ATA 2015; [10] ATA MTC 2015; [11] ATA ATC 2021",
`SCRIPT (about 60 seconds)
"Three short illustrative cases.
First, a thirty-four-year-old woman with a two-centimetre TI-RADS five nodule; FNA is Bethesda six: papillary cancer. She is under fifty-five, so stage one and low risk. Thyroidectomy, no routine radioiodine; at one year thyroglobulin is undetectable: an excellent response.
Second, a forty-two-year-old man with a nodule, a neck node, diarrhoea and a family history of thyroid cancer and severe hypertension: medullary cancer, probably MEN2A. Exclude phaeochromocytoma first, then total thyroidectomy with neck dissection, and follow calcitonin and CEA.
Third, a seventy-four-year-old woman with a hard mass growing over three weeks and hoarseness. Think anaplastic cancer, but exclude lymphoma by core biopsy and test BRAF. She needs airway assessment, urgent multidisciplinary care and early palliative input.
Thank you."

BACKUP Q&A (not read aloud)
Q: Case 1: why no routine radioiodine?
A: In ATA low-risk disease it does not improve survival or meaningfully reduce recurrence; ATA 2025 is even more selective.

Q: Case 1: what if Tg were 5 ng/mL with a clear ultrasound?
A: Biochemical incomplete response: keep TSH mildly suppressed (0.1–0.5), follow the trend, image if rising.

Q: Case 2: why exclude phaeochromocytoma first?
A: In MEN2A about half develop phaeochromocytoma; surgery can trigger a catecholamine crisis. Adrenalectomy with alpha-blockade comes first.

Q: Case 2: why no radioiodine?
A: C cells do not take up iodine and are not TSH-driven.

Q: Case 2: family?
A: Offer RET testing; carriers get prophylactic thyroidectomy timed by mutation risk (e.g. M918T/MEN2B within the first year; codon 634 by age 5).

Q: Case 3: how to tell anaplastic from lymphoma?
A: Core biopsy with immunohistochemistry (± flow cytometry); lymphoma often arises in Hashimoto and is very treatable.

Q: Case 3: treatment?
A: Resect only if complete resection is realistic; dabrafenib + trametinib if BRAF V600E; radiotherapy; early palliative care. Median survival about 5 months.`);
  const cases = [
    { k: "Case 1", t: "Woman, 34", c: TEAL, sc: "2 cm nodule, TI-RADS 5; FNA Bethesda VI", th: "Papillary cancer; under 55: stage I, low risk", de: "Thyroidectomy; no routine RAI. Tg undetectable at 1 year: excellent response" },
    { k: "Case 2", t: "Man, 42", c: ORANGE, sc: "Nodule + neck node; diarrhoea; father: thyroid cancer, severe hypertension", th: "Medullary cancer, likely MEN2A", de: "Exclude phaeochromocytoma → total thyroidectomy + neck dissection; follow calcitonin/CEA" },
    { k: "Case 3", t: "Woman, 74", c: RED, sc: "Hard mass growing over 3 weeks; hoarseness, breathlessness", th: "Anaplastic vs lymphoma: core biopsy, BRAF", de: "Airway; urgent MDT; BRAF-targeted therapy; early palliative care" },
  ];
  cases.forEach((c, i) => {
    const x = X0 + i * 4.1, w = 3.93;
    box(s, x, 1.65, w, 5.15); topbar(s, x, 1.65, w, c.c);
    label(s, c.k, x + 0.25, 1.9, 2, c.c, 10.5);
    s.addText(c.t, { x: x + 0.25, y: 2.17, w: w - 0.5, h: 0.4, fontSize: 18, bold: true, color: NAVY, valign: "middle", margin: 0, isTextBox: true });
    const ix = x + 0.2, iw = w - 0.4;
    [[2.7, 1.3, WHITE, BORDER, "Scenario", TEAL, BODY, c.sc, false], [4.07, 1.2, TEAL_T, "BFDDD8", "Think", TEAL, BODY, c.th, false], [5.34, 1.32, NAVY, NAVY, "Decide", MINT, WHITE, c.de, true]].forEach(([y, h, f, l, lab, lc, tc, t, b]) => {
      s.addShape(pres.shapes.RECTANGLE, { x: ix, y, w: iw, h, fill: { color: f }, line: { color: l, width: 0.75 } });
      label(s, lab, ix + 0.13, y + 0.06, 2, lc, 9.5);
      s.addText(t, { x: ix + 0.13, y: y + 0.32, w: iw - 0.26, h: h - 0.38, fontSize: 13.5, bold: b, color: tc, valign: "top", margin: 0, isTextBox: true });
    });
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
