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
  s.addText(title, title.length > 52 ? { placeholder: "title", fontSize: 23 } : { placeholder: "title" });
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

// ===== helpers for teammate content =====
function tbl(s, x, y, colW, rows, o = {}) {
  const rh = o.rowH || 0.62, hh = o.headH || 0.5, fs = o.fontSize || 14;
  let yy = y;
  rows.forEach((r, i) => {
    const head = i === 0 && !o.noHeader;
    const h = head ? hh : rh;
    s.addShape(pres.shapes.RECTANGLE, { x, y: yy, w: colW.reduce((a, b) => a + b, 0), h, fill: { color: head ? NAVY : (i % 2 ? WHITE : ROWALT) }, line: { color: head ? NAVY : BORDER, width: 0.5 } });
    let cx = x;
    r.forEach((t, j) => {
      s.addText(t, { x: cx + 0.15, y: yy, w: colW[j] - 0.25, h, fontSize: head ? 12 : fs, bold: head || j === 0, color: head ? WHITE : (j === 0 ? NAVY : BODY), charSpacing: head ? 1 : 0, valign: "middle", margin: 0, isTextBox: true });
      cx += colW[j];
    });
    yy += h;
  });
  return yy;
}
function plainCard(s, x, y, w, h, color, kicker, items, o = {}) {
  box(s, x, y, w, h); topbar(s, x, y, w, color);
  label(s, kicker, x + 0.28, y + 0.28, w - 0.5, color, 10.5);
  const top = y + 0.68;
  s.addText(bullets(items, o.size || 15), { x: x + 0.28, y: top, w: w - 0.5, h: y + h - top - 0.15, valign: "top", margin: 0, isTextBox: true });
}
function img(s, file, x, y, w, ar, alt) {
  const h = w / ar;
  box(s, x, y, w, h);
  s.addImage({ path: path.join(__dirname, file), x: x + 0.03, y: y + 0.03, w: w - 0.06, h: h - 0.06, altText: alt });
  return h;
}
function blank(secIdx) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: SECTIONS[secIdx] });
  s.addNotes(`Placeholder: ${SECTIONS[secIdx]} content to be added here.`);
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
// 3. Intro — Introduction + Benign vs malignant (teammate content)
// =====================================================================
{
  const s = content(0, "Introduction to Thyroid Tumors", null, "Teammate content (Objective 1). Sources: [1] ATA 2025; [2] ETA 2023.");
  plainCard(s, X0, 1.65, 4.75, 5.15, TEAL, "Definition", [
    "Thyroid tumors are abnormal growths arising from thyroid gland cells.",
    "They can be benign or malignant.",
    "They may present as a solitary nodule, multiple nodules, or an enlarged thyroid gland.",
    [{ text: "Most thyroid nodules are benign; only a minority are malignant.", b: true }],
  ], { size: 17 });
  label(s, "Benign vs. Malignant Thyroid Tumors", 5.6, 1.65, 7, NAVY, 10.5);
  tbl(s, 5.6, 2.0, [1.65, 2.35, 3.13], [
    ["Feature", "Benign", "Malignant"],
    ["Nature", "Non-cancerous", "Cancerous"],
    ["Growth", "Usually localized", "Variable, potentially invasive"],
    ["Local invasion", "Absent", "May be present"],
    ["Metastasis", "Absent", "Possible"],
    ["Examples", "Follicular adenoma, cyst, colloid nodule", "Papillary, follicular, medullary, anaplastic carcinoma"],
  ], { rowH: 0.8 });
}

// =====================================================================
// 4. Intro — Epidemiology + Risk factors (teammate content)
// =====================================================================
{
  const s = content(0, "Epidemiology and Risk Factors", null, "Teammate content (Objective 1). Sources: [2] ETA 2023; [5] Saudi Cancer Incidence Report 2023.");
  plainCard(s, X0, 1.65, 5.6, 3.55, TEAL, "Global epidemiology", [
    "Thyroid nodules are common, particularly in adults.",
    "Ultrasound detects thyroid nodules in approximately 19–68% of selected adult populations.",
    "Thyroid cancer is more common in females than males.",
    "Papillary thyroid carcinoma is the most common type of thyroid cancer.",
    "Increased detection through imaging contributes to rising reported incidence.",
  ], { size: 13.5 });
  label(s, "Saudi Arabia – 2023", X0, 5.38, 4, RED, 10.5);
  [["2nd", "Among cancers in Saudi females"], ["7th", "Among cancers in Saudi males"]].forEach(([n, t], i) => {
    const x = X0 + i * 2.85;
    box(s, x, 5.72, 2.75, 1.08);
    s.addText(n, { x: x + 0.2, y: 5.72, w: 0.95, h: 1.08, fontFace: HEAD, fontSize: 26, bold: true, color: RED, valign: "middle", margin: 0, isTextBox: true });
    s.addText(t, { x: x + 1.15, y: 5.72, w: 1.5, h: 1.08, fontSize: 13, color: BODY, valign: "middle", margin: 0, isTextBox: true });
  });
  label(s, "Risk Factors", 6.45, 1.65, 6, NAVY, 10.5);
  tbl(s, 6.45, 2.0, [2.2, 4.08], [
    ["Risk factor", "Clinical significance"],
    ["Ionizing radiation", "Especially childhood exposure to head and neck radiation"],
    ["Family history", "First-degree relatives with thyroid cancer"],
    ["Genetic mutations", "RET mutations, MEN2 syndromes"],
    ["Female sex", "Thyroid cancer is more frequently diagnosed in females"],
    ["Increasing age", "Thyroid nodules become more common with age"],
    ["Iodine deficiency", "Associated with goiter and nodular thyroid disease"],
  ], { rowH: 0.72 });
}

divider(1);
// =====================================================================
// 5. Anatomy — Gross anatomy
// =====================================================================
{
  const s = content(1, "Gross Anatomy of the Thyroid Gland", null, "Teammate content (Objective 2). Sources: [3] Gray's Anatomy; [4] Schwartz.");
  box(s, X0, 1.65, 5.3, 5.15); topbar(s, X0, 1.65, 5.3, TEAL);
  s.addText([
    ...bullets(["Butterfly-shaped endocrine gland located in the anterior neck.", "Lies anterior and lateral to the upper trachea.", "Extends approximately from C5 to T1.", "Consists of:"], 15),
    { text: "", options: { breakLine: true } },
    ...["Right lobe.", "Left lobe.", "Isthmus.", "Pyramidal lobe (variable)."].map((t) => ({ text: t, options: { bullet: { indent: 14 }, indentLevel: 1, fontSize: 14, color: BODY, breakLine: true, paraSpaceAfter: 3 } })),
    { text: "Moves during swallowing due to its attachment to the larynx and trachea.", options: { bullet: { indent: 14 }, fontSize: 15, bold: true, color: NAVY } },
  ], { x: X0 + 0.28, y: 2.0, w: 4.8, h: 4.65, valign: "top", margin: 0, isTextBox: true });
  img(s, "team_img1.jpg", 6.15, 1.65, 6.58, 2048 / 1506, "Labelled anterior view of the thyroid gland and larynx");
}

// =====================================================================
// 6. Anatomy — Blood supply
// =====================================================================
{
  const s = content(1, "Blood Supply of the Thyroid Gland", null, "Teammate content (Objective 2). Sources: [3] Gray's Anatomy.");
  label(s, "Arterial supply", X0, 1.65, 5, RED, 10.5);
  tbl(s, X0, 1.98, [2.55, 2.85], [
    ["Artery", "Origin"],
    ["Superior thyroid artery", "External carotid artery"],
    ["Inferior thyroid artery", "Thyrocervical trunk"],
    ["Thyroid ima artery", "Brachiocephalic trunk or aortic arch (variable)"],
  ], { rowH: 0.7 });
  label(s, "Venous drainage", X0, 4.65, 5, "3E6FB0", 10.5);
  tbl(s, X0, 4.98, [2.55, 2.85], [
    ["Vein", "Drains into"],
    ["Superior thyroid vein", "Internal jugular vein"],
    ["Middle thyroid vein", "Internal jugular vein"],
    ["Inferior thyroid veins", "Brachiocephalic veins"],
  ], { rowH: 0.45, headH: 0.42 });
  img(s, "team_img2.jpg", 6.3, 1.65, 6.43, 2228 / 1581, "Arterial blood supply of the thyroid gland, anterior view");
}

// =====================================================================
// 7. Anatomy — Lymphatic drainage
// =====================================================================
{
  const s = content(1, "Lymphatic Drainage", null, "Teammate content (Objective 2). Sources: [3] Gray's Anatomy; [4] Schwartz.");
  tbl(s, X0, 1.9, [2.9, 3.1], [
    ["Lymph node group", ""],
    ["Central compartment (Level VI)", "Prelaryngeal, pretracheal, paratracheal nodes"],
    ["Lateral compartment (Levels II–V)", "Cervical jugular and posterior triangle nodes"],
    ["Superior mediastinal (Level VII)", "Upper mediastinal nodes"],
  ], { rowH: 1.0 });
  img(s, "team_img3.jpg", 6.85, 1.9, 5.88, 611 / 363, "Cervical lymph node levels and superior mediastinal nodes");
}

// =====================================================================
// 8. Anatomy — Relations and surgical importance
// =====================================================================
{
  const s = content(1, "Anatomical Relations and Surgical Importance", null, "Teammate content (Objective 2). Sources: [4] Schwartz.");
  tbl(s, X0, 1.65, [2.75, 3.65], [
    ["Anatomical structure", "Surgical importance"],
    ["Recurrent laryngeal nerve (RLN)", "Injury may cause vocal cord paralysis and hoarseness"],
    ["External branch of superior laryngeal nerve", "Injury may affect voice pitch and projection"],
    ["Parathyroid glands", "Injury or devascularization may cause hypocalcemia"],
    ["Trachea", "May be compressed or invaded by advanced tumors"],
    ["Esophagus", "May be compressed or invaded by advanced tumors"],
  ], { rowH: 0.88 });
  img(s, "team_img4.jpg", 7.25, 1.65, 5.48, 1000 / 747, "Relations of the thyroid: recurrent laryngeal nerve, external branch of superior laryngeal nerve, parathyroid glands");
}

// =====================================================================
// 9. Physiology — functions + HPT axis
// =====================================================================
{
  const s = content(1, "Thyroid Hormone Regulation (HPT Axis)", null, "Teammate content (Objective 2). Sources: [6] Williams Textbook of Endocrinology.");
  plainCard(s, X0, 1.65, 5.3, 2.5, TEAL, "Main physiological functions of T3 and T4", [
    "Increase basal metabolic rate.", "Increase heart rate and cardiac output.", "Support normal growth and development.", "Regulate body temperature.", "Maintain normal neurological function.",
  ], { size: 13 });
  plainCard(s, X0, 4.3, 5.3, 2.5, RED, "Hormonal regulation", [
    "Hypothalamus releases TRH.", "Anterior pituitary releases TSH.", "TSH stimulates thyroid hormone production.", "Thyroid gland releases T3 and T4.", "Increased T3 and T4 suppress TRH and TSH through negative feedback.",
  ], { size: 13 });
  img(s, "team_img5.jpg", 6.1, 2.3, 6.63, 1408 / 768, "Hypothalamic–pituitary–thyroid axis with negative feedback");
}

divider(2);
// =====================================================================
// 10. Classification (WHO 2022) — verbatim
// =====================================================================
{
  const s = content(2, "Classification of Thyroid Tumors (WHO 2022) [7]", null, "Sources: [7] Jung CK, Bychkov A, Kakudo K. Endocrinol Metab 2022 (2022 WHO classification).");
  plainCard(s, X0, 1.65, 4.8, 2.9, TEAL, "A. Benign Tumors / Lesions", ["Follicular adenoma", "Oncocytic adenoma", "Thyroid cysts (non-neoplastic lesions)"], { size: 16 });
  plainCard(s, X0, 4.7, 4.8, 2.1, ORANGE, "B. Low-risk Neoplasms", ["Non-invasive follicular thyroid neoplasm with papillary-like nuclear features (NIFTP)"], { size: 15 });
  const cx = 5.6, cw = 7.13;
  box(s, cx, 1.65, cw, 5.15); topbar(s, cx, 1.65, cw, RED);
  label(s, "C. Malignant Tumors", cx + 0.28, 1.93, 5, RED, 10.5);
  const sub = (t, y, items, cols) => {
    s.addText(t, { x: cx + 0.28, y, w: cw - 0.5, h: 0.36, fontSize: 15, bold: true, color: NAVY, margin: 0, valign: "middle", isTextBox: true });
    const per = Math.ceil(items.length / cols), colw = (cw - 0.6) / cols;
    for (let c = 0; c < cols; c++) {
      const part = items.slice(c * per, (c + 1) * per);
      if (part.length) s.addText(bullets(part, 14), { x: cx + 0.35 + c * colw, y: y + 0.4, w: colw - 0.1, h: per * 0.42, valign: "top", margin: 0, isTextBox: true });
    }
  };
  sub("1. Follicular Cell-Derived:", 2.3, ["Papillary thyroid carcinoma (PTC)", "Follicular thyroid carcinoma (FTC)", "Oncocytic carcinoma", "Poorly differentiated thyroid carcinoma (PDTC)", "Differentiated high-grade thyroid carcinoma (DHGTC)", "Anaplastic thyroid carcinoma (ATC)"], 1);
  s.addShape(pres.shapes.LINE, { x: cx + 0.28, y: 5.25, w: cw - 0.56, h: 0, line: { color: BORDER, width: 0.75 } });
  sub("2. C-cell-Derived:", 5.32, ["Medullary thyroid carcinoma (MTC)"], 1);
  s.addText("3. Other Malignancies:", { x: cx + 3.6, y: 5.32, w: 3.3, h: 0.36, fontSize: 15, bold: true, color: NAVY, margin: 0, valign: "middle", isTextBox: true });
  s.addText(bullets(["Primary thyroid lymphoma", "Metastatic tumors (e.g., kidney, lung, breast)"], 14), { x: cx + 3.67, y: 5.72, w: 3.35, h: 0.95, valign: "top", margin: 0, isTextBox: true });
}

divider(3);
// =====================================================================
// 11. Pathophysiology — verbatim
// =====================================================================
{
  const s = content(3, "Pathophysiology of Thyroid Tumors", null, "Teammate content (Objective 4). Sources: [7] Jung 2022.\nPresentation tip: keep the slide text concise and explain the details verbally.");
  plainCard(s, X0, 1.65, 6.6, 5.15, TEAL, "Pathophysiology", [
    "Thyroid tumors develop due to genetic mutations and abnormal proliferation of thyroid cells.",
    "Accumulation of genetic alterations leads to uncontrolled cell growth and tumor formation.",
    "Most thyroid carcinomas originate from follicular cells.",
    "Medullary thyroid carcinoma originates from parafollicular C cells.",
    "Anaplastic carcinoma is associated with loss of cellular differentiation (dedifferentiation).",
  ], { size: 17 });
  const kx = 7.4, kw = 5.33;
  box(s, kx, 1.65, kw, 5.15); topbar(s, kx, 1.65, kw, RED);
  label(s, "Key Genetic Mutations", kx + 0.28, 1.93, 4.5, RED, 10.5);
  [["Papillary carcinoma", "BRAF, RET/PTC"], ["Follicular carcinoma", "RAS, PAX8-PPARG"], ["Medullary carcinoma", "RET"], ["Anaplastic carcinoma", "TP53"]].forEach(([t, g], i) => {
    const y = 2.4 + i * 1.08;
    s.addShape(pres.shapes.RECTANGLE, { x: kx + 0.28, y, w: kw - 0.56, h: 0.92, fill: { color: i % 2 ? WHITE : ROWALT }, line: { color: BORDER, width: 0.5 } });
    s.addText(t + ":", { x: kx + 0.45, y, w: 2.6, h: 0.92, fontSize: 15, bold: true, color: NAVY, valign: "middle", margin: 0, isTextBox: true });
    s.addText(g + ".", { x: kx + 3.05, y, w: 1.9, h: 0.92, fontSize: 16, bold: true, color: RED, valign: "middle", margin: 0, isTextBox: true });
  });
}

// =====================================================================
// 12. Patterns of spread — verbatim
// =====================================================================
{
  const s = content(3, "Patterns of Spread of Thyroid Malignancies", null, "Teammate content (Objective 4). Sources: [7] Jung 2022.\nPresentation tip: emphasize the contrast between lymphatic spread in papillary carcinoma and hematogenous spread in follicular carcinoma.");
  s.addText("Thyroid cancers spread through three main pathways:", { x: X0, y: 1.55, w: 10, h: 0.35, fontSize: 15, color: BODY, margin: 0, isTextBox: true });
  [["Direct invasion", "Invasion of surrounding structures such as the trachea, esophagus, and recurrent laryngeal nerve.", RED],
   ["Lymphatic spread", "Spread to regional cervical lymph nodes.", TEAL],
   ["Hematogenous spread", "Spread through blood vessels to distant organs.", ORANGE]].forEach(([t, d, c], i) => {
    const x = X0 + i * 4.1, w = 3.93;
    box(s, x, 2.0, w, 1.72); topbar(s, x, 2.0, w, c);
    s.addText(t + ":", { x: x + 0.25, y: 2.2, w: w - 0.45, h: 0.38, fontSize: 16, bold: true, color: NAVY, valign: "middle", margin: 0, isTextBox: true });
    s.addText(d, { x: x + 0.25, y: 2.62, w: w - 0.45, h: 1.0, fontSize: 13.5, color: BODY, valign: "top", margin: 0, isTextBox: true });
  });
  label(s, "Patterns of Spread According to Tumor Type", X0, 3.95, 8, NAVY, 10.5);
  tbl(s, X0, 4.27, [2.4, 4.4, 5.33], [
    ["Tumor Type", "Main Pattern of Spread", "Common Sites"],
    ["Papillary", "Lymphatic", "Cervical lymph nodes"],
    ["Follicular", "Hematogenous", "Bones, lungs"],
    ["Medullary", "Both lymphatic and hematogenous", "Lymph nodes, liver, lungs, bones"],
    ["Anaplastic", "Direct invasion and distant metastasis", "Trachea, esophagus, lungs, bones, brain"],
  ], { rowH: 0.5, headH: 0.45 });
}

divider(4);
// ===== Objectives 5 and 6 (teammate content, rebuilt from the final PDF; source lines removed) =====
function subtitle(s, t) {
  s.addText(t, { x: X0, y: 1.32, w: 11, h: 0.3, fontSize: 12.5, color: TEAL, charSpacing: 1, margin: 0, valign: "middle", isTextBox: true });
}
function banner(s, lead, t, y = 6.3, h = 0.58) {
  box(s, X0, y, CW, h, NAVY, NAVY);
  s.addText([{ text: lead + "  /  ", options: { bold: true, color: WHITE } }, { text: t, options: { bold: true, color: WHITE } }],
    { x: X0 + 0.25, y, w: CW - 0.4, h, fontSize: 15, valign: "middle", margin: 0, isTextBox: true });
}
function footnote(s, t) {
  s.addText(t, { x: X0, y: 6.93, w: 10.5, h: 0.28, fontSize: 10, color: MUTED, margin: 0, valign: "middle", isTextBox: true });
}
function sqBullets(items, size) {
  return items.map((t, i) => ({ text: t, options: { bullet: { code: "25AA", indent: 16 }, fontSize: size, color: BODY, breakLine: i < items.length - 1, paraSpaceAfter: 12 } }));
}

// ---- Slide: Clinical Presentation of Thyroid Tumors ----
{
  const s = content(4, "Clinical Presentation of Thyroid Tumors", null, "Teammate content (Objective 5). Sources: [1] ATA 2025; [2] ETA 2023; [21] Moorthy 2026.");
  subtitle(s, "COMMON FINDINGS  /  SUSPICIOUS CLUES  /  LOCAL INVASION");
  [["COMMON PRESENTATION", TEAL, ["Painless thyroid nodule or neck swelling", "Frequently found incidentally", "Most patients are euthyroid", "Mass may move with swallowing"]],
   ["MALIGNANCY WARNING SIGNS", RED, ["Firm or fixed mass; progressive growth", "Suspicious cervical lymph nodes", "Persistent hoarseness or voice change", "Childhood irradiation or thyroid cancer family history"]],
   ["LOCAL / ADVANCED DISEASE", ORANGE, ["Dysphagia: possible esophageal involvement", "Dyspnea or stridor: airway compromise", "Voice change: possible nerve invasion", "Pain can occur but is nonspecific"]]].forEach(([h, c, items], i) => {
    const x = X0 + i * 4.1, w = 3.93, y = 1.78, hh = 3.5;
    box(s, x, y, w, hh); topbar(s, x, y, w, c);
    s.addText(h, { x: x + 0.25, y: y + 0.25, w: w - 0.4, h: 0.4, fontSize: 16, bold: true, color: NAVY, margin: 0, valign: "middle", isTextBox: true });
    s.addText(sqBullets(items, 14.5), { x: x + 0.25, y: y + 0.85, w: w - 0.4, h: hh - 1.0, valign: "top", margin: 0, isTextBox: true });
  });
  box(s, X0, 5.42, CW, 0.78, TEAL_T, "BFDDD8");
  s.addText([{ text: "CLINICAL ACTION", options: { bold: true, color: TEAL, fontSize: 14, breakLine: true } }, { text: "Abnormal voice: laryngeal examination when indicated.", options: { color: BODY, fontSize: 13.5 } }], { x: X0 + 0.25, y: 5.42, w: 5.8, h: 0.78, valign: "middle", margin: 0, isTextBox: true });
  s.addText([{ text: "AIRWAY RED FLAG", options: { bold: true, color: RED, fontSize: 14, breakLine: true } }, { text: "Stridor or respiratory distress needs urgent assessment.", options: { color: BODY, fontSize: 13.5 } }], { x: X0 + 6.35, y: 5.42, w: 5.6, h: 0.78, valign: "middle", margin: 0, isTextBox: true });
  banner(s, "REMEMBER", "Rapid swelling can also reflect hemorrhage into a benign nodule.");
}

// ---- Slide: Clinical Features by Tumor Type ----
{
  const s = content(4, "Clinical Features by Tumor Type", null, "Teammate content (Objective 5). Sources: [2] ETA 2023; [15] ATA MTC 2015; [21] Moorthy 2026.");
  subtitle(s, "PRESENTATION VARIES WITH TUMOR BIOLOGY AND SPREAD");
  const rows = [
    ["Papillary carcinoma", "Usually painless; a cervical node can be the first sign.", "Lymphatic spread to neck nodes", TEAL],
    ["Follicular carcinoma", "Usually solitary painless nodule; distant symptoms are later.", "Blood-borne spread to bone/lung", TEAL],
    ["Medullary carcinoma", "Nodule or neck nodes; diarrhea/flushing are uncommon.", "May be hereditary (MEN type 2)", TEAL],
    ["Anaplastic carcinoma", "Rapidly enlarging hard mass with compressive symptoms.", "Urgent airway assessment", RED],
    ["Thyroid lymphoma", "Rapidly enlarging thyroid; may cause dysphagia or dyspnea.", "Often with Hashimoto thyroiditis", RED],
  ];
  const colX = [X0 + 0.25, X0 + 3.0, X0 + 8.4], colW = [2.7, 5.3, 3.6];
  box(s, X0, 1.78, CW, 0.5, NAVY, NAVY);
  ["TUMOR TYPE", "TYPICAL PRESENTATION", "IMPORTANT CLUE"].forEach((t, j) => s.addText(t, { x: colX[j], y: 1.78, w: colW[j], h: 0.5, fontSize: 13, bold: true, color: WHITE, valign: "middle", margin: 0, isTextBox: true }));
  rows.forEach(([a, b, c, col], i) => {
    const y = 2.28 + i * 0.6;
    box(s, X0, y, CW, 0.6, i % 2 ? ROWALT : WHITE, BORDER);
    s.addShape(pres.shapes.RECTANGLE, { x: X0, y, w: 0.08, h: 0.6, fill: { color: col }, line: { type: "none" } });
    [a, b, c].forEach((t, j) => s.addText(t, { x: colX[j], y, w: colW[j], h: 0.6, fontSize: j === 0 ? 14.5 : 13.5, bold: j === 0, color: j === 0 ? NAVY : BODY, valign: "middle", margin: 0, isTextBox: true }));
  });
  box(s, X0, 5.3, CW, 0.5, "F6ECEA", "F6ECEA");
  s.addText("MEN type 2 = multiple endocrine neoplasia type 2.  Suspected thyroid lymphoma may need core biopsy.", { x: X0 + 0.25, y: 5.3, w: CW - 0.4, h: 0.5, fontSize: 13, color: BODY, valign: "middle", margin: 0, isTextBox: true });
  banner(s, "CLINICAL TRAP", "Rapid growth: consider anaplastic cancer, lymphoma, or nodule hemorrhage.", 5.98, 0.6);
}

divider(5);
// ---- Slide: Diagnostic Approach to a Thyroid Nodule ----
{
  const s = content(5, "Diagnostic Approach to a Thyroid Nodule", null, "Teammate content (Objective 6). Sources: [1] ATA 2025; [2] ETA 2023; [15] ATA MTC 2015; [21] Moorthy 2026.");
  subtitle(s, "INITIAL EVALUATION  /  FUNCTION  /  STRUCTURAL RISK  /  TARGETED TESTS");
  const sub = (x, y, w, h, fill, head, headColor, lines) => {
    box(s, x, y, w, h, fill, fill);
    const runs = [{ text: head, options: { bold: true, color: headColor, fontSize: 14.5, breakLine: true } }];
    lines.forEach((l, i) => runs.push({ text: l, options: { color: BODY, fontSize: 13, breakLine: i < lines.length - 1, paraSpaceAfter: 3 } }));
    s.addText(runs, { x: x + 0.15, y, w: w - 0.3, h, valign: "middle", margin: 0, isTextBox: true });
  };
  const lx = X0, lw = 5.85, rx = X0 + 6.05, rw = 6.08, y0 = 1.78, hh = 3.35;
  box(s, lx, y0, lw, hh); topbar(s, lx, y0, lw, TEAL);
  s.addText("1. FIRST-LINE EVALUATION", { x: lx + 0.22, y: y0 + 0.2, w: lw - 0.4, h: 0.38, fontSize: 16, bold: true, color: NAVY, margin: 0, valign: "middle", isTextBox: true });
  sub(lx + 0.22, y0 + 0.68, lw - 0.44, 0.78, "EEF4F4", "HISTORY + EXAMINATION", NAVY, ["Growth, risk factors, voice, compression and neck nodes"]);
  sub(lx + 0.22, y0 + 1.54, lw - 0.44, 0.78, "EEF4F4", "NECK ULTRASOUND", NAVY, ["Nodule morphology, size and central/lateral neck nodes"]);
  sub(lx + 0.22, y0 + 2.4, lw - 0.44, 0.78, "EEF4F4", "SERUM TSH", NAVY, ["Thyroid-stimulating hormone assesses function"]);
  box(s, rx, y0, rw, hh); topbar(s, rx, y0, rw, ORANGE);
  s.addText("2. INTERPRET TSH + ULTRASOUND", { x: rx + 0.22, y: y0 + 0.2, w: rw - 0.4, h: 0.38, fontSize: 16, bold: true, color: NAVY, margin: 0, valign: "middle", isTextBox: true });
  sub(rx + 0.22, y0 + 0.68, rw - 0.44, 1.18, TEAL_T, "LOW TSH", TEAL, ["Check free thyroxine (T4); free triiodothyronine (T3) if needed.", "Radionuclide scan: a functioning (hot) nodule rarely needs biopsy."]);
  sub(rx + 0.22, y0 + 1.96, rw - 0.44, 1.22, "F7F1EA", "NORMAL / HIGH TSH", NAVY, ["Ultrasound risk category + size determine needle aspiration.", "If high TSH: free T4; assess thyroid autoimmunity as appropriate."]);
  box(s, X0, 5.25, CW, 0.95); topbar(s, X0, 5.25, CW, RED);
  s.addText("3. SELECTIVE ADDITIONAL TESTS", { x: X0 + 0.22, y: 5.38, w: 6, h: 0.3, fontSize: 14.5, bold: true, color: NAVY, margin: 0, valign: "middle", isTextBox: true });
  s.addText("Voice abnormality: laryngoscopy. Suspected invasion or bulky nodes: contrast-enhanced CT/MRI.", { x: X0 + 0.22, y: 5.68, w: 5.8, h: 0.48, fontSize: 12.5, color: BODY, margin: 0, valign: "top", isTextBox: true });
  s.addText("Suspected medullary cancer: calcitonin / CEA; RET genetic testing if diagnosed.", { x: X0 + 6.3, y: 5.68, w: 5.6, h: 0.48, fontSize: 12.5, color: BODY, margin: 0, valign: "top", isTextBox: true });
  banner(s, "IMPORTANT", "Suspicious neck nodes may need aspiration regardless of nodule size.", 6.3, 0.55);
  footnote(s, "CT = computed tomography; MRI = magnetic resonance imaging; CEA = carcinoembryonic antigen.");
}

// ---- Slide: Ultrasound Risk and Bethesda Cytology ----
{
  const s = content(5, "Ultrasound Risk and Bethesda Cytology", null, "Teammate content (Objective 6). Sources: [2] ETA 2023; [16] Bethesda 2023; [20] ACR TI-RADS; [21] Moorthy 2026.");
  subtitle(s, "AMERICAN COLLEGE OF RADIOLOGY TI-RADS  /  2023 BETHESDA SYSTEM");
  const lx = X0, lw = 5.85, rx = X0 + 6.05, rw = 6.08, y0 = 1.78, hh = 3.55;
  const t = (x, y, colW, rows, rh, firstColor) => {
    rows.forEach((r, i) => {
      const head = i === 0, w = colW.reduce((a, b) => a + b, 0);
      box(s, x, y + (head ? 0 : 0.45 + (i - 1) * rh), w, head ? 0.45 : rh, head ? NAVY : (i % 2 ? ROWALT : WHITE), head ? NAVY : (i % 2 ? ROWALT : WHITE));
      let cx = x;
      r.forEach((v, j) => {
        s.addText(v, { x: cx + 0.12, y: y + (head ? 0 : 0.45 + (i - 1) * rh), w: colW[j] - 0.15, h: head ? 0.45 : rh, fontSize: head ? 12 : 13.5, bold: head || j === 0 || (firstColor && j === 2), color: head ? WHITE : (firstColor && j === 2 ? TEAL : (j === 0 ? NAVY : BODY)), valign: "middle", margin: 0, isTextBox: true });
        cx += colW[j];
      });
    });
  };
  box(s, lx, y0, lw, hh); topbar(s, lx, y0, lw, TEAL);
  s.addText("ACR TI-RADS: WHEN TO SAMPLE?", { x: lx + 0.22, y: y0 + 0.2, w: lw - 0.4, h: 0.36, fontSize: 15.5, bold: true, color: NAVY, margin: 0, valign: "middle", isTextBox: true });
  t(lx + 0.22, y0 + 0.65, [1.75, 1.8, 1.86], [["CATEGORY", "ASPIRATE (FNA)", "US FOLLOW-UP"], ["TR1-TR2", "None", "None"], ["TR3", "at least 2.5 cm", "at least 1.5 cm"], ["TR4", "at least 1.5 cm", "at least 1.0 cm"], ["TR5", "at least 1.0 cm", "at least 0.5 cm"]], 0.48, false);
  s.addText("High-risk features: punctate foci, irregular margin, taller-than-wide shape or extrathyroidal extension.", { x: lx + 0.22, y: y0 + 3.02, w: lw - 0.44, h: 0.48, fontSize: 12, color: BODY, margin: 0, valign: "top", isTextBox: true });
  box(s, rx, y0, rw, hh); topbar(s, rx, y0, rw, ORANGE);
  s.addText("BETHESDA: CYTOLOGY -> DECISION", { x: rx + 0.22, y: y0 + 0.2, w: rw - 0.4, h: 0.36, fontSize: 15.5, bold: true, color: NAVY, margin: 0, valign: "middle", isTextBox: true });
  t(rx + 0.22, y0 + 0.65, [0.72, 2.0, 0.72, 2.2], [["CLASS", "CYTOLOGY", "ROM", "USUAL NEXT STEP"], ["I", "Nondiagnostic", "13%", "Repeat FNA"], ["II", "Benign", "4%", "Ultrasound follow-up"], ["III", "Atypia (AUS)", "22%", "Repeat FNA / testing"], ["IV", "Follicular neoplasm", "30%", "Testing / lobectomy"], ["V", "Suspicious", "74%", "Surgical assessment"], ["VI", "Malignant", "97%", "Individualized treatment"]], 0.33, true);
  s.addText("ROM = mean risk of malignancy in surgically selected adults; AUS = atypia of undetermined significance.", { x: rx + 0.22, y: y0 + 3.13, w: rw - 0.44, h: 0.38, fontSize: 10, color: MUTED, margin: 0, valign: "top", isTextBox: true });
  box(s, X0, 5.45, CW, 0.5, TEAL_T, TEAL_T);
  s.addText("FNA = fine-needle aspiration; ultrasound and cytology must agree. Follicular invasion requires histology.", { x: X0 + 0.22, y: 5.45, w: CW - 0.4, h: 0.5, fontSize: 13.5, color: BODY, margin: 0, valign: "middle", isTextBox: true });
  banner(s, "IMPORTANT", "Benign cytology + highly suspicious ultrasound: consider repeat FNA.", 6.05, 0.58);
  footnote(s, "TI-RADS = Thyroid Imaging Reporting and Data System.");
}

divider(6);
// =====================================================================
// 13. Management — benign nodules (verbatim)
// =====================================================================
{
  const s = content(6, "Management of Benign Thyroid Nodules", null, "Teammate content (Objective 7). Sources: [2] ETA 2023.");
  const cards = [
    ["1. Observation", TEAL, ["Asymptomatic benign nodules → Ultrasound surveillance.", "Reassess if suspicious features or significant growth develop."]],
    ["2. Surgical Management", RED, ["Indications: Compressive symptoms, cosmetic concerns, or selected suspicious nodules.", [{ text: "Lobectomy: ", b: true }, { text: "Preferred for symptomatic unilateral nodules." }]]],
    ["3. Non-Surgical Options", ORANGE, ["Ethanol or thermal ablation for selected symptomatic benign nodules.", "Radioactive iodine (RAI) for hyperfunctioning nodules."]],
  ];
  cards.forEach(([k, c, items], i) => {
    const x = X0 + i * 4.1, w = 3.93;
    box(s, x, 1.65, w, 5.15); topbar(s, x, 1.65, w, c);
    s.addText(k, { x: x + 0.28, y: 1.95, w: w - 0.5, h: 0.45, fontSize: 19, bold: true, color: NAVY, valign: "middle", margin: 0, isTextBox: true });
    s.addText(bullets(items, 18), { x: x + 0.28, y: 2.6, w: w - 0.5, h: 4.0, valign: "top", margin: 0, isTextBox: true });
  });
}

// =====================================================================
// 14. Management — surgical DTC (verbatim)
// =====================================================================
{
  const s = content(6, "Surgical Management of Differentiated Thyroid Cancer", null, "Teammate content (Objective 7). Sources: [1] ATA 2025; [8] NCCN 2026.");
  label(s, "Papillary and Follicular Thyroid Carcinomas", X0, 1.65, 7, RED, 10.5);
  tbl(s, X0, 2.0, [3.9, 3.0], [
    ["Tumor Characteristics", "Recommended Surgery"],
    ["≤2 cm, low-risk, unilateral", "Lobectomy preferred"],
    [">2–4 cm, low-risk, unilateral", "Lobectomy or total thyroidectomy"],
    [">4 cm, gross invasion, or clinically apparent metastases", "Total thyroidectomy"],
  ], { rowH: 1.4, fontSize: 15 });
  const nx = 7.75, nw = 4.98;
  box(s, nx, 1.65, nw, 5.15); topbar(s, nx, 1.65, nw, TEAL);
  s.addText("Neck Dissection", { x: nx + 0.28, y: 1.95, w: nw - 0.5, h: 0.45, fontSize: 19, bold: true, color: NAVY, valign: "middle", margin: 0, isTextBox: true });
  s.addText(bullets([
    [{ text: "Central: ", b: true }, { text: "For involved central nodes (cN1a); also ipsilaterally with lateral nodal disease." }],
    [{ text: "Lateral: ", b: true }, { text: "For involved lateral nodes (cN1b)." }],
    [{ text: "Prophylactic: ", b: true }, { text: "Not routine in small, low-risk, node-negative cancers." }],
  ], 17), { x: nx + 0.28, y: 2.6, w: nw - 0.5, h: 4.0, valign: "top", margin: 0, isTextBox: true });
}

// =====================================================================
// 15. Management — non-surgical DTC (verbatim)
// =====================================================================
{
  const s = content(6, "Non-Surgical Management of Differentiated Thyroid Cancer", null, "Teammate content (Objective 7). Sources: [1] ATA 2025; [8] NCCN 2026.");
  const cards = [
    ["1. Radioactive Iodine (RAI)", TEAL, [[{ text: "Low risk: ", b: true }, { text: "Not routinely recommended." }], [{ text: "Intermediate risk: ", b: true }, { text: "Consider individually." }], [{ text: "High risk: ", b: true }, { text: "Generally recommended after total thyroidectomy." }]]],
    ["2. Levothyroxine", TEAL, ["Hormone replacement after total thyroidectomy.", "TSH suppression individualized by risk and response."]],
    ["3. External Beam Radiotherapy (EBRT)", ORANGE, ["Selected unresectable, residual, or recurrent disease."]],
    ["4. Systemic Therapy", RED, ["Progressive RAI-refractory cancer.", "Lenvatinib or mutation-specific targeted therapy."]],
  ];
  cards.forEach(([k, c, items], i) => {
    const x = X0 + i * 3.073, w = 2.92;
    box(s, x, 1.65, w, 5.15); topbar(s, x, 1.65, w, c);
    s.addText(k, { x: x + 0.25, y: 1.92, w: w - 0.45, h: 0.75, fontSize: 17, bold: true, color: NAVY, valign: "top", margin: 0, isTextBox: true });
    s.addText(bullets(items, 16), { x: x + 0.25, y: 2.8, w: w - 0.45, h: 3.85, valign: "top", margin: 0, isTextBox: true });
  });
}

// =====================================================================
// 16. Management — other malignancies (verbatim)
// =====================================================================
{
  const s = content(6, "Management of Other Thyroid Malignancies", null, "Teammate content (Objective 7). Sources: [8] NCCN 2026; [9] Saudi CHI 2023.");
  const cards = [
    ["A. Medullary Thyroid Carcinoma (MTC)", ORANGE, [[{ text: "Surgery: ", b: true }, { text: "Total thyroidectomy + central neck dissection; lateral dissection if involved." }], [{ text: "Preoperative: ", b: true }, { text: "RET testing and pheochromocytoma screening." }], [{ text: "Advanced disease: ", b: true }, { text: "Targeted therapy (e.g., selpercatinib for RET-mutated MTC)." }], "RAI is ineffective."]],
    ["B. Anaplastic Thyroid Carcinoma (ATC)", RED, [[{ text: "Resectable: ", b: true }, { text: "Surgery followed by radiotherapy ± chemotherapy." }], [{ text: "Unresectable: ", b: true }, { text: "Radiotherapy and systemic therapy." }], [{ text: "BRAF V600E-positive: ", b: true }, { text: "Dabrafenib + trametinib." }], "Urgent airway assessment and supportive care."]],
    ["C. Other Rare Malignancies", TEAL, [[{ text: "Thyroid lymphoma: ", b: true }, { text: "Chemotherapy ± radiotherapy." }], [{ text: "Thyroid metastases: ", b: true }, { text: "Treat primary cancer; selective surgery." }]]],
  ];
  cards.forEach(([k, c, items], i) => {
    const x = X0 + i * 4.1, w = 3.93;
    box(s, x, 1.65, w, 5.15); topbar(s, x, 1.65, w, c);
    s.addText(k, { x: x + 0.28, y: 1.92, w: w - 0.5, h: 0.75, fontSize: 17, bold: true, color: NAVY, valign: "top", margin: 0, isTextBox: true });
    s.addText(bullets(items, 15.5), { x: x + 0.28, y: 2.78, w: w - 0.5, h: 3.9, valign: "top", margin: 0, isTextBox: true });
  });
}

divider(7);
// =====================================================================
// 21. MY PART — Prognosis and outcomes (single slide)
// =====================================================================
{
  const s = content(7, "Prognosis: follow-up, risk factors and survival", "[10] ATA 2015; [1] ATA 2025; [11] AJCC 8th; [12] SEER; [13] Roman 2006; [14] ATA ATC 2021",
`SCRIPT (about 2 minutes; follows the slide left → middle → right → banner)
"Now we move to prognosis and outcomes. This slide has three parts: follow-up, prognostic factors, and survival.

[left column] On the left is follow-up in differentiated thyroid cancer, meaning papillary and follicular. After surgery we follow three things: thyroglobulin, anti-thyroglobulin antibodies, and a neck ultrasound. Thyroglobulin is made only by thyroid tissue, so after the gland is removed it should be very low. If it rises, some thyroid tissue, usually cancer, is still there. We check the antibodies with it because they can make the result falsely low.
At six to twelve months, the patient is placed in one of four response groups. Excellent: no evidence of disease, and the recurrence risk is only about one to four percent. Indeterminate: the results are unclear, and most of these patients stay stable. Biochemical incomplete: thyroglobulin is raised, but imaging is negative. Structural incomplete: disease is visible on imaging, and this is the highest-risk group.

[middle column] In the middle are the prognostic factors. Age is the most important, with fifty-five years as the AJCC cut-off. Distant metastases are the worst factor. For the tumour itself, gross extension outside the gland and aggressive histology worsen the outcome, and so does an incomplete resection.

[right column] On the right is five-year survival from SEER. It includes all thyroid cancers, but most of them are differentiated. Survival is almost one hundred percent for localized and regional disease, and about fifty percent with distant spread.
Below the chart: medullary cancer is followed with calcitonin and CEA instead of thyroglobulin, because it comes from C cells. Its ten-year survival falls from ninety-six percent when localized to forty percent with distant spread. Anaplastic cancer has a median survival of only about five months.

[banner] So the key message is: thyroid cancer rarely causes death, but it often recurs, mostly in the neck nodes within the first ten years. That is why follow-up is long."

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
  label(s, "DTC follow-up", X0 + 0.25, 1.92, 3.5, TEAL, 10);
  s.addText("Tg + anti-Tg antibodies and neck US; response at 6–12 months", { x: X0 + 0.25, y: 2.22, w: 3.5, h: 0.5, fontSize: 12.5, color: BODY, margin: 0, valign: "top", isTextBox: true });
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
  s.addText("All thyroid cancer, mostly DTC", { x: cx + 0.25, y: 2.18, w: 3.5, h: 0.28, fontSize: 11, italic: true, color: MUTED, margin: 0, isTextBox: true });
  s.addChart(pres.charts.BAR, [{ name: "5-year survival", labels: ["Localized", "Regional", "Distant"], values: [99.9, 98.1, 50.2] }], {
    x: cx + 0.1, y: 2.4, w: cw - 0.2, h: 2.3, barDir: "col", chartColors: [TEAL, TEAL, RED], barGapWidthPct: 55,
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 13, dataLabelFontBold: true, dataLabelColor: NAVY, dataLabelFormatCode: '0.0"%"',
    valAxisHidden: true, valAxisMaxVal: 118, valAxisMinVal: 0, valGridLine: { style: "none" }, catGridLine: { style: "none" },
    catAxisLabelColor: NAVY, catAxisLabelFontSize: 12, catAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", showLegend: false, showTitle: false,
  });
  s.addText([{ text: "MTC: ", options: { bold: true, color: ORANGE } }, { text: "calcitonin + CEA; 10-yr 96% local → 40% distant", options: { color: BODY, breakLine: true } },
    { text: "Anaplastic: ", options: { bold: true, color: RED } }, { text: "median survival ≈ 5 months", options: { color: BODY } }], { x: cx + 0.25, y: 4.78, w: cw - 0.45, h: 0.95, fontSize: 12.5, valign: "middle", margin: 0, isTextBox: true, paraSpaceAfter: 4 });
  remember(s, "recurrence is more common than death: most recur in neck nodes within the first decade", 6.12, "Key: ");
}

// =====================================================================
// 21b. MY PART — Surgical outcomes and post-op morbidity
// =====================================================================
{
  const s = content(7, "Surgical Outcomes & Post-op Morbidity", "[17] Bailey & Love; [1] ATA 2025; [18] Adam 2017",
`SCRIPT (about 1.5 minutes; follows the slide left → middle → right → banner)
"The next slide looks at surgical outcomes: the complications of thyroid surgery, and how surgery affects cancer control.

[left column] On the left is nerve injury. The recurrent laryngeal nerve moves the vocal cord, so injury causes hoarseness. It is temporary in three to eight percent of patients and permanent in less than one to two percent. If both nerves are injured, both vocal cords can close, and the airway is at risk. The external branch of the superior laryngeal nerve controls voice pitch, so its injury causes loss of the high-pitched voice and voice fatigue.

[middle column] In the middle are metabolic and airway complications. The parathyroid glands lie behind the thyroid and can be injured or lose their blood supply during surgery. This causes low calcium, the most common complication, in ten to thirty percent of patients. It is usually temporary, and permanent in less than one to two percent. A neck hematoma is rare, about one percent, but it can compress the airway, so it needs immediate evacuation.

[right column] On the right is oncologic control. An R0 resection means the tumour is completely removed, with nothing left behind. It is the strongest factor the surgeon controls to prevent recurrence. After surgery, the ATA risk group guides two decisions: whether the patient needs radioactive iodine, and how much we suppress TSH.

[banner] Finally, complication rates are clearly lower when the surgeon performs more than about twenty-five thyroid operations a year."

BACKUP Q&A (not read aloud)
Q: How do you manage a post-thyroidectomy haematoma with stridor?
A: Open the wound at the bedside (skin and strap-muscle sutures) to decompress, then return to theatre for haemostasis; secure the airway.

Q: How do you detect and treat hypocalcaemia?
A: Symptoms: perioral tingling, cramps; Chvostek and Trousseau signs. Check post-operative PTH or calcium; oral calcium ± calcitriol; IV calcium gluconate if severe or symptomatic.

Q: What happens with bilateral RLN palsy?
A: Both cords in a paramedian position: stridor after extubation; reintubation and often tracheostomy.

Q: Where does the >25 cases/year figure come from?
A: Adam et al., Ann Surg 2017: complication risk fell with surgeon volume up to about 25–26 total thyroidectomies per year; low-volume surgeons had about 1.5 times the odds of complications.

Q: How do you reduce nerve injury?
A: Routine visual identification of the RLN (landmarks: Berry's ligament, Zuckerkandl's tubercle), ligation of superior pole vessels close to the gland for the EBSLN, and intraoperative nerve monitoring as an adjunct; laryngoscopy before and after surgery.`);
  card(s, { x: X0, y: 1.65, w: 3.95, h: 4.2, color: TEAL, kicker: "Nerve injury", title: null, rows: [
    { l: "RLN injury", t: "Transient 3–8%, permanent < 1–2%; bilateral risks airway compromise", b: true },
    { l: "EBSLN injury", t: "High-pitched voice loss / vocal fatigue" } ] });
  card(s, { x: X0 + 4.09, y: 1.65, w: 3.95, h: 4.2, color: ORANGE, kicker: "Metabolic & airway", title: null, rows: [
    { l: "Hypoparathyroidism", t: "Transient hypocalcemia 10–30% (most common); permanent < 1–2%" },
    { l: "Neck hematoma", t: "Rare (~1%); life-threatening airway emergency requiring immediate evacuation", b: true } ] });
  card(s, { x: X0 + 8.18, y: 1.65, w: 3.95, h: 4.2, color: RED, kicker: "Oncologic control", title: null, rows: [
    { l: "R0 resection", t: "Complete surgical clearance: strongest surgeon-controlled predictor against recurrence", b: true },
    { l: "ATA risk stratification", t: "Guides post-op RAI ablation and TSH suppression targets" } ] });
  remember(s, "complication rates drop significantly in the hands of high-volume thyroid surgeons (> 25 cases/year)", 6.12, "Key: ");
}

divider(8);
// =====================================================================
// 23. MY PART — Case discussion (single slide)
// =====================================================================
{
  const s = content(8, "Case discussion: three nodules, three outcomes", "[16] Bethesda 2023; [10] ATA 2015; [15] ATA MTC 2015; [14] ATA ATC 2021",
`SCRIPT (about 2 minutes, interactive; 9 clicks in the order Scenario → Think → Decide)
"Finally, let's apply this to three cases. For each one, I'll show the scenario first, give you a moment to think, and then show the answer.

[click 1: Case 1 Scenario] Case one is a thirty-four-year-old woman with a two-centimetre nodule. The ultrasound is TI-RADS five, which is highly suspicious, and the FNA is Bethesda six, which means malignant. What is the diagnosis, and what is her risk?
[pause, then click 2: Think] This is papillary cancer. Because she is under fifty-five, she is stage one, and her risk is low.
[click 3: Decide] She had a thyroidectomy without routine radioactive iodine. One year later her thyroglobulin is undetectable, which is an excellent response.

[click 4: Case 2 Scenario] Case two is a forty-two-year-old man with a thyroid nodule and a neck node. He has diarrhoea, and his father had thyroid cancer and episodes of severe high blood pressure. What should we suspect?
[pause, then click 5: Think] Medullary cancer, most likely as part of MEN2A.
[click 6: Decide] Before surgery we must exclude a pheochromocytoma, because operating without knowing about it can cause a dangerous rise in blood pressure. Then we do a total thyroidectomy with neck dissection, and we follow him with calcitonin and CEA.

[click 7: Case 3 Scenario] Case three is a seventy-four-year-old woman with a hard neck mass that has grown over three weeks. She is hoarse and short of breath. What are we worried about?
[pause, then click 8: Think] Anaplastic cancer, but we must also consider lymphoma, because lymphoma is treatable. So we take a core biopsy and test for BRAF.
[click 9: Decide] Management starts with securing the airway, then an urgent multidisciplinary discussion, BRAF-targeted therapy if the test is positive, and early palliative care.

These three cases show how different thyroid cancers can be. Thank you."

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
      const on = `step_${i}_${lab}`;
      s.addShape(pres.shapes.RECTANGLE, { x: ix, y, w: iw, h, fill: { color: f }, line: { color: l, width: 0.75 }, objectName: on + "_box" });
      s.addText(lab.toUpperCase(), { x: ix + 0.13, y: y + 0.06, w: 2, h: 0.26, fontSize: 9.5, bold: true, color: lc, charSpacing: 2.5, margin: 0, valign: "middle", isTextBox: true, objectName: on + "_lbl" });
      s.addText(t, { x: ix + 0.13, y: y + 0.32, w: iw - 0.26, h: h - 0.38, fontSize: 13.5, bold: b, color: tc, valign: "top", margin: 0, isTextBox: true, objectName: on + "_txt" });
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
    "Ringel MD, Sosa JA, Baloch Z, et al. 2025 American Thyroid Association management guidelines for adult patients with differentiated thyroid cancer. Thyroid. 2025;35(8):841–985.",
    "Durante C, Hegedüs L, Czarniecka A, et al. 2023 European Thyroid Association clinical practice guidelines for thyroid nodule management. Eur Thyroid J. 2023;12(5).",
    "Standring S. Gray's Anatomy: The Anatomical Basis of Clinical Practice.",
    "Brunicardi FC, et al. Schwartz's Principles of Surgery.",
    "Saudi Health Council. Saudi Cancer Incidence Report 2023.",
    "Melmed S, et al. Williams Textbook of Endocrinology.",
    "Jung CK, Bychkov A, Kakudo K. Update from the 2022 World Health Organization classification of thyroid tumors: a standardized diagnostic approach. Endocrinol Metab. 2022;37(5):703–718.",
    "National Comprehensive Cancer Network. NCCN Clinical Practice Guidelines in Oncology: Thyroid Carcinoma. 2026.",
    "Council of Health Insurance (CHI), Saudi Arabia. Thyroid cancer treatment algorithm. 2023.",
    "Haugen BR, Alexander EK, Bible KC, et al. 2015 ATA management guidelines for thyroid nodules and differentiated thyroid cancer. Thyroid. 2016;26(1):1–133.",
    "Amin MB, Edge SB, Greene FL, et al., eds. AJCC Cancer Staging Manual. 8th ed. Springer; 2017.",
    "National Cancer Institute. SEER Cancer Stat Facts: Thyroid Cancer. seer.cancer.gov/statfacts/html/thyro.html",
    "Roman S, Lin R, Sosa JA. Prognosis of medullary thyroid carcinoma: predictors of survival in 1252 patients. Cancer. 2006;107(9):2134–2142.",
    "Bible KC, Kebebew E, Brierley J, et al. 2021 ATA guidelines for management of patients with anaplastic thyroid cancer. Thyroid. 2021;31(3):337–386.",
    "Wells SA Jr, Asa SL, Dralle H, et al. Revised ATA guidelines for the management of medullary thyroid carcinoma. Thyroid. 2015;25(6):567–610.",
    "Ali SZ, Baloch ZW, Cochand-Priollet B, et al. The 2023 Bethesda System for Reporting Thyroid Cytopathology. Thyroid. 2023;33(9):1039–1044.",
    "Williams NS, O'Connell PR, McCaskie AW, eds. Bailey & Love's Short Practice of Surgery. 28th ed. CRC Press; 2023.",
    "Adam MA, Thomas S, Youngwirth L, et al. Is there a minimum number of thyroidectomies a surgeon should perform to optimize patient outcomes? Ann Surg. 2017;265(2):402–407.",
    "Hay ID, Bergstralh EJ, Goellner JR, et al. Predicting outcome in papillary thyroid carcinoma (MACIS). Surgery. 1993;114(6):1050–1058.",
    "American College of Radiology. ACR TI-RADS ultrasound reporting template. 2022.",
    "Moorthy R, Balasubramanian SP, Farnell K, et al. Evaluation and management of thyroid nodules: joint consensus statement. Clin Endocrinol. 2026;104:682-692.",
  ];
  const half = 11;
  [refs.slice(0, half), refs.slice(half)].forEach((col, ci) => {
    const runs = [];
    col.forEach((r, i) => {
      runs.push({ text: `${ci * half + i + 1}  `, options: { bold: true, color: TEAL } });
      runs.push({ text: r, options: { color: BODY, breakLine: i < col.length - 1, paraSpaceAfter: 6 } });
    });
    s.addText(runs, { x: X0 + ci * 6.18, y: 1.55, w: 5.95, h: 5.4, fontSize: 10, valign: "top", margin: 0, isTextBox: true });
  });
  s.addNotes("References 1–9 and 20–21 come from the teammates' sections; 10–19 support the prognosis and case sections. Where several sources state the same fact, one is cited.");
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
