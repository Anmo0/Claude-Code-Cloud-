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
// 3. Intro — Introduction + Benign vs malignant (teammate content)
// =====================================================================
{
  const s = content(0, "Introduction to Thyroid Tumors", null, "Teammate content (Objective 1). Sources: [1] ATA 2025; [2] ETA 2023.");
  plainCard(s, X0, 1.65, 4.75, 5.15, TEAL, "Definition", [
    "Thyroid tumors are abnormal growths arising from thyroid gland cells.",
    "They can be benign or malignant.",
    "They may present as a solitary nodule, multiple nodules, or an enlarged thyroid gland.",
    [{ text: "Most thyroid nodules are benign; only a minority are malignant.", b: true }],
  ], { size: 16 });
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
  ], { size: 16 });
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
blank(4);
divider(5);
blank(5);
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
    s.addText(bullets(items, 16), { x: x + 0.28, y: 2.6, w: w - 0.5, h: 4.0, valign: "top", margin: 0, isTextBox: true });
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
  ], 16), { x: nx + 0.28, y: 2.6, w: nw - 0.5, h: 4.0, valign: "top", margin: 0, isTextBox: true });
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
    s.addText(bullets(items, 15), { x: x + 0.25, y: 2.8, w: w - 0.45, h: 3.85, valign: "top", margin: 0, isTextBox: true });
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
    s.addText(bullets(items, 14), { x: x + 0.28, y: 2.78, w: w - 0.5, h: 3.9, valign: "top", margin: 0, isTextBox: true });
  });
}

