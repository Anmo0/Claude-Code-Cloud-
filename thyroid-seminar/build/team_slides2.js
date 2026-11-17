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
  t(rx + 0.22, y0 + 0.65, [0.72, 2.0, 0.72, 2.2], [["CLASS", "CYTOLOGY", "ROM", "USUAL NEXT STEP"], ["I", "Nondiagnostic", "13%", "Repeat FNA"], ["II", "Benign", "4%", "Ultrasound follow-up"], ["III", "Atypia (AUS)", "22%", "Repeat FNA / testing"], ["IV", "Follicular neoplasm", "30%", "Testing / lobectomy"], ["V", "Suspicious", "74%", "Surgical assessment"], ["VI", "Malignant", "97%", "Individualized treatment"]], 0.36, true);
  s.addText("ROM = mean risk of malignancy in surgically selected adults; AUS = atypia of undetermined significance.", { x: rx + 0.22, y: y0 + 3.08, w: rw - 0.44, h: 0.42, fontSize: 10.5, color: MUTED, margin: 0, valign: "top", isTextBox: true });
  box(s, X0, 5.45, CW, 0.5, TEAL_T, TEAL_T);
  s.addText("FNA = fine-needle aspiration; ultrasound and cytology must agree. Follicular invasion requires histology.", { x: X0 + 0.22, y: 5.45, w: CW - 0.4, h: 0.5, fontSize: 13.5, color: BODY, margin: 0, valign: "middle", isTextBox: true });
  banner(s, "IMPORTANT", "Benign cytology + highly suspicious ultrasound: consider repeat FNA.", 6.05, 0.58);
  footnote(s, "TI-RADS = Thyroid Imaging Reporting and Data System.");
}

