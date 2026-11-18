// Presentation 4: status update (user 2026-11-17: "a new presentation of status update (presentation 4, 3-8 slides)").
//   04_CPAT-AI_Status_latest.pptx: Excel and Python fully updated and matching; the initial development phase complete;
//   the key features of the new spreadsheet; the new phase from Monday (distributed work, as business as usual in the
//   old CPAT); the open question of introducing CPAT-AI to the IMF; the functions to add later.
//   Same theme as decks 01-03 (build_decks.js).
//
//   NODE_PATH=<node_modules> PPTX_SKILL=<pptx skill dir> node docs/presentations/build/build_deck04.js
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const RDS = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");

const OUT = path.join(__dirname, "..", "04_CPAT-AI_Status_latest.pptx");
const THEME = {
  name: "CPAT-AI",
  headFontFace: "Arial",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1E2622", lt1: "FFFFFF", dk2: "0D3B2B", lt2: "EEF6EF",
    accent1: "00AE49", accent2: "92D050", accent3: "FFF59D", accent4: "2E75B6",
    accent5: "ED7D31", accent6: "6B7570", hlink: "00843A", folHlink: "5E7A6A",
  },
};
const H = THEME.colors;
const W = 13.333, M = 0.6;

// The figures of the overnight release (filled from the test reports; see the speaker notes)
const FIG = {
  version: "2.0.3",
  date: "18 November 2026",
  codes: "18,700+",          // 18,718 output codes compared (Mitigation reference 2.0.3), Excel = Python, every reference scenario
  policyChecks: "81",       // tests by policy (one policy at a time against the baseline)
  countries: "204",         // countries of the Pigouvian dataset
  emm: "27,783",            // codes per country identical on the EMM edition and Main
};

async function icon(Comp, color) {
  const svg = RDS.renderToStaticMarkup(React.createElement(Comp, { color: "#" + color, size: 256 }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

async function main() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  pres.title = "CPAT-AI status update";
  pres.author = "WB CPAT Team";
  pres.company = "World Bank";
  const C = pres.SchemeColor;

  pres.defineSlideMaster({
    title: "DARK",
    background: { color: H.dk2 },
    objects: [
      { placeholder: { options: { name: "kicker", type: "body", x: M, y: 0.7, w: 9, h: 0.4, fontSize: 14, bold: true,
        color: C.accent2, margin: 0 }, text: "" } },
      { placeholder: { options: { name: "title", type: "title", x: M, y: 1.2, w: 8.4, h: 2.2, fontSize: 40, bold: true,
        color: C.background1, align: "left", valign: "top", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "body", type: "body", x: M, y: 3.6, w: 7.6, h: 2.8, fontSize: 18,
        color: C.background2, valign: "top", margin: 0 }, text: "" } },
    ],
  });
  pres.defineSlideMaster({
    title: "LIGHT",
    background: { color: H.lt1 },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: M, y: 0.4, w: W - 2 * M, h: 0.9, fontSize: 32,
        bold: true, color: C.text2, align: "left", valign: "middle", margin: 0 }, text: "" } },
      { text: { text: "CPAT-AI status update · " + FIG.date, options: { x: M, y: 7.0, w: 8, h: 0.3, fontSize: 10,
        color: C.accent6, margin: 0 } } },
    ],
    slideNumber: { x: W - M - 0.6, y: 7.0, w: 0.6, h: 0.3, fontSize: 10, color: C.accent6, align: "right" },
  });

  const card = (s, x, y, w, h, fill) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08,
    fill: { color: fill }, line: { color: fill } });

  // 1. Title
  pres.addSection({ title: "Status" });
  let s = pres.addSlide({ masterName: "DARK", sectionTitle: "Status" });
  s.addText("STATUS UPDATE · " + FIG.date.toUpperCase(), { placeholder: "kicker" });
  s.addText("CPAT-AI " + FIG.version + ": the initial development phase is complete", { placeholder: "title" });
  s.addText([
    { text: "Excel and Python fully updated, and they match", options: { bullet: true, breakLine: true } },
    { text: "The Excel can now be used in country work", options: { bullet: true, breakLine: true } },
    { text: "From Monday: a new phase, a distributed work model", options: { bullet: true } },
  ], { placeholder: "body", paraSpaceAfter: 8 });
  s.addImage({ data: await icon(fa.FaFlagCheckered, H.accent2), x: 9.9, y: 2.0, w: 2.6, h: 2.6 });
  s.addNotes("CPAT-AI " + FIG.version + " released overnight: the Excel CPAT (Mitigation, Distribution, Air pollution, " +
    "Transport co-benefits, the unified workbook and the MT), Coded CPAT (Python) and the app, all on main.");

  // 2. Excel and Python match
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "Status" });
  s.addText("Excel and Python have been fully updated, and they match", { placeholder: "title" });
  const stats = [
    [FIG.codes, "output codes equal in Excel and Python, every reference scenario (within 1e-9)"],
    [FIG.policyChecks, "checks in 31 tests by policy (each policy alone against the baseline): all pass, Excel = Python"],
    [FIG.countries, "countries with Pigouvian charges from one Python-built dataset"],
    [FIG.emm, "codes per country identical on the EMM data edition and the full data"],
  ];
  stats.forEach(([n, lab], i) => {
    const x = M + i * 3.05;
    card(s, x, 1.7, 2.85, 3.0, H.lt2);
    s.addText(n, { x: x + 0.2, y: 1.95, w: 2.45, h: 1.1, fontSize: 40, bold: true, color: C.accent1, margin: 0,
      isTextBox: true });
    s.addText(lab, { x: x + 0.2, y: 3.15, w: 2.45, h: 1.8, fontSize: 15, color: C.text1, valign: "top", margin: 0,
      isTextBox: true });
  });
  s.addText("One model, two forms: the same equations, the same codes, tested against each other release by release. " +
    "The tests by policy check that each policy moves prices, emissions and revenue as expected.",
    { x: M, y: 5.0, w: W - 2 * M, h: 1.0, fontSize: 16, color: C.text1, margin: 0, isTextBox: true });
  s.addNotes("Figures from the release checks: Coded CPAT's comparison with the Excel references (Mitigation, " +
    "engineer, instruments variant, Distribution, Transport co-benefits, air pollution, unified), the tests by policy " +
    "(cpat_testing/policy_scenarios.py; Excel-AI checks/check_policies.py), the EMM edition check.");

  // 3. Ready for country work
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "Status" });
  s.addText("The initial development phase is complete: the Excel can be used in country work", { placeholder: "title" });
  const ready = [
    [fa.FaFileExcel, "Excel CPAT " + FIG.version, "Mitigation, Distribution, Air pollution and Transport co-benefits, " +
      "one unified workbook, the MT for multiple scenarios"],
    [fa.FaPython, "Coded CPAT", "The same model in Python, with a Shiny app on the same policy page"],
    [fa.FaDatabase, "Data editions", "Main (full data), Default (public data), EMM (Egypt, Morocco, Mexico)"],
    [fa.FaCheckDouble, "Tested", "Every release built, checked and compared with Python and with legacy CPAT"],
  ];
  for (let i = 0; i < ready.length; i++) {
    const [Ic, head, txt] = ready[i];
    const y = 1.6 + i * 1.3;
    s.addShape(pres.shapes.OVAL, { x: M, y, w: 0.9, h: 0.9, fill: { color: H.accent1 }, line: { color: H.accent1 } });
    s.addImage({ data: await icon(Ic, "FFFFFF"), x: M + 0.2, y: y + 0.2, w: 0.5, h: 0.5 });
    s.addText(head, { x: M + 1.2, y, w: 3.4, h: 0.9, fontSize: 20, bold: true, color: C.text2, valign: "middle",
      margin: 0, isTextBox: true });
    s.addText(txt, { x: M + 4.7, y, w: 7.4, h: 0.9, fontSize: 16, color: C.text1, valign: "middle", margin: 0,
      isTextBox: true });
  }
  s.addNotes("Where: the GitHub repository's main branch (excel_ai/unified/releases, excel_ai/master/releases, " +
    "coded/). Known limits and assumptions: core/ISSUES.md (major issues and assumptions), " +
    "core/Legacy_differences.md (every difference from legacy CPAT that is not a bug fix).");

  // 4. Key features
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "Status" });
  s.addText("Key features of the new spreadsheet", { placeholder: "title" });
  const feats = [
    [fa.FaColumns, "One policy page", "Every policy on one User-defined page; settings in one list; pre-defined scenarios"],
    [fa.FaBalanceScale, "Carbon pricing", "Coverage by sector and fuel, enhancement and output-based rebates, ETS by % auctioned"],
    [fa.FaGasPump, "Subsidies and taxes", "Reform by price line, price controls, excises, VAT, a price build-up by fuel"],
    [fa.FaCalendarAlt, "Start and target years", "One timing rule; Linear or Constant after the target year"],
    [fa.FaBolt, "Power", "Capacity targets and phase-outs by technology in the engineer model"],
    [fa.FaLungs, "Co-benefits", "Air pollution and road externalities for every country; the Pigouvian charges"],
  ];
  for (let i = 0; i < feats.length; i++) {
    const [Ic, head, txt] = feats[i];
    const x = M + (i % 3) * 4.1, y = 1.55 + Math.floor(i / 3) * 2.65;
    card(s, x, y, 3.85, 2.4, H.lt2);
    s.addShape(pres.shapes.OVAL, { x: x + 0.25, y: y + 0.25, w: 0.7, h: 0.7, fill: { color: H.accent1 },
      line: { color: H.accent1 } });
    s.addImage({ data: await icon(Ic, "FFFFFF"), x: x + 0.4, y: y + 0.4, w: 0.4, h: 0.4 });
    s.addText(head, { x: x + 1.1, y: y + 0.25, w: 2.6, h: 0.7, fontSize: 18, bold: true, color: C.text2,
      valign: "middle", margin: 0, isTextBox: true });
    s.addText(txt, { x: x + 0.25, y: y + 1.1, w: 3.35, h: 1.15, fontSize: 14, color: C.text1, valign: "top",
      margin: 0, isTextBox: true });
  }
  s.addNotes("Also new: non-energy taxes and subsidies by GHG category; green spending in the fiscal accounts and " +
    "GDP; settings scenarios and settings runs; the current year beside the year of interest on the Dashboard; the " +
    "VSL by income group (a major assumption).");

  // 5. Next phase
  pres.addSection({ title: "Next" });
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "Next" });
  s.addText("From Monday: a new phase, a distributed work model", { placeholder: "title" });
  card(s, M, 1.6, 6.0, 4.9, H.lt2);
  s.addText("How we work", { x: M + 0.3, y: 1.8, w: 5.4, h: 0.5, fontSize: 20, bold: true, color: C.text2, margin: 0,
    isTextBox: true });
  s.addText([
    { text: "Development moves to a distributed model, close to business as usual in the old CPAT", options: { bullet: true, breakLine: true } },
    { text: "Country teams use the Excel in their work and report issues", options: { bullet: true, breakLine: true } },
    { text: "Changes go through the repository: one version, built, tested and released", options: { bullet: true } },
  ], { x: M + 0.3, y: 2.4, w: 5.4, h: 3.8, fontSize: 16, color: C.text1, valign: "top", margin: 0, paraSpaceAfter: 10,
    isTextBox: true });
  card(s, M + 6.4, 1.6, W - 2 * M - 6.4, 4.9, H.dk2);
  s.addImage({ data: await icon(fa.FaQuestionCircle, H.accent2), x: M + 6.75, y: 1.9, w: 0.8, h: 0.8 });
  s.addText("Outstanding issue to discuss", { x: M + 6.75, y: 2.85, w: 5.0, h: 0.5, fontSize: 20, bold: true,
    color: C.background1, margin: 0, isTextBox: true });
  s.addText("How to introduce CPAT-AI to the IMF: who presents it, to whom, and how the two institutions work on it " +
    "together.", { x: M + 6.75, y: 3.45, w: 5.0, h: 2.6, fontSize: 18, color: C.background2, valign: "top", margin: 0,
    isTextBox: true });
  s.addNotes("For discussion: the introduction to the IMF (timing, audience, demonstration on a country), and the " +
    "possibility of IMF modules (next slide).");

  // 6. Remaining functions
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "Next" });
  s.addText("Functions still to add, in the next phase", { placeholder: "title" });
  const later = [
    [fa.FaIndustry, "ETS and CBAM module", "Intensity-based targets and benchmark-based systems; downstream pricing and CBAM"],
    [fa.FaSeedling, "Public (green) investment and subsidies", "With CDPM / SiSePuede: their data incorporated into CPAT"],
    [fa.FaPercent, "Cost of capital (derisking)", "Policies that lower the cost of capital of clean investment"],
    [fa.FaTools, "Tailoring", "For Policy Crediting, GHG accounting and other uses"],
    [fa.FaHandshake, "IMF modules", "We could work with the IMF to add their modules for different things"],
  ];
  for (let i = 0; i < later.length; i++) {
    const [Ic, head, txt] = later[i];
    const y = 1.5 + i * 1.05;
    s.addShape(pres.shapes.OVAL, { x: M, y, w: 0.8, h: 0.8, fill: { color: i === 4 ? H.accent4 : H.accent1 },
      line: { color: i === 4 ? H.accent4 : H.accent1 } });
    s.addImage({ data: await icon(Ic, "FFFFFF"), x: M + 0.18, y: y + 0.18, w: 0.44, h: 0.44 });
    s.addText(head, { x: M + 1.1, y, w: 4.6, h: 0.8, fontSize: 18, bold: true, color: C.text2, valign: "middle",
      margin: 0, isTextBox: true });
    s.addText(txt, { x: M + 5.8, y, w: 6.3, h: 0.8, fontSize: 15, color: C.text1, valign: "middle", margin: 0,
      isTextBox: true });
  }
  s.addNotes("On the policy page these are visible and greyed (not yet in the model): C2-C3 benchmark-based " +
    "instruments, D4 guarantees, green spending's physical effects. The FTT technology models remain a separate " +
    "track.");

  // 7. Close
  s = pres.addSlide({ masterName: "DARK", sectionTitle: "Next" });
  s.addText("FOR DISCUSSION", { placeholder: "kicker" });
  s.addText("Introducing CPAT-AI to the IMF", { placeholder: "title" });
  s.addText([
    { text: "Who presents, to whom, and when", options: { bullet: true, breakLine: true } },
    { text: "A country demonstration with the new Excel", options: { bullet: true, breakLine: true } },
    { text: "Joint work: IMF modules in CPAT", options: { bullet: true } },
  ], { placeholder: "body", paraSpaceAfter: 8 });
  s.addImage({ data: await icon(fa.FaComments, H.accent2), x: 9.9, y: 2.0, w: 2.6, h: 2.6 });

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote " + OUT);
}

main().catch((e) => { console.error(e); process.exit(1); });
