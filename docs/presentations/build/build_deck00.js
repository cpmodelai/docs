// Presentation 0: how to use CPAT and the MT for country work (user 2026-11-18: "a quick userguide for how to use CPAT
//   for country work ... where to get it, activation, how to input scenarios, how settings work, how MT works, and its
//   macros ... slice CPAT for a single country ... the different folders (master to activate, active folder afterwards
//   (gitignored), published version) ... for the non-technical audience ... an overview of the other PPTs").
//   00_CPAT-AI_CountryWork_latest.pptx. Same theme as decks 01-04 (build_decks.js).
//
//   NODE_PATH=<node_modules with pptxgenjs, react, react-dom, react-icons, sharp> \
//   PPTX_SKILL=<pptx skill dir> node docs/presentations/build/build_deck00.js
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const RDS = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");

const OUT = path.join(__dirname, "..", "00_CPAT-AI_CountryWork_latest.pptx");
const IMG = (f) => path.join(__dirname, "img", f);
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
const FOOT = "How to use CPAT and the MT for country work · CPAT-AI 2.0";

async function icon(Comp, color) {
  const svg = RDS.renderToStaticMarkup(React.createElement(Comp, { color: "#" + color, size: 256 }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

async function main() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  pres.title = "How to use CPAT and the MT for country work";
  pres.author = "WB CPAT Team";
  pres.company = "World Bank";
  const C = pres.SchemeColor;

  pres.defineSlideMaster({
    title: "DARK",
    background: { color: H.dk2 },
    objects: [
      { placeholder: { options: { name: "kicker", type: "body", x: M, y: 0.7, w: 9, h: 0.4, fontSize: 14, bold: true,
        color: C.accent2, charSpacing: 2, margin: 0 }, text: "" } },
      { placeholder: { options: { name: "title", type: "title", x: M, y: 1.2, w: 8.4, h: 2.2, fontSize: 40, bold: true,
        color: C.background1, align: "left", valign: "top", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "body", type: "body", x: M, y: 3.6, w: 7.8, h: 2.9, fontSize: 18,
        color: C.background2, valign: "top", margin: 0 }, text: "" } },
      { text: { text: FOOT, options: { x: M, y: 6.95, w: 9, h: 0.3, fontSize: 10, color: C.accent2, margin: 0 } } },
    ],
  });
  pres.defineSlideMaster({
    title: "CONTENT",
    background: { color: H.lt1 },
    objects: [
      { placeholder: { options: { name: "kicker", type: "body", x: M, y: 0.35, w: 10, h: 0.3, fontSize: 12, bold: true,
        color: C.accent1, charSpacing: 2, margin: 0 }, text: "" } },
      { placeholder: { options: { name: "title", type: "title", x: M, y: 0.65, w: W - 2 * M, h: 0.75, fontSize: 28,
        bold: true, color: C.text2, align: "left", valign: "top", margin: 0 }, text: "" } },
      { text: { text: FOOT, options: { x: M, y: 7.05, w: 10, h: 0.3, fontSize: 10, color: C.accent6, margin: 0 } } },
    ],
    slideNumber: { x: W - M - 0.5, y: 7.05, w: 0.5, h: 0.3, fontSize: 10, color: H.accent6, align: "right" },
  });

  // ---------- helpers ----------
  const card = (s, x, y, w, h, fill) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08,
    fill: { color: fill || H.lt2 }, line: { color: fill || H.lt2 } });
  const slide = (kicker, title, section) => {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
    s.addText(kicker, { placeholder: "kicker" });
    s.addText(title, { placeholder: "title" });
    return s;
  };
  const txt = (s, t, x, y, w, h, o = {}) => s.addText(t, { x, y, w, h, fontSize: o.fs || 15, color: o.color || H.dk1,
    bold: o.bold, italic: o.italic, valign: o.valign || "top", align: o.align, margin: 0, isTextBox: true,
    paraSpaceAfter: o.space || 4, fontFace: o.mono ? "Courier New" : undefined });
  const bl = (items, o = {}) => items.map((t, i) => {
    const runs = Array.isArray(t) ? t : [{ text: t }];
    return runs.map((r, j) => ({ text: r.text, options: { ...(r.options || {}),
      bullet: j === 0 ? { indent: 14 } : undefined, breakLine: j === runs.length - 1 && i < items.length - 1 } }));
  }).flat();
  const B = (t) => ({ text: t, options: { bold: true } });
  const T = (t) => ({ text: t });
  const CODE = (t) => ({ text: t, options: { fontFace: "Courier New", fontSize: 13, color: H.dk2 } });
  const step = (s, x, y, w, n, head, body, o = {}) => {
    const d = 0.5;
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: o.fill || H.accent1 }, line: { color: o.fill || H.accent1 } });
    s.addText(String(n), { x, y, w: d, h: d, align: "center", valign: "middle", fontSize: 16, bold: true,
      color: "FFFFFF", margin: 0, isTextBox: true });
    s.addText([{ text: head, options: { bold: true, fontSize: o.hs || 16, color: H.dk2, breakLine: true } },
      ...(Array.isArray(body) ? body : [{ text: body }]).map((r) => ({ text: r.text,
        options: { fontSize: o.fs || 13.5, color: H.dk1, ...(r.options || {}) } }))],
    { x: x + d + 0.2, y: y - 0.04, w: w - d - 0.2, h: o.h || 0.9, valign: "top", margin: 0, isTextBox: true });
  };
  const iconHead = async (s, Ic, x, y, head, o = {}) => {
    const d = o.d || 0.6;
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: o.fill || H.accent1 }, line: { color: o.fill || H.accent1 } });
    s.addImage({ data: await icon(Ic, "FFFFFF"), x: x + d * 0.22, y: y + d * 0.22, w: d * 0.56, h: d * 0.56 });
    s.addText(head, { x: x + d + 0.15, y, w: o.w || 3, h: d, fontSize: o.hs || 17, bold: true, color: H.dk2,
      valign: "middle", margin: 0, isTextBox: true });
  };
  const table = (s, rows, x, y, w, colW, o = {}) => {
    const fs = o.fs || 12;
    const data = rows.map((r, i) => r.map((c) => ({ text: c, options: {
      bold: i === 0 || (o.boldFirst && c === r[0]), color: i === 0 ? "FFFFFF" : H.dk1,
      fill: { color: i === 0 ? H.dk2 : (i % 2 ? "FFFFFF" : H.lt2) }, fontSize: fs, valign: "middle" } })));
    s.addTable(data, { x, y, w, colW, border: { type: "solid", pt: 0.5, color: "C9D3CC" }, margin: 0.06,
      rowH: o.rowH || 0.32, autoPage: false });
  };
  const note = (s, t, y = 6.35) => {
    card(s, M, y, W - 2 * M, 0.55, "FFF9C4");
    txt(s, t, M + 0.2, y + 0.08, W - 2 * M - 0.4, 0.42, { fs: 13, valign: "middle" });
  };

  // ===================== 1. Title
  pres.addSection({ title: "Start" });
  let s = pres.addSlide({ masterName: "DARK", sectionTitle: "Start" });
  s.addText("00 · START HERE FOR COUNTRY WORK", { placeholder: "kicker" });
  s.addText("How to use CPAT and the MT for country work", { placeholder: "title" });
  s.addText([
    { text: "Where to get CPAT, and which folder is which", options: { bullet: true, breakLine: true } },
    { text: "Activate it, and slice it down to your country", options: { bullet: true, breakLine: true } },
    { text: "Enter scenarios and settings, read the results", options: { bullet: true, breakLine: true } },
    { text: "Run many scenarios with the MT and its macros", options: { bullet: true } },
  ], { placeholder: "body", paraSpaceAfter: 8 });
  s.addImage({ data: await icon(fa.FaGlobeAfrica, H.accent2), x: 9.9, y: 2.0, w: 2.6, h: 2.6 });
  s.addNotes("For country economists and analysts. No knowledge of git, Python or VBA is needed. The other decks " +
    "(01-04) cover development, the interface in depth, Python and the project status.");

  // ===================== 2. Two files
  s = slide("THE BIG PICTURE", "Two Excel files do the work: CPAT and the MT", "Start");
  card(s, M, 1.6, 5.9, 4.4);
  await iconHead(s, fa.FaFileExcel, M + 0.3, 1.8, "CPAT (the model)", { w: 4.5, hs: 20 });
  txt(s, bl([
    "One workbook: Mitigation with Distribution, Air pollution and Transport co-benefits inside",
    "One country and up to four scenarios at a time; results on its Dashboard",
    "Where you enter policies (User-defined, Pre-defined) and settings",
    "No macros: it is ordinary Excel",
  ]), M + 0.3, 2.6, 5.3, 3.3, { fs: 15, space: 8 });
  card(s, M + 6.3, 1.6, W - 2 * M - 6.3, 4.4);
  await iconHead(s, fa.FaCogs, M + 6.6, 1.8, "The MT (Multiscenario Tool)", { w: 4.8, hs: 20 });
  txt(s, bl([
    "The control panel, and the only file with macros",
    "Activates CPAT (prepares it for use) and loads the data",
    "Runs many countries x scenarios and stores every result as values",
    "Compares stored runs: tables and graphs for one or many countries",
  ]), M + 6.6, 2.6, 5.3, 3.3, { fs: 15, space: 8 });
  note(s, "Rule of thumb: you edit policies in CPAT; you press buttons in the MT. For a single country and a few scenarios, CPAT alone is enough.", 6.25);

  // ===================== 3. Where to get it
  pres.addSection({ title: "Get it" });
  s = slide("WHERE TO GET IT", "Get CPAT from the GitHub repository", "Get it");
  step(s, M, 1.65, 6.0, 1, "Ask for access", "A GitHub account with access to github.com/cpmodelai/cpat_ai (ask the CPAT team).");
  step(s, M, 2.7, 6.0, 2, "Get a copy on your laptop (once)", [T("Easiest: GitHub Desktop, File > Clone repository, into "), CODE("C:\\Users\\<you>\\Repos"), T(". Setup guide: repo-tools/SetUpGuide.docx.")], { h: 1.1 });
  step(s, M, 3.95, 6.0, 3, "Keep it up to date", [T("Double-click "), CODE("SyncRepo.bat"), T(" at the top of the repository (or Fetch / Pull in GitHub Desktop) before each session.")], { h: 1.1 });
  step(s, M, 5.2, 6.0, 4, "Just one file?", "On github.com, open the file and press the download button (Download raw file). Fine for a quick look; you will not get updates.", { h: 1.1 });
  card(s, M + 6.4, 1.6, W - 2 * M - 6.4, 4.75);
  txt(s, "Which file?", M + 6.7, 1.8, 5.2, 0.45, { fs: 18, bold: true, color: H.dk2 });
  txt(s, bl([
    [B("Always take the file named _latest. "), T("Older versions sit in each folder's releases/.")],
    [B("Quickest start: "), T("the files in "), CODE("excel_ai/published_cpat/"), T(", already activated by the development team (next slide).")],
    [B("Your own, freshest copy: "), T("open "), CODE("excel_ai/master/CPAT-AI-MT_latest.xlsm"), T(" and activate it.")],
    [B("Documentation: "), T("cpmodelai.github.io/docs (the website), and the decks in docs/presentations/.")],
  ]), M + 6.7, 2.35, 5.2, 3.9, { fs: 14, space: 10 });

  // ===================== 4. The three folders
  s = slide("FIND YOUR WAY · THE FOLDERS", "Three folders matter: master, active_workbooks, published_cpat", "Get it");
  const fx = [M, M + 4.15, M + 8.3], fw = 3.85;
  const folders = [
    [fa.FaBoxOpen, "master/", "The source (shared)", H.accent4,
      [[CODE("CPAT-AI-MT_latest.xlsm")], [T("The MT as released, with every macro. You open it to "), B("activate"), T(".")],
        [T("Tracked by git: the team's copy. Do not save your work here.")]]],
    [fa.FaUserEdit, "active_workbooks/", "Your workspace (private)", H.accent1,
      [[T("Made by Activate: "), CODE("CPAT-AI-MT.xlsb"), T(" and "), CODE("CPAT_latest-4s.xlsb")],
        [T("This is where you work, run and save.")],
        [B("Git-ignored: "), T("never uploaded, never overwritten (next slide).")]]],
    [fa.FaStar, "published_cpat/", "Ready to use (shared)", H.accent5,
      [[T("The last version "), B("activated by the development team"), T(": "), CODE("CPAT-AI-MT.xlsb"), T(", "), CODE("CPAT_latest-3s.xlsb")],
        [T("No activation needed. Copy the files out first, then open the copies.")],
        [T("Earlier ones in Old/.")]]],
  ];
  for (let i = 0; i < 3; i++) {
    const [Ic, name, sub, col, lines] = folders[i];
    card(s, fx[i], 1.6, fw, 4.55);
    s.addShape(pres.shapes.RECTANGLE, { x: fx[i], y: 1.6, w: fw, h: 0.12, fill: { color: col }, line: { color: col } });
    s.addShape(pres.shapes.OVAL, { x: fx[i] + 0.25, y: 1.95, w: 0.65, h: 0.65, fill: { color: col }, line: { color: col } });
    s.addImage({ data: await icon(Ic, "FFFFFF"), x: fx[i] + 0.4, y: 2.1, w: 0.35, h: 0.35 });
    txt(s, [{ text: "excel_ai/" + name, options: { bold: true, fontSize: 17, color: H.dk2, breakLine: true } },
      { text: sub, options: { fontSize: 13, color: col, bold: true } }], fx[i] + 1.05, 1.95, fw - 1.2, 0.8);
    txt(s, bl(lines), fx[i] + 0.25, 2.9, fw - 0.45, 3.2, { fs: 13.5, space: 8 });
    if (i < 2) s.addShape(pres.shapes.RIGHT_ARROW, { x: fx[i] + fw + 0.02, y: 3.55, w: 0.26, h: 0.45,
      fill: { color: H.accent6 }, line: { color: H.accent6 } });
  }
  note(s, "Flow: master (open the MT) → Activate → active_workbooks (work here). Short cut: copy published_cpat's two files into active_workbooks and open the MT copy there.");
  s.addNotes("The activated CPAT is named by its number of live scenario groups: CPAT_latest-2s, -3s or -4s.xlsb. " +
    "The published files are tracked by git: if you open and save them inside the repository, the next sync will " +
    "see a change and may clash with the team's next upload. Copy them out first.");

  // ===================== 5. Git-ignored
  s = slide("FIND YOUR WAY · GIT-IGNORED", "Your workspace is \"git-ignored\": what that means", "Get it");
  txt(s, [{ text: "Git is the system that keeps the team's shared copy of the repository in step: SyncRepo.bat downloads " +
    "everyone's changes and uploads yours. A git-ignored folder is one git deliberately does not look at.",
    options: { fontSize: 16 } }], M, 1.55, W - 2 * M, 0.8);
  const gi = [
    [fa.FaShieldAlt, "Safe", "A team update never overwrites your activated files; Reset and Stash leave them alone."],
    [fa.FaUserSecret, "Private", "Your workbooks (which may hold real, confidential data) are never uploaded to GitHub."],
    [fa.FaCloudUploadAlt, "Not backed up", "If your laptop dies, they are gone. Keep copies of important work on OneDrive or SharePoint."],
    [fa.FaShareAlt, "Not shared", "Colleagues do not see them. To share, send the file (a reduced one: slide 10)."],
    [fa.FaSyncAlt, "Not upgraded", "A new release does not change your copy: activate again. The old one moves to active_workbooks/old."],
  ];
  for (let i = 0; i < gi.length; i++) {
    const [Ic, head, t] = gi[i];
    const y = 2.5 + i * 0.8;
    await iconHead(s, Ic, M, y, head, { d: 0.55, w: 2.2, hs: 16, fill: i < 2 ? H.accent1 : H.accent5 });
    txt(s, t, M + 3.0, y + 0.05, 9.1, 0.55, { fs: 15, valign: "middle" });
  }
  s.addNotes("Technically: .gitignore lists excel_ai/active_workbooks/* (only its README is tracked). ResetToMain.bat " +
    "runs git clean -fd, which never deletes ignored files.");

  // ===================== 6. Before you start
  s = slide("BEFORE YOU START", "What you need on your laptop", "Get it");
  const need = [
    [fa.FaWindows, "Excel 365, 64-bit", "On Windows. The full workbook is large; 64-bit Excel and 16 GB of memory help."],
    [fa.FaUnlock, "Macros allowed for the MT", "Open the MT, press Enable Content. If Excel blocks it: right-click the file, Properties, tick Unblock."],
    [fa.FaDatabase, "The data", "The full data (Main) sit in the repository; the MT's Update data loads them. Anything shared outside the Bank uses the public Default edition."],
    [fa.FaClock, "Some patience", "Activation takes several minutes; a run of many scenarios longer. Estimate run time tells you first."],
  ];
  for (let i = 0; i < need.length; i++) {
    const [Ic, head, t] = need[i];
    const x = M + (i % 2) * 6.15, y = 1.65 + Math.floor(i / 2) * 2.35;
    card(s, x, y, 5.95, 2.1);
    await iconHead(s, Ic, x + 0.25, y + 0.25, head, { w: 4.8 });
    txt(s, t, x + 0.25, y + 1.0, 5.45, 1.0, { fs: 14.5 });
  }
  s.addNotes("Six data tables are proprietary (Balances, EnergyCons, Prices_dom, IO_GTAP, HHSurvey, Tran_Scatter). " +
    "Do not send a workbook with the full data outside the team without checking.");

  // ===================== 7. Activation
  pres.addSection({ title: "Activate and slice" });
  s = slide("ACTIVATION", "Activate: turn the released files into your working CPAT", "Activate and slice");
  step(s, M, 1.65, 6.4, 1, "Sync, then open the MT", [T("Run SyncRepo.bat. Open "), CODE("excel_ai/master/CPAT-AI-MT_latest.xlsm"), T("; Enable Content.")]);
  step(s, M, 2.65, 6.4, 2, "Go to tab Macros, check the settings", "Live scenarios (2-4: how many scenarios CPAT shows side by side; fewer = faster), Last year (2041), Formulas (leave Mix).", { h: 1.1 });
  step(s, M, 3.85, 6.4, 3, "Press Activate Unified", "The MT saves itself as CPAT-AI-MT.xlsb in active_workbooks, opens CPAT, fills the years, adds the scenario groups, recalculates and saves CPAT_latest-<n>s.xlsb beside it.", { h: 1.3 });
  step(s, M, 5.25, 6.4, 4, "From now on, work in active_workbooks", "Close the .xlsm. Open CPAT-AI-MT.xlsb and CPAT_latest-<n>s.xlsb from active_workbooks next time.", { h: 1.0 });
  card(s, M + 6.8, 1.6, W - 2 * M - 6.8, 4.9);
  txt(s, "Why activate?", M + 7.1, 1.8, 4.6, 0.4, { fs: 18, bold: true, color: H.dk2 });
  txt(s, bl([
    "The released files are kept small: formulas for the first years only, one scenario group. Activation builds the full workbook on your laptop.",
    ".xlsb (binary) files open and save much faster than .xlsx.",
    "When: once, and again after each new release (the MT's Cover says the version).",
    "Do not edit the activated file to keep a change in the model: changes to the model go through the team.",
  ]), M + 7.1, 2.3, 4.6, 4.1, { fs: 14, space: 9 });

  // ===================== 8. Slice for one country: why
  s = slide("SLICE CPAT FOR YOUR COUNTRY · 1 OF 3", "Why slice? Smaller, faster, easier to share", "Activate and slice");
  const why = [
    ["~27 MB", "the full activated CPAT holds every country's data tables"],
    ["7 tables", "hold most of that: Balances, EnergyCons, IO_GTAP, HHSurvey, Prices_dom, EF_PM, Tran_Scatter"],
    ["2 dials", "make CPAT lighter: the countries kept, and the number of live scenarios (2, 3 or 4)"],
  ];
  why.forEach(([n, l], i) => {
    const x = M + i * 4.1;
    card(s, x, 1.65, 3.85, 2.2);
    txt(s, n, x + 0.25, 1.85, 3.4, 0.8, { fs: 34, bold: true, color: H.accent1 });
    txt(s, l, x + 0.25, 2.7, 3.4, 1.1, { fs: 14.5 });
  });
  txt(s, "What slicing does", M, 4.15, 6, 0.4, { fs: 18, bold: true, color: H.dk2 });
  txt(s, bl([
    "Deletes the rows of every other country from the seven large data tables",
    "Keeps regional aggregates (e.g. GTAP regions) and rows without a country, so lookups still work",
    "Leaves the model, the scenarios and the settings exactly as they are: your country's numbers do not change",
  ]), M, 4.6, 6.0, 1.7, { fs: 14.5, space: 6 });
  txt(s, "What it does not do", M + 6.4, 4.15, 6, 0.4, { fs: 18, bold: true, color: H.dk2 });
  txt(s, bl([
    "It cannot be undone in that file: to add a country back, activate again",
    "It does not make the file public: your country's real data are still inside",
    "It does not choose the country on CPAT's Dashboard: you do that (next slides)",
  ]), M + 6.4, 4.6, 5.7, 1.7, { fs: 14.5, space: 6 });

  // ===================== 9. Slice: how
  s = slide("SLICE CPAT FOR YOUR COUNTRY · 2 OF 3", "Slice in four steps, all in the MT", "Activate and slice");
  step(s, M, 1.65, 6.2, 1, "List your country on the MT's Dashboard", "Countries and scenarios to run: type your country (e.g. Egypt) as the only country. Delete the others.", { h: 1.0 });
  step(s, M, 2.8, 6.2, 2, "Choose the number of scenarios", "Tab Macros, Live scenarios: 2, 3 or 4 (3 is the team's published choice). Then Activate Unified, if not done yet.", { h: 1.0 });
  step(s, M, 3.95, 6.2, 3, "Press Reduce to countries", [T("Tab Macros, section Countries. The activated CPAT keeps only your country, is saved, and a copy is written: "), CODE("CPAT_<country>_<n>s.xlsb"), T(".")], { h: 1.1 });
  step(s, M, 5.15, 6.2, 4, "Pick the country in CPAT", "CPAT's Dashboard: the country picker at the top. Check the Current situation tiles look right.", { h: 1.0 });
  card(s, M + 6.6, 1.6, W - 2 * M - 6.6, 4.8);
  txt(s, "Make it stick", M + 6.9, 1.8, 5, 0.4, { fs: 18, bold: true, color: H.dk2 });
  txt(s, bl([
    [B("Keep data for: Dashboard countries "), T("(tab Macros, under Reduce to countries): every future activation is reduced to your countries automatically.")],
    [B("Update data, Rows to load: Country "), T("loads only the chosen country's rows when you refresh the data: faster.")],
    [B("A few countries? "), T("List them all on the Dashboard: Reduce keeps each of them (the copy is named after them).")],
    [B("Name it: "), T("the copy CPAT_<country>_<n>s.xlsb is the file to keep and share.")],
  ]), M + 6.9, 2.3, 5.0, 4.0, { fs: 14, space: 10 });

  // ===================== 10. Slice: sharing & recipe
  s = slide("SLICE CPAT FOR YOUR COUNTRY · 3 OF 3", "Sharing a one-country CPAT", "Activate and slice");
  table(s, [
    ["You want to…", "Do this"],
    ["Send CPAT to a colleague in the Bank", "Send the reduced copy CPAT_<country>_<n>s.xlsb (plus the MT only if they will run batches)"],
    ["Send it to a government counterpart", "Check with the team first: real data from proprietary sources are inside. Outside the Bank, use the public Default edition (Update data, Data source: Default)"],
    ["Add a second country later", "Activate again (or Update data with all countries), list both countries, Reduce to countries"],
    ["Move to a new CPAT release", "Sync, activate the new MT, reduce again, then re-enter your scenario (keep a note of your inputs)"],
    ["Keep a safe copy", "Copy the file to OneDrive or SharePoint: active_workbooks is not backed up"],
  ], M, 1.6, W - 2 * M, [4.0, W - 2 * M - 4.0], { fs: 14, rowH: 0.62, boldFirst: true });
  note(s, "Tip: write your scenario as a column of tab Pre-defined (slide 13). A column can be copied into a new release in one paste.", 5.9);

  // ===================== 11. Map of CPAT
  pres.addSection({ title: "Use CPAT" });
  s = slide("FIND YOUR WAY · INSIDE CPAT", "CPAT's tabs, in the order you use them", "Use CPAT");
  const tabs = [
    ["Dashboard", "Start and end here: country, scenarios 2-4, year of interest; the key results and charts", H.accent1],
    ["SCENARIOS ›", "User-defined (your own policy package) · Pre-defined (the scenario table, one column each) · Lists", H.accent4],
    ["OUTPUTS ›", "Comparison (scenarios side by side) · Climate (diagnostic) · Fiscal (macro-fiscal diagnostic) · Output Data (every result, by code)", H.accent4],
    ["ADVANCED ›", "Cover (version) · Settings (assumptions shared by all scenarios) · Manual Inputs · Mappings", H.accent5],
    ["CALCS ›", "Mitigation · Distribution · Transport · AirPollution: the model itself, read like a table. Look, do not touch", H.accent6],
    ["DATA ›", "The data tables (Balances, Prices, GHGs, …), loaded by the MT's Update data", H.accent6],
  ];
  tabs.forEach(([n, t, col], i) => {
    const y = 1.55 + i * 0.75;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: M, y, w: 2.5, h: 0.6, rectRadius: 0.06, fill: { color: col }, line: { color: col } });
    txt(s, n, M, y, 2.5, 0.6, { fs: 15, bold: true, color: "FFFFFF", align: "center", valign: "middle" });
    txt(s, t, M + 2.75, y, W - 2 * M - 2.75, 0.6, { fs: 14.5, valign: "middle" });
  });
  card(s, M, 6.15, W - 2 * M, 0.75, "FFF9C4");
  s.addShape(pres.shapes.RECTANGLE, { x: M + 0.25, y: 6.33, w: 0.45, h: 0.38, fill: { color: "FFF59D" }, line: { color: "BFA900" } });
  txt(s, "Lemon cells are inputs: type there. Grey cells come from the model. Struck-through items are not yet in the model. Tab groups (…›) are just dividers.", M + 0.9, 6.2, W - 2 * M - 1.1, 0.65, { fs: 13.5, valign: "middle" });

  // ===================== 12. User-defined
  s = slide("ENTER A SCENARIO · 1 OF 2", "Your own policy package: tab User-defined", "Use CPAT");
  s.addImage({ path: IMG("user_defined.png"), x: M, y: 1.55, w: 6.2, h: 6.2 * 0.62, sizing: { type: "contain", w: 6.2, h: 6.2 * 0.62 },
    line: { color: H.accent6, width: 0.75 } });
  txt(s, "Tab User-defined: section A, carbon pricing (public-data render)", M, 5.5, 6.2, 0.3, { fs: 11, italic: true, color: H.accent6 });
  txt(s, bl([
    [B("1 Name it. "), T("Selection name (pickers, up to 28 characters) and short name (charts, up to 12).")],
    [B("2 Pick its settings. "), T("Settings scenario: DEF, USR, HPR or LPR (slide 14).")],
    [B("3 Fill the sections. "), T("A carbon tax, exemptions, ETS, feebates · B fossil-fuel subsidy reform · C, D power and road transport · revenue use (tax cuts, transfers, investment).")],
    [B("4 Watch the charts "), T("beside each section: they update as you type.")],
    [B("5 Show it. "), T("It is scenario 2, \"User defined\": pick it on the Dashboard.")],
  ]), M + 6.5, 1.6, W - 2 * M - 6.5, 4.6, { fs: 14, space: 9 });
  note(s, "Each instrument has a start year and a target year; prices are in real US$. Hover over a cell's note for its meaning and unit.", 6.2);

  // ===================== 13. Pre-defined
  s = slide("ENTER A SCENARIO · 2 OF 2", "Several packages: tab Pre-defined, one column per scenario", "Use CPAT");
  table(s, [
    ["Scenario ID", "BASE", "P1", "P2", "P3", "P4 …", "U1 …"],
    ["Name", "Baseline", "User defined", "Tax to $50 + feebates", "Carbon tax $50", "ETS power + industry", "yours"],
    ["Role", "always scenario 1", "linked to User-defined", "preset", "preset", "preset", "add as many as you like"],
  ], M, 1.6, W - 2 * M, [1.7, 1.6, 1.85, 2.0, 1.75, 1.65, W - 2 * M - 10.55], { fs: 13, rowH: 0.45 });
  txt(s, "Add a scenario", M, 3.2, 6, 0.4, { fs: 18, bold: true, color: H.dk2 });
  txt(s, bl([
    "Copy the last column one to the right",
    "Give it an ID, a selection name and a short name",
    "Change the lemon cells (the policy values)",
    "Run? Yes if the MT should run it in a batch",
    "Pick it on the Dashboard as scenario 2, 3 or 4",
  ]), M, 3.65, 6.0, 2.6, { fs: 14.5, space: 6 });
  txt(s, "Good practice for country work", M + 6.4, 3.2, 6, 0.4, { fs: 18, bold: true, color: H.dk2 });
  txt(s, bl([
    "Leave BASE as it is: everything is measured against it",
    "Build your package on User-defined, then copy it into a new Pre-defined column to keep it",
    "Use clear names: they appear on every chart and in the MT",
    "The MT keeps its own copy of this table (its tab Scenarios); Sync scenarios sends it to CPAT",
  ]), M + 6.4, 3.65, 5.7, 2.8, { fs: 14.5, space: 6 });

  // ===================== 14. Settings
  s = slide("SETTINGS", "Settings: the assumptions every scenario shares", "Use CPAT");
  txt(s, [{ text: "Policies say what the government does. Settings say how the world and the model behave: existing taxes " +
    "and ETS, price controls, baseline subsidy phase-outs, elasticities, international prices, and more. They live on tab Settings, one row each.",
    options: { fontSize: 15.5 } }], M, 1.55, W - 2 * M, 0.85);
  const sets = [
    ["DEF", "Default", "The CPAT settings as shipped. Leave this column alone."],
    ["USR", "User defined", "Your own: type country-specific values in this column."],
    ["HPR", "High prices", "High international energy prices (blank = as DEF)."],
    ["LPR", "Low prices", "Low international energy prices (blank = as DEF)."],
  ];
  sets.forEach(([c, n, t], i) => {
    const x = M + i * 3.07;
    card(s, x, 2.55, 2.87, 2.0);
    txt(s, c, x + 0.2, 2.7, 1.2, 0.6, { fs: 28, bold: true, color: i === 1 ? H.accent5 : H.accent1 });
    txt(s, n, x + 1.3, 2.75, 1.5, 0.5, { fs: 15, bold: true, color: H.dk2, valign: "middle" });
    txt(s, t, x + 0.2, 3.4, 2.5, 1.1, { fs: 13.5 });
  });
  txt(s, bl([
    [B("Choose per scenario: "), T("each scenario names its settings code (User-defined: Settings scenario; Pre-defined: row SettingsScenario). So the same policy can run under high and under low prices.")],
    [B("Column In use "), T("shows the value the model uses now; the cell at the top shows the settings scenario in use.")],
    [B("Changing a setting changes the baseline too "), T("(when the baseline uses that code): check the Climate and Fiscal diagnostics again.")],
  ]), M, 4.8, W - 2 * M, 1.6, { fs: 14.5, space: 8 });

  // ===================== 15. Results
  s = slide("READ THE RESULTS", "Where the answers are", "Use CPAT");
  s.addImage({ path: IMG("dashboard_top.png"), x: M, y: 1.55, w: 6.4, h: 3.6, sizing: { type: "contain", w: 6.4, h: 3.6 },
    line: { color: H.accent6, width: 0.75 } });
  txt(s, "CPAT's Dashboard (public-data render)", M, 5.2, 6.4, 0.3, { fs: 11, italic: true, color: H.accent6 });
  table(s, [
    ["Tab", "Use it for"],
    ["Dashboard", "One country, four scenarios, one year: emissions, revenue, households, health"],
    ["Climate", "The country diagnostic before any policy: where emissions come from"],
    ["Fiscal", "The fiscal starting point: subsidies, energy revenue, effective carbon rates"],
    ["Comparison", "Choosing between packages: levels and changes vs the baseline"],
    ["Output Data", "Every number, by code, to copy into your own tables"],
  ], M + 6.7, 1.55, W - 2 * M - 6.7, [1.5, W - 2 * M - 8.2], { fs: 12.5, rowH: 0.55, boldFirst: true });
  note(s, "Charts look empty? Excel has not finished calculating: wait, or press F9. Copy charts into Word or PowerPoint as pictures to freeze them.", 5.75);

  // ===================== 16. The MT: what it does
  pres.addSection({ title: "The MT" });
  s = slide("THE MT", "The MT: run and compare many scenarios", "The MT");
  table(s, [
    ["MT tab", "What it is for"],
    ["Dashboard", "The countries and scenario IDs to run; Run scenarios; status; results at a glance"],
    ["Macros", "The buttons and their settings: activation, data, scenarios, run time, countries"],
    ["Scenarios", "The master scenario table (as CPAT's Pre-defined); Run? flags; Sync scenarios sends it to CPAT"],
    ["StoredResults", "Every run, stored as values (one row per output, year and scenario)"],
    ["Comparison", "The stored scenarios side by side"],
    ["OneCountry · ManyCountries · OneScenario", "Graphs of stored runs: one country over scenarios, countries side by side, one run in detail"],
    ["Indicators · Cover", "The indicators stored; the version and the modules"],
  ], M, 1.55, W - 2 * M, [3.6, W - 2 * M - 3.6], { fs: 13.5, rowH: 0.52, boldFirst: true });
  note(s, "Stored results are values: they do not change when you edit CPAT later. That is what makes them safe to compare and to report.", 6.0);

  // ===================== 17. MT workflow
  s = slide("THE MT · A BATCH RUN", "Run a batch in five steps", "The MT");
  const flow = [
    [fa.FaListUl, "List", "Countries and scenario IDs on the MT's Dashboard (e.g. Egypt: BASE, P1, P2, P3)"],
    [fa.FaExchangeAlt, "Sync", "Edited the MT's Scenarios tab? Press Sync scenarios (tab Macros)"],
    [fa.FaStopwatch, "Estimate", "Estimate run time: how long the batch will take"],
    [fa.FaPlay, "Run", "Run scenarios (Dashboard): CPAT calculates each, the MT stores the results"],
    [fa.FaChartBar, "Compare", "Comparison, OneCountry, ManyCountries, OneScenario; optional .csv"],
  ];
  for (let i = 0; i < flow.length; i++) {
    const [Ic, head, t] = flow[i];
    const x = M + i * 2.45;
    card(s, x, 1.7, 2.25, 3.6);
    s.addShape(pres.shapes.OVAL, { x: x + 0.72, y: 1.9, w: 0.8, h: 0.8, fill: { color: H.accent1 }, line: { color: H.accent1 } });
    s.addImage({ data: await icon(Ic, "FFFFFF"), x: x + 0.92, y: 2.1, w: 0.4, h: 0.4 });
    txt(s, (i + 1) + ". " + head, x + 0.15, 2.85, 1.95, 0.45, { fs: 17, bold: true, color: H.dk2, align: "center" });
    txt(s, t, x + 0.15, 3.35, 1.95, 1.9, { fs: 13, align: "center" });
    if (i < flow.length - 1) s.addShape(pres.shapes.RIGHT_ARROW, { x: x + 2.27, y: 3.3, w: 0.16, h: 0.35,
      fill: { color: H.accent6 }, line: { color: H.accent6 } });
  }
  txt(s, bl([
    "The baseline is run once per country and stored; the MT runs the scenarios in groups of the live scenarios (2-4 at a time).",
    "Results go to StoredResults; with Save results to .csv = Yes (tab Macros, Run settings) also to CPAT-AI_StoredResults.csv in active_workbooks.",
    "ETS scenarios are solved automatically (the permit price that meets the cap), so they take longer.",
  ]), M, 5.5, W - 2 * M, 1.4, { fs: 14, space: 5 });

  // ===================== 18. Macros
  s = slide("THE MT · ITS MACROS", "Every button, what it does and when", "The MT");
  table(s, [
    ["Button (tab)", "What it does", "When"],
    ["Activate Unified (Macros)", "Prepares CPAT from the release: years to 2041, live scenario groups, recalculation; saves both files into active_workbooks", "Once, and after each new release"],
    ["Update data (Macros)", "Loads the data tables into CPAT from the chosen source (Data source; Rows to load: Country or All countries)", "When the data change, or for another country"],
    ["Sync scenarios (Macros)", "Writes the MT's Scenarios tab into CPAT", "After editing the MT's Scenarios tab"],
    ["Estimate run time (Macros)", "Times one recalculation and estimates the batch", "Before a long run"],
    ["Reduce to countries (Macros)", "Keeps only the Dashboard's countries in the large data tables; saves a copy CPAT_<countries>_<n>s", "To slice for your country"],
    ["Expand advanced outputs (Macros)", "Fills the advanced rows of Output Data across every year", "When you need those rows over time"],
    ["Run scenarios (Dashboard)", "Runs every country x scenario listed and stores the results", "For batches and reports"],
  ], M, 1.5, W - 2 * M, [3.0, 6.3, W - 2 * M - 9.3], { fs: 12.5, rowH: 0.6, boldFirst: true });
  note(s, "Settings above each button on tab Macros: Live scenarios, Last year, Formulas · Data source, Rows to load · Keep data for · Save results to .csv.", 6.45);

  // ===================== 19. Data
  s = slide("THE DATA", "Data editions: which numbers are in your CPAT", "The MT");
  const eds = [
    ["Main", "The full real dataset (from CPAT 1.0). The default for country work inside the Bank.", H.accent1],
    ["Default", "The public edition: real public sources, proprietary tables replaced. Use it for anything shared outside.", H.accent4],
    ["EMM", "A small edition for Egypt, Morocco and Mexico, for tests and demonstrations.", H.accent6],
  ];
  eds.forEach(([n, t, col], i) => {
    const x = M + i * 4.1;
    card(s, x, 1.6, 3.85, 2.3);
    s.addShape(pres.shapes.RECTANGLE, { x, y: 1.6, w: 3.85, h: 0.12, fill: { color: col }, line: { color: col } });
    txt(s, n, x + 0.25, 1.85, 3.4, 0.55, { fs: 24, bold: true, color: col });
    txt(s, t, x + 0.25, 2.45, 3.4, 1.4, { fs: 14 });
  });
  txt(s, bl([
    [B("Choose: "), T("tab Macros, Data source, then Update data. Rows to load: Country (fast) or All countries (needed before a run over many countries).")],
    [B("Your own numbers? "), T("Do not overwrite the data tabs: use the settings (USR column) or the scenario inputs, and record the source.")],
    [B("Spotted a data problem? "), T("Tell the team: data problems are registered as issues and fixed in the shared data, for every user.")],
  ]), M, 4.2, W - 2 * M, 2.2, { fs: 15, space: 10 });

  // ===================== 20. Country work recipe
  pres.addSection({ title: "In practice" });
  s = slide("IN PRACTICE", "A country mission, start to finish", "In practice");
  const rec = [
    ["Get CPAT", "Sync; activate the MT (or copy published_cpat's files into active_workbooks)"],
    ["Slice", "Your country on the MT's Dashboard; Reduce to countries; pick it on CPAT's Dashboard"],
    ["Check the baseline", "Climate and Fiscal diagnostics against national statistics (emissions, subsidies, revenue)"],
    ["Adjust settings", "Country-specific assumptions in the USR column; set the scenarios to USR"],
    ["Build scenarios", "User-defined for drafting; keep each package as a Pre-defined column"],
    ["Run and store", "MT: list the scenario IDs, Run scenarios; every run stored as values"],
    ["Report", "Dashboard, Comparison and the MT's graphs; numbers from Output Data or StoredResults"],
    ["Keep and share", "Copy the reduced CPAT and the MT to OneDrive; report anything odd to the team"],
  ];
  rec.forEach(([h, t], i) => {
    const col = i < 4 ? 0 : 1, row = i % 4;
    step(s, M + col * 6.2, 1.65 + row * 1.2, 5.9, i + 1, h, t, { h: 1.0, fs: 13.5 });
  });

  // ===================== 21. Troubleshooting
  s = slide("WHEN THINGS GO WRONG", "Common problems and quick fixes", "In practice");
  table(s, [
    ["Problem", "Fix"],
    ["The buttons do nothing", "Macros are off: Enable Content; or right-click the file, Properties, Unblock. Only the MT has macros"],
    ["Excel is slow or runs out of memory", "Reduce to countries; fewer live scenarios (2 or 3); close other workbooks; 64-bit Excel"],
    ["Charts are empty or #N/A", "Calculation not finished: wait or press F9; check the country is picked on the Dashboard"],
    ["The MT cannot find CPAT", "Keep both files in the same folder (active_workbooks); activate again"],
    ["SyncRepo complains about a file", "You saved a tracked file (e.g. in published_cpat or master). Work on copies in active_workbooks"],
    ["A new release came out", "Sync, activate again, reduce again; re-enter your scenario from your Pre-defined column"],
    ["A number looks wrong", "Tell the CPAT team with the country, scenario, tab and cell: it is checked against the bug and issue registers"],
  ], M, 1.5, W - 2 * M, [3.8, W - 2 * M - 3.8], { fs: 13, rowH: 0.6, boldFirst: true });

  // ===================== 22. Other decks
  pres.addSection({ title: "More" });
  s = slide("THE OTHER DECKS", "Where to go next: the decks in docs/presentations", "More");
  const decks = [
    ["00", "How to use CPAT and the MT for country work", "This deck: get it, activate, slice, scenarios, settings, the MT", "Country teams", H.accent5],
    ["01", "CPAT-AI user guide", "The repository and its folders, key files, working with Claude and GitHub Copilot, pulling and pushing, bugs and issues, example prompts", "Anyone developing CPAT", H.accent1],
    ["02", "CPAT-AI Excel: pitch and guide", "What is new in 2.0 and better than CPAT 1.0, the tabs, User-defined page, Dashboard, Comparison, Climate and Fiscal diagnostics", "Users and managers", H.accent1],
    ["03", "CPAT-AI Python", "The same model in Python (Coded CPAT): how it works, testing, the Shiny app", "Analysts who code", H.accent4],
    ["04", "CPAT-AI status update", "Where the project stands, the next phase, the functions still to add", "Management", H.accent6],
  ];
  decks.forEach(([n, t, d, who, col], i) => {
    const y = 1.55 + i * 0.98;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: M, y, w: 0.85, h: 0.82, rectRadius: 0.06, fill: { color: col }, line: { color: col } });
    txt(s, n, M, y, 0.85, 0.82, { fs: 22, bold: true, color: "FFFFFF", align: "center", valign: "middle" });
    txt(s, [{ text: t, options: { bold: true, fontSize: 16, color: H.dk2, breakLine: true } },
      { text: d, options: { fontSize: 12.5 } }], M + 1.05, y, 8.6, 0.85);
    txt(s, who, M + 9.8, y, 2.3, 0.82, { fs: 13, italic: true, color: H.accent6, valign: "middle" });
  });
  txt(s, "Also: the website cpmodelai.github.io/docs (method, model, Excel, Python) · QuickStart.docx at the top of the repository · earlier decks in docs/presentations/archive.",
    M, 6.45, W - 2 * M, 0.45, { fs: 13, color: H.accent6 });

  // ===================== 23. In short
  s = pres.addSlide({ masterName: "DARK", sectionTitle: "More" });
  s.addText("IN SHORT", { placeholder: "kicker" });
  s.addText("Sync, activate, slice, scenario, run", { placeholder: "title" });
  s.addText([
    { text: "Get the files: SyncRepo.bat, or published_cpat for a quick start", options: { bullet: true, breakLine: true } },
    { text: "master → Activate → work in active_workbooks (yours, git-ignored, back it up)", options: { bullet: true, breakLine: true } },
    { text: "Your country on the MT's Dashboard, then Reduce to countries", options: { bullet: true, breakLine: true } },
    { text: "Policies on User-defined / Pre-defined; assumptions in Settings (USR)", options: { bullet: true, breakLine: true } },
    { text: "Batches and stored results: the MT's Run scenarios", options: { bullet: true } },
  ], { placeholder: "body", paraSpaceAfter: 8 });
  s.addImage({ data: await icon(fa.FaRoute, H.accent2), x: 9.9, y: 2.0, w: 2.6, h: 2.6 });

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote " + OUT);
}

main().catch((e) => { console.error(e); process.exit(1); });
