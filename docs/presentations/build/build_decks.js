// Two decks (user 2026-11-17: "simplify and update the slide decks"):
//   07_CPAT-AI_UserGuide_latest.pptx       the simplified user guide: folders, activating CPAT, working with Claude
//                                          and GitHub Copilot, push and pull, credits, bugs/issues/norms, key files,
//                                          modules, where the latest CPAT is, example prompts
//   08_CPAT-AI_Excel_PitchGuide_latest.pptx the pitch and guide to CPAT-AI Excel, for the new interface (2.0.0): tabs,
//                                          the User-defined page, Dashboard, Comparison, Climate and Macro-Fiscal
// Screenshots in img/: the workbook views are renders of the public-data release CPAT_v1.9.58 (edition Default, no
// proprietary values; same layout as 2.0.0) printed by LibreOffice; the Copilot screens are from GitHub's
// "GitHub Copilot app for Beginners" (MIT License), as in repo-tools/AI_Tools_Quick_Start.pptx.
//
//   NODE_PATH=<node_modules with pptxgenjs, react, react-dom, react-icons, sharp> \
//   PPTX_SKILL=<pptx skill dir> node docs/presentations/build/build_guide_decks.js
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const RDS = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");

const HERE = __dirname;
const IMG = (f) => path.join(HERE, "img", f);
const OUTDIR = path.join(HERE, "..");
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

async function icon(Comp, color) {
  const svg = RDS.renderToStaticMarkup(React.createElement(Comp, { color: "#" + color, size: 256 }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

function makePres(title, foot) {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  pres.title = title;
  pres.author = "WB CPAT Team";
  pres.company = "World Bank";
  const C = pres.SchemeColor;
  pres.defineSlideMaster({
    title: "DARK",
    background: { color: H.dk2 },
    objects: [
      { placeholder: { options: { name: "kicker", type: "body", x: M, y: 0.7, w: 7, h: 0.4, fontSize: 14, bold: true,
        color: C.accent2, charSpacing: 2, margin: 0 }, text: "" } },
      { placeholder: { options: { name: "title", type: "title", x: M, y: 1.2, w: 6.6, h: 2.4, fontSize: 40, bold: true,
        color: C.background1, align: "left", valign: "top", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "body", type: "body", x: M, y: 3.8, w: 6.2, h: 2.6, fontSize: 16,
        color: C.background2, valign: "top", margin: 0 }, text: "" } },
      { text: { text: foot, options: { x: M, y: 6.95, w: 9, h: 0.3, fontSize: 10, color: C.accent2, margin: 0 } } },
    ],
  });
  pres.defineSlideMaster({
    title: "CONTENT",
    background: { color: H.lt1 },
    margin: [0.5, 0.6, 0.6, 0.6],
    objects: [
      { placeholder: { options: { name: "kicker", type: "body", x: M, y: 0.35, w: 10, h: 0.3, fontSize: 12, bold: true,
        color: C.accent1, charSpacing: 2, margin: 0 }, text: "" } },
      { placeholder: { options: { name: "title", type: "title", x: M, y: 0.65, w: W - 2 * M, h: 0.75, fontSize: 30,
        bold: true, color: C.text2, align: "left", valign: "top", margin: 0 }, text: "" } },
      { text: { text: foot, options: { x: M, y: 7.05, w: 10, h: 0.3, fontSize: 10, color: C.accent6, margin: 0 } } },
    ],
    slideNumber: { x: W - M - 0.5, y: 7.05, w: 0.5, h: 0.3, fontSize: 10, color: H.accent6, align: "right" },
  });
  return { pres, C };
}

// ---------- helpers ----------
function head(s, kicker, title) {
  s.addText(kicker, { placeholder: "kicker" });
  s.addText(title, { placeholder: "title" });
}
function card(pres, C, s, x, y, w, h, opts = {}) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08,
    fill: { color: opts.fill || C.background2 }, line: { color: opts.line || opts.fill || C.background2, width: 1 },
    objectName: opts.name || "card" });
}
function bullets(items, o = {}) {
  return items.map((t, i) => {
    const runs = Array.isArray(t) ? t : [{ text: t }];
    return runs.map((r, j) => ({ text: r.text, options: { ...(r.options || {}),
      bullet: j === 0 ? (o.bullet === false ? false : { indent: 14 }) : undefined,
      breakLine: j === runs.length - 1 && i < items.length - 1, paraSpaceAfter: o.space || 6 } }));
  }).flat();
}
const B = (t) => ({ text: t, options: { bold: true } });
const T = (t) => ({ text: t });
const CODE = (t) => ({ text: t, options: { fontFace: "Courier New", fontSize: 13 } });

// a numbered circle + header + text row
function stepRow(pres, C, s, x, y, w, n, title, text, opts = {}) {
  const d = 0.5;
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: opts.circle || C.accent1 }, line: { color: opts.circle || C.accent1 } });
  s.addText(String(n), { x, y, w: d, h: d, align: "center", valign: "middle", fontSize: 16, bold: true, color: C.background1, margin: 0, isTextBox: true });
  s.addText([{ text: title, options: { bold: true, fontSize: opts.ts || 16, color: C.text2, breakLine: true } },
    { text, options: { fontSize: opts.fs || 14, color: C.text1 } }],
  { x: x + d + 0.2, y: y - 0.05, w: w - d - 0.2, h: opts.h || 0.95, valign: "top", margin: 0, isTextBox: true, paraSpaceAfter: 2 });
}
function iconCircle(pres, C, s, data, x, y, d = 0.6, fill) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill || C.accent1 }, line: { color: fill || C.accent1 } });
  s.addImage({ data, x: x + d * 0.22, y: y + d * 0.22, w: d * 0.56, h: d * 0.56 });
}
function shot(s, file, x, y, w, pw, ph, alt, C, caption) {
  const h = w * ph / pw;
  s.addImage({ path: IMG(file), x, y, w, h, altText: alt, line: { color: H.accent6, width: 0.75 } });
  if (caption) s.addText(caption, { x, y: y + h + 0.08, w, h: 0.3, fontSize: 11, italic: true, color: C.accent6, margin: 0, isTextBox: true });
  return h;
}

// a slide of 2-4 cards (icon, title, text) and an optional line under them
function cardsSlide(pres, C, sec, kicker, title, items, foot) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: sec });
  head(s, kicker, title);
  const n = items.length, gap = 0.25, w = (W - 2 * M - gap * (n - 1)) / n, h = 3.6;
  items.forEach(([t, d, im, col], i) => {
    const x = M + i * (w + gap);
    card(pres, C, s, x, 1.65, w, h, { name: t });
    iconCircle(pres, C, s, im, x + 0.3, 1.9, 0.75, col);
    s.addText(t, { x: x + 0.3, y: 2.85, w: w - 0.6, h: 0.75, fontSize: 19, bold: true, color: C.text2, valign: "top", margin: 0, isTextBox: true });
    const body = Array.isArray(d) ? bullets(d, { space: 6 }) : d;
    s.addText(body, { x: x + 0.3, y: 3.6, w: w - 0.6, h: h - 2.05, fontSize: 14, color: C.text1, valign: "top", margin: 0, isTextBox: !Array.isArray(d) });
  });
  if (foot) s.addText(foot, { x: M, y: 5.55, w: W - 2 * M, h: 0.7, fontSize: 15, color: C.text1, margin: 0, isTextBox: true });
  return s;
}
// a picture on the left, bullets on the right (or the picture across the top and bullets below)
function picSlide(pres, C, sec, kicker, title, file, pw, ph, alt, caption, items, wide) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: sec });
  head(s, kicker, title);
  if (wide) {
    const h = shot(s, file, M, 1.6, W - 2 * M, pw, ph, alt, C, caption);
    s.addText(bullets(items, { space: 6 }), { x: M, y: 1.6 + h + 0.5, w: W - 2 * M, h: 6.85 - (1.6 + h + 0.5), fontSize: 14, color: C.text1, valign: "top", margin: 0 });
  } else {
    let w = 7.2; if (w * ph / pw > 4.9) w = 4.9 * pw / ph;
    shot(s, file, M, 1.6, w, pw, ph, alt, C, caption);
    s.addText(bullets(items, { space: 8 }), { x: M + w + 0.4, y: 1.6, w: W - 2 * M - w - 0.4, h: 5.2, fontSize: 14, color: C.text1, valign: "top", margin: 0 });
  }
  return s;
}
// big numbers
function statsSlide(pres, C, sec, kicker, title, stats, foot) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: sec });
  head(s, kicker, title);
  const n = stats.length, cols = n > 4 ? 3 : n, gap = 0.25, w = (W - 2 * M - gap * (cols - 1)) / cols;
  const rowsN = Math.ceil(n / cols), h = rowsN > 1 ? 2.1 : 3.0;
  stats.forEach(([v, l], i) => {
    const x = M + (i % cols) * (w + gap), y = 1.65 + Math.floor(i / cols) * (h + 0.2);
    card(pres, C, s, x, y, w, h, { name: "stat " + v });
    s.addText(v, { x: x + 0.3, y: y + 0.15, w: w - 0.6, h: 1.0, fontSize: 44, bold: true, color: C.accent1, margin: 0, isTextBox: true, fontFace: "Arial" });
    s.addText(l, { x: x + 0.3, y: y + 1.15, w: w - 0.6, h: h - 1.25, fontSize: 14, color: C.text1, valign: "top", margin: 0, isTextBox: true });
  });
  if (foot) s.addText(foot, { x: M, y: 6.3, w: W - 2 * M, h: 0.6, fontSize: 13, italic: true, color: C.accent6, margin: 0, isTextBox: true });
  return s;
}
// a two-column table slide
function tableSlide(pres, C, sec, kicker, title, rows, colW, foot) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: sec });
  head(s, kicker, title);
  s.addTable(rows.map((r, i) => r.map((c, j) => ({ text: c, options: {
    bold: i === 0 || j === 0, color: i === 0 ? H.lt1 : H.dk1, fill: { color: i === 0 ? H.dk2 : (i % 2 ? H.lt1 : H.lt2) }, fontSize: 15 } }))),
  { x: M, y: 1.6, w: W - 2 * M, colW, rowH: 0.52, border: { type: "solid", pt: 0.5, color: "D5DDD8" }, valign: "middle", margin: [3, 6, 3, 6] });
  if (foot) s.addText(foot, { x: M, y: 6.45, w: W - 2 * M, h: 0.45, fontSize: 13, italic: true, color: C.accent6, margin: 0, isTextBox: true });
  return s;
}

// =====================================================================================
// DECK 1: the user guide
// =====================================================================================
async function deckGuide() {
  const FOOT = "CPAT-AI user guide  ·  WB CPAT Team  ·  November 2026  ·  github.com/cpmodelai/cpat_ai";
  const { pres, C } = makePres("CPAT-AI: a simple user guide", FOOT);
  const ic = {
    folder: await icon(fa.FaFolderOpen, H.lt1), file: await icon(fa.FaFileAlt, H.lt1), play: await icon(fa.FaPlay, H.lt1),
    robot: await icon(fa.FaRobot, H.lt1), github: await icon(fa.FaGithub, H.lt1), sync: await icon(fa.FaSyncAlt, H.lt1),
    bug: await icon(fa.FaBug, H.lt1), cubes: await icon(fa.FaCubes, H.lt1), star: await icon(fa.FaStar, H.lt1),
    coins: await icon(fa.FaCoins, H.lt1), warn: await icon(fa.FaExclamationTriangle, H.lt1), chat: await icon(fa.FaComments, H.lt1),
    book: await icon(fa.FaBook, H.lt1), cloud: await icon(fa.FaCloud, H.lt1), laptop: await icon(fa.FaLaptopCode, H.lt1),
  };  Object.assign(ic, { db: await icon(fa.FaDatabase, H.lt1), lock: await icon(fa.FaLock, H.lt1), globe: await icon(fa.FaGlobe, H.lt1),
    wrench: await icon(fa.FaWrench, H.lt1), py: await icon(fa.FaPython, H.lt1), excel: await icon(fa.FaFileExcel, H.lt1),
    user: await icon(fa.FaUserCog, H.lt1), life: await icon(fa.FaLifeRing, H.lt1), flag: await icon(fa.FaFlag, H.lt1) });


  // 1 title
  pres.addSection({ title: "Start" });
  let s = pres.addSlide({ masterName: "DARK", sectionTitle: "Start" });
  s.addText("CPAT-AI · USER GUIDE", { placeholder: "kicker" });
  s.addText("Working on CPAT-AI with Claude and GitHub Copilot", { placeholder: "title" });
  s.addText("Where things are, how to open the latest CPAT, and how to develop it with an AI assistant without losing work.", { placeholder: "body" });
  const steps = [["Find", "the repo's folders and key files", ic.folder], ["Open", "the latest CPAT from the MT", ic.play],
    ["Develop", "with Claude Code or GitHub Copilot", ic.robot], ["Share", "pull before, push after, every time", ic.sync]];
  steps.forEach(([t, d, im], i) => {
    const y = 1.3 + i * 1.3;
    iconCircle(pres, C, s, im, 8.2, y, 0.75, C.accent1);
    s.addText([{ text: t, options: { bold: true, fontSize: 20, color: C.background1, breakLine: true } },
      { text: d, options: { fontSize: 14, color: C.background2 } }], { x: 9.15, y: y - 0.05, w: 3.7, h: 0.9, margin: 0, valign: "middle", isTextBox: true });
  });
  s.addNotes("A short, practical guide. The detail is in README.md, QuickStart.docx and the website; this deck only points to them.");

  cardsSlide(pres, C, "Start", "WHAT IS NEW", "CPAT-AI 2.0 at a glance", [
    ["One unified workbook", ["Mitigation with Distribution, Air pollution and Transport co-benefits inside", "Run from the MT, the only file with macros"], ic.excel, C.accent1],
    ["A new interface", ["Dashboard, User-defined page, Comparison, Climate and Macro-Fiscal diagnostics", "Settings scenarios: DEF, USR, HPR, LPR"], ic.star, C.accent4],
    ["Real data inside", ["The Main edition (real values) is the default in the repository", "Anything shared outside uses the public Default edition"], ic.db, C.accent5],
    ["Python twin", ["Coded CPAT 2.0.0p2.3 replicates every module of 2.0.0", "104 tests pass"], ic.py, C.accent6],
  ], [B("Three decks now: "), T("this user guide, the CPAT-AI Excel pitch and guide, and the CPAT-AI Python deck (docs/presentations/).")]);

  // 2 folder structure
  pres.addSection({ title: "The repository" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "The repository" });
  head(s, "THE REPOSITORY", "Six folders: know where each thing lives");
  const folders = [
    ["excel_ai/", "The Excel prototypes: the MT (master/), the modules (mitigation/, distribution/, air_pollution/, transport_cobenefits/), the unified CPAT (unified/), the rules (core/)"],
    ["coded/", "Coded CPAT: the Python model and its Shiny app, the twin of the Excel model"],
    ["data/", "All data, as editions (Default public, Main real), the catalogue and the data scripts"],
    ["docs/", "The documentation website and the latest slide decks (docs/presentations/)"],
    ["users/", "Personal working folders: your own runs and notes"],
    ["other/", "Research papers, earlier prototypes and notes"],
  ];
  folders.forEach(([f, d], i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = M + col * 6.15, y = 1.65 + row * 1.6;
    card(pres, C, s, x, y, 5.95, 1.4, { name: "folder " + f });
    iconCircle(pres, C, s, ic.folder, x + 0.25, y + 0.3, 0.6);
    s.addText([{ text: f, options: { bold: true, fontSize: 18, color: C.text2, fontFace: "Courier New", breakLine: true } },
      { text: d, options: { fontSize: 13, color: C.text1 } }], { x: x + 1.05, y: y + 0.12, w: 4.75, h: 1.2, valign: "top", margin: 0, isTextBox: true });
  });
  s.addText([B("At the root: "), T("README.md (start here), AGENTS.md (what the AI reads first), QuickStart.docx, and SyncRepo.bat / ResetToMain.bat / StashChanges.bat.")],
    { x: M, y: 6.5, w: W - 2 * M, h: 0.4, fontSize: 14, color: C.text1, margin: 0, isTextBox: true });

  // 3 key human-readable files
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "The repository" });
  head(s, "THE REPOSITORY", "The key human-readable files, and where they are");
  const rows = [
    ["What", "File", "Where"],
    ["Start here", "README.md, QuickStart.docx", "root"],
    ["The rules (norms)", "NORMS.md", "excel_ai/core/ · coded/"],
    ["Legacy bugs, all fixed", "BUGS.md", "excel_ai/core/"],
    ["Legacy issues, kept and tracked", "ISSUES.md", "excel_ai/core/"],
    ["The to-do list", "TODO.md", "excel_ai/core/ · coded/"],
    ["What was done, task by task", "CAVEATS.md, CHANGELOG.md", "excel_ai/core/ · coded/"],
    ["Every code and input", "CPAT-AI_CodesMap_latest.xlsx", "excel_ai/core/"],
    ["Bug book and tracker", "Legacy_CPAT_Bugs_v*.docx, Tracking_Legacy_Bugs_v*.xlsx", "excel_ai/bugs_and_legacy/"],
    ["Data editions and tables", "README.md, editions.csv, catalog.csv", "data/"],
    ["The website", "cpmodelai.github.io/docs (offline: index.html)", "docs/"],
  ];
  s.addTable(rows.map((r, i) => r.map((c, j) => ({ text: c, options: {
    bold: i === 0 || j === 0, color: i === 0 ? H.lt1 : H.dk1, fill: { color: i === 0 ? H.dk2 : (i % 2 ? H.lt1 : H.lt2) },
    fontFace: j === 1 && i > 0 ? "Courier New" : undefined, fontSize: j === 1 && i > 0 ? 12 : 13 } }))),
  { x: M, y: 1.6, w: W - 2 * M, colW: [3.4, 5.3, 3.43], rowH: 0.42, border: { type: "solid", pt: 0.5, color: "D5DDD8" }, valign: "middle", margin: [2, 6, 2, 6] });
  s.addText("The .md files are plain text: GitHub shows them formatted, and the AI reads the same files, so people and AI follow the same rules.",
    { x: M, y: 6.45, w: W - 2 * M, h: 0.4, fontSize: 13, italic: true, color: C.accent6, margin: 0, isTextBox: true });

  cardsSlide(pres, C, "The repository", "THE REPOSITORY · DATA", "Data come in editions; real data stay inside", [
    ["Main", ["The full CPAT dataset with real values (from CPAT 1.0pre_468b)", "data/Main: the default since 2.0"], ic.db, C.accent5],
    ["Default", ["Public tables and open estimates for the six proprietary tables", "data/Default: for anything shared"], ic.globe, C.accent1],
    ["The rule", ["Never put real data in anything shared outside: decks, the website, the Shiny app, emails", "Check the changed files before every push"], ic.lock, C.text2],
  ], [B("Proprietary tables: "), T("Balances, EnergyCons, Prices_dom, IO_GTAP, HHSurvey, Tran_Scatter. data/editions.csv lists every edition; data/README.md explains them.")]);
  picSlide(pres, C, "The repository", "THE REPOSITORY · DOCUMENTATION", "The website: the documentation in one place",
    "website_home.png", 2100, 1290, "The documentation website", "cpmodelai.github.io/docs (offline: docs/index.html)", [
      [B("Home: "), T("start here, the three decks, recent changes")],
      [B("Methodology: "), T("the CPAT method")],
      [B("CPAT-AI Model: "), T("each block of the model, shared by Excel and Python")],
      [B("CPAT Excel-AI: "), T("the workbooks, bugs and issues, how to use them")],
      [B("CPAT Python: "), T("Coded CPAT: quick start, testing, every equation")],
    ]);

  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "The repository" });
  head(s, "GETTING SET UP · ONCE", "What you need on your laptop");
  [["GitHub account", "with access to cpmodelai/cpat_ai (ask the team)"], ["Git and GitHub Desktop", "installed; see repo-tools/SetUpGuide.docx"],
   ["The repos", "repo-tools/CloneRepos.bat puts every repo in C:\\Users\\<you>\\Repos"], ["Excel 365", "macros enabled for the MT"],
   ["GitHub Copilot app", "from the Microsoft Store, signed in with GitHub"], ["Claude", "an account on claude.ai with Claude Code, connected to GitHub"]].forEach(([t, d], i) => {
    const col = i % 2, row = Math.floor(i / 2);
    stepRow(pres, C, s, M + col * 6.15, 1.75 + row * 1.5, 5.9, i + 1, t, d, { h: 1.2 });
  });
  s.addText("Python (optional): conda env create -f coded/environment.yml, or Run_CPAT_App.bat for the app.", { x: M, y: 6.4, w: W - 2 * M, h: 0.4, fontSize: 13, italic: true, color: C.accent6, margin: 0, isTextBox: true });

  // 4 where the latest CPAT is + activate
  pres.addSection({ title: "Using CPAT" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Using CPAT" });
  head(s, "USING CPAT", "Where the latest CPAT is, and how to activate it");
  card(pres, C, s, M, 1.6, 5.4, 4.9, { name: "where" });
  iconCircle(pres, C, s, ic.star, M + 0.3, 1.85, 0.6);
  s.addText("Always the file named _latest", { x: M + 1.05, y: 1.9, w: 4.2, h: 0.5, fontSize: 18, bold: true, color: C.text2, margin: 0, isTextBox: true });
  s.addText(bullets([
    [B("The MT (start here): "), CODE("excel_ai/master/CPAT-AI-MT_latest.xlsm")],
    [B("The unified CPAT: "), CODE("excel_ai/unified/CPAT_latest.xlsx")],
    [B("Each module: "), T("its folder, e.g. "), CODE("mitigation/CPAT-AI-Mitigation_latest.xlsx")],
    [B("Older versions: "), T("each module's releases/ folder")],
    [B("Python: "), CODE("coded/"), T(" (Run_CPAT_App.bat opens the Shiny app)")],
  ]), { x: M + 0.3, y: 2.6, w: 4.85, h: 3.7, fontSize: 14, color: C.text1, valign: "top", margin: 0 });
  const act = [
    ["Pull first", "Run SyncRepo.bat (or Fetch in GitHub Desktop) so you have the latest files."],
    ["Open the MT", "excel_ai/master/CPAT-AI-MT_latest.xlsm in Excel 365, macros on."],
    ["Activate", "Tab Settings and Macros: Activate Unified. It saves your working copies in excel_ai/active_workbooks/."],
    ["Work in the .xlsb", "From then on open active_workbooks/: git never tracks it, so updates never clash with your runs."],
  ];
  act.forEach(([t, d], i) => stepRow(pres, C, s, 6.4, 1.7 + i * 1.2, 6.3, i + 1, t, d));
  s.addNotes("When a newer Master arrives with a pull, the MT's Dashboard says 'Newer Master available': open the new one from excel_ai/master/ and activate again. Never commit workbooks holding real data.");

  // 5 modules stitched
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Using CPAT" });
  head(s, "USING CPAT", "Modules are developed on their own, then stitched together");
  const mods = [["Mitigation", "the hub: prices, energy, emissions, revenue, GDP"], ["Distribution", "households by decile"],
    ["Air pollution", "deaths and welfare"], ["Transport co-benefits", "congestion, accidents"]];
  mods.forEach(([t, d], i) => {
    const y = 1.7 + i * 1.15;
    card(pres, C, s, M, y, 4.0, 0.95, { fill: i === 0 ? C.accent1 : C.background2, name: "module " + t });
    s.addText([{ text: t, options: { bold: true, fontSize: 16, color: i === 0 ? C.background1 : C.text2, breakLine: true } },
      { text: d, options: { fontSize: 12, color: i === 0 ? C.background1 : C.text1 } }],
    { x: M + 0.2, y, w: 3.7, h: 0.95, valign: "middle", margin: 0, isTextBox: true });
    s.addShape(pres.shapes.LINE, { x: M + 4.0, y: y + 0.475, w: 0.5, h: 0, line: { color: H.accent6, width: 1.5 } });
  });
  s.addShape(pres.shapes.LINE, { x: M + 4.5, y: 2.175, w: 0, h: 3.45, line: { color: H.accent6, width: 1.5 } });
  s.addShape(pres.shapes.LINE, { x: M + 4.5, y: 3.4, w: 0.5, h: 0, line: { color: H.accent6, width: 1.5, endArrowType: "triangle" } });
  card(pres, C, s, 5.6, 2.6, 2.8, 1.6, { fill: C.text2, name: "unified" });
  s.addText([{ text: "Unified CPAT", options: { bold: true, fontSize: 18, color: C.background1, breakLine: true } },
    { text: "one workbook, built by build_unified.py at every release", options: { fontSize: 12, color: C.background2 } }],
  { x: 5.75, y: 2.6, w: 2.5, h: 1.6, valign: "middle", align: "center", margin: 0, isTextBox: true });
  s.addText(bullets([
    "Each module has its own folder, builder, checks, version and releases/: one person or one AI session can change it without touching the others.",
    "Modules talk only through named export blocks (Mitigation section 11 for Distribution and Transport, section 15 for Air pollution), by code.",
    "At a release the builder stitches them into the unified CPAT, and Coded CPAT is checked number for number against it.",
  ], { space: 10 }), { x: 8.8, y: 1.7, w: 3.95, h: 4.6, fontSize: 14, color: C.text1, valign: "top", margin: 0 });
  s.addNotes("Arrows: every module feeds the unified workbook. The MT (excel_ai/master) activates the unified CPAT by default.");

  cardsSlide(pres, C, "Using CPAT", "USING CPAT · PYTHON", "The Python version runs the same model", [
    ["The app", ["Double-click coded/Run_CPAT_App.bat", "Tabs: User-defined scenario, Scenarios and run, Data, Dashboard, Results"], ic.laptop, C.accent1],
    ["The command line", ["python run_model.py", "cpat-run --countries EGY MAR --scenarios 1 2"], ic.py, C.accent4],
    ["Its own deck", ["03_CPAT-AI_Python_latest.pptx", "and the website's CPAT Python tab"], ic.book, C.accent6],
  ]);

  // 6 two AI tools + credits
  pres.addSection({ title: "Developing with AI" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Developing with AI" });
  head(s, "DEVELOPING WITH AI", "Two assistants, one set of rules");
  const tools = [
    ["GitHub Copilot app", ic.laptop, C.accent4, [
      "Runs locally on your laptop, on your copy of the repo",
      "Pick Claude models (Sonnet everyday, Opus for hard work)",
      "Interactive mode: it asks before it acts",
      "Credits: you get USD 19; type /context to see what a session has used"]],
    ["Claude Code", ic.cloud, C.accent1, [
      "Runs in the cloud (claude.ai/code) on a fresh clone of the repo",
      "Works on its own branch: you merge to main when happy",
      "Good for long tasks: builds, checks, releases",
      "In our experience Claude credits seem to last longer (we are not sure why)"]],
  ];
  tools.forEach(([t, im, col, items], i) => {
    const x = M + i * 6.15;
    card(pres, C, s, x, 1.6, 5.95, 4.2, { name: t });
    iconCircle(pres, C, s, im, x + 0.3, 1.85, 0.7, col);
    s.addText(t, { x: x + 1.2, y: 1.95, w: 4.5, h: 0.5, fontSize: 22, bold: true, color: C.text2, margin: 0, isTextBox: true });
    s.addText(bullets(items, { space: 10 }), { x: x + 0.3, y: 2.8, w: 5.4, h: 2.9, fontSize: 15, color: C.text1, valign: "top", margin: 0 });
  });
  s.addText([B("Both read AGENTS.md first, "), T("which sends them to NORMS.md, BUGS.md, ISSUES.md and TODO.md: the AI follows the same rules as people and logs what it did.")],
    { x: M, y: 6.1, w: W - 2 * M, h: 0.6, fontSize: 14, color: C.text1, margin: 0, isTextBox: true });

  // 7 copilot: current checkout, not worktree
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Developing with AI" });
  head(s, "GITHUB COPILOT · LOCAL MODE", "Run Copilot on the repo itself, not on a worktree");
  shot(s, "copilot_checkout.png", M, 1.65, 6.6, 1000, 480, "Copilot workspace selector with Current checkout highlighted", C,
    "Under the prompt box: (1) the workspace selector, (2) Current checkout. Screenshot: GitHub, MIT License");
  iconCircle(pres, C, s, ic.warn, 7.6, 1.7, 0.6, C.accent5);
  s.addText("Choose Current checkout", { x: 8.35, y: 1.75, w: 4.4, h: 0.5, fontSize: 20, bold: true, color: C.text2, margin: 0, isTextBox: true });
  s.addText(bullets([
    "New worktree (the default) makes a hidden copy of the repo in another folder for each session. Your Excel, SyncRepo.bat and GitHub Desktop never see it, and work there is easily lost.",
    "Current checkout edits C:\\Users\\<you>\\Repos\\cpat_ai directly: what Copilot changes is what you open, test and push.",
    "Add the repo once: + next to Projects, Open folder, pick Repos\\cpat_ai.",
    "Old worktrees piling up? repo-tools\\Advanced\\PruneWorktrees.bat removes them (it warns first).",
  ], { space: 9 }), { x: 7.6, y: 2.5, w: 5.15, h: 4.2, fontSize: 14, color: C.text1, valign: "top", margin: 0 });

  picSlide(pres, C, "Developing with AI", "GITHUB COPILOT · EVERY SESSION", "Set up a Copilot session in three clicks",
    "copilot_addproject.png", 1300, 400, "Adding a project folder in the Copilot app", "+ next to Projects, then Open folder. Screenshot: GitHub, MIT License", [
      [B("1 Project: "), T("Repos\\cpat_ai, added once; then + next to it for a new session")],
      [B("2 Workspace: "), T("Current checkout (previous slide)")],
      [B("3 Model: "), T("switch Auto off; Claude Sonnet for everyday work, Claude Opus for hard problems")],
      [B("4 Mode: "), T("Interactive, so it asks before it acts")],
      [B("Then: "), T("one task per prompt; read the Changes view before you keep anything")],
    ]);
  cardsSlide(pres, C, "Developing with AI", "CREDITS", "AI credits are limited: spend them well", [
    ["USD 19 to start", ["Copilot gives you USD 19 of AI credits", "Type /context to see what a session has used"], ic.coins, C.accent5],
    ["Claude seems to last longer", ["In our experience Claude credits go further for the same work", "We are not sure why: compare with /context and the usage page"], ic.cloud, C.accent1],
    ["Save credits", ["Short, precise prompts; one task each", "Opus only for hard problems", "Start a new session for a new task"], ic.star, C.accent4],
  ]);

  // 8 claude code
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Developing with AI" });
  head(s, "CLAUDE CODE", "Claude Code works on a branch: get it into main");
  const cc = [
    ["Start a session", "claude.ai/code, pick the cpmodelai/cpat_ai repository, describe the task."],
    ["Let it work", "It clones the repo, edits, runs the checks, and commits to its own claude/... branch."],
    ["Ask it to push", "Nothing it does survives the session unless it is pushed: say \"commit and push\"."],
    ["Merge to main", "Open a pull request and merge it (or ask Claude to). Only main is what everyone, and the next session, starts from."],
    ["Pull on your laptop", "Run SyncRepo.bat so your copy and your Excel have the new files."],
  ];
  cc.forEach(([t, d], i) => stepRow(pres, C, s, M, 1.65 + i * 0.98, 7.2, i + 1, t, d, { h: 0.9 }));
  card(pres, C, s, 8.4, 1.65, 4.33, 4.7, { fill: C.text2, name: "why main" });
  s.addText([{ text: "Why main matters", options: { bold: true, fontSize: 20, color: C.background1, breakLine: true } },
    { text: " ", options: { fontSize: 8, breakLine: true } },
    { text: "A branch nobody merges is invisible: the next Claude session, Copilot and your colleagues all start from main.", options: { fontSize: 14, color: C.background2, breakLine: true } },
    { text: " ", options: { fontSize: 8, breakLine: true } },
    { text: "Two branches that drift apart end in merge conflicts, which are hardest in workbooks.", options: { fontSize: 14, color: C.background2, breakLine: true } },
    { text: " ", options: { fontSize: 8, breakLine: true } },
    { text: "So: small tasks, merged to main the same day.", options: { fontSize: 14, bold: true, color: C.accent2 } }],
  { x: 8.7, y: 1.85, w: 3.8, h: 4.3, valign: "top", margin: 0, isTextBox: true });

  // 9 push and pull
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Developing with AI" });
  head(s, "SHARING WORK", "Pull before you start, push when you stop");
  const loop = [["Pull", "get everyone's latest", ic.sync], ["Work", "you, Copilot or Claude", ic.robot], ["Push", "send yours to main", ic.github]];
  loop.forEach(([t, d, im], i) => {
    const x = M + 0.2 + i * 2.15;
    iconCircle(pres, C, s, im, x + 0.45, 1.75, 0.9, i === 1 ? C.accent4 : C.accent1);
    s.addText([{ text: t, options: { bold: true, fontSize: 18, color: C.text2, breakLine: true } }, { text: d, options: { fontSize: 13, color: C.text1 } }],
      { x, y: 2.75, w: 1.8, h: 0.9, align: "center", margin: 0, isTextBox: true });
    if (i < 2) s.addShape(pres.shapes.LINE, { x: x + 1.5, y: 2.2, w: 0.75, h: 0, line: { color: H.accent6, width: 2, endArrowType: "triangle" } });
  });
  s.addText(bullets([
    "Pull at the start of every session, push at the end: work only on your laptop is lost to the team.",
    "Never commit the files Excel makes while a workbook is open (~$...), nor anything holding real data.",
    "A conflict on a generated .xlsb: take either side and activate again.",
  ], { space: 9 }), { x: M, y: 4.0, w: 6.3, h: 2.6, fontSize: 14, color: C.text1, valign: "top", margin: 0 });
  const tl = [
    ["repo-tools (double-click)", "SyncRepo.bat: pull, commit and push this repo\nSyncAllRepos.bat: every repo in RepoDirectory.csv\nStashChanges.bat, ResetToMain.bat: put work aside, or start clean"],
    ["GitHub Desktop (the same, with buttons)", "Fetch / Pull origin, then Commit to main, then Push origin. Shows clashing files when a merge conflicts."],
    ["Or ask the AI", "\"Pull, then commit and push with a clear message.\" It shows what it will commit: check it."],
  ];
  tl.forEach(([t, d], i) => {
    const y = 1.65 + i * 1.65;
    card(pres, C, s, 7.2, y, 5.53, 1.5, { name: t });
    s.addText([{ text: t, options: { bold: true, fontSize: 15, color: C.text2, breakLine: true } }, { text: d, options: { fontSize: 12, color: C.text1 } }],
      { x: 7.4, y: y + 0.1, w: 5.15, h: 1.3, valign: "top", margin: 0, isTextBox: true });
  });

  tableSlide(pres, C, "Developing with AI", "SHARING WORK · WHEN THINGS GO WRONG", "When git gets in the way", [
    ["Problem", "What to do"],
    ["Push rejected: someone pushed first", "Pull (SyncRepo.bat does it), then push again"],
    ["Merge conflict in a .md or code file", "GitHub Desktop shows the files; keep both sides' changes, or ask the AI to resolve it"],
    ["Conflict in a generated .xlsb", "Take either side, then Activate again"],
    ["Local changes you do not want", "StashChanges.bat puts them aside; ResetToMain.bat starts clean from main"],
    ["Too many Copilot worktrees", "repo-tools\\Advanced\\PruneWorktrees.bat"],
    ["Still stuck", "repo-tools: Git_Basics.docx and Git_Rescue_Cheat_Sheet.docx; or ask Claude"],
  ], [4.2, 7.93]);

  // 10 bugs, issues, norms
  pres.addSection({ title: "Rules" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Rules" });
  head(s, "RULES", "Bugs, issues and norms");
  const bi = [
    ["Bug", "CPAT 1.0 does not compute what it means to (a wrong row, a unit twice).", "Never replicated: fixed, with a cell plan. BUGS.md", C.accent5],
    ["Issue", "CPAT 1.0 does what it means to, but the choice is open to question.", "Always replicated as CPAT 1.0, and tracked. ISSUES.md", C.accent4],
    ["Upgrade", "Newer data or a better method.", "Proposed and tracked (U codes); the team decides. Never silent.", C.accent1],
  ];
  bi.forEach(([t, d, r, col], i) => {
    const x = M + i * 4.1;
    card(pres, C, s, x, 1.6, 3.9, 2.75, { name: t });
    iconCircle(pres, C, s, ic.bug, x + 0.25, 1.8, 0.6, col);
    s.addText(t, { x: x + 1.0, y: 1.85, w: 2.7, h: 0.5, fontSize: 22, bold: true, color: C.text2, margin: 0, isTextBox: true });
    s.addText([{ text: d, options: { fontSize: 14, color: C.text1, breakLine: true } }, { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: r, options: { fontSize: 14, bold: true, color: C.text2 } }], { x: x + 0.25, y: 2.55, w: 3.45, h: 1.7, valign: "top", margin: 0, isTextBox: true });
  });
  s.addText("Norms to know first (NORMS.md)", { x: M, y: 4.6, w: 6, h: 0.4, fontSize: 18, bold: true, color: C.text2, margin: 0, isTextBox: true });
  s.addText(bullets([
    "When in doubt, call it an issue: replicate it and ask.",
    "Never edit a released workbook: build the next version.",
    "Same input names and codes in Excel and Python (e.g. CPLevelTarget).",
  ]), { x: M, y: 5.05, w: 6.0, h: 1.8, fontSize: 14, color: C.text1, valign: "top", margin: 0 });
  s.addText(bullets([
    "Short formulas (target 90 characters), CPAT colours: inputs lemon, codes light blue.",
    "Real data never leave the repository: shared files use the public Default edition.",
    "Every finished task is logged (CAVEATS.md, CHANGELOG.md).",
  ]), { x: 6.75, y: 5.05, w: 6.0, h: 1.8, fontSize: 14, color: C.text1, valign: "top", margin: 0 });

  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Rules" });
  head(s, "RULES · RAISING A BUG OR ISSUE", "Found something odd? Register it the same day");
  [["Check", "Is it already in BUGS.md or ISSUES.md? Search the code or the description."],
   ["Decide", "Bug (not what CPAT 1.0 means to do) or issue (open to question)? In doubt: an issue."],
   ["Register", "Ask the AI: \u201CRegister this as an issue, with the rows and an example.\u201D It goes in the register, which writes the tracker and BUGS.md / ISSUES.md."],
   ["Fix or track", "A bug is fixed in the next release, with a cell plan for CPAT 1.0; an issue is replicated and waits for a team decision."]].forEach(([t, d], i) =>
    stepRow(pres, C, s, M, 1.7 + i * 1.2, 7.4, i + 1, t, d, { h: 1.1 }));
  card(pres, C, s, 8.6, 1.7, 4.13, 4.6, { name: "codes" });
  s.addText([{ text: "The codes", options: { bold: true, fontSize: 18, color: C.text2, breakLine: true } },
    ...bullets(["L: legacy bug", "M: prototype bug, CB: Python bug", "MId / MIm: major issue, data / method", "I: issue", "MD: major difference", "U: upgrade", "DOC: documentation bug"], { space: 4 })],
  { x: 8.85, y: 1.85, w: 3.7, h: 4.3, fontSize: 14, color: C.text1, valign: "top", margin: 0 });

  // 11 example prompts
  pres.addSection({ title: "Example prompts" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Example prompts" });
  head(s, "EXAMPLE PROMPTS", "Ask about the model");
  const p1 = [
    "Explain how the retail price of diesel is built up, with the cells in the mitigation workbook.",
    "Which legacy bugs affect revenue? Show the CPAT 1.0 formula and ours side by side.",
    "Where does the Distribution module get its fuel price changes from?",
    "What changed between 1.9.61 and 2.0.0? Read CHANGELOG.md and CAVEATS.md.",
    "Is this difference from CPAT 1.0 a bug or an issue? Register it.",
    "What does the code egy.mit.co2.tot.tot.2 mean? Use the codes map.",
  ];
  p1.forEach((p, i) => {
    const col = i % 2, row = Math.floor(i / 2), x = M + col * 6.15, y = 1.65 + row * 1.6;
    card(pres, C, s, x, y, 5.95, 1.4, { name: "prompt" });
    iconCircle(pres, C, s, ic.chat, x + 0.2, y + 0.35, 0.6, col ? C.accent4 : C.accent1);
    s.addText("\u201C" + p + "\u201D", { x: x + 1.0, y: y + 0.1, w: 4.8, h: 1.2, fontSize: 14, italic: true, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
  });

  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Example prompts" });
  head(s, "EXAMPLE PROMPTS", "Ask it to change something, then check");
  const p2 = [
    ["Change", "Add a chart of revenue by fuel to the Macro-Fiscal Diagnostic, as the Dashboard's. Follow NORMS.md. Do not change anything else."],
    ["Fix", "The household chart is empty for Morocco. Find out why, fix it, run the quick check and log it in CAVEATS.md."],
    ["Python", "Make Coded CPAT replicate the new excise input, then run the tests and compare with the Excel workbook."],
    ["Release", "Build the next release of Mitigation and the unified CPAT, run the full checks, and tell me what changed."],
    ["Save", "Pull, then commit and push to main with a clear message."],
  ];
  p2.forEach(([t, p], i) => {
    const y = 1.6 + i * 1.0;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: M, y, w: 1.6, h: 0.8, rectRadius: 0.08, fill: { color: C.accent1 }, line: { color: C.accent1 } });
    s.addText(t, { x: M, y, w: 1.6, h: 0.8, align: "center", valign: "middle", fontSize: 16, bold: true, color: C.background1, margin: 0, isTextBox: true });
    card(pres, C, s, M + 1.8, y, W - 2 * M - 1.8, 0.8, { name: "prompt " + t });
    s.addText("\u201C" + p + "\u201D", { x: M + 2.0, y, w: W - 2 * M - 2.2, h: 0.8, fontSize: 14, italic: true, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
  });
  s.addText("Tips: one task per prompt; say what not to change; ask for the check to be run; read the diff before you keep it.",
    { x: M, y: 6.55, w: W - 2 * M, h: 0.4, fontSize: 13, italic: true, color: C.accent6, margin: 0, isTextBox: true });

  // 13 summary
  s = pres.addSlide({ masterName: "DARK", sectionTitle: "Example prompts" });
  s.addText("IN SHORT", { placeholder: "kicker" });
  s.addText("Pull, work, push to main", { placeholder: "title" });
  s.addText("Everything else (rules, logs, codes) the AI already reads from the repository.", { placeholder: "body" });
  const sum = ["Open the _latest MT and Activate Unified", "Copilot: Current checkout, never a worktree", "Claude Code: merge its branch to main",
    "Bugs fixed, issues kept and tracked", "Watch your credits: /context (USD 19)"];
  sum.forEach((t, i) => {
    const y = 1.2 + i * 1.0;
    s.addShape(pres.shapes.OVAL, { x: 8.0, y, w: 0.55, h: 0.55, fill: { color: C.accent1 }, line: { color: C.accent1 } });
    s.addText(String(i + 1), { x: 8.0, y, w: 0.55, h: 0.55, align: "center", valign: "middle", fontSize: 16, bold: true, color: C.background1, margin: 0, isTextBox: true });
    s.addText(t, { x: 8.75, y: y - 0.1, w: 4.2, h: 0.75, fontSize: 16, color: C.background1, valign: "middle", margin: 0, isTextBox: true });
  });

  const out = path.join(OUTDIR, "01_CPAT-AI_UserGuide_latest.pptx");
  await pres.writeFile({ fileName: out });
  await applyTheme(out, THEME);
  return out;
}

// =====================================================================================
// DECK 2: pitch and guide to CPAT-AI Excel
// =====================================================================================
async function deckExcel() {
  const FOOT = "CPAT-AI Excel  ·  WB CPAT Team  ·  November 2026  ·  Screens: public Default data";
  const { pres, C } = makePres("CPAT-AI Excel: pitch and guide", FOOT);
  const ic = {
    eye: await icon(fa.FaEye, H.lt1), check: await icon(fa.FaCheckCircle, H.lt1), bolt: await icon(fa.FaBolt, H.lt1),
    copy: await icon(fa.FaCopy, H.lt1), layer: await icon(fa.FaLayerGroup, H.lt1), chart: await icon(fa.FaChartBar, H.lt1),
    db: await icon(fa.FaDatabase, H.lt1), code: await icon(fa.FaPython, H.lt1),
    flag: await icon(fa.FaFlag, H.lt1), globe: await icon(fa.FaGlobe, H.lt1), cog: await icon(fa.FaCog, H.lt1),
    bug: await icon(fa.FaBug, H.lt1), lock: await icon(fa.FaLock, H.lt1), book: await icon(fa.FaBook, H.lt1),
  };

  // 1 title
  pres.addSection({ title: "Pitch" });
  let s = pres.addSlide({ masterName: "DARK", sectionTitle: "Pitch" });
  s.addText("CPAT-AI EXCEL · RELEASE 2.0", { placeholder: "kicker" });
  s.addText("CPAT you can read, check and run at scale", { placeholder: "title" });
  s.addText("The World Bank-IMF Climate Policy Assessment Tool rebuilt as readable Excel, with a new interface: one Dashboard, a scenario page and three diagnostic views.", { placeholder: "body" });
  s.addImage({ path: IMG("dashboard.png"), x: 7.3, y: 1.0, w: 5.5, h: 5.5 * 1092 / 1577, altText: "The CPAT Dashboard", line: { color: H.accent2, width: 1 } });

  // 2 why
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Pitch" });
  head(s, "WHY CPAT-AI", "Climate-policy advice rests on CPAT's numbers");
  const why = [["Readable", "Every row named, with its unit and source; short formulas; one formula per row, dragged across the years.", ic.eye],
    ["Correct", "Every bug of CPAT 1.0 fixed, with a cell plan to fix CPAT 1.0 itself; every issue kept as it is and tracked.", ic.check],
    ["Fast", "Many countries and scenarios at once from the MT; a new fuel, year or scenario is a paste, not a rebuild.", ic.bolt]];
  why.forEach(([t, d, im], i) => {
    const x = M + i * 4.1;
    card(pres, C, s, x, 1.7, 3.9, 3.6, { name: t });
    iconCircle(pres, C, s, im, x + 0.3, 1.95, 0.8);
    s.addText(t, { x: x + 0.3, y: 2.95, w: 3.3, h: 0.5, fontSize: 24, bold: true, color: C.text2, margin: 0, isTextBox: true });
    s.addText(d, { x: x + 0.3, y: 3.5, w: 3.3, h: 1.7, fontSize: 15, color: C.text1, valign: "top", margin: 0, isTextBox: true });
  });
  s.addText([B("And checked twice: "), T("an independent Python model (Coded CPAT) reads the same data and replicates the workbook number for number.")],
    { x: M, y: 5.65, w: W - 2 * M, h: 0.5, fontSize: 15, color: C.text1, margin: 0, isTextBox: true });

  cardsSlide(pres, C, "Pitch", "WHAT IS NEW IN 2.0", "CPAT-AI 2.0: the whole model, a new interface, real data", [
    ["All modules in one", ["Mitigation 2.0.0 with Distribution 0.6.31, Air pollution 0.1.5 and Transport co-benefits 0.1.6", "Pigouvian charges with no circular references"], ic.layer, C.accent1],
    ["New interface", ["Dashboard, User-defined page, Comparison, Climate and Macro-Fiscal diagnostics", "Three names per scenario"], ic.chart, C.accent4],
    ["New instruments", ["ETS by sector, CBAM, carbon tax enhancement, output-based feebates", "Capacity targets, green spending"], ic.cog, C.accent5],
    ["Main edition", ["The full real dataset (CPAT 1.0pre_468b) is the default", "Settings scenarios DEF, USR, HPR, LPR"], ic.db, C.text2],
  ], [B("Checked: "), T("Coded CPAT 2.0.0p2.3 replicates every module of 2.0.0 on Main within 1e-9; 104 tests pass.")]);

  // 3 eight improvements
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Pitch" });
  head(s, "BETTER THAN CPAT 1.0", "Eight improvements");
  const imp = [["Human readable", "named rows, units, sources", ic.eye], ["Consistent", "one formula per row", ic.check],
    ["LAMBDA or not", "direct formulas or named functions", ic.layer], ["Copy-pastable", "across years, scenarios, fuels", ic.copy],
    ["Small and fast", "shared formulas, .xlsb working copies", ic.bolt], ["Many scenarios", "the MT runs countries x scenarios", ic.layer],
    ["New interface", "Dashboard, Comparison, two diagnostics", ic.chart], ["Data editions", "public Default, real Main, linked", ic.db]];
  imp.forEach(([t, d, im], i) => {
    const col = i % 4, row = Math.floor(i / 4), x = M + col * 3.07, y = 1.7 + row * 2.45;
    card(pres, C, s, x, y, 2.87, 2.2, { name: t });
    iconCircle(pres, C, s, im, x + 0.25, y + 0.25, 0.65, row ? C.accent4 : C.accent1);
    s.addText([{ text: t, options: { bold: true, fontSize: 17, color: C.text2, breakLine: true } }, { text: d, options: { fontSize: 13, color: C.text1 } }],
      { x: x + 0.25, y: y + 1.05, w: 2.45, h: 1.05, valign: "top", margin: 0, isTextBox: true });
  });

  picSlide(pres, C, "Pitch", "BETTER THAN CPAT 1.0 · READABLE", "The model sheet reads like a table",
    "mitigation_layout.png", 2740, 597, "The Mitigation tab", "Tab Mitigation: settings on the left, macro variables on top, years to the right, four scenario groups side by side", [
      [B("Settings on the left: "), T("each row's parameters (elasticities, base-year data) in named cells")],
      [B("Macro variables on top: "), T("year, scenario, deflators, GDP growth, read by every row below")],
      [B("Drag right: "), T("the base year is its own column; then one formula per row for every later year")],
      [B("Scenarios left to right: "), T("four groups, the same rows and formulas; every row has a code, a unit and a source")],
    ], true);
  picSlide(pres, C, "Pitch", "BETTER THAN CPAT 1.0 · FORMULAS", "One row, CPAT 1.0 against CPAT-AI",
    "formulas_cpat10_now.png", 2434, 678, "The same fuel-use row in CPAT 1.0 and CPAT-AI", "The same road fuel-use row: CPAT 1.0 above, CPAT-AI below", [
      [B("Short formulas: "), T("target 90 characters, at most 115; one formula per block")],
      [B("Direct or LAMBDA: "), T("direct cell formulas as CPAT 1.0, or named LAMBDA functions (FUELUSE, ...), chosen at activation")],
      [B("The same numbers either way, "), T("and the same in Python")],
    ], true);
  statsSlide(pres, C, "Pitch", "WHAT WE FOUND", "Rebuilding CPAT found hundreds of bugs and issues", [
    ["261", "legacy bugs found, every one fixed in CPAT-AI, each with a cell plan for CPAT 1.0"],
    ["183", "still open in CPAT 1.0 (468b): the fix plans are ready for the CPAT team"],
    ["10", "major issues: five of data (MId), five of method (MIm), for the team to decide first"],
    ["158", "other issues: replicated as CPAT 1.0 and tracked, never changed silently"],
    ["27", "prototype bugs, found mostly by the Python replication, all fixed"],
    ["15", "upgrades proposed (newer data, better methods), waiting for a decision"],
  ], "Counts from excel_ai/core/BUGS.md and ISSUES.md (register 1.9.60), generated at each release.");
  tableSlide(pres, C, "Pitch", "WHAT WE FOUND · MAJOR ISSUES", "The open questions to decide first", [
    ["Code", "Issue"],
    ["MId3", "Combustion emission factors use only the GAINS base and final factors; year-by-year paths are needed"],
    ["MId2", "Blank VAT cells in the IMF price dataset"],
    ["MId1", "Proprietary tables need complete country coverage"],
    ["MIm1", "Fuel use grows with total GDP"],
    ["MIm2", "No partial adjustment of fuel use"],
    ["MIm3", "Average prices mix price and composition changes"],
    ["MIm4", "The GDP impulse"],
    ["MIm5", "Pigouvian netting"],
  ], [1.5, 10.63], "All replicated as CPAT 1.0 and tracked in ISSUES.md; a method change is the CPAT team's decision.");

  // 4 the workbook's tabs
  pres.addSection({ title: "Guide" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Guide" });
  head(s, "GUIDE · THE WORKBOOK", "The tabs, in the order you use them");
  const tabs = [
    ["Dashboard", "pick the country, the four scenarios and the year of interest; read the results", C.accent1],
    ["SCENARIOS ›", "User-defined (your own policy package) · Pre-defined (the scenario table)", C.accent4],
    ["OUTPUTS ›", "Comparison · Climate · Fiscal (Macro-Fiscal Diagnostic) · Output Data (every output by code)", C.accent4],
    ["ADVANCED ›", "Cover · Settings (model settings scenarios DEF, USR, HPR, LPR) · Manual Inputs · Mappings", C.accent6],
    ["CALCS ›", "Mitigation · Distribution · Transport · AirPollution: the model itself, read like a table", C.accent6],
    ["DATA ›", "the data tables, loaded by Update data from the chosen edition", C.accent6],
  ];
  tabs.forEach(([t, d, col], i) => {
    const y = 1.6 + i * 0.85;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: M, y, w: 2.6, h: 0.65, rectRadius: 0.08, fill: { color: col }, line: { color: col } });
    s.addText(t, { x: M, y, w: 2.6, h: 0.65, align: "center", valign: "middle", fontSize: 15, bold: true, color: C.background1, margin: 0, isTextBox: true });
    s.addText(d, { x: M + 2.85, y, w: 9.4, h: 0.65, valign: "middle", fontSize: 15, color: C.text1, margin: 0, isTextBox: true });
  });
  s.addText("Colour key everywhere: lemon = input, grey = information from the model, struck through = not yet in the model.",
    { x: M, y: 6.75 - 0.2, w: W - 2 * M, h: 0.35, fontSize: 13, italic: true, color: C.accent6, margin: 0, isTextBox: true });

  // 5 get started (MT)
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Guide" });
  head(s, "GUIDE · GET STARTED", "From the MT to your first results, in four steps");
  const st = [
    ["Open the MT", "excel_ai/master/CPAT-AI-MT_latest.xlsm (Excel 365, macros on)."],
    ["Activate Unified", "Tab Settings and Macros. Your working copies go to excel_ai/active_workbooks/."],
    ["Update data", "Choose the edition and the countries, then Update data."],
    ["Run", "On the CPAT Dashboard: country and scenarios. For many countries x scenarios: the MT's Dashboard, Run scenarios; every run is stored as values."],
  ];
  st.forEach(([t, d], i) => {
    const x = M + i * 3.07;
    card(pres, C, s, x, 1.7, 2.87, 3.4, { name: t });
    s.addShape(pres.shapes.OVAL, { x: x + 0.25, y: 1.95, w: 0.7, h: 0.7, fill: { color: C.accent1 }, line: { color: C.accent1 } });
    s.addText(String(i + 1), { x: x + 0.25, y: 1.95, w: 0.7, h: 0.7, align: "center", valign: "middle", fontSize: 22, bold: true, color: C.background1, margin: 0, isTextBox: true });
    s.addText([{ text: t, options: { bold: true, fontSize: 18, color: C.text2, breakLine: true } }, { text: d, options: { fontSize: 14, color: C.text1 } }],
      { x: x + 0.25, y: 2.85, w: 2.45, h: 2.15, valign: "top", margin: 0, isTextBox: true });
  });
  s.addText(bullets([
    "The MT's other macros: Sync scenarios (the scenario table into the model), Estimate run time, Reduce to countries.",
    "Stored runs: tabs StoredResults, Scenario comparison, OneCountry and ManyCountries of the MT.",
  ]), { x: M, y: 5.4, w: W - 2 * M, h: 1.3, fontSize: 14, color: C.text1, valign: "top", margin: 0 });

  // 6 user-defined
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Guide" });
  head(s, "GUIDE · SCENARIOS", "Define a policy package on the User-defined page");
  shot(s, "user_defined.png", M, 1.6, 7.4, 1060, 600, "The User-defined page", C, "Tab User-defined: section A, direct carbon pricing and feebates");
  s.addText(bullets([
    [B("A "), T("Carbon tax, exemptions, ETS, feebates, methane fee, coverage by fuel and sector")],
    [B("B "), T("Fossil fuel subsidy reform")],
    [B("C, D "), T("Power and road transport (EV subsidy, distance tax, purchase taxes)")],
    [B("Revenue use "), T("PIT and CIT cuts, public investment, transfers, green spending")],
    [B("Charts beside each section "), T("read the model live: carbon price, coverage, energy prices by component")],
    [B("Three names "), T("per scenario: the picker name, the number, a short name for the charts")],
  ], { space: 8 }), { x: 8.3, y: 1.6, w: 4.45, h: 5.0, fontSize: 14, color: C.text1, valign: "top", margin: 0 });

  cardsSlide(pres, C, "Guide", "GUIDE · SETTINGS", "Settings scenarios: model settings apart from policies", [
    ["DEF Default", ["The CPAT settings as shipped"], ic.cog, C.accent1],
    ["USR User defined", ["Your own settings, edited on tab Settings"], ic.cog, C.accent4],
    ["HPR High prices", ["High international energy prices"], ic.cog, C.accent5],
    ["LPR Low prices", ["Low international energy prices"], ic.cog, C.accent6],
  ], [B("Each scenario names its settings code "), T("(row SettingsScenario), so the same policy can be run under high and low prices. Settings hold the values every scenario shares; policies live on the User-defined and Pre-defined pages.")]);
  tableSlide(pres, C, "Guide", "GUIDE · NEW IN 2.0", "New policy instruments", [
    ["Instrument", "What it does"],
    ["ETS by sector", "Coverage by sub-sector, a benchmark for power, the share auctioned"],
    ["CBAM", "The industries exposed to a border adjustment, 0% auctioned"],
    ["Carbon tax enhancement", "Scales the carbon tax by sub-sector"],
    ["Output-based feebate", "Feebates by sub-sector, rebated on output"],
    ["Capacity targets", "A path to a target capacity for each power technology"],
    ["Green spending", "Spending and investment from the revenue, with its own GDP impulse"],
  ], [3.4, 8.73], "All on the User-defined page; Coded CPAT 2.0.0p2.3 replicates each of them.");

  // 7 dashboard
  pres.addSection({ title: "The dashboards" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "The dashboards" });
  head(s, "DASHBOARDS · 1 OF 4", "The Dashboard: one country, four scenarios, one year");
  shot(s, "dashboard.png", M, 1.55, 7.2, 1577, 1092, "The CPAT Dashboard", C, "Charts left empty here have not been recalculated in the render");
  s.addText(bullets([
    [B("Pickers: "), T("country, scenarios 2-4 (scenario 1 is the baseline), Year of Interest for every chart and table")],
    [B("Current situation: "), T("six tiles (GHG, per capita, NDC gap, fuel subsidies, effective carbon rate, net energy imports)")],
    [B("Scenarios table: "), T("emissions, revenue, bottom decile, deaths averted, welfare")],
    [B("Over time: "), T("carbon price, energy price impacts, revenue, GHG vs Paris pledge, GDP growth")],
    [B("One policy scenario: "), T("carbon price by component, households by decile, revenue by fuel, deaths averted, net welfare")],
    [B("Below: "), T("energy prices and externalities by fuel and sector")],
  ], { space: 6 }), { x: 8.45, y: 1.55, w: 4.3, h: 5.3, fontSize: 13, color: C.text1, valign: "top", margin: 0 });

  // 8 prices and externalities
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "The dashboards" });
  head(s, "DASHBOARDS · 1 OF 4", "Energy prices against the efficient price");
  shot(s, "prices.png", M, 1.7, W - 2 * M, 1576, 237, "Energy prices and externalities by fuel and sector", C,
    "Bottom of the Dashboard: liquid fuels, coal, natural gas and electricity, each on its own unit axis");
  const pr = [["Left: energy prices by fuel and sector", "Supply cost, existing taxes and VAT, new policies, for the chosen scenario. Black diamond: the end-user price."],
    ["Right: externalities by fuel and sector", "Climate, air pollution and road costs. Red diamond: the efficient price, all externalities included."],
    ["Read the gap", "Where the black diamond sits below the red one, energy is under-priced: the case for the policy."]];
  pr.forEach(([t, d], i) => {
    const x = M + i * 4.1;
    card(pres, C, s, x, 4.0, 3.9, 2.4, { name: t });
    s.addText([{ text: t, options: { bold: true, fontSize: 16, color: C.text2, breakLine: true } }, { text: d, options: { fontSize: 14, color: C.text1 } }],
      { x: x + 0.25, y: 4.15, w: 3.4, h: 2.15, valign: "top", margin: 0, isTextBox: true });
  });

  // 9 comparison
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "The dashboards" });
  head(s, "DASHBOARDS · 2 OF 4", "Comparison: the live scenarios side by side");
  shot(s, "comparison.png", M, 1.55, 7.2, 1577, 1091, "The Comparison tab", C, "Tab Comparison (OUTPUTS)");
  s.addText(bullets([
    [B("Key results table: "), T("each indicator as a level and as a change vs the baseline, for every scenario")],
    [B("Over time: "), T("carbon price, CO2 emissions, additional revenue, GDP effect, one line per scenario")],
    [B("Where the change comes from: "), T("carbon price by component, emission cuts by sector, revenue by fuel, GDP decomposition, one bar per scenario")],
    [B("Use it to "), T("choose between packages: which cuts most, raises most, costs least")],
  ], { space: 8 }), { x: 8.45, y: 1.55, w: 4.3, h: 5.3, fontSize: 14, color: C.text1, valign: "top", margin: 0 });

  // 10 climate
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "The dashboards" });
  head(s, "DASHBOARDS · 3 OF 4", "Climate Diagnostic: where the emissions come from");
  shot(s, "climate.png", M, 1.55, 7.6, 1576, 1029, "The Climate Diagnostic", C, "Tab Climate (OUTPUTS); the scenario is picked at the top");
  s.addText(bullets([
    [B("Headline tiles: "), T("GHG per capita (vs the G20 average), carbon intensity, NDC gap, coal and renewables in power, fossil share")],
    [B("Emissions and targets: "), T("GHG by gas, CO2 by sector, per capita vs G20 peers, path vs the NDC")],
    [B("Energy system: "), T("primary energy by fuel, the power mix, energy prices by fuel (dashed: baseline), household energy by decile")],
    [B("Use it for "), T("the country diagnostic before any policy")],
  ], { space: 8 }), { x: 8.45, y: 1.55, w: 4.3, h: 5.3, fontSize: 14, color: C.text1, valign: "top", margin: 0 });

  // 11 macro-fiscal
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "The dashboards" });
  head(s, "DASHBOARDS · 4 OF 4", "Macro-Fiscal Diagnostic: the fiscal starting point");
  shot(s, "fiscal.png", M, 1.5, 10.4, 1576, 583, "The Macro-Fiscal Diagnostic", C, null);
  const mf = [["Headline tiles", "Explicit energy subsidies, import dependence, energy revenue net of subsidies, GDP growth, effective carbon rate, energy trade balance."],
    ["Four charts", "GDP and population growth; energy tax revenue by fuel; energy revenue and subsidies over time; effective carbon rates by sector."],
    ["Use it for", "the fiscal case: how much revenue is at stake and where carbon is under-priced today."]];
  mf.forEach(([t, d], i) => {
    const x = M + i * 4.1;
    card(pres, C, s, x, 5.6, 3.9, 1.3, { name: t });
    s.addText([{ text: t, options: { bold: true, fontSize: 14, color: C.text2, breakLine: true } }, { text: d, options: { fontSize: 12, color: C.text1 } }],
      { x: x + 0.2, y: 5.65, w: 3.5, h: 1.2, valign: "top", margin: 0, isTextBox: true });
  });

  cardsSlide(pres, C, "The dashboards", "MANY COUNTRIES, MANY SCENARIOS", "The MT runs the batches and keeps the results", [
    ["Choose", ["Countries and scenario IDs (BASE, P1-P3, U1, ...) on the MT's Dashboard", "Reduce to countries keeps the files small"], ic.flag, C.accent1],
    ["Run", ["Run scenarios: batches over the live scenario groups", "Estimate run time first"], ic.bolt, C.accent4],
    ["Compare", ["Every run stored as values: StoredResults", "Scenario comparison, OneCountry, ManyCountries"], ic.chart, C.accent5],
  ]);
  pres.addSection({ title: "Close" });
  cardsSlide(pres, C, "Close", "DOCUMENTATION", "Everything is written down, for people and for the AI", [
    ["The website", ["cpmodelai.github.io/docs", "Methodology, the model block by block, Excel-AI, Python"], ic.globe, C.accent1],
    ["The bug book", ["Every legacy bug with its cell plan: cells, formula now, new formula", "Legacy_CPAT_Bugs_v*.docx"], ic.bug, C.accent5],
    ["The codes map", ["Every input, output and variable code, shared by Excel and Python", "CPAT-AI_CodesMap_latest.xlsx"], ic.book, C.accent4],
  ]);
  cardsSlide(pres, C, "Close", "NEXT STEPS", "Where CPAT-AI goes from 2.0", [
    ["Test in Excel 365", ["The team runs 2.0 on Main for its countries", "Feedback on the new interface"], ic.check, C.accent1],
    ["Decide the issues", ["The ten major issues first", "Then the upgrades"], ic.flag, C.accent5],
    ["Fix CPAT 1.0", ["Apply the cell plans of the open legacy bugs"], ic.bug, C.accent4],
    ["Grow the modules", ["Distribution with World Bank experts", "Air pollution and transport co-benefits"], ic.layer, C.text2],
  ]);

  // 12 where to find
  
  s = pres.addSlide({ masterName: "DARK", sectionTitle: "Close" });
  s.addText("WHERE TO FIND IT", { placeholder: "kicker" });
  s.addText("Open the MT, activate, and read the Dashboard", { placeholder: "title" });
  s.addText("Repository github.com/cpmodelai/cpat_ai · documentation cpmodelai.github.io/docs", { placeholder: "body" });
  const fin = [["The MT", "excel_ai/master/CPAT-AI-MT_latest.xlsm"], ["The unified CPAT", "excel_ai/unified/CPAT_latest.xlsx"],
    ["Bugs and issues", "excel_ai/core/BUGS.md, ISSUES.md"], ["Every code", "excel_ai/core/CPAT-AI_CodesMap_latest.xlsx"],
    ["Python twin", "coded/ (Shiny app: Run_CPAT_App.bat)"]];
  fin.forEach(([t, d], i) => {
    const y = 1.1 + i * 1.05;
    s.addText([{ text: t, options: { bold: true, fontSize: 16, color: C.accent2, breakLine: true } },
      { text: d, options: { fontSize: 14, color: C.background1, fontFace: "Courier New" } }], { x: 7.7, y, w: 5.1, h: 0.9, margin: 0, valign: "top", isTextBox: true });
  });

  const out = path.join(OUTDIR, "02_CPAT-AI_Excel_latest.pptx");
  await pres.writeFile({ fileName: out });
  await applyTheme(out, THEME);
  return out;
}

// =====================================================================================
// DECK 3: CPAT-AI Python (Coded CPAT)
// =====================================================================================
async function deckPython() {
  const FOOT = "CPAT-AI Python (Coded CPAT)  ·  WB CPAT Team  ·  November 2026  ·  Screens: public Default data";
  const { pres, C } = makePres("CPAT-AI Python: Coded CPAT", FOOT);
  const ic = {
    py: await icon(fa.FaPython, H.lt1), check: await icon(fa.FaCheckCircle, H.lt1), bolt: await icon(fa.FaBolt, H.lt1),
    eye: await icon(fa.FaEye, H.lt1), bug: await icon(fa.FaBug, H.lt1), db: await icon(fa.FaDatabase, H.lt1),
    term: await icon(fa.FaTerminal, H.lt1), laptop: await icon(fa.FaLaptopCode, H.lt1), globe: await icon(fa.FaGlobe, H.lt1),
    box: await icon(fa.FaBox, H.lt1), layer: await icon(fa.FaLayerGroup, H.lt1), flag: await icon(fa.FaFlag, H.lt1),
    cog: await icon(fa.FaCog, H.lt1), book: await icon(fa.FaBook, H.lt1), shield: await icon(fa.FaShieldAlt, H.lt1),
  };

  pres.addSection({ title: "Why" });
  let s = pres.addSlide({ masterName: "DARK", sectionTitle: "Why" });
  s.addText("CPAT-AI PYTHON · CODED CPAT 2.0.0p2.3", { placeholder: "kicker" });
  s.addText("The CPAT model in Python, number for number", { placeholder: "title" });
  s.addText("The twin of CPAT-AI Excel 2.0: the same data, inputs, codes and equations, written to be read, tested on every change, fast enough for many countries.", { placeholder: "body" });
  [["1e-9", "agreement with the Excel workbooks"], ["104", "tests passing"], ["~1 s", "one country, two scenarios"]].forEach(([v, l], i) => {
    const y = 1.3 + i * 1.7;
    s.addText(v, { x: 8.2, y, w: 4.5, h: 0.9, fontSize: 54, bold: true, color: C.accent2, margin: 0, isTextBox: true, fontFace: "Arial" });
    s.addText(l, { x: 8.2, y: y + 0.9, w: 4.5, h: 0.5, fontSize: 16, color: C.background1, margin: 0, isTextBox: true });
  });

  cardsSlide(pres, C, "Why", "IN ONE SLIDE", "One model, two forms, the same numbers", [
    ["Exact", ["Every output code of Excel-AI 2.0.0 for Egypt, both scenario groups, 2022-2041, within 1e-9"], ic.check, C.accent1],
    ["Fast", ["One country and two scenarios in about a second; reruns compute only what changed"], ic.bolt, C.accent4],
    ["Readable", ["One small function per equation, named with the Excel codes; a documentation book"], ic.eye, C.accent5],
    ["Honest", ["Bugs are never copied: the replication found and fixed them in both models"], ic.shield, C.text2],
  ]);

  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Why" });
  head(s, "WHAT IT IS", "Excel-AI is the specification; Python is its twin");
  [["CPAT-AI Excel", "the unified workbook: readable cells", C.accent1], ["Coded CPAT", "the same rows, codes and equations in Python", C.accent4], ["Ways in", "command line, Python, the Shiny app", C.text2]].forEach(([t, d, col], i) => {
    const x = M + i * 4.1;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.7, w: 3.6, h: 1.4, rectRadius: 0.08, fill: { color: col }, line: { color: col } });
    s.addText([{ text: t, options: { bold: true, fontSize: 20, color: C.background1, breakLine: true } }, { text: d, options: { fontSize: 13, color: C.background1 } }],
      { x: x + 0.2, y: 1.7, w: 3.2, h: 1.4, valign: "middle", margin: 0, isTextBox: true });
    if (i < 2) s.addShape(pres.shapes.LINE, { x: x + 3.65, y: 2.4, w: 0.4, h: 0, line: { color: H.accent6, width: 2, endArrowType: "triangle" } });
  });
  s.addText(bullets([
    [B("What it computes: "), T("energy prices, power, fuel use, emissions, revenues, GDP effects, welfare and efficient prices for 2022-2041; households by decile; air pollution deaths and welfare; transport co-benefits")],
    [B("How it stays equal: "), T("every Excel row has an output code (e.g. egy.mit.co2.tot.tot.2). The tests compare Python with the workbooks code by code and year by year; a block is done only when all its codes match.")],
    [B("Why have both: "), T("Excel for people to read and check; Python for speed, many countries, testing and apps. Each checks the other.")],
  ], { space: 10 }), { x: M, y: 3.5, w: W - 2 * M, h: 3.2, fontSize: 15, color: C.text1, valign: "top", margin: 0 });

  pres.addSection({ title: "What is new" });
  cardsSlide(pres, C, "What is new", "WHAT IS NEW IN 2.0", "Coded CPAT 2.0.0p2.3 replicates CPAT-AI 2.0", [
    ["Every module", ["Mitigation 2.0.0, Distribution 0.6.31, Air pollution 0.1.5, Transport co-benefits 0.1.6"], ic.layer, C.accent1],
    ["New instruments", ["Carbon tax enhancement, output-based feebates, capacity targets, green spending in GDP, ETS by sector, CBAM"], ic.cog, C.accent4],
    ["Main edition", ["Tests run on the real-data Main edition; settings scenarios by code"], ic.db, C.accent5],
    ["Better code", ["Several countries in one call, installable package, ruff and mypy on every push"], ic.py, C.text2],
  ], [B("Version naming: "), T("<Excel-AI version>p<Python version>: 2.0.0p2.3 is Python version 2.3 of the replication of Excel-AI 2.0.0.")]);

  statsSlide(pres, C, "What is new", "STATUS", "Every number checked against the workbooks", [
    ["104", "tests passing (11 skipped: old Public01 references); 23 data tests"],
    ["18,124", "Mitigation reference codes compared, Egypt, scenarios 1-2 (1.9.60)"],
    ["58,830", "cells of the AirPollution tab equal to the unified workbook"],
    ["1e-9", "tolerance: relative above 1, absolute below"],
    ["0", "known differences with Excel-AI"],
    ["5", "modules: mitigation, distribution, air pollution, transport co-benefits, settings runs"],
  ], "Source: coded/CHANGELOG.md and coded/README.md.");

  pres.addSection({ title: "How it works" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "How it works" });
  head(s, "HOW A RUN WORKS", "Baseline first, then each scenario");
  const flow = ["Scenario set", "Prepare", "Year loop", "Finish", "Modules", "Results"];
  const fd = ["inputs by name, baseline first", "all years: prices, policies, costs", "2022-2041: power, fuel use, emissions, GDP", "all-year rows (NDC)", "distribution, air pollution, transport", "by output code"];
  flow.forEach((t, i) => {
    const w = 1.85, x = M + i * 2.03;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.8, w, h: 1.6, rectRadius: 0.08, fill: { color: i === 2 ? C.accent1 : C.background2 }, line: { color: i === 2 ? C.accent1 : C.background2 } });
    s.addText([{ text: t, options: { bold: true, fontSize: 15, color: i === 2 ? C.background1 : C.text2, breakLine: true } }, { text: fd[i], options: { fontSize: 12, color: i === 2 ? C.background1 : C.text1 } }],
      { x: x + 0.12, y: 1.8, w: w - 0.24, h: 1.6, valign: "middle", margin: 0, isTextBox: true });
  });
  s.addText(bullets([
    [B("One order: "), T("cpat_model/mitigation/chain.py is the only place the order of the blocks is written, as the Mitigation sheet")],
    [B("Arrays inside: "), T("block rows stay numpy arrays; they become labelled arrays only when read from outside")],
    [B("Cached: "), T("the baseline, the data steps and every scenario already run, so a rerun computes only what changed")],
    [B("Modules after Mitigation: "), T("air pollution and transport co-benefits run after it and feed their rows back, as the unified workbook links them")],
  ], { space: 9 }), { x: M, y: 3.85, w: W - 2 * M, h: 2.9, fontSize: 15, color: C.text1, valign: "top", margin: 0 });

  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "How it works" });
  head(s, "SPEED", "From five seconds to under one");
  s.addChart(pres.charts.BAR, [{ name: "Seconds", labels: ["Before (two scenarios)", "Baseline cached", "Unchanged rerun"], values: [5.3, 0.8, 0.05] }], {
    x: M, y: 1.6, w: 7.2, h: 4.9, barDir: "bar", chartColors: [H.accent1], showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 14, dataLabelFormatCode: "0.0#",
    dataLabelColor: H.dk1, catAxisLabelColor: H.dk1, valAxisLabelColor: H.accent6, catAxisLabelFontSize: 14, valAxisLabelFontSize: 12,
    valGridLine: { color: "D5DDD8", size: 0.5 }, catGridLine: { style: "none" }, showLegend: false, showTitle: true, title: "Seconds, Egypt, two scenarios",
    titleFontSize: 14, titleColor: H.dk2, catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", titleFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", catAxisOrientation: "maxMin" });
  s.addText(bullets([
    [B("Target: "), T("one country and two scenarios in under a second. Met with the baseline and data cached.")],
    [B("Checked by a test: "), T("test_speed.py: a cached run equals a fresh one exactly, and is fast")],
    [B("Many countries: "), T("cpat-run --countries A B runs them together; each equals its run alone within 1e-9")],
  ], { space: 10 }), { x: 8.2, y: 1.8, w: 4.55, h: 4.7, fontSize: 15, color: C.text1, valign: "top", margin: 0 });

  cardsSlide(pres, C, "How it works", "THE DATA", "The same data as Excel, as tidy tables", [
    ["Tidy CSV", ["One table per file, keys first (country, year, codes), then the value", "Converted to Excel's layout and back, exactly"], ic.db, C.accent1],
    ["Editions", ["Main (real) and Default (public) in data/", "data/editions.csv lists them; CPAT_DATA_DIR points elsewhere"], ic.layer, C.accent4],
    ["The rule", ["Results from Main stay inside", "The deployed Shiny app and anything shared use Default"], ic.shield, C.accent5],
  ]);

  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "How it works" });
  head(s, "THE CODE", "Where everything is in coded/");
  const tree = [["cpat_model/core/", "index sets, the variable registry (every Excel row and its code), the run loop and cache"],
    ["cpat_model/mitigation/", "the blocks of the Mitigation sheet in order: prices, power, fuel use, road, emissions, GDP, welfare"],
    ["cpat_model/distribution/", "the Distribution sheet: households by decile"],
    ["cpat_model/air_pollution/", "concentrations, health burden, valuation, unit damages"],
    ["cpat_model/transport_cobenefits/", "accidents, congestion, road damage"],
    ["cpat_model/results/", "Results: by_code, distribution_by_code, tcb_by_code"],
    ["cpat_shiny/, app.py", "the Shiny app"],
    ["cpat_testing/", "the pytest suite and the Excel reference values"],
    ["cpat_documentation/", "spec and plan, the equation list, the app's texts"]];
  s.addTable(tree.map(([a, b], i) => [{ text: a, options: { fontFace: "Courier New", bold: true, fontSize: 13, color: H.dk2, fill: { color: i % 2 ? H.lt1 : H.lt2 } } },
    { text: b, options: { fontSize: 13, color: H.dk1, fill: { color: i % 2 ? H.lt1 : H.lt2 } } }]),
  { x: M, y: 1.6, w: W - 2 * M, colW: [4.2, 7.93], border: { type: "solid", pt: 0.5, color: "D5DDD8" }, valign: "middle", margin: [3, 6, 3, 6], rowH: 0.48 });

  cardsSlide(pres, C, "How it works", "READABLE FIRST", "One equation, one function, Excel's names", [
    ["The equation list", ["cpat_documentation/equations/: every equation (Equations.qmd), generated from equations.yaml", "Each with its function"], ic.book, C.accent1],
    ["Excel's codes", ["Variables carry the Excel-AI codes and input names (CPLevelTarget, ...)", "The codes map is shared"], ic.eye, C.accent4],
    ["Norms", ["coded/NORMS.md: no sums over sets, one name one meaning, PEP 8 at 120 characters", "ruff and mypy on every push"], ic.check, C.accent5],
  ]);

  pres.addSection({ title: "Testing" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Testing" });
  head(s, "TESTING", "Every block against the workbooks");
  s.addText(bullets([
    [B("Mitigation: "), T("every output code, Egypt, scenarios 1-2, 2022-2041; also the engineer road model and the new instruments switched on")],
    [B("Distribution: "), T("every code, basket and ten deciles")],
    [B("Air pollution: "), T("every cell of the AirPollution tab")],
    [B("Transport co-benefits: "), T("modes Legacy and Engineer, both scenario groups")],
    [B("Several countries, the package, speed, data: "), T("test_multi_country, test_package, test_speed, the data tests")],
  ], { space: 10 }), { x: M, y: 1.7, w: 7.3, h: 4.9, fontSize: 15, color: C.text1, valign: "top", margin: 0 });
  card(pres, C, s, 8.3, 1.7, 4.43, 4.6, { fill: C.text2, name: "run tests" });
  s.addText([{ text: "Run them", options: { bold: true, fontSize: 20, color: C.background1, breakLine: true } }, { text: " ", options: { fontSize: 8, breakLine: true } },
    { text: "cd coded", options: { fontFace: "Courier New", fontSize: 15, color: C.accent2, breakLine: true } },
    { text: "pytest", options: { fontFace: "Courier New", fontSize: 15, color: C.accent2, breakLine: true } }, { text: " ", options: { fontSize: 8, breakLine: true } },
    { text: "About a minute, before every commit. The references are re-extracted from the workbooks at each Excel release.", options: { fontSize: 14, color: C.background2 } }],
  { x: 8.6, y: 1.9, w: 3.9, h: 4.2, valign: "top", margin: 0, isTextBox: true });

  tableSlide(pres, C, "Testing", "BUGS FOUND BY WRITING IT TWICE", "The replication found bugs in both models", [
    ["Bug", "What", "Status"],
    ["M15", "VAT on efficient prices at the 2022 rate in every year", "Fixed in Excel and Python"],
    ["M16", "Freight BEV feebate at the ICEV truck's distance", "Fixed"],
    ["M17", "Plug-in hybrid cost mixed two price bases", "Fixed"],
    ["M19", "Nominal GDP divided by a billion twice", "Fixed"],
    ["M20", "LNG and European gas prices swapped", "Fixed"],
    ["M21", "Methane fee: division by zero", "Fixed"],
    ["M22", "Distribution: a % added to amounts in LCU", "Fixed"],
    ["CB1-CB3", "Python's own: several countries in one call; methane inventory year; a sum over a set", "Fixed (2.2)"],
  ], [1.4, 7.73, 3.0], "The rule is the same as Excel: bugs are never copied; issues are replicated and tracked.");

  pres.addSection({ title: "Using it" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Using it" });
  head(s, "USING IT · COMMAND LINE AND PYTHON", "A few lines to a result");
  const code1 = ["conda env create -f environment.yml", "conda activate cpat_p", "python run_model.py", "python run_model.py my_set.csv --scenarios 1 2 --xlsx", "cpat-run --countries EGY MAR --settings-runs", "pytest"];
  const code2 = ["import cpat_model", "from cpat_model.inputs import read_scenario_set", "results = cpat_model.run(read_scenario_set())", "results.by_code('EGY').loc['egy.mit.co2.tot.tot.2']", "results.distribution_by_code('EGY')", "results.tcb_by_code('EGY')"];
  [[code1, "Command line (from coded/)"], [code2, "From Python"]].forEach(([lines, t], i) => {
    const x = M + i * 6.15;
    s.addText(t, { x, y: 1.6, w: 5.95, h: 0.4, fontSize: 16, bold: true, color: C.text2, margin: 0, isTextBox: true });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 2.1, w: 5.95, h: 3.4, rectRadius: 0.06, fill: { color: C.text2 }, line: { color: C.text2 } });
    s.addText(lines.map((l, j) => ({ text: l, options: { breakLine: j < lines.length - 1 } })), { x: x + 0.25, y: 2.25, w: 5.5, h: 3.1, fontFace: "Courier New", fontSize: 12, color: C.background1, valign: "top", margin: 0, paraSpaceAfter: 6, isTextBox: true });
  });
  s.addText("Results come by output code and year: the same codes as the Excel workbook and the codes map.", { x: M, y: 5.8, w: W - 2 * M, h: 0.5, fontSize: 14, italic: true, color: C.accent6, margin: 0, isTextBox: true });

  picSlide(pres, C, "Using it", "USING IT · THE APP", "The Shiny app: the User-defined page, generated from Excel",
    "app_userdefined.png", 1400, 1000, "The User-defined scenario tab of the Shiny app", "Tab User-defined scenario (public data)", [
      [B("Start: "), T("double-click coded/Run_CPAT_App.bat, or shiny run app.py")],
      [B("Same page as Excel: "), T("its controls are generated from the Excel page specification, so the two never drift")],
      [B("Then: "), T("add the scenario to the set, run on Scenarios and run")],
      [B("Tabs: "), T("User-defined scenario, Scenarios and run (with settings runs), Data, Dashboard, Results")],
    ]);
  picSlide(pres, C, "Using it", "USING IT · THE APP", "The app's Dashboard and Results",
    "app_dashboard.png", 1400, 1100, "The Dashboard tab of the Shiny app", "Tab Dashboard (public data)", [
      [B("Key results: "), T("carbon price, effective carbon price, CO2, revenue, GDP, road deaths averted")],
      [B("Charts: "), T("carbon price, CO2, revenue, GDP, energy prices, households by decile, transport co-benefits")],
      [B("Results: "), T("every output code, downloadable as CSV or a workbook")],
      [B("Several countries "), T("in one run")],
    ]);
  cardsSlide(pres, C, "Using it", "USING IT · AS A PACKAGE", "Install it anywhere, bring your own data", [
    ["Install", ["pip install from GitHub (subdirectory coded); cpat[app] for the app", "Needs read access to the repository"], ic.box, C.accent1],
    ["Point at data", ["Set CPAT_DATA_DIR to a folder laid out as data/", "The package holds no data"], ic.db, C.accent4],
    ["Run", ["cpat-run --scenarios 1 2 --out results/", "Version <Excel>.<Python>, e.g. 2.0.0.2.3"], ic.term, C.accent5],
  ]);

  pres.addSection({ title: "Close" });
  cardsSlide(pres, C, "Close", "DEVELOPING IT", "How Python follows Excel", [
    ["Excel releases", ["A new Excel-AI version is built and checked"], ic.layer, C.accent1],
    ["Re-extract", ["The references are re-extracted from the workbooks"], ic.db, C.accent4],
    ["Replicate", ["The AI changes the blocks until every code matches"], ic.py, C.accent5],
    ["Log", ["CHANGELOG.md, TODO.md, a git tag coded-v<version>"], ic.book, C.text2],
  ], [B("Ask: "), T("\u201CMake Coded CPAT replicate Excel-AI 2.0.1, run the tests and update the changelog.\u201D The rules are in coded/NORMS.md; the open work in coded/TODO.md.")]);
  s = pres.addSlide({ masterName: "DARK", sectionTitle: "Close" });
  s.addText("NEXT STEPS", { placeholder: "kicker" });
  s.addText("Every country, every module, in one call", { placeholder: "title" });
  s.addText("Code: coded/ · documentation: the website's CPAT Python tab (cpmodelai.github.io/docs) · app: coded/Run_CPAT_App.bat", { placeholder: "body" });
  ["Every country on the Main edition", "The open code-quality items (TODO D14-D16, E16)", "The Pigouvian stored block and road rows of the unified workbook", "Deploy the app on public data", "Follow each Excel release the same day"].forEach((t, i) => {
    const y = 1.2 + i * 1.0;
    s.addShape(pres.shapes.OVAL, { x: 8.0, y, w: 0.55, h: 0.55, fill: { color: C.accent1 }, line: { color: C.accent1 } });
    s.addText(String(i + 1), { x: 8.0, y, w: 0.55, h: 0.55, align: "center", valign: "middle", fontSize: 16, bold: true, color: C.background1, margin: 0, isTextBox: true });
    s.addText(t, { x: 8.75, y: y - 0.1, w: 4.2, h: 0.75, fontSize: 15, color: C.background1, valign: "middle", margin: 0, isTextBox: true });
  });

  const out = path.join(OUTDIR, "03_CPAT-AI_Python_latest.pptx");
  await pres.writeFile({ fileName: out });
  await applyTheme(out, THEME);
  return out;
}

(async () => {
  console.log(await deckGuide());
  console.log(await deckExcel());
  console.log(await deckPython());
})();
