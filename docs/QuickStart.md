<!-- Source of ../QuickStart.docx. Regenerate from the repository root: cp QuickStart.docx /tmp/qs_ref.docx && pandoc docs/QuickStart.md -o QuickStart.docx --reference-doc=/tmp/qs_ref.docx -->

---
title: "CPAT-AI: Quick Start for People"
---

A short, human-readable orientation to the cpat_ai repository, mostly about the Excel prototypes (Excel-AI). Written 2026-10-29, updated 2026-11-02 for release 1.5 (versions since: Mitigation, the unified CPAT and the MT, the Master, 1.9.11, Distribution 0.6.5; the unified CPAT built with every release and activated by default; 1.9.11: the Dashboard's five charts a row, CopyProprietaryData in the MT): Mitigation and Master 1.5.1 (the Master runs many countries and scenarios from its Dashboard and keeps the results; the mitigation workbook opens on its Dashboard), Distribution 0.5.1 (its own 0.x version line), data edition Public01. The detail lives in the documents listed in section 1; this guide only points to them.

# 1. Key documents (start here)

Paths are relative to the repository root. Files named `_v<version>` carry the current version (today 1.9.11 for Mitigation, the unified CPAT and the MT, 0.6.5 for Distribution; the bug book and the tracker carry the version of the Mitigation release they were made with); older versions are kept with their module (`releases/` folders).

**The website** <https://cpmodelai.github.io/cpat_ai/> has a home page and four tabs: Methodology (`#methodology`, the CPAT methodology), CPAT-AI Model (`#model`, the method of every block, shared by Excel-AI and Python), CPAT Excel-AI (`#excel-ai`, the Excel prototypes) and CPAT Python (`#python`, the Python model). Offline: `docs/index.html`.

## Human-readable documents

| **Document** | **What it is for** |
|---|---|
| `QuickStart.docx` | This guide (root folder). |
| `excel_ai/master/CPAT-AI-MT_latest.xlsm` | The MT, Multiscenario Tool (the Master workbook; 1.9.5, was `CPAT-AI-Master_UserGuide_latest.xlsm`; Activate all saves it into `excel_ai/active_workbooks/` as `CPAT-AI-MT.xlsb`, your working copy): every macro. Tab Dashboard: Run scenarios, the countries and scenarios to run, the status, the results at a glance; tab Activation and Macros: Activate all with its settings (Live scenarios, Last year, Formulas), Update data with its settings, Real data (proprietary): the copy of the real data tables from the legacy CPAT into the activated workbook (1.9.11; log on tab Real data log), and the other macros, each explained. |
| `excel_ai/mitigation/docs/Documentation.md` | The guide to the mitigation module: what it is, quick start through the Master, how it works section by section, results, limits (also chapter 2 of the website's Excel-AI tab; Word copy `CPAT-AI-Mitigation_Documentation.docx`). |
| `excel_ai/mitigation/docs/CPAT-AI-Mitigation_Methods.docx` | Methods volume: one chapter per block, with equations. |
| `excel_ai/bugs_and_legacy/Tracking_Legacy_Bugs_v<version>.xlsx` | **The master list of bugs and issues**, every module, with the CPAT team's own tabs. |
| `excel_ai/bugs_and_legacy/Legacy_CPAT_Bugs_v<version>.docx` | Every legacy bug found, with its cell-level fix plan for legacy CPAT. |
| `excel_ai/core/CPAT-AI_CodesMap_v<version>.xlsx` | Codes, variables and every input (legacy row, name, Python key). |
| `other/SetupGuide.docx` | Getting set up: GitHub, the repos, the AI tools, git in brief. |

## Key PowerPoints

| **Deck** | **What it is for** |
|---|---|
| `docs/presentations/06_CPAT-AI_Consolidated_latest.pptx` | **Start here.** The consolidated CPAT-AI deck: terms, what it is and its status, what it does better than CPAT 1.0, bugs and issues found, the rules, the data, the documentation, how to use, develop and collaborate. |
| `docs/presentations/earlier_versions/01_CPAT-AI_Overview_latest.pptx` | Earlier (1.8.1) overview deck, replaced by the consolidated deck: status, how to use the prototype, results. |
| `docs/presentations/earlier_versions/04_CPAT-AI_Pitch_latest.pptx` | Earlier (1.8.1) short pitch, replaced by the consolidated deck. |
| `docs/presentations/earlier_versions/05_CPAT-AI_Setup_latest.pptx` | Earlier (1.8.1) setup deck, replaced by the consolidated deck. |
| `coded/cpat_documentation/Coded-CPAT_Overview_v<version>.pptx` | The Coded CPAT (Python) deck. |
| `other/presentations/AIClimate-FiscalTools_v2.6.pptx` | The overarching deck on the team's AI climate-fiscal tools, CPAT-AI among them. |

## Computer-readable files

Plain-text (Markdown) files that the AI and the scripts read; GitHub shows them formatted, so people can read them too. The Excel-AI rules and records are in `excel_ai/core/`.

| **File** | **What it is for** |
|---|---|
| `README.md`, `excel_ai/README.md` | Front page of the repo and of the Excel prototypes: who reads what, folder map. Each module folder has its own README. |
| `AGENTS.md` (and `CLAUDE.md`) | The entry point for any AI working in the repo. |
| `excel_ai/core/NORMS.md` | The rules for every task (people and AI alike). |
| `excel_ai/core/CAVEATS.md` | Log of every finished task and its caveats (newest at the bottom). |
| `excel_ai/core/CHANGELOG.md` | One line per release of the prototype. |
| `excel_ai/core/BUGS.md`, `excel_ai/core/ISSUES.md` | Short lists of bugs and issues (generated from the register). |
| `excel_ai/core/TODO.md` | The single list of open work; one file per open plan in `excel_ai/core/plans/`. |
| `data/README.md` | The data editions, every table, the scripts and the tests. |
| `coded/README.md` | The Python model (Coded CPAT): layout, running, tests. |

# 2. What the Excel prototype is, and the plan

The World Bank/IMF **Climate Policy Assessment Tool (CPAT)** models the fiscal, economic, energy, emissions and distributional effects of carbon pricing and related policies. The legacy CPAT (version CPAT 1.0pre_465) is a very large single workbook; the repository holds it as a shell, `data/legacy/CPAT 1.0pre_465_Shell.xlsb`, with its six proprietary tabs empty (section 4).

**CPAT-AI** rebuilds it, with AI help, as readable modular Excel workbooks (Excel-AI). The mitigation prototype answers: what do carbon taxes, ETSs, feebates, excises, subsidy reform and a methane fee do to prices, energy use, emissions, revenue and GDP, year by year to 2041, for one country at a time?

## The plan

1. **Replicate CPAT-legacy in Excel (making the ai-generated new ‘Excel Prototype’).** Match legacy results, but without its bugs (every bug is fixed). Issues are replicated as legacy has them, not yet upgraded: improvements are proposed and tracked, then decided later.
2. **Make Python replicate the Excel prototype** (`coded/`, the Coded CPAT and its Shiny app), using the same codes and input names, so the two can be compared number for number. Coded CPAT replicates Excel-AI Mitigation 1.5.0 and Distribution 0.5.1 on Public01 (its version and tests: `coded/README.md`).

## Next steps

- A user interface (specification in `excel_ai/bugs_and_legacy/ui_spec/`: a mock-up workbook and the core graphs of country reports).
- Model settings split from policy inputs.
- The other modules: distribution (started), air pollution, transport co-benefits.
- The full set of core graphs in the Master, which replaces the legacy Multiscenario Tool (kept in `excel_ai/bugs_and_legacy/legacy_multiscenario_tool/`).
- The team's test of the Master's runs in Excel (many countries, ETS scenarios, Achieve NDCs).

## Key improvements over legacy

- Every known legacy bug fixed, and documented with a fix plan the CPAT team can apply to legacy.
- Readable: formulas are short (target 90 characters, at most 115) and one formula serves a whole block; Mitigation's sections are numbered in sheet order, the exports to other modules marked by blue bands.
- Any country: formulas contain no country numbers; all country data come from the data tabs, for the country chosen on the mitigation workbook's Dashboard.
- Linked data: the data tables refresh from a data workbook (Power Query) instead of being pasted in.
- Scenarios: one scenario table with IDs (BASE, P1-P3, U1, ...), kept in the Master and synced into the module; the Master runs every country listed with every scenario listed, in batches over four live scenario groups, and keeps every run as values (StoredResults, Scenario comparison, OneCountry, ManyCountries); every output has a code shared with Python (e.g. `egy.mit.co2.tot.tot.2`).
- Fast: the modules ship with their values and formulas to 2035; repeated formulas are stored once (shared formulas); the Master saves the working copy as a binary `.xlsb`; calculation is automatic.
- Modular: one workbook per module (Mitigation, Distribution, Air Pollution; Transport Co-benefits to follow), without macros; one Master workbook holds every macro.
- Checked: an independent Python model (Coded CPAT), a quick check after every change, and a full check at milestones.

# 3. The modules and who is responsible

| **Module** | **Workbook (activated by the Master into `excel_ai/active_workbooks/`)** | **Responsible** |
|---|---|---|
| MT, the Master (1.9.11) | `CPAT-AI-MT_latest.xlsm` in `excel_ai/master/`, activated into `excel_ai/active_workbooks/CPAT-AI-MT.xlsb`: every macro, the scenario table and the stored results; sources in `excel_ai/master/` | Stephen |
| Mitigation (1.9.11) | `CPAT-AI-Mitigation_latest.xlsx` in `excel_ai/mitigation/`, activated as `CPAT-AI-Mitigation_latest-<n>s.xlsb` (n the live scenarios, 1.9.11) | Stephen (others joining soon) |
| Distribution (0.6.0) | `CPAT-AI-Distribution_latest.xlsx` in `excel_ai/distribution/`, activated as `..._latest.xlsb` | Faustyna (Victor joining soon) |
| Unified CPAT (1.9.11) | `CPAT-AI-Unified_latest.xlsx` in `excel_ai/unified/` (Mitigation with Distribution inside, read live; built with every full release), activated by default as `CPAT-AI-Unified_latest-<n>s.xlsb` (2s, 3s or 4s by Live scenarios; the MT works on the last one activated); the Master's **Build unified CPAT** (merge of the separate modules) is the fallback | Stephen |
| Air Pollution (AP) (scaffold v0.3, a separate track) | `CPAT-AI-AirPollution_latest.xlsx` in `excel_ai/air_pollution/` | Paulina |
| Transport Co-benefits | Planned (`excel_ai/core/plans/Transport_cobenefits.md`) | Alexandra |

## The active workbooks

Open **`excel_ai/master/CPAT-AI-MT_latest.xlsm`** (the MT, Multiscenario Tool) first (Excel 365, macros on). It is the only workbook with macros, and it opens on its **Dashboard**. `excel_ai/active_workbooks/` is your side: git never tracks it, so updates never clash with your work there.

1. **Activate all** (tab *Activation and Macros*, with the settings Live scenarios, Last year and Formulas). Opened from `excel_ai/master/`, the MT first saves itself into `excel_ai/active_workbooks/` as `CPAT-AI-MT.xlsb` (an earlier one goes to `active_workbooks/old/`) and carries on there; from then on open that `.xlsb`. When a newer Master has come with an update, its Dashboard says "Newer Master available: <version>": open the new one in `excel_ai/master/` and Activate all again. The Master then opens each module's `.xlsx` (shipped with formulas to 2035, so it is small and quick to open), fills the formulas across to 2041, adds the live scenario groups (two by default; the module's Dashboard shows scenarios 3-4 when they are live, else in a collapsed section under its pickers), writes its scenario table into the module, recalculates and saves `<Name>_latest.xlsb` next to itself in `active_workbooks/`. Work in the `.xlsb` (calculation is automatic).
2. **Choose the countries and the scenarios** on the Dashboard (one per row; scenario IDs BASE, P1-P3, U1, ... from the Master's tab Scenarios), (the data source and rows are on *Activation and Macros*, with Update data). Every country listed runs every scenario listed.
3. **Update data** (all countries, for a run of many), then **Run scenarios**: the runs go in batches over the live groups and are stored as values in StoredResults.
4. **Read the results.** Live: the mitigation workbook's **Dashboard** (country and policy scenario by name, key numbers and charts), Outputs and Diagnostic. Stored: the Master's StoredResults, Scenario comparison, OneCountry and ManyCountries.
5. **Households in the same file (1.9.7: the default).** With *Activate: Unified* (the default, tab *Activation and Macros*), Activate all activates the unified release `excel_ai/unified/CPAT-AI-Unified_latest.xlsx` (built with every release) as `CPAT-AI-Unified_latest-<n>s.xlsb` beside the Master (n the live scenarios): the mitigation workbook with Distribution inside, one Distribution group per live scenario, read live (no Distribution from Mitigation step); its Dashboard's household-decile chart (the stack of the chosen policy scenario) and tile and the Diagnostic's household energy chart show model values (the mitigation workbook alone leaves them empty: no dummy values). Run scenarios, Send and Get scenario, the stored results and Update data work on it. *Separate modules* activates the mitigation and distribution workbooks as before; **Build unified CPAT** then merges them (the fallback: it crashed one user's Excel, hence the release).

The other macros (send or get a scenario, sync the scenario table, store the live scenarios, estimate the run time, the ETS goal seek, Achieve NDCs, Distribution from Mitigation, save to .csv) are on the MT's tab *Activation and Macros*, each with what it does. The `.xlsb` are generated: on a git conflict take either side and activate again; never commit the `~$` lock files Excel writes while a workbook is open (`.gitignore` skips them). Run a module for a particular country there, on your local drive; keep copies holding real data local and never commit them. Every version is kept in the module's `releases/` folder (`excel_ai/mitigation/releases/`, ...) as `<Name>_v<version>`.

## How the modules connect and how data flow

- Mitigation is the hub. The blocks other modules read are named in their section bands (blue for the exports): section 11 (block 11a for Distribution, block 11b for Transport Co-benefits) and section 15 for Air Pollution.
- Distribution receives its block from the Master: **Update Distribution from Mitigation** copies block 11a of the activated mitigation workbook into the activated distribution workbook, by code. In the unified CPAT (1.9.1) the same block is read live by formula inside one file.
- Data tables: the Master's **Update data** loads them from the data workbook `data/Public01/legacy_xlsx/CPAT_LegacyData.xlsx` (or GitHub, or another data repository) into the module's data tabs by Power Query, for the module's country or all countries, as chosen on the MT's tab Activation and Macros (the module's Cover can override).
- Results: Outputs in the mitigation workbook shows the live scenarios side by side, by topic in Mitigation's row order; the Master stores every run as values (StoredResults) and compares them in tables and charts (Scenario comparison, OneCountry, ManyCountries).

# 4. Data and information security

- **All data live in `data/`, organised in editions.** An edition is a folder with the same structure as every other; `data/editions.csv` is the registry and `data/catalog.csv` the catalogue of every table. The scripts that convert the data are in `data/scripts/`.
- **Public01** (in this public repository, the default edition) holds the public tables plus the **six proprietary tables**, Balances, EnergyCons, Prices_dom, IO_GTAP, HHSurvey and Tran_Scatter, as perturbed values (±3%, identities kept): not the real data. The real data are only in the private edition Full01, never in this repository.
- **Tidy CSV is the source of truth:** `data/Public01/tidy/`. The data workbook the prototype's data links read, `data/Public01/legacy_xlsx/CPAT_LegacyData.xlsx` (tables `L_<Tab>` with a FilterISO3 column), is written from it and never edited by hand.
- **The legacy workbook:** `data/legacy/CPAT 1.0pre_465_Shell.xlsb` has the six proprietary tabs empty; `data/Public01/legacy/CPAT 1.0pre_465_Public01.xlsb` is the shell filled with Public01's data. Its cached formula results are still from the real data until it is recalculated and saved in Excel.
- **Documents shared beyond the team** (the bugs document, the tracker, the codes map and the decks) must hold no proprietary data or values derived from it.
- **Never upload any form of CPAT containing real data to the repo** (a workbook, an export, a CSV or a screenshot). That breaks information security. Real data stay in the private data repository or on your local drive (see `data/README.md` for pointing the scripts at them); check the changed files before every commit.

## Loading data into the mitigation module

On the MT's tab **Activation and Macros**: choose the **data source** and the **rows** to load (the country, or all countries); then click **Update data**. The mitigation workbook's Cover can override the choice (source, all countries, a file or URL); "From the Master" there uses the Master's choice. The data sources and their addresses are listed on the MT's tab *Activation and Macros*.

| **Source** | **What it does** |
|---|---|
| Embedded (shipped data, no update) | Uses the data shipped inside the workbook (Public01); the default. |
| Public01 - GitHub (public repository) | Reads `data/Public01/legacy_xlsx/CPAT_LegacyData.xlsx` from the public repository. |
| Public01 - local file | Reads `CPAT_LegacyData.xlsx` in your clone's `data/Public01/legacy_xlsx/` (or a path you type). |
| Data repository (type its URL) | Reads a data workbook from a separate data repository; a workbook you then save with real data must not be committed. |
| SharePoint (type its URL) | Reads a data workbook from SharePoint. |

# 5. Bugs, issues and how they are tracked

**The rule (`excel_ai/core/NORMS.md` section 0): never replicate a legacy bug; always replicate a legacy issue, and track it.**

| | **Bug** | **Issue** |
|---|---|---|
| Meaning | Legacy does not do what it means to: a wrong row, a term applied twice, a sign error. | Legacy does what it means to, but the choice is open to question: old data, a simplification, a missing feature. |
| We | Fix it. Never replicated. | Replicate it exactly; depart only by explicit user decision. |
| Example | L1: services electricity demand grows with the industrial multiple after 2023. | MId3: combustion emission factors use only the GAINS base and final factors, as legacy. |

Codes: L legacy bug, M prototype bug, DOC documentation bug, MId major issue (data), MIm major issue (methodology), I issue (data problems are issues), LD departure from legacy, U upgrade. The ten major issues (MId1-5, MIm1-5) were I codes before; the register keeps the former codes. One register for every module (`excel_ai/bugs_and_legacy/register/`) generates the tracker, the bugs document, `BUGS.md` and `ISSUES.md` at every release.

## Where things stand

The current counts of every code (legacy bugs, prototype bugs, major issues, issues, departures, upgrades, documentation bugs) and the status of each legacy bug in legacy 465 are at the top of `excel_ai/core/BUGS.md` and `excel_ai/core/ISSUES.md`, and in chapter 1 of the website's CPAT Excel-AI tab. They are generated from the register at each release.

## Major open issues

- **MId3 (Extremely high):** combustion emission factors use only the GAINS base and final factors; the year-by-year paths are needed.
- **MId2 (Very high):** blank VAT cells in the IMF price dataset (legacy fix L170 is in the prototype).
- **MId1:** the proprietary tables are Egypt only, so other countries need data.
- **I1:** the prototype has not yet had a full run in Excel 365 by the team (it is built and checked in LibreOffice).
- **MIm1-MIm5:** the major questions of method, replicated as legacy: fuel use grows with total GDP (MIm1), no partial adjustment of fuel use (MIm2), average prices mix price and composition changes (MIm3), the GDP impulse (MIm4), Pigouvian netting (MIm5).
- CBAM process emissions, green public investment, industrial emissions intensity and the Air Pollution and Transport Co-benefits modules are upgrades U15-U19, outside the current scope.

To raise a bug or issue: add it to the tracker (or ask the AI to register it). It must appear in `BUGS.md` or `ISSUES.md` the same day.

# 6. Key norms, with examples

`excel_ai/core/NORMS.md` holds all of them. The ones to know first:

- **Bugs vs issues** (section 0): as above. When in doubt, treat it as an issue, replicate it and ask.
- **Never improve legacy silently:** newer data or a better method is an upgrade (U) for the team to decide, e.g. a new population source.
- **Versions** (section 5): never edit a released workbook; build the next version. Workbooks are `<Name>_v<version>` in the module's `releases/`, e.g. `CPAT-AI-Mitigation_v1.5.1.xlsx` (to 1.1.0: `CPAT-AI-Mitigation-Prototype_v1.1.0.xlsm`).
- **Latest and active workbooks** (sections 5a, 9.7): the latest `.xlsx` of each module sits in its folder; the latest Master, the MT, is `excel_ai/master/CPAT-AI-MT_latest.xlsm`; `excel_ai/active_workbooks/` (git-ignored) holds the activated MT `CPAT-AI-MT.xlsb` and the activated modules.
- **Inputs** (section 4): use legacy MTInputs names, e.g. CPLevelTarget for the carbon price in the target year, the same in Excel and Python.
- **Style** (sections 1-2): fixed column layout and CPAT colours, e.g. input cells light green, code cells light blue, calculation cells unfilled; formulas short (target 90 characters, at most 115).
- **Data** (section 8): all data live in `data/`, as editions; tidy is the source of truth; only public or perturbed data are committed.
- **Every finished task is logged** in `excel_ai/core/CAVEATS.md` (section 6); methodology and version notes are separate documents (section 7).

# 7. How the AI works here

Any AI tool (Claude, Codex, Copilot, ...) starts from the root `AGENTS.md` (Claude's `CLAUDE.md` just points to it). `AGENTS.md` routes the AI to each project's rules: for Excel-AI, `excel_ai/core/AGENTS.md`, which points to `NORMS.md`, `BUGS.md`, `ISSUES.md`, `TODO.md` and `CAVEATS.md` in `excel_ai/core/`; for Coded CPAT, `coded/README.md`, `coded/NORMS.md` and `coded/TODO.md`. So it follows the same rules as people and logs what it did. To change how the AI behaves, change the project's rules (`NORMS.md`), not the AI's own file.

## Queries you can ask

- "Which legacy bugs are still open in legacy 465, and which affect revenue?"
- "Explain how the retail price of diesel is built up, with the cells in the prototype."
- "Is this difference from legacy a bug or an issue? Register it."
- "What changed between 1.0.4 and 1.1.0?" (it reads `CHANGELOG.md` and `CAVEATS.md`)
- "Where does the Distribution module get its fuel price changes from?"
- "Run the quick check and tell me if anything broke."

# 8. The repos and git

The repos sit under the GitHub organisation **cpmodelai**:

| **Repo** | **Purpose** |
|---|---|
| cpat_ai (public) | This repo: the CPAT Excel prototypes, the Python model, the public data edition and documents. Public, so never put real data or internal material here. |
| climate-fiscal-ai-tools (public) | The wider set of AI climate-fiscal tools beyond CPAT (e.g. the excise diagnostic, the tax dashboard). Public, so never put CPAT data or internal material there. |
| repo-tools (public) | The helper scripts that make git easy: CloneRepos.bat, SyncRepo.bat, SyncAllRepos.bat, RepoDirectory.csv (in Repos.zip). |

The real data (edition Full01) are in a separate private data repository.

Getting set up and a short guide to git are in `other/SetupGuide.docx`. In short: double-click SyncAllRepos.bat to pull and push every repo; if it stops with a merge conflict, **GitHub Desktop** shows the clashing files and helps you resolve them.

# 9. Further information

- Folder map: `README.md` (root) and `excel_ai/README.md`; each module folder has its own README.
- Documentation: the website <https://cpmodelai.github.io/cpat_ai/> (sources in `docs/`).
- How-to guides: the setup and tools guide `other/SetupGuide.docx`; the quick start in the mitigation Documentation; the data editions and scripts in `data/README.md`; module guides and prompts in `excel_ai/core/prompts/`.
- Supporting material in `other/`: research papers, decks (`other/presentations/`), distribution notes and earlier prototypes.
- Plans: `excel_ai/core/plans/` (e.g. `Modules.md` for the modular design).
- Python model (Coded CPAT): `coded/README.md` (`run_model.py`, the Shiny app `app.py`, tests in `coded/cpat_testing/`), its deck and the Coded CPAT tab of the website.
- Personal working folders: `users/`.
