# OSCE Sustainable Course Guide

Student-facing course guide for Texas A&M sustainability-classified courses.
Built with static HTML/CSS/JS — no server, build step, or framework required.
`index.html` is the entire site; everything else exists to produce the CSV
data it reads.

## Quick reference: what to do and when

This section is written for whoever is doing the yearly/semester data
update — no coding background assumed. It assumes two things are already
true on your computer:

- You have this repository downloaded and open in a terminal, inside the
  project folder (if you're not sure, run `pwd` — it should end in
  `sustainable-courses-tamu`).
- You have `git` installed and are able to push to this repository (if
  `git push` asks you to log in and you don't know the credentials, stop
  and ask whoever set up the repo).

Every set of steps below ends the same way: you push to `main`, and
GitHub does the rest automatically (usually within 30 seconds) — turning
your `.xlsx` file into `.csv` files, rebuilding `data/manifest.json`, and
redeploying the live site. You never need to touch `index.html` or run
anything by hand to make new data show up.

### A. Mid-year: you get one updated spreadsheet

Two kinds of files show up in `data/`, and the routine is slightly
different for each one.

**A file named like `<Semester> Sorted.xlsx`** (a per-semester syllabus
link lookup file):
1. In Finder (or your file browser), drag the new file into the `data/`
   folder, replacing the old file with the same name.
2. In the terminal, run:
   ```bash
   git add "data/Fall 2025 Sorted.xlsx"   # use the actual filename
   git commit -m "Update Fall 2025 Sorted lookup"
   git push
   ```
3. Done. Nothing else to do.

**A file named like `... Syllabi Review.xlsx`** (the consolidated review
workbook):
1. If OSCE sent you a new version of the *same* file (new sheets/tabs
   filled in), drag it into `data/`, replacing the old one under the exact
   same filename.
2. If OSCE sent you a *new* file with a different name (e.g. it now covers
   a later range of semesters), delete the old one from `data/` instead of
   leaving both — two consolidated workbooks sitting in `data/` at once
   will double-count any semester they both cover.
3. In the terminal:
   ```bash
   git add data/
   git commit -m "Update Spring 2026 syllabi review"
   git push
   ```
4. On GitHub, click the **Actions** tab and open the run that just started.
   It prints every sheet name and row count it found — confirm the
   semester you updated shows a nonzero row count and there's no
   `WARNING:` line (a warning means two files defined the same semester —
   see step 2).

### B. Once a year, when OSCE starts a new review cycle

This is the same idea as above, done for a whole batch of new files at
once, plus yearly cleanup. Do these in order:

1. Open a terminal in the project folder and get the latest version of the
   site first:
   ```bash
   git pull
   ```
2. Drag this year's new files into `data/`: the new consolidated review
   workbook, and the new per-semester `*Sorted.xlsx` lookup files.
3. **Delete last year's `*Sorted.xlsx` files.** Once a semester's syllabus
   links are baked into its `.csv`, the `Sorted.xlsx` that produced them is
   no longer needed — keeping it around only adds clutter and file size.
   ```bash
   git rm "data/Fall 2024 Sorted.xlsx" "data/Spring 2025 Sorted.xlsx"   # use last year's actual filenames
   ```
4. **Move last year's syllabi review workbook to the archive folder** —
   don't delete it, `git mv` it so it's kept for reference but ignored by
   the pipeline:
   ```bash
   git mv "data/Old Syllabi Review.xlsx" "data/archive/Old Syllabi Review.xlsx"
   ```
   (Skip this step if OSCE simply added new tabs to the *same* consolidated
   workbook you already have — same filename, nothing to archive.)
5. **Leave every `.csv` file in `data/` exactly where it is.** Do not
   delete, rename, or move old `.csv` files — the site needs them to keep
   showing previous semesters, and it already sorts/groups them by
   semester automatically.
6. Stage, commit, and push everything:
   ```bash
   git add data/
   git commit -m "Add 2026-2027 syllabi review data"
   git push
   ```
7. On GitHub, open the **Actions** tab and confirm the run succeeded with
   no `WARNING:` lines. Then open the live site and spot-check that the
   new semester appears.

### C. Fixing a mistake after you've already pushed

If something looks wrong on the live site after a push (a missing course,
a bad syllabus link, a typo), you don't need to redo the whole process —
just fix the spreadsheet and push again following section A or B above. If
you'd like to double-check the fix *before* pushing, you can run the
conversion script yourself on your own computer:

1. **First time only** — create a Python virtual environment (an isolated
   place to install the one package the script needs, so it doesn't affect
   anything else on your computer) and install that package:
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate          # on Windows (Command Prompt): .venv\Scripts\activate
   pip install openpyxl
   ```
2. **Every time** you come back to test something, re-activate it (you'll
   see `(.venv)` appear at the start of your terminal prompt when it's
   active):
   ```bash
   source .venv/bin/activate          # on Windows: .venv\Scripts\activate
   ```
3. Run the same conversion the website's automation runs, using everything
   currently in `data/`:
   ```bash
   python3 scripts/xlsx_to_csv.py data/*.xlsx
   ```
   This prints one line per sheet found (with its row count) and rewrites
   the `.csv` files in `data/` — open one in a spreadsheet program to
   double check the fix looks right.
4. When you're done testing, leave the virtual environment:
   ```bash
   deactivate
   ```
5. Commit and push the result, same as before:
   ```bash
   git add data/
   git commit -m "Fix syllabus link for ENGL 393"
   git push
   ```

You can skip steps 1–4 entirely and just push straight from section A or
B — the GitHub Action runs this exact same script automatically. Testing
locally first is only useful if you want to see the result before it goes
live.

### What's already automated (no manual step needed, ever)

- Converting every `.xlsx` in `data/` into per-semester `.csv` files
- Rebuilding `data/manifest.json`
- Filling in syllabus links by cross-referencing lookup workbooks
- Warning you (in the Actions log) if two workbooks conflict on the same
  semester
- Deploying the live site — GitHub Pages redeploys automatically whenever
  `main` changes

## Repo Structure

```
index.html                     ← the entire site (markup + CSS + JS in one file)
data/
  *.xlsx                       ← active source spreadsheets (see "Input files" below)
  archive/*.xlsx                ← retired/backup review workbooks — kept for reference, NOT read by the pipeline
  *.csv                        ← auto-generated, one per semester — DO NOT hand-edit
  manifest.json                ← auto-generated list of CSVs for the site to load — DO NOT hand-edit
logos/
  *.png                        ← OSCE / TAMU branding assets referenced by index.html
scripts/
  xlsx_to_csv.py                ← converts data/*.xlsx → data/*.csv (see below)
.github/workflows/
  build-data.yml                ← runs xlsx_to_csv.py + rebuilds manifest.json on every push
README.md                      ← this file
```

Nothing outside `data/` needs to change for a routine semester update. New
top-level folders (e.g. a future `docs/` or `tests/`) should get a one-line
entry added to the tree above and, if they hold anything auto-generated,
a note in this README about what generates them and what not to hand-edit.

## How the site works, end to end

1. **Input**: one or more `.xlsx` files live in `data/`. There are two kinds
   the pipeline understands (see "Input file layouts" below) — everything
   else in `data/` is ignored.
2. **Conversion**: `scripts/xlsx_to_csv.py` reads every `.xlsx` in `data/`,
   turns each semester/level sheet in the review-layout workbook(s) into rows
   of a per-semester CSV (`data/<Semester>.csv`), and — if any workbook is in
   the "raw" lookup layout — uses it to fill in syllabus links for review rows
   that don't have their own.
3. **Manifest**: a list of the resulting CSV filenames is written to
   `data/manifest.json` so the static site knows what to fetch without
   needing a directory listing (static hosting can't list files).
4. **Automation**: `.github/workflows/build-data.yml` runs steps 2–3
   automatically on every push to `main` that touches `data/**.xlsx`, then
   commits the generated CSVs + manifest back to `main` as `github-actions[bot]`.
5. **Serving**: `index.html` fetches `data/manifest.json`, then fetches and
   parses each listed CSV client-side (hand-rolled parser, no dependencies),
   and renders everything in the browser. GitHub Pages redeploys automatically
   whenever `main` changes, so the bot's commit in step 4 triggers a redeploy
   with the new data ~30 seconds later.

Nothing in this chain requires a manual edit to `index.html` when data
changes — only when you want to change how the site looks or behaves.

## Adding a new semester

For the actual step-by-step commands, see "Quick reference: what to do and
when" at the top of this file. This section just explains the *format* of
the annual review workbook, for context.

The annual syllabi review is delivered as a single **consolidated workbook**
covering multiple semesters at once — one sheet per semester/level, e.g.
`"Summer 2025 - Undergraduate"`, `"Fall 2025 - Graduate"`, `"Spring 2026 - Undergraduate"`,
all inside one file like `data/Summer 2025-Spring 2026 Syllabi Review.xlsx`. As new
semesters are reviewed, that same file gets updated in place with the new sheets
filled in (or reissued as a new file covering a later range) — you don't need to
touch the pipeline for that.

## Input file layouts

The converter recognizes two distinct spreadsheet shapes. Anything that
matches neither is silently skipped (safe to keep other reference
spreadsheets in `data/`).

**1. Review layout** — the standard STARS-style syllabi review. One sheet per
semester/level, sheet name formatted `"<Semester> - <Level>"` (e.g.
`"Fall 2026 - Graduate"`). Header row must include: `Course Title`,
`Department(s)`, `Course`, `Level`, `Course Description`, `Type`
(classification: Focused/Inclusive), and the 17 `Goal N. …` SDG columns
(`yes` in a cell = that SDG applies). An optional column whose header
contains both "syllabus" and "url"/"link" is picked up automatically as a
syllabus link. Cross-listed departments are expanded into one row per
department, and both cell conventions are supported: `Department(s)` may hold
short code(s) (`AFST/ENGL`) with `Course` holding just the number (`393`), or
`Department(s)` may hold full name(s) (cross-listed ones joined by `x`/`X`,
e.g. `Africana Studies x English`) with `Course` holding the code(s) and
number together (`AFST/ENGL 393`) — the converter detects which layout a row
uses and fills `department_full` when full names are available.

**2. Raw lookup layout** — a CRN/keyword-scan export whose header row
includes `SUBJECT`, `COURSE`, and `URL` columns (department, course number,
and a syllabus link — extra columns like keyword-match counts are ignored).
This isn't converted into a CSV of its own; instead it's scanned to build a
`(subject, course number) → url` lookup, which fills in `syllabus_url` for
review rows that don't already have one. This is how syllabus links get
attached even when the review spreadsheet itself has no URL column — keep a
file like this in `data/` alongside the review workbook if you have one.

## Primary vs. backup source files

`scripts/xlsx_to_csv.py` derives each output CSV's semester purely from sheet
*names* (`"<Semester> - <Level>"`), not from which file they live in — so a
single workbook can (and now does) cover many semesters as separate tabs, and
that's the primary format going forward: one consolidated review workbook
(e.g. `data/Summer 2025-Spring 2026 Syllabi Review.xlsx`) with a sheet per
semester/level. Per-semester `*Sorted.xlsx` raw-lookup files are unaffected by
this — keep supplying one per semester as before.

Because the converter runs against every `.xlsx` directly in `data/`
(`data/*.xlsx`, non-recursive) and combines rows across files by matching
sheet name, having an **old-format, one-workbook-per-semester** review file
loose in `data/` at the same time as the consolidated workbook — both with a
sheet like `"Spring 2026 - Undergraduate"` — would duplicate/conflict rows for
that semester. Retired review workbooks must go in `data/archive/` instead:
`data/*.xlsx` doesn't match paths inside subfolders, so anything archived
there is automatically excluded from every pipeline run (local or CI) while
staying committed to the repo as backup/reference. Use `git mv` so history is
preserved.

If a collision like this is ever accidentally introduced anyway (e.g. an old
workbook left unarchived), the script prints a `WARNING:` line to the console
identifying both files and the affected semester/level instead of silently
duplicating rows — check the Action log or local console output if row
counts on the site look doubled.

## Regenerating CSVs manually

See "Fixing a mistake after you've already pushed" in the Quick reference
section above for how to set up the Python virtual environment first. Once
it's activated, these are the other ways to run the script beyond the
basic `python3 scripts/xlsx_to_csv.py data/*.xlsx`:

```bash
# Convert only the sheets in one workbook (writes data/<Semester>.csv for each sheet found)
python3 scripts/xlsx_to_csv.py "data/Summer 2025-Spring 2026 Syllabi Review.xlsx"

# Convert just one semester to a specific output path
python3 scripts/xlsx_to_csv.py "data/Summer 2025-Spring 2026 Syllabi Review.xlsx" --semester "Spring 2026" --output "data/Spring 2026.csv"
```

`--output` requires `--semester` (ambiguous otherwise) and only works with a
single input file. If you convert manually, either also rebuild
`data/manifest.json` yourself:

```bash
python3 -c "
import json, glob
files = sorted(f'data/{f.split(\"/\")[-1]}' for f in glob.glob('data/*.csv'))
json.dump(files, open('data/manifest.json', 'w'), indent=2)
"
```

...or just push — the next run of `build-data.yml` will regenerate it for
you regardless of whether you edited CSVs by hand or via the script.

See the docstring at the top of `scripts/xlsx_to_csv.py` for the exact column
positions and matching rules if you need to debug why a row didn't come
through as expected.

## CSV schema

Each generated `data/<Semester>.csv` has these columns, one row per
(course × cross-listed department):

| column            | meaning                                                    |
|-------------------|-------------------------------------------------------------|
| `course_title`    | course name                                                  |
| `department`      | one department code for this row (e.g. `AFST`)               |
| `department_full` | full department name for this row's code, if known (e.g. `Africana Studies`); blank when the source workbook only provided a short code |
| `all_departments` | cross-listed department codes joined by `/` (e.g. `AFST/ENGL`)      |
| `course_number`   | course number, as text                                        |
| `level`           | `Undergraduate` or `Graduate`                                 |
| `description`     | full course description                                       |
| `classification`  | `Focused` or `Inclusive`                                       |
| `sdgs`            | comma-separated SDG numbers, e.g. `4,13,15`                    |
| `syllabus_url`    | link to the syllabus, or blank if none found                   |
| `semester`        | e.g. `Spring 2026`                                              |

`index.html` never trusts anything about column order — it reads by header
name — so reordering columns is safe, but renaming one requires updating the
matching field name in `index.html`'s `buildCard()`/`render()` functions too.

## Front-end features (`index.html`)

- **Tabs**: Undergraduate / Graduate filter the course list by level; a third
  "New Course or Reclassification?" tab swaps the whole view for an embedded
  Google Form instead of filtering — see the `isReclass` branch in the
  tab-click handler if you need to add a fourth tab.
- **Search, filters, sort, SDG pills**: all operate client-side over the data
  already loaded — no additional requests.
- **Click-to-expand**: clicking a card (or pressing Enter/Space while it's
  focused) expands its full description; other cards in the same visual row
  expand together for a tidy layout (`toggleRowExpansion` in the script).
- **Disclaimer banner**: the strip at the top of the page — edit the text
  directly in the `<div class="disclaimer-banner">` element if the wording
  needs to change.
- **Reclassification form**: the iframe `src` in `#reclass-view` points at a
  Google Form's embed URL (`.../viewform?embedded=true`). To point it at a
  different form, replace that URL.

## Color / Brand

Colors from the OSCE brand guide, defined once in `:root` of `index.html` and
used throughout via CSS variables:
- `--green-dark: #3a6124` — main accent, card left border, badges
- `--green-mid: #4a7a2e` — secondary, SDG tags
- `--green-forest: #194007` — disclaimer banner background
- `--green-deeper: #013120` — header background, footer
- `--green-light: #a6ce3f` — hero highlights
- `--maroon: #500000` — hero background (TAMU maroon)

## Troubleshooting

- **A course is missing from the site**: check the Action's log output (or
  run the script locally) — it prints how many rows each sheet produced. A
  row is dropped if its `Course Title` cell is blank, or if the sheet's
  header row wasn't found (the script looks for a cell containing "course
  title" to locate it).
- **No syllabus link showing**: the course wasn't found in the review
  workbook's own URL column or in any raw-lookup workbook's `(SUBJECT,
  COURSE)` pairs in `data/`. This is expected until a link becomes available
  — cards simply omit the link.
- **Site shows stale/no data locally**: `index.html` fetches
  `data/manifest.json` over HTTP, so opening it directly via `file://` will
  fail that fetch (browsers block local file fetches) and fall back to the
  hardcoded `DATA_FILES` list in the script section of `index.html`. Serve
  the folder locally instead, e.g. `python3 -m http.server` from the repo
  root, then visit `http://localhost:8000`.
- **GitHub Action didn't run**: it only triggers on pushes that touch
  `data/**.xlsx` on `main`. You can also trigger it manually from the
  **Actions** tab via "Run workflow" (`workflow_dispatch`).
