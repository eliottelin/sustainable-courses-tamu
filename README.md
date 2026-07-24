# OSCE Sustainable Course Guide

Student-facing course guide for Texas A&M sustainability-classified courses.
Built with static HTML/CSS/JS — no server, build step, or framework required.
`index.html` is the entire site; everything else exists to produce the CSV
data it reads.

## Repo Structure

```
index.html                     ← the entire site (markup + CSS + JS in one file)
data/
  *.xlsx                       ← source spreadsheets you drop in (see "Input files" below)
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

## Upkeep: adding a new semester (the normal yearly task)

1. Get the syllabi review spreadsheet into the "review" layout described
   below — a sheet per semester/level named like `"Fall 2026 - Undergraduate"`.
2. Drop it into `data/` (any filename ending in `.xlsx`, e.g.
   `data/Fall 2026.xlsx`) and push to `main`.
3. GitHub Actions does the rest (see "How the site works" above) — no other
   steps needed. Check the **Actions** tab on GitHub if you want to confirm
   the run succeeded; it prints exactly which sheets it found and how many
   rows each produced.
4. Old semester `.xlsx`/`.csv` files can just stay in `data/` — the site
   groups and sorts courses by semester automatically, so nothing needs to be
   removed when a new one is added. Delete old source `.xlsx` files only if
   you no longer want them kept in the repo for reference.

You only need the manual commands below if you're testing a conversion
locally before pushing, debugging why a workbook didn't produce a CSV, or
regenerating a single semester's CSV without touching the others.

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
syllabus link. Cross-listed departments (e.g. `AFST/ENGL`) are expanded into
one row per department.

**2. Raw lookup layout** — a CRN/keyword-scan export whose header row
includes `SUBJECT`, `COURSE`, and `URL` columns (department, course number,
and a syllabus link — extra columns like keyword-match counts are ignored).
This isn't converted into a CSV of its own; instead it's scanned to build a
`(subject, course number) → url` lookup, which fills in `syllabus_url` for
review rows that don't already have one. This is how syllabus links get
attached even when the review spreadsheet itself has no URL column — keep a
file like this in `data/` alongside the review workbook if you have one.

## Regenerating CSVs manually

Useful for local testing, debugging a conversion, or fixing one semester
without re-running the full pipeline.

```bash
# Convert every semester/level sheet found in a workbook (writes data/<Semester>.csv for each)
python3 scripts/xlsx_to_csv.py "data/(Eliotte) Spring 2026 Syllabi Review.xlsx"

# Convert all xlsx files in data/ together, so syllabus URLs join across files
# (this is what the GitHub Action runs)
python3 scripts/xlsx_to_csv.py data/*.xlsx

# Convert just one semester to a specific output path
python3 scripts/xlsx_to_csv.py "data/(Eliotte) Spring 2026 Syllabi Review.xlsx" --semester "Spring 2026" --output "data/Spring 2026.csv"
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
| `all_departments` | full cross-listing string as it appeared in the sheet (`AFST/ENGL`) |
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
