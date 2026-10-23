# Numbers Converter

Convert Apple Numbers (`.numbers`) spreadsheets to Excel (`.xlsx`) on Windows, entirely offline.

![Icon](build/icon.png)

## Install

1. Download **`NumbersConverter-Setup-x.y.z.exe`** from the [latest release](https://github.com/silent-diffusion/numbers-converter/releases/latest).
2. Double-click it. It installs for the current user (no admin needed), opens the app, and keeps it up to date (see [Updates](#updates)).

> The installer isn't code-signed yet, so Windows SmartScreen may say "Windows protected your PC". Click **More info → Run anyway**.

Or install it with [Scoop](https://scoop.sh). Scoop installs a copy of its own, which `scoop update numbers-converter` keeps up to date:

```bash
scoop bucket add numbers-converter https://github.com/silent-diffusion/numbers-converter
scoop install numbers-converter
```

Or don't install it at all: unzip **`NumbersConverter-x.y.z-win-x64.zip`** from the release anywhere and open `Numbers Converter.exe`. That copy doesn't update itself; get the new zip when there's a new version.

**Checking a download:** each release lists the SHA-256 of its installer and zip in `SHA256SUMS.txt` (compare with `Get-FileHash <file>` in PowerShell), and GitHub keeps a signed record that they were built from this repository by its release workflow. To check that record:

```bash
gh attestation verify NumbersConverter-Setup-x.y.z.exe --repo silent-diffusion/numbers-converter
```

## Use

- Drag `.numbers` files onto the window, or click **Choose files** (Ctrl+O). Each file opens in a preview of its sheets, cells, formatting, charts, text boxes and images as they'll be in Excel. Files a Mac saved as a *package* (a folder ending in `.numbers`) can be dragged in too, or opened with **Choose a folder**.
- Drag in a folder, or choose one with **Choose a folder** (or **⋯ → Add a folder**), and every Numbers file in it and its subfolders opens, up to 200 at a time. For more at once, use the [command line](#command-line).
- Before you save, the button beside **Convert & save** says whether the file **converts fully** or how many **changes** to check. It opens **What changes in Excel**, which lists:
  - formulas that keep their last value, with why and in which cells (**Show** jumps to each one);
  - things that can't be converted, such as bubble charts, pictures in a format Excel can't show, or highlighting rules Excel can't calculate;
  - smaller differences, such as Mac fonts Windows doesn't have.
- Click **Convert & save** (Ctrl+S) to save an `.xlsx` next to the original, or in a folder you choose in Settings. Existing files are never overwritten: you get `Budget (1).xlsx` instead. **Convert all** (Ctrl+Shift+S) saves every file that's ready.
- **Settings** (the gear at the top right, or Ctrl+,):
  - **Convert files as soon as they open**, to skip the button.
  - **Save Excel files**: next to the original, or in a folder you choose.
  - **Sheets with more than one table**: **Separate sheets** (the default) puts each table on an Excel sheet of its own; **One sheet** keeps a Numbers sheet's tables and charts together on one Excel sheet, laid out as in Numbers. The preview has the same switch for the sheet you're looking at, so each sheet can go either way, and the app remembers your choice for that file.
  - **Add a Contents sheet**: starts the Excel file with a sheet linking to every table, like the summary sheet Numbers can add when it exports.
- You can also right-click a `.numbers` file → **Open with → Numbers Converter**.
- If a file can't be read, the app says why, shows the picture Numbers saves inside every file, and lets you **Try again** or **Report a problem**.
- Keyboard: Ctrl+O open files · Ctrl+S convert & save · Ctrl+Shift+S convert all · ↑ ↓ previous or next file · Delete remove the file from the list · Ctrl+PgUp / PgDn previous or next sheet · Ctrl+, settings · Esc close a panel or go back · Ctrl+= / Ctrl+- / Ctrl+0 zoom.

### What carries over

Each Numbers table becomes an Excel sheet (with the table's name in a title row above it when Numbers shows it). Or, with **One sheet**, each Numbers sheet becomes one Excel sheet with every table where it sits in Numbers: tables side by side share rows, tables one above the other share columns (as wide as the widest of them needs, so nothing gets cut off), and the gaps between them become rows and columns of the same size. Formulas that read another table follow it wherever it ends up. Either way, tables keep:

- **Values** — numbers, text, dates, durations, checkboxes
- **Table styles** — the table theme's lines, header/footer colours and banded rows, the sheet's page colour, and Numbers' own row heights and column widths (in points, so text fits as it does in Numbers). Excel's gridlines are turned off, as Numbers shows none.
- **Formatting** — fonts, sizes, bold/italic/underline/strikethrough, text and fill colours, cell borders, alignment, text wrapping, merged cells, header rows/columns frozen where Numbers freezes them
- **Number formats** — currency (including accounting style), percentages, decimals, thousands separators, negative styles, fractions, durations, date and time formats
- **Cell extras** — comments (as Excel notes), links, bold/italic/coloured words within a cell, bullet and numbered lists, hidden rows and columns, filters (as Excel AutoFilter)
- **Form controls** — checkboxes (Excel checkboxes), pop-up menus (dropdown lists), sliders, steppers and star ratings (as number-range rules)
- **Formulas** — rewritten as Excel formulas so they keep calculating, including references to other tables and to columns by their header name (`SUM(Amount)`). Formulas that use Apple-only functions, functions that behave differently in Excel, or approximate lookups on unsorted data (where Excel's answer can differ) keep their last value instead, and the app tells you how many.
- **Pivot tables** — become real Excel PivotTables that read the converted source table, placed where Numbers places them, with the same row, column and value fields, summaries (sum, average, count, minimum, maximum), grouping (by value, or by a date part Excel can group by itself, such as year, quarter or month; any other grouping keeps Numbers' groups, made of the same values), sort order (A to Z, Z to A, by a value's totals, or as Numbers lists them), subtotals and grand totals, and Numbers' labels, number formats and colours (as a PivotTable style, so they stay when Excel refreshes the table). Newer versions of Numbers don't save pivot totals in the file, so the app works them out from the pivot's rules and its source data: the PivotTable opens showing exactly the totals Excel gets when it refreshes it, with no repair prompt. A pivot Excel can't show the same way is reported and its totals are left blank rather than guessed (e.g. "show as percentage" or running totals), or saved as values (a pivot of a table grouped into categories).

- **Charts** — column, bar, stacked, line, area, pie, donut, scatter, radar, 2-axis and mixed charts become real Excel charts that read the converted cells, so they update when the data changes. Each keeps its place beside, above or below its table (rows are added above a table to make room for charts over it, as Apple's own export does), its colours, title, legend, gaps, value labels, axis titles and gridlines. The preview shows them where they'll be. 3D charts come out flat, and interactive charts show all their data sets at once. Bubble charts aren't converted yet; the app tells you when a file has one.
- **Text boxes and shapes** — titles, notes and drawings keep their text (paragraphs, fonts, bold/italic/coloured words, alignment, spacing, bullets and numbering, padding), fill, border and line, and their outline: rectangles, rounded rectangles, ovals, lines, arrows, stars, polygons, speech bubbles and freehand shapes, turned and flipped as in Numbers, including inside groups. They keep their place as charts do; a title or note above a table gets room made above it, as in Apple's own export.
- **Images** — pictures keep their place, size, crop (including rounded or oval masks), border, transparency, description and link. For a picture Excel can't show (HEIC, TIFF or PDF), the app uses the preview Numbers keeps of it where there is one. A video shows its poster frame.
- **Conditional highlighting** — each rule becomes an Excel conditional formatting rule on the same cells, with the same text colour, fill, bold, italic, underline or strikethrough: comparisons with a value as "Cell Value" rules, anything else Excel can calculate as formula rules. As in Numbers, the first rule a cell matches is the one that applies.
- **Categories** — a table grouped into categories comes out grouped, as Numbers shows it and as Apple's own export writes it: a category column with each group's name, and each group's row above its rows (with groups inside it indented under it), showing its summaries (sum, average, minimum, maximum, count) as formulas that keep calculating, under a label row that names them. Each group is also an Excel outline group, so it collapses with the buttons beside the rows as it does in Numbers (not on **One sheet**, where tables beside it share its rows). Formulas that read the table follow its rows to where they went.

Not carried over yet: bubble charts, pictures Excel can't show that have no preview, and the occasional highlighting rule Excel can't calculate. When a file has any of these, **What changes in Excel** says so before you save.

Password-protected files ask for the password and show its hint; the password is only kept in memory while the app is open.

## Command line

The app comes with a `numbers-converter` command, which converts files and whole folders without opening the
window, for batches and scripts. It saves the same Excel files the app does, on your PC, and never writes over a
file unless you ask it to.

- Installed with Scoop, it's on your PATH: run `numbers-converter --help`.
- Otherwise it's in the app's folder, and **Settings → Help → Command line** shows how to run it, ready to copy. For
  the installer's copy, in PowerShell:

  ```powershell
  & "$env:LOCALAPPDATA\Programs\numbers-converter\numbers-converter.cmd" --help
  ```

  To run it as just `numbers-converter`, add that folder to your PATH.

```powershell
numbers-converter Budget.numbers                     # Budget.xlsx, next to it
numbers-converter C:\Reports --out D:\Excel          # every Numbers file in Reports and its subfolders, into the same subfolders in D:\Excel
numbers-converter *.numbers --layout one --contents  # wildcards work; One sheet, with a Contents sheet
numbers-converter C:\Reports --dry-run               # what would be saved, and what changes in Excel, saving nothing
numbers-converter C:\Reports --json > summary.json   # a summary for scripts
```

- `--out FOLDER`: save the Excel files there instead of next to the originals. Files found in subfolders go in the
  same subfolders there.
- `--layout separate` or `--layout one`: how sheets with more than one table are saved, as in Settings. Separate
  sheets is the default.
- `--contents`: start each Excel file with a Contents sheet.
- `--overwrite`: write over Excel files that are already there, instead of adding " (1)" to the name.
- `--password-file FILE`: the password for locked files, on the first line of FILE. In a terminal, the command asks
  for any password it needs instead.
- `--dry-run`: say what would be saved where, and what changes in Excel, without saving anything.
- `--json`: print a summary in JSON. For each file, it has its status (`saved`; `ready` in a dry run; `locked` or
  `failed`, with why), where its copy went, and the changes **What changes in Excel** lists.

It exits with 0 when every file was saved, 1 when a file couldn't be converted or needs a password, and 2 when no
Numbers files were found or the command has a mistake in it. It doesn't use the app's settings: it does what the
command says.

## Updates

**Settings → Updates**

- **Check for updates automatically** (on by default): checks GitHub Releases at startup and every 6 hours, downloads new versions in the background, and installs them the next time you restart. When one is ready, an **Update ready · Restart** button appears at the top of the window.
- Turn it off and the app never goes online on its own; it only checks when you click **Check for updates**.

Copies installed with Scoop, or unzipped from the zip, don't update themselves, as that would install a second copy beside them. Scoop updates its copy (`scoop update numbers-converter`), and a zip copy is replaced by the new zip; **Settings → Updates** says which applies.

## Privacy

Numbers Converter does all of its work on your PC.

- **Your files are never uploaded.** Previews and conversions run in the app's own converter on your computer, with no internet connection needed.
- **No account, no ads, no tracking.** The app collects no usage data.
- **The only time it goes online is to check GitHub for updates:** at startup and every 6 hours while automatic updates are on, or when you click **Check for updates**.
- **Report a problem** opens a new GitHub issue in your browser, filled in with the app version, the Windows version and the error message. File names and contents are left out, and nothing is sent until you post the issue yourself.
- **Problems are written to a log on your PC** (`%APPDATA%\Numbers Converter\logs`, or **Settings → Help → Open log folder**). It stays there unless you send it to someone.
- **Passwords** for protected files are kept in memory only, and forgotten when you close the app.

## Develop

The window is plain HTML, CSS and JavaScript in [`src/renderer`](src/renderer), and the Electron main process is
[`src/main.js`](src/main.js). The Numbers reader and Excel writer is a Python helper
([`py/numbers_tool.py`](py/numbers_tool.py), which also works out the "What changes in Excel" report, with pivot
recalculation in [`py/pivots.py`](py/pivots.py) and Excel PivotTables in [`py/pivot_tables.py`](py/pivot_tables.py),
charts in [`py/charts.py`](py/charts.py), text boxes, shapes and images in [`py/drawings.py`](py/drawings.py),
conditional highlighting in [`py/highlights.py`](py/highlights.py), tables grouped into categories in
[`py/categories.py`](py/categories.py) and sheet layout in [`py/layout.py`](py/layout.py), built on
[numbers-parser](https://github.com/masaccio/numbers-parser) and [XlsxWriter](https://xlsxwriter.readthedocs.io/)).
The installer ships it as a standalone PyInstaller build, so users don't need Python. Started without `--serve`,
the same build is the `numbers-converter` command ([`py/cli.py`](py/cli.py), which also finds the Numbers files in
folders opened in the app), run by [`build/numbers-converter.cmd`](build/numbers-converter.cmd) beside the app.

```bash
npm install
py -3.11 -m venv .venv
.venv\Scripts\pip install -r py/requirements.txt
npm start        # run the app
npm test          # helper, Node bridge and command-line tests (uses real Numbers files in test/fixtures)
npm run dist      # build the helper, then the installer and the portable zip in dist/
npm run test:ui   # UI smoke test: drives the window of the app built in dist/win-unpacked
npm run test:cli  # the numbers-converter command as the app built in dist/win-unpacked ships it
```

The UI smoke test ([`test/ui/smoke.test.js`](test/ui/smoke.test.js)) runs on every CI build, and nothing is released
unless it passes. It opens every file in `test/fixtures` the ways a user would, previews and saves each one and goes
through the app's screens, with a profile of its own, so it never touches your settings or an open copy of the app.
`npm run test:ui -- --source` tests `src/` without building. To test more files, drop them in a folder in
`test/fixtures`; [`test/ui/expectations.js`](test/ui/expectations.js) lists what particular files must show.

### Measuring against Apple's own export

[`tools/compare_with_apple.py`](tools/compare_with_apple.py) converts every `X.numbers` in a folder that also has
Apple's `X.xlsx` export (File → Export To → Excel in Numbers; name exports with Numbers' other worksheet option
`X 2.xlsx`), laid out the way that export was (a worksheet per table, or one per Numbers sheet), opens both in
Excel, and scores how closely each cell matches: text, formulas, fonts, fills, borders, alignment, number formats,
row heights and column widths; each chart: its type, and the values and labels Excel reads for each series; and
each text box: its text, font and (where both lay a Numbers sheet out on one worksheet) its place, plus lines and
other shapes; and each PivotTable, once Excel has refreshed both: its fields and summaries, grand totals, labels and
values, and whether ours showed those same cells before the refresh. Needs Windows with Excel.

```bash
.venv\Scripts\python tools/compare_with_apple.py path\to\folder
```

### Publishing a new version

GitHub Actions ([`.github/workflows/build.yml`](.github/workflows/build.yml)) builds and tests every push and pull
request on Windows, and publishes a release when a version tag is pushed:

1. Bump `version` in `package.json` and commit it to `main`.
2. Tag that commit with the version, with the release notes as the tag's message, and push the tag:

   ```bash
   git tag -a v1.0.1 -m "What changed"
   git push origin v1.0.1
   ```

The workflow checks that the tag matches `package.json`, then:

- creates the release with the installer, its `.blockmap` and `latest.yml` (installed copies read `latest.yml` to find
  updates), the portable zip and `SHA256SUMS.txt`, and records build provenance for the installer and the zip;
- writes the version's Scoop manifest to `bucket/numbers-converter.json` on `main` (that folder is what makes this
  repository a Scoop bucket);
- sends the version to winget as a pull request to [microsoft/winget-pkgs](https://github.com/microsoft/winget-pkgs),
  once the repository has a `WINGET_TOKEN` secret: a classic personal access token with the `public_repo` scope, from
  an account with a fork of winget-pkgs. The package's first version goes in by hand, from the manifests in
  [`packaging/winget`](packaging/winget).

**Actions → Build → Run workflow** goes through the release steps as a dry run, without publishing anything.

A release made by hand (`npm run dist`, then `gh release create v1.0.1 dist/NumbersConverter-Setup-1.0.1.exe
dist/NumbersConverter-Setup-1.0.1.exe.blockmap dist/latest.yml`) is left as it is, so that version gets no zip,
Scoop manifest or winget update.

Built with [Electron](https://www.electronjs.org/), [numbers-parser](https://github.com/masaccio/numbers-parser), [XlsxWriter](https://xlsxwriter.readthedocs.io/) and [electron-updater](https://www.electron.build/auto-update).

## License

MIT
