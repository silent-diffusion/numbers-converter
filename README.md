# Numbers Converter

Convert Apple Numbers (`.numbers`) spreadsheets to Excel (`.xlsx`) on Windows, entirely offline.

![Icon](build/icon.png)

## Install

1. Download **`NumbersConverter-Setup-x.y.z.exe`** from the [latest release](https://github.com/silent-diffusion/numbers-converter/releases/latest).
2. Double-click it. It installs for the current user (no admin needed) and opens the app.

> The installer isn't code-signed yet, so Windows SmartScreen may say "Windows protected your PC". Click **More info → Run anyway**.

## Use

- Drag `.numbers` files onto the window, or click **Choose files**. Each file opens in a preview showing its sheets, cells and formatting.
- Click **Convert & save** to save it as an `.xlsx` next to the original (change the folder in **Settings → General**). Existing files are never overwritten: you get `Budget (1).xlsx` instead.
- To skip the button, turn on **Settings → General → Convert automatically**. Files are then saved as soon as they're opened.
- **Sheets with more than one table** (Settings → General): **Separate sheets** (the default) puts each table on an Excel sheet of its own; **One sheet** keeps a Numbers sheet's tables and charts together on one Excel sheet, laid out as in Numbers. The preview has the same switch for the sheet you're looking at, so each sheet can go either way.
- You can also right-click a `.numbers` file → **Open with → Numbers Converter**.

### What carries over

Each Numbers table becomes an Excel sheet (with the table's name in a title row above it when Numbers shows it). Or, with **One sheet**, each Numbers sheet becomes one Excel sheet with every table where it sits in Numbers: tables side by side share rows, tables one above the other share columns (as wide as the widest of them needs, so nothing gets cut off), and the gaps between them become rows and columns of the same size. Formulas that read another table follow it wherever it ends up. Either way, tables keep:

- **Values** — numbers, text, dates, durations, checkboxes
- **Table styles** — the table theme's lines, header/footer colours and banded rows, the sheet's page colour, and Numbers' own row heights and column widths (in points, so text fits as it does in Numbers). Excel's gridlines are turned off, as Numbers shows none.
- **Formatting** — fonts, sizes, bold/italic/underline/strikethrough, text and fill colours, cell borders, alignment, text wrapping, merged cells, header rows/columns frozen where Numbers freezes them
- **Number formats** — currency (including accounting style), percentages, decimals, thousands separators, negative styles, fractions, durations, date and time formats
- **Cell extras** — comments (as Excel notes), links, bold/italic/coloured words within a cell, bullet and numbered lists, hidden rows and columns, filters (as Excel AutoFilter)
- **Form controls** — checkboxes (Excel checkboxes), pop-up menus (dropdown lists), sliders, steppers and star ratings (as number-range rules)
- **Formulas** — rewritten as Excel formulas so they keep calculating, including references to other tables and to columns by their header name (`SUM(Amount)`). Formulas that use Apple-only functions, functions that behave differently in Excel, or approximate lookups on unsorted data (where Excel's answer can differ) keep their last value instead, and the app tells you how many.
- **Pivot tables** — newer versions of Numbers don't save pivot totals in the file, so the app recalculates them from the pivot's rules and its source data (grouping by value, date part, etc.; sum, average, count, minimum, maximum). Pivot results are saved as values, not as an Excel pivot table. Any pivot it can't reproduce exactly (e.g. "show as percentage" or running totals) is reported and its totals are left blank rather than guessed.

- **Charts** — column, bar, stacked, line, area, pie, donut, scatter, radar, 2-axis and mixed charts become real Excel charts that read the converted cells, so they update when the data changes. Each keeps its place beside, above or below its table (rows are added above a table to make room for charts over it, as Apple's own export does), its colours, title, legend, gaps, value labels, axis titles and gridlines. The preview shows them where they'll be. 3D charts come out flat, and interactive charts show all their data sets at once. Bubble charts aren't converted yet; the app tells you when a file has one.

Not carried over yet: **images, shapes and text boxes** (the app tells you when a file has images or shapes), conditional highlighting, and the grouped view of tables with categories (their rows come out in their original order).

Password-protected files ask for the password (and show its hint); it's only kept in memory while the app is open.
Files a Mac saved as a *package* (a folder ending in `.numbers`) can be dragged onto the window.

## Updates

**Settings → Updates**

- **Automatic updates** (on by default): checks GitHub Releases at startup and every 6 hours, downloads new versions in the background, and installs them when you close the app or click **Restart to update**.
- Turn it off to update only when you click **Check for updates**.

Conversion itself never uses the internet; only the update check does.

## Develop

The Numbers reader and Excel writer is a Python helper ([`py/numbers_tool.py`](py/numbers_tool.py), with pivot
recalculation in [`py/pivots.py`](py/pivots.py), charts in [`py/charts.py`](py/charts.py) and sheet layout in
[`py/layout.py`](py/layout.py), built on
[numbers-parser](https://github.com/masaccio/numbers-parser) and [XlsxWriter](https://xlsxwriter.readthedocs.io/)).
The installer ships it as a standalone PyInstaller build, so users don't need Python.

```bash
npm install
py -3.11 -m venv .venv
.venv\Scripts\pip install -r py/requirements.txt
npm start        # run the app
npm test         # helper tests + Node bridge tests (uses real Numbers files in test/fixtures)
npm run dist     # build the helper and dist/NumbersConverter-Setup-<version>.exe
```

### Measuring against Apple's own export

[`tools/compare_with_apple.py`](tools/compare_with_apple.py) converts every `X.numbers` in a folder that also has
Apple's `X.xlsx` export (File → Export To → Excel in Numbers; name exports with Numbers' other worksheet option
`X 2.xlsx`), laid out the way that export was (a worksheet per table, or one per Numbers sheet), opens both in
Excel, and scores how closely each cell matches: text, formulas, fonts, fills, borders, alignment, number formats,
row heights and column widths; and each chart: its type, and the values and labels Excel reads for each series.
Needs Windows with Excel. Pivot tables are skipped (Apple writes Excel pivot tables, laid out Excel's way).

```bash
.venv\Scripts\python tools/compare_with_apple.py path\to\folder
```

### Publishing a new version

1. Bump `version` in `package.json` and commit.
2. `npm run dist`
3. Create the release with the installer and update metadata attached (installed apps read `latest.yml` to find updates):

   ```bash
   gh release create v1.0.1 dist/NumbersConverter-Setup-1.0.1.exe dist/NumbersConverter-Setup-1.0.1.exe.blockmap dist/latest.yml --title "v1.0.1" --notes "What changed"
   ```

Built with [Electron](https://www.electronjs.org/), [numbers-parser](https://github.com/masaccio/numbers-parser), [XlsxWriter](https://xlsxwriter.readthedocs.io/) and [electron-updater](https://www.electron.build/auto-update).

## License

MIT
