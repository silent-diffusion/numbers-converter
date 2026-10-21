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
- You can also right-click a `.numbers` file → **Open with → Numbers Converter**.

### What carries over

Each Numbers table becomes an Excel sheet, keeping:

- **Values** — numbers, text, dates, durations, checkboxes
- **Formatting** — fonts, sizes, bold/italic/underline/strikethrough, text and fill colours, borders, alignment, text wrapping, merged cells, column widths, row heights, frozen header rows/columns
- **Number formats** — currency, percentages, decimals, thousands separators, negative styles, date and time formats
- **Formulas** — rewritten as Excel formulas so they keep calculating. Formulas that use Numbers-only features (header-name references, Apple-only functions, functions that behave differently in Excel) keep their last value instead, and the app tells you how many.
- **Pivot tables** — newer versions of Numbers don't save pivot totals in the file, so the app recalculates them from the pivot's rules and its source data (grouping by value, date part, etc.; sum, average, count, minimum, maximum). Pivot results are saved as values, not as an Excel pivot table. Any pivot it can't reproduce exactly (e.g. "show as percentage" or running totals) is reported and its totals are left blank rather than guessed.

Not carried over: **charts, images and shapes** (the app tells you when a file has them), conditional highlighting, comments.
To keep charts, export from Numbers itself (**File → Export To → Excel** on a Mac, iPhone, iPad or iCloud.com).

Files saved by Numbers as a *package* (a folder) must be re-saved as a single file or zipped first. Password-protected files need the password removed in Numbers first.

## Updates

**Settings → Updates**

- **Automatic updates** (on by default): checks GitHub Releases at startup and every 6 hours, downloads new versions in the background, and installs them when you close the app or click **Restart to update**.
- Turn it off to update only when you click **Check for updates**.

Conversion itself never uses the internet; only the update check does.

## Develop

The Numbers reader and Excel writer is a Python helper ([`py/numbers_tool.py`](py/numbers_tool.py), with pivot
recalculation in [`py/pivots.py`](py/pivots.py), built on
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
