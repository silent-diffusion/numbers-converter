# Numbers Converter

Convert Apple Numbers (`.numbers`) spreadsheets to Excel (`.xlsx`) on Windows, entirely offline.

![Icon](build/icon.png)

## Install

1. Download **`NumbersConverter-Setup-x.y.z.exe`** from the [latest release](https://github.com/silent-diffusion/numbers-converter/releases/latest).
2. Double-click it. It installs for the current user (no admin needed) and opens the app.

> The installer isn't code-signed yet, so Windows SmartScreen may say "Windows protected your PC". Click **More info → Run anyway**.

## Use

- Drag `.numbers` files onto the window, or click **Choose files**.
- Each file is saved as an `.xlsx` next to the original (change this in **Settings → General**). Existing files are never overwritten: you get `Budget (1).xlsx` instead.
- You can also right-click a `.numbers` file → **Open with → Numbers Converter**.

Every table becomes a sheet with its values. Charts, images and some formatting are not carried over.
Files saved by Numbers as a *package* (a folder) must be re-saved as a single file or zipped first.

## Updates

**Settings → Updates**

- **Automatic updates** (on by default): checks GitHub Releases at startup and every 6 hours, downloads new versions in the background, and installs them when you close the app or click **Restart to update**.
- Turn it off to update only when you click **Check for updates**.

Conversion itself never uses the internet; only the update check does.

## Develop

```bash
npm install
npm start        # run the app
npm test         # conversion round-trip tests
npm run dist     # build dist/NumbersConverter-Setup-<version>.exe
```

### Publishing a new version

1. Bump `version` in `package.json` and commit.
2. `npm run dist`
3. Create the release with the installer and update metadata attached (installed apps read `latest.yml` to find updates):

   ```bash
   gh release create v1.0.1 dist/NumbersConverter-Setup-1.0.1.exe dist/NumbersConverter-Setup-1.0.1.exe.blockmap dist/latest.yml --title "v1.0.1" --notes "What changed"
   ```

Built with [Electron](https://www.electronjs.org/), [SheetJS](https://sheetjs.com/) and [electron-updater](https://www.electron.build/auto-update).

## License

MIT
