# OpenCells

**An open-source spreadsheet app for Android phones and tablets.**
Open, edit and save `.xlsx`, `.xlsm` and `.csv` files from your phone storage, with formulas, charts,
pivot tables, JavaScript macros and a Power-Query-style data transform panel, all on the device.

License: **GNU GPL v3** (see [LICENSE](LICENSE)). Third-party components: [THIRD_PARTY_LICENSES.md](THIRD_PARTY_LICENSES.md).

> Download: ready-built APKs are in [`releases/`](releases/). See [Install](#install) below.

<!-- Screenshots: replace the placeholders below with real device captures -->
| Phone | Tablet | Dark mode |
|---|---|---|
| ![phone screenshot placeholder](docs/screenshots/phone.png) | ![tablet screenshot placeholder](docs/screenshots/tablet.png) | ![dark mode placeholder](docs/screenshots/dark.png) |

---

## Features

| # | Feature | Notes |
|---|---|---|
| 1 | **Open / edit / save** `.xlsx`, `.xlsm`, `.csv` (also opens `.xls`, `.ods`) | Android Storage Access Framework picker (Downloads, Drive, SD card…). **Save** writes back to the same file, **Save as** creates a new one. "Open with OpenCells" from file managers. |
| 2 | **Formulas** with 420+ functions via HyperFormula | SUM, AVERAGE, IF, IFS, XLOOKUP, VLOOKUP, INDEX/MATCH, SUMIFS, COUNTIFS, AVERAGEIFS, TEXT (Excel number formats), DATE, TODAY, EOMONTH, EDATE, DATEDIF, NETWORKDAYS… Cross-sheet references, function suggestions in the formula bar. |
| 3 | **Sheets, freeze panes, sort, filter, find & replace** | Add / rename / duplicate / delete sheets (long-press a tab). Freeze top row, first column or at selection. Sort keeps header rows. Column filters. |
| 4 | **Formatting** | Number formats (incl. custom patterns), bold/italic/underline/strike, text and fill colours, borders, alignment, wrap, merge cells, conditional formatting (greater/less/between/equal, text contains, duplicates, colour scale, data bar). Saved to xlsx. |
| 5 | **Charts** | Bar, line and pie from any range; update live as data changes; stored in the file. |
| 6 | **Pivot tables** | Row fields × column field with sum, count, average, min or max, plus totals. Refreshable. |
| 7 | **Undo / redo and autosave** | Unlimited-style undo stack; the session autosaves to app storage on every change (and when the app goes to background) and is restored at next launch. |
| 8 | **Macros** | JavaScript macros with `getCell`, `setCell`, `getRange`, `setRange`, `addSheet`, `sheetNames`, `useSheet`, `log`. Run in a sandboxed Web Worker (no network or storage) with a time limit; changes are undoable. **VBA in `.xlsm` files is preserved on save but not executed** (VBA cannot run on Android). |
| 9 | **Data transforms** (Power Query-like) | Import CSV/xlsx or use a sheet range, then record steps: filter rows, remove columns, change type, merge tables (joins), group by, pivot, unpivot, sort, rename. Steps are saved in the workbook and can be **refreshed** later. |

### Design
* Original look (not modelled on any office suite): indigo/violet accent, rounded touch targets of ≥ 48 dp.
* **Phone:** bottom toolbar (Edit · Format · Insert · Data · Tools) opening bottom sheets; the formula bar is
  docked at the bottom so it sits right above the keyboard.
* **Tablet (≥ 760 dp wide):** formula bar on top and a side panel with the same tools.
* Pinch-to-zoom, inertial panning, light and dark mode (follows the system, toggle in the ⋮ menu).

### Performance and device support
* Android **8.0+** (minSdk 26), target SDK 36. Universal APK for **arm64-v8a, armeabi-v7a and x86_64**.
  The app contains no native code, so the same APK runs on Snapdragon, MediaTek Dimensity/Helio,
  Samsung Exynos, Google Tensor and x86 emulators/Chromebooks.
* At startup the app reads total RAM and CPU cores (native `DeviceInfo` plugin) and picks a tier:

  | Tier | RAM | Load chunk | Workers | Effects | Warn above 100k rows |
  |---|---|---|---|---|---|
  | Low | ≤ 3 GB | 1,000 rows | 1 | reduced motion | yes |
  | Mid | 4–6 GB | 4,000 rows | 2 | normal | yes |
  | High | 8 GB+ | 10,000 rows | up to 4 | smooth, 90/120 Hz requested | no |

  Phones report slightly less RAM than advertised (an 8 GB phone shows about 7.4 GB), so the value is
  rounded up to the nearest standard size. Devices with 4 or fewer cores drop one tier.
  You can force a tier for testing with `?tier=low|mid|high`.
* **Everything heavy runs in Web Workers:** formulas (HyperFormula), file parsing and writing
  (ExcelJS / SheetJS), transforms and pivots (Arquero, worker pool sized by tier) and macros (one throw-away worker per run).
* **Virtualised canvas grid** (Univer): only visible cells are drawn. Render data for hidden sheets is released.
* **Incremental recalculation:** only cells depending on an edit are recomputed (HyperFormula dependency graph).
* **Chunked loading** with a progress bar; files are transferred to workers without copying (Transferables).
* Debounced input (formula suggestions, chart redraws). Release build is minified and code-split.
* Measured results for a 50,000 × 20 file: [tests/perf/RESULTS.md](tests/perf/RESULTS.md).

---

## Install

1. Download `releases/OpenCells-1.0.0-release.apk` (or the `-debug` one) from this repository on your phone.
2. Open it and allow "Install unknown apps" for your browser or file manager when Android asks.
3. Launch **OpenCells** and tap the folder icon to open a spreadsheet.

The release APK in `releases/` is signed with a key created for this build. If you later build and sign your
own release with a different key, uninstall the old app first (Android refuses updates signed by another key).

---

## Build from source

### Requirements
* Node.js 20+ (22 recommended) and npm
* JDK 21
* Android SDK with platform 36 and build-tools 35+ (Android Studio, or the command line tools below)

```bash
# Android SDK without Android Studio (Linux/macOS)
export ANDROID_HOME=$HOME/android-sdk
mkdir -p $ANDROID_HOME/cmdline-tools && cd $ANDROID_HOME/cmdline-tools
curl -LO https://dl.google.com/android/repository/commandlinetools-linux-14742923_latest.zip
unzip commandlinetools-linux-*_latest.zip && mv cmdline-tools latest
yes | latest/bin/sdkmanager --licenses
latest/bin/sdkmanager "platform-tools" "platforms;android-36" "build-tools;36.0.0"
```

### Debug APK

```bash
git clone https://github.com/aiman-959/Excel-apk.git opencells && cd opencells
npm ci
npm run build                 # type-check + minified, code-split web build into dist/
npx cap sync android          # copy dist/ into the Android project
cd android
echo "sdk.dir=$ANDROID_HOME" > local.properties
./gradlew assembleDebug
# -> android/app/build/outputs/apk/debug/app-debug.apk
adb install -r app/build/outputs/apk/debug/app-debug.apk   # optional, phone connected with USB debugging
```

### Release APK (signed)

```bash
# 1. one-time: create your signing key (keep it safe and never commit it)
keytool -genkeypair -v -keystore ~/opencells-release.jks -alias opencells \
        -keyalg RSA -keysize 4096 -validity 10000

# 2. tell Gradle where it is (android/keystore.properties is git-ignored)
cat > android/keystore.properties <<'EOF'
storeFile=/home/you/opencells-release.jks
storePassword=YOUR_STORE_PASSWORD
keyAlias=opencells
keyPassword=YOUR_KEY_PASSWORD
EOF

# 3. build
npm ci && npm run build && npx cap sync android
cd android && ./gradlew assembleRelease
# -> android/app/build/outputs/apk/release/app-release.apk  (R8-minified, resources shrunk, signed)
# Google Play wants an App Bundle instead:  ./gradlew bundleRelease
```

Instead of `keystore.properties` you can set the environment variables `OPENCELLS_KEYSTORE`,
`OPENCELLS_KEYSTORE_PASSWORD`, `OPENCELLS_KEY_ALIAS` and `OPENCELLS_KEY_PASSWORD`.
Without either, `assembleRelease` produces an unsigned APK.

### Continuous integration
`.github/workflows/android.yml` builds both APKs on every push and uploads them as workflow artifacts.
Pushing a tag such as `v1.0.1` also creates a GitHub Release with the APKs attached. Add the repository secrets
`KEYSTORE_BASE64` (`base64 -w0 opencells-release.jks`), `KEYSTORE_PASSWORD`, `KEY_ALIAS` and `KEY_PASSWORD`
to get a signed release APK.

### Develop in the browser

```bash
npm run dev            # http://localhost:5173  (?mobile, ?desktop, ?tier=low|mid|high, ?fresh)
npm test               # unit tests (formulas, xlsx/csv/VBA round trip, transforms, pivots, macros)
npm run build && npm run preview &                  # serve the production build on :4173
BASE_URL=http://localhost:4173/ npm run test:e2e   # end-to-end test of features 1-9 (18 checks)
BASE_URL=http://localhost:4173/ npm run test:perf  # 50k x 20 load / scroll / memory test -> tests/perf/RESULTS.md
```
Playwright needs a Chromium: `npx playwright install chromium`, or point `CHROMIUM_PATH` at an existing one.

---

## Project layout

```
src/
  univer/        Univer grid setup (desktop + mobile plugins), conditional-format mapping
  engine/        HyperFormula worker, formula bridge (incremental recalc), RPC helper
  io/            xlsx read/write (ExcelJS), csv/xls/ods (SheetJS), VBA preservation, file access
  device/        performance tier detection, native plugin wrappers
  state/         document controller (open/save/autosave), IndexedDB storage
  ui/            app chrome, toolbars, dialogs (touch-first)
  charts/ pivot/ macros/ transform/   feature panels (lazy-loaded) and their workers
android/app/src/main/java/org/opencells/app/
  MainActivity.java, DeviceInfoPlugin.java (RAM/cores/SoC/refresh rate), DocumentsPlugin.java (SAF)
tests/unit  tests/e2e  tests/perf
```

## Known limitations
* VBA macros are kept but never executed. JavaScript macros cover the automation use case.
* Charts, pivot definitions, saved macros and transform queries are stored in a hidden worksheet
  (`__opencells_meta`). Other apps keep them but do not show them as native charts or pivots.
* xlsx features outside the list above (images, comments, data validation, sparklines, external links) are not imported.
* Very large files (500k+ cells on Low tier) can exceed WebView memory. The app warns before opening files over 100k rows on Low and Mid tiers.

## Contributing
Issues and pull requests are welcome. By contributing you agree your work is licensed under GPL-3.0.

## License
OpenCells is free software: you can redistribute it and/or modify it under the terms of the GNU General
Public License version 3 as published by the Free Software Foundation. It is distributed WITHOUT ANY WARRANTY.
See [LICENSE](LICENSE).
