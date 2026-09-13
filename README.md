# Student Batch Mailer

Electron app for matching student feedback files to roster entries and mailing them through Outlook.

## Getting started (fresh machine)

### Prerequisites
1. **Homebrew** – https://brew.sh
2. **Node.js 24 LTS** and **pnpm**:
   ```
   brew install node@24 pnpm
   brew link --overwrite --force node@24
   node --version   # should print v24.x
   ```
   > Electron 44 requires Node ≥ 22.12; Node 24 LTS is recommended.
   > Note: Homebrew's pnpm cannot manage Node itself (`pnpm env use` won't work), so install Node via Homebrew as above.
3. **Microsoft Outlook for macOS** (see Platform support below).

### Install & run
```
git clone <repo-url>
cd student-batch-mailer
pnpm install
pnpm run start
```

Notes on the setup:
- `package.json` has a `postinstall` script that runs `install-electron`, which downloads the ~110 MB Electron binary into `node_modules/electron`. Electron 44+ no longer does this from its own postinstall, so without this hook `pnpm run start` would fail.
- `pnpm-workspace.yaml` sets `nodeLinker: hoisted` so `node_modules` has a flat, npm-style layout. `@electron/packager --prune` (used by `pnpm run dist`) cannot follow pnpm's default symlinked layout.

### Build a distributable
```
pnpm run dist
```
Produces `dist/Student Batch Mailer-darwin-arm64/Student Batch Mailer.app` and a zipped copy `dist/Student Batch Mailer-mac.zip` (~130 MB zipped, ~315 MB unpacked).

- **Why so large?** The app bundles its own Chromium + Node runtime (`Electron Framework.framework`, ~285 MB). That is the price of every Electron app and is not affected by the project's own code (~30 MB, mostly the `exceljs` dependency). Only the runtime code, `renderer/`, `report/` and `outlook.scpt` are packaged; `sample-data/`, `scripts/`, `build/`, docs and lockfiles are excluded.
- **App icon:** the source is [`build/icon.svg`](./build/icon.svg). Run `pnpm run icon` after editing it to regenerate `build/icon.icns` (uses only macOS built-in tools: `qlmanage`, `sips`, `iconutil`). `pnpm run dist` picks the `.icns` up automatically.

### Sending the app to a colleague
Send them `dist/Student Batch Mailer-mac.zip`. On their Mac they need:

1. **Apple Silicon (M1 or newer).** The build is arm64-only. For an Intel Mac run `pnpm run dist -- --arch=x64` and send that build instead.
2. **Microsoft Outlook for Mac**, signed in. The app does **not** use the system default mail client — it drives Outlook specifically via AppleScript (`outlook.scpt`). Apple Mail, Thunderbird or web-only Outlook will not work.
3. **First launch:** the app is not code-signed or notarized, so macOS Gatekeeper will refuse to open it with a *"cannot be opened because the developer cannot be verified"* or *"is damaged"* message. The colleague should **right-click the app → Open**, then click **Open** in the dialog (or go to *System Settings → Privacy & Security* and click *Open Anyway*). If macOS still says the app is damaged, run `xattr -cr "/path/to/Student Batch Mailer.app"` once in Terminal to strip the quarantine flag.
4. **Automation permission:** the first time an email is sent, macOS asks *"Student Batch Mailer wants to control Microsoft Outlook"*. Click **Allow**. If it was denied by accident, re-enable it under *System Settings → Privacy & Security → Automation*.

No other installation is needed — Node, pnpm and the repository are only required to build the app, not to run it.

### Troubleshooting
- **`node: command not found` during `pnpm install`** – Node.js isn't installed or not on your PATH; see Prerequisites.
- **`Electron failed to install correctly`** – the Electron binary wasn't downloaded. Run `pnpm exec install-electron`, or do a clean reinstall:
  ```
  rm -rf node_modules ~/Library/Caches/electron
  pnpm install
  ls node_modules/electron/dist   # should contain Electron.app
  ```
- **`Cannot find module ...` when starting** – a previous install was interrupted. `rm -rf node_modules && pnpm install`.

## Platform support
- The distributed build targets macOS on Apple Silicon (`--platform=darwin --arch=arm64`), i.e., M1/M2/M3 machines. Intel Macs require rebuilding with `pnpm run dist -- --arch=x64` or similar.
- Microsoft Outlook for macOS must be installed and allowed to run AppleScript, because sending relies on `osascript outlook.scpt`. Other mail clients are not supported.
- Uploads are cached under `~/Library/Application Support/student-batch-mailer/upload-cache/` for the duration of each session so pathless drag-and-drop files can be attached.

## Sent logs
After every send the app automatically writes a timestamped HTML report (e.g. `sent-2026-09-13_14-05-32.html`) to `~/Library/Application Support/student-batch-mailer/sent-logs/`. Each report lists:

- the emails that were sent (time, student, address, attachment),
- students that were matched but unchecked,
- students in the roster without a matching file,
- files that didn't match any student.

The reports share one stylesheet, `sent-logs/css/style.css`, which is copied from [`report/style.css`](./report/style.css) on every app start — edit that file to restyle all reports. Use the **Open sent logs folder** button in the app to view them. Saved message templates live in the same Application Support folder as `templates.json`.

### Quick self-checks
- `uname -m` should print `arm64` to use the prebuilt binary. If it prints `x86_64`, rebuild with `pnpm run dist -- --arch=x64`.
- `osascript -e 'id of application "Microsoft Outlook"'` should print the Outlook bundle id (e.g., `com.microsoft.Outlook`) to confirm Outlook is installed and AppleScript-accessible.

## Sample data

### Quick try-out (included in the repo)
The [`sample-data/`](./sample-data/) folder contains everything needed to test the app in a minute:

- Roster: [`sample-data/student-roster.xlsx`](./sample-data/student-roster.xlsx) (8 students)
- Feedback PDFs to drag & drop: [`sample-data/feedback-files/`](./sample-data/feedback-files/)

Load the roster, drag the PDFs onto the drop zone, and all 8 files should match. The roster's addresses are Gmail `+` aliases (`vinzzz81+emmaanderson@gmail.com`, …) so test emails all arrive in one inbox — replace `vinzzz81` with your own Gmail name to receive them yourself. See [`sample-data/README.md`](./sample-data/README.md) for step-by-step instructions.

### Larger generated set
To create a 60-student sample set that mails to your own Gmail inbox, run:

```
python3 scripts/create_sample_set.py your.gmail+alias@gmail.com
```

- Generates 60 PDFs in `sample-set/feedback-files/`
- Generates a matching `sample-set/student-sampleset.xlsx`
- Emails are based on Gmail `+` addressing so all messages land in your inbox
