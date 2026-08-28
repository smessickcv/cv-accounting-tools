# Cardone Ventures — Internal Tools Hub

Static GitHub Pages site hosting internal finance & accounting tools.

**Live:** https://smessickcv.github.io/cv-accounting-tools/

---

## ⚠ Branch model — read this first

**GitHub Pages deploys from `master`. `master` is the live branch.**

`main` is still the repo's *default* branch, so the GitHub web UI ("Add files via
upload", the online editor, new PRs) will target `main` unless you change it —
and anything landed there **will not go live**. This has already happened once:
`fam_converter.html` and `xml to xlsx converter.html` were uploaded to `main` on
2026-07-13 and never deployed from that push.

`main` and `master` have diverged and `main` does not contain `index.html`.

- Always commit to `master`.
- If you use the GitHub web UI, confirm the branch selector says `master`.
- Fix at the source when convenient: Settings → General → Default branch → `master`.

---

## Repository layout

| File | Purpose |
|------|---------|
| `index.html` | Hub landing page — tool cards are defined in the `LOCAL_TOOLS` array inside this file |
| `version.json` | Version manifest that drives the "update available" banners |
| `version-check.js` | Shared banner script, included by most tool pages |
| `*.html` | One self-contained tool per file |
| `.gitignore` | Keeps local-only data (`_maps/`, employee maps, `.py`, `.xlsx`, superseded tool versions) out of the public repo |

Everything is a single self-contained HTML file. There is no build step — a push
to `master` is live in about a minute.

---

## Tool registry

`key` is the identifier that ties a tool's `TOOL_VERSION`, its `version.json`
entry, and its `index.html` card together. All three must agree. Every tool below is wired into all three.

| Key | File | Notes |
|-----|------|-------|
| `ns-import` | `ns-import-tool.html` | NetSuite cash sale import |
| `bofa-netsuite` | `BofA_to_NetSuite.html` | BofA → NetSuite bank import |
| `cc-recon` | `CC_Reconciliation_Tool-1.html` | Credit card reconciliation |
| `cc-statement-recon` | `CC_Statement_Recon.html` | CC statement reconciliation |
| `due-to-from` | `DueToFrom_Master_Generator.html` | Due-to/from master generator |
| `intercompany-recon` | `IntercompanyRecon_SharePoint_21.html` | Subsidiary intercompany recon |
| `payroll-je-subsidiary` | `Payroll JE Generator.html` | Payroll JE — subsidiaries |
| `payroll-je-hvac` | `payroll_je_tool.html` | Payroll JE — 10X HVAC |
| `payroll-je-generator` | `payroll-je-generator-hvac.html` | Payroll JE — 10X HVAC (Pyodide) |
| `loan-calculator` | `loan-interest-calculator_17_review.html` | Loan interest calculator |
| `rev-share-je` | `rev share je generator.html` | Rev share JE generator |
| `shopify-netsuite` | `shopify-to-netsuite_5.html` | Shopify → NetSuite |
| `importrange-builder` | `importrange-query-builder.html` | IMPORTRANGE query builder |
| `netsuite-processor` | `netsuite-export-processor.html` | NetSuite export → Due_ToFrom_Master (Pyodide) |
| `ap-dashboard` | `AP_Dashboard.html` | AP dashboard — **not local-only, see Data privacy** |
| `fam-converter` | `fam_converter.html` | FAM XML → XLSX converter |
| `fulfillment-builder` | `fulfillment_builder.html` | Item fulfillment builder |
| *(not listed)* | `xml_to_xlsx_converter.html` | Duplicate of `fam_converter.html`; deliberately absent from the hub — see Known issues |

---

## Shipping an update

A tool's version lives in **three** places. Miss one and users either get a red
"update available" banner for a version that does not exist, or get no banner at
all for one that does.

1. **The tool file** — bump the constant near the bottom, and the `v1.x` badge in
   the page header if it has one:
   ```html
   <script>const TOOL_KEY='ns-import'; const TOOL_VERSION='1.1';</script>
   ```
2. **`version.json`** — same value, and bump `"updated"` to today:
   ```json
   "ns-import": "1.1"
   ```
3. **`index.html`** — the `version:` field on that tool's entry in `LOCAL_TOOLS`.

Then commit to `master`, push, and post in **#internal-tools** so people reload.

### How the banner logic works

`version-check.js` fetches `version.json` and compares it to the page's
`TOOL_VERSION`. If the manifest is **higher**, it injects a red sticky banner.
So `version.json` must never run ahead of what is actually deployed.

`index.html` runs its own copy of the same comparison to put a `↑` on each card,
which is why the card's `version:` field has to be kept in sync too.

---

## Data privacy

Most tools parse uploads entirely in the browser via the File API — nothing
leaves the machine, and `localStorage` holds only UI preferences and saved
mappings.

**`AP_Dashboard.html` is the exception.** It signs the user into Google, requests
the full `https://www.googleapis.com/auth/spreadsheets` scope, and reads/writes a
live Google Sheet over the network. Its OAuth client ID and the target sheet ID
are hardcoded in this public repo. Treat it differently from the offline tools,
and do not repeat the blanket "nothing is sent to a server" claim about it.

---

## Known issues

Tracked so they are not rediscovered.

**Open**

- `xml_to_xlsx_converter.html` is a byte-identical copy of `fam_converter.html`
  (same MD5, same `<title>`). The file is still on `master` so nothing that
  deep-links it breaks, but it is no longer advertised on the hub. Either drop
  in the real XML → XLSX tool or delete the file.
- Commit `de48bac` is titled "v1.3" but `DueToFrom_Master_Generator.html` still
  reports `1.2` in both its header badge and `TOOL_VERSION`. The manifest has
  been set to `1.2` to match the file that actually shipped. Confirm which is
  correct before the next bump.
- `AP_Dashboard.html` requests the full `.../auth/spreadsheets` scope. Google has
  no way to narrow the Sheets API to a single file without switching to a
  `drive.file` picker flow, so this is accepted rather than fixed — but it means
  the tool can read and write every spreadsheet the signed-in user owns.
- No tool sanitises leading `=`, `+`, `-` or `@` in exported CSV fields. Low risk
  while the data originates in NetSuite, but it is a formula-injection vector if
  a CSV is ever opened in Excel instead of imported.

**Fixed 2026-08-28**

- `version.json` listed `importrange-builder` and `netsuite-processor` at `1.1`
  while the deployed files were `1.0`, firing a false "update available" banner
  on both. Manifest corrected to `1.0`.
- `due-to-from`, `rev-share-je` and `shopify-netsuite` were ahead of the
  manifest; the `ns-import` and `payroll-je-hvac` cards in `index.html` were
  behind their files. All three sources now agree for all 17 tools.
- `AP_Dashboard`, `fam_converter` and `fulfillment_builder` had no version
  wiring at all — added `TOOL_KEY`/`TOOL_VERSION`, the `version-check.js` tag,
  and manifest entries.
- `AP_Dashboard.html` never cleared its 45-second refresh interval on sign-out,
  and stacked a second one on every re-sign-in. Now managed through
  `startAutoRefresh()` / `stopAutoRefresh()`.
- `AP_Dashboard.html` called `signIn()` from `gFetch` on a 401. That fires from
  the background timer, so the browser blocked the popup and the dashboard died
  silently after roughly an hour. Replaced with `handleSessionExpired()`, which
  stops the timer and returns the user to the sign-in card with an explanation,
  so re-auth happens on a real click.
- `AP_Dashboard.html` read a fixed `A3:Z1000` range in three places. Now a single
  `DATA_RANGE` constant at `A3:ZZ10000`.
- `AP_Dashboard.html` threw `Cannot read properties of null` on every page load
  from a leftover `#config-banner` reference to an element that no longer exists.
  Because it sat above `initAuth()` in the `load` handler, it aborted the rest of
  init on every visit. Reference removed.
- `index.html` styled `#cv-global-update-banner` while `version-check.js` creates
  `#cv-update-banner`, and `index.html` never loaded that script. Dead CSS and an
  unused `anyUpdates` variable removed.


---

Internal use only.
