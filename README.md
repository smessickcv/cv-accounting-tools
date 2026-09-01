# Cardone Ventures — Internal Tools Hub

Static GitHub Pages site hosting internal finance & accounting tools.

**Live:** https://smessickcv.github.io/cv-accounting-tools/

---

## Branch model

**`master` is both the default branch and the branch GitHub Pages deploys.** Push
to `master` and it is live in about a minute.

`main` is a stale leftover: it diverged long ago, does not contain `index.html`,
and is not deployed. It only still exists because two files were once uploaded
there by mistake (`fam_converter.html` and `xml to xlsx converter.html`, on
2026-07-13) and never went live from that push — both are on `master` now.
Nothing on `main` is missing from `master`. It can be deleted whenever you want.

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
| `ns-import` | `ns-import-tool.html` | NetSuite cash sale import — Shopify / Square / Stripe tabs |
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
| `401k-converter` | `401k-converter.html` | 401k payroll converter — **stores SSN/DOB locally, see Data privacy** |
| `register-builder` | `register-builder.html` | Weekly cash register builder — **ships empty, needs the settings bundle, see below** |
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

`ns-import-tool.html` additionally carries a `VERSION` constant near the top that
drives its on-screen badges. It logs a console warning if it ever falls out of
step with `TOOL_VERSION`, so bump both.

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

**`401k-converter.html` holds employee PII.** Its saved roster (`401k_roster`)
contains SSNs, dates of birth and hire dates. That data never leaves the browser
— it is not transmitted anywhere and none of it is in this repo — but it does
persist in `localStorage` on whatever machine the tool is used on. Use it on a
company device, and use the *Clear Roster from Storage* button on a shared one.

**`register-builder.html` ships deliberately empty.** Its coding rules and code
list are not in this repo. The original carried an 88KB rule corpus — roughly
1,800 bank transaction descriptions naming vendors and individual payees — plus
four real bank account numbers, none of which belongs in a public repository.

The published file starts with no rules and no codes, shows a banner saying so,
and codes every row `REVIEW` until someone loads the data. Everything lives in a
single `Register Builder settings.json`, distributed through **#internal-tools**
or SharePoint, and loaded with **Settings → Import settings**. That one file
carries rules, codes, sources (including account numbers), the export layout,
flags, learned corrections, hidden rules and the learn toggle — so **Export
settings** is also a complete backup. Treat the bundle as confidential; it is
gitignored here so it cannot be committed by accident.

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

**Accepted — reviewed and deliberately not changed**

- `ns-import-tool.html` keeps `state.shopify.gateway` when a new orders export is
  dropped in, so a stale gift-card map can carry over between files.
- Stripe line `Qty` is always `1`. `parseStripeProducts` strips the leading count
  off `"3 TeamWork Book"` but does not use it; the even rate split keeps the
  transaction total correct.
- `resolveItemBySku` falls back to progressively stripping `-` segments and then
  accepts `name.startsWith(base) && !name.includes(' ')`. A wrong match suppresses
  the "unresolved SKU" warning, but this has not caused a problem in practice.
- Shopify always emits a shipping row, including at $0.
- No tool sanitises leading `=`, `+`, `-` or `@` in exported CSV fields. Low risk
  while the data originates in NetSuite; it is a formula-injection vector only if
  an export is opened in Excel rather than imported.

**Fixed 2026-08-31 — ns-import-tool v1.2**

- Stripe read the charge-currency columns (`Amount` / `Amount Refunded`, cols C
  and D). It now reads the settlement-converted columns (`Converted Amount` /
  `Converted Amount Refunded`, cols G and H) via `stripeAmount()` /
  `stripeRefund()`, so a 1,703.45 MXN charge imports as its 99.01 USD equivalent.
  `applyStripeRule()` and the $497 threshold use the converted figure too — the
  sample MXN charge was previously judged above the threshold on its peso value.
- Any charge whose `Currency` (col E) is not USD is now flagged **ALTERNATE
  CURRENCY**, forced out of auto-include into manual review, and counted in its
  own stat badge. A **Currency** column was added to the Stripe review table
  showing the currency and, for non-USD, the pre-conversion amount. The exported
  CSV columns are unchanged — the flag is a review-stage check only.
- The Stripe Cash Sale Date auto-detect silently overwrote a manually entered
  date. It now prompts when auto-detection disagrees, remembers the answer so
  regenerating does not re-ask, and keeps a manual date on multi-day files
  instead of clearing it.
- The GitHub self-update block (`GH_OWNER`/`GH_REPO`/`GH_TOKEN`, `checkForUpdates`)
  is commented out and its Setup button removed. It never worked — the constants
  were still `YOUR_...` placeholders — and it called for embedding a PAT in a file
  served from a public repo. `version-check.js` covers this without credentials.
- Version strings were disagreeing: `VERSION` said `1.0.0` while `TOOL_VERSION`
  and the navbar badge said `1.1`, so Setup reported a different version from the
  header. Both badges now render from `VERSION`, all four sources read `1.2`, and
  `init()` logs a console warning on any future drift.
- Shopify money fields (`Discount Amount`, `Shipping`, `Lineitem price`, and the
  gateway report's `Gross payments`) used bare `parseFloat`, which reads
  `"1,234.56"` as `1`. They now use `parseMoney()`, which was also hardened to
  never return `NaN`.
- The inline entry panels emitted bare `cust-id-0` / `sku-id-0` element IDs from
  both the Shopify and Stripe tabs. With results open on both, `getElementById`
  returned the wrong tab's inputs. IDs are now namespaced per platform.

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
