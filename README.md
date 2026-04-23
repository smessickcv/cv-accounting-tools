# Cardone Ventures — Internal Tools Hub

A GitHub Pages site hosting internal finance & accounting tools for the Cardone Ventures team.

## Live Site

`https://smessickcv.github.io/<repo-name>/`

---

## One-Time GitHub Setup

1. Create a new **public** repository on your personal GitHub account (e.g. `cv-tools`).
2. Push all files in this folder to the `main` branch:

```bash
git init
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/smessickcv/cv-tools.git
git push -u origin main
```

3. In the repo on GitHub → **Settings → Pages**:
   - Source: **Deploy from a branch**
   - Branch: `main` / `/ (root)`
   - Click **Save**

GitHub Pages will be live within ~1 minute. The URL will be:
`https://smessickcv.github.io/cv-tools/`

---

## Pushing Tool Updates

1. Edit the tool HTML file and increment its version constant at the bottom:
   ```html
   <script>const TOOL_KEY='ns-import'; const TOOL_VERSION='1.1';</script>
   ```
2. Update `version.json` to match:
   ```json
   "ns-import": "1.1"
   ```
3. Commit and push — GitHub Pages deploys automatically within ~1 minute.
4. Post the update in **#internal-tools** so users know to reload.

---

## Files

| File | Purpose |
|------|---------|
| `index.html` | Hub landing page |
| `version.json` | Version manifest — controls update banners |
| `version-check.js` | Shared script injected into every tool |
| `BofA_to_NetSuite.html` | BofA → NetSuite bank import |
| `CC_Reconciliation_Tool-1.html` | Credit card reconciliation |
| `CC_Statement_Recon.html` | CC statement reconciliation |
| `DueToFrom_Master_Generator.html` | Due-to/from master generator |
| `IntercompanyRecon_SharePoint_21.html` | Intercompany recon |
| `Payroll JE Generator.html` | Payroll JE (subsidiaries) |
| `payroll_je_tool.html` | Payroll JE (10X HVAC) |
| `loan-interest-calculator_17_review.html` | Loan interest calculator |
| `ns-import-tool.html` | NetSuite cash sale import |
| `rev share je generator.html` | Rev share JE generator |
| `shopify-to-netsuite_5.html` | Shopify → NetSuite |

---

## Data Privacy

All CSV imports are processed locally in the browser via the File API. No customer data is transmitted to GitHub servers. `localStorage` is used only for non-sensitive UI preferences (theme, in-progress work).
