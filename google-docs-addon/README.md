# WordToWP — Google Docs add-on (Apps Script)

Sidebar add-on that mirrors the Word add-in: WordPress publishing and formatting (wired up in later steps).

## Prerequisites

- [Node.js](https://nodejs.org/) (LTS is fine)
- Google account with access to the Apps Script project you created (step 2)

**Avoid `npm install -g @google/clasp`** if you see permission errors under `/usr/local/`. Use the **local install** below instead (everything goes into `node_modules/` in this folder).

## One-time: link this repo to your script

1. From the `google-docs-addon/` folder, install dependencies (installs `clasp` locally):

   ```bash
   cd google-docs-addon
   npm install
   ```

2. Log in (uses the local `clasp`):

   ```bash
   npx clasp login
   ```

3. Copy the example config and paste your **Script ID** from [script.google.com](https://script.google.com) → Project Settings:

   ```bash
   cp .clasp.json.example .clasp.json
   # Edit .clasp.json — replace YOUR_SCRIPT_ID_FROM_SCRIPT_GOOGLE_COM
   ```

4. Push files to Google:

   ```bash
   npx clasp push
   ```

   Or: `npm run push`

5. In the Apps Script editor, open **Deploy** → **Test deployments** → select **Editor Add-on** (or install the test add-on for Docs). Open a **Google Doc** → **Extensions** → your add-on → **Open WordToWP**.

## What step 3 includes

- `appsscript.json` — V8 runtime, OAuth scopes for current document, sidebar UI, and external HTTP (for WordPress later).
- `Code.gs` — `onOpen` menu, `showSidebar`, and `getActiveDocumentInfo()` sample.
- `Sidebar.html` — sidebar with **Load document preview** to verify `google.script.run` ↔ server.

## Git

- Commit everything **except** `.clasp.json` if you prefer not to store the Script ID in the repo (optional). Otherwise commit `.clasp.json` so the team shares one project ID.

Add to the parent repo’s `.gitignore` (optional):

```
google-docs-addon/.clasp.json
```

## Next steps (M2+)

- Save WordPress URL / username / app password in `PropertiesService`.
- Implement publish + HTML export + formatting buttons.
