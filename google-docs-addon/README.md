# WordToWP — Google Docs add-on (Apps Script)

Sidebar add-on that mirrors the Word add-in: WordPress publishing and paragraph styles (Rubrik, Underrubrik, etc.).

## Prerequisites

- [Node.js](https://nodejs.org/) (LTS is fine)
- Google account with access to the Apps Script project you created (step 2)

**Avoid `npm install -g @google/clasp`** if you see permission errors under `/usr/local/`. Use the **local install** below instead (everything goes into `node_modules/` in this folder).

## Project layout

- **`src/`** — the only files **clasp** uploads (`Code.gs`, `Sidebar.html`, `appsscript.json`).  
  `npm` / `node_modules` stay **outside** `src/`, so `clasp push` never scans them (fixes duplicate-name errors from `node_modules`).

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
   # Edit .clasp.json: set scriptId and ensure "rootDir": "src"
   ```

   If you already had `.clasp.json` with `"rootDir": "."`, change it to **`"rootDir": "src"`** after pulling this layout.

4. Push files to Google:

   ```bash
   npx clasp push
   ```

   Or: `npm run push`

5. In the Apps Script editor, open **Deploy** → **Test deployments** → select **Editor Add-on** (or install the test add-on for Docs). Open a **Google Doc** → **Extensions** → your add-on → **Open WordToWP**.

## Git

- Commit everything **except** `.clasp.json` if you prefer not to store the Script ID in the repo (optional). Otherwise commit `.clasp.json` so the team shares one project ID.

Add to the parent repo’s `.gitignore` (optional):

```
google-docs-addon/.clasp.json
```

## Implemented (M2 + M3 + formatting)

- **Settings:** WordPress URL, username, and application password stored in `PropertiesService` (per user). **Save** / **Update application password** match the Word add-in flow.
- **Publish:** Converts the document body to HTML (paragraphs, headings h1–h6, bold, italic, simple lists, **blockquote** for paragraphs with **Subtitle** style) and creates a **draft** via `POST /wp-json/wp/v2/posts`. Opens the WordPress post editor in a new tab.
- **Styles (sidebar):** **Rubrik** (Title), **Mellemrubrik** (Heading 2), **Brødtekst** (Normal), **Underrubrik** (Normal + bold selection), **Citat** (inserts selected text as a new paragraph with **Subtitle** style; Google Docs has no block-quote enum, so this maps to `<blockquote>` on export). Avoid using the built-in Subtitle style for non-quotes if you need those paragraphs to stay regular body text in WordPress.

## Sharing the add-on without the public Workspace Marketplace

You do **not** have to publish to the public [Google Workspace Marketplace](https://workspace.google.com/marketplace) for a small team or pilot.

1. **Test deployment (typical for collaborators)**  
   In Apps Script: **Deploy** → **Test deployments** → create or select a **Test** deployment of the **Editor Add-on** (Google Docs). Add each user’s Google account as a **test user** (or use the project’s test-install flow your Google Cloud / Apps Script version shows). Those users install the add-on from the **test install link** or **Extensions** → **Add-ons** → **Get add-ons** → **My add-ons** / internal testing, depending on the UI. They use the add-on like a normal install; it is not listed publicly.

2. **Google Workspace domain only (organization)**  
   A Google Workspace **admin** can deploy and assign add-ons for your domain without a public listing (domain-wide or allowlisted apps). Exact steps depend on your admin console version; search for “domain install Google Workspace add-on” in your admin help.

3. **Private / unlisted Marketplace listing**  
   Some organizations use a **private** Marketplace app (visible only to your domain). That still goes through Marketplace infrastructure but is not “public to the world.”

For day-to-day development, **Test deployments** plus listed test users is usually enough until you want a formal listing.

## Next steps

- Richer HTML (links, tables) to align with the Word add-in where needed.
