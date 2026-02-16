# Word to WordPress — Office Add-in

A Word add-in that creates a draft WordPress post from your document via the REST API.

---

## For Users: How to Install

### Step 1 — Download the manifest file

Download the `manifest.xml` file from this page (click it, then click the Download button).

### Step 2 — Add it to Word

**Word for Windows:**
1. Open Word
2. Go to **Insert** → **Get Add-ins** (or **My Add-ins**)
3. Click **Upload My Add-in**
4. Browse to the `manifest.xml` file you downloaded and click **OK**

**Word for Mac:**
1. Open Finder and press **Cmd+Shift+G**
2. Paste this path (replace `<username>` with your Mac username):
   ```
   /Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef
   ```
3. If the `wef` folder doesn't exist, create it
4. Copy the `manifest.xml` file into this `wef` folder
5. Restart Word (quit and reopen), then open a document
6. Go to **Home** tab → click **Add-ins** → your add-in should appear

**Word Online (Office 365):**
1. Open a document in Word Online
2. Go to **Insert** → **Office Add-ins**
3. Click **Upload My Add-in** (top right)
4. Browse to the `manifest.xml` file you downloaded

### Step 3 — You're done!

A **"Publish to WP"** button appears in the **Home** tab of the Word ribbon. Click it to open the add-in panel.

---

## How to Use

1. **First time only:** Create an Application Password in WordPress:
   - Go to **Users** → **Profile** → **Application Passwords**
   - Create a new app password (e.g. "Word Add-in") and copy it
2. Enter your **WordPress site URL**, **username**, and **Application Password** in the add-in, then click **Save**. These are stored locally so you don't need to enter them again.
3. Click **Publish to WordPress**. The add-in creates a draft post directly via the WordPress REST API and opens it in your browser.
4. Review the post in WordPress and click **Publish** when ready.
5. To change your password (e.g. after rotating credentials), click **Update Application Password**, enter the new password, and click **Save**.

## What gets transferred

- Headings (H1–H6)
- Paragraphs with formatting (bold, italic, underline, strikethrough)
- Ordered and unordered lists
- Hyperlinks
- Block quotes
- The post title is extracted from the first heading in your document

## Limitations

- Images in the Word document are not transferred (upload them separately via the WordPress Media Library)
- Complex tables and custom styles may not convert perfectly
- Application Passwords require WordPress 5.6 or later

---

## For Developers: Hosting Setup

The add-in files (HTML, CSS, JS) must be hosted on an HTTPS server. The easiest free option is **GitHub Pages**.

### Option A — GitHub Pages (recommended)

1. Fork or clone this repository
2. In your GitHub repo, go to **Settings** → **Pages** → set source to **main branch**
3. Your add-in is now live at `https://YOURUSERNAME.github.io/wordtowp/`
4. Open `manifest.xml` and replace all instances of `YOURDOMAIN.github.io/wordtowp` with your actual GitHub Pages URL
5. Share the updated `manifest.xml` with your users

### Option B — Netlify / Vercel / Cloudflare Pages

1. Connect your Git repo to any of these services
2. They provide automatic HTTPS on a custom domain
3. Update the URLs in `manifest.xml` to point to your hosted domain

### Option C — Any HTTPS web server

Upload these files to any web server that supports HTTPS:
- `taskpane.html`
- `taskpane.css`
- `taskpane.js`
- `ooxml-to-html.js`
- `assets/icon-16.png`, `icon-32.png`, `icon-80.png`

Then update the URLs in `manifest.xml` accordingly.

### For organizations (zero-click install for end users)

A Microsoft 365 admin can deploy the add-in to all users automatically:
1. Go to **Microsoft 365 Admin Center** → **Settings** → **Integrated Apps**
2. Upload the `manifest.xml`
3. Assign to users or groups
4. The add-in appears in everyone's Word automatically — no action needed from end users
