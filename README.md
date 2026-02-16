# Word to WordPress — Office Add-in

A Word add-in that publishes your document as a new WordPress post with one click.

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
1. Open Word
2. Go to **Insert** → **Add-ins** → **My Add-ins**
3. Click the dropdown arrow and choose **Upload My Add-in**
4. Browse to the `manifest.xml` file you downloaded and click **Upload**

**Word Online (Office 365):**
1. Open a document in Word Online
2. Go to **Insert** → **Office Add-ins**
3. Click **Upload My Add-in** (top right)
4. Browse to the `manifest.xml` file you downloaded

### Step 3 — You're done!

A **"Publish to WP"** button appears in the **Home** tab of the Word ribbon. Click it to open the add-in panel.

---

## How to Use

1. **First time only:** Enter your WordPress site URL (e.g., `https://mysite.com`) and click **Save**.
2. Click **Publish to WordPress**.
3. WordPress opens in your browser with the post title and content already filled in.
4. If you're not logged in, log in first — you'll be redirected to the new post automatically.
5. Review the post and click **Publish** in WordPress.

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
- Very long documents may hit browser URL length limits

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
