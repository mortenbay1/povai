/* global Office, Word, OoxmlToHtml */

const STORAGE_KEY = "wordtowp_site_url";

let wpUrlInput;
let saveUrlBtn;
let urlStatus;
let publishBtn;
let publishText;
let publishStatus;

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        initUI();
    }
});

function initUI() {
    wpUrlInput = document.getElementById("wp-url");
    saveUrlBtn = document.getElementById("save-url-btn");
    urlStatus = document.getElementById("url-status");
    publishBtn = document.getElementById("publish-btn");
    publishText = document.getElementById("publish-text");
    publishStatus = document.getElementById("publish-status");

    // Load saved URL
    const savedUrl = localStorage.getItem(STORAGE_KEY);
    if (savedUrl) {
        wpUrlInput.value = savedUrl;
        publishBtn.disabled = false;
        setStatus(urlStatus, "Saved", "success");
    }

    saveUrlBtn.addEventListener("click", saveUrl);
    publishBtn.addEventListener("click", publish);
    wpUrlInput.addEventListener("keydown", (e) => {
        if (e.key === "Enter") saveUrl();
    });
}

function saveUrl() {
    const url = normalizeUrl(wpUrlInput.value.trim());
    if (!url) {
        setStatus(urlStatus, "Please enter a valid URL.", "error");
        return;
    }
    wpUrlInput.value = url;
    localStorage.setItem(STORAGE_KEY, url);
    publishBtn.disabled = false;
    setStatus(urlStatus, "Saved", "success");
}

function normalizeUrl(url) {
    if (!url) return "";
    if (!/^https?:\/\//i.test(url)) {
        url = "https://" + url;
    }
    return url.replace(/\/+$/, "");
}

async function publish() {
    const wpUrl = localStorage.getItem(STORAGE_KEY);
    if (!wpUrl) {
        setStatus(publishStatus, "Please save a WordPress URL first.", "error");
        return;
    }

    publishBtn.disabled = true;
    publishText.textContent = "Reading document...";
    setStatus(publishStatus, "", "info");

    try {
        // Step 1: Read document OOXML
        let ooxmlContent;
        await Word.run(async (context) => {
            const body = context.document.body;
            const ooxml = body.getOoxml();
            await context.sync();
            ooxmlContent = ooxml.value;
        });

        if (!ooxmlContent) {
            throw new Error("Could not read document content.");
        }

        // Step 2: Convert OOXML to HTML
        publishText.textContent = "Converting...";
        const html = OoxmlToHtml.convert(ooxmlContent);

        if (!html || !html.trim()) {
            throw new Error("Document appears to be empty.");
        }

        // Step 3: Extract title from first heading
        const title = extractTitle(html);

        // Step 4: Remove the title heading from the body content so it isn't duplicated
        const bodyHtml = removeTitleFromHtml(html, title);

        // Step 5: Build the WordPress URL with content as query parameters
        // WordPress post-new.php accepts ?post_title=...&content=...
        publishText.textContent = "Opening WordPress...";
        const newPostUrl = buildWordPressUrl(wpUrl, title, bodyHtml);

        // Open in the system browser. If the user is already logged in,
        // they land directly on the new post editor with content pre-filled.
        // If not logged in, WP redirects to login, and after login redirects
        // back to the pre-filled new post page.
        Office.context.ui.openBrowserWindow(newPostUrl);

        setStatus(
            publishStatus,
            'Opened WordPress with your post "' + title + '". Review and click Publish in WordPress.',
            "success"
        );
    } catch (err) {
        setStatus(publishStatus, "Error: " + err.message, "error");
    } finally {
        publishBtn.disabled = false;
        publishText.textContent = "Publish to WordPress";
    }
}

function buildWordPressUrl(wpUrl, title, content) {
    // Build the post-new.php URL with pre-filled title and content
    const newPostPath = "/wp-admin/post-new.php";
    const params = new URLSearchParams({
        post_title: title,
        content: content,
    });
    const newPostUrl = wpUrl + newPostPath + "?" + params.toString();

    // Wrap in wp-login.php?redirect_to=... so that:
    // - If logged in: WP skips login and redirects straight to post-new.php
    // - If not logged in: WP shows login, then redirects to post-new.php after
    const loginUrl =
        wpUrl +
        "/wp-login.php?redirect_to=" +
        encodeURIComponent(newPostPath + "?" + params.toString());

    return loginUrl;
}

function extractTitle(html) {
    const match = html.match(/<h[12][^>]*>(.*?)<\/h[12]>/i);
    if (match) {
        return match[1].replace(/<[^>]+>/g, "").trim();
    }
    const pMatch = html.match(/<p[^>]*>(.*?)<\/p>/i);
    if (pMatch) {
        const text = pMatch[1].replace(/<[^>]+>/g, "").trim();
        if (text.length > 60) return text.substring(0, 57) + "...";
        return text;
    }
    return "Untitled Post";
}

function removeTitleFromHtml(html, title) {
    // Remove the first h1 or h2 that matches the extracted title
    // so it doesn't appear twice (once as WP title, once in body)
    return html.replace(/<h[12][^>]*>.*?<\/h[12]>\s*/i, "");
}

function setStatus(el, message, type) {
    el.textContent = message;
    el.className = "status-text";
    if (type) el.classList.add("status-" + type);
}
