/* global Office, Word, OoxmlToHtml */

const STORAGE_KEY_URL = "wordtowp_site_url";
const STORAGE_KEY_USER = "wordtowp_username";
const STORAGE_KEY_APP_PASS = "wordtowp_app_password";

let wpUrlInput;
let wpUsernameInput;
let wpAppPasswordInput;
let passwordSection;
let passwordSavedSection;
let updatePasswordBtn;
let saveBtn;
let settingsStatus;
let publishBtn;
let publishText;
let publishStatus;
let markTitleBtn;
let markH2Btn;
let markNormalBtn;
let markUnderrubrikBtn;
let markBlockquoteBtn;
let formatStatus;

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        initUI();
    }
});

function initUI() {
    wpUrlInput = document.getElementById("wp-url");
    wpUsernameInput = document.getElementById("wp-username");
    wpAppPasswordInput = document.getElementById("wp-app-password");
    passwordSection = document.getElementById("password-section");
    passwordSavedSection = document.getElementById("password-saved-section");
    updatePasswordBtn = document.getElementById("update-password-btn");
    saveBtn = document.getElementById("save-btn");
    settingsStatus = document.getElementById("settings-status");
    publishBtn = document.getElementById("publish-btn");
    publishText = document.getElementById("publish-text");
    publishStatus = document.getElementById("publish-status");
    markTitleBtn = document.getElementById("mark-title-btn");
    markH2Btn = document.getElementById("mark-h2-btn");
    markNormalBtn = document.getElementById("mark-normal-btn");
    markUnderrubrikBtn = document.getElementById("mark-underrubrik-btn");
    markBlockquoteBtn = document.getElementById("mark-blockquote-btn");
    formatStatus = document.getElementById("format-status");

    loadSavedSettings();

    markTitleBtn.addEventListener("click", () => applyStyle("Title"));
    markH2Btn.addEventListener("click", () => applyStyle("Heading 2"));
    markNormalBtn.addEventListener("click", () => applyStyle("Normal"));
    markUnderrubrikBtn.addEventListener("click", applyUnderrubrik);
    markBlockquoteBtn.addEventListener("click", applyBlockquote);
    saveBtn.addEventListener("click", saveSettings);
    updatePasswordBtn.addEventListener("click", showUpdatePassword);
    publishBtn.addEventListener("click", publish);
    wpUrlInput.addEventListener("keydown", (e) => {
        if (e.key === "Enter") saveSettings();
    });
    wpUsernameInput.addEventListener("keydown", (e) => {
        if (e.key === "Enter") saveSettings();
    });
    wpAppPasswordInput.addEventListener("keydown", (e) => {
        if (e.key === "Enter") saveSettings();
    });
}

function loadSavedSettings() {
    const savedUrl = localStorage.getItem(STORAGE_KEY_URL);
    const savedUser = localStorage.getItem(STORAGE_KEY_USER);
    const savedPass = localStorage.getItem(STORAGE_KEY_APP_PASS);

    if (savedUrl) wpUrlInput.value = savedUrl;
    if (savedUser) wpUsernameInput.value = savedUser;

    if (savedPass) {
        passwordSection.classList.add("hidden");
        passwordSavedSection.classList.remove("hidden");
        publishBtn.disabled = false;
        setStatus(settingsStatus, "Settings saved", "success");
    } else {
        passwordSection.classList.remove("hidden");
        passwordSavedSection.classList.add("hidden");
        publishBtn.disabled = true;
    }
}

function showUpdatePassword() {
    localStorage.removeItem(STORAGE_KEY_APP_PASS);
    wpAppPasswordInput.value = "";
    passwordSection.classList.remove("hidden");
    passwordSavedSection.classList.add("hidden");
    publishBtn.disabled = true;
    wpAppPasswordInput.focus();
    setStatus(settingsStatus, "Enter new application password and click Save", "info");
}

function saveSettings() {
    const url = normalizeUrl(wpUrlInput.value.trim());
    const username = wpUsernameInput.value.trim();
    const appPassword = wpAppPasswordInput.value.trim().replace(/\s/g, "");

    if (!url) {
        setStatus(settingsStatus, "Please enter a valid WordPress URL.", "error");
        return;
    }
    if (!username) {
        setStatus(settingsStatus, "Please enter your WordPress username.", "error");
        return;
    }
    if (!appPassword) {
        setStatus(settingsStatus, "Please enter a valid application password.", "error");
        return;
    }

    wpUrlInput.value = url;
    localStorage.setItem(STORAGE_KEY_URL, url);
    localStorage.setItem(STORAGE_KEY_USER, username);
    localStorage.setItem(STORAGE_KEY_APP_PASS, appPassword);

    passwordSection.classList.add("hidden");
    passwordSavedSection.classList.remove("hidden");
    wpAppPasswordInput.value = "";
    publishBtn.disabled = false;
    setStatus(settingsStatus, "Settings saved", "success");
}

async function applyStyle(styleName) {
    setStatus(formatStatus, "", "info");
    try {
        await Word.run(async (context) => {
            const range = context.document.getSelection();
            range.paragraphs.load("items");
            await context.sync();
            const paragraphs = range.paragraphs.items;
            if (paragraphs.length === 0) {
                setStatus(formatStatus, "Select text first, then click a style.", "info");
                return;
            }
            for (const para of paragraphs) {
                para.style = styleName;
            }
            await context.sync();
            setStatus(formatStatus, `Applied ${styleName} to ${paragraphs.length} paragraph(s).`, "success");
        });
    } catch (err) {
        setStatus(formatStatus, "Error: " + err.message, "error");
    }
}

async function applyUnderrubrik() {
    setStatus(formatStatus, "", "info");
    try {
        await Word.run(async (context) => {
            const range = context.document.getSelection();
            range.load("text");
            await context.sync();
            if (!range.text || !range.text.trim()) {
                setStatus(formatStatus, "Select text first, then click Underrubrik.", "info");
                return;
            }
            range.font.bold = true;
            await context.sync();
            setStatus(formatStatus, "Applied bold (Underrubrik) to selection.", "success");
        });
    } catch (err) {
        setStatus(formatStatus, "Error: " + err.message, "error");
    }
}

async function applyBlockquote() {
    setStatus(formatStatus, "", "info");
    try {
        await Word.run(async (context) => {
            const range = context.document.getSelection();
            range.load("text");
            range.paragraphs.load("items");
            await context.sync();
            const selectedText = range.text ? range.text.trim() : "";
            if (!selectedText) {
                setStatus(formatStatus, "Select text first, then click Citat.", "info");
                return;
            }
            const paragraphs = range.paragraphs.items;
            const lastPara = paragraphs[paragraphs.length - 1];
            const insertRange = lastPara.getRange("End");
            const newPara = insertRange.insertParagraph(selectedText, Word.InsertLocation.after);
            newPara.style = "Quote";
            await context.sync();
            setStatus(formatStatus, "Inserted blockquote below selection.", "success");
        });
    } catch (err) {
        setStatus(formatStatus, "Error: " + err.message, "error");
    }
}

function normalizeUrl(url) {
    if (!url) return "";
    if (!/^https?:\/\//i.test(url)) {
        url = "https://" + url;
    }
    return url.replace(/\/+$/, "");
}

async function publish() {
    const wpUrl = localStorage.getItem(STORAGE_KEY_URL);
    const username = localStorage.getItem(STORAGE_KEY_USER);
    const appPassword = localStorage.getItem(STORAGE_KEY_APP_PASS);

    if (!wpUrl || !username || !appPassword) {
        setStatus(publishStatus, "Please save your WordPress URL, username, and application password first.", "error");
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

        // Step 5: Create post via WordPress REST API
        publishText.textContent = "Publishing...";
        const postUrl = wpUrl + "/wp-json/wp/v2/posts";
        const auth = btoa(username + ":" + appPassword);

        const response = await fetch(postUrl, {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                "Authorization": "Basic " + auth,
            },
            body: JSON.stringify({
                title: title,
                content: bodyHtml,
                status: "draft",
            }),
        });

        if (!response.ok) {
            const errBody = await response.text();
            let errMsg = "WordPress API error: " + response.status;
            try {
                const errJson = JSON.parse(errBody);
                if (errJson.message) errMsg = errJson.message;
                else if (errJson.code) errMsg = errJson.code + ": " + (errJson.message || errMsg);
            } catch (_) {}
            if (response.status === 401) {
                errMsg = "Invalid credentials. Check your username and application password, then use Update Application Password.";
            } else if (response.status === 403) {
                errMsg = "Access denied. Ensure the application password has permission to create posts.";
            }
            throw new Error(errMsg);
        }

        const post = await response.json();
        const editUrl = wpUrl + "/wp-admin/post.php?post=" + post.id + "&action=edit";

        setStatus(
            publishStatus,
            'Draft created: "' + title + '". Opening in browser...',
            "success"
        );

        Office.context.ui.openBrowserWindow(editUrl);
    } catch (err) {
        setStatus(publishStatus, "Error: " + err.message, "error");
    } finally {
        publishBtn.disabled = false;
        publishText.textContent = "Publish to WordPress";
    }
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
    return html.replace(/<h[12][^>]*>.*?<\/h[12]>\s*/i, "");
}

function setStatus(el, message, type) {
    el.textContent = message;
    el.className = "status-text";
    if (type) el.classList.add("status-" + type);
}
