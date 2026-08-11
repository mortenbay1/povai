/**
 * WordToWP — Google Docs add-on
 * Settings (PropertiesService), HTML export, WordPress draft via REST API.
 */

var PROP_WP_URL = "WP_URL";
var PROP_WP_USER = "WP_USER";
var PROP_WP_PASS = "WP_PASS";

function onOpen() {
  DocumentApp.getUi()
    .createAddonMenu()
    .addItem("Open WordToWP", "showSidebar")
    .addToUi();
}

function onInstall() {
  onOpen();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("Sidebar")
    .setTitle("WordToWP")
    .setWidth(400);
  DocumentApp.getUi().showSidebar(html);
}

// --- Formatting (matches Word add-in behaviour) ---

/**
 * @returns {{ ok: boolean, error?: string, count?: number }}
 */
function formatRubrik() {
  return applyHeadingToSelection_(DocumentApp.ParagraphHeading.TITLE);
}

/**
 * @returns {{ ok: boolean, error?: string, count?: number }}
 */
function formatMellemrubrik() {
  return applyHeadingToSelection_(DocumentApp.ParagraphHeading.HEADING2);
}

/**
 * @returns {{ ok: boolean, error?: string, count?: number }}
 */
function formatBrodtekst() {
  return applyHeadingToSelection_(DocumentApp.ParagraphHeading.NORMAL);
}

/**
 * Normal paragraph + bold on selection (Underrubrik).
 * @returns {{ ok: boolean, error?: string }}
 */
function formatUnderrubrik() {
  try {
    const doc = DocumentApp.getActiveDocument();
    const sel = doc.getSelection();
    if (!sel) {
      return { ok: false, error: "Select text first, then click Underrubrik." };
    }
    const elems = sel.getRangeElements();
    const paras = [];
    for (var i = 0; i < elems.length; i++) {
      const p = getParagraphAncestor_(elems[i].getElement());
      if (p && paras.indexOf(p) === -1) {
        paras.push(p);
      }
    }
    if (paras.length === 0) {
      return { ok: false, error: "Select text first, then click Underrubrik." };
    }
    for (var j = 0; j < paras.length; j++) {
      paras[j].setHeading(DocumentApp.ParagraphHeading.NORMAL);
    }
    for (var k = 0; k < elems.length; k++) {
      const re = elems[k];
      const el = re.getElement();
      if (el.getType() === DocumentApp.ElementType.TEXT) {
        const text = el.asText();
        const start = re.getStartOffset();
        const end = re.getEndOffsetInclusive();
        if (start >= 0 && end >= start) {
          text.setBold(start, end, true);
        }
      }
    }
    return { ok: true };
  } catch (e) {
    return { ok: false, error: String(e.message || e) };
  }
}

/**
 * Duplicates selected text as a new paragraph below with Quote style (Citat).
 * @returns {{ ok: boolean, error?: string }}
 */
function formatCitat() {
  try {
    const doc = DocumentApp.getActiveDocument();
    const sel = doc.getSelection();
    if (!sel) {
      return { ok: false, error: "Select text first, then click Citat." };
    }
    const elems = sel.getRangeElements();
    const parts = [];
    for (var i = 0; i < elems.length; i++) {
      const el = elems[i].getElement();
      const re = elems[i];
      if (el.getType() === DocumentApp.ElementType.TEXT) {
        const t = el.asText();
        const s = re.getStartOffset();
        const e = re.getEndOffsetInclusive();
        if (s >= 0 && e >= s) {
          parts.push(t.getText().substring(s, e + 1));
        }
      }
    }
    const selectedText = parts.join("").trim();
    if (!selectedText) {
      return { ok: false, error: "Select text first, then click Citat." };
    }
    var lastPara = null;
    for (var j = elems.length - 1; j >= 0; j--) {
      const p = getParagraphAncestor_(elems[j].getElement());
      if (p) {
        lastPara = p;
        break;
      }
    }
    if (!lastPara) {
      return { ok: false, error: "Could not find a paragraph for the selection." };
    }
    const body = doc.getBody();
    const top = getTopLevelChildOfBody_(lastPara);
    const idx = body.getChildIndex(top);
    body.insertParagraph(idx + 1, selectedText);
    const newChild = body.getChild(idx + 1);
    if (newChild.getType() === DocumentApp.ElementType.PARAGRAPH) {
      // Google Docs has no ParagraphHeading.QUOTE; Subtitle is reserved for Citat export as <blockquote>.
      newChild.asParagraph().setHeading(DocumentApp.ParagraphHeading.SUBTITLE);
    }
    return { ok: true };
  } catch (e) {
    return { ok: false, error: String(e.message || e) };
  }
}

/**
 * @param {GoogleAppsScript.Document.ParagraphHeading} heading
 * @returns {{ ok: boolean, error?: string, count?: number }}
 */
function applyHeadingToSelection_(heading) {
  try {
    const sel = DocumentApp.getActiveDocument().getSelection();
    if (!sel) {
      return { ok: false, error: "Select text first, then click a style." };
    }
    const elems = sel.getRangeElements();
    const paras = [];
    for (var i = 0; i < elems.length; i++) {
      const p = getParagraphAncestor_(elems[i].getElement());
      if (p && paras.indexOf(p) === -1) {
        paras.push(p);
      }
    }
    if (paras.length === 0) {
      return { ok: false, error: "Select text first, then click a style." };
    }
    for (var j = 0; j < paras.length; j++) {
      paras[j].setHeading(heading);
    }
    return { ok: true, count: paras.length };
  } catch (e) {
    return { ok: false, error: String(e.message || e) };
  }
}

/**
 * @param {GoogleAppsScript.Document.Element} element
 * @returns {GoogleAppsScript.Document.Paragraph|null}
 */
function getParagraphAncestor_(element) {
  var n = element;
  var guard = 0;
  while (n && guard++ < 30) {
    const t = n.getType();
    if (t === DocumentApp.ElementType.PARAGRAPH) {
      return n.asParagraph();
    }
    if (t === DocumentApp.ElementType.LIST_ITEM) {
      return n.asListItem();
    }
    n = n.getParent();
  }
  return null;
}

/**
 * Walks up to the element that is a direct child of Body (paragraph, list, table, …).
 * @param {GoogleAppsScript.Document.Element} element
 * @returns {GoogleAppsScript.Document.Element}
 */
function getTopLevelChildOfBody_(element) {
  var n = element;
  var guard = 0;
  while (n && guard++ < 30) {
    const p = n.getParent();
    if (p.getType() === DocumentApp.ElementType.BODY) {
      return n;
    }
    n = p;
  }
  return element;
}

/**
 * @returns {{ url: string, username: string, hasPassword: boolean }}
 */
function loadSettings() {
  const props = PropertiesService.getUserProperties();
  return {
    url: props.getProperty(PROP_WP_URL) || "",
    username: props.getProperty(PROP_WP_USER) || "",
    hasPassword: !!props.getProperty(PROP_WP_PASS),
  };
}

/**
 * @param {string} url
 * @param {string} username
 * @param {string} appPassword
 * @returns {{ ok: boolean, error?: string }}
 */
function saveSettings(url, username, appPassword) {
  try {
    url = normalizeUrl_(String(url || "").trim());
    username = String(username || "").trim();
    appPassword = String(appPassword || "").trim().replace(/\s/g, "");

    if (!url) {
      return { ok: false, error: "Please enter a valid WordPress URL." };
    }
    if (!username) {
      return { ok: false, error: "Please enter your WordPress username." };
    }
    if (!appPassword) {
      return { ok: false, error: "Please enter an application password." };
    }

    const props = PropertiesService.getUserProperties();
    props.setProperty(PROP_WP_URL, url);
    props.setProperty(PROP_WP_USER, username);
    props.setProperty(PROP_WP_PASS, appPassword);
    return { ok: true };
  } catch (e) {
    return { ok: false, error: String(e.message || e) };
  }
}

/**
 * Clears stored application password so user can enter a new one.
 */
function clearStoredPassword() {
  PropertiesService.getUserProperties().deleteProperty(PROP_WP_PASS);
  return { ok: true };
}

/**
 * Preview: document name + short text (optional debug).
 */
function getActiveDocumentInfo() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const full = body.getText();
  const preview = full.length > 400 ? full.substring(0, 400) + "…" : full;
  return {
    name: doc.getName(),
    preview: preview || "(empty document)",
  };
}

/**
 * Builds HTML from the active document, creates a draft post, returns edit URL or error.
 * @returns {{ ok: boolean, editUrl?: string, title?: string, error?: string }}
 */
function publishDraft() {
  const props = PropertiesService.getUserProperties();
  const wpUrl = props.getProperty(PROP_WP_URL);
  const username = props.getProperty(PROP_WP_USER);
  const appPassword = props.getProperty(PROP_WP_PASS);

  if (!wpUrl || !username || !appPassword) {
    return {
      ok: false,
      error: "Save your WordPress URL, username, and application password first.",
    };
  }

  try {
    const doc = DocumentApp.getActiveDocument();
    const fullHtml = documentBodyToHtml_(doc.getBody());
    if (!fullHtml || !fullHtml.trim()) {
      return { ok: false, error: "Document appears to be empty." };
    }

    const title = extractTitleFromHtml_(fullHtml);
    const bodyHtml = removeFirstTitleBlockFromHtml_(fullHtml, title);

    const postUrl = wpUrl + "/wp-json/wp/v2/posts";
    const auth = Utilities.base64Encode(username + ":" + appPassword);
    const payload = {
      title: title,
      content: bodyHtml,
      status: "draft",
    };

    const response = UrlFetchApp.fetch(postUrl, {
      method: "post",
      contentType: "application/json",
      headers: {
        Authorization: "Basic " + auth,
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });

    const code = response.getResponseCode();
    const responseText = response.getContentText();

    if (code < 200 || code >= 300) {
      let msg = "WordPress API error: " + code;
      try {
        const errJson = JSON.parse(responseText);
        if (errJson.message) msg = errJson.message;
      } catch (ignore) {}
      if (code === 401) {
        msg =
          "Invalid credentials. Check username and application password, then save again.";
      } else if (code === 403) {
        msg =
          "Access denied. Ensure the application password can create posts.";
      }
      return { ok: false, error: msg };
    }

    const post = JSON.parse(responseText);
    const editUrl = wpUrl + "/wp-admin/post.php?post=" + post.id + "&action=edit";
    return { ok: true, editUrl: editUrl, title: title };
  } catch (e) {
    return { ok: false, error: String(e.message || e) };
  }
}

// --- HTML conversion (paragraphs, headings, bold / italic) ---

/**
 * @param {GoogleAppsScript.Document.Body} body
 * @returns {string}
 */
function documentBodyToHtml_(body) {
  const parts = [];
  const n = body.getNumChildren();
  var i = 0;
  while (i < n) {
    const child = body.getChild(i);
    const type = child.getType();
    if (type === DocumentApp.ElementType.PARAGRAPH) {
      parts.push(paragraphToHtml_(child.asParagraph()));
      i++;
    } else if (type === DocumentApp.ElementType.LIST_ITEM) {
      const items = [];
      var ordered = null;
      while (i < n && body.getChild(i).getType() === DocumentApp.ElementType.LIST_ITEM) {
        const li = body.getChild(i).asListItem();
        if (ordered === null) {
          ordered = isOrderedListItem_(li);
        }
        items.push("<li>" + paragraphRichTextToHtml_(li) + "</li>");
        i++;
      }
      const tag = ordered ? "ol" : "ul";
      parts.push("<" + tag + ">" + items.join("") + "</" + tag + ">");
    } else if (type === DocumentApp.ElementType.TABLE) {
      parts.push("<p><em>[Table not exported — add in WordPress]</em></p>");
      i++;
    } else {
      i++;
    }
  }
  return parts.filter(Boolean).join("\n");
}

/**
 * @param {GoogleAppsScript.Document.ListItem} listItem
 * @returns {boolean}
 */
function isOrderedListItem_(listItem) {
  try {
    const g = listItem.getGlyphType();
    return (
      g === DocumentApp.GlyphType.NUMBER ||
      g === DocumentApp.GlyphType.LATIN_NUMBER ||
      g === DocumentApp.GlyphType.ROMAN_NUMBER
    );
  } catch (e) {
    return false;
  }
}

/**
 * @param {GoogleAppsScript.Document.Paragraph} paragraph
 * @returns {string}
 */
function paragraphToHtml_(paragraph) {
  const raw = paragraph.getText();
  if (!raw.trim()) {
    return "";
  }
  const inner = paragraphRichTextToHtml_(paragraph);
  const h = paragraph.getHeading();
  // Subtitle = Citat from the sidebar (no QUOTE enum in DocumentApp).
  if (h === DocumentApp.ParagraphHeading.SUBTITLE) {
    return "<blockquote><p>" + inner + "</p></blockquote>";
  }
  const tag = paragraphHeadingToTag_(paragraph);
  return "<" + tag + ">" + inner + "</" + tag + ">";
}

/**
 * @param {GoogleAppsScript.Document.Paragraph} block
 * @returns {string}
 */
function paragraphRichTextToHtml_(block) {
  const text = block.editAsText();
  const s = text.getText();
  if (!s) {
    return "";
  }
  var out = [];
  var i = 0;
  while (i < s.length) {
    var bold = text.isBold(i) === true;
    var italic = text.isItalic(i) === true;
    var j = i + 1;
    while (j < s.length) {
      var b = text.isBold(j) === true;
      var it = text.isItalic(j) === true;
      if (b !== bold || it !== italic) break;
      j++;
    }
    var chunk = escapeHtml_(s.substring(i, j));
    if (bold) chunk = "<strong>" + chunk + "</strong>";
    if (italic) chunk = "<em>" + chunk + "</em>";
    out.push(chunk);
    i = j;
  }
  return out.join("");
}

/**
 * @param {GoogleAppsScript.Document.Paragraph} paragraph
 * @returns {string}
 */
function paragraphHeadingToTag_(paragraph) {
  const h = paragraph.getHeading();
  switch (h) {
    case DocumentApp.ParagraphHeading.TITLE:
    case DocumentApp.ParagraphHeading.HEADING1:
      return "h1";
    case DocumentApp.ParagraphHeading.SUBTITLE:
    case DocumentApp.ParagraphHeading.HEADING2:
      return "h2";
    case DocumentApp.ParagraphHeading.HEADING3:
      return "h3";
    case DocumentApp.ParagraphHeading.HEADING4:
      return "h4";
    case DocumentApp.ParagraphHeading.HEADING5:
      return "h5";
    case DocumentApp.ParagraphHeading.HEADING6:
      return "h6";
    default:
      return "p";
  }
}

/**
 * @param {string} html
 * @returns {string}
 */
function extractTitleFromHtml_(html) {
  var m = html.match(/<h[12][^>]*>([\s\S]*?)<\/h[12]>/i);
  if (m) {
    return m[1].replace(/<[^>]+>/g, "").trim();
  }
  m = html.match(/<p[^>]*>([\s\S]*?)<\/p>/i);
  if (m) {
    var t = m[1].replace(/<[^>]+>/g, "").trim();
    if (t.length > 60) return t.substring(0, 57) + "...";
    return t;
  }
  return "Untitled Post";
}

/**
 * @param {string} html
 * @param {string} title
 * @returns {string}
 */
function removeFirstTitleBlockFromHtml_(html, title) {
  return html.replace(/<h[12][^>]*>[\s\S]*?<\/h[12]>\s*/i, "");
}

/**
 * @param {string} str
 * @returns {string}
 */
function escapeHtml_(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/**
 * @param {string} url
 * @returns {string}
 */
function normalizeUrl_(url) {
  if (!url) return "";
  if (!/^https?:\/\//i.test(url)) {
    url = "https://" + url;
  }
  return url.replace(/\/+$/, "");
}
