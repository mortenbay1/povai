/**
 * Converts Word OOXML body content to clean HTML suitable for WordPress.
 * Handles paragraphs, headings, bold, italic, underline, strikethrough,
 * ordered/unordered lists, and hyperlinks.
 */

const OoxmlToHtml = (() => {
    const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";

    function convert(ooxmlString) {
        const parser = new DOMParser();
        const doc = parser.parseFromString(ooxmlString, "application/xml");

        // Extract relationships for hyperlinks
        const relationships = extractRelationships(doc);

        // Find the document body
        const body = doc.getElementsByTagNameNS(W_NS, "body")[0];
        if (!body) {
            // If no body, try parsing the OOXML as just paragraphs
            return convertParagraphs(doc.documentElement, relationships);
        }

        return convertParagraphs(body, relationships);
    }

    function extractRelationships(doc) {
        const rels = {};
        const relElements = doc.getElementsByTagNameNS(REL_NS, "Relationship");
        for (let i = 0; i < relElements.length; i++) {
            const rel = relElements[i];
            rels[rel.getAttribute("Id")] = {
                type: rel.getAttribute("Type"),
                target: rel.getAttribute("Target"),
            };
        }
        return rels;
    }

    function convertParagraphs(container, relationships) {
        const paragraphs = container.getElementsByTagNameNS(W_NS, "p");
        const htmlParts = [];
        let listType = null; // "ol" or "ul"
        let inList = false;

        for (let i = 0; i < paragraphs.length; i++) {
            const para = paragraphs[i];
            const style = getParagraphStyle(para);
            const listInfo = getListInfo(para);
            const runs = para.getElementsByTagNameNS(W_NS, "r");
            let inlineHtml = "";

            for (let j = 0; j < runs.length; j++) {
                inlineHtml += convertRun(runs[j], relationships);
            }

            // Handle lists
            if (listInfo) {
                const newListType = listInfo.ordered ? "ol" : "ul";
                if (!inList || listType !== newListType) {
                    if (inList) {
                        htmlParts.push(`</${listType}>`);
                    }
                    listType = newListType;
                    htmlParts.push(`<${listType}>`);
                    inList = true;
                }
                htmlParts.push(`<li>${inlineHtml}</li>`);
                continue;
            }

            // Close list if we were in one
            if (inList) {
                htmlParts.push(`</${listType}>`);
                inList = false;
                listType = null;
            }

            // Skip empty paragraphs
            if (!inlineHtml.trim()) {
                continue;
            }

            // Map Word styles to HTML elements
            if (style && /^Heading(\d)$/.test(style)) {
                const level = RegExp.$1;
                htmlParts.push(`<h${level}>${inlineHtml}</h${level}>`);
            } else if (style === "Title") {
                htmlParts.push(`<h1>${inlineHtml}</h1>`);
            } else if (style === "Subtitle") {
                htmlParts.push(`<h2>${inlineHtml}</h2>`);
            } else if (style === "Quote" || style === "IntenseQuote" || style === "Citat" || style === "Intensivt citat" || style === "BlockQuotation") {
                htmlParts.push(`<blockquote><p>${inlineHtml}</p></blockquote>`);
            } else {
                htmlParts.push(`<p>${inlineHtml}</p>`);
            }
        }

        // Close any open list
        if (inList) {
            htmlParts.push(`</${listType}>`);
        }

        return htmlParts.join("\n");
    }

    function getParagraphStyle(para) {
        const pPr = para.getElementsByTagNameNS(W_NS, "pPr")[0];
        if (!pPr) return null;
        const pStyle = pPr.getElementsByTagNameNS(W_NS, "pStyle")[0];
        if (!pStyle) return null;
        return pStyle.getAttributeNS(W_NS, "val") || pStyle.getAttribute("w:val");
    }

    function getListInfo(para) {
        const pPr = para.getElementsByTagNameNS(W_NS, "pPr")[0];
        if (!pPr) return null;
        const numPr = pPr.getElementsByTagNameNS(W_NS, "numPr")[0];
        if (!numPr) return null;
        const numId = numPr.getElementsByTagNameNS(W_NS, "numId")[0];
        if (!numId) return null;

        const id = numId.getAttributeNS(W_NS, "val") || numId.getAttribute("w:val");
        if (!id || id === "0") return null;

        // Heuristic: numId of 1 is typically bullet list, 2+ is numbered
        // This is a simplification — full accuracy would require parsing numbering.xml
        return { ordered: parseInt(id) > 1 };
    }

    function convertRun(run, relationships) {
        const rPr = run.getElementsByTagNameNS(W_NS, "rPr")[0];
        const textNodes = run.getElementsByTagNameNS(W_NS, "t");
        let text = "";
        for (let i = 0; i < textNodes.length; i++) {
            text += textNodes[i].textContent;
        }

        // Check for tabs and breaks
        const tabs = run.getElementsByTagNameNS(W_NS, "tab");
        const breaks = run.getElementsByTagNameNS(W_NS, "br");
        if (tabs.length > 0) text += "\t";
        if (breaks.length > 0) text += "<br>";

        if (!text) return "";

        // Apply formatting
        if (rPr) {
            if (hasElement(rPr, "b")) text = `<strong>${text}</strong>`;
            if (hasElement(rPr, "i")) text = `<em>${text}</em>`;
            if (hasElement(rPr, "u")) text = `<u>${text}</u>`;
            if (hasElement(rPr, "strike")) text = `<del>${text}</del>`;
        }

        // Check for hyperlinks — the run may be inside a w:hyperlink element
        const parent = run.parentNode;
        if (parent && parent.localName === "hyperlink" && parent.namespaceURI === W_NS) {
            const rId = parent.getAttributeNS(R_NS, "id") || parent.getAttribute("r:id");
            if (rId && relationships[rId]) {
                text = `<a href="${escapeHtml(relationships[rId].target)}">${text}</a>`;
            }
        }

        return text;
    }

    function hasElement(rPr, localName) {
        const el = rPr.getElementsByTagNameNS(W_NS, localName)[0];
        if (!el) return false;
        // Check if explicitly turned off: w:val="0" or w:val="false"
        const val = el.getAttributeNS(W_NS, "val") || el.getAttribute("w:val");
        if (val === "0" || val === "false") return false;
        return true;
    }

    function escapeHtml(str) {
        return str
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;");
    }

    return { convert };
})();
