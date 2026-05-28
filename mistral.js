/* global Office, Word */

// =====================================================================
// mistral.js — Auto-redigér med sporede ændringer
// POV International Word Add-in
//
// Læser dokumentet afsnitsvis, sender til Mistral API og indsætter
// rettelser som sporede ændringer direkte i det åbne dokument.
// API-nøglen gemmes i localStorage på samme måde som WordPress-indstillinger.
// =====================================================================

// ── Konfiguration ────────────────────────────────────────────────────
// Skift denne URL hvis du placerer stilguiden et andet sted i repo'et.
// Format: https://raw.githubusercontent.com/[bruger]/[repo]/[branch]/[sti]
const STILGUIDE_URL =
    "https://raw.githubusercontent.com/mortenbay1/povai/main/pov-stilguide.md";

const MISTRAL_KEY_STORAGE = "mistral_api_key";

// ── Basis-prompt (fallback hvis GitHub ikke kan naas) ─────────────────
    const MISTRAL_BASE_PROMPT = `Du er en erfaren korrekturlæser på det danske nyhedsmedie POV International. Din opgave er udelukkende at rette fejl i den tekst, du modtager — ikke at omskrive den.

Tekstens struktur er ukrænkelig:
- Bevar alle afsnit i deres oprindelige rækkefølge
- Bevar alle afsnit — slet eller flet ikke afsnit
- Bevar tekstens længde — du må ikke forkorte eller kondensere indhold

Du retter kun:
- Stavefejl og slåfejl
- Forkerte bøjninger og grammatiske fejl (fx forkert køn, tal eller kasus — ikke ordvalg)
- Manglende eller forkerte tegnsætningstegn

Du må aldrig:
- Erstatte et ord med et andet ord, uanset om du mener det er mere naturligt eller korrekt
- Slette eller omskrive faktuelle oplysninger, tal, stednavne eller personnavne
- Omformulere sætninger der allerede er korrekte og velfungerende
- Tilføje indhold der ikke findes i originalen
- Ændre genrekoder som "NYHED // OVERBLIK" eller lignende

Når du er i tvivl om noget er en fejl eller et bevidst stilistisk valg, lader du det stå uændret.

Returner kun den rettede tekst uden kommentarer, forklaringer eller indledning.
Brug ALDRIG markdown-formatering i dit svar — ingen **, ingen *, ingen #, ingen -.
Dit svar må KUN indeholde den rettede tekst med almindelige tegn. Hvis du ser ** eller * i din output, har du begået en fejl.`;

// ── Hent stilguide fra GitHub ────────────────────────────────────────
// Returnerer den fulde prompt-streng med stilguide injiceret,
// eller basis-prompt alene hvis GitHub ikke kan naas.
// Resultatet caches i sessionStorage saa det kun hentes een gang pr. session.
const STILGUIDE_CACHE_KEY = "pov_stilguide_cache_v4";

async function buildSystemPrompt() {
    const cached = sessionStorage.getItem(STILGUIDE_CACHE_KEY);
    if (cached) return cached;

    try {
        const response = await fetch(STILGUIDE_URL, { cache: "no-cache" });
        if (!response.ok) throw new Error("HTTP " + response.status);

        const stilguide = await response.text();
        const prompt = MISTRAL_BASE_PROMPT +
            "\n\n---\n\nStilguiden herunder er formateret i markdown til læsbarhed — " +
            "det er KUN referencemateriale. Dit output må aldrig indeholde markdown.\n\n" +
            stilguide +
            "\n\n---\n\nABSOLUT REGEL: Du må kun rette stavefejl og grammatiske fejl. " +
            "Du må ikke omskrive, omformulere, erstatte ord eller ændre sætningsstruktur — " +
            "hverken ud fra stilguiden ovenfor eller din egen vurdering. " +
            "Hvis du er i tvivl, lader du teksten stå uændret.";

        sessionStorage.setItem(STILGUIDE_CACHE_KEY, prompt);
        return prompt;

    } catch (err) {
        console.warn("Stilguide kunne ikke hentes fra GitHub - koerer med basis-prompt.", err);
        return MISTRAL_BASE_PROMPT;
    }
}

// ── Elementer ────────────────────────────────────────────────────────
let autoRedigérBtn;
let autoRedigérText;
let korrekturStatus;
let mistralApiKeyInput;
let saveApiKeyBtn;
let apiKeyStatus;

// ── Init — kaldes fra taskpane.js via initMistralUI() ────────────────
function initMistralUI() {
    autoRedigérBtn     = document.getElementById("autoredigér-btn");
    autoRedigérText    = document.getElementById("autoredigér-text");
    korrekturStatus    = document.getElementById("korrektur-status");
    mistralApiKeyInput = document.getElementById("mistral-api-key");
    saveApiKeyBtn      = document.getElementById("save-api-key-btn");
    apiKeyStatus       = document.getElementById("api-key-status");

    if (!autoRedigérBtn) return;

    if (localStorage.getItem(MISTRAL_KEY_STORAGE)) {
        mistralApiKeyInput.placeholder = "••••••••••••••••••••• (gemt)";
        setStatus(apiKeyStatus, "✓ API-nøgle er gemt", "success");
    }

    autoRedigérBtn.addEventListener("click", autoRedigér);

    saveApiKeyBtn.addEventListener("click", () => {
        const key = mistralApiKeyInput.value.trim();
        if (!key) {
            setStatus(apiKeyStatus, "Indsæt en nøgle først.", "error");
            return;
        }
        localStorage.setItem(MISTRAL_KEY_STORAGE, key);
        mistralApiKeyInput.value = "";
        mistralApiKeyInput.placeholder = "••••••••••••••••••••• (gemt)";
        setStatus(apiKeyStatus, "✓ Nøgle gemt", "success");
    });

    mistralApiKeyInput.addEventListener("keydown", (e) => {
        if (e.key === "Enter") saveApiKeyBtn.click();
    });
}


// ── Post-processing ──────────────────────────────────────────────────
// Retter mønstre Mistral konsekvent fejler pga. engelsk træning:
// 1. Komma inden for anførselstegn → uden for (dansk regel)
// 2. Rettet anførselstegn → buede anførselstegn (POV's typografi)
function postProcess(text) {
    // Fjern markdown hvis Mistral alligevel indsætter det
    text = text.replace(/\*\*([^*]+)\*\*/g, '$1');
    text = text.replace(/\*([^*]+)\*/g, '$1');
    // Komma uden for anførselstegn (dansk regel)
    text = text.replace(/,(”|")/g, '$1,');
    // Rettet anførselstegn → buede
    text = text.replace(/"([^"]+)"/g, '“$1”');
    return text;
}

// ── Word-level diff ───────────────────────────────────────────────────
// Tokeniserer tekst i ord+whitespace og returnerer LCS-baseret diff.
function tokenize(text) {
    return text.match(/\S+|\s+/g) || [];
}

function diffTokens(origTokens, revTokens) {
    const m = origTokens.length;
    const n = revTokens.length;
    const dp = Array.from({ length: m + 1 }, () => new Uint16Array(n + 1));
    for (let i = 1; i <= m; i++) {
        for (let j = 1; j <= n; j++) {
            dp[i][j] = origTokens[i-1] === revTokens[j-1]
                ? dp[i-1][j-1] + 1
                : Math.max(dp[i-1][j], dp[i][j-1]);
        }
    }
    const ops = [];
    let i = m, j = n;
    while (i > 0 || j > 0) {
        if (i > 0 && j > 0 && origTokens[i-1] === revTokens[j-1]) {
            ops.push({ op: 'equal', orig: origTokens[i-1], rev: origTokens[i-1] });
            i--; j--;
        } else if (j > 0 && (i === 0 || dp[i][j-1] >= dp[i-1][j])) {
            ops.push({ op: 'insert', orig: '', rev: revTokens[j-1] });
            j--;
        } else {
            ops.push({ op: 'delete', orig: origTokens[i-1], rev: '' });
            i--;
        }
    }
    ops.reverse();
    // Grupper delete+insert → replace
    const grouped = [];
    for (let k = 0; k < ops.length; k++) {
        if (ops[k].op === 'delete' && k + 1 < ops.length && ops[k+1].op === 'insert') {
            grouped.push({ op: 'replace', orig: ops[k].orig, rev: ops[k+1].rev });
            k++;
        } else {
            grouped.push(ops[k]);
        }
    }
    return grouped;
}

// Anvender diff som præcise sporede ændringer i et Word-afsnit.
async function applyDiffToParagraph(context, para, origText, revisedText) {
    const ops = diffTokens(tokenize(origText), tokenize(revisedText));
    let anyChange = false;

    for (const op of ops) {
        if (op.op === 'equal' || /^\s+$/.test(op.orig || op.rev || '')) continue;
        anyChange = true;

        if (op.op === 'replace') {
            const results = para.search(op.orig, { matchCase: true, matchWholeWord: false });
            results.load('items');
            await context.sync();
            if (results.items.length > 0) {
                results.items[0].insertText(op.rev, Word.InsertLocation.replace);
                await context.sync();
            }
        }
        // insert/delete logges men ignoreres foreløbig — sjældne ved stavekorrektion
    }
    return anyChange;
}

// ── Mistral API-kald ─────────────────────────────────────────────────
async function callMistral(text, apiKey, systemPrompt) {
    const response = await fetch("https://api.mistral.ai/v1/chat/completions", {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
            "Authorization": "Bearer " + apiKey
        },
        body: JSON.stringify({
            model: "mistral-medium-latest",
            temperature: 0,
            messages: [
                { role: "system", content: systemPrompt },
                { role: "user",   content: text }
            ]
        })
    });

    if (!response.ok) {
        const err = await response.json().catch(() => ({}));
        const msg = err.message || response.statusText;
        throw new Error(`${response.status}: ${msg}`);
    }

    const data = await response.json();
    return data.choices[0].message.content.trim();
}

// ── Hoved-funktion ───────────────────────────────────────────────────
async function autoRedigér() {
    const apiKey = localStorage.getItem(MISTRAL_KEY_STORAGE);
    if (!apiKey) {
        setStatus(korrekturStatus,
            "Ingen API-nøgle. Åbn 'API-nøgle (Mistral)' herunder og gem din nøgle.",
            "error");
        return;
    }

    autoRedigérBtn.disabled = true;
    autoRedigérText.textContent = "Redigerer…";
    setStatus(korrekturStatus, "Henter stilguide…", "info");

    // Byg systemprompt — henter stilguide fra GitHub (eller bruger cache)
    const systemPrompt = await buildSystemPrompt();

    try {
        await Word.run(async (context) => {
            const paragraphs = context.document.body.paragraphs;
            paragraphs.load("items");
            await context.sync();

            const items = paragraphs.items;

            // Aktiver spor ændringer
            context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
            await context.sync();

            let changed = 0;

            for (let i = 0; i < items.length; i++) {
                const para = items[i];
                para.load("text, styleBuiltIn");
                await context.sync();

                const originalText = para.text.trim();

                // Spring tomme og meget korte afsnit over
                if (!originalText || originalText.length < 15) continue;

                // Spring overskrifter over — de redigeres ikke
                const style = para.styleBuiltIn || "";
                if (
                    style === Word.BuiltInStyleName.heading1 ||
                    style === Word.BuiltInStyleName.heading2 ||
                    style === Word.BuiltInStyleName.heading3 ||
                    style === Word.BuiltInStyleName.title
                ) continue;

                setStatus(korrekturStatus,
                    `Analyserer afsnit ${i + 1} af ${items.length}…`,
                    "info");

                const rawText = await callMistral(originalText, apiKey, systemPrompt);
                const revisedText = rawText ? postProcess(rawText) : rawText;

                if (revisedText && revisedText !== originalText) {
                    const hadChanges = await applyDiffToParagraph(context, para, originalText, revisedText);
                    await context.sync();
                    if (hadChanges) changed++;
                }
            }

            await context.sync();

            if (changed === 0) {
                setStatus(korrekturStatus,
                    "✓ Ingen rettelser fundet — teksten ser fin ud!",
                    "success");
            } else {
                setStatus(korrekturStatus,
                    `✓ Færdig! ${changed} afsnit rettet. Gennemgå med Words "Acceptér/Afvis"-knapper.`,
                    "success");
            }
        });

    } catch (err) {
        console.error("Auto-redigér fejl:", err);
        const msg = err.message || "";
        if (msg.includes("401")) {
            setStatus(korrekturStatus,
                "Ugyldig API-nøgle. Tjek nøglen under 'API-nøgle (Mistral)'.",
                "error");
        } else if (msg.includes("429")) {
            setStatus(korrekturStatus,
                "Rate limit nået. Vent et øjeblik og prøv igen.",
                "error");
        } else {
            setStatus(korrekturStatus, "Fejl: " + msg, "error");
        }
    } finally {
        autoRedigérBtn.disabled = false;
        autoRedigérText.textContent = "✦ Auto-redigér";
    }
}

// ── Genbruger setStatus fra taskpane.js (tilgængelig globalt) ────────
