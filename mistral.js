/* global Office, Word */

// =====================================================================
// mistral.js — Auto-redigér med sporede ændringer
// POV International Word Add-in
//
// Læser dokumentet afsnitsvis, sender til Mistral API og indsætter
// rettelser som sporede ændringer direkte i det åbne dokument.
// API-nøglen gemmes i localStorage på samme måde som WordPress-indstillinger.
//
// Word-level diff sikrer at kun de faktisk ændrede ord vises som
// sporede ændringer — ikke hele afsnittet.
// =====================================================================

// ── Konfiguration ────────────────────────────────────────────────────
const STILGUIDE_URL =
    "https://raw.githubusercontent.com/mortenbay1/povai/main/pov-stilguide.md";

const MISTRAL_KEY_STORAGE = "mistral_api_key";

// ── Basis-prompt (fallback hvis GitHub ikke kan naas) ─────────────────
const MISTRAL_BASE_PROMPT = `Du er korrekturlæser på POV International. Du må KUN rette stavefejl, bøjningsfejl og tegnsætning.

ABSOLUT FORBUD — du må aldrig erstatte et ord med et andet:
- "monitorere" forbliver "monitorere" (ikke "overvåge")
- "implementere" forbliver "implementere" (ikke "indføre")
- "fokusere" forbliver "fokusere" (ikke "koncentrere sig om")
- Selv hvis et andet ord lyder mere dansk eller naturligt: lad det stå.

Tekstens struktur er ukrænkelig:
- Bevar alle afsnit i deres oprindelige rækkefølge og længde
- Slet, flet eller forkort ikke afsnit

Du retter kun:
- Stavefejl og slåfejl (fx "hsue" → "huse")
- Bøjningsfejl (fx forkert køn, tal, kasus, verbalform)
- Manglende eller forkert tegnsætning

Du må aldrig:
- Erstatte et ord med et synonym, uanset hvor "bedre" det lyder
- Omformulere sætninger der allerede er korrekte
- Tilføje eller fjerne indhold
- Ændre genrekoder som "NYHED // OVERBLIK" eller lignende
- Slette eller omskrive faktuelle oplysninger, tal, stednavne eller personnavne

Når du er i tvivl: lad teksten stå uændret.

Returner kun den rettede tekst — ingen kommentarer, ingen forklaringer, ingen indledning.
Brug ALDRIG markdown-formatering: ingen **, ingen *, ingen #, ingen -.`;

// ── Hent stilguide fra GitHub ────────────────────────────────────────
const STILGUIDE_CACHE_KEY = "pov_stilguide_cache_v5";

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
function postProcess(text) {
    // 0. Fjern markdown-formatering
    text = text.replace(/\*\*([^*]+)\*\*/g, '$1');              // **fed** → fed
    text = text.replace(/(?<!\*)\*([^*]+)\*(?!\*)/g, '$1');     // *kursiv* → kursiv
    text = text.replace(/`([^`]+)`/g, '$1');                    // `kode` → kode
    text = text.replace(/^#{1,6}\s+/gm, '');                    // # overskrift → overskrift

    // 1. Rettet anførselstegn → buede
    text = text.replace(/"([^"]+)"/g, '“$1”');

    // 2. Komma uden for anførselstegn (dansk regel)
    text = text.replace(/,”/g, '”,');

    return text;
}

// ── Tokenizer ────────────────────────────────────────────────────────
// Splitter tekst i atomare tokens: hele ord, whitespace, og enkelttegns-
// tegnsætning. Diff'en arbejder KUN på disse tokens — aldrig delstrenge.
// Dette forhindrer "mangel"→"dergel"-fælden hvor karakterdiff splitter
// ord op i mindre dele.
function tokenize(text) {
    const re = /(\s+|[.,;:!?"”“„''()\[\]—–\-]|\S+)/g;
    return text.match(re) || [];
}

// ── Word-level diff (LCS-baseret) ────────────────────────────────────
// Returnerer en sekvens af operationer: {op: 'keep'|'delete'|'insert', text}
function diffTokens(aTokens, bTokens) {
    const n = aTokens.length;
    const m = bTokens.length;

    // Byg LCS-tabel
    const dp = Array.from({ length: n + 1 }, () => new Int32Array(m + 1));
    for (let i = 1; i <= n; i++) {
        for (let j = 1; j <= m; j++) {
            if (aTokens[i - 1] === bTokens[j - 1]) {
                dp[i][j] = dp[i - 1][j - 1] + 1;
            } else {
                dp[i][j] = Math.max(dp[i - 1][j], dp[i][j - 1]);
            }
        }
    }

    // Backtrack for at bygge operationssekvensen
    const ops = [];
    let i = n, j = m;
    while (i > 0 && j > 0) {
        if (aTokens[i - 1] === bTokens[j - 1]) {
            ops.unshift({ op: 'keep', text: aTokens[i - 1] });
            i--; j--;
        } else if (dp[i - 1][j] >= dp[i][j - 1]) {
            ops.unshift({ op: 'delete', text: aTokens[i - 1] });
            i--;
        } else {
            ops.unshift({ op: 'insert', text: bTokens[j - 1] });
            j--;
        }
    }
    while (i > 0) ops.unshift({ op: 'delete', text: aTokens[--i] });
    while (j > 0) ops.unshift({ op: 'insert', text: bTokens[--j] });

    return ops;
}

// ── Konsolidér diff-operationer ──────────────────────────────────────
// Slår sammenhængende delete/insert sammen så Word viser fx "mangel" → "der"
// som ÉN erstatning, ikke flere mikro-ændringer. Delete+insert = replace.
function consolidateOps(ops) {
    const result = [];
    let i = 0;
    while (i < ops.length) {
        if (ops[i].op === 'keep') {
            let text = '';
            while (i < ops.length && ops[i].op === 'keep') {
                text += ops[i].text;
                i++;
            }
            result.push({ op: 'keep', text });
        } else {
            let deleteText = '';
            while (i < ops.length && ops[i].op === 'delete') {
                deleteText += ops[i].text;
                i++;
            }
            let insertText = '';
            while (i < ops.length && ops[i].op === 'insert') {
                insertText += ops[i].text;
                i++;
            }

            if (deleteText && insertText) {
                result.push({ op: 'replace', deleteText, insertText });
            } else if (deleteText) {
                result.push({ op: 'delete', text: deleteText });
            } else if (insertText) {
                result.push({ op: 'insert', text: insertText });
            }
        }
    }
    return result;
}

// ── Hjælpefunktion: søg og erstat første forekomst i et afsnit ──────
async function searchAndReplace(para, searchText, replaceText, context) {
    if (!searchText || searchText === replaceText) return false;

    try {
        const searchResults = para.search(searchText, {
            matchCase: true,
            matchWholeWord: false
        });
        searchResults.load("items");
        await context.sync();

        if (searchResults.items.length === 0) {
            console.warn("Kunne ikke finde søgetekst i afsnit:", searchText.slice(0, 50));
            return false;
        }

        const range = searchResults.items[0];
        range.insertText(replaceText, Word.InsertLocation.replace);
        await context.sync();
        return true;
    } catch (err) {
        console.warn("search/replace fejlede for:", searchText.slice(0, 50), err.message);
        return false;
    }
}

// ── Anvend diff på et afsnit via Word ranges ─────────────────────────
// Itererer gennem konsoliderede ops. For hver ændring bygges en søgestreng
// med kontekstanker (tekst lige før/efter) der unikt identificerer
// positionen, så vi rammer den rigtige forekomst hvis et ord optræder
// flere gange i afsnittet.
async function applyDiffToParagraph(para, ops, context) {
    let anyChanges = false;

    for (let opIdx = 0; opIdx < ops.length; opIdx++) {
        const op = ops[opIdx];
        if (op.op === 'keep') continue;

        // Hent kontekst fra omkringliggende keeps som anker
        const prevKeep = opIdx > 0 && ops[opIdx - 1].op === 'keep' ? ops[opIdx - 1].text : '';
        const nextKeep = opIdx < ops.length - 1 && ops[opIdx + 1].op === 'keep' ? ops[opIdx + 1].text : '';

        // Tag op til 20 tegn fra hver side
        const prevAnchor = prevKeep.slice(-20);
        const nextAnchor = nextKeep.slice(0, 20);

        let searchText, replaceText;

        if (op.op === 'replace') {
            searchText  = prevAnchor + op.deleteText + nextAnchor;
            replaceText = prevAnchor + op.insertText + nextAnchor;
        } else if (op.op === 'delete') {
            searchText  = prevAnchor + op.text + nextAnchor;
            replaceText = prevAnchor + nextAnchor;
        } else if (op.op === 'insert') {
            searchText  = prevAnchor + nextAnchor;
            replaceText = prevAnchor + op.text + nextAnchor;
        }

        const ok = await searchAndReplace(para, searchText, replaceText, context);
        if (ok) anyChanges = true;
    }

    return anyChanges;
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

                // Spring overskrifter over
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

                if (!revisedText || revisedText === originalText) continue;

                // Word-level diff
                const aTokens = tokenize(originalText);
                const bTokens = tokenize(revisedText);
                const rawOps = diffTokens(aTokens, bTokens);
                const ops = consolidateOps(rawOps);

                const numChanges = ops.filter(o => o.op !== 'keep').length;
                if (numChanges === 0) continue;

                // Sikkerhedstjek: hvis mindre end halvdelen af afsnittet
                // bevares, har Mistral sandsynligvis omskrevet for meget.
                // Falder tilbage til hel-erstatning så den redigerende
                // tydeligt kan se at noget er gået galt.
                const keepLength = ops.filter(o => o.op === 'keep')
                    .reduce((sum, o) => sum + o.text.length, 0);
                const tooManyChanges = keepLength < originalText.length * 0.5;

                let success = false;
                if (!tooManyChanges) {
                    success = await applyDiffToParagraph(para, ops, context);
                }

                if (!success) {
                    // Fallback: erstat hele afsnittet (gammel adfærd)
                    para.insertText(revisedText, Word.InsertLocation.replace);
                    await context.sync();
                }

                changed++;
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
