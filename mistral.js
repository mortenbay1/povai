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
// Returnerer den fulde prompt-streng med stilguide injiceret,
// eller basis-prompt alene hvis GitHub ikke kan naas.
// Resultatet caches i sessionStorage saa det kun hentes een gang pr. session.
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
// Rydder op i Mistrals output før det skrives til dokumentet:
// 0. Fjerner markdown-formatering (Mistral indsætter det indimellem trods prompt)
// 1. Konverterer rettede anførselstegn til buede (POV's typografi)
// 2. Flytter komma uden for anførselstegn (dansk regel)
function postProcess(text) {
    // 0. Fjern markdown-formatering
    text = text.replace(/\*\*([^*]+)\*\*/g, '$1');              // **fed** → fed
    text = text.replace(/(?<!\*)\*([^*]+)\*(?!\*)/g, '$1');     // *kursiv* → kursiv
    text = text.replace(/`([^`]+)`/g, '$1');                    // `kode` → kode
    text = text.replace(/^#{1,6}\s+/gm, '');                    // # overskrift → overskrift

    // 1. Rettet anførselstegn → buede (gør først, så trin 2 kun behøver håndtere buede)
    text = text.replace(/"([^"]+)"/g, '“$1”');

    // 2. Komma uden for anførselstegn (dansk regel)
    text = text.replace(/,”/g, '”,');

    return text;
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
                    para.insertText(revisedText, Word.InsertLocation.replace);
                    await context.sync();
                    changed++;
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
