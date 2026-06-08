/* global Office, Word */

// =====================================================================
// mistral.js — Auto-redigér + Rubrik-forslag
// POV International Word Add-in
//
// To funktioner:
// 1. Auto-redigér: korrektur via Mistral, indsættes som sporede ændringer
//    med word-level diff så kun de faktisk ændrede ord markeres.
// 2. Foreslå rubrikker: hovedrubrik + mellemrubrikker i journalistisk stil,
//    med valgbar tone (Faktuel, Varm, Humoristisk).
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

// =====================================================================
// RUBRIK-FORSLAG — Tone-specifikke prompts
// =====================================================================

const TONE_BESKRIVELSER = {
    faktuel: `TONE: FAKTUEL
- Brug et alvorligt, fagligt og nøgternt journalistisk sprog
- Hold dig til kendsgerninger og det centrale i artiklen
- Ingen ordspil, ingen følelsesladet sprog
- Tænk klassiske avis-rubrikker (Politiken, Berlingske, Information)

Eksempler på faktuelle HOVEDRUBRIKKER (op til 15 ord, fyldige nok til at give kontekst):
Én sætning:
- "Brændstofkrise truer europæisk luftfart efter lukning af Hormuzstrædet"
- "IEA-direktør advarer EU om akut mangel på flybrændstof inden sommerferien"
To sætninger:
- "Hormuzstrædet er lukket. Nu står Europas luftfart over for en akut brændstofkrise"
- "Lagrene tømmes hurtigt. Birol advarer om aflyste afgange allerede i juni"

Eksempler på faktuelle MELLEMRUBRIKKER (kort, 3-5 ord, aktive verber):
- "Birol advarer mod brændstofkrise"
- "EU forbereder nødplan"
- "Lufthavne aflyser hundredvis af afgange"
- "Priserne stiger hurtigt"`,

    varm: `TONE: VARM
- Brug et empatisk, menneskeligt og indlevende sprog
- Du må gerne udtrykke følelserne der ligger i artiklen
- Egnet til livsstil, personportrætter, og emner med menneskelig kerne
- Tænk magasin- og featurejournalistik

Eksempler på varme HOVEDRUBRIKKER (op til 15 ord, fyldige nok til at give kontekst):
Én sætning:
- "Hun mistede sin søn — nu bygger hun et fristed for andre forældre i sorg"
- "Tre generationer i samme køkken og opskriften der binder familien sammen"
To sætninger:
- "Hun troede livet var slut. I haven fandt hun en ny begyndelse"
- "Landsbyen mistede sin købmand. Så åbnede naboerne deres egne døre"

Eksempler på varme MELLEMRUBRIKKER (kort, 3-5 ord, aktive verber):
- "Familien venter stadig på svar"
- "Hun finder ro i haven"
- "Naboer rækker hånden frem"
- "Drømmen vokser hver dag"`,

    humoristisk: `TONE: HUMORISTISK
- Brug gerne ordspil, dobbelttydigheder eller let humor
- Stadig respektfuld over for emnet — humor må aldrig blive på bekostning af artiklens substans
- Egnet til kulturstof, kommentarer, lette emner
- Tænk klummeagtig tone

Eksempler på humoristiske HOVEDRUBRIKKER (op til 15 ord, fyldige nok til at give kontekst):
Én sætning:
- "Når Mellemøsten nyser, får europæiske lufthavne lungebetændelse"
- "Sommerferien hænger i en tynd tråd — og en endnu tyndere brændstofslange"
To sætninger:
- "Hormuzstrædet siger nej tak. Og så pakker sommerferien sin kuffert ud igen"
- "Flyene stod klar. Brændstoffet glemte at møde op"

Eksempler på humoristiske MELLEMRUBRIKKER (kort, 3-5 ord, aktive verber):
- "Brændstoffet damper væk"
- "EU-Kommissionen sveder over kortet"
- "Flyene står stille — og stille"
- "Sommerferien går i stå"`,

    kreativ: `TONE: KREATIV
- Tag chancer — overrask læseren med uventede vinkler, billedsprog eller utraditionelle formuleringer
- Du må gerne være metaforisk, lyrisk, eller bryde med klassiske journalistiske mønstre
- Tonen kan skifte mellem alvor og leg — det vigtigste er at rubrikken har en distinkt stemme
- Egnet til essays, portrætter, kulturstof, og artikler der fortjener noget særligt
- Tænk litterær journalistik, magasin-essays, eller stærke kommentarer

Eksempler på kreative HOVEDRUBRIKKER (op til 15 ord, fyldige nok til at give kontekst):
Én sætning:
- "Et stræde lukker, og sommerens drømme falder som dominobrikker over Europa"
- "Brændstoffet forsvandt — og med det forsvandt også illusionen om en grænseløs verden"
To sætninger:
- "Et smalt stræde. Et enormt problem"
- "Engang fløj vi for at undslippe verden. Nu kan verden ikke længere undslippe os"

Eksempler på kreative MELLEMRUBRIKKER (kort, 3-5 ord, aktive verber):
- "Strædet kvæler kontinentet"
- "Tankene tørrer ud"
- "Birol tegner kortet om"
- "Sommeren venter i lufthavnen"`
};

// Temperatur pr. tone. De tre første beholder samme moderate temperatur;
// "kreativ" får højere temperatur for mere overraskende output.
const TONE_TEMPERATUR = {
    faktuel: 0.4,
    varm: 0.4,
    humoristisk: 0.4,
    kreativ: 0.8
};

function buildRubrikPrompt(tone, harEksisterendeMellemrubrikker, antalEksisterendeMellemrubrikker) {
    const toneBeskrivelse = TONE_BESKRIVELSER[tone] || TONE_BESKRIVELSER.faktuel;

    let mellemrubrikInstruks;
    if (harEksisterendeMellemrubrikker) {
        mellemrubrikInstruks = `Artiklen indeholder allerede ${antalEksisterendeMellemrubrikker} mellemrubrik(ker). Du skal foreslå PRÆCIS ${antalEksisterendeMellemrubrikker} nye mellemrubrikker — én pr. eksisterende mellemrubrik, i samme rækkefølge. Hver ny mellemrubrik skal beskrive den samme tekstsektion som den eksisterende mellemrubrik introducerer.`;
    } else {
        mellemrubrikInstruks = `Artiklen har INGEN eksisterende mellemrubrikker. Foreslå mellemrubrikker så de bryder teksten op i passende sektioner.

OBLIGATORISK AFSTANDSREGEL:
- Der SKAL være MINIMUM 5 paragraphs fra dokumentets start (eller hovedrubrikken) til den FØRSTE mellemrubrik — artiklen skal have plads til at etablere sig før første brud
- Der SKAL være MINIMUM 4 paragraphs mellem to mellemrubrikker
- Der MÅ MAKSIMUM være 6 paragraphs mellem to mellemrubrikker
- Tæl kun brødtekst-paragraphs — hovedrubrikker og eksisterende overskrifter tæller ikke

Eksempel på korrekt afstand: hvis artiklen har 20 brødtekst-paragraphs, vil et godt resultat være mellemrubrikker ved paragraph 6, 12 og 18 (første kommer EFTER paragraph 5, derefter afstande på 5-6).

Hver mellemrubrik skal beskrive det afsnit der følger lige efter den. Du skal angive efter hvilken paragraph-index (0-baseret) hver mellemrubrik skal indsættes.`;
    }

    return `Du er en erfaren redaktør på det danske nyhedsmedie POV International. Din opgave er at foreslå rubrikker (overskrifter) til den artikel, du modtager.

${toneBeskrivelse}

KRAV TIL HOVEDRUBRIK:
- Maksimum 15 ord, gerne færre — men typisk 8-15 ord, IKKE så kort som en mellemrubrik
- Må gerne være én eller to sætninger
- Skal virke inviterende og have blikfang
- Skal vække nysgerrighed hos læseren
- Skal afspejle artiklens kerne
- Journalistisk stil tilpasset den valgte tone
- VIGTIGT: Hovedrubrikken må ALDRIG forveksles med en mellemrubrik. Mellemrubrikker er 3-5 ord; hovedrubrikken er fyldigere og giver mere kontekst.

KRAV TIL MELLEMRUBRIKKER:
- Maksimum 5 ord pr. mellemrubrik
- Skal indeholde et BØJET UDSAGNSORD (finit verbum) som det centrale element
- Skal beskrive en HANDLING eller UDVIKLING fra det følgende afsnit — ikke et tema eller en kategori
- Samme tone som hovedrubrikken

KRITISK — undgå "liste-agtige" mellemrubrikker:
Mellemrubrikker må IKKE være rene substantiv-fraser der fungerer som indholdsfortegnelse.
De skal læses som mini-overskrifter med et tydeligt subjekt og verbum.

UNDGÅ (substantiv-fraser, virker som kategorier):
- "Krigens konsekvenser"
- "Mellemøstlige spændinger"
- "Den økonomiske påvirkning"
- "Birol og IEA's advarsel"

BRUG (aktive sætninger med bøjet udsagnsord):
- "Krigen rammer flytrafikken"
- "Spændingen vokser i Mellemøsten"
- "Økonomien mærker konsekvenserne"
- "Birol advarer EU-ledere"

Tommelfingerregel: hvis du kan sætte "om" foran din mellemrubrik og den giver mening som indholdsfortegnelse-punkt ("om krigens konsekvenser"), så er den for passiv. Lav den om så den indeholder et bøjet verbum.

${mellemrubrikInstruks}

SVARFORMAT:
Du SKAL svare med ren JSON i præcis dette format — intet andet, ingen markdown, ingen indledning:

{
  "hovedrubrik": "Forslag til hovedrubrik her",
  "mellemrubrikker": [
    {"efter_paragraph_index": 2, "tekst": "Mellemrubrik 1"},
    {"efter_paragraph_index": 5, "tekst": "Mellemrubrik 2"}
  ]
}

Hvis der ingen eksisterende mellemrubrikker er: brug efter_paragraph_index til at angive hvor i teksten mellemrubrikken skal indsættes (0-baseret index på den paragraph den skal stå FØR).
Hvis der ER eksisterende mellemrubrikker: efter_paragraph_index ignoreres — mellemrubrikkerne placeres automatisk over de eksisterende i den rækkefølge du leverer dem.

Returner KUN JSON — ingen forklaringer, ingen markdown-kodeblok, ingen tekst før eller efter.`;
}

// ── Elementer ────────────────────────────────────────────────────────
let autoRedigérBtn;
let autoRedigérText;
let korrekturStatus;
let mistralApiKeyInput;
let saveApiKeyBtn;
let apiKeyStatus;

// Rubrik-elementer
let rubrikBtn;
let rubrikText;
let rubrikStatus;
let rubrikToneSelect;

// ── Init — kaldes fra taskpane.js via initMistralUI() ────────────────
function initMistralUI() {
    autoRedigérBtn     = document.getElementById("autoredigér-btn");
    autoRedigérText    = document.getElementById("autoredigér-text");
    korrekturStatus    = document.getElementById("korrektur-status");
    mistralApiKeyInput = document.getElementById("mistral-api-key");
    saveApiKeyBtn      = document.getElementById("save-api-key-btn");
    apiKeyStatus       = document.getElementById("api-key-status");

    // Rubrik-UI
    rubrikBtn        = document.getElementById("rubrik-btn");
    rubrikText       = document.getElementById("rubrik-text");
    rubrikStatus     = document.getElementById("rubrik-status");
    rubrikToneSelect = document.getElementById("rubrik-tone");

    if (!autoRedigérBtn) return;

    if (localStorage.getItem(MISTRAL_KEY_STORAGE)) {
        mistralApiKeyInput.placeholder = "••••••••••••••••••••• (gemt)";
        setStatus(apiKeyStatus, "✓ API-nøgle er gemt", "success");
    }

    autoRedigérBtn.addEventListener("click", autoRedigér);

    if (rubrikBtn) {
        rubrikBtn.addEventListener("click", foreslåRubrikker);
    }

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
function tokenize(text) {
    const re = /(\s+|[.,;:!?"”“„''()\[\]—–\-]|\S+)/g;
    return text.match(re) || [];
}

// ── Word-level diff (LCS-baseret) ────────────────────────────────────
function diffTokens(aTokens, bTokens) {
    const n = aTokens.length;
    const m = bTokens.length;

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
async function applyDiffToParagraph(para, ops, context) {
    let anyChanges = false;

    for (let opIdx = 0; opIdx < ops.length; opIdx++) {
        const op = ops[opIdx];
        if (op.op === 'keep') continue;

        const prevKeep = opIdx > 0 && ops[opIdx - 1].op === 'keep' ? ops[opIdx - 1].text : '';
        const nextKeep = opIdx < ops.length - 1 && ops[opIdx + 1].op === 'keep' ? ops[opIdx + 1].text : '';

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
async function callMistral(text, apiKey, systemPrompt, opts = {}) {
    const body = {
        model: opts.model || "mistral-medium-latest",
        temperature: opts.temperature !== undefined ? opts.temperature : 0,
        messages: [
            { role: "system", content: systemPrompt },
            { role: "user",   content: text }
        ]
    };

    // Brug JSON mode hvis ønsket (for rubrik-forslag)
    if (opts.jsonMode) {
        body.response_format = { type: "json_object" };
    }

    const response = await fetch("https://api.mistral.ai/v1/chat/completions", {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
            "Authorization": "Bearer " + apiKey
        },
        body: JSON.stringify(body)
    });

    if (!response.ok) {
        const err = await response.json().catch(() => ({}));
        const msg = err.message || response.statusText;
        throw new Error(`${response.status}: ${msg}`);
    }

    const data = await response.json();
    return data.choices[0].message.content.trim();
}

// ── Hoved-funktion: Auto-redigér ─────────────────────────────────────
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

            context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
            await context.sync();

            let changed = 0;

            for (let i = 0; i < items.length; i++) {
                const para = items[i];
                para.load("text, styleBuiltIn");
                await context.sync();

                const originalText = para.text.trim();

                if (!originalText || originalText.length < 15) continue;

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

                const aTokens = tokenize(originalText);
                const bTokens = tokenize(revisedText);
                const rawOps = diffTokens(aTokens, bTokens);
                const ops = consolidateOps(rawOps);

                const numChanges = ops.filter(o => o.op !== 'keep').length;
                if (numChanges === 0) continue;

                const keepLength = ops.filter(o => o.op === 'keep')
                    .reduce((sum, o) => sum + o.text.length, 0);
                const tooManyChanges = keepLength < originalText.length * 0.5;

                let success = false;
                if (!tooManyChanges) {
                    success = await applyDiffToParagraph(para, ops, context);
                }

                if (!success) {
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

// =====================================================================
// RUBRIK-FORSLAG
// =====================================================================

// Hjælpefunktion: er denne style en hovedrubrik (Heading 1 eller Title)?
function erHovedrubrikStyle(styleBuiltIn) {
    return styleBuiltIn === Word.BuiltInStyleName.heading1 ||
           styleBuiltIn === Word.BuiltInStyleName.title;
}

// Hjælpefunktion: er denne style en mellemrubrik (Heading 2 eller 3)?
function erMellemrubrikStyle(styleBuiltIn) {
    return styleBuiltIn === Word.BuiltInStyleName.heading2 ||
           styleBuiltIn === Word.BuiltInStyleName.heading3;
}

// Filtrer mellemrubrikker så afstandsreglen (min 4, max 7 paragraphs) overholdes.
// Kaldes KUN når der ikke er eksisterende mellemrubrikker — så har Mistral selv
// valgt placeringerne, og vi skal validere dem.
//
// Strategi:
// 1. Sortér Mistrals forslag efter paragraph-index
// 2. Filtrer dem der ligger for tæt (< 4) på den forrige accepterede
// 3. Hvis hul > 7 mellem to accepterede, kan vi ikke selv lave nye mellemrubrikker
//    (vi har ingen tekst at basere dem på), men vi advarer i konsollen
//
// MIN_AFSTAND og MAX_AFSTAND er antal BRØDTEKST-paragraphs imellem, ikke
// absolutte paragraph-indeks. Vi bruger paragrafTekster til at omregne.
function filtrerMellemrubrikkerMedAfstand(mellemrubrikker, paragrafTekster) {
    const MIN_AFSTAND_FRA_START = 5;  // første mellemrubrik må tidligst komme efter 5 brødtekst-paragraphs
    const MIN_AFSTAND = 4;             // mellem to mellemrubrikker
    const MAX_AFSTAND = 6;             // maks mellem to mellemrubrikker

    // Byg en mapping: paragraph-index → "brødtekst-position"
    // Kun brødtekst (ikke-headings) tæller med i afstandsberegningen
    const brødtekstPositioner = {};
    let brødtekstCount = 0;
    for (const p of paragrafTekster) {
        if (!p.isHovedrubrik && !p.isMellemrubrik) {
            brødtekstPositioner[p.index] = brødtekstCount;
            brødtekstCount++;
        }
    }

    // Sortér forslag efter paragraph-index
    const sorterede = [...mellemrubrikker]
        .filter(m => typeof m.efter_paragraph_index === "number")
        .sort((a, b) => a.efter_paragraph_index - b.efter_paragraph_index);

    const accepterede = [];
    // Brug en sentinel-værdi der gør at første mellemrubrik først tillades
    // fra brødtekst-position MIN_AFSTAND_FRA_START (dvs. efter mindst 5 brødtekst-paragraphs).
    // Når accepterede.length === 0, bruger vi MIN_AFSTAND_FRA_START; ellers MIN_AFSTAND.
    let sidstAccepteretBrødtekstPos = -1;

    for (const forslag of sorterede) {
        const paraIdx = forslag.efter_paragraph_index;

        // Find brødtekst-positionen for denne paragraph
        // (Hvis paragraphen ikke er brødtekst, find nærmeste følgende brødtekst)
        let brødtekstPos = brødtekstPositioner[paraIdx];
        if (brødtekstPos === undefined) {
            // Mistral peger på en heading-paragraph — find næste brødtekst
            for (const p of paragrafTekster) {
                if (p.index >= paraIdx && !p.isHovedrubrik && !p.isMellemrubrik) {
                    brødtekstPos = brødtekstPositioner[p.index];
                    break;
                }
            }
        }
        if (brødtekstPos === undefined) continue; // ingen brødtekst efter dette punkt

        // Første mellemrubrik: kræver MIN_AFSTAND_FRA_START fra start (position 0)
        // Efterfølgende: kræver MIN_AFSTAND fra forrige accepterede
        let kravOpfyldt;
        if (accepterede.length === 0) {
            kravOpfyldt = brødtekstPos >= MIN_AFSTAND_FRA_START;
            if (!kravOpfyldt) {
                console.info(
                    `Filtrerede første mellemrubrik fra (for tæt på start): "${forslag.tekst}" ` +
                    `ved brødtekst-position ${brødtekstPos} < ${MIN_AFSTAND_FRA_START}`
                );
            }
        } else {
            const afstand = brødtekstPos - sidstAccepteretBrødtekstPos;
            kravOpfyldt = afstand >= MIN_AFSTAND;
            if (!kravOpfyldt) {
                console.info(
                    `Filtrerede mellemrubrik fra (for tæt på forrige): "${forslag.tekst}" ` +
                    `ved paragraph ${paraIdx} — afstand ${afstand} < ${MIN_AFSTAND}`
                );
            }
        }

        if (kravOpfyldt) {
            accepterede.push(forslag);
            sidstAccepteretBrødtekstPos = brødtekstPos;
        }
    }

    // Tjek for huller > MAX_AFSTAND og log det (vi kan ikke selv generere ny tekst)
    for (let i = 0; i < accepterede.length - 1; i++) {
        const a = brødtekstPositioner[accepterede[i].efter_paragraph_index];
        const b = brødtekstPositioner[accepterede[i + 1].efter_paragraph_index];
        if (a !== undefined && b !== undefined && (b - a) > MAX_AFSTAND) {
            console.warn(
                `Hul mellem mellemrubrikker er ${b - a} paragraphs ` +
                `(maks. ${MAX_AFSTAND}) — Mistral har ikke foreslået nok mellemrubrikker.`
            );
        }
    }

    return accepterede;
}

// ── Hale-justering ───────────────────────────────────────────────────
// Kaldes KUN når der ikke er eksisterende mellemrubrikker.
// Beslutningstræ:
//   1. Består halen af flere end 6 paragraphs?
//      → Nej: gør intet
//      → Ja: gå til 2
//   2. Består sektionen før sidste mellemrubrik af 6 paragraphs?
//      → Ja: gør intet (vi kan ikke optimere uden at bryde min-reglen)
//      → Nej: gå til 3
//   3. Er sektionen før sidste mellemrubrik ≤ 5 paragraphs?
//      → Ja: flyt sidste mellemrubrik ned så sektionen før bliver præcis 6,
//            og kald Mistral igen for at omformulere mellemrubrikken så den
//            passer til den nye (kortere) hale-tekst
async function justérHaleHvisNødvendigt(forslag, paragrafTekster, tone, apiKey) {
    const MAX_HALE = 6;
    const MÅL_SEKTION_FØR = 6;

    const mellemrubrikker = forslag.mellemrubrikker || [];
    if (mellemrubrikker.length === 0) return;

    // Byg brødtekst-positioner (samme logik som i filteret)
    const brødtekstPositioner = {};
    const brødtekstParagraphs = []; // [{originalIndex, brødtekstPos, text}, ...]
    let brødtekstCount = 0;
    for (const p of paragrafTekster) {
        if (!p.isHovedrubrik && !p.isMellemrubrik) {
            brødtekstPositioner[p.index] = brødtekstCount;
            brødtekstParagraphs.push({
                originalIndex: p.index,
                brødtekstPos: brødtekstCount,
                text: p.text
            });
            brødtekstCount++;
        }
    }

    const totalBrødtekst = brødtekstCount;
    const sidsteMellem = mellemrubrikker[mellemrubrikker.length - 1];
    const sidstePos = brødtekstPositioner[sidsteMellem.efter_paragraph_index];
    if (sidstePos === undefined) return;

    // Trin 1: Beregn hale-længde
    const haleLængde = totalBrødtekst - sidstePos - 1;
    if (haleLængde <= MAX_HALE) {
        console.info(`Hale OK: ${haleLængde} paragraphs (≤ ${MAX_HALE}).`);
        return; // intet at gøre
    }

    // Trin 2: Beregn sektionen før sidste mellemrubrik
    const næstSidstePos = mellemrubrikker.length > 1
        ? brødtekstPositioner[mellemrubrikker[mellemrubrikker.length - 2].efter_paragraph_index]
        : -1; // ingen tidligere mellemrubrik → start af dokumentet
    const sektionFørLængde = sidstePos - næstSidstePos - 1;

    if (sektionFørLængde === MÅL_SEKTION_FØR) {
        console.info(
            `Hale ${haleLængde} paragraphs accepteret — sektion før sidste mellemrubrik ` +
            `er allerede ${sektionFørLængde}. Kan ikke optimere.`
        );
        return;
    }

    // Trin 3: Er sektionen før ≤ 5? (skulle altid være sandt nu, men vi tjekker)
    if (sektionFørLængde > 5) {
        console.warn(
            `Uventet: sektion før er ${sektionFørLængde} (>5 men ≠ 6). ` +
            `Springer hale-justering over.`
        );
        return;
    }

    // Flyt sidste mellemrubrik: ny brødtekst-position = næstSidstePos + 1 + 6 = næstSidstePos + 7
    // (præcis 6 paragraphs i sektionen før, så mellemrubrikken står FØR paragraph nummer 7)
    const nyBrødtekstPos = næstSidstePos + 1 + MÅL_SEKTION_FØR;

    // Find original paragraph-index for den nye position
    const nyBrødtekst = brødtekstParagraphs.find(b => b.brødtekstPos === nyBrødtekstPos);
    if (!nyBrødtekst) {
        console.warn(`Kunne ikke finde brødtekst på position ${nyBrødtekstPos}.`);
        return;
    }

    const gammelPos = sidsteMellem.efter_paragraph_index;
    const gammelTekst = sidsteMellem.tekst;
    sidsteMellem.efter_paragraph_index = nyBrødtekst.originalIndex;

    console.info(
        `Flytter sidste mellemrubrik: paragraph ${gammelPos} → ${nyBrødtekst.originalIndex} ` +
        `(sektion før: ${sektionFørLængde} → ${MÅL_SEKTION_FØR}, hale: ${haleLængde} → ${totalBrødtekst - nyBrødtekstPos - 1})`
    );

    // Kald Mistral igen for at omformulere rubrikken så den passer til den nye hale-tekst
    const haleTekst = brødtekstParagraphs
        .filter(b => b.brødtekstPos >= nyBrødtekstPos)
        .map(b => b.text)
        .join("\n\n");

    if (!haleTekst.trim()) {
        console.warn("Hale-tekst er tom — beholder oprindelig rubrik.");
        return;
    }

    try {
        const toneBeskrivelse = TONE_BESKRIVELSER[tone] || TONE_BESKRIVELSER.faktuel;
        const omformuleringsPrompt = `Du er en erfaren redaktør på POV International. Skriv ÉN mellemrubrik der beskriver den tekstsektion du modtager.

${toneBeskrivelse}

KRAV:
- Maksimum 5 ord
- Skal indeholde et BØJET UDSAGNSORD (finit verbum) som det centrale element
- Skal beskrive en handling eller udvikling fra teksten — ikke et tema eller en kategori
- Må ALDRIG være en ren substantiv-frase (fx "Krigens konsekvenser") — BRUG aktive sætninger (fx "Krigen rammer flytrafikken")

SVARFORMAT:
Returnér KUN selve mellemrubrikken som ren tekst — ingen anførselstegn, ingen forklaringer, ingen markdown, ingen indledning.`;

        const valgtTemperatur = TONE_TEMPERATUR[tone] !== undefined ? TONE_TEMPERATUR[tone] : 0.4;
        const nyRubrik = await callMistral(haleTekst, apiKey, omformuleringsPrompt, {
            temperature: valgtTemperatur
        });

        // Ryd op: fjern evt. anførselstegn eller markdown der måtte være sneget sig ind
        const renset = nyRubrik
            .replace(/^["“„'`*#\s]+|["”'`*\s]+$/g, "")
            .trim();

        if (renset && renset.split(/\s+/).length <= 8) {
            sidsteMellem.tekst = renset;
            console.info(`Omformuleret: "${gammelTekst}" → "${renset}"`);
        } else {
            console.warn(`Omformulering forkastet (for lang eller tom): "${renset}". Beholder original.`);
        }
    } catch (err) {
        console.warn("Omformulering fejlede — beholder oprindelig rubrik:", err.message);
    }
}

// Hoved-funktion: foreslå rubrikker
async function foreslåRubrikker() {
    const apiKey = localStorage.getItem(MISTRAL_KEY_STORAGE);
    if (!apiKey) {
        setStatus(rubrikStatus,
            "Ingen API-nøgle. Åbn 'API-nøgle (Mistral)' under Auto-redigér og gem din nøgle.",
            "error");
        return;
    }

    const tone = rubrikToneSelect ? rubrikToneSelect.value : "faktuel";

    rubrikBtn.disabled = true;
    rubrikText.textContent = "Læser artiklen…";
    setStatus(rubrikStatus, "", "info");

    try {
        // Trin 1: Læs dokumentet og kortlæg eksisterende rubrikker
        let dokumentTekst = "";
        let eksisterendeHovedrubrik = null;  // {paraIndex, text}
        let eksisterendeMellemrubrikker = []; // [{paraIndex, text}, ...]
        let paragrafTekster = [];             // [{index, text, isHeading}, ...]

        await Word.run(async (context) => {
            const paragraphs = context.document.body.paragraphs;
            paragraphs.load("items");
            await context.sync();

            const items = paragraphs.items;
            for (let i = 0; i < items.length; i++) {
                items[i].load("text, styleBuiltIn");
            }
            await context.sync();

            for (let i = 0; i < items.length; i++) {
                const para = items[i];
                const text = para.text.trim();
                const style = para.styleBuiltIn || "";

                if (!text) continue;

                const isHovedrubrik = erHovedrubrikStyle(style);
                const isMellemrubrik = erMellemrubrikStyle(style);

                paragrafTekster.push({
                    index: i,
                    text: text,
                    isHovedrubrik: isHovedrubrik,
                    isMellemrubrik: isMellemrubrik
                });

                if (isHovedrubrik && !eksisterendeHovedrubrik) {
                    eksisterendeHovedrubrik = { paraIndex: i, text: text };
                } else if (isMellemrubrik) {
                    eksisterendeMellemrubrikker.push({ paraIndex: i, text: text });
                }
            }
        });

        if (paragrafTekster.length === 0) {
            setStatus(rubrikStatus, "Dokumentet er tomt.", "error");
            return;
        }

        // Byg artikelteksten der sendes til Mistral, med paragraph-indeks som
        // anker så modellen kan referere til positioner præcist
        dokumentTekst = paragrafTekster.map((p, idx) => {
            let prefix = `[Paragraph ${idx}]`;
            if (p.isHovedrubrik) prefix += " (HOVEDRUBRIK)";
            else if (p.isMellemrubrik) prefix += " (MELLEMRUBRIK)";
            return `${prefix}: ${p.text}`;
        }).join("\n\n");

        // Trin 2: Byg prompt og kald Mistral
        rubrikText.textContent = "Genererer forslag…";
        setStatus(rubrikStatus, `Mistral genererer rubrikker (tone: ${tone})…`, "info");

        const harEksisterendeMellemrubrikker = eksisterendeMellemrubrikker.length > 0;
        const prompt = buildRubrikPrompt(
            tone,
            harEksisterendeMellemrubrikker,
            eksisterendeMellemrubrikker.length
        );

        const valgtTemperatur = TONE_TEMPERATUR[tone] !== undefined ? TONE_TEMPERATUR[tone] : 0.4;
        const rawResponse = await callMistral(dokumentTekst, apiKey, prompt, {
            jsonMode: true,
            temperature: valgtTemperatur
        });

        // Trin 3: Parse JSON-svaret
        let forslag;
        try {
            // Fjern evt. markdown-kodeblokke
            const cleaned = rawResponse.replace(/```json\s*|\s*```/g, "").trim();
            forslag = JSON.parse(cleaned);
        } catch (parseErr) {
            console.error("Kunne ikke parse Mistral-svar som JSON:", rawResponse);
            throw new Error("Mistral returnerede ugyldigt format. Prøv igen.");
        }

        if (!forslag.hovedrubrik) {
            throw new Error("Intet hovedrubrik-forslag modtaget.");
        }

        // Hvis der ikke er eksisterende mellemrubrikker: håndhæv 4-7 afstandsregel
        // (Hvis der ER eksisterende, er antallet bundet og placeringen givet — så
        // springer vi filteret over.)
        if (!harEksisterendeMellemrubrikker && forslag.mellemrubrikker) {
            const før = forslag.mellemrubrikker.length;
            forslag.mellemrubrikker = filtrerMellemrubrikkerMedAfstand(
                forslag.mellemrubrikker,
                paragrafTekster
            );
            const efter = forslag.mellemrubrikker.length;
            if (efter < før) {
                console.info(`Filter: ${før} → ${efter} mellemrubrikker (afstandsregel)`);
            }

            // Hale-justering: tjek om sidste sektion er for lang, og flyt evt.
            // sidste mellemrubrik + omformulér via nyt Mistral-kald
            rubrikText.textContent = "Justerer hale…";
            await justérHaleHvisNødvendigt(forslag, paragrafTekster, tone, apiKey);
        }

        // Trin 4: Indsæt forslagene i dokumentet
        rubrikText.textContent = "Indsætter forslag…";
        await indsætRubrikker(
            forslag,
            eksisterendeHovedrubrik,
            eksisterendeMellemrubrikker,
            paragrafTekster
        );

        const antalMellem = forslag.mellemrubrikker ? forslag.mellemrubrikker.length : 0;
        setStatus(rubrikStatus,
            `✓ Forslag indsat: 1 hovedrubrik + ${antalMellem} mellemrubrik(ker). Gennemgå med Acceptér/Afvis.`,
            "success");

    } catch (err) {
        console.error("Rubrik-forslag fejl:", err);
        const msg = err.message || "";
        if (msg.includes("401")) {
            setStatus(rubrikStatus,
                "Ugyldig API-nøgle. Tjek nøglen under 'API-nøgle (Mistral)'.",
                "error");
        } else if (msg.includes("429")) {
            setStatus(rubrikStatus,
                "Rate limit nået. Vent et øjeblik og prøv igen.",
                "error");
        } else {
            setStatus(rubrikStatus, "Fejl: " + msg, "error");
        }
    } finally {
        rubrikBtn.disabled = false;
        rubrikText.textContent = "✦ Foreslå rubrikker";
    }
}

// ── Indsæt rubrikker i dokumentet som sporede ændringer ──────────────
async function indsætRubrikker(forslag, eksisterendeHovedrubrik, eksisterendeMellemrubrikker, paragrafTekster) {
    await Word.run(async (context) => {
        // Aktiver tracked changes så indsættelserne markeres
        context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
        await context.sync();

        const paragraphs = context.document.body.paragraphs;
        paragraphs.load("items");
        await context.sync();
        const items = paragraphs.items;

        // ── Mellemrubrikker først (fra slutningen og bagud, så indeks ikke flytter sig) ──
        const mellemrubrikker = forslag.mellemrubrikker || [];

        if (eksisterendeMellemrubrikker.length > 0) {
            // Eksisterende mellemrubrikker findes — placer forslag OVER hver af dem.
            // Vi går baglæns så vi ikke forstyrrer paragraph-indeks for ikke-behandlede.
            const antal = Math.min(mellemrubrikker.length, eksisterendeMellemrubrikker.length);
            for (let k = antal - 1; k >= 0; k--) {
                const eksisterende = eksisterendeMellemrubrikker[k];
                const nyTekst = mellemrubrikker[k].tekst;
                const targetPara = items[eksisterende.paraIndex];
                if (!targetPara) continue;

                // Indsæt nyt paragraph FØR eksisterende mellemrubrik
                const newPara = targetPara.insertParagraph(nyTekst, Word.InsertLocation.before);
                newPara.styleBuiltIn = Word.BuiltInStyleName.heading2;
                await context.sync();
            }
        } else {
            // Ingen eksisterende mellemrubrikker — brug efter_paragraph_index fra Mistral.
            // Sortér descending så vi indsætter bagfra og bevarer indekserne.
            const sortedMellem = [...mellemrubrikker]
                .filter(m => typeof m.efter_paragraph_index === "number")
                .sort((a, b) => b.efter_paragraph_index - a.efter_paragraph_index);

            for (const m of sortedMellem) {
                const targetPara = items[m.efter_paragraph_index];
                if (!targetPara) continue;

                // Vi indsætter FØR den paragraph mellemrubrikken skal beskrive
                const newPara = targetPara.insertParagraph(m.tekst, Word.InsertLocation.before);
                newPara.styleBuiltIn = Word.BuiltInStyleName.heading2;
                await context.sync();
            }
        }

        // ── Hovedrubrik til sidst (indeks kan have flyttet sig, men vi placerer enten
        // over eksisterende eller helt øverst, så det er ok) ──
        const hovedrubrikTekst = forslag.hovedrubrik;

        // Genindlæs paragraphs for at få aktuel state
        const refreshedParagraphs = context.document.body.paragraphs;
        refreshedParagraphs.load("items");
        await context.sync();
        const refreshedItems = refreshedParagraphs.items;

        if (eksisterendeHovedrubrik && refreshedItems.length > 0) {
            // Find igen den eksisterende hovedrubrik via tekst-match
            // (paraIndex kan have ændret sig pga. mellemrubrik-indsættelser)
            let targetIdx = -1;
            for (let i = 0; i < refreshedItems.length; i++) {
                refreshedItems[i].load("text, styleBuiltIn");
            }
            await context.sync();

            for (let i = 0; i < refreshedItems.length; i++) {
                const p = refreshedItems[i];
                if (erHovedrubrikStyle(p.styleBuiltIn) &&
                    p.text.trim() === eksisterendeHovedrubrik.text) {
                    targetIdx = i;
                    break;
                }
            }

            if (targetIdx >= 0) {
                const targetPara = refreshedItems[targetIdx];
                const newPara = targetPara.insertParagraph(hovedrubrikTekst, Word.InsertLocation.before);
                newPara.styleBuiltIn = Word.BuiltInStyleName.heading1;
                await context.sync();
            } else {
                // Fallback: indsæt øverst i dokumentet
                const firstPara = refreshedItems[0];
                const newPara = firstPara.insertParagraph(hovedrubrikTekst, Word.InsertLocation.before);
                newPara.styleBuiltIn = Word.BuiltInStyleName.heading1;
                await context.sync();
            }
        } else if (refreshedItems.length > 0) {
            // Ingen eksisterende hovedrubrik — indsæt øverst
            const firstPara = refreshedItems[0];
            const newPara = firstPara.insertParagraph(hovedrubrikTekst, Word.InsertLocation.before);
            newPara.styleBuiltIn = Word.BuiltInStyleName.heading1;
            await context.sync();
        }

        await context.sync();
    });
}

// ── Genbruger setStatus fra taskpane.js (tilgængelig globalt) ────────
