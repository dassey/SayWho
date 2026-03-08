/* ============================================================
   app.js — SayWho Speaker Tagger Engine + UI
   ============================================================ */

/* ─── PALETTE ────────────────────────────────────────────── */
const PALETTE = [
    { color: '#ff6b35', bg: 'rgba(255,107,53,0.18)' },    // ember orange
    { color: '#f5c842', bg: 'rgba(245,200,66,0.15)' },    // golden wheat
    { color: '#7ec850', bg: 'rgba(126,200,80,0.15)' },    // sage green
    { color: '#e07840', bg: 'rgba(224,120,64,0.18)' },    // terracotta
    { color: '#50b8a0', bg: 'rgba(80,184,160,0.15)' },    // dusty teal
    { color: '#d4885a', bg: 'rgba(212,136,90,0.18)' },    // clay
    { color: '#f0a060', bg: 'rgba(240,160,96,0.18)' },    // peach flame
    { color: '#d8c050', bg: 'rgba(216,192,80,0.15)' },    // amber
    { color: '#90b840', bg: 'rgba(144,184,64,0.15)' },    // moss
    { color: '#e85535', bg: 'rgba(232,85,53,0.18)' },     // deep rust
    { color: '#c8a030', bg: 'rgba(200,160,48,0.15)' },    // antique gold
    { color: '#70a870', bg: 'rgba(112,168,112,0.15)' },   // faded forest
];
const NOBODY_PAL = { color: '#7a6a58', bg: 'rgba(100,90,75,0.15)' };

function getPal(s) {
    if (s.name === 'Nobody') return NOBODY_PAL;
    const override = speakerColorOverrides.get(s.name);
    if (override) return override;
    return PALETTE[s.colorIndex % PALETTE.length];
}

/* ─── THEMES ─────────────────────────────────────────────── */
const THEMES = {
    'dark-walnut': {
        '--bg': '#1e1a16', '--bg-surface': '#2c2620', '--bg-elevated': '#3a3028',
        '--border': '#5c4e3e', '--border-strong': '#8a7260',
        '--text-primary': '#ede5da', '--text-secondary': '#b0988a', '--text-muted': '#7a6558',
        '--rust': '#a84020', '--rust-neon': '#ff6b35',
        '--ochre': '#8a6a18', '--ochre-neon': '#f5c842',
        '--olive': '#4a5a18', '--olive-neon': '#a8c850',
    },
    'parchment': {
        '--bg': '#ede8de', '--bg-surface': '#e2dbd0', '--bg-elevated': '#d8d0c4',
        '--border': '#b8a888', '--border-strong': '#8a7860',
        '--text-primary': '#2a2018', '--text-secondary': '#5a4e3a', '--text-muted': '#8a7a68',
        '--rust': '#8b3a1e', '--rust-neon': '#b83a10',
        '--ochre': '#7a5c1e', '--ochre-neon': '#9a6800',
        '--olive': '#4a4a1e', '--olive-neon': '#5a7a00',
    },
    'slate-dark': {
        '--bg': '#161820', '--bg-surface': '#20222e', '--bg-elevated': '#282c3a',
        '--border': '#3e4258', '--border-strong': '#5a6080',
        '--text-primary': '#dde0ee', '--text-secondary': '#8890b0', '--text-muted': '#5a6070',
        '--rust': '#8b3a1e', '--rust-neon': '#ff6b35',
        '--ochre': '#8a6a18', '--ochre-neon': '#f5c842',
        '--olive': '#4a5a18', '--olive-neon': '#a8c850',
    },
    'burnt-sienna': {
        '--bg': '#1a1210', '--bg-surface': '#281a16', '--bg-elevated': '#362420',
        '--border': '#5e3830', '--border-strong': '#8a5848',
        '--text-primary': '#f0e0d8', '--text-secondary': '#c09080', '--text-muted': '#806050',
        '--rust': '#8b3a1e', '--rust-neon': '#ff8040',
        '--ochre': '#8a6a18', '--ochre-neon': '#f5c842',
        '--olive': '#4a5a18', '--olive-neon': '#a8c850',
    },
    'black-coffee': {
        '--bg': '#0e0c0a', '--bg-surface': '#1a1612', '--bg-elevated': '#241e18',
        '--border': '#3a3028', '--border-strong': '#5a4e3e',
        '--text-primary': '#e8e0d4', '--text-secondary': '#a09080', '--text-muted': '#685850',
        '--rust': '#8b3a1e', '--rust-neon': '#ff6b35',
        '--ochre': '#8a6a18', '--ochre-neon': '#f5c842',
        '--olive': '#4a5a18', '--olive-neon': '#a8c850',
    },
};

function applyTheme(key) {
    currentTheme = key;
    const vars = THEMES[key];
    if (!vars) return;
    for (const [k, v] of Object.entries(vars)) {
        document.documentElement.style.setProperty(k, v);
    }
    // Persist
    localStorage.setItem('sw_theme', key);
    // Update theme buttons
    document.querySelectorAll('.theme-btn').forEach(b => {
        b.classList.toggle('active', b.dataset.theme === key);
    });
    // Update color pickers to reflect new values
    syncColorPickers();
}

function syncColorPickers() {
    document.querySelectorAll('.color-row input[type=color]').forEach(inp => {
        const varName = inp.dataset.var;
        const val = getComputedStyle(document.documentElement).getPropertyValue(varName).trim();
        try { inp.value = hexFromCss(val); } catch { }
    });
}

function hexFromCss(val) {
    // Convert css color (hex or rgb) → hex
    if (val.startsWith('#')) {
        if (val.length === 4) return '#' + val[1] + val[1] + val[2] + val[2] + val[3] + val[3];
        return val.slice(0, 7);
    }
    return '#888888'; // fallback
}

/* ─── SPEECH ENGINE ──────────────────────────────────────── */
const SPEECH_VERBS = new Set([
    'said', 'says', 'say', 'asked', 'ask', 'asks', 'replied', 'reply', 'replies',
    'answered', 'answer', 'answers', 'shouted', 'shout', 'shouts',
    'whispered', 'whisper', 'whispers', 'called', 'calls', 'call',
    'cried', 'cry', 'cries', 'muttered', 'mutter', 'mutters',
    'added', 'add', 'adds', 'continued', 'continue', 'continues',
    'declared', 'declared', 'announced', 'stated', 'explained', 'noted',
    'responded', 'respond', 'responds', 'interrupted', 'growled', 'growl',
    'cut', 'repeated', 'confirmed', 'told', 'tell', 'tells',
    'breathed', 'snapped', 'snap', 'snaps',
]);

const STOP_WORDS = new Set([
    'the', 'a', 'an', 'and', 'or', 'but', 'not', 'with', 'at', 'to', 'from', 'in', 'on',
    'of', 'for', 'it', 'its', "it's", 'he', 'she', 'they', 'we', 'i', 'you', 'his', 'her',
    'their', 'our', 'my', 'your', 'then', 'as', 'if', 'unless', 'until', 'before',
    'after', 'during', 'while', 'because', 'though', 'although', 'first', 'second',
    'third', 'next', 'finally', 'also', 'just', 'still', 'even', 'other', 'another',
    'there', 'here', 'those', 'these', 'that', 'this', 'which', 'who', 'what', 'when',
    'already', 'suddenly', 'slowly', 'carefully', 'quickly', 'quietly', 'loudly',
    'gently', 'slightly', 'briefly', 'once', 'twice', 'always', 'never', 'often',
    'now', 'later', 'earlier', 'soon', 'almost', 'nearly', 'exactly', 'clearly',
    'simply', 'directly', "he'd", "she'd", "they'd", "they're", "he's", "she's",
    'captain', 'officer', 'director', 'doctor', 'dr', 'mr', 'mrs', 'ms', 'sir',
    "logic's", 'edge', 'ues', 'engineering', 'mars',
]);

const TITLE_PREFIXES = /^(Captain|Captain\s+|Dr\.?\s+|Mr\.?\s+|Mrs\.?\s+|Ms\.?\s+|Director\s+|Lieutenant\s+|Lt\.?\s+|Commander\s+|Chair\s+|Chief\s+|Senior\s+|Engineer\s+|Councilmember\s+|Councilwoman\s+|Councilman\s+)/i;

class SpeakerEngine {
    constructor(text) {
        this.rawText = text;
        this.speakerMap = new Map();
        this.taggedText = '';
        this.segments = [];
    }

    analyze() {
        this._extractSpeakers();
        this._buildTaggedOutput();
        return {
            speakers: [...this.speakerMap.entries()].map(([name, info]) => ({
                name, count: info.count, colorIndex: info.colorIndex,
            })).sort((a, b) => b.count - a.count),
            segments: this.segments,
        };
    }

    _extractSpeakers() {
        const text = this.rawText;
        const ATTR_BEFORE = /\b(\w[\w\s''-]{0,30})\s+(said|asked|replied|answered|shouted|whispered|called|cried|muttered|added|continued|declared|announced|stated|explained|noted|responded|interrupted|growled|repeated|confirmed|told|breathed|snapped|cut)[^,\n]*[,:]?\s+"([^"]{1,500})"/gi;
        const ATTR_AFTER = /"([^"]{1,500})"[,.]?\s{0,3}(said|asked|replied|answered|shouted|whispered|called|cried|muttered|added|continued|declared|announced|stated|explained|noted|responded|interrupted|growled|repeated|confirmed|told|breathed|snapped)[,\s]{0,5}(?:he|she|it|they)?\s*(\w[\w\s''-]{0,24})\b/gi;

        const candidates = new Map();
        let m;
        ATTR_BEFORE.lastIndex = 0;
        while ((m = ATTR_BEFORE.exec(text)) !== null) {
            const name = this._cleanName(m[1].trim());
            if (name) candidates.set(name, (candidates.get(name) || 0) + 1);
        }
        ATTR_AFTER.lastIndex = 0;
        while ((m = ATTR_AFTER.exec(text)) !== null) {
            const name = this._cleanName(m[3].trim());
            if (name) candidates.set(name, (candidates.get(name) || 0) + 1);
        }

        const lines = text.split(/\r?\n/);
        for (const line of lines) {
            const cleaned = line.trim();
            if (!cleaned) continue;
            const lm = cleaned.match(/^([A-Z][a-z]+(?:\s+[A-Z][a-z]+)?)\s+(said|asked|replied|answered|shouted|whispered|called|muttered|added|continued|responded|interrupted|growled|repeated|breathed|snapped|told)\b/);
            if (lm) { const n = this._cleanName(lm[1]); if (n) candidates.set(n, (candidates.get(n) || 0) + 2); }
            const pm = cleaned.match(/[.!?,"']\s+(said|asked|replied|answered|shouted|whispered|called|muttered|added|responded|snapped|told)\s+([A-Z][a-z]+(?:\s+[A-Z][a-z]+)?)\b/);
            if (pm) { const n = this._cleanName(pm[2]); if (n) candidates.set(n, (candidates.get(n) || 0) + 2); }
        }

        let colorIdx = 0;
        for (const [name, count] of [...candidates.entries()].sort((a, b) => b[1] - a[1])) {
            if (count < 1) continue;
            this.speakerMap.set(name, { count, colorIndex: colorIdx % PALETTE.length });
            colorIdx++;
        }
    }

    _cleanName(raw) {
        if (!raw) return null;
        let name = raw.replace(TITLE_PREFIXES, '').trim();
        if (!/^[A-Z]/.test(name)) return null;
        if (/^(A|An|The)\s/i.test(name)) return null;
        const REJECT = /^(Instead|Despite|Although|Because|Before|After|During|While|However|Meanwhile|Therefore|Otherwise|Perhaps|Suddenly|Slowly|Carefully|Then|There|Now|At|In|On|By|With|For|As|No|Not)\b/i;
        if (REJECT.test(name)) return null;
        const parts = name.split(/\s+/).slice(0, 3);
        name = parts.join(' ');
        if (STOP_WORDS.has(name.toLowerCase())) return null;
        if (name.length < 2) return null;
        if (/^[\d\W]+$/.test(name)) return null;
        if (/^[A-Z]{2,}$/.test(name)) return null;
        if (parts.length >= 2 && /^[a-z]/.test(parts[1])) return null;
        return name;
    }

    _buildTaggedOutput() {
        if (this.speakerMap.size === 0) {
            this.taggedText = this.rawText;
            this.segments = [{ type: 'text', content: this.rawText }];
            return;
        }

        const lines = this.rawText.split(/\r?\n/);
        const result = [];
        this.segments = [];

        let lastMentionedChar = null;
        let lastGenderChar = { m: null, f: null };
        let lastMentionedUsedCount = 0;
        const SPEECH_VERBS_RE = /\b(said|asked|replied|answered|shouted|whispered|called|cried|muttered|added|continued|declared|announced|stated|explained|noted|responded|interrupted|growled|repeated|confirmed|told|breathed|snapped|cut|nodded|turned)\b/i;
        const names = [...this.speakerMap.keys()].sort((a, b) => b.length - a.length);

        for (const line of lines) {
            const hasQuote = /"/.test(line);

            if (!hasQuote) {
                this._scanProse(line, names, lastGenderChar, (name) => {
                    lastMentionedChar = name; lastMentionedUsedCount = 0;
                });
                this.segments.push({ type: 'text', content: line });
                result.push(line);
                continue;
            }

            const firstQ = line.indexOf('"');
            if (firstQ > 0) {
                this._scanProse(line.slice(0, firstQ), names, lastGenderChar, (name) => {
                    lastMentionedChar = name; lastMentionedUsedCount = 0;
                });
            }

            const speaker = this._resolve4Tier(line, names, SPEECH_VERBS_RE, lastMentionedChar, lastMentionedUsedCount, lastGenderChar);

            if (speaker && speaker !== 'Nobody') {
                if (/\b(he|his|him)\b/i.test(line)) lastGenderChar.m = speaker;
                if (/\b(she|her|hers)\b/i.test(line)) lastGenderChar.f = speaker;
                if (speaker === lastMentionedChar) lastMentionedUsedCount++;
                else { lastMentionedChar = speaker; lastMentionedUsedCount = 0; }
            }

            const effectiveSpeaker = speaker || 'Nobody';
            const tagged = line.replace(/"([^"]+)"/g, (_, content) => `[${effectiveSpeaker}]${content}[/${effectiveSpeaker}]`);
            this.segments.push({ type: 'speech', speaker: effectiveSpeaker, content: tagged });
            result.push(tagged);
        }
        this.taggedText = result.join('\n');
    }

    _resolve4Tier(line, names, SPEECH_VERBS_RE, lastMentionedChar, lastMentionedUsedCount, lastGenderChar) {
        for (const name of names) {
            const esc = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
            const pre = new RegExp(`\\b${esc}\\b.{0,80}${SPEECH_VERBS_RE.source}`, 'i');
            const post = new RegExp(`${SPEECH_VERBS_RE.source}.{0,40}\\b${esc}\\b`, 'i');
            if (pre.test(line) || post.test(line)) return name;
        }
        if (/\b(he|him)\b.{0,20}(said|asked|replied|answered|shouted|whispered|called|muttered|added|continued|responded|snapped)\b/i.test(line) ||
            /\b(said|asked|replied|answered|shouted|whispered|called|muttered|added|continued|responded|snapped)\b.{0,20}\b(he|him)\b/i.test(line)) {
            if (lastGenderChar.m) return lastGenderChar.m;
        }
        if (/\b(she|her)\b.{0,20}(said|asked|replied|answered|shouted|whispered|called|muttered|added|continued|responded|snapped)\b/i.test(line) ||
            /\b(said|asked|replied|answered|shouted|whispered|called|muttered|added|continued|responded|snapped)\b.{0,20}\b(she|her)\b/i.test(line)) {
            if (lastGenderChar.f) return lastGenderChar.f;
        }
        if (lastMentionedChar && lastMentionedUsedCount === 0) return lastMentionedChar;
        return 'Nobody';
    }

    _scanProse(text, names, lastGenderChar, onNameFound) {
        for (const name of names) {
            const esc = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
            if (new RegExp(`\\b${esc}\\b`).test(text)) {
                if (/\b(he|his|him)\b/i.test(text)) lastGenderChar.m = name;
                if (/\b(she|her|hers)\b/i.test(text)) lastGenderChar.f = name;
                onNameFound(name);
                break;
            }
        }
    }
}

/* ─── FILE READING ───────────────────────────────────────── */
async function readFile(file) {
    const ext = file.name.split('.').pop().toLowerCase();
    if (ext === 'txt' || ext === 'md') return readAsText(file);
    if (ext === 'docx' || ext === 'doc') return readDocx(file);
    if (ext === 'pdf') return readPdf(file);
    return readAsText(file);
}
function readAsText(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = e => resolve(e.target.result);
        reader.onerror = reject;
        reader.readAsText(file, 'utf-8');
    });
}
async function readDocx(file) {
    if (typeof mammoth === 'undefined') throw new Error('mammoth.js not loaded');
    const ab = await file.arrayBuffer();
    const result = await mammoth.extractRawText({ arrayBuffer: ab });
    return result.value;
}
async function readPdf(file) {
    if (typeof pdfjsLib === 'undefined') throw new Error('pdf.js not loaded');
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js';
    const ab = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: ab }).promise;
    const parts = [];
    for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const content = await page.getTextContent();
        parts.push(content.items.map(item => item.str).join(' '));
    }
    return parts.join('\n\n');
}

/* ─── STATE ──────────────────────────────────────────────── */
let currentText = '';
let analysisResult = null;
let activeSpeakers = new Set();
let dialogueOnly = false;
let searchTerm = '';
let searchHitIndex = 0;
let tagOverrides = new Map();       // contentSlice → speakerName
let speakerColorOverrides = new Map(); // name → {color, bg}
let currentTheme = 'dark-walnut';

/* ─── DOM ────────────────────────────────────────────────── */
const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');
const pasteInput = document.getElementById('paste-input');
const clearPasteBtn = document.getElementById('clear-paste-btn');
const analyzeBtn = document.getElementById('analyze-btn');
const resetBtn = document.getElementById('reset-btn');
const resultsSection = document.getElementById('results-section');
const speakersGrid = document.getElementById('speakers-grid');
const statsBar = document.getElementById('stats-bar');
const outputPreview = document.getElementById('output-preview');
const copyBtn = document.getElementById('copy-btn');
const downloadBtn = document.getElementById('download-btn');
const selectAllBtn = document.getElementById('select-all-btn');
const deselectAllBtn = document.getElementById('deselect-all-btn');
const highlightToggle = document.getElementById('highlight-toggle');
const toastContainer = document.getElementById('toast-container');
const dialogueOnlyBtn = document.getElementById('dialogue-only-btn');
const searchInput = document.getElementById('search-input');
const searchClearBtn = document.getElementById('search-clear-btn');
const searchMatchCount = document.getElementById('search-match-count');
const settingsBtn = document.getElementById('settings-btn');
const settingsPanel = document.getElementById('settings-panel');
const settingsClose = document.getElementById('settings-close');
const tagPopup = document.getElementById('tag-popup');
const chipColorMenu = document.getElementById('chip-color-menu');
const popupBackdrop = document.getElementById('popup-backdrop');

/* ─── DRAG & DROP ────────────────────────────────────────── */
['dragenter', 'dragover'].forEach(evt => dropZone.addEventListener(evt, e => { e.preventDefault(); dropZone.classList.add('drag-over'); }));
['dragleave', 'drop'].forEach(evt => dropZone.addEventListener(evt, e => { e.preventDefault(); dropZone.classList.remove('drag-over'); }));
dropZone.addEventListener('drop', async e => { const f = e.dataTransfer.files[0]; if (f) await loadFile(f); });
dropZone.addEventListener('click', () => fileInput.click());
fileInput.addEventListener('change', async () => { const f = fileInput.files[0]; if (f) await loadFile(f); });

async function loadFile(file) {
    showToast(`📖 Loading "${file.name}"…`, 'info');
    try {
        const text = await readFile(file);
        pasteInput.value = text;
        showToast(`✓ "${file.name}" loaded — ${text.length.toLocaleString()} characters`, 'success');
    } catch (err) { showToast(`✗ ${err.message}`, 'error'); }
}

clearPasteBtn.addEventListener('click', () => { pasteInput.value = ''; pasteInput.focus(); });

/* ─── TEXT NORMALIZATION ────────────────────────────────── */
/**
 * Collapse word-wrapped lines (e.g. text pasted from PDFs or editors
 * that insert hard line-breaks every ~70 chars) into proper paragraphs.
 * Real paragraph breaks (blank lines) are preserved.
 * Dialogue that straddles two wrapped lines is rejoined into one line
 * so the attribution regexes can see the full quote + attribution.
 */
function normalizeText(raw) {
    // Standardise endings
    let t = raw.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    // Protect real paragraph breaks (2+ newlines) by a placeholder
    t = t.replace(/\n{2,}/g, '\x00');
    // Collapse single newlines (word-wraps) → space
    t = t.replace(/([^\x00])\n([^\x00])/g, '$1 $2');
    // Tidy up double spaces
    t = t.replace(/ {2,}/g, ' ');
    // Restore paragraph breaks
    t = t.replace(/\x00/g, '\n\n');
    return t.trim();
}

/* ─── ANALYSIS ───────────────────────────────────────────── */
analyzeBtn.addEventListener('click', runAnalysis);

async function runAnalysis() {
    const raw = pasteInput.value.trim();
    if (!raw) { showToast('⚠ Please upload or paste text first.', 'error'); return; }

    // Normalize word-wrapped text so engine sees full lines
    const text = normalizeText(raw);
    currentText = text;
    tagOverrides.clear();
    analyzeBtn.disabled = true;
    analyzeBtn.innerHTML = '<div class="spinner"></div> Analyzing…';
    await new Promise(r => setTimeout(r, 60));

    try {
        const engine = new SpeakerEngine(text);
        analysisResult = engine.analyze();

        if (analysisResult.speakers.length === 0) {
            showToast('No speakers detected. Try text with "Quote," Name said. patterns.', 'error');
            analyzeBtn.disabled = false;
            analyzeBtn.innerHTML = '⏱ Analyze Speakers';
            return;
        }

        // Re-run with full speaker map to get accurate tagging
        const finalEngine = new SpeakerEngine(text);
        finalEngine.speakerMap = new Map(
            analysisResult.speakers.map(s => [s.name, { count: s.count, colorIndex: s.colorIndex }])
        );
        finalEngine._buildTaggedOutput();
        analysisResult.taggedText = finalEngine.taggedText;
        analysisResult.segments = finalEngine.segments;

        // Count Nobody occurrences
        const nobodyMatches = finalEngine.taggedText.match(/\[Nobody\]/g);
        const nobodyCount = nobodyMatches ? nobodyMatches.length : 0;
        if (nobodyCount > 0) {
            analysisResult.speakers.push({ name: 'Nobody', count: nobodyCount, taggedCount: nobodyCount, colorIndex: 99 });
        }

        // Count actual tagged lines per speaker
        const taggedLines = finalEngine.taggedText.split('\n');
        for (const s of analysisResult.speakers) {
            if (s.name === 'Nobody') continue;
            const re = new RegExp(`\\[${s.name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\]`, 'g');
            let cnt = 0;
            for (const ln of taggedLines) { const m = ln.match(re); if (m) cnt += m.length; }
            s.taggedCount = cnt;
        }
        analysisResult.speakers.sort((a, b) => (b.taggedCount ?? b.count) - (a.taggedCount ?? a.count));

        activeSpeakers = new Set(analysisResult.speakers.map(s => s.name));
        renderResults();
        showToast(`✓ Found ${analysisResult.speakers.length} speaker${analysisResult.speakers.length !== 1 ? 's' : ''}`, 'success');
    } catch (err) {
        showToast(`Error: ${err.message}`, 'error');
        console.error(err);
    }
    analyzeBtn.disabled = false;
    analyzeBtn.innerHTML = '<svg width="14" height="14" viewBox="0 0 14 14" fill="none"><circle cx="7" cy="7" r="6" stroke="currentColor" stroke-width="1.5"/><path d="M7 4.5v3l1.5 1.5" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg> Analyze Speakers';
}

/* ─── RENDER ─────────────────────────────────────────────── */
function renderResults() {
    resultsSection.classList.remove('hidden');
    renderSpeakers();
    renderStats();
    renderOutput();
    resultsSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function renderSpeakers() {
    speakersGrid.innerHTML = '';
    for (const s of analysisResult.speakers) {
        const pal = getPal(s);
        const chip = document.createElement('div');
        chip.className = 'speaker-chip active' + (s.name === 'Nobody' ? ' chip--nobody' : '');
        chip.dataset.name = s.name;
        chip.style.setProperty('--chip-color', pal.color);
        chip.style.setProperty('--chip-bg', pal.bg);
        chip.innerHTML = `
          <span class="chip-swatch"></span>
          <span class="chip-name">${escHtml(s.name)}</span>
          <span class="chip-count">${s.taggedCount !== undefined ? s.taggedCount : s.count}</span>
        `;
        chip.addEventListener('click', () => toggleSpeaker(s.name, chip));
        chip.addEventListener('contextmenu', e => { e.preventDefault(); showChipColorMenu(e, s); });
        speakersGrid.appendChild(chip);
    }
}

function toggleSpeaker(name, chip) {
    if (activeSpeakers.has(name)) { activeSpeakers.delete(name); chip.classList.remove('active'); }
    else { activeSpeakers.add(name); chip.classList.add('active'); }
    renderOutput();
}

selectAllBtn.addEventListener('click', () => {
    activeSpeakers = new Set(analysisResult.speakers.map(s => s.name));
    document.querySelectorAll('.speaker-chip').forEach(c => c.classList.add('active'));
    renderOutput();
});
deselectAllBtn.addEventListener('click', () => {
    activeSpeakers.clear();
    document.querySelectorAll('.speaker-chip').forEach(c => c.classList.remove('active'));
    renderOutput();
});

function renderStats() {
    const totalLines = currentText.split('\n').length;
    const totalSpeakers = analysisResult.speakers.length;
    const totalDialogue = analysisResult.speakers.reduce((a, s) => a + (s.taggedCount ?? s.count), 0);
    statsBar.innerHTML = `
      <div class="stat-item"><span class="stat-label">Speakers</span><span class="stat-value accent">${totalSpeakers}</span></div>
      <div class="stat-item"><span class="stat-label">Dialogue Lines</span><span class="stat-value">${totalDialogue}</span></div>
      <div class="stat-item"><span class="stat-label">Total Lines</span><span class="stat-value">${totalLines.toLocaleString()}</span></div>
      <div class="stat-item"><span class="stat-label">Characters</span><span class="stat-value">${currentText.length.toLocaleString()}</span></div>
    `;
}

function renderOutput() {
    const highlight = highlightToggle.checked;
    let taggedText = generateTaggedText();

    if (dialogueOnly) {
        taggedText = taggedText.split('\n').filter(l => /\[[^\]]+\]/.test(l)).join('\n');
    }

    let html = highlight ? buildHighlightedHTML(taggedText) : escHtml(taggedText);

    // Search highlighting
    const term = searchTerm.trim();
    if (term) {
        const safe = term.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        html = html.replace(new RegExp(`(${safe})`, 'gi'), '<mark class="search-hit">$1</mark>');
    }

    outputPreview.innerHTML = html;

    if (term) {
        const hits = outputPreview.querySelectorAll('.search-hit');
        if (!hits.length) { searchMatchCount.textContent = 'No matches'; searchMatchCount.style.color = 'var(--rust-neon)'; }
        else {
            searchHitIndex = Math.min(searchHitIndex, hits.length - 1);
            hits.forEach((h, i) => h.classList.toggle('current', i === searchHitIndex));
            hits[searchHitIndex]?.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
            searchMatchCount.textContent = `${searchHitIndex + 1} / ${hits.length}`;
            searchMatchCount.style.color = '';
        }
        searchClearBtn.classList.remove('hidden');
    } else {
        searchMatchCount.textContent = '';
        searchClearBtn.classList.add('hidden');
    }
}

/* ─── GENERATE TAGGED TEXT (with overrides) ──────────────── */
function generateTaggedText() {
    if (!analysisResult) return '';

    const lines = currentText.split(/\r?\n/);
    const result = [];
    const speakersSorted = analysisResult.speakers.map(s => s.name).filter(n => n !== 'Nobody').sort((a, b) => b.length - a.length);
    const SPEECH_VERBS_RE = /\b(said|asked|replied|answered|shouted|whispered|called|cried|muttered|added|continued|declared|announced|stated|explained|noted|responded|interrupted|growled|repeated|confirmed|told|breathed|snapped|cut|nodded|turned)\b/i;

    let lastMentionedChar = null;
    let lastGenderChar = { m: null, f: null };
    let lastMentionedUsedCount = 0;

    for (const line of lines) {
        const hasQuote = /"/.test(line);

        if (!hasQuote) {
            scanProse(line, speakersSorted, lastGenderChar, (name) => { lastMentionedChar = name; lastMentionedUsedCount = 0; });
            result.push(line);
            continue;
        }

        const firstQ = line.indexOf('"');
        if (firstQ > 0) {
            scanProse(line.slice(0, firstQ), speakersSorted, lastGenderChar, (name) => { lastMentionedChar = name; lastMentionedUsedCount = 0; });
        }

        const speaker = resolve4Tier(line, speakersSorted, SPEECH_VERBS_RE, lastMentionedChar, lastMentionedUsedCount, lastGenderChar);
        const resolvedSpeaker = speaker || 'Nobody';

        if (resolvedSpeaker !== 'Nobody') {
            if (/\b(he|his|him)\b/i.test(line)) lastGenderChar.m = resolvedSpeaker;
            if (/\b(she|her|hers)\b/i.test(line)) lastGenderChar.f = resolvedSpeaker;
            if (resolvedSpeaker === lastMentionedChar) lastMentionedUsedCount++;
            else { lastMentionedChar = resolvedSpeaker; lastMentionedUsedCount = 0; }
        }

        // Apply content-key overrides
        const tagged = line.replace(/"([^"]+)"/g, (_, content) => {
            const slice = content.trim().slice(0, 60);
            const override = tagOverrides.get(slice);
            const effectiveSpeaker = override ?? resolvedSpeaker;
            if (!activeSpeakers.has(effectiveSpeaker)) return `"${content}"`;
            return `[${effectiveSpeaker}]${content}[/${effectiveSpeaker}]`;
        });
        result.push(tagged);
    }
    return result.join('\n');
}

function resolve4Tier(line, names, SPEECH_VERBS_RE, lastMentionedChar, lastMentionedUsedCount, lastGenderChar) {
    for (const name of names) {
        const esc = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        const pre = new RegExp(`\\b${esc}\\b.{0,80}${SPEECH_VERBS_RE.source}`, 'i');
        const post = new RegExp(`${SPEECH_VERBS_RE.source}.{0,40}\\b${esc}\\b`, 'i');
        if (pre.test(line) || post.test(line)) return name;
    }
    if (/\b(he|him)\b.{0,20}(said|asked|replied|answered|shouted|whispered|called|muttered|added|continued|responded|snapped)\b/i.test(line) ||
        /\b(said|asked|replied|answered|shouted|whispered|called|muttered|added|continued|responded|snapped)\b.{0,20}\b(he|him)\b/i.test(line)) {
        if (lastGenderChar.m) return lastGenderChar.m;
    }
    if (/\b(she|her)\b.{0,20}(said|asked|replied|answered|shouted|whispered|called|muttered|added|continued|responded|snapped)\b/i.test(line) ||
        /\b(said|asked|replied|answered|shouted|whispered|called|muttered|added|continued|responded|snapped)\b.{0,20}\b(she|her)\b/i.test(line)) {
        if (lastGenderChar.f) return lastGenderChar.f;
    }
    if (lastMentionedChar && lastMentionedUsedCount === 0) return lastMentionedChar;
    return null; // Nobody will be applied by caller
}

function scanProse(text, names, lastGenderChar, onNameFound) {
    for (const name of names) {
        const esc = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        if (new RegExp(`\\b${esc}\\b`).test(text)) {
            if (/\b(he|his|him)\b/i.test(text)) lastGenderChar.m = name;
            if (/\b(she|her|hers)\b/i.test(text)) lastGenderChar.f = name;
            onNameFound(name);
            break;
        }
    }
}

/* ─── HIGHLIGHTED HTML (interactive tags) ────────────────── */
function buildHighlightedHTML(text) {
    if (!analysisResult) return escHtml(text);

    const speakers = analysisResult.speakers.filter(s => activeSpeakers.has(s.name))
        .sort((a, b) => b.name.length - a.name.length);

    let html = escHtml(text);

    for (const s of speakers) {
        const pal = getPal(s);
        const escapedName = escHtml(s.name).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        const re = new RegExp(`\\[${escapedName}\\](.*?)\\[\\/${escapedName}\\]`, 'gs');
        html = html.replace(re, (_, content) => {
            // Decode the content back to get the raw text for the content key
            const rawContent = content.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"');
            const contentKey = escHtml(rawContent.trim().slice(0, 60));
            return `<button class="tag-btn" data-speaker="${escHtml(s.name)}" data-content-key="${contentKey}" style="border-left-color:${pal.color}; background:rgba(237,229,218,0.07);" title="Click to reassign speaker"><span class="tag-label" style="color:${pal.color};">[${escHtml(s.name)}]</span>${content}<span class="tag-label" style="color:${pal.color};">[/${escHtml(s.name)}]</span></button>`;
        });
    }
    return html;
}

/* ─── TAG CLICK → REASSIGNMENT POPUP ────────────────────── */
outputPreview.addEventListener('click', e => {
    const btn = e.target.closest('.tag-btn');
    if (!btn || !analysisResult) return;
    e.stopPropagation();
    const contentKey = btn.dataset.contentKey;
    const currentSpeaker = btn.dataset.speaker;
    showTagPopup(e.clientX, e.clientY, contentKey, currentSpeaker);
});

function showTagPopup(x, y, contentKey, currentSpeaker) {
    tagPopup.innerHTML = '';
    const title = document.createElement('div');
    title.className = 'popup-title';
    title.textContent = 'Reassign to:';
    tagPopup.appendChild(title);

    const grid = document.createElement('div');
    grid.className = 'popup-speaker-grid';

    for (const s of analysisResult.speakers) {
        const pal = getPal(s);
        const btn = document.createElement('button');
        btn.className = 'popup-speaker-btn' + (s.name === currentSpeaker ? ' current' : '');
        btn.style.setProperty('--chip-color', pal.color);
        btn.style.setProperty('--chip-bg', pal.bg);
        btn.innerHTML = `<span style="background:${pal.color}; width:8px; height:8px; border-radius:1px; display:inline-block; flex-shrink:0;"></span> ${escHtml(s.name)}`;
        btn.addEventListener('click', () => {
            tagOverrides.set(contentKey, s.name);
            closePopups();
            renderOutput();
        });
        grid.appendChild(btn);
    }
    tagPopup.appendChild(grid);

    // Custom name input
    const customRow = document.createElement('div');
    customRow.className = 'popup-custom-row';
    customRow.innerHTML = `<input class="popup-custom-input" placeholder="Custom name…" /><button class="popup-custom-ok btn btn--primary btn--sm">OK</button>`;
    tagPopup.appendChild(customRow);
    const inp = customRow.querySelector('.popup-custom-input');
    const ok = customRow.querySelector('.popup-custom-ok');
    ok.addEventListener('click', () => {
        const name = inp.value.trim();
        if (!name) return;
        tagOverrides.set(contentKey, name);
        closePopups();
        renderOutput();
    });
    inp.addEventListener('keydown', e => { if (e.key === 'Enter') ok.click(); });

    // Position popup
    positionPopup(tagPopup, x, y);
    tagPopup.classList.remove('hidden');
    popupBackdrop.classList.remove('hidden');
    inp.focus();
}

/* ─── CHIP RIGHT-CLICK → COLOR PICKER ───────────────────── */
function showChipColorMenu(e, speaker) {
    chipColorMenu.innerHTML = '';
    const title = document.createElement('div');
    title.className = 'popup-title';
    title.textContent = `Color: ${speaker.name}`;
    chipColorMenu.appendChild(title);

    const swatches = document.createElement('div');
    swatches.className = 'color-swatches';
    for (const pal of PALETTE) {
        const sw = document.createElement('button');
        sw.className = 'color-swatch';
        sw.style.background = pal.color;
        sw.title = pal.color;
        sw.addEventListener('click', () => {
            speakerColorOverrides.set(speaker.name, pal);
            closePopups();
            renderSpeakers();
            renderOutput();
        });
        swatches.appendChild(sw);
    }
    chipColorMenu.appendChild(swatches);

    // Custom color input
    const customRow = document.createElement('div');
    customRow.className = 'popup-custom-row';
    const currentPal = getPal(speaker);
    customRow.innerHTML = `<label class="popup-custom-label">Custom</label><input type="color" class="chip-color-input" value="${currentPal.color}" />`;
    chipColorMenu.appendChild(customRow);
    customRow.querySelector('.chip-color-input').addEventListener('input', ev => {
        const col = ev.target.value;
        const bg = col + '30'; // rough alpha
        speakerColorOverrides.set(speaker.name, { color: col, bg });
        renderSpeakers();
        renderOutput();
    });

    // Reset option
    const resetBtn2 = document.createElement('button');
    resetBtn2.className = 'btn btn--ghost btn--sm';
    resetBtn2.style.marginTop = '8px';
    resetBtn2.style.width = '100%';
    resetBtn2.textContent = 'Reset to default';
    resetBtn2.addEventListener('click', () => {
        speakerColorOverrides.delete(speaker.name);
        closePopups();
        renderSpeakers();
        renderOutput();
    });
    chipColorMenu.appendChild(resetBtn2);

    positionPopup(chipColorMenu, e.clientX, e.clientY);
    chipColorMenu.classList.remove('hidden');
    popupBackdrop.classList.remove('hidden');
}

function positionPopup(el, x, y) {
    el.style.left = '0';
    el.style.top = '0';
    el.classList.remove('hidden');
    const rect = el.getBoundingClientRect();
    const vw = window.innerWidth, vh = window.innerHeight;
    el.style.left = Math.min(x, vw - rect.width - 8) + 'px';
    el.style.top = Math.min(y + 8, vh - rect.height - 8) + 'px';
    el.classList.add('hidden');
}

function closePopups() {
    tagPopup.classList.add('hidden');
    chipColorMenu.classList.add('hidden');
    popupBackdrop.classList.add('hidden');
}

popupBackdrop.addEventListener('click', closePopups);
document.addEventListener('keydown', e => { if (e.key === 'Escape') closePopups(); });

/* ─── SETTINGS PANEL ─────────────────────────────────────── */
settingsBtn.addEventListener('click', () => {
    settingsPanel.classList.toggle('open');
    syncColorPickers();
});
settingsClose.addEventListener('click', () => settingsPanel.classList.remove('open'));

document.querySelectorAll('.theme-btn').forEach(btn => {
    btn.addEventListener('click', () => applyTheme(btn.dataset.theme));
});

document.querySelectorAll('.color-row input[type=color]').forEach(inp => {
    inp.addEventListener('input', () => {
        document.documentElement.style.setProperty(inp.dataset.var, inp.value);
        localStorage.setItem('sw_custom_' + inp.dataset.var, inp.value);
    });
});

// Load saved theme + custom colors on startup
(function loadSaved() {
    const saved = localStorage.getItem('sw_theme');
    if (saved && THEMES[saved]) applyTheme(saved);
    else applyTheme('dark-walnut');
    // Custom color overrides
    for (const key of Object.keys(localStorage)) {
        if (key.startsWith('sw_custom_')) {
            const varName = key.slice('sw_custom_'.length);
            const val = localStorage.getItem(key);
            document.documentElement.style.setProperty(varName, val);
        }
    }
})();

/* ─── OTHER CONTROLS ─────────────────────────────────────── */
highlightToggle.addEventListener('change', () => { if (analysisResult) renderOutput(); });

dialogueOnlyBtn.addEventListener('click', () => {
    dialogueOnly = !dialogueOnly;
    dialogueOnlyBtn.classList.toggle('active', dialogueOnly);
    if (analysisResult) renderOutput();
});

let searchDebounce;
searchInput.addEventListener('input', () => {
    clearTimeout(searchDebounce);
    searchDebounce = setTimeout(() => { searchTerm = searchInput.value; searchHitIndex = 0; if (analysisResult) renderOutput(); }, 120);
});
searchInput.addEventListener('keydown', e => {
    if (e.key !== 'Enter') return;
    e.preventDefault();
    const hits = outputPreview.querySelectorAll('.search-hit');
    if (!hits.length) return;
    searchHitIndex = e.shiftKey ? (searchHitIndex - 1 + hits.length) % hits.length : (searchHitIndex + 1) % hits.length;
    hits.forEach((h, i) => h.classList.toggle('current', i === searchHitIndex));
    hits[searchHitIndex]?.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
    searchMatchCount.textContent = `${searchHitIndex + 1} / ${hits.length}`;
});
searchClearBtn.addEventListener('click', () => {
    searchInput.value = ''; searchTerm = ''; searchHitIndex = 0;
    searchClearBtn.classList.add('hidden'); searchMatchCount.textContent = '';
    if (analysisResult) renderOutput();
    searchInput.focus();
});

copyBtn.addEventListener('click', async () => {
    try { await navigator.clipboard.writeText(generateTaggedText()); showToast('✓ Copied!', 'success'); }
    catch { showToast('Copy failed — select and copy manually.', 'error'); }
});

downloadBtn.addEventListener('click', () => {
    const blob = new Blob([generateTaggedText()], { type: 'text/plain; charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a'); a.href = url; a.download = 'tagged_book.txt'; a.click();
    URL.revokeObjectURL(url);
    showToast('✓ Downloaded tagged_book.txt', 'success');
});

/* ─── SEND TO TTS-STORY ──────────────────────────────────── */
const sendTtsBtn = document.getElementById('send-tts-btn');
if (sendTtsBtn) {
    sendTtsBtn.addEventListener('click', sendToTTSStory);
}

async function sendToTTSStory() {
    if (!analysisResult) { showToast('⚠ Analyze your book first.', 'error'); return; }

    const tagged = generateTaggedText();
    const hostInput = document.getElementById('tts-host-input');
    const host = (hostInput ? hostInput.value.trim() : 'http://localhost:5000').replace(/\/$/, '');

    // Reliable clipboard copy that works on file:// pages
    let copied = false;
    try {
        // Modern API first
        await navigator.clipboard.writeText(tagged);
        copied = true;
    } catch {
        // Fallback: textarea + execCommand (works on file://)
        try {
            const ta = document.createElement('textarea');
            ta.value = tagged;
            ta.style.cssText = 'position:fixed;top:0;left:0;opacity:0;pointer-events:none;';
            document.body.appendChild(ta);
            ta.focus(); ta.select();
            copied = document.execCommand('copy');
            document.body.removeChild(ta);
        } catch { /* truly failed */ }
    }

    // Open TTS-Story — user may need to start it first, that's fine
    window.open(host, '_blank');

    showToast(
        copied
            ? `✓ Tagged text copied — paste it in TTS-Story → Generate tab`
            : `✓ Opened TTS-Story — manually copy the output and paste in Generate`,
        'success'
    );
}

resetBtn.addEventListener('click', () => {
    pasteInput.value = ''; currentText = ''; analysisResult = null;
    activeSpeakers.clear(); resultsSection.classList.add('hidden'); fileInput.value = '';
    dialogueOnly = false; dialogueOnlyBtn.classList.remove('active');
    searchTerm = ''; searchInput.value = ''; searchMatchCount.textContent = ''; searchClearBtn.classList.add('hidden');
    tagOverrides.clear(); speakerColorOverrides.clear();
});

/* ─── HELPERS ────────────────────────────────────────────── */
function escHtml(str) {
    return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
function showToast(msg, type = 'info') {
    const icons = { success: '✓', error: '✗', info: 'ℹ' };
    const t = document.createElement('div');
    t.className = `toast toast--${type}`;
    t.innerHTML = `<span class="toast-icon">${icons[type]}</span>${escHtml(msg)}`;
    toastContainer.appendChild(t);
    setTimeout(() => t.remove(), 4000);
}

/* ─── KEYBOARD SHORTCUTS ─────────────────────────────────── */
document.addEventListener('keydown', e => {
    if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') { e.preventDefault(); analyzeBtn.click(); }
    if ((e.ctrlKey || e.metaKey) && e.key === ',') { e.preventDefault(); settingsPanel.classList.toggle('open'); }
});
