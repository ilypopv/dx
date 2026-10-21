// The .docx and .doc viewer. The format is decided by the bytes, not by the file name:
// a .doc that is really a .docx (or the other way round) still opens, read-only when dx cannot write it back.
// Editing is deliberately small: text inside a paragraph can be typed, deleted and pasted (as plain text);
// paragraphs cannot be split, joined or reformatted, so the structure of the document always survives a save.
import { readDocx, writeDocx } from '../core/docx';
import { readDoc } from '../core/doc';
import { Renderer } from './render';
import {
    Shell,
    ViewState,
    readOptions,
    post,
    h,
    icon,
    button,
    ICONS,
    errorText,
    toBytes,
    interceptLinks,
    FindQuery
} from './shell';

const options = readOptions();
const shell = new Shell(document.getElementById('app')!);
const viewState = new ViewState(options, '1');
const PX_PER_PT = 4 / 3;

const scroller = h('div', 'dx-scroll');
scroller.tabIndex = 0;
const pagesEl = h('div', 'dx-pages');
scroller.appendChild(pagesEl);
shell.main.appendChild(scroller);

let docx: any;
let pages: HTMLElement[] = [];
let images: string[] = [];
let words = 0;
let pageWidthPt = 595.3;
let zoom = 1;
let fitWidth = false;
let canEdit = false;
let editing = false;
let changeReported = false;
const dirty = new Set<number>();
let saved = viewState.load();

// ---------- toolbar ----------

const editBtn = button('dx-labeled', 'Edit the text of the document', icon(ICONS.pen));
editBtn.appendChild(h('span', '', 'Edit'));
editBtn.style.display = 'none';
const fitBtn = button('', 'Fit width', icon(ICONS.fitWidth));
const readOnlyTag = h('span', 'dx-tag', 'Read-only');
readOnlyTag.style.display = 'none';
shell.left.append(editBtn, fitBtn, readOnlyTag);

editBtn.addEventListener('click', () => setEditing(!editing));
fitBtn.addEventListener('click', () => {
    fitWidth = true;
    applyZoom(fitZoom());
});

function setEditing(on: boolean) {
    editing = on && canEdit;
    pagesEl.contentEditable = editing ? 'true' : 'false';
    pagesEl.spellcheck = false;
    pagesEl.classList.toggle('dx-editing', editing);
    editBtn.classList.toggle('on', editing);
    editBtn.setAttribute('aria-pressed', String(editing));
    shell.setMessage(editing ? 'Editing: text changes only, Ctrl+S saves' : canEdit ? 'Ready' : 'Read-only');
}

// ---------- zoom and position ----------

function fitZoom(): number {
    const available = scroller.clientWidth - 48;
    return available > 0 ? Math.min(4, Math.max(0.25, available / (pageWidthPt * PX_PER_PT))) : 1;
}

function applyZoom(z: number) {
    const anchor = scroller.scrollTop / Math.max(1, scroller.scrollHeight);
    zoom = z;
    (pagesEl.style as any).zoom = String(z);
    scroller.scrollTop = anchor * scroller.scrollHeight;
    shell.showZoom(z);
    updatePage();
    saveState();
}

shell.enableZoom({
    get: () => zoom,
    set: z => {
        fitWidth = false;
        applyZoom(z);
    },
    reset: () => {
        fitWidth = false;
        applyZoom(1);
    }
});

new ResizeObserver(() => {
    if (fitWidth) {
        applyZoom(fitZoom());
    }
}).observe(scroller);

function saveState() {
    viewState.save({
        zoom: fitWidth ? 'fit' : zoom,
        scroll: scroller.scrollTop / Math.max(1, scroller.scrollHeight)
    });
}

let pageFrame = 0;
function updatePage() {
    if (pageFrame) {
        return;
    }
    pageFrame = requestAnimationFrame(() => {
        pageFrame = 0;
        const view = scroller.getBoundingClientRect();
        const middle = view.top + view.height / 3;
        let current = 1;
        for (let i = 0; i < pages.length; i++) {
            if (pages[i].getBoundingClientRect().top <= middle) {
                current = i + 1;
            }
        }
        const count = words.toLocaleString() + (words === 1 ? ' word' : ' words');
        shell.setInfo(pages.length ? `Page ${current} of ${pages.length}  ·  ${count}` : '');
    });
}

scroller.addEventListener(
    'scroll',
    () => {
        updatePage();
        saveState();
    },
    { passive: true }
);

// ---------- find ----------

const highlights: any = (globalThis as any).CSS && (CSS as any).highlights;
const HighlightCtor: any = (globalThis as any).Highlight;
let matches: Range[] = [];
let matchIndex = -1;
let matchKey = '';

function escapeRegExp(s: string) {
    return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function search(q: FindQuery) {
    matches = [];
    const flags = 'gu' + (q.caseSensitive ? '' : 'i');
    const body = escapeRegExp(q.query);
    const re = new RegExp(q.entireWord ? `(?<![\\p{L}\\p{N}_])${body}(?![\\p{L}\\p{N}_])` : body, flags);
    for (const p of Array.from(pagesEl.querySelectorAll('.dx-p'))) {
        const nodes: Text[] = [];
        const starts: number[] = [];
        let text = '';
        const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT, {
            acceptNode: n =>
                n.parentElement && n.parentElement.closest('.dx-marker')
                    ? NodeFilter.FILTER_REJECT
                    : NodeFilter.FILTER_ACCEPT
        });
        for (let n = walker.nextNode(); n; n = walker.nextNode()) {
            // Nested paragraphs (text boxes) are searched on their own.
            if (n.parentElement!.closest('.dx-p') !== p) {
                continue;
            }
            nodes.push(n as Text);
            starts.push(text.length);
            text += (n as Text).data;
        }
        if (!text) {
            continue;
        }
        const locate = (offset: number, end: boolean): [Text, number] => {
            let i = nodes.length - 1;
            while (i > 0 && (starts[i] > offset || (end && starts[i] === offset))) {
                i--;
            }
            return [nodes[i], offset - starts[i]];
        };
        re.lastIndex = 0;
        let m: RegExpExecArray | null;
        while ((m = re.exec(text)) && matches.length < 10000) {
            if (!m[0].length) {
                re.lastIndex++;
                continue;
            }
            const range = document.createRange();
            const [sn, so] = locate(m.index, false);
            const [en, eo] = locate(m.index + m[0].length, true);
            range.setStart(sn, so);
            range.setEnd(en, eo);
            matches.push(range);
        }
    }
    if (highlights && HighlightCtor) {
        highlights.set('dx-find', new HighlightCtor(...matches));
    }
}

function showMatch(index: number) {
    matchIndex = index;
    const range = matches[index];
    if (highlights && HighlightCtor) {
        if (range) {
            highlights.set('dx-find-current', new HighlightCtor(range));
        } else {
            highlights.delete('dx-find-current');
        }
    }
    shell.setFindCount(range ? index + 1 : 0, matches.length);
    if (range) {
        const r = range.getBoundingClientRect();
        const view = scroller.getBoundingClientRect();
        if (r.top < view.top + 20 || r.bottom > view.bottom - 20) {
            scroller.scrollTop += r.top - view.top - view.height / 3;
        }
        if (r.left < view.left || r.right > view.right) {
            scroller.scrollLeft += r.left - view.left - view.width / 3;
        }
    }
}

function runFind(q: FindQuery) {
    search(q);
    matchKey = JSON.stringify(q);
    // Start from the first match below the top of the view.
    const top = scroller.getBoundingClientRect().top;
    let first = matches.findIndex(r => r.getBoundingClientRect().bottom >= top);
    showMatch(matches.length ? Math.max(0, first) : -1);
}

shell.findHandler = {
    find: runFind,
    next: (q, previous) => {
        if (JSON.stringify(q) !== matchKey) {
            runFind(q);
            return;
        }
        if (!matches.length) {
            shell.setFindCount(0, 0);
            return;
        }
        showMatch((matchIndex + (previous ? matches.length - 1 : 1)) % matches.length);
    },
    clear: () => {
        matches = [];
        matchIndex = -1;
        matchKey = '';
        if (highlights) {
            highlights.delete('dx-find');
            highlights.delete('dx-find-current');
        }
    }
};

// ---------- editing ----------

const paragraphOf = (node: Node | null): HTMLElement | null => {
    const el = node && (node.nodeType === Node.ELEMENT_NODE ? (node as Element) : node.parentElement);
    return el ? (el.closest('.dx-p[data-p]') as HTMLElement | null) : null;
};

let inputParagraph: HTMLElement | null = null;

pagesEl.addEventListener('beforeinput', (e: InputEvent) => {
    inputParagraph = null;
    if (!editing) {
        e.preventDefault();
        return;
    }
    const type = e.inputType;
    if (type === 'historyUndo' || type === 'historyRedo') {
        return;
    }
    if (
        type === 'insertParagraph' ||
        type === 'insertLineBreak' ||
        type.startsWith('format') ||
        type === 'insertFromDrop' ||
        type === 'deleteByDrag' ||
        type === 'insertLink' ||
        type === 'insertOrderedList' ||
        type === 'insertUnorderedList' ||
        type === 'insertHorizontalRule'
    ) {
        e.preventDefault();
        return;
    }
    // Every change must stay inside one paragraph: no joins across paragraph ends, no multi-paragraph deletes.
    const ranges: StaticRange[] = e.getTargetRanges();
    const selection = window.getSelection();
    const pairs: [Node, Node][] = ranges.length
        ? ranges.map(r => [r.startContainer, r.endContainer] as [Node, Node])
        : selection && selection.rangeCount
          ? [[selection.getRangeAt(0).startContainer, selection.getRangeAt(0).endContainer]]
          : [];
    for (const [start, end] of pairs) {
        const a = paragraphOf(start);
        if (!a || a !== paragraphOf(end)) {
            e.preventDefault();
            return;
        }
        inputParagraph = a;
    }
    if (!pairs.length) {
        e.preventDefault();
    }
});

pagesEl.addEventListener('input', () => {
    const selection = window.getSelection();
    for (const p of [inputParagraph, selection ? paragraphOf(selection.anchorNode) : null]) {
        if (p && p.dataset.p !== undefined) {
            dirty.add(+p.dataset.p);
        }
    }
    inputParagraph = null;
    if (!changeReported) {
        changeReported = true;
        post({ changed: true, reason: 'Edit text' });
    }
    if (matchKey) {
        // The text moved under the find results.
        const q = JSON.parse(matchKey);
        search(q);
        showMatch(Math.min(matchIndex, matches.length - 1));
    }
});

pagesEl.addEventListener('paste', e => {
    e.preventDefault();
    if (!editing || !e.clipboardData) {
        return;
    }
    const text = e.clipboardData.getData('text/plain').replace(/\s*[\r\n]+\s*/g, ' ');
    if (text) {
        document.execCommand('insertText', false, text);
    }
});
pagesEl.addEventListener('drop', e => e.preventDefault());

document.addEventListener('keydown', e => {
    if (e.key === 'Escape' && editing && !shell.findOpen) {
        setEditing(false);
    }
});

// The pieces of an edited paragraph in document order: rendered pieces by slot, everything else as loose text.
function segments(p: HTMLElement) {
    const out: { slot?: number; text: string }[] = [];
    const seen = new Set<number>();
    const walk = (node: Node) => {
        for (const c of Array.from(node.childNodes)) {
            if (c.nodeType === Node.TEXT_NODE) {
                out.push({ text: (c as Text).data });
            } else if (c.nodeType === Node.ELEMENT_NODE) {
                const el = c as HTMLElement;
                if (el.classList.contains('dx-marker')) {
                    continue;
                }
                const slot = el.dataset.s;
                if (slot !== undefined && !seen.has(+slot)) {
                    seen.add(+slot);
                    out.push({ slot: +slot, text: el.textContent || '' });
                } else if (el.tagName !== 'BR') {
                    walk(el);
                }
            }
        }
    };
    walk(p);
    return out;
}

function getData(requestId: number) {
    try {
        if (!docx) {
            throw new Error('This document cannot be saved by dx.');
        }
        const edits = [];
        for (const id of dirty) {
            const p = pagesEl.querySelector(`.dx-p[data-p="${id}"]`) as HTMLElement | null;
            if (p) {
                edits.push({ p: id, segments: segments(p) });
            }
        }
        const body = writeDocx(docx, edits);
        changeReported = false;
        post({ response: true, requestId, body });
    } catch (error) {
        post({ response: true, requestId, error: errorText(error) });
    }
}

interceptLinks(pagesEl, () => !editing);

// ---------- loading ----------

function detect(bytes: Uint8Array): 'docx' | 'doc' | 'rtf' | 'unknown' {
    if (bytes.length >= 4 && bytes[0] === 0x50 && bytes[1] === 0x4b) {
        return 'docx';
    }
    if (bytes.length >= 8 && bytes[0] === 0xd0 && bytes[1] === 0xcf && bytes[2] === 0x11 && bytes[3] === 0xe0) {
        return 'doc';
    }
    if (String.fromCharCode(...Array.from(bytes.subarray(0, 5))) === '{\\rtf') {
        return 'rtf';
    }
    return 'unknown';
}

function load(bytes: Uint8Array, editable: boolean) {
    const keep = pagesEl.childNodes.length
        ? { zoom: fitWidth ? 'fit' : zoom, scroll: scroller.scrollTop / Math.max(1, scroller.scrollHeight) }
        : null;
    for (const url of images) {
        URL.revokeObjectURL(url);
    }
    shell.findHandler!.clear();
    pagesEl.textContent = '';
    dirty.clear();
    changeReported = false;
    docx = undefined;
    let model;
    let files;
    const format = detect(bytes);
    try {
        if (format === 'docx') {
            docx = readDocx(bytes);
            model = docx.model;
            files = docx.files;
        } else if (format === 'doc') {
            model = readDoc(bytes);
        } else if (format === 'rtf') {
            throw new Error('This is an RTF document saved with a Word extension; dx does not read RTF.');
        } else if (bytes.length === 0) {
            throw new Error('The file is empty.');
        } else {
            throw new Error('The file is neither a Word 97-2003 (.doc) nor a Word 2007+ (.docx) document.');
        }
    } catch (error) {
        shell.showError('The document could not be opened.', errorText(error));
        post({ error: `Could not open '${options.name}': ${errorText(error)}` });
        return;
    }
    // Only the OOXML format can be written back, and only to a .docx file.
    canEdit = editable && format === 'docx' && options.kind === 'docx';
    if (!canEdit) {
        docx = undefined;
    }
    if (!scroller.isConnected) {
        shell.main.textContent = '';
        shell.main.appendChild(scroller);
    }
    pagesEl.classList.toggle('dx-plain', format === 'doc');
    const rendered = new Renderer(pagesEl, files, canEdit).render(model);
    pages = rendered.pages;
    images = rendered.images;
    words = rendered.words;
    pageWidthPt = (model.page && model.page.width) || 595.3;

    editBtn.style.display = canEdit ? '' : 'none';
    readOnlyTag.style.display = canEdit ? 'none' : '';
    readOnlyTag.title =
        format === 'doc' ? 'dx shows Word 97-2003 documents read-only' : 'This document cannot be saved by dx';
    setEditing(editing);

    const state = keep || saved;
    saved = null;
    fitWidth = !state || state.zoom === 'fit';
    applyZoom(fitWidth ? Math.min(1, fitZoom()) : +state.zoom || 1);
    if (fitWidth && fitZoom() >= 1) {
        // A wide panel shows the page at its real size; narrow ones follow the width.
        fitWidth = false;
    }
    if (state && state.scroll) {
        scroller.scrollTop = state.scroll * scroller.scrollHeight;
    }
    updatePage();
    scroller.focus();
    post({ loaded: { kind: format, pages: pages.length, words, editable: canEdit } });
}

// Integration tests: type into the first editable paragraph the way the browser would.
function testEdit() {
    setEditing(true);
    const p = pagesEl.querySelector('.dx-p[data-p]') as HTMLElement | null;
    const piece = p && (p.querySelector('[data-s]') as HTMLElement | null);
    if (!p || !piece) {
        return;
    }
    piece.textContent = 'Edited by dx';
    inputParagraph = p;
    pagesEl.dispatchEvent(new Event('input'));
}

window.addEventListener('message', (event: MessageEvent) => {
    const message = event.data;
    if (!message || typeof message !== 'object') {
        return;
    }
    if (message.type === 'getData') {
        getData(message.requestId);
    } else if (message.type === 'test' && options.test) {
        testEdit();
    } else if (message.refresh || message.content) {
        const bytes = toBytes(message.content);
        if (!bytes) {
            shell.showError('The document did not reach the viewer.');
            return;
        }
        load(bytes, !!message.editable);
    }
});

post({ refresh: true });
