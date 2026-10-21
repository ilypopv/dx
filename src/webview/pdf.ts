// The PDF viewer: pdf.js and its viewer components, framed by the dx shell.
// The pdf.js worker is baked into this bundle as a string and started as a same-origin blob worker,
// the file bytes arrive from the extension host, so the viewer never loads the document over the network.
import { pdfjsLib } from './pdfjs-global';
import {
    EventBus,
    PDFLinkService,
    PDFFindController,
    PDFViewer,
    FindState,
    LinkTarget
} from 'pdfjs-dist/web/pdf_viewer.mjs';
import workerSource from 'pdfjs-dist/build/pdf.worker.min.mjs';
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
    interceptLinks
} from './shell';

const { AnnotationEditorType, AnnotationMode, PasswordResponses, GlobalWorkerOptions, getDocument } = pdfjsLib;

const options = readOptions();
const shell = new Shell(document.getElementById('app')!);
const viewState = new ViewState(options, '1');

let worker: Worker | undefined;
let workerUrl: string | undefined;
let pdfDocument: any;
let loadingTask: any;
let editable = false;
let changeReported = false;
let restore: { page: number; scale: string | number; rotation: number } | null = viewState.load();

function ensureWorker() {
    if (!worker) {
        workerUrl = URL.createObjectURL(new Blob([workerSource], { type: 'text/javascript' }));
        worker = new Worker(workerUrl, { type: 'module' });
        GlobalWorkerOptions.workerPort = worker;
    }
}

// ---------- layout ----------

const container = h('div', 'dx-pdf');
container.tabIndex = 0;
const viewerEl = h('div', 'pdfViewer');
container.appendChild(viewerEl);
shell.main.appendChild(container);

const eventBus = new EventBus();
const linkService = new PDFLinkService({ eventBus, externalLinkTarget: LinkTarget.BLANK });
const findController = new PDFFindController({ eventBus, linkService });
const viewer = new PDFViewer({
    container,
    viewer: viewerEl,
    eventBus,
    linkService,
    findController,
    textLayerMode: 1,
    annotationMode: AnnotationMode.ENABLE_FORMS,
    annotationEditorMode: AnnotationEditorType.NONE,
    annotationEditorHighlightColors: 'yellow=#FFFF98,green=#53FFBC,blue=#80EBFF,pink=#FFCBE6,red=#FF4F5F',
    enableHighlightFloatingButton: true,
    removePageBorders: false
});
linkService.setViewer(viewer);
interceptLinks(container);

// ---------- toolbar ----------

const prev = button('', 'Previous page', icon(ICONS.up));
const next = button('', 'Next page', icon(ICONS.down));
const pageIn = h('input', 'dx-input dx-page-in');
pageIn.setAttribute('aria-label', 'Page number');
pageIn.spellcheck = false;
const pageOf = h('span', 'dx-of', '');
const fitWidth = button('', 'Fit width', icon(ICONS.fitWidth));
const fitPage = button('', 'Fit page', icon(ICONS.fitPage));
const rotate = button('', 'Rotate clockwise', icon(ICONS.rotate));
const sep = () => h('span', 'dx-sep');
const tools = h('div', 'dx-group dx-tools');
const toolButtons = new Map<number, HTMLButtonElement>([
    [AnnotationEditorType.HIGHLIGHT, button('', 'Highlight text', icon(ICONS.highlight))],
    [AnnotationEditorType.FREETEXT, button('', 'Add text', icon(ICONS.text))],
    [AnnotationEditorType.INK, button('', 'Draw', icon(ICONS.ink))]
]);
tools.append(sep(), ...toolButtons.values());
tools.style.display = 'none';
shell.left.append(prev, next, pageIn, pageOf, sep(), fitWidth, fitPage, rotate, tools);

prev.addEventListener('click', () => viewer.previousPage());
next.addEventListener('click', () => viewer.nextPage());
pageIn.addEventListener('keydown', e => {
    if (e.key === 'Enter') {
        e.preventDefault();
        const n = parseInt(pageIn.value, 10);
        if (pdfDocument && n >= 1 && n <= pdfDocument.numPages) {
            viewer.currentPageNumber = n;
            container.focus();
        } else {
            pageIn.classList.add('bad');
            setTimeout(() => pageIn.classList.remove('bad'), 400);
            pageIn.value = String(viewer.currentPageNumber);
        }
    } else if (e.key === 'Escape') {
        pageIn.value = String(viewer.currentPageNumber);
        container.focus();
    }
});
pageIn.addEventListener('focus', () => pageIn.select());
fitWidth.addEventListener('click', () => (viewer.currentScaleValue = 'page-width'));
fitPage.addEventListener('click', () => (viewer.currentScaleValue = 'page-fit'));
rotate.addEventListener('click', () => {
    viewer.pagesRotation = (viewer.pagesRotation + 90) % 360;
    saveState();
});
for (const [mode, b] of toolButtons) {
    b.addEventListener('click', () => {
        if (!pdfDocument) {
            return;
        }
        const current = viewer.annotationEditorMode;
        viewer.annotationEditorMode = { mode: current === mode ? AnnotationEditorType.NONE : mode };
    });
}
document.addEventListener('keydown', e => {
    if (
        e.key === 'Escape' &&
        !shell.findOpen &&
        pdfDocument &&
        viewer.annotationEditorMode !== AnnotationEditorType.NONE
    ) {
        viewer.annotationEditorMode = { mode: AnnotationEditorType.NONE };
    }
});

// ---------- viewer events ----------

function saveState() {
    if (!pdfDocument) {
        return;
    }
    viewState.save({ page: viewer.currentPageNumber, scale: viewer.currentScaleValue, rotation: viewer.pagesRotation });
}

function updatePage() {
    const n = viewer.currentPageNumber;
    const total = pdfDocument ? pdfDocument.numPages : 0;
    if (document.activeElement !== pageIn) {
        pageIn.value = String(n);
    }
    pageOf.textContent = '/ ' + total;
    prev.disabled = n <= 1;
    next.disabled = n >= total;
    shell.setInfo(total ? `Page ${n} of ${total}` : '');
}

// What the viewer shows, reported to the host (the integration tests wait for it).
const reportLoaded = (text: boolean) =>
    pdfDocument && post({ loaded: { kind: 'pdf', pages: pdfDocument.numPages, rendered: true, text } });
eventBus.on('pagerendered', (e: any) => e.pageNumber === viewer.currentPageNumber && reportLoaded(false));
eventBus.on('textlayerrendered', (e: any) => {
    if (e.pageNumber === viewer.currentPageNumber) {
        reportLoaded(!!container.querySelector('.textLayer span'));
    }
});
eventBus.on('pagesinit', () => {
    const r = restore;
    restore = null;
    viewer.currentScaleValue = r && r.scale ? String(r.scale) : 'auto';
    if (r && r.rotation) {
        viewer.pagesRotation = r.rotation;
    }
    if (r && r.page > 1 && r.page <= pdfDocument.numPages) {
        viewer.currentPageNumber = r.page;
    }
    updatePage();
    shell.showZoom(viewer.currentScale);
});
eventBus.on('pagechanging', () => {
    updatePage();
    saveState();
});
eventBus.on('scalechanging', (e: any) => {
    shell.showZoom(e.scale);
    saveState();
});
eventBus.on('updatefindmatchescount', (e: any) => {
    shell.setFindCount(e.matchesCount.current, e.matchesCount.total);
});
eventBus.on('updatefindcontrolstate', (e: any) => {
    if (e.state === FindState.PENDING) {
        return;
    }
    if (e.state === FindState.NOT_FOUND) {
        shell.setFindCount(0, 0);
    } else if (e.matchesCount) {
        shell.setFindCount(e.matchesCount.current, e.matchesCount.total);
    }
});
eventBus.on('annotationeditormodechanged', (e: any) => {
    for (const [mode, b] of toolButtons) {
        b.classList.toggle('on', e.mode === mode);
        b.setAttribute('aria-pressed', String(e.mode === mode));
    }
});

const find = (query: any, again: boolean, previous: boolean) =>
    eventBus.dispatch('find', {
        source: shell,
        type: again ? 'again' : '',
        query: query.query,
        caseSensitive: query.caseSensitive,
        entireWord: query.entireWord,
        highlightAll: true,
        findPrevious: previous,
        matchDiacritics: false
    });
shell.findHandler = {
    find: q => find(q, false, false),
    next: (q, previous) => find(q, true, previous),
    clear: () => eventBus.dispatch('findbarclose', { source: shell })
};
shell.enableZoom({
    get: () => viewer.currentScale || 1,
    set: zoom => (viewer.currentScale = zoom),
    reset: () => (viewer.currentScaleValue = '1')
});

// ---------- loading ----------

function askPassword(incorrect: boolean): Promise<string> {
    return new Promise((resolve, reject) => {
        const box = h('form', 'dx-password');
        box.appendChild(
            h('div', 'dx-error-title', incorrect ? 'Incorrect password' : 'This PDF is password protected')
        );
        const input = h('input', 'dx-input');
        input.type = 'password';
        input.placeholder = 'Password';
        const ok = h('button', 'dx-btn dx-primary', 'Open');
        ok.type = 'submit';
        const row = h('div', 'dx-row');
        row.append(input, ok);
        box.appendChild(row);
        box.addEventListener('submit', e => {
            e.preventDefault();
            box.remove();
            resolve(input.value);
        });
        input.addEventListener('keydown', e => {
            if (e.key === 'Escape') {
                box.remove();
                reject(new Error('No password was given.'));
            }
        });
        shell.main.appendChild(box);
        shell.setMessage('Password required');
        input.focus();
    });
}

async function closeDocument() {
    if (loadingTask) {
        const task = loadingTask;
        loadingTask = undefined;
        await task.destroy().catch(() => undefined);
    }
    if (pdfDocument) {
        const doc = pdfDocument;
        pdfDocument = undefined;
        viewer.setDocument(null);
        linkService.setDocument(null);
        await doc.destroy().catch(() => undefined);
    }
}

async function load(bytes: Uint8Array) {
    // A reload (revert, change on disk) keeps the position.
    if (pdfDocument && !restore) {
        restore = { page: viewer.currentPageNumber, scale: viewer.currentScaleValue, rotation: viewer.pagesRotation };
    }
    await closeDocument();
    shell.setMessage('Opening…');
    if (!container.isConnected) {
        shell.main.textContent = '';
        shell.main.appendChild(container);
    }
    try {
        ensureWorker();
        const task = getDocument({
            data: bytes,
            cMapUrl: options.media + '/pdfjs/cmaps/',
            cMapPacked: true,
            standardFontDataUrl: options.media + '/pdfjs/standard_fonts/',
            useWorkerFetch: false,
            isEvalSupported: false,
            enableXfa: true
        });
        loadingTask = task;
        task.onPassword = (update: (password: string) => void, reason: number) => {
            askPassword(reason === PasswordResponses.INCORRECT_PASSWORD).then(update, error => {
                task.destroy();
                shell.showError('The PDF was not opened.', errorText(error));
            });
        };
        const doc = await task.promise;
        if (loadingTask !== task) {
            doc.destroy();
            return;
        }
        loadingTask = undefined;
        pdfDocument = doc;
        changeReported = false;
        doc.annotationStorage.onSetModified = () => {
            if (editable && !changeReported) {
                changeReported = true;
                post({ changed: true, reason: 'Annotation' });
            }
        };
        viewer.setDocument(doc);
        linkService.setDocument(doc, null);
        tools.style.display = editable ? '' : 'none';
        shell.setMessage(editable ? 'Ready' : 'Read-only');
        updatePage();
        container.focus();
    } catch (error) {
        if (error && (error as any).name === 'PasswordException') {
            return;
        }
        const name = error && (error as any).name;
        const detail =
            name === 'InvalidPDFException' ? 'The file is damaged or is not a PDF document.' : errorText(error);
        shell.showError('The PDF could not be opened.', detail);
        post({ error: `Could not open '${options.name}': ${detail}` });
    }
}

async function getData(requestId: number) {
    try {
        if (!pdfDocument) {
            throw new Error('No document is open.');
        }
        shell.setMessage('Saving…');
        const data: Uint8Array = await pdfDocument.saveDocument();
        changeReported = false;
        post({ response: true, requestId, body: data });
        shell.setMessage('Ready');
    } catch (error) {
        shell.setMessage('Save failed');
        post({ response: true, requestId, error: errorText(error) });
    }
}

window.addEventListener('message', (event: MessageEvent) => {
    const message = event.data;
    if (!message || typeof message !== 'object') {
        return;
    }
    if (message.type === 'getData') {
        getData(message.requestId);
    } else if (message.type === 'test' && options.test && pdfDocument) {
        // Integration tests: a free text annotation, as if typed with the text tool.
        pdfDocument.annotationStorage.setValue('pdfjs_internal_editor_test', {
            annotationType: AnnotationEditorType.FREETEXT,
            color: [0, 0, 0],
            fontSize: 12,
            value: 'dx test note',
            pageIndex: 0,
            rect: [50, 50, 200, 80],
            rotation: 0,
            id: null
        });
    } else if (message.refresh || message.content) {
        const bytes = toBytes(message.content);
        if (!bytes) {
            shell.showError('The document did not reach the viewer.');
            return;
        }
        editable = !!message.editable;
        // pdf.js transfers the buffer to its worker: keep the copy the host sent intact.
        load(bytes.slice());
    }
});

window.addEventListener('unload', () => {
    closeDocument();
    if (worker) {
        worker.terminate();
    }
    if (workerUrl) {
        URL.revokeObjectURL(workerUrl);
    }
});

post({ refresh: true });
