// Renders the block model of doc.js / docx.js as pages of HTML, laid out by the layout engine (core/layout.js),
// which computes the pages and the lines the way Word does. Every line of a paragraph is drawn as the engine broke
// it: its height, where its text starts, how its spaces are stretched for justification. While the text is edited
// the lines flow freely again (see .dx-editing in the stylesheet).
// Paragraphs carry data-p (the paragraph id of an editable .docx), rendered pieces carry data-s (their slot),
// which is how an edit in the page finds its way back into the XML.
import { h } from './shell';
import { createMeasure, fontStack, lineRatio, DEFAULT_STACK } from './metrics';
import { createLayout, cellWidths, runLength } from '../core/layout';

export interface Page {
    width: number;
    height: number;
    top: number;
    right: number;
    bottom: number;
    left: number;
    // Distances of the header and footer from the edge of the page.
    header?: number;
    footer?: number;
}

// Headers and footers of one section, by kind: 'default', 'first' (title page) and 'even'.
export interface Section {
    header: { [kind: string]: any[] };
    footer: { [kind: string]: any[] };
    titlePg?: boolean;
    start?: number;
    page?: Page;
}

export interface Model {
    blocks: any[];
    notes?: any[];
    sections?: Section[];
    evenOdd?: boolean;
    page?: Page;
    font?: { family?: string; size?: number };
}

export { fontStack };

// A4 with 2 cm margins for the binary format, which keeps its page setup in the section properties dx does not read.
const DEFAULT_PAGE: Page = { width: 595.3, height: 841.9, top: 56.7, right: 56.7, bottom: 56.7, left: 56.7 };

export interface Rendered {
    pages: HTMLElement[];
    images: string[];
    words: number;
}

interface SheetParts {
    page: Page;
    section: number;
    number: number;
    kind: string;
    // The areas above and below the body: the header and footer live there.
    top: HTMLElement;
    body: HTMLElement;
    bottom: HTMLElement;
}

const pt = (v: number) => +v.toFixed(3) + 'pt';

export class Renderer {
    private pages: HTMLElement[] = [];
    private images: string[] = [];
    private words = 0;
    private model: Model | undefined;
    private layout: any;
    private readonly parts = new WeakMap<HTMLElement, SheetParts>();
    private readonly lineCache = new WeakMap<any, Map<number, any[]>>();
    private observer: IntersectionObserver | undefined;
    private readonly pending = new Map<Element, () => void>();
    private disposed = false;

    constructor(
        private readonly root: HTMLElement,
        private readonly files: { [name: string]: Uint8Array } | undefined,
        private readonly editable: boolean,
        // The scrolling ancestor: pictures far outside of it are not inflated and decoded yet.
        private readonly viewport?: Element
    ) {}

    // Stops loading pictures; the blob URLs stay valid until the caller revokes `images`.
    dispose() {
        this.disposed = true;
        if (this.observer) {
            this.observer.disconnect();
            this.observer = undefined;
        }
    }

    get pageCount(): number {
        return this.pages.length;
    }

    // Renders everything at once.
    render(model: Model): Rendered {
        const steps = this.steps(model);
        while (!steps.next().done) {
            // drained
        }
        return this.result();
    }

    // Live view of the result: the arrays grow while the steps run.
    result(): Rendered {
        return { pages: this.pages, images: this.images, words: this.words };
    }

    // Lays the document out and renders it page by page; every step hands the main thread back.
    *steps(model: Model): Generator<void> {
        this.model = model;
        const font = model.font || {};
        this.root.style.setProperty('--dx-font', fontStack(font.family) || DEFAULT_STACK);
        this.root.style.setProperty('--dx-font-size', (font.size || 11) + 'pt');
        const layout = (this.layout = createLayout(model, createMeasure()));
        const engine = layout.steps();
        let shown = 0;
        for (;;) {
            const done = !!engine.next().done;
            // The page in progress may still grow; the ones before it are final.
            const ready = done ? layout.pages.length : layout.pages.length - 1;
            while (shown < ready) {
                this.renderPage(layout.pages[shown++]);
                yield;
            }
            if (done) {
                break;
            }
        }
        if (model.notes && model.notes.length && this.pages.length) {
            const sheet = this.pages[this.pages.length - 1];
            const parts = this.parts.get(sheet)!;
            const notes = h('div', 'dx-notes');
            notes.contentEditable = 'false';
            const width = this.columnOf(parts.page);
            for (const block of model.notes) {
                if (block.type === 'p') {
                    notes.appendChild(this.paragraph(block, width));
                }
            }
            parts.body.appendChild(notes);
        }
    }

    // ---------- pages ----------

    private pageOf(section: number): Page {
        const sections = this.model && this.model.sections;
        return (
            (sections && sections[section] && sections[section].page) || (this.model && this.model.page) || DEFAULT_PAGE
        );
    }

    private columnOf(page: Page): number {
        return page.width - page.left - page.right;
    }

    private renderPage(page: any) {
        const blocks = this.model!.blocks;
        const sheet = this.makeSheet(page);
        const parts = this.parts.get(sheet)!;
        const width = this.columnOf(parts.page);
        let prev: { block: any; el: HTMLElement } | null = null;
        for (const part of page.parts) {
            const block = blocks[part.block];
            if (part.empty) {
                continue;
            }
            if (part.rows) {
                parts.body.appendChild(this.table(block, part.rows[0], part.rows[1]));
                prev = null;
                continue;
            }
            if (part.whole) {
                parts.body.appendChild(this.box(block, width));
                prev = null;
                continue;
            }
            const el = this.paragraph(block, width, part.lines, part.first, part.last);
            if (part.first) {
                if (!prev && page.hard) {
                    // A page that a manual break started drops the space before.
                    el.style.marginTop = '0';
                }
                if (prev && block.sid && prev.block.sid === block.sid) {
                    // Contextual spacing: no space between paragraphs of one style.
                    if (block.layout && block.layout.contextual) {
                        el.style.marginTop = '0';
                    }
                    if (prev.block.layout && prev.block.layout.contextual) {
                        prev.el.style.marginBottom = '0';
                    }
                }
            }
            parts.body.appendChild(el);
            prev = { block, el };
        }
    }

    private makeSheet(page: any): HTMLElement {
        const p = this.pageOf(page.section);
        const sheet = h('div', 'dx-sheet');
        sheet.style.setProperty('--dx-page-width', p.width + 'pt');
        sheet.style.setProperty('--dx-page-height', p.height + 'pt');
        sheet.style.paddingLeft = p.left + 'pt';
        sheet.style.paddingRight = p.right + 'pt';
        const top = h('div', 'dx-top dx-hf');
        top.style.minHeight = p.top + 'pt';
        top.style.paddingTop = (p.header ?? 36) + 'pt';
        const body = h('div', 'dx-body');
        const bottom = h('div', 'dx-bottom dx-hf');
        bottom.style.minHeight = p.bottom + 'pt';
        bottom.style.paddingBottom = (p.footer ?? 36) + 'pt';
        top.contentEditable = bottom.contentEditable = 'false';
        sheet.append(top, body, bottom);
        this.parts.set(sheet, {
            page: p,
            section: page.section,
            number: page.number,
            kind: page.kind,
            top,
            body,
            bottom
        });
        const sections = this.model!.sections;
        if (sections && sections[page.section]) {
            this.attachHeaderFooter(sheet, sections[page.section], page.kind);
        }
        this.root.appendChild(sheet);
        this.pages.push(sheet);
        return sheet;
    }

    // Fills in the page counters of the headers and footers; the total grows while the document is laid out.
    numberPages() {
        const total = String(this.layout ? this.layout.pages.length : this.pages.length);
        for (const sheet of this.pages) {
            const parts = this.parts.get(sheet)!;
            for (const span of Array.from(sheet.querySelectorAll<HTMLElement>('.dx-hf [data-field]'))) {
                const text = span.dataset.field === 'page' ? String(parts.number) : total;
                if (span.textContent !== text) {
                    span.textContent = text;
                }
            }
        }
    }

    private attachHeaderFooter(sheet: HTMLElement, section: Section, kind: string) {
        const parts = this.parts.get(sheet)!;
        const width = this.columnOf(parts.page);
        const fill = (el: HTMLElement, blocks: any[] | undefined, edge: 'top' | 'bottom', distance: number) => {
            el.replaceChildren();
            // Text of headers and footers does not count as words of the document.
            const words = this.words;
            for (const b of blocks || []) {
                if (b.type === 'pagebreak' || b.type === 'section') {
                    continue;
                }
                const child =
                    b.type === 'table'
                        ? this.table(b)
                        : b.type === 'box'
                          ? this.box(b, width)
                          : this.paragraph(b, width);
                if (b.layout && b.layout.frame) {
                    // A framed paragraph (the page number) is placed by its frame and takes no room.
                    child.classList.add('dx-framed');
                    child.style[edge] = pt(distance);
                }
                el.appendChild(child);
            }
            this.words = words;
        };
        fill(parts.top, section.header[kind], 'top', parts.page.header ?? 36);
        fill(parts.bottom, section.footer[kind], 'bottom', parts.page.footer ?? 36);
    }

    // ---------- paragraphs ----------

    private linesOf(block: any, width: number): any[] {
        let byWidth = this.lineCache.get(block);
        if (!byWidth) {
            byWidth = new Map();
            this.lineCache.set(block, byWidth);
        }
        let lines = byWidth.get(width);
        if (!lines) {
            lines = this.layout.linesOf(block, width);
            byWidth.set(width, lines!);
        }
        return lines!;
    }

    // The font and size that set the height of the lines while the text flows freely (editing): the ones most of
    // its text uses, or those of the paragraph mark when there is no text.
    private dominant(block: any): { font?: string; size?: number } {
        const weight = new Map<string, { font?: string; size?: number; n: number }>();
        for (const run of block.runs) {
            const n = run.text && !run.br && !run.image ? run.text.length : 0;
            if (n) {
                const key = (run.font || '') + '|' + (run.size || '');
                const entry = weight.get(key) || { font: run.font, size: run.size, n: 0 };
                entry.n += n;
                weight.set(key, entry);
            }
        }
        let best: { font?: string; size?: number; n: number } | undefined;
        for (const entry of weight.values()) {
            if (!best || entry.n > best.n) {
                best = entry;
            }
        }
        return best || { font: block.markFont, size: block.markSize };
    }

    // A paragraph, or the lines [from, to) of it that one page shows, drawn line by line for a column `width`
    // points wide.
    private paragraph(block: any, width: number, range?: [number, number], first = true, last = true): HTMLElement {
        const p = h('div', 'dx-p' + (block.style ? ' dx-' + block.style : '') + (first ? '' : ' dx-cont'));
        if (block.id !== undefined && this.editable) {
            p.dataset.p = String(block.id);
        }
        const s = p.style;
        const layout = block.layout || {};
        const left = block.list && !layout.left ? 18 * (block.list.level + 1) : layout.left || 0;
        if (block.align) {
            s.textAlign = block.align;
        }
        if (left) {
            s.marginLeft = pt(left);
        }
        if (layout.right) {
            s.marginRight = pt(layout.right);
        }
        // Space before belongs to the start of a paragraph, space after to its end.
        if (layout.before !== undefined && first) {
            s.marginTop = pt(layout.before);
        }
        if (layout.after !== undefined && last) {
            s.marginBottom = pt(layout.after);
        }
        if (layout.first && first && !block.list) {
            s.textIndent = pt(layout.first);
        }
        if (layout.shade) {
            s.background = layout.shade;
        }
        // How the text flows while it is edited.
        const dom = this.dominant(block);
        const model = this.model;
        const stack = fontStack(dom.font) || fontStack(model && model.font && model.font.family) || DEFAULT_STACK;
        if (dom.font) {
            s.fontFamily = stack;
        }
        if (dom.size) {
            s.fontSize = dom.size + 'pt';
        }
        const spacing = layout.line;
        s.lineHeight =
            spacing && spacing.pt
                ? pt(spacing.pt)
                : String(lineRatio(stack) * (spacing ? Math.max(0.1, spacing.factor) : 1));

        const lines = this.linesOf(block, width);
        const lastLine = lines[lines.length - 1];
        const [from, to] = range || [0, lines.length];
        let text = '';
        for (let k = from; k < to; k++) {
            const l = lines[k];
            const ln = h('div', 'dx-ln');
            const ls = ln.style;
            ls.height = pt(l.height);
            ls.lineHeight = pt(l.single);
            let shift = 0;
            const free = l.room - l.width;
            if (block.align === 'center') {
                shift = free / 2;
            } else if (block.align === 'right') {
                shift = free;
            } else if (block.align === 'justify' && l !== lastLine && l.spaces > 0 && free > 0) {
                // Word justifies by stretching the spaces of the line.
                ls.wordSpacing = pt(free / l.spaces);
            }
            const indent = l.offset - left + Math.max(0, shift);
            if (indent) {
                ls.paddingLeft = pt(indent);
            }
            if (k === 0 && block.list) {
                ln.appendChild(this.marker(block));
            }
            const pieces = this.slice(block.runs, l.start, l.end);
            for (const piece of pieces) {
                const node = this.run(piece.run);
                if (piece.run.text === '\t') {
                    const tab = l.tabs.find((t: any) => t.at === piece.at);
                    if (tab) {
                        (node as HTMLElement).style.width = pt(tab.width);
                    }
                }
                ln.appendChild(node);
                if (piece.run.text && !piece.run.field) {
                    text += piece.run.text;
                }
            }
            if (!pieces.length) {
                // An empty line keeps its height, and it survives a copy.
                ln.appendChild(h('br'));
            }
            p.appendChild(ln);
        }
        const words = text.match(/[\p{L}\p{N}]+(?:['’-][\p{L}\p{N}]+)*/gu);
        this.words += words ? words.length : 0;
        return p;
    }

    private marker(block: any): HTMLElement {
        const layout = block.layout || {};
        const marker = h('span', 'dx-marker', block.list.marker);
        marker.contentEditable = 'false';
        const f = block.list.format;
        if (f) {
            if (f.font) {
                marker.style.fontFamily = fontStack(f.font);
            }
            if (f.size) {
                marker.style.fontSize = f.size + 'pt';
            }
            if (f.b) {
                marker.style.fontWeight = 'bold';
            }
            if (f.i) {
                marker.style.fontStyle = 'italic';
            }
            if (f.color) {
                marker.style.color = f.color;
            }
        }
        // The marker hangs in the indent unless the paragraph says otherwise.
        const hang = layout.first !== undefined && layout.first < 0 ? -layout.first : 18;
        marker.style.width = pt(hang);
        marker.style.marginLeft = pt(-hang);
        return marker;
    }

    // The runs of characters `from` to `to`, cut where the range cuts them; a piece keeps its run's slot, so the
    // parts of one run on two lines or pages still save as one.
    private slice(runs: any[], from: number, to: number): { run: any; at: number }[] {
        const out: { run: any; at: number }[] = [];
        let pos = 0;
        for (const run of runs) {
            const len = runLength(run);
            const end = pos + len;
            if (len === 0) {
                if (run.field && pos >= from && pos < Math.max(to, from + 1)) {
                    out.push({ run, at: pos });
                }
            } else if (end > from && pos < to) {
                if (run.text && !run.atom && (pos < from || end > to)) {
                    const a = Math.max(from, pos) - pos;
                    const b = Math.min(to, end) - pos;
                    out.push({ run: Object.assign({}, run, { text: run.text.slice(a, b) }), at: pos + a });
                } else {
                    out.push({ run, at: pos });
                }
            }
            pos = end;
        }
        return out;
    }

    private run(run: any): Node {
        if (run.br) {
            const br = h('br');
            if (run.slot !== undefined) {
                br.dataset.s = String(run.slot);
            }
            return br;
        }
        if (run.image) {
            return this.image(run);
        }
        const span = h('span', run.text === '\t' ? 'dx-tab' : undefined, run.text);
        if (run.slot !== undefined) {
            span.dataset.s = String(run.slot);
        }
        if (run.atom || run.field) {
            span.contentEditable = 'false';
        }
        if (run.field) {
            span.dataset.field = run.field;
            span.textContent = '0';
        }
        const s = span.style;
        if (run.b) {
            s.fontWeight = 'bold';
        }
        if (run.i) {
            s.fontStyle = 'italic';
        }
        const decorations = [];
        if (run.u) {
            decorations.push('underline');
        }
        if (run.s) {
            decorations.push('line-through');
        }
        if (decorations.length) {
            s.textDecoration = decorations.join(' ');
        }
        if (run.size) {
            s.fontSize = run.size + 'pt';
        }
        if (run.sup || run.sub) {
            s.verticalAlign = run.sup ? 'super' : 'sub';
            s.fontSize = (run.size || 11) * 0.65 + 'pt';
            s.lineHeight = '0';
        }
        if (run.color) {
            s.color = run.color;
        }
        if (run.highlight) {
            s.background = run.highlight;
        }
        if (run.font) {
            s.fontFamily = fontStack(run.font);
        }
        if (run.caps) {
            s.textTransform = 'uppercase';
        }
        if (run.smallCaps) {
            s.fontVariant = 'small-caps';
        }
        if (run.href && /^(https?:|mailto:)/i.test(run.href)) {
            const a = h('a', 'dx-link');
            a.href = run.href;
            a.title = run.href + (this.editable ? '\nCtrl+click to open' : '');
            a.appendChild(span);
            return a;
        }
        return span;
    }

    private image(run: any): Node {
        const { part, type, width, height } = run.image;
        const files = this.files;
        let node: HTMLElement;
        if (files && type && part in files) {
            const img = h('img', 'dx-img');
            img.alt = '';
            img.decoding = 'async';
            node = img;
            this.lazyImage(img, () => {
                // Reading the part inflates it; only pictures that come into view pay for it.
                const data = files[part];
                if (data) {
                    const url = URL.createObjectURL(new Blob([data as BlobPart], { type }));
                    this.images.push(url);
                    img.src = url;
                }
            });
        } else {
            // EMF/WMF and missing parts cannot be shown by the browser.
            node = h('span', 'dx-img-missing', 'image');
        }
        node.contentEditable = 'false';
        if (width) {
            node.style.width = width + 'pt';
        }
        if (height) {
            node.style.height = height + 'pt';
        }
        if (run.slot !== undefined) {
            node.dataset.s = String(run.slot);
        }
        return node;
    }

    private lazyImage(img: HTMLElement, load: () => void) {
        if (!this.viewport || typeof IntersectionObserver !== 'function') {
            load();
            return;
        }
        if (!this.observer) {
            this.observer = new IntersectionObserver(
                entries => {
                    for (const entry of entries) {
                        if (entry.isIntersecting) {
                            this.observer && this.observer.unobserve(entry.target);
                            const pending = this.pending.get(entry.target);
                            this.pending.delete(entry.target);
                            if (pending && !this.disposed) {
                                pending();
                            }
                        }
                    }
                },
                { root: this.viewport, rootMargin: '1500px 0px' }
            );
        }
        this.pending.set(img, load);
        this.observer.observe(img);
    }

    // ---------- tables and boxes ----------

    private box(block: any, width: number): HTMLElement {
        const box = h('div', 'dx-box');
        box.contentEditable = 'false';
        if (block.width) {
            box.style.maxWidth = block.width + 'pt';
        }
        const inner = Math.min(width, block.width || width);
        for (const b of block.blocks) {
            if (b.type === 'table') {
                box.appendChild(this.table(b));
            } else if (b.type === 'p') {
                box.appendChild(this.paragraph(b, inner));
            }
        }
        return box;
    }

    private row(block: any, r: number, widths: number[][]): HTMLElement {
        const tr = h('tr');
        block.rows[r].forEach((cell: any, c: number) => {
            // The .doc reader gives cells as plain block lists.
            const blocks = Array.isArray(cell) ? cell : cell.blocks;
            const td = h('td');
            if (cell.colspan) {
                td.colSpan = cell.colspan;
            }
            if (cell.rowspan) {
                td.rowSpan = cell.rowspan;
            }
            if (cell.fill) {
                td.style.background = cell.fill;
            }
            for (const b of blocks) {
                if (b.type === 'table') {
                    td.appendChild(this.table(b));
                } else if (b.type === 'box') {
                    td.appendChild(this.box(b, widths[r][c]));
                } else if (b.type === 'p') {
                    td.appendChild(this.paragraph(b, widths[r][c]));
                }
            }
            tr.appendChild(td);
        });
        return tr;
    }

    // The rows `from` to `to` of a table; a part that continues the table on another page starts with its header
    // rows again. They are a copy for the eye only: no paragraph ids or slots, so edits and saves never see them twice.
    private table(block: any, from = 0, to = block.rows.length): HTMLElement {
        const wrap = h('div', 'dx-table-wrap');
        const table = h('table', 'dx-table' + (block.borders === false ? '' : ' dx-borders'));
        if (block.grid && block.grid.length && block.grid.every((w: number) => w > 0)) {
            table.classList.add('dx-fixed');
            table.style.width = pt(block.grid.reduce((a: number, b: number) => a + b, 0));
        }
        if (block.grid && block.grid.length) {
            const colgroup = h('colgroup');
            for (const w of block.grid) {
                const col = h('col');
                if (w) {
                    col.style.width = w + 'pt';
                }
                colgroup.appendChild(col);
            }
            table.appendChild(colgroup);
        }
        const body = h('tbody');
        const widths = cellWidths(block);
        if (from > 0 && block.headerRows) {
            const words = this.words;
            for (let r = 0; r < block.headerRows; r++) {
                const tr = this.row(block, r, widths);
                tr.classList.add('dx-rep');
                tr.contentEditable = 'false';
                for (const el of Array.from(tr.querySelectorAll<HTMLElement>('[data-p], [data-s]'))) {
                    delete el.dataset.p;
                    delete el.dataset.s;
                }
                body.appendChild(tr);
            }
            this.words = words;
        }
        for (let r = from; r < to; r++) {
            body.appendChild(this.row(block, r, widths));
        }
        table.appendChild(body);
        wrap.appendChild(table);
        return wrap;
    }
}
