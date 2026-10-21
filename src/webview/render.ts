// Renders the block model of doc.js / docx.js as pages of HTML.
// Paragraphs carry data-p (the paragraph id of an editable .docx), rendered pieces carry data-s (their slot),
// which is how an edit in the page finds its way back into the XML.
import { h } from './shell';

export interface Page {
    width: number;
    height: number;
    top: number;
    right: number;
    bottom: number;
    left: number;
}

export interface Model {
    blocks: any[];
    notes?: any[];
    page?: Page;
    font?: { family?: string; size?: number };
}

// A4 with 2 cm margins for the binary format, which keeps its page setup in the section properties dx does not read.
const DEFAULT_PAGE: Page = { width: 595.3, height: 841.9, top: 56.7, right: 56.7, bottom: 56.7, left: 56.7 };

const GENERIC_SERIF = /times|georgia|cambria|garamond|serif|book|minion|palatino/i;
const GENERIC_MONO = /courier|consolas|mono|menlo/i;

export function fontStack(name: string | undefined): string {
    if (!name) {
        return '';
    }
    const generic = GENERIC_MONO.test(name) ? 'monospace' : GENERIC_SERIF.test(name) ? 'serif' : 'sans-serif';
    return `"${name.replace(/["\\]/g, '')}", ${generic}`;
}

export interface Rendered {
    pages: HTMLElement[];
    images: string[];
    words: number;
}

export class Renderer {
    private pages: HTMLElement[] = [];
    private sheet!: HTMLElement;
    private images: string[] = [];
    private words = 0;

    constructor(
        private readonly root: HTMLElement,
        private readonly files: { [name: string]: Uint8Array } | undefined,
        private readonly editable: boolean
    ) {}

    render(model: Model): Rendered {
        const page = model.page || DEFAULT_PAGE;
        const font = model.font || {};
        this.root.style.setProperty('--dx-page-width', page.width + 'pt');
        this.root.style.setProperty('--dx-page-height', page.height + 'pt');
        this.root.style.setProperty(
            '--dx-page-padding',
            `${page.top}pt ${page.right}pt ${page.bottom}pt ${page.left}pt`
        );
        this.root.style.setProperty('--dx-font', fontStack(font.family) || 'Calibri, Carlito, sans-serif');
        this.root.style.setProperty('--dx-font-size', (font.size || 11) + 'pt');
        this.newSheet();
        for (const block of model.blocks) {
            if (block.type === 'pagebreak') {
                this.newSheet();
                continue;
            }
            if (block.pageStart && this.sheet.childNodes.length) {
                this.newSheet();
            }
            this.sheet.appendChild(this.block(block));
        }
        if (model.notes && model.notes.length) {
            const notes = h('div', 'dx-notes');
            notes.contentEditable = 'false';
            for (const block of model.notes) {
                if (block.type !== 'pagebreak') {
                    notes.appendChild(this.block(block));
                }
            }
            this.sheet.appendChild(notes);
        }
        // A trailing page break leaves an empty last sheet.
        if (this.pages.length > 1 && !this.sheet.childNodes.length) {
            this.sheet.remove();
            this.pages.pop();
        }
        return { pages: this.pages, images: this.images, words: this.words };
    }

    private newSheet() {
        this.sheet = h('div', 'dx-sheet');
        this.root.appendChild(this.sheet);
        this.pages.push(this.sheet);
    }

    private block(block: any): HTMLElement {
        switch (block.type) {
            case 'table':
                return this.table(block);
            case 'box': {
                const box = h('div', 'dx-box');
                box.contentEditable = 'false';
                if (block.width) {
                    box.style.maxWidth = block.width + 'pt';
                }
                for (const b of block.blocks) {
                    if (b.type !== 'pagebreak') {
                        box.appendChild(this.block(b));
                    }
                }
                return box;
            }
            default:
                return this.paragraph(block);
        }
    }

    private paragraph(block: any): HTMLElement {
        const p = h('div', 'dx-p' + (block.style ? ' dx-' + block.style : ''));
        if (block.id !== undefined && this.editable) {
            p.dataset.p = String(block.id);
        }
        const s = p.style;
        if (block.align) {
            s.textAlign = block.align;
        }
        const layout = block.layout || {};
        let left = layout.left || 0;
        if (block.list) {
            if (!layout.left) {
                left = 18 * (block.list.level + 1);
            }
            p.classList.add('dx-li');
            const marker = h('span', 'dx-marker', block.list.marker);
            marker.contentEditable = 'false';
            // The marker hangs in the indent unless the paragraph says otherwise.
            const hang = layout.first !== undefined && layout.first < 0 ? -layout.first : 18;
            marker.style.width = hang + 'pt';
            marker.style.marginLeft = -hang + 'pt';
            p.appendChild(marker);
        } else if (layout.first) {
            s.textIndent = layout.first + 'pt';
        }
        if (block.indent) {
            left = Math.max(left, block.indent * 0.75);
        }
        if (left) {
            s.marginLeft = left + 'pt';
        }
        if (layout.right) {
            s.marginRight = layout.right + 'pt';
        }
        if (layout.before !== undefined) {
            s.marginTop = layout.before + 'pt';
        }
        if (layout.after !== undefined) {
            s.marginBottom = layout.after + 'pt';
        }
        if (layout.line) {
            // Word's single spacing is the font line height, a little more than 1.
            s.lineHeight = layout.line.pt ? layout.line.pt + 'pt' : String(Math.max(0.6, layout.line.factor) * 1.08);
        }
        if (layout.shade) {
            s.background = layout.shade;
        }
        let text = '';
        for (const run of block.runs) {
            p.appendChild(this.run(run));
            if (run.text) {
                text += run.text;
            }
        }
        const words = text.match(/[\p{L}\p{N}]+(?:['’-][\p{L}\p{N}]+)*/gu);
        this.words += words ? words.length : 0;
        return p;
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
        if (run.atom) {
            span.contentEditable = 'false';
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
        const data = this.files && this.files[part];
        let node: HTMLElement;
        if (data && type) {
            const url = URL.createObjectURL(new Blob([data as BlobPart], { type }));
            this.images.push(url);
            const img = h('img', 'dx-img');
            img.src = url;
            img.alt = '';
            node = img;
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

    private table(block: any): HTMLElement {
        const wrap = h('div', 'dx-table-wrap');
        const table = h('table', 'dx-table' + (block.borders === false ? '' : ' dx-borders'));
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
        for (const row of block.rows) {
            const tr = h('tr');
            for (const cell of row) {
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
                    if (b.type !== 'pagebreak') {
                        td.appendChild(this.block(b));
                    }
                }
                tr.appendChild(td);
            }
            body.appendChild(tr);
        }
        table.appendChild(body);
        wrap.appendChild(table);
        return wrap;
    }
}
