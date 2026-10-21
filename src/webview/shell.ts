// The frame every dx viewer lives in, the same layout and look as the xl shell:
// a toolbar on top, the document in the middle with a floating find widget, a status bar at the bottom.

declare function acquireVsCodeApi(): any;

export interface Options {
    kind: 'pdf' | 'docx' | 'doc';
    name: string;
    uri: string;
    media: string;
    state: any;
    test?: boolean;
}

export const vscode = acquireVsCodeApi();

export function readOptions(): Options {
    const node = document.getElementById('dx-options');
    try {
        return JSON.parse((node && node.textContent) || '{}');
    } catch {
        return {} as Options;
    }
}

export function post(message: any): void {
    try {
        vscode.postMessage(message);
    } catch {
        // The webview is being torn down.
    }
}

export function errorText(error: any): string {
    if (!error) {
        return 'unknown error';
    }
    return typeof error === 'string' ? error : String(error.message || error);
}

export function h<K extends keyof HTMLElementTagNameMap>(
    tag: K,
    cls?: string,
    text?: string
): HTMLElementTagNameMap[K] {
    const e = document.createElement(tag);
    if (cls) {
        e.className = cls;
    }
    if (text !== undefined) {
        e.textContent = text;
    }
    return e;
}

export function icon(path: string): SVGSVGElement {
    const s = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    s.setAttribute('viewBox', '0 0 16 16');
    s.setAttribute('width', '14');
    s.setAttribute('height', '14');
    const p = document.createElementNS('http://www.w3.org/2000/svg', 'path');
    p.setAttribute('d', path);
    p.setAttribute('fill', 'none');
    p.setAttribute('stroke', 'currentColor');
    p.setAttribute('stroke-width', '1.5');
    p.setAttribute('stroke-linecap', 'round');
    p.setAttribute('stroke-linejoin', 'round');
    s.appendChild(p);
    return s;
}

export const ICONS = {
    search: 'M7 12A5 5 0 1 0 7 2a5 5 0 0 0 0 10zM11 11l3.5 3.5',
    up: 'M3 10l5-5 5 5',
    down: 'M3 6l5 5 5-5',
    close: 'M4 4l8 8M12 4l-8 8',
    fitWidth: 'M2 3v10M14 3v10M5 8h6M5 8l2-2M5 8l2 2M11 8l-2-2M11 8l-2 2',
    fitPage: 'M4 2.5h8v11H4zM6.5 6h3M6.5 8.5h3M6.5 11h1.5',
    rotate: 'M13 8a5 5 0 1 1-1.5-3.5M13 2.5v3h-3',
    pen: 'M10.5 2.5l3 3-7.5 7.5H3v-3z',
    highlight: 'M2.5 13.5h4M4 12l1-3 5.5-5.5 2 2L7 11zM9.5 4.5l2 2',
    text: 'M3 3.5h10M8 3.5v9.5M6 13h4',
    ink: 'M2 12c2-4 3.5-6.5 5-3.5s2 3 3.5-.5 2-3 3.5-2'
};

export function button(cls: string, title: string, content: string | SVGSVGElement): HTMLButtonElement {
    const b = h('button', 'dx-btn' + (cls ? ' ' + cls : ''));
    b.title = title;
    b.type = 'button';
    if (typeof content === 'string') {
        b.textContent = content;
    } else {
        b.appendChild(content);
    }
    return b;
}

export interface FindQuery {
    query: string;
    caseSensitive: boolean;
    entireWord: boolean;
}

export interface FindHandler {
    // A new or changed query; the viewer jumps to the first match from the current position.
    find(q: FindQuery): void;
    next(q: FindQuery, previous: boolean): void;
    clear(): void;
}

export interface ZoomHandler {
    get(): number;
    set(zoom: number): void;
    reset(): void;
}

const ZOOM_MIN = 0.25;
const ZOOM_MAX = 5;
const ZOOM_STEP = 1.15;

export class Shell {
    readonly app: HTMLElement;
    readonly bar: HTMLElement;
    readonly left: HTMLElement;
    readonly right: HTMLElement;
    readonly main: HTMLElement;
    readonly status: HTMLElement;
    private readonly msg: HTMLElement;
    private readonly info: HTMLElement;
    private readonly zoomBox: HTMLElement;
    private readonly zVal: HTMLElement;
    private readonly find: {
        el: HTMLElement;
        input: HTMLInputElement;
        count: HTMLElement;
        mc: HTMLButtonElement;
        ww: HTMLButtonElement;
    };
    private findTimer: any = 0;
    findHandler: FindHandler | undefined;
    zoomHandler: ZoomHandler | undefined;

    constructor(app: HTMLElement) {
        this.app = app;
        app.classList.add('dx-app');

        this.bar = h('div', 'dx-bar');
        this.left = h('div', 'dx-group');
        this.right = h('div', 'dx-group dx-end');
        const btnFind = button('', 'Find (Ctrl+F)', icon(ICONS.search));
        btnFind.addEventListener('click', () => this.openFind());
        this.right.appendChild(btnFind);
        this.bar.append(this.left, this.right);

        this.main = h('div', 'dx-main');
        this.find = this.buildFind();

        this.status = h('div', 'dx-status');
        this.msg = h('span', 'dx-msg', 'Opening…');
        this.info = h('span', 'dx-info');
        this.zoomBox = h('span', 'dx-zoom');
        const zOut = button('dx-z', 'Zoom out (Ctrl+wheel)', '−');
        this.zVal = h('span', 'dx-zval', '100%');
        this.zVal.title = 'Reset zoom';
        const zIn = button('dx-z', 'Zoom in (Ctrl+wheel)', '+');
        this.zoomBox.append(zOut, this.zVal, zIn);
        this.zoomBox.style.display = 'none';
        this.status.append(this.msg, this.info, this.zoomBox);

        const frame = h('div', 'dx-frame');
        frame.append(this.main, this.find.el);
        app.append(this.bar, frame, this.status);

        zOut.addEventListener('click', () => this.zoomBy(1 / ZOOM_STEP));
        zIn.addEventListener('click', () => this.zoomBy(ZOOM_STEP));
        this.zVal.addEventListener('click', () => this.zoomHandler && this.zoomHandler.reset());
        this.main.addEventListener(
            'wheel',
            e => {
                if ((e.ctrlKey || e.metaKey) && this.zoomHandler) {
                    e.preventDefault();
                    // Trackpad pinches arrive as small ctrl+wheel deltas: zoom proportionally.
                    this.zoomBy(Math.exp(-Math.max(-50, Math.min(50, e.deltaY)) / 200));
                }
            },
            { passive: false }
        );
        document.addEventListener('keydown', e => this.onKey(e));
    }

    private buildFind() {
        const el = h('div', 'dx-find');
        el.style.display = 'none';
        const input = h('input', 'dx-find-in');
        input.placeholder = 'Find';
        input.spellcheck = false;
        input.setAttribute('aria-label', 'Find');
        const count = h('span', 'dx-find-count');
        const mc = button('dx-tog', 'Match case', 'Aa');
        const ww = button('dx-tog', 'Match whole word', 'ab');
        ww.style.textDecoration = 'underline';
        const prev = button('', 'Previous match (Shift+Enter)', icon(ICONS.up));
        const next = button('', 'Next match (Enter)', icon(ICONS.down));
        const close = button('', 'Close (Escape)', icon(ICONS.close));
        el.append(input, count, mc, ww, prev, next, close);

        input.addEventListener('input', () => {
            clearTimeout(this.findTimer);
            this.findTimer = setTimeout(() => {
                this.findTimer = 0;
                this.runFind();
            }, 160);
        });
        input.addEventListener('keydown', e => {
            if (e.key === 'Enter') {
                e.preventDefault();
                if (this.findTimer) {
                    clearTimeout(this.findTimer);
                    this.findTimer = 0;
                    this.runFind();
                } else {
                    this.findNext(e.shiftKey);
                }
            } else if (e.key === 'Escape') {
                e.preventDefault();
                this.closeFind();
            }
        });
        const toggle = (b: HTMLButtonElement) => {
            b.classList.toggle('on');
            b.setAttribute('aria-pressed', String(b.classList.contains('on')));
            this.runFind();
        };
        mc.addEventListener('click', () => toggle(mc));
        ww.addEventListener('click', () => toggle(ww));
        prev.addEventListener('click', () => this.findNext(true));
        next.addEventListener('click', () => this.findNext(false));
        close.addEventListener('click', () => this.closeFind());
        return { el, input, count, mc, ww };
    }

    private query(): FindQuery {
        return {
            query: this.find.input.value,
            caseSensitive: this.find.mc.classList.contains('on'),
            entireWord: this.find.ww.classList.contains('on')
        };
    }

    private runFind() {
        if (!this.findHandler) {
            return;
        }
        const q = this.query();
        if (!q.query) {
            this.findHandler.clear();
            this.setFindCount(0, 0, false);
            return;
        }
        this.findHandler.find(q);
    }

    private findNext(previous: boolean) {
        const q = this.query();
        if (this.findHandler && q.query) {
            this.findHandler.next(q, previous);
        }
    }

    openFind() {
        const f = this.find;
        f.el.style.display = '';
        const selection = window.getSelection();
        const selected = selection ? selection.toString().trim() : '';
        if (selected && selected.length < 200 && selected.indexOf('\n') < 0 && selected !== f.input.value) {
            f.input.value = selected;
            this.runFind();
        }
        f.input.focus();
        f.input.select();
    }

    closeFind() {
        this.find.el.style.display = 'none';
        if (this.findHandler) {
            this.findHandler.clear();
        }
        this.setFindCount(0, 0, false);
        this.main.focus();
    }

    get findOpen(): boolean {
        return this.find.el.style.display !== 'none';
    }

    // current is 1-based; total 0 with a query means nothing was found.
    setFindCount(current: number, total: number, searched = true) {
        const c = this.find.count;
        c.classList.toggle('none', searched && total === 0 && !!this.find.input.value);
        if (!searched || !this.find.input.value) {
            c.textContent = '';
        } else if (total === 0) {
            c.textContent = 'No results';
        } else {
            c.textContent = (current > 0 ? current : '?') + ' of ' + total;
        }
    }

    private onKey(e: KeyboardEvent) {
        const mod = e.ctrlKey || e.metaKey;
        if (mod && !e.altKey && (e.key === 'f' || e.key === 'F')) {
            e.preventDefault();
            this.openFind();
        } else if (e.key === 'F3' || (mod && (e.key === 'g' || e.key === 'G'))) {
            e.preventDefault();
            if (!this.findOpen) {
                this.openFind();
            } else {
                this.findNext(e.shiftKey);
            }
        } else if (e.key === 'Escape' && this.findOpen) {
            e.preventDefault();
            this.closeFind();
        }
    }

    enableZoom(handler: ZoomHandler) {
        this.zoomHandler = handler;
        this.zoomBox.style.display = '';
        this.showZoom(handler.get());
    }

    zoomBy(factor: number) {
        if (!this.zoomHandler) {
            return;
        }
        const next = Math.min(ZOOM_MAX, Math.max(ZOOM_MIN, this.zoomHandler.get() * factor));
        this.zoomHandler.set(Math.round(next * 1000) / 1000);
    }

    showZoom(zoom: number) {
        this.zVal.textContent = Math.round(zoom * 100) + '%';
    }

    setMessage(text: string) {
        this.msg.textContent = text;
        this.msg.title = text;
    }

    setInfo(text: string) {
        this.info.textContent = text;
    }

    // A full-size message in place of the document, like the editor placeholder of VS Code.
    showError(title: string, detail?: string) {
        this.main.textContent = '';
        const box = h('div', 'dx-error');
        box.appendChild(h('div', 'dx-error-title', title));
        if (detail) {
            box.appendChild(h('div', 'dx-error-detail', detail));
        }
        this.main.appendChild(box);
        this.setMessage(title);
        this.zoomBox.style.display = 'none';
    }
}

// Remembers the view position of the document, in the webview state and in the workspace state of the host.
export class ViewState {
    private timer: any;

    constructor(
        private readonly options: Options,
        private readonly version: string
    ) {}

    load(): any {
        const s = vscode.getState() || this.options.state;
        return s && s.version === this.version && s.uri === this.options.uri ? s : null;
    }

    save(data: any) {
        clearTimeout(this.timer);
        this.timer = setTimeout(() => {
            const state = Object.assign({ version: this.version, uri: this.options.uri }, data);
            vscode.setState(state);
            post({ save: true, state });
        }, 250);
    }
}

// Converts whatever carried the file bytes through postMessage into a Uint8Array.
export function toBytes(body: any): Uint8Array | undefined {
    if (body instanceof Uint8Array) {
        return body;
    }
    if (body instanceof ArrayBuffer) {
        return new Uint8Array(body);
    }
    if (ArrayBuffer.isView(body)) {
        return new Uint8Array(body.buffer, body.byteOffset, body.byteLength);
    }
    if (body && typeof body === 'object' && Array.isArray(body.data)) {
        return Uint8Array.from(body.data);
    }
    return undefined;
}

// Opens http(s) and mailto links through the host, the webview itself never navigates.
export function interceptLinks(root: HTMLElement, allow: () => boolean = () => true) {
    root.addEventListener(
        'click',
        e => {
            const target = e.target as Element;
            const a = target && target.closest ? (target.closest('a[href]') as HTMLAnchorElement) : null;
            if (!a) {
                return;
            }
            const href = a.getAttribute('href') || '';
            if (/^(https?:|mailto:)/i.test(href)) {
                e.preventDefault();
                e.stopPropagation();
                if (allow() || e.ctrlKey || e.metaKey) {
                    post({ openLink: href });
                }
            }
        },
        true
    );
}
