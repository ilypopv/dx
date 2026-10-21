'use strict';
import { window, ExtensionContext, ExtensionMode, Memento, WebviewPanel, CancellationToken, Disposable } from 'vscode';
import { URI, Utils } from 'vscode-uri';
import { safeJson, makeNonce, csp } from './html';
import { openExternalLink } from './links';
import type { DxDocument, Kind } from './editor';

const DATA_TIMEOUT_MS = 60000;
const MAX_STATE_CHARS = 1024 * 1024;
const MAX_FILE_BYTES = 1024 * 1024 * 1024;

function isSaneState(state: any): boolean {
    if (state === null) {
        return true;
    }
    if (typeof state !== 'object') {
        return false;
    }
    try {
        return JSON.stringify(state).length <= MAX_STATE_CHARS;
    } catch (err) {
        return false;
    }
}

// The bytes the webview hands back for a save, checked against the signature of the format.
export function toFileBytes(body: any, kind: Kind): Uint8Array | undefined {
    let out: Uint8Array | undefined;
    if (body instanceof Uint8Array) {
        out = body;
    } else if (body instanceof ArrayBuffer) {
        out = new Uint8Array(body);
    } else if (ArrayBuffer.isView(body)) {
        out = new Uint8Array(body.buffer, body.byteOffset, body.byteLength);
    } else if (body && typeof body === 'object' && Number.isInteger(body.length) && body.length > 0) {
        out = Uint8Array.from(body as ArrayLike<number>);
    }
    if (!out || out.length < 5 || out.length > MAX_FILE_BYTES) {
        return undefined;
    }
    if (kind === 'docx' && !(out[0] === 0x50 && out[1] === 0x4b)) {
        return undefined;
    }
    if (kind === 'pdf' && String.fromCharCode(out[0], out[1], out[2], out[3], out[4]) !== '%PDF-') {
        return undefined;
    }
    if (kind === 'doc') {
        return undefined;
    }
    return out;
}

// One webview showing one document. The webview parses and renders the bytes itself;
// the host only delivers them, persists the view state and asks for the edited bytes on save.
export default class DocumentView {
    // Every open view; the integration tests read their status through a test-only command.
    static readonly all = new Set<DocumentView>();
    public status: any = null;
    private readonly _test: boolean;
    private readonly _panel: WebviewPanel;
    private readonly _storage: Memento;
    private readonly _extensionUri: URI;
    private readonly _disposables: Disposable[] = [];
    private readonly _pending = new Map<number, { resolve: (b: Uint8Array) => void; reject: (e: Error) => void }>();
    private _requestId = 0;
    private _disposed = false;

    constructor(
        context: ExtensionContext,
        private readonly _document: DxDocument,
        panel: WebviewPanel
    ) {
        this._panel = panel;
        this._storage = context.workspaceState;
        this._extensionUri = context.extensionUri;
        this._test = context.extensionMode === ExtensionMode.Test;
        DocumentView.all.add(this);
        panel.webview.options = {
            enableScripts: true,
            localResourceRoots: [Utils.joinPath(this._extensionUri, 'out'), Utils.joinPath(this._extensionUri, 'media')]
        };
        panel.onDidDispose(() => this.dispose(), this, this._disposables);
        panel.webview.onDidReceiveMessage(e => this.onMessage(e), this, this._disposables);
        panel.webview.html = this.getHtml();
    }

    get document(): DxDocument {
        return this._document;
    }

    get isDisposed(): boolean {
        return this._disposed;
    }

    private onMessage(e: any) {
        if (!e || typeof e !== 'object') {
            return;
        }
        if (e.loaded) {
            this.status = e.loaded;
        } else if (e.refresh) {
            this.refresh();
        } else if (e.save) {
            if (isSaneState(e.state)) {
                this._storage.update(this.stateKey, e.state);
            }
        } else if (e.changed) {
            this._document.change(typeof e.reason === 'string' ? e.reason.slice(0, 100) : 'Edit', this);
        } else if (e.response) {
            const p = this._pending.get(e.requestId);
            if (!p) {
                return;
            }
            this._pending.delete(e.requestId);
            if (e.error) {
                p.reject(new Error(String(e.error)));
            } else {
                const bytes = toFileBytes(e.body, this._document.kind);
                if (bytes) {
                    p.resolve(bytes);
                } else {
                    p.reject(new Error('The editor returned no valid document data; the file was not written.'));
                }
            }
        } else if (typeof e.error === 'string') {
            this.status = { error: e.error };
            if (!this._test) {
                window.showErrorMessage(e.error);
            }
        } else if (typeof e.openLink === 'string') {
            openExternalLink(e.openLink);
        }
    }

    // The custom editor shows the document model: that is what restores hot-exit backups and reverts.
    public refresh(): void {
        if (!this._disposed) {
            this._panel.webview.postMessage({
                refresh: true,
                content: this._document.documentData,
                editable: this._document.editable
            });
        }
    }

    public postTest(action: string) {
        this._panel.webview.postMessage({ type: 'test', action });
    }

    // Asks the webview for the current bytes; rejects on webview error, cancellation, close or timeout.
    public requestData(token?: CancellationToken): Promise<Uint8Array> {
        if (this._disposed) {
            return Promise.reject(new Error('The editor was closed.'));
        }
        const id = ++this._requestId;
        return new Promise<Uint8Array>((resolve, reject) => {
            let sub: Disposable | undefined;
            const done = () => {
                clearTimeout(timer);
                if (sub) {
                    sub.dispose();
                }
                this._pending.delete(id);
            };
            const timer = setTimeout(() => {
                done();
                reject(new Error('Timed out waiting for the editor to produce the file.'));
            }, DATA_TIMEOUT_MS);
            if (token) {
                sub = token.onCancellationRequested(() => {
                    done();
                    reject(new Error('Cancelled.'));
                });
            }
            this._pending.set(id, {
                resolve: b => {
                    done();
                    resolve(b);
                },
                reject: err => {
                    done();
                    reject(err);
                }
            });
            this._panel.webview.postMessage({ type: 'getData', requestId: id });
        });
    }

    public dispose() {
        if (this._disposed) {
            return;
        }
        this._disposed = true;
        DocumentView.all.delete(this);
        const pending = Array.from(this._pending.values());
        this._pending.clear();
        pending.forEach(p => p.reject(new Error('The editor was closed.')));
        while (this._disposables.length) {
            const item = this._disposables.pop();
            if (item) {
                item.dispose();
            }
        }
        this._document.detach(this);
    }

    private get stateKey(): string {
        return 'editor:dx.' + this._document.kind + ':' + this._document.uri.toString();
    }

    private getHtml(): string {
        const webview = this._panel.webview;
        const out = webview.asWebviewUri(Utils.joinPath(this._extensionUri, 'out'));
        const media = webview.asWebviewUri(Utils.joinPath(this._extensionUri, 'media'));
        const kind = this._document.kind;
        const nonce = makeNonce();
        const options = {
            kind,
            name: Utils.basename(this._document.uri),
            uri: this._document.uri.toString(),
            media: media.toString(),
            state: this._test ? null : this._storage.get(this.stateKey, null),
            test: this._test
        };
        const pdfCss =
            kind === 'pdf' ? `<link href="${media}/pdfjs/pdf_viewer.css" rel="stylesheet" type="text/css" />` : '';
        return `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="utf-8">
    <meta http-equiv="Content-Security-Policy" content="${csp(webview.cspSource, nonce)}">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    ${pdfCss}
    <link href="${media}/styles/vscode.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <div id="app"></div>
    <script id="dx-options" nonce="${nonce}" type="application/json">${safeJson(options)}</script>
    <script nonce="${nonce}" src="${out}/${kind === 'pdf' ? 'pdf' : 'word'}.js"></script>
</body>
</html>`;
    }
}
