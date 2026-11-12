import * as vscode from 'vscode';
import { Utils } from 'vscode-uri';
import { Disposable, disposeAll } from './dispose';
import DocumentView from './view';

export type Kind = 'pdf' | 'docx' | 'doc';

function sameBytes(a: Uint8Array, b: Uint8Array): boolean {
    if (a.length !== b.length) {
        return false;
    }
    for (let i = 0; i < a.length; i++) {
        if (a[i] !== b[i]) {
            return false;
        }
    }
    return true;
}

async function readFile(uri: vscode.Uri): Promise<Uint8Array> {
    if (uri.scheme === 'untitled') {
        return new Uint8Array();
    }
    return vscode.workspace.fs.readFile(uri);
}

// The document behind one or more dx editors. It holds the last known bytes (disk, backup or the latest save)
// and, while edited, asks the webview that made the edits for the new bytes on save and backup.
export class DxDocument extends Disposable implements vscode.CustomDocument {
    // The file is read while the webview boots; consumers of documentData wait for `ready`.
    static create(uri: vscode.Uri, kind: Kind, backupId: string | undefined): DxDocument {
        const dataFile = typeof backupId === 'string' ? vscode.Uri.parse(backupId) : uri;
        return new DxDocument(uri, kind, readFile(dataFile), typeof backupId === 'string');
    }

    private _documentData: Uint8Array = new Uint8Array();
    private _ready: Promise<void>;
    private readonly _views = new Set<DocumentView>();
    private _lastView: DocumentView | undefined;
    private _dirty: boolean;
    private _watchTimer: any;
    private _warned = false;

    private readonly _onDidDispose = this._register(new vscode.EventEmitter<void>());
    public readonly onDidDispose = this._onDidDispose.event;

    private readonly _onDidChange = this._register(new vscode.EventEmitter<void>());
    public readonly onDidChange = this._onDidChange.event;

    private constructor(
        public readonly uri: vscode.Uri,
        public readonly kind: Kind,
        initialContent: Promise<Uint8Array>,
        // true when the model was restored from a hot-exit backup, so the disk file is not the content to show
        public readonly restored: boolean
    ) {
        super();
        this._dirty = restored;
        const ready: Promise<void> = initialContent.then(data => {
            // A revert before the first read finished has already replaced the model.
            if (this._ready === ready) {
                this._documentData = data;
            }
        });
        this._ready = ready;
        ready.catch(() => undefined);
        this.watch();
    }

    get ready(): Promise<void> {
        return this._ready;
    }

    get documentData(): Uint8Array {
        return this._documentData;
    }

    // .doc is shown read-only: dx cannot write the binary Word format.
    get editable(): boolean {
        return this.kind !== 'doc';
    }

    get dirty(): boolean {
        return this._dirty;
    }

    attach(view: DocumentView) {
        this._views.add(view);
    }

    detach(view: DocumentView) {
        this._views.delete(view);
        if (this._lastView === view) {
            this._lastView = undefined;
        }
    }

    private watch() {
        if (this.uri.scheme === 'untitled') {
            return;
        }
        try {
            const watcher = vscode.workspace.createFileSystemWatcher(
                new vscode.RelativePattern(Utils.dirname(this.uri) as any, Utils.basename(this.uri))
            );
            const onDisk = () => {
                clearTimeout(this._watchTimer);
                this._watchTimer = setTimeout(() => this.diskChanged().catch(() => undefined), 300);
            };
            this._register(watcher);
            this._register(watcher.onDidChange(onDisk));
            this._register(watcher.onDidCreate(onDisk));
            this._register({ dispose: () => clearTimeout(this._watchTimer) });
        } catch {
            // No watcher support (virtual file systems).
        }
    }

    // Reloads a clean document when the file changes on disk; a dirty one only warns, it never loses edits.
    private async diskChanged() {
        await this._ready;
        const disk = await readFile(this.uri);
        if (sameBytes(disk, this._documentData)) {
            return;
        }
        if (!this._dirty) {
            this._documentData = disk;
            this.refreshViews();
            return;
        }
        if (this._warned) {
            return;
        }
        this._warned = true;
        const pick = await vscode.window.showWarningMessage(
            `'${Utils.basename(this.uri)}' changed on disk while it has unsaved edits. Saving will overwrite the disk version.`,
            'Revert to disk version'
        );
        this._warned = false;
        if (pick) {
            await vscode.commands.executeCommand('workbench.action.files.revert');
        }
    }

    private refreshViews() {
        for (const view of this._views) {
            view.refresh();
        }
    }

    change(_label: string, view: DocumentView) {
        if (!this.editable) {
            return;
        }
        this._dirty = true;
        this._lastView = view;
        this._onDidChange.fire();
    }

    private async currentData(token: vscode.CancellationToken): Promise<Uint8Array> {
        await this._ready;
        const view =
            this._lastView && !this._lastView.isDisposed
                ? this._lastView
                : Array.from(this._views).find(v => !v.isDisposed);
        if (!view) {
            // No webview is open (a restored backup that was never shown): the bytes already are the edited state.
            return this._documentData;
        }
        return view.requestData(token);
    }

    async save(cancellation: vscode.CancellationToken): Promise<void> {
        await this.saveAs(this.uri, cancellation);
    }

    async saveAs(target: vscode.Uri, cancellation: vscode.CancellationToken, backup = false): Promise<void> {
        // Rejects when the webview cannot produce the bytes, so VS Code reports the failed save and keeps the document dirty.
        const data = await this.currentData(cancellation);
        if (cancellation.isCancellationRequested) {
            return;
        }
        if (!data || data.length < 5) {
            throw new Error('Refusing to write an empty document.');
        }
        await vscode.workspace.fs.writeFile(target, data);
        if (backup) {
            return;
        }
        if (target.toString() === this.uri.toString()) {
            this._dirty = false;
        }
        // Remember the latest bytes: a re-created webview must not fall back to older data.
        this._documentData = data;
        // The other editors of the same file show the saved bytes; the one that made the edits already does.
        for (const view of this._views) {
            if (view !== this._lastView) {
                view.refresh();
            }
        }
    }

    async revert(_cancellation: vscode.CancellationToken): Promise<void> {
        this._documentData = await readFile(this.uri);
        this._ready = Promise.resolve();
        this._dirty = false;
        this.refreshViews();
    }

    async backup(
        destination: vscode.Uri,
        cancellation: vscode.CancellationToken
    ): Promise<vscode.CustomDocumentBackup> {
        await this.saveAs(destination, cancellation, true);
        return {
            id: destination.toString(),
            delete: async () => {
                try {
                    await vscode.workspace.fs.delete(destination);
                } catch {
                    // Already gone.
                }
            }
        };
    }

    dispose(): void {
        this._onDidDispose.fire();
        super.dispose();
    }
}

const VIEW_TYPES: { [kind in Kind]: string } = {
    pdf: 'dx-pdf-editor',
    docx: 'dx-docx-editor',
    doc: 'dx-doc-editor'
};

export class DxEditorProvider implements vscode.CustomEditorProvider<DxDocument> {
    public static register(context: vscode.ExtensionContext, kind: Kind): vscode.Disposable {
        return vscode.window.registerCustomEditorProvider(VIEW_TYPES[kind], new DxEditorProvider(context, kind), {
            webviewOptions: {
                enableFindWidget: false,
                retainContextWhenHidden: true
            },
            supportsMultipleEditorsPerDocument: true
        });
    }

    private readonly _onDidChangeCustomDocument = new vscode.EventEmitter<
        vscode.CustomDocumentContentChangeEvent<DxDocument>
    >();
    public readonly onDidChangeCustomDocument = this._onDidChangeCustomDocument.event;

    constructor(
        private readonly context: vscode.ExtensionContext,
        private readonly kind: Kind
    ) {}

    openCustomDocument(
        uri: vscode.Uri,
        openContext: vscode.CustomDocumentOpenContext,
        _token: vscode.CancellationToken
    ): DxDocument {
        const document = DxDocument.create(uri, this.kind, openContext.backupId);
        const listeners: vscode.Disposable[] = [];
        // A content change (not an edit with undo/redo): the webview keeps its own undo history.
        listeners.push(document.onDidChange(() => this._onDidChangeCustomDocument.fire({ document })));
        document.onDidDispose(() => disposeAll(listeners));
        return document;
    }

    async resolveCustomEditor(
        document: DxDocument,
        webviewPanel: vscode.WebviewPanel,
        _token: vscode.CancellationToken
    ): Promise<void> {
        document.attach(new DocumentView(this.context, document, webviewPanel));
    }

    saveCustomDocument(document: DxDocument, cancellation: vscode.CancellationToken): Thenable<void> {
        return document.save(cancellation);
    }

    saveCustomDocumentAs(
        document: DxDocument,
        destination: vscode.Uri,
        cancellation: vscode.CancellationToken
    ): Thenable<void> {
        return document.saveAs(destination, cancellation);
    }

    revertCustomDocument(document: DxDocument, cancellation: vscode.CancellationToken): Thenable<void> {
        return document.revert(cancellation);
    }

    backupCustomDocument(
        document: DxDocument,
        context: vscode.CustomDocumentBackupContext,
        cancellation: vscode.CancellationToken
    ): Thenable<vscode.CustomDocumentBackup> {
        return document.backup(context.destination, cancellation);
    }
}
