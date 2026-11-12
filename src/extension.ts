'use strict';
import { commands, window, ExtensionContext, ExtensionMode, Uri } from 'vscode';
import { DxEditorProvider } from './editor';
import DocumentView from './view';

// Opens a file in the dx editor of its kind, also when another editor is the one VS Code uses for it.
function preview(what: string, viewType: (uri: Uri) => string) {
    return (uri: unknown) => {
        if (!(uri instanceof Uri)) {
            window.showInformationMessage(
                `Use the explorer context menu or editor title menu to preview ${what} files.`
            );
            return;
        }
        return commands.executeCommand('vscode.openWith', uri, viewType(uri));
    };
}

export function activate(context: ExtensionContext) {
    context.subscriptions.push(DxEditorProvider.register(context, 'pdf'));
    context.subscriptions.push(DxEditorProvider.register(context, 'docx'));
    context.subscriptions.push(DxEditorProvider.register(context, 'doc'));
    context.subscriptions.push(
        commands.registerCommand(
            'pdf.preview',
            preview('PDF', () => 'dx-pdf-editor')
        ),
        commands.registerCommand(
            'word.preview',
            preview('Word', uri => (/\.doc$/i.test(uri.path) ? 'dx-doc-editor' : 'dx-docx-editor'))
        )
    );

    if (context.extensionMode === ExtensionMode.Test) {
        // Integration test hooks: what every webview reported, and a scripted edit.
        context.subscriptions.push(
            commands.registerCommand('dx._test.views', () =>
                Array.from(DocumentView.all).map(v => ({ uri: v.document.uri.toString(), status: v.status }))
            ),
            commands.registerCommand('dx._test.edit', (uri: string) => {
                for (const view of DocumentView.all) {
                    if (view.document.uri.toString() === uri) {
                        view.postTest('edit');
                    }
                }
            })
        );
    }
}

export function deactivate() {}
