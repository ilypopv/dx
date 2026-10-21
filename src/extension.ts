'use strict';
import { commands, ExtensionContext, ExtensionMode } from 'vscode';
import { DxEditorProvider } from './editor';
import DocumentView from './view';

export function activate(context: ExtensionContext) {
    context.subscriptions.push(DxEditorProvider.register(context, 'pdf'));
    context.subscriptions.push(DxEditorProvider.register(context, 'docx'));
    context.subscriptions.push(DxEditorProvider.register(context, 'doc'));

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
