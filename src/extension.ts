'use strict';
import type { ExtensionContext, TextDocument } from 'vscode';
import { window, workspace, commands, Uri, ViewColumn } from 'vscode';
import CsvDocumentView from './csv/csvDocumentView';
import ExcelDocumentView from './excel/excelDocumentView';
import { documentViewManager } from './common/documentViewManager';
import { Utils } from 'vscode-uri';
import { CsvSerializer, ExcelSerializer } from './serializer';
import { CsvEditorProvider } from './csv/csvEditor';
import { ExcelEditorProvider } from './excel/excelEditor';

export function activate(context: ExtensionContext) {
    // CSV: Open Preview
    const csvCommand = commands.registerCommand('csv.preview', (uri: unknown) => {
        let resource: Uri;
        let viewColumn: ViewColumn = getViewColumn();
        if (uri instanceof Uri) {
            resource = uri;
        } else if (window.activeTextEditor) {
            resource = window.activeTextEditor.document.uri;
            viewColumn = window.activeTextEditor.viewColumn ?? ViewColumn.One;
        } else {
            void window.showInformationMessage('Open a CSV file first to show a preview.');
            return;
        }
        let preview = documentViewManager.find(resource, false);
        if (preview) {
            preview.reveal();
            return preview.webview;
        }
        preview = CsvDocumentView.create(context, resource, viewColumn);
        return preview.webview;
    });

    // Excel: Open Preview
    const excelCommand = commands.registerCommand('excel.preview', (uri: unknown) => {
        const resource = uri;
        const viewColumn = getViewColumn();
        if (!(resource instanceof Uri)) {
            void window.showInformationMessage(
                'Use the explorer context menu or editor title menu to preview Excel files.'
            );
            return;
        }
        let preview = documentViewManager.find(resource, false);
        if (preview) {
            preview.reveal();
            return preview.webview;
        }
        preview = ExcelDocumentView.create(context, resource, viewColumn);
        return preview.webview;
    });

    // CSV: Clear Preview State
    const clearCommand = commands.registerCommand('csv.clearState', () => {
        const preview = documentViewManager.active();
        if (preview) {
            preview.clearState();
        }
    });

    // CSV: Refresh
    const refreshCommand = commands.registerCommand('csv.refresh', () => {
        const preview = documentViewManager.active();
        if (preview) {
            preview.userRefresh();
        }
    });

    // Add disposables to subscriptions array
    context.subscriptions.push(csvCommand);
    context.subscriptions.push(excelCommand);
    context.subscriptions.push(clearCommand);
    context.subscriptions.push(refreshCommand);
    context.subscriptions.push(CsvEditorProvider.register(context));
    context.subscriptions.push(ExcelEditorProvider.register(context));

    // Register serializers for persistent settings
    context.subscriptions.push(window.registerWebviewPanelSerializer('xl-csv-preview', new CsvSerializer(context)));
    context.subscriptions.push(window.registerWebviewPanelSerializer('xl-excel-preview', new ExcelSerializer(context)));

    // Reset all previews when the configuration changes
    context.subscriptions.push(
        workspace.onDidChangeConfiguration(e => {
            if (e.affectsConfiguration('csv-preview')) {
                documentViewManager.configure();
            }
        })
    );

    // Automatically preview content piped from stdin (when VSCode is already open)
    context.subscriptions.push(
        workspace.onDidOpenTextDocument(document => {
            if (isStdinFile(document)) {
                void commands.executeCommand('csv.preview', document.uri);
            }
        })
    );

    // Automatically preview content piped from stdin (when VSCode first starts up)
    if (window.activeTextEditor) {
        const document = window.activeTextEditor.document;
        if (isStdinFile(document)) {
            void commands.executeCommand('csv.preview', document.uri);
        }
    }
}

export function deactivate() {}

function isStdinFile(document: TextDocument): boolean {
    const allowed = workspace.getConfiguration('csv-preview').get<boolean>('openStdin', false);
    return allowed && new RegExp('code-stdin-[^.]+.txt').test(Utils.basename(document.uri));
}

function getViewColumn(): ViewColumn {
    const active = window.activeTextEditor;
    return active?.viewColumn ?? ViewColumn.One;
}
