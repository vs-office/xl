'use strict';
import type { WebviewPanel, WebviewPanelSerializer, ExtensionContext } from 'vscode';
import { Uri, workspace } from 'vscode';
import type BaseDocumentView from './common/baseDocumentView';
import { isRecord } from './common/guards';
import CsvDocumentView from './csv/csvDocumentView';
import ExcelDocumentView from './excel/excelDocumentView';
import { documentViewManager } from './common/documentViewManager';

type Revive = (context: ExtensionContext, uri: Uri, panel: WebviewPanel) => BaseDocumentView;

async function restore(context: ExtensionContext, panel: WebviewPanel, state: unknown, revive: Revive) {
    let uri: Uri | undefined;
    try {
        uri = isRecord(state) && typeof state.uri === 'string' ? Uri.parse(state.uri, true) : undefined;
    } catch {
        uri = undefined;
    }
    // Without a persisted uri there is nothing to show; drop the dead panel instead of opening file:///undefined.
    if (!uri || documentViewManager.find(uri, false)) {
        panel.dispose();
        return;
    }
    try {
        await workspace.fs.stat(uri);
    } catch {
        panel.dispose();
        return;
    }
    revive(context, uri, panel);
}

export class CsvSerializer implements WebviewPanelSerializer {
    constructor(private readonly context: ExtensionContext) {}

    public async deserializeWebviewPanel(webviewPanel: WebviewPanel, state: unknown) {
        await restore(this.context, webviewPanel, state, (c, u, p) => CsvDocumentView.revive(c, u, p));
    }
}

export class ExcelSerializer implements WebviewPanelSerializer {
    constructor(private readonly context: ExtensionContext) {}

    public async deserializeWebviewPanel(webviewPanel: WebviewPanel, state: unknown) {
        await restore(this.context, webviewPanel, state, (c, u, p) => ExcelDocumentView.revive(c, u, p));
    }
}
