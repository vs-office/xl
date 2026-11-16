'use strict';
import type { WebviewPanel, ExtensionContext, ViewColumn, CancellationToken, Disposable } from 'vscode';
import { window, workspace, RelativePattern } from 'vscode';
import type { URI } from 'vscode-uri';
import { Utils } from 'vscode-uri';
import { safeJson, makeNonce, csp } from '../common/html';
import { errorMessage, isRecord } from '../common/guards';
import { toWorkbookBytes } from './workbookBytes';
import BaseDocumentView from '../common/baseDocumentView';
import type { ExcelDocument } from './excelEditor';

const DATA_TIMEOUT_MS = 60000;

export interface ExcelOptions {
    customEditor: boolean;
    uri: string;
    scriptUri: string;
    directUrl: string | undefined;
    previewUri: string;
    state: unknown;
}

export default class ExcelDocumentView extends BaseDocumentView {
    static create(context: ExtensionContext, uri: URI, viewColumn: ViewColumn): ExcelDocumentView {
        const preview = new ExcelDocumentView(context, uri);
        preview.scheme = 'excel-preview';
        preview.initWebviewPanel(viewColumn);
        preview.initialize();
        return preview;
    }

    static revive(context: ExtensionContext, uri: URI, webviewPanel: WebviewPanel): ExcelDocumentView {
        const preview = new ExcelDocumentView(context, uri);
        preview.scheme = 'excel-preview';
        preview.attachWebviewPanel(webviewPanel);
        preview.initialize();
        return preview;
    }

    public getOptions(): ExcelOptions {
        return {
            customEditor: this.hasCustomEditor,
            uri: this.uri.toString(),
            scriptUri: this.scriptUri.toString(),
            // a clean, non-restored local file is fetched by the webview straight from disk (no host round trip)
            directUrl: this._direct && this._document && !this._document.dirty ? this.previewUri.toString() : undefined,
            previewUri: this.previewUri.toString(),
            state: this.state
        };
    }

    private _document: ExcelDocument | undefined;
    private _requestId = 0;
    private readonly _pending = new Map<number, { resolve: (b: Uint8Array) => void; reject: (e: Error) => void }>();

    private _direct = false;

    public enableEditing(document: ExcelDocument) {
        this._document = document;
        if (document.uri.scheme === 'file' && !document.restored && !document.dirty) {
            this._direct = true;
            this.allowLocalRoots([Utils.dirname(document.uri)]);
        }
        this.webview.onDidReceiveMessage(
            (e: unknown) => {
                if (!isRecord(e)) {
                    return;
                }
                if (e.changed) {
                    document.change(typeof e.reason === 'string' ? e.reason.slice(0, 100) : 'Edit', this);
                } else if (e.response) {
                    const requestId = e.requestId;
                    const p = typeof requestId === 'number' ? this._pending.get(requestId) : undefined;
                    if (!p || typeof requestId !== 'number') {
                        return;
                    }
                    this._pending.delete(requestId);
                    if (e.error) {
                        p.reject(new Error(errorMessage(e.error)));
                    } else {
                        const bytes = toWorkbookBytes(e.body);
                        if (bytes) {
                            p.resolve(bytes);
                        } else {
                            p.reject(
                                new Error('The Excel editor returned no valid workbook data; the file was not written.')
                            );
                        }
                    }
                }
            },
            this,
            this._disposables
        );
    }

    // Asks the webview for the current workbook bytes; rejects on webview error, cancellation, close or timeout.
    public requestData(token?: CancellationToken): Promise<Uint8Array> {
        if (this._disposed) {
            return Promise.reject(new Error('The Excel editor was closed.'));
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
                reject(new Error('Timed out waiting for the Excel editor to produce the file.'));
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
                reject: e => {
                    done();
                    reject(e);
                }
            });
            void this.webview.postMessage({ type: 'getData', requestId: id });
        });
    }

    public override initialize() {
        super.initialize();
        if (!this.hasCustomEditor) {
            // read-only preview: follow the file on disk
            let timer: ReturnType<typeof setTimeout> | undefined;
            const watcher = workspace.createFileSystemWatcher(
                new RelativePattern(Utils.dirname(this.uri), Utils.basename(this.uri))
            );
            const reload = () => {
                clearTimeout(timer);
                timer = setTimeout(() => {
                    if (!this._disposed) {
                        this.refresh();
                    }
                }, 300);
            };
            this._disposables.push(watcher, watcher.onDidChange(reload), watcher.onDidCreate(reload), {
                dispose: () => clearTimeout(timer)
            });
        }
    }

    public override get hasUnsavedEdits(): boolean {
        return this._document?.dirty === true;
    }

    public override dispose() {
        const pending = Array.from(this._pending.values());
        this._pending.clear();
        super.dispose();
        pending.forEach(p => p.reject(new Error('The Excel editor was closed.')));
    }

    refresh(): void {
        // A custom editor shows the document model (this is what restores hot-exit backups); previews read the disk.
        if (this._document) {
            const doc = this._document;
            void doc.ready.then(
                () => {
                    if (!this._disposed) {
                        void this.webview.postMessage({ refresh: true, content: doc.documentData });
                    }
                },
                (reason: unknown) => {
                    void window.showErrorMessage(errorMessage(reason));
                }
            );
            return;
        }
        void workspace.fs.readFile(this.uri).then(
            buffer => {
                if (!this._disposed) {
                    void this.webview.postMessage({
                        refresh: true,
                        content: buffer
                    });
                }
            },
            (reason: unknown) => {
                void window.showInformationMessage(errorMessage(reason));
            }
        );
    }

    undo(): void {
        void this.webview.postMessage({
            undo: true
        });
    }

    redo(): void {
        void this.webview.postMessage({
            redo: true
        });
    }

    getHtml(ignoreState: boolean = false): string {
        const s = this.scriptUri;
        const nonce = makeNonce();
        return `
        <!DOCTYPE html>
        <html>
        <head>
        <meta http-equiv="Content-Security-Policy" content="${csp(this.webview.cspSource, nonce)}">
            <link href="${s.toString()}/style.css" rel="stylesheet" type="text/css" />
        </head>
        <body>
            <div id="app"></div>
            <script nonce="${nonce}" src="${s.toString()}/excel.js"></script>
            <script nonce="${nonce}" type="text/javascript">
                xlPage.start(${safeJson(this.getOptions())}, ${ignoreState === true});
            </script>
        </body>
        </html>`;
    }

    get viewType(): string {
        return 'xl-excel-preview';
    }

    get configurable(): boolean {
        return false;
    }
}
