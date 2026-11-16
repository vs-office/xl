import * as vscode from 'vscode';
import CsvDocumentView from './csvDocumentView';

export class CsvEditorProvider implements vscode.CustomTextEditorProvider {
    public static register(context: vscode.ExtensionContext): vscode.Disposable {
        const provider = new CsvEditorProvider(context);
        const providerRegistration = vscode.window.registerCustomEditorProvider(CsvEditorProvider.viewType, provider, {
            webviewOptions: {
                enableFindWidget: false,
                retainContextWhenHidden: true
            }
        });
        return providerRegistration;
    }

    private static readonly viewType = 'xl-csv-editor';

    private readonly context: vscode.ExtensionContext;

    constructor(context: vscode.ExtensionContext) {
        this.context = context;
    }

    public resolveCustomTextEditor(document: vscode.TextDocument, webviewPanel: vscode.WebviewPanel): void {
        const editor = new CsvDocumentView(this.context, document.uri);
        editor.languageId = document.languageId;
        editor.panel = webviewPanel;
        editor.initialize();
        editor.enableEditing(document);
    }
}
