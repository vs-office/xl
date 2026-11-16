import fs from 'node:fs';
import Module, { createRequire } from 'node:module';
import path from 'node:path';
import ts from 'typescript';
import type * as VscodeUri from 'vscode-uri';
import type { ExtensionContext, TextDocument, WebviewPanel } from 'vscode';
import type CsvDocumentView from '../../csv/csvDocumentView.ts';
import type { Shell, ShellOptions } from '../../webview/core/app-shell.ts';
import type { CsvProvider } from '../../webview/core/csv-provider.ts';
import type { GridContextMenuEvent, GridMenuItem, GridSelection } from '../../webview/core/grid.ts';
import type { CsvPageOptions, HostMessage, OxCsvApi, PageMessage } from '../../webview/csv.ts';

// The page schedules work through MessageChannel when it sees one; real ports would keep the test process alive,
// and the sandboxed page used to have none, so it falls back to timers here too.
// (Deleted before the page modules are imported below: import statements would be hoisted above it.)
delete (globalThis as { MessageChannel?: unknown }).MessageChannel;

const { parseCsv } = await import('../../webview/core/csv-parser.ts');
const csvPage = await import('../../webview/csv.ts');
export { parseCsv };

const R = path.resolve(import.meta.dirname, '../../..');
const require = createRequire(import.meta.url);
// vscode-uri is CommonJS: Node cannot see its named exports from an ES import, so load it like the host does.
export const { URI } = require('vscode-uri') as typeof VscodeUri;
type URI = VscodeUri.URI;

// ---- fake vscode
export class Position {
    line: number;
    character: number;
    constructor(l: number, c: number) {
        this.line = l;
        this.character = c;
    }
}
export class Range {
    start: Position;
    end: Position;
    constructor(a: Position, b: Position);
    constructor(a: number, b: number, c: number, d: number);
    constructor(a: Position | number, b: Position | number, c?: number, d?: number) {
        if (typeof a === 'number' && typeof b === 'number') {
            this.start = new Position(a, b);
            this.end = new Position(c ?? 0, d ?? 0);
        } else if (a instanceof Position && b instanceof Position) {
            this.start = a;
            this.end = b;
        } else {
            throw new TypeError('bad Range arguments');
        }
    }
}
export interface TextOp {
    range: Range;
    text: string;
}
export class WorkspaceEdit {
    ops: TextOp[] = [];
    replace(_uri: URI, range: Range, text: string): void {
        this.ops.push({ range, text });
    }
    insert(_uri: URI, pos: Position, text: string): void {
        this.ops.push({ range: new Range(pos, pos), text });
    }
}
export const EndOfLine = { LF: 1, CRLF: 2 };

export class FakeDoc {
    version = 1;
    uri: URI;
    text: string;
    eol: number;
    languageId = 'csv';
    constructor(uri: URI, text: string, eol?: number) {
        this.uri = uri;
        this.text = text;
        this.eol = eol || EndOfLine.LF;
    }
    get lines(): string[] {
        return this.text.split(/\r\n|\n|\r/);
    }
    get lineCount(): number {
        return this.lines.length;
    }
    getText(r?: Range): string {
        if (!r) return this.text;
        const a = this.offsetAt(r.start),
            b = this.offsetAt(r.end);
        return this.text.slice(a, b);
    }
    lineAt(i: number): { text: string; range: Range } {
        const ls = this.lines;
        if (i < 0 || i >= ls.length) throw new Error('Illegal value for line');
        const t = ls[i];
        return { text: t, range: new Range(i, 0, i, t.length) };
    }
    validatePos(p: Position): Position {
        const ls = this.lines;
        const l = p.line,
            c = p.character;
        if (l < 0) {
            return new Position(0, 0);
        }
        if (l >= ls.length) {
            return new Position(ls.length - 1, ls[ls.length - 1].length);
        }
        return new Position(l, Math.max(0, Math.min(c, ls[l].length)));
    }
    validateRange(r: Range): Range {
        return new Range(this.validatePos(r.start), this.validatePos(r.end));
    }
    offsetAt(p: Position): number {
        const re = /\r\n|\n|\r/g;
        let m: RegExpExecArray | null;
        const starts = [0];
        while ((m = re.exec(this.text))) starts.push(m.index + m[0].length);
        return starts[p.line] + p.character;
    }
    apply(ops: TextOp[]): void {
        // VS Code semantics: validate ranges, error on overlap, same-pos in order
        const o = ops.map((op, i) => {
            const r = this.validateRange(op.range);
            return { s: this.offsetAt(r.start), e: this.offsetAt(r.end), t: op.text, i };
        });
        o.sort((a, b) => a.s - b.s || a.i - b.i);
        for (let k = 1; k < o.length; k++)
            if (o[k].s < o[k - 1].e) throw new Error('Overlapping ranges are not allowed!');
        let out = '',
            pos = 0;
        for (const x of o) {
            out += this.text.slice(pos, x.s) + x.t;
            pos = x.e;
        }
        out += this.text.slice(pos);
        this.text = out;
        this.version++;
    }
}

interface DocChangeEvent {
    document: FakeDoc;
    contentChanges: object[];
}
type DocListener = (e: DocChangeEvent) => void;
/** The slice of the `vscode` module the host classes use, plus hooks the tests drive. */
export interface FakeVscode {
    Range: typeof Range;
    Position: typeof Position;
    WorkspaceEdit: typeof WorkspaceEdit;
    EndOfLine: typeof EndOfLine;
    ViewColumn: { One: number };
    window: {
        createWebviewPanel(): never;
        showErrorMessage(): void;
        showInformationMessage(): void;
        activeTextEditor: { document: { languageId: string } } | null;
    };
    workspace: {
        getConfiguration(): { get(k: string): unknown };
        textDocuments: { uri: URI; languageId: string }[];
        onDidChangeTextDocument(fn: DocListener, thisArg: object, disp?: unknown): { dispose(): void };
        applyEdit(we: WorkspaceEdit): Promise<boolean>;
        openTextDocument(): Promise<FakeDoc>;
    };
    /** applyEdit waits this many ms first */
    _delay?: number;
    /** runs inside applyEdit before the edit lands */
    _hook?: (() => void) | null;
    /** runs once inside applyEdit after the edit landed (one-shot) */
    _after?: (() => void) | null;
}

let currentDoc = new FakeDoc(URI.parse('file:///t.csv'), '');
let docListeners: DocListener[] = [];
let cfg: Record<string, string | boolean> = {};
export const vscode: FakeVscode = {
    Range,
    Position,
    WorkspaceEdit,
    EndOfLine,
    ViewColumn: { One: 1 },
    window: {
        createWebviewPanel(): never {
            throw new Error('createWebviewPanel is not available in the simulation');
        },
        showErrorMessage() {},
        showInformationMessage() {},
        activeTextEditor: null
    },
    workspace: {
        getConfiguration: () => ({ get: k => cfg[k] }),
        textDocuments: [],
        onDidChangeTextDocument(fn, thisArg) {
            docListeners.push(fn.bind(thisArg));
            return { dispose() {} };
        },
        async applyEdit(we) {
            if (vscode._delay) await new Promise(r => setTimeout(r, vscode._delay));
            if (vscode._hook) vscode._hook();
            currentDoc.apply(we.ops);
            for (const f of docListeners) f({ document: currentDoc, contentChanges: [{}] });
            if (vscode._after) {
                const a = vscode._after;
                vscode._after = null;
                a();
            }
            return true;
        },
        openTextDocument() {
            return Promise.resolve(currentDoc);
        }
    }
};

// The host classes are extensionless TypeScript with `import ... from 'vscode'`: make CommonJS `require` load them
// (transpiled by the TypeScript API) and answer 'vscode' with the fake above.
interface CompilableModule {
    _compile(code: string, filename: string): void;
}
interface ModuleInternals {
    _load: (request: string, parent: unknown, isMain: boolean) => unknown;
    _extensions: Record<string, (m: CompilableModule, filename: string) => void>;
}
// Module's private loader hooks are not part of the public typings.
const internals = Module as unknown as ModuleInternals;
const origLoad = internals._load;
internals._load = function (this: unknown, req: string, parent: unknown, isMain: boolean): unknown {
    if (req === 'vscode') return vscode;
    return origLoad.call(this, req, parent, isMain);
};
internals._extensions['.ts'] = (m, f) => {
    m._compile(
        ts.transpileModule(fs.readFileSync(f, 'utf8'), {
            compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2019 }
        }).outputText,
        f
    );
};
const hostModule = require(R + '/src/csv/csvDocumentView.ts') as { default: typeof CsvDocumentView };
const CsvDocumentViewCtor = hostModule.default;

/** Loads another host source file (path relative to the repo root) through the same hooks as the view. */
export const loadHost = (rel: string): unknown => require(R + rel);

// ---- fake page shell
export interface FakeSelection {
    ranges: { r1: number; r2: number; c1: number; c2: number }[];
}
export class FakeGrid {
    /** contextmenu listener registered by the page */
    contextmenu: ((e: GridContextMenuEvent) => void) | undefined;
    /** replaced by tests to capture the menu */
    _showMenu: (items: readonly GridMenuItem[]) => void = () => {};
    getSelection: () => FakeSelection | GridSelection = () => ({ ranges: [] });
    on(n: string, f: (e: GridContextMenuEvent) => void): void {
        if (n === 'contextmenu') this.contextmenu = f;
    }
    invalidate(): void {}
    refresh(): void {}
    select(): void {}
    scrollTo(): void {}
}
export class FakeShell {
    /** the options the page passed (onEdit is always provided for an editable page) */
    o: ShellOptions & { onEdit: NonNullable<ShellOptions['onEdit']> };
    grid = new FakeGrid();
    bar = { insertBefore() {} };
    btnClear = {};
    provider: CsvProvider | null = null;
    loading = false;
    constructor(o: ShellOptions) {
        if (!o.onEdit) throw new Error('page created a read-only shell');
        this.o = { ...o, onEdit: o.onEdit };
    }
    setProvider(p: CsvProvider): void {
        this.provider = p;
    }
    applyViewState(): void {}
    getViewState(): null {
        return null;
    }
    setMessage(): void {}
    _updateAll(): void {}
    isDark(): boolean {
        return false;
    }
    clearView(): void {}
}

/** The fake `window` the page sees. */
export interface FakeWindow {
    addEventListener(n: string, f: (e: { data: HostMessage }) => void): void;
    oxCsv?: OxCsvApi;
}
export interface WebviewHandle {
    view: CsvDocumentView;
    win: FakeWindow;
    shell(): FakeShell;
    doc: FakeDoc;
    provider(): OxCsvApi | undefined;
}

/** Narrows a possibly missing value; throws when it is absent. */
export function defined<T>(v: T | null | undefined, what = 'value'): T {
    if (v === null || v === undefined) throw new Error(what + ' is not available');
    return v;
}

// ---- webview side
export function makeWebview(text: string, eol?: number, optsOver?: Record<string, string | boolean>): WebviewHandle {
    currentDoc = new FakeDoc(URI.parse('file:///t.csv'), text, eol);
    docListeners = [];
    cfg = {
        separator: ',',
        quoteMark: '"',
        hasHeaders: true,
        capitalizeHeaders: false,
        resizeColumns: 'none',
        lineNumbers: 'none',
        commentCharacter: '#',
        skipComments: false,
        formatValues: 'never',
        numberFormat: '',
        ...(optsOver || {})
    };
    const hostHandlers: ((m: unknown) => void)[] = [];
    const webview = {
        cspSource: 'csp',
        asWebviewUri: (u: URI) => u,
        onDidReceiveMessage(fn: (m: unknown) => void, t: object) {
            hostHandlers.push(fn.bind(t));
            return { dispose() {} };
        },
        postMessage(m: HostMessage) {
            setImmediate(() => deliver(m));
            return Promise.resolve(true);
        },
        set html(_h: string) {},
        get html() {
            return '';
        }
    };
    const panel = { webview, onDidDispose() {}, visible: true };
    // VS Code's ExtensionContext / WebviewPanel are huge; the host only touches the members faked here.
    const context = {
        extensionUri: URI.parse('file:///ext'),
        workspaceState: { get: () => null, update() {} }
    } as unknown as ExtensionContext;
    const view = new CsvDocumentViewCtor(context, currentDoc.uri);
    view.languageId = 'csv';
    view.panel = panel as unknown as WebviewPanel;
    view.initialize();
    view.enableEditing(currentDoc as unknown as TextDocument);
    // the page runs in this process against fake browser globals; each call gets a fresh window
    const msgL: ((e: { data: HostMessage }) => void)[] = [];
    const win: FakeWindow = {
        addEventListener(n, f) {
            if (n === 'message') msgL.push(f);
        }
    };
    // The page only needs these few members of the browser globals.
    const g = globalThis as unknown as Record<string, unknown>;
    g.window = win;
    g.acquireVsCodeApi = () => ({
        postMessage(m: PageMessage) {
            for (const h of hostHandlers) h(m);
        },
        getState: () => null,
        setState() {}
    });
    g.document = {
        getElementById: () => ({}),
        createElement: () => ({ addEventListener() {}, style: {} })
    };
    function deliver(m: HostMessage): void {
        for (const f of msgL) f({ data: m });
    }
    let created: FakeShell | undefined;
    csvPage.handleEvents();
    // the page reads the options through the same JSON round trip the real HTML embeds
    const pageOptions: CsvPageOptions = view.getOptions();
    csvPage.initPage(pageOptions, false, (_app, o) => {
        created = new FakeShell(o);
        // the page only uses the members FakeShell provides
        return created as unknown as Shell;
    });
    return {
        view,
        win,
        shell: () => defined(created, 'shell'),
        doc: currentDoc,
        provider: () => win.oxCsv
    };
}

export const fire = (): void => {
    for (const f of docListeners) f({ document: currentDoc, contentChanges: [{}] });
};
export const externalEdit = (text: string): void => {
    currentDoc.text = text;
    currentDoc.version++;
    for (const f of docListeners) f({ document: currentDoc, contentChanges: [{}] });
};
