import { Shell } from './core/app-shell.ts';
import type { ShellOptions, ShellViewState, ShellViewStateInput } from './core/app-shell.ts';
import { CsvIndex } from './core/csv-parser.ts';
import type { CsvParseOptions } from './core/csv-parser.ts';
import { CsvProvider } from './core/csv-provider.ts';
import type { FormatValuesMode, LineNumbersMode } from './core/csv-provider.ts';
import type { GridContextMenuEvent, GridEdit } from './core/grid.ts';

/** Page options the extension host passes to `start` (webview-side counterpart of the host's CsvOptions). */
export interface CsvPageOptions {
    separator: string;
    languageId: string | undefined;
    quoteMark: string;
    hasHeaders: boolean;
    capitalizeHeaders: boolean;
    resizeColumns: string;
    lineNumbers: LineNumbersMode;
    commentCharacter: string;
    skipComments: boolean;
    formatValues: FormatValuesMode;
    numberFormat: string;
    customEditor: boolean;
    uri: string;
    previewUri: string;
    state: unknown;
}

/** Host -> page: the whole document as one string (small files). */
export interface RefreshMessage {
    refresh: true;
    content?: string;
    version?: number;
}
/** Host -> page: a chunked transfer begins; a newer start cancels the old transfer. */
export interface ChunkStartMessage {
    refreshChunk: 'start';
    id?: number;
    version?: number;
    chars: number;
    lines?: number;
}
export interface ChunkDataMessage {
    refreshChunk: 'data';
    id?: number;
    version?: number;
    seq: number;
    text: string;
}
export interface ChunkEndMessage {
    refreshChunk: 'end';
    id?: number;
    version?: number;
    count: number;
}
export type ChunkMessage = ChunkStartMessage | ChunkDataMessage | ChunkEndMessage;
export type HostMessage = RefreshMessage | ChunkMessage;

/** One document edit in line coordinates (start/end are line numbers, end exclusive). */
export interface LineEdit {
    start: number;
    end: number;
    text: string;
}
/** View state persisted by the page (webview state and host `save`). */
export interface PageState {
    version: string;
    uri: string;
    view: ShellViewState;
}
/** Page -> host messages. */
export interface RequestRefreshMessage {
    refresh: true;
}
export interface SaveStateMessage {
    save: true;
    state: PageState;
}
export interface SeparatorMessage {
    separator: string;
}
export interface LineEditsMessage {
    lineEdits: LineEdit[];
    version: number | null;
}
export type PageMessage = RequestRefreshMessage | SaveStateMessage | SeparatorMessage | LineEditsMessage;

/** Test seam: builds the Shell for the page. */
export type CreateShell = (app: HTMLElement, o: ShellOptions) => Shell;

/** Programmatic API of the page, exposed as `window.oxCsv`. */
export interface OxCsvApi {
    load(text: string | null | undefined, version?: number): void;
    chunk(m: ChunkMessage): void;
    shell: Shell;
}
declare global {
    interface Window {
        oxCsv?: OxCsvApi;
    }
}

/** A chunked transfer / load in progress. */
export interface LoadJob {
    id: number | null | undefined;
    seq: number;
    cancelled: boolean;
    final: boolean;
    pumping: boolean;
    swapped: boolean;
    sync: boolean;
    buf: Uint8Array;
    filled: number;
    expected: number;
    ix: CsvIndex;
    t0: number;
    lastPos: number;
    lastInv: number;
    prev: ShellViewState | null;
    prevProvider: CsvProvider | null;
    state0: string | null;
    version?: number | undefined;
    reapply?: boolean;
    vs?: ShellViewStateInput | null;
}

function isRecord(v: unknown): v is Record<string, unknown> {
    return typeof v === 'object' && v !== null;
}

function isOptionalNumber(v: unknown): v is number | undefined {
    return v === undefined || typeof v === 'number';
}

/** Narrows an untyped window message to a host message; anything else is ignored. */
export function isHostMessage(d: unknown): d is HostMessage {
    if (!isRecord(d)) return false;
    if (d.refreshChunk) {
        if (!isOptionalNumber(d.id) || !isOptionalNumber(d.version)) return false;
        switch (d.refreshChunk) {
            case 'start':
                return typeof d.chars === 'number';
            case 'data':
                return typeof d.seq === 'number' && typeof d.text === 'string';
            case 'end':
                return typeof d.count === 'number';
            default:
                return false;
        }
    }
    return !!d.refresh && (d.content === undefined || typeof d.content === 'string') && isOptionalNumber(d.version);
}

function isChunkMessage(m: HostMessage): m is ChunkMessage {
    return 'refreshChunk' in m;
}

function isViewStateInput(v: unknown): v is ShellViewStateInput {
    if (!isRecord(v)) return false;
    const { zoom, scroll, active } = v;
    if (zoom !== undefined && typeof zoom !== 'number') return false;
    if (scroll !== undefined && scroll !== null) {
        if (!isRecord(scroll)) return false;
        for (const k of [scroll.top, scroll.left])
            if (k !== undefined && k !== null && typeof k !== 'number') return false;
    }
    if (active !== undefined && active !== null) {
        if (!isRecord(active) || typeof active.r !== 'number' || typeof active.c !== 'number') return false;
    }
    return true;
}

interface SavedState {
    version: string;
    view?: unknown;
}

function isSavedState(v: unknown, version: string): v is SavedState {
    return isRecord(v) && v.version === version;
}

let measureCtx: CanvasRenderingContext2D | null = null;

// createShell is a seam for tests, which drive the page without a DOM
export function initPage(
    options: CsvPageOptions,
    ignoreState?: boolean,
    createShell: CreateShell = (app, o) => new Shell(app, o)
): void {
    const vscode = acquireVsCodeApi();
    const STATE_VERSION = '6';
    const readOnly = !options.customEditor;
    const sendMessage = function (m: PageMessage): void {
        vscode.postMessage(m);
    };

    const app = document.getElementById('app');
    if (!app) throw new Error('Missing #app element');
    let provider: CsvProvider | null = null,
        sepSent = false;
    // docVersion: document version the displayed rows were computed from (echoed with edits); transferring: a chunked refresh is in flight
    let docVersion: number | null = null,
        transferring = false;

    const saved = (function (): SavedState | null {
        if (ignoreState) return null;
        const s = vscode.getState() || options.state;
        return isSavedState(s, STATE_VERSION) ? s : null;
    })();

    // the provider exists whenever rows are displayed; callers reach this only after makeProvider
    function need(): CsvProvider {
        if (!provider) throw new Error('No CSV provider yet');
        return provider;
    }

    function measure(text: string): number {
        let c = measureCtx;
        if (!c) {
            c = document.createElement('canvas').getContext('2d');
            if (!c) throw new Error('2D canvas context unavailable');
            measureCtx = c;
            c.font = '13px system-ui';
        }
        return c.measureText(text).width;
    }

    const shell: Shell = createShell(app, {
        readOnly: readOnly,
        tabs: false,
        headerRows: function () {
            return provider ? provider.frozenRows : 0;
        },
        onEdit: applyEdits,
        onViewState: function (vs) {
            const state: PageState = { version: STATE_VERSION, uri: options.uri, view: vs };
            vscode.setState(state);
            sendMessage({ save: true, state: state });
        },
        onLink: function () {}
    });

    if (!readOnly) {
        const add = document.createElement('button');
        add.className = 'ox-btn ox-add';
        add.textContent = 'Add row';
        add.title = 'Append a row at the end';
        add.addEventListener('click', addRow);
        shell.bar.insertBefore(add, shell.btnClear);
        shell.grid.on('contextmenu', onRowMenu);
    }

    function sourceRows(view: { ranges: { r1: number; r2: number }[] }): number[] {
        const order = shell._order,
            rows = new Set<number>();
        for (const rg of view.ranges) {
            const r2 = Math.min(rg.r2, (order ? order.length : need().rowCount) - 1);
            for (let r = rg.r1; r <= r2; r++) rows.add(order ? order[r] : r);
        }
        return Array.from(rows);
    }

    function onRowMenu(e: GridContextMenuEvent): void {
        const p = provider;
        if (e.zone !== 'rowhdr' || !p) return;
        e.handled = true;
        const sel = shell.grid.getSelection();
        const rows = sourceRows(sel).filter(function (r) {
            return !(p.hasHeaders && r === 0);
        });
        const g = shell.grid;
        g._showMenu(
            [
                [
                    'Copy',
                    '',
                    function () {
                        g.copy(false);
                    }
                ],
                null,
                [
                    'Delete ' + (rows.length > 1 ? rows.length + ' rows' : 'row'),
                    '',
                    function () {
                        deleteRows(rows);
                    },
                    !rows.length
                ],
                ['Add row at end', '', addRow]
            ],
            e.clientX,
            e.clientY
        );
    }

    function post(edits: LineEdit[]): void {
        if (!edits.length) return;
        // an empty or comment-only file has no header line (-1): clamp so the host never sees negative lines
        const clean = edits.map(function (e) {
            const s = Math.max(0, e.start | 0);
            return { start: s, end: Math.max(s, e.end | 0), text: e.text };
        });
        sendMessage({ lineEdits: clean, version: docVersion });
        // the host applies a batch as exactly one document version
        if (docVersion !== null) docVersion++;
    }

    function busy(): boolean {
        if (!shell.loading && !transferring && docVersion !== null) return false;
        shell.setMessage('Still loading: editing is available when the file is fully loaded');
        shell.grid.refresh();
        return true;
    }

    function applyEdits(edits: GridEdit[]): void {
        if (readOnly || !provider || busy()) return;
        const p = provider,
            changed = new Map<number, ReturnType<CsvProvider['getSourceRange']>>(),
            appended: { r: number; range: ReturnType<CsvProvider['getSourceRange']> }[] = [];
        let maxR = -1;
        edits.forEach(function (e) {
            if (e.r > maxR) maxR = e.r;
        });
        const startCount = p.rowCount;
        const viewActive = !!shell._order;
        while (p.rowCount <= maxR) {
            const r = p.appendRow();
            appended.push({ r: r, range: p.getSourceRange(r) });
        }
        const ordered = edits.filter(function (e) {
            return e.r >= startCount || p.getRaw(e.r, e.c) !== e.text;
        });
        const lineEdits: LineEdit[] = [];
        ordered.forEach(function (e) {
            if (e.r < startCount && !changed.has(e.r)) changed.set(e.r, p.getSourceRange(e.r));
        });
        ordered.forEach(function (e) {
            p.setCell(e.r, e.c, e.text);
        });
        changed.forEach(function (range, r) {
            lineEdits.push({ start: range.start, end: range.end, text: p.serializeRow(r) });
        });
        if (appended.length) {
            const text = appended
                .map(function (a) {
                    return p.serializeRow(a.r);
                })
                .join('\n');
            lineEdits.push({ start: appended[0].range.start, end: appended[0].range.start, text: text });
        }
        // New rows are not part of a sort/filter order; drop the view so they are visible.
        if (appended.length && viewActive) {
            shell.clearView();
            shell.setMessage('Sort/filter cleared: rows were added');
        }
        shell.grid.invalidate();
        shell.grid.refresh();
        shell._updateAll();
        post(lineEdits);
    }

    function addRow(): void {
        if (readOnly || !provider || busy()) return;
        const r = provider.appendRow();
        const range = provider.getSourceRange(r);
        post([{ start: range.start, end: range.start, text: provider.serializeRow(r) }]);
        if (shell._order) shell.clearView();
        shell.grid.invalidate();
        shell.grid.select(r, 0);
        shell.grid.scrollTo(r, 0);
    }

    function deleteRows(rows: number[]): void {
        if (readOnly || !provider || !rows.length || busy()) return;
        const ranges = provider.deleteRows(rows);
        const merged: LineEdit[] = [];
        ranges.forEach(function (g) {
            const last = merged[merged.length - 1];
            if (last && last.end === g.start) last.end = g.end;
            else merged.push({ start: g.start, end: g.end, text: '' });
        });
        post(merged);
        shell.clearView();
        shell.grid.invalidate();
        shell.grid.refresh();
    }

    const later =
        typeof MessageChannel !== 'undefined' &&
        typeof window !== 'undefined' &&
        typeof window.addEventListener === 'function'
            ? (function () {
                  const ch = new MessageChannel(),
                      q: (() => void)[] = [];
                  ch.port1.onmessage = function () {
                      const f = q.shift();
                      if (f) f();
                  };
                  return function (fn: () => void): void {
                      q.push(fn);
                      ch.port2.postMessage(0);
                  };
              })()
            : function (fn: () => void): void {
                  setTimeout(fn, 0);
              };
    const SYNC_LIMIT = 4000000,
        FIRST_ROWS = 2000;
    let job: LoadJob | null = null,
        enc: TextEncoder | null = null;

    function parseOptions(): CsvParseOptions {
        return {
            separator: options.separator,
            quote: options.quoteMark || '"',
            hasHeaders: options.hasHeaders,
            commentChar: options.commentCharacter,
            skipComments: options.skipComments
        };
    }

    function beginLoad(expectedChars: number, id: number | null | undefined): LoadJob {
        if (job) job.cancelled = true;
        transferring = id !== null && id !== undefined;
        const prevProvider = provider;
        job = {
            id: id,
            seq: 0,
            cancelled: false,
            final: false,
            pumping: false,
            swapped: false,
            sync: false,
            buf: new Uint8Array(Math.max(1024, expectedChars | 0)),
            filled: 0,
            expected: expectedChars | 0,
            ix: new CsvIndex(parseOptions()),
            t0: performance.now(),
            lastPos: -1,
            lastInv: 0,
            prev: prevProvider ? shell.getViewState() : null,
            prevProvider: prevProvider,
            state0: null
        };
        return job;
    }

    function feed(j: LoadJob, text: string): void {
        enc = enc || new TextEncoder();
        let pos = 0;
        while (pos < text.length) {
            const rest = pos ? text.slice(pos) : text;
            const r = enc.encodeInto(rest, j.buf.subarray(j.filled));
            j.filled += r.written;
            pos += r.read;
            if (r.read < rest.length) {
                const nb = new Uint8Array(
                    Math.max(Math.ceil(j.buf.length * 1.2), j.filled + (rest.length - r.read) + 65536)
                );
                nb.set(j.buf.subarray(0, j.filled));
                j.buf = nb;
            }
        }
        j.ix.setSource(j.buf, j.filled, false);
    }

    function makeProvider(j: LoadJob): void {
        const ix = j.ix;
        const p = new CsvProvider(ix, {
            hasHeaders: options.hasHeaders,
            capitalizeHeaders: options.capitalizeHeaders,
            formatValues: options.formatValues,
            numberFormat: options.numberFormat,
            lineNumbers: options.lineNumbers,
            dark: shell.isDark(),
            measure: measure
        });
        provider = p;
        if (j.prevProvider)
            j.prevProvider._userWidths.forEach(function (w, c) {
                p.setColWidth(c, w);
            });
        shell.setProvider(p);
        j.swapped = true;
        if (ix.done) autosize();
        if (ix.sepDirective && !sepSent) {
            sepSent = true;
            sendMessage({ separator: ix.sepDirective });
        }
        const vs = j.prev || (saved && isViewStateInput(saved.view) ? saved.view : null) || null;
        shell.applyViewState(vs);
        j.state0 = JSON.stringify(shell.getViewState());
        j.vs = vs;
    }

    function status(j: LoadJob, done: boolean): void {
        if (done) {
            const n = need().rowCount - (need().hasHeaders ? 1 : 0);
            shell.setMessage(
                n.toLocaleString() +
                    ' rows, ' +
                    need().colCount +
                    ' columns  |  parsed in ' +
                    Math.round(performance.now() - j.t0) +
                    ' ms'
            );
        } else {
            const total = j.final ? j.filled : Math.max(j.expected, j.filled);
            const pct = total ? Math.min(99, Math.floor((100 * j.ix.pos) / total)) : 0;
            shell.setMessage(
                'Parsing ' +
                    pct +
                    '%  |  ' +
                    Math.max(0, provider ? provider.rowCount - (provider.hasHeaders ? 1 : 0) : 0).toLocaleString() +
                    ' rows'
            );
        }
    }

    function autosize(): void {
        if (options.resizeColumns !== 'all') return;
        need().autosizeAll();
        shell.grid.invalidate();
    }

    function complete(j: LoadJob): void {
        const p = need();
        p.sync();
        if (j.prevProvider && j.prevProvider.dataRowCount === p.dataRowCount) {
            j.prevProvider._userHeights.forEach(function (h, r) {
                p.setRowHeight(r, h);
            });
        }
        shell.loading = false;
        transferring = false;
        docVersion = j.version === undefined ? null : j.version;
        if (j.reapply) {
            autosize();
            shell.grid.invalidate();
            if (j.vs && JSON.stringify(shell.getViewState()) === j.state0) shell.applyViewState(j.vs);
        }
        status(j, true);
        job = null;
    }

    // Indexes available data in slices; the grid shows the first rows as soon as they exist.
    function step(j: LoadJob): void {
        if (j.cancelled) return;
        const ix = j.ix;
        ix.run(j.sync ? 0 : 12);
        if (!j.swapped && (ix.done || ((ix.header !== null || !options.hasHeaders) && ix.n >= FIRST_ROWS))) {
            makeProvider(j);
            j.reapply = !ix.done;
            shell.loading = !ix.done;
        } else if (j.swapped) {
            need().sync();
            const t = performance.now();
            if (!ix.done && t - j.lastInv > 300) {
                j.lastInv = t;
                shell.grid.invalidate();
            }
        }
        if (ix.done) {
            complete(j);
            return;
        }
        status(j, false);
        if (ix.pos !== j.lastPos) {
            j.lastPos = ix.pos;
            later(function () {
                step(j);
            });
        } else j.pumping = false;
    }

    function pump(j: LoadJob): void {
        if (j.pumping || j.cancelled) return;
        j.pumping = true;
        j.lastPos = -1;
        later(function () {
            step(j);
        });
    }

    function finishLoad(j: LoadJob): void {
        j.final = true;
        j.ix.setSource(j.buf, j.filled, true);
        if (j.filled <= SYNC_LIMIT) {
            j.sync = true;
            j.pumping = true;
            step(j);
        } else pump(j);
    }

    // Whole document as one string (small files, tests, tools).
    function load(text: string | null | undefined, version?: number): void {
        text = text || '';
        const j = beginLoad(0, null);
        j.version = version;
        enc = enc || new TextEncoder();
        j.buf = enc.encode(text);
        j.filled = j.buf.length;
        j.expected = j.filled;
        finishLoad(j);
    }

    // Large documents arrive as ordered chunks: {refreshChunk:'start'|'data'|'end', id, ...}; a newer start cancels the old transfer.
    function chunk(m: ChunkMessage): void {
        if (m.refreshChunk === 'start') {
            beginLoad(m.chars, m.id).version = m.version;
            shell.setMessage('Loading...');
            return;
        }
        const j = job;
        if (!j || j.id !== m.id || j.final) return;
        if (m.refreshChunk === 'data') {
            if (m.seq !== j.seq) {
                j.cancelled = true;
                job = null;
                transferring = false;
                sendMessage({ refresh: true });
                return;
            }
            j.seq++;
            feed(j, m.text);
            if (!j.swapped)
                shell.setMessage('Loading ' + Math.floor((100 * j.filled) / Math.max(1, j.expected)) + '%...');
            pump(j);
        } else if (m.refreshChunk === 'end') {
            if (m.count !== j.seq) {
                j.cancelled = true;
                job = null;
                transferring = false;
                sendMessage({ refresh: true });
                return;
            }
            finishLoad(j);
        }
    }

    window.oxCsv = { load: load, chunk: chunk, shell: shell };
    shell.setMessage('Opening...');
    sendMessage({ refresh: true });
}

export function handleEvents(): void {
    window.addEventListener('message', function (event: MessageEvent<unknown>) {
        const d = event.data;
        if (!window.oxCsv || !isHostMessage(d)) return;
        if (isChunkMessage(d)) window.oxCsv.chunk(d);
        else window.oxCsv.load(d.content, d.version);
    });
}

// Entry point called by the inline script of the page: xlPage.start(options, ignoreState)
export function start(options: CsvPageOptions, ignoreState?: boolean): void {
    handleEvents();
    initPage(options, ignoreState);
}
