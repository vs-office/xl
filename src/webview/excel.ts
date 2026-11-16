import { Shell } from './core/app-shell.ts';
import type { ShellEditSource, ShellViewState, ShellViewStateInput } from './core/app-shell.ts';
import type * as EditNamespace from './edit.ts';
import type { Calc, CalcResult, CellPos } from './core/calc.ts';
import type { GridEdit, GridLinkEvent } from './core/grid.ts';
import { XlsxProvider } from './core/xlsx-provider.ts';
import type { EditCell, EditError, PendingEdit, SlotCell } from './core/xlsx-provider.ts';
import { readXlsx } from './core/xlsx-reader.ts';
import type { CellStyle, Hyperlink, Sheet, Workbook } from './core/xlsx-reader.ts';

/** The options object the extension host passes to the page (counterpart of the host's ExcelOptions). */
export interface ExcelOptions {
    customEditor: boolean;
    uri: string;
    scriptUri: string;
    directUrl: string | undefined;
    previewUri: string;
    /** the host's persisted state; untrusted, validated before use */
    state: unknown;
}

/** Per-sheet view states (zoom, scroll, active cell) keyed by sheet index. */
export type ExcelViews = Record<string, ShellViewStateInput>;

/** The page state persisted through vscode.setState and sent to the host with `save`. */
export interface ExcelPageState {
    version: string;
    uri: string;
    sheet: number;
    views: ExcelViews;
}

/** The state restored from vscode.getState() / options.state once validated. */
export interface ExcelSavedState {
    version: string;
    sheet?: number | undefined;
    views?: ExcelViews | undefined;
}

// ---- host -> page messages ----
/** The host asks the page to (re)load the workbook; `content` is the file's bytes (see load). */
export interface ExcelRefreshMessage {
    refresh: true;
    content?: unknown;
}
/** The host asks for the current file bytes (a save / backup); answered with an ExcelResponseMessage. */
export interface ExcelGetDataMessage {
    type: 'getData';
    requestId: number;
}
export interface ExcelUndoMessage {
    undo: true;
}
export interface ExcelRedoMessage {
    redo: true;
}
/** Legacy form of a reload: file bytes without a type. */
export interface ExcelContentMessage {
    content: unknown;
    type?: undefined;
}
export type ExcelHostMessage =
    ExcelRefreshMessage | ExcelGetDataMessage | ExcelUndoMessage | ExcelRedoMessage | ExcelContentMessage;

// ---- page -> host messages ----
export interface ExcelRefreshRequest {
    refresh: true;
}
export interface ExcelSaveMessage {
    save: true;
    state: ExcelPageState;
}
export interface ExcelOpenLinkMessage {
    openLink: string;
}
export interface ExcelErrorMessage {
    error: string;
}
export interface ExcelChangedMessage {
    changed: true;
    reason: 'Paste' | 'Edit';
}
export interface ExcelDataResponse {
    response: true;
    requestId: number;
    body: Uint8Array | null;
}
export interface ExcelErrorResponse {
    response: true;
    requestId: number;
    error: string;
}
export type ExcelResponseMessage = ExcelDataResponse | ExcelErrorResponse;
export type ExcelPageMessage =
    | ExcelRefreshRequest
    | ExcelSaveMessage
    | ExcelOpenLinkMessage
    | ExcelErrorMessage
    | ExcelChangedMessage
    | ExcelResponseMessage;

/** The page API the host messages drive (window.oxExcel). */
export interface ExcelApi {
    load(content: unknown): void;
    undo(): void;
    redo(): void;
    getBytes(): Uint8Array | null;
    getBytesAsync(): Promise<Uint8Array | null>;
    /** performance.now() at page start */
    t0?: number | undefined;
}

declare global {
    interface Window {
        oxExcel?: ExcelApi | undefined;
    }
}

/** The lazily loaded editing chunk (formula engine and writer). */
export type ExcelEditModule = typeof EditNamespace;

/** One edited cell of an undo transaction; `si` is set when the cell is on another sheet than the transaction's. */
export interface UndoItem {
    si?: number | undefined;
    r: number;
    c: number;
    before: SlotCell | null | undefined;
    after: SlotCell | null | undefined;
}
export interface UndoTransaction {
    si: number;
    items: UndoItem[];
}

type Dict = Record<string, unknown>;

function isDict(v: unknown): v is Dict {
    return typeof v === 'object' && v !== null;
}

function isNumberList(v: unknown): v is ArrayLike<number> {
    return isDict(v) && typeof v.length === 'number';
}

function toBytes(raw: unknown): Uint8Array {
    if (raw instanceof Uint8Array) return raw;
    if (raw instanceof ArrayBuffer) return new Uint8Array(raw);
    if (isNumberList(raw)) return new Uint8Array(raw);
    return new Uint8Array(0);
}

/** The message text of an error-like value, else the value as text. */
function errText(err: unknown): string {
    if (isDict(err)) {
        const m = err.message;
        if (typeof m === 'string' && m) return m;
    }
    return String(err);
}

function parseViewState(v: unknown): ShellViewStateInput | null {
    if (!isDict(v)) return null;
    const out: ShellViewStateInput = {};
    if (typeof v.zoom === 'number') out.zoom = v.zoom;
    const sc = v.scroll;
    if (isDict(sc)) {
        const scroll: { top?: number; left?: number } = {};
        if (typeof sc.top === 'number') scroll.top = sc.top;
        if (typeof sc.left === 'number') scroll.left = sc.left;
        out.scroll = scroll;
    }
    const a = v.active;
    if (isDict(a) && typeof a.r === 'number' && typeof a.c === 'number') out.active = { r: a.r, c: a.c };
    return out;
}

/** Validates a persisted page state of the current version; null for anything else. */
export function parseSavedState(s: unknown, version: string): ExcelSavedState | null {
    if (!isDict(s) || s.version !== version) return null;
    const out: ExcelSavedState = { version };
    if (typeof s.sheet === 'number') out.sheet = s.sheet;
    if (isDict(s.views)) {
        const views: ExcelViews = {};
        for (const k of Object.keys(s.views)) {
            const vs = parseViewState(s.views[k]);
            if (vs) views[k] = vs;
        }
        out.views = views;
    }
    return out;
}

export function isRefreshMessage(d: unknown): d is ExcelRefreshMessage {
    return isDict(d) && !!d.refresh;
}
export function isGetDataMessage(d: unknown): d is ExcelGetDataMessage {
    return isDict(d) && d.type === 'getData' && typeof d.requestId === 'number';
}
export function isUndoMessage(d: unknown): d is ExcelUndoMessage {
    return isDict(d) && !!d.undo;
}
export function isRedoMessage(d: unknown): d is ExcelRedoMessage {
    return isDict(d) && !!d.redo;
}
export function isContentMessage(d: unknown): d is ExcelContentMessage {
    return isDict(d) && !!d.content && !d.type;
}
/** Narrows the untyped `event.data` of a window message to a host message. */
export function isExcelHostMessage(d: unknown): d is ExcelHostMessage {
    return isRefreshMessage(d) || isGetDataMessage(d) || isUndoMessage(d) || isRedoMessage(d) || isContentMessage(d);
}

let sendMessage: (m: ExcelPageMessage) => void = () => undefined;

export function initPage(options: ExcelOptions, ignoreState?: boolean): void {
    const vscode = acquireVsCodeApi();
    // the lazily loaded editing chunk is fetched next to the page script
    if (typeof __webpack_public_path__ !== 'undefined' && options.scriptUri)
        __webpack_public_path__ = options.scriptUri + '/';
    const STATE_VERSION = '6';
    const readOnly = !options.customEditor;
    sendMessage = function (m: ExcelPageMessage) {
        vscode.postMessage(m);
    };

    const app = document.getElementById('app');
    let wb: Workbook | null = null,
        bytes: Uint8Array | null = null,
        active = 0;
    let calc: Calc | null = null;
    let edit: ExcelEditModule | null = null; // editing chunk namespace once loaded: { Calc, patchXlsx, validateFormula }
    const providers = new Map<number, XlsxProvider>();
    // a sheet whose load was cancelled keeps its first rows, read only
    const cancelledPartial = new WeakSet<XlsxProvider>();
    const views: ExcelViews = {};
    let undoStack: UndoTransaction[] = [],
        redoStack: UndoTransaction[] = [],
        undoCells = 0;
    let shell: Shell | null = null,
        loadCtl: AbortController | null = null,
        cancelBtn: HTMLButtonElement | null = null;
    // repositions the floating sheet pictures (installed by createShell)
    let layoutPics: () => void = () => undefined;

    function requireShell(): Shell {
        if (!shell) throw new TypeError('The grid shell is not created yet');
        return shell;
    }
    function requireWb(): Workbook {
        if (!wb) throw new TypeError('No workbook is loaded');
        return wb;
    }
    function requireEdit(): ExcelEditModule {
        if (!edit) throw new TypeError('The editing module is not loaded');
        return edit;
    }
    // the provider shown by the shell; the page only ever mounts XlsxProviders
    function shownProvider(): XlsxProvider | null {
        const p = requireShell().provider;
        return p instanceof XlsxProvider ? p : null;
    }

    const saved: ExcelSavedState | null = (function () {
        if (ignoreState) return null;
        return parseSavedState(vscode.getState() || options.state, STATE_VERSION);
    })();

    let measureCtx: CanvasRenderingContext2D | null = null;
    function measure(text: string, style: CellStyle | null): number {
        const c = measureCtx || (measureCtx = document.createElement('canvas').getContext('2d'));
        if (!c) throw new TypeError('Cannot get a 2D canvas context');
        c.font =
            (style && style.b ? 'bold ' : '') +
            (style && style.size ? (style.size * 4) / 3 : 13) +
            'px ' +
            ((style && style.font) || 'system-ui');
        return c.measureText(text).width;
    }

    function createShell(): void {
        if (!app) throw new TypeError('The page has no #app element');
        const sh = (shell = new Shell(app, {
            readOnly: readOnly,
            tabs: true,
            headerRows: function () {
                const p = shell && shell.provider;
                return p instanceof XlsxProvider ? p.frozenRows || 1 : 0;
            },
            onEdit: applyEdits,
            onViewState: function (vs: ShellViewState) {
                views[active] = vs;
                const state: ExcelPageState = { version: STATE_VERSION, uri: options.uri, sheet: active, views: views };
                vscode.setState(state);
                vscode.postMessage({ save: true, state: state });
            },
            onLink: onLink
        }));
        const picLayer = document.createElement('div');
        picLayer.className = 'ox-pics';
        sh.grid.el.appendChild(picLayer);
        // floating sheet pictures (drawing anchors): read-only overlay positioned over the canvas
        layoutPics = function () {
            const g = sh.grid,
                p = sh.provider instanceof XlsxProvider ? sh.provider : null;
            const pics = (p && p.sheet && p.sheet.pictures) || [];
            while (picLayer.children.length > pics.length && picLayer.lastChild)
                picLayer.removeChild(picLayer.lastChild);
            if (!pics.length) return;
            const W = g.el.clientWidth,
                H = g.el.clientHeight,
                z = g.getZoom();
            for (let i = 0; i < pics.length; i++) {
                const q = pics[i];
                const child = picLayer.children[i];
                let img: HTMLImageElement;
                if (child instanceof HTMLImageElement) img = child;
                else {
                    img = document.createElement('img');
                    if (child) picLayer.replaceChild(img, child);
                    else picLayer.appendChild(img);
                }
                if (img.getAttribute('src') !== q.url) img.setAttribute('src', q.url);
                img.alt = '';
                const rc = g.cellRect(q.r, q.c);
                const L = Math.round(rc.x + q.dx * z),
                    T = Math.round(rc.y + q.dy * z),
                    Wd = Math.max(1, Math.round(q.w * z)),
                    Ht = Math.max(1, Math.round(q.h * z));
                img.style.cssText = 'left:' + L + 'px;top:' + T + 'px;width:' + Wd + 'px;height:' + Ht + 'px;';
                img.style.display = L + Wd < 0 || T + Ht < 0 || L > W || T > H ? 'none' : '';
            }
        };
        sh.grid.scroller.addEventListener(
            'scroll',
            function () {
                layoutPics();
            },
            { passive: true }
        );
        sh.grid.on('zoom', function () {
            layoutPics();
        });
        sh.grid.on('resizeCol', function () {
            layoutPics();
        });
        sh.grid.on('resizeRow', function () {
            layoutPics();
        });
        document.addEventListener('keydown', function (e) {
            if (
                (e.ctrlKey || e.metaKey) &&
                !e.shiftKey &&
                !e.altKey &&
                wb &&
                (e.key === 'PageDown' || e.key === 'PageUp')
            ) {
                const vis = visibleSheets();
                const i = vis.indexOf(active);
                const n = vis[(i + (e.key === 'PageDown' ? 1 : vis.length - 1)) % vis.length];
                if (n !== undefined && n !== active) {
                    e.preventDefault();
                    showSheet(n);
                }
            }
        });
    }

    function onLink(e: GridLinkEvent): void {
        if (!e || !wb) return;
        const p = shownProvider();
        const link: Hyperlink = (p && p.sheet && p.sheet.hyperlinks && p.sheet.hyperlinks.get(e.sr + ',' + e.c)) || {};
        const url = e.url || link.url,
            loc = e.location || link.location;
        if (url) sendMessage({ openLink: url });
        else if (loc) goLocation(loc);
    }

    function goLocation(loc: string): void {
        const shell = requireShell(),
            wb = requireWb();
        const m = /^(?:'((?:[^']|'')+)'|([^!']+))!(.+)$/.exec(String(loc).trim());
        let ref = String(loc).trim(),
            idx = active;
        if (m) {
            const name = (m[1] ? m[1].replace(/''/g, "'") : m[2]).toLowerCase();
            const sh = wb.sheets.filter(function (s) {
                return s.name.toLowerCase() === name && s.state === 'visible';
            })[0];
            if (!sh) {
                shell.setMessage('Link target sheet not found: ' + (m[1] || m[2]));
                return;
            }
            idx = sh.index;
            ref = m[3];
        }
        if (idx !== active) showSheet(idx);
        if (!shell._goto(ref)) shell.setMessage('Cannot follow link: ' + loc);
    }

    function visibleSheets(): number[] {
        const out: number[] = [];
        requireWb().sheets.forEach(function (s) {
            if (s.state === 'visible') out.push(s.index);
        });
        return out.length ? out : [0];
    }

    function renderTabs(): void {
        const shell = requireShell(),
            wb = requireWb();
        shell.tabs.textContent = '';
        visibleSheets().forEach(function (i) {
            const info = wb.sheets[i];
            const b = document.createElement('button');
            b.className = 'ox-tab' + (i === active ? ' active' : '');
            b.textContent = info.name;
            if (info.tabColor) b.style.boxShadow = 'inset 0 -3px 0 ' + info.tabColor;
            b.addEventListener('click', function () {
                if (i !== active) showSheet(i);
            });
            shell.tabs.appendChild(b);
        });
    }

    function newProvider(sheet: Sheet): XlsxProvider {
        return new XlsxProvider(requireWb(), sheet, {
            measure: measure,
            validateFormula: function (f) {
                return edit ? edit.validateFormula(f) : null;
            }
        });
    }

    function loadError(i: number, err: unknown): void {
        const wb = requireWb();
        const msg = 'Cannot load sheet "' + (wb.sheets[i] ? wb.sheets[i].name : i) + '": ' + errText(err);
        requireShell().setMessage(msg);
        sendMessage({ error: 'xl: ' + msg });
    }

    function setCancel(fn: (() => void) | null): void {
        const shell = requireShell();
        if (!cancelBtn) {
            cancelBtn = document.createElement('button');
            cancelBtn.className = 'ox-tab';
            cancelBtn.textContent = 'Cancel';
            cancelBtn.style.cssText = 'margin-left:8px;cursor:pointer';
            shell.status.insertBefore(cancelBtn, shell.msg.nextSibling);
        }
        cancelBtn.onclick = fn;
        cancelBtn.style.display = fn ? '' : 'none';
    }

    function mount(p: XlsxProvider, i: number, viewState?: ShellViewStateInput, loading?: boolean): void {
        const shell = requireShell();
        shell.setProvider(p);
        if (!loading) shell.setMessage('Ready');
        shell.applyViewState(viewState || views[i] || null);
        layoutPics();
        shell.grid.focus();
    }

    // streams the sheet in; the first rows are shown as soon as they are parsed, the rest keeps loading in time slices
    function showSheet(i: number, viewState?: ShellViewStateInput): void {
        const shell = requireShell(),
            wb = requireWb();
        if (loadCtl) {
            loadCtl.abort();
            loadCtl = null;
            setCancel(null);
        }
        const prev = active;
        active = i;
        renderTabs();
        const done = providers.get(i);
        if (done) {
            mount(done, i, viewState);
            return;
        }
        const ctl = new AbortController(),
            name = wb.sheets[i].name;
        let p: XlsxProvider | null = null,
            lastInv = 0;
        loadCtl = ctl;
        shell.setMessage('Loading ' + name + '... 0%');
        setCancel(function () {
            ctl.abort();
            if (p) {
                cancelledPartial.add(p);
                shell.setMessage(
                    'Loading cancelled: showing the first ' + p.sheet.rowCount.toLocaleString() + ' rows (read only)'
                );
                loadCtl = null;
                setCancel(null);
            } else if (prev !== i && providers.has(prev)) {
                loadCtl = null;
                setCancel(null);
                showSheet(prev);
                shell.setMessage('Loading cancelled');
            } else {
                loadCtl = null;
                setCancel(null);
                shell.setMessage('Loading cancelled');
            }
        });
        wb.loadSheetAsync(i, {
            signal: ctl.signal,
            onProgress: function (info) {
                if (ctl.signal.aborted || loadCtl !== ctl) return;
                if (info.partial)
                    shell.setMessage(
                        'Loading ' +
                            name +
                            '... ' +
                            Math.floor(info.fraction * 100) +
                            '%  (' +
                            info.rows.toLocaleString() +
                            ' rows)'
                    );
                else return;
                if (!p && info.rows > 0) {
                    p = newProvider(info.sheet);
                    mount(p, i, viewState, true);
                    lastInv = Date.now();
                } else if (p && Date.now() - lastInv > 500) {
                    shell.grid.invalidate();
                    lastInv = Date.now();
                }
            }
        })
            .then(function (sh) {
                if (loadCtl !== ctl) return;
                loadCtl = null;
                setCancel(null);
                if (!p) p = newProvider(sh);
                providers.set(i, p);
                if (!readOnly)
                    setTimeout(() => {
                        void ensureEdit().catch(() => undefined);
                    }, 0);
                if (shell.provider === p) {
                    shell.grid.invalidate();
                    layoutPics();
                    if (shell.view && (shell.view.sort || shell.view.filters.size)) shell.recomputeView();
                    shell._updateAll();
                    shell.setMessage(sh.warning ? sh.warning : 'Ready');
                } else mount(p, i, viewState);
            })
            .catch(function (err: unknown) {
                if (isDict(err) && err.name === 'AbortError') return;
                if (loadCtl === ctl) {
                    loadCtl = null;
                    setCancel(null);
                }
                loadError(i, err);
            });
    }

    function load(content: unknown): void {
        const shell = requireShell();
        const raw = isDict(content) && content.data !== undefined ? content.data : content;
        const reloading = !!wb;
        // parse first: a file that fails to open must not wipe the workbook (and unsaved edits) on screen
        let nb: Workbook, nbytes: Uint8Array;
        try {
            nbytes = toBytes(raw);
            nb = readXlsx(nbytes);
        } catch (err) {
            const why = errText(err);
            shell.setMessage(
                (reloading ? 'Cannot reload file (keeping the previous content): ' : 'Cannot open file: ') + why
            );
            sendMessage({ error: 'xl: ' + why });
            return;
        }
        console.log(
            '[xl] workbook parsed ' +
                Math.round(performance.now() - (window.oxExcel?.t0 ?? NaN)) +
                ' ms after start, ' +
                nbytes.length +
                ' bytes'
        );
        bytes = nbytes;
        wb = nb;
        if (loadCtl) {
            loadCtl.abort();
            loadCtl = null;
            setCancel(null);
        }
        providers.clear();
        calc = null;
        undoStack = [];
        redoStack = [];
        undoCells = 0;
        let idx: number;
        const vis = visibleSheets();
        if (saved && saved.views)
            Object.keys(saved.views).forEach(function (k) {
                const sv = saved.views?.[k];
                if (!views[k] && sv) views[k] = sv;
            });
        // On a host refresh keep the sheet the user is on; the saved state only decides the first load.
        if (reloading && vis.indexOf(active) >= 0) idx = active;
        else if (saved && saved.sheet !== undefined && vis.indexOf(saved.sheet) >= 0) idx = saved.sheet;
        else idx = vis[0];
        try {
            showSheet(idx);
        } catch (err) {
            loadError(idx, err);
        }
    }

    function providerFor(si: number): XlsxProvider {
        let q = providers.get(si);
        if (!q) {
            q = newProvider(requireWb().loadSheet(si));
            providers.set(si, q);
        }
        return q;
    }

    function getCalc(ed: ExcelEditModule): Calc {
        if (calc) return calc;
        const wb = requireWb();
        calc = new ed.Calc({
            sheetNames: wb.sheets.map(function (x) {
                return x.name;
            }),
            date1904: wb.date1904,
            sheetIndexByName: function (n: string) {
                const l = String(n).toLowerCase();
                for (let i = 0; i < wb.sheets.length; i++) if (wb.sheets[i].name.toLowerCase() === l) return i;
                return -1;
            },
            getCell: function (si, r, c) {
                return providerFor(si).cellAt(r, c);
            },
            getExtent: function (si) {
                const q = providerFor(si);
                return { rowCount: q.rowCount, colCount: q.colCount };
            },
            forEachCell: function (si, r1, c1, r2, c2, cb) {
                providerFor(si).forEachCell(r1, c1, r2, c2, function (cell, r, c) {
                    cb(r, c, cell);
                });
            }
        });
        return calc;
    }

    function inArrayFormula(q: XlsxProvider, r: number, c: number): boolean {
        const arr = q.sheet && q.sheet.arrays;
        if (!arr) return false;
        for (let i = 0; i < arr.length; i++) {
            const a = arr[i];
            if (r >= a.r1 && r <= a.r2 && c >= a.c1 && c <= a.c2) return true;
        }
        return false;
    }

    // writes a calculated value into the cell, keeping its formula and style; recorded in the undo transaction
    function putCalculated(
        q: XlsxProvider,
        si: number,
        r: number,
        c: number,
        res: CalcResult,
        items: UndoItem[],
        seen: Map<string, UndoItem>
    ): boolean {
        const cell = q.cellAt(r, c);
        if (!cell || !cell.f || inArrayFormula(q, r, c)) return false;
        let out: CalcResult = res;
        if (res.t === 'e' && res.v === '#CIRC!') out = { t: 'n', v: 0 };
        if (cell.t === out.t && cell.v === out.v) return false;
        const key = si + ':' + r + ':' + c;
        let it = seen.get(key);
        const before = q.setCell(r, c, calculatedCell(out, cell.f, cell.s));
        const after = q.overlay.has(r * 16384 + c) ? q.overlay.get(r * 16384 + c) : undefined;
        if (it) it.after = after;
        else {
            it = { si: si, r: r, c: c, before: before, after: after };
            seen.set(key, it);
            items.push(it);
        }
        return true;
    }

    // the cell written for a calculated value: the result plus the cell's formula and style
    function calculatedCell(res: CalcResult, f: string, s: number): EditCell {
        switch (res.t) {
            case 'n':
                return { t: 'n', v: res.v, f: f, s: s };
            case 's':
                return { t: 's', v: res.v, f: f, s: s };
            case 'b':
                return { t: 'b', v: res.v, f: f, s: s };
            case 'e':
                return { t: 'e', v: res.v, f: f, s: s };
        }
    }

    // result cells inside array formulas are not recomputed (the engine cannot spill); these are reported to the user
    let staleArrays = 0;
    function noteArray(q: XlsxProvider, r: number, c: number): void {
        if (inArrayFormula(q, r, c)) staleArrays++;
    }

    function recalculate(q: XlsxProvider, si: number, items: UndoItem[]): number {
        if (!edit) return 0;
        let n = 0;
        staleArrays = 0;
        try {
            const engine = getCalc(edit);
            const seen = new Map<string, UndoItem>();
            items.forEach(function (it) {
                seen.set(si + ':' + it.r + ':' + it.c, it);
            });
            items.slice().forEach(function (it) {
                const cell = q.cellAt(it.r, it.c);
                if (cell && cell.f) {
                    const res = engine.evaluate(si, it.r, it.c, cell.f);
                    if (putCalculated(q, si, it.r, it.c, res, items, seen)) n++;
                }
            });
            let changed: CellPos[] = items.map(function (it) {
                return { sheet: it.si !== undefined ? it.si : si, r: it.r, c: it.c };
            });
            for (let round = 0; round < 40 && changed.length; round++) {
                const res = engine.recalcDependents(changed, { maxEvals: 500000, timeBudgetMs: 250, volatile: false });
                res.forEach(function (val) {
                    const vq = providerFor(val.sheet);
                    noteArray(vq, val.r, val.c);
                    if (putCalculated(vq, val.sheet, val.r, val.c, val, items, seen)) n++;
                });
                changed = res.partial && res.pending ? res.pending : [];
            }
        } catch (err) {
            requireShell().setMessage('Formulas were not recalculated: ' + errText(err));
            calc = null;
            return -1;
        }
        return n;
    }

    // the formula engine and the writer are a separate chunk, fetched once editing is possible
    let editReady: Promise<void> | null = null;
    function hasEditModule(): boolean {
        return edit !== null;
    }
    function ensureEdit(): Promise<void> {
        if (edit) return Promise.resolve();
        if (!editReady)
            editReady = import(/* webpackChunkName: "excel-edit" */ './edit.ts').then(
                function (m) {
                    edit = m;
                },
                function () {
                    editReady = null;
                    throw new Error('Cannot load the editing module');
                }
            );
        return editReady;
    }

    function applyEdits(edits: GridEdit[], source: ShellEditSource): void {
        if (readOnly || !wb) return;
        const shell = requireShell();
        if (!hasEditModule()) {
            shell.setMessage('Loading editor...');
            ensureEdit().then(
                function () {
                    applyEdits(edits, source);
                },
                function (err) {
                    shell.setMessage('Edit failed: ' + errText(err));
                    shell.grid.refresh();
                }
            );
            return;
        }
        const p = shownProvider(),
            si = active,
            items: UndoItem[] = [];
        if (!p) throw new TypeError('No sheet is shown');
        if (p.sheet && p.sheet.partial) {
            shell.setMessage(
                cancelledPartial.has(p)
                    ? 'Loading was cancelled: this sheet is read only. Switch sheets and back to load it again.'
                    : 'The sheet is still loading: wait for it to finish before editing.'
            );
            shell.grid.refresh();
            return;
        }
        let firstError: string | null = null,
            rejected = 0;
        for (const e of edits) {
            if (p.getRaw(e.r, e.c) === e.text) continue;
            let cell: EditCell | EditError | null | undefined = p.editFromText(e.r, e.c, e.text);
            if (cell === undefined) continue;
            if (cell && 'error' in cell) {
                rejected++;
                if (!firstError) firstError = cell.error;
                continue;
            }
            if (cell && cell.cleared) cell = null;
            const before = p.setCell(e.r, e.c, cell);
            items.push({
                r: e.r,
                c: e.c,
                before: before,
                after: p.overlay.has(e.r * 16384 + e.c) ? p.overlay.get(e.r * 16384 + e.c) : undefined
            });
        }
        if (firstError)
            shell.setMessage('Edit rejected: ' + firstError + (rejected > 1 ? ' (' + rejected + ' cells)' : ''));
        if (!items.length) {
            shell.grid.refresh();
            return;
        }
        const recalced = recalculate(p, si, items);
        undoStack.push({ si: si, items: items });
        undoCells += items.length;
        while (undoStack.length > 1 && (undoStack.length > 2000 || undoCells > 500000)) {
            const oldest = undoStack.shift();
            if (oldest) undoCells -= oldest.items.length;
        }
        redoStack = [];
        shell.grid.invalidate();
        shell.grid.refresh();
        shell._updateAll();
        if (!firstError && staleArrays > 0)
            shell.setMessage(
                'Array formula results are not recalculated here and may be out of date; they update when the file is opened in Excel.'
            );
        else if (!firstError && recalced > 0)
            shell.setMessage('Recalculated ' + recalced.toLocaleString() + ' formula' + (recalced === 1 ? '' : 's'));
        sendMessage({ changed: true, reason: source === 'paste' ? 'Paste' : 'Edit' });
    }

    function undo(): void {
        const shell = requireShell();
        const tx = undoStack.pop();
        if (!tx) return;
        undoCells -= tx.items.length;
        if (tx.si !== active) showSheet(tx.si);
        for (let i = tx.items.length - 1; i >= 0; i--) {
            const it = tx.items[i];
            providerFor(it.si !== undefined ? it.si : tx.si).restoreSlot(it.r, it.c, it.before);
        }
        calc = null;
        redoStack.push(tx);
        shell.grid.refresh();
        shell._updateAll();
    }

    function redo(): void {
        const shell = requireShell();
        const tx = redoStack.pop();
        if (!tx) return;
        if (tx.si !== active) showSheet(tx.si);
        for (const it of tx.items) providerFor(it.si !== undefined ? it.si : tx.si).restoreSlot(it.r, it.c, it.after);
        calc = null;
        undoStack.push(tx);
        undoCells += tx.items.length;
        shell.grid.refresh();
        shell._updateAll();
    }

    function getBytes(): Uint8Array | null {
        const edits: PendingEdit[] = [];
        providers.forEach(function (p, si) {
            if (p.dirty) Array.prototype.push.apply(edits, p.pendingEdits(si));
        });
        if (!edits.length) return bytes;
        try {
            if (!bytes) throw new Error('originalBytes must be a Uint8Array');
            return requireEdit().patchXlsx(bytes, edits);
        } catch (err) {
            const msg = 'Could not save your edits: ' + errText(err);
            requireShell().setMessage(msg);
            throw new Error(msg, { cause: err });
        }
    }

    function getBytesAsync(): Promise<Uint8Array | null> {
        let dirty = false;
        providers.forEach(function (p) {
            if (p.dirty) dirty = true;
        });
        return (dirty ? ensureEdit() : Promise.resolve()).then(getBytes);
    }

    const api: ExcelApi = {
        load: load,
        undo: undo,
        redo: redo,
        getBytes: getBytes,
        getBytesAsync: getBytesAsync
    };
    window.oxExcel = api;

    createShell();
    requireShell().setMessage('Opening...');
    const t0 = performance.now();
    api.t0 = t0;
    if (options.directUrl) {
        // read the file straight from disk; any failure falls back to the host round trip
        void fetch(options.directUrl, { cache: 'no-store' })
            .then(function (r) {
                if (!r.ok) throw new Error('HTTP ' + r.status);
                return r.arrayBuffer();
            })
            .then(
                function (buf) {
                    return new Uint8Array(buf);
                },
                function () {
                    return null;
                }
            )
            .then(function (u8) {
                if (u8) {
                    console.log('[xl] file fetched in ' + Math.round(performance.now() - t0) + ' ms');
                    load(u8);
                } else vscode.postMessage({ refresh: true });
            });
    } else vscode.postMessage({ refresh: true });
}

export function handleEvents(): void {
    window.addEventListener('message', function (event: MessageEvent<unknown>) {
        const d = event.data,
            app = window.oxExcel;
        if (!app) return;
        if (isRefreshMessage(d)) app.load(d.content);
        else if (isGetDataMessage(d)) {
            app.getBytesAsync().then(
                function (body) {
                    sendMessage({ response: true, requestId: d.requestId, body: body });
                },
                function (err) {
                    sendMessage({
                        response: true,
                        requestId: d.requestId,
                        error: errText(err)
                    });
                }
            );
        } else if (isUndoMessage(d)) app.undo();
        else if (isRedoMessage(d)) app.redo();
        else if (isContentMessage(d)) app.load(d.content);
    });
}

// Entry point called by the inline script of the page: xlPage.start(options, ignoreState)
export function start(options: ExcelOptions, ignoreState?: boolean): void {
    handleEvents();
    initPage(options, ignoreState);
}
