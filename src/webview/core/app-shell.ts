import { Grid } from './grid.ts';
import type {
    GridActiveCell,
    GridContextMenuEvent,
    GridEdit,
    GridEditSource,
    GridFindState,
    GridLinkEvent,
    GridMenuItem,
    GridProvider,
    GridScroll,
    GridScrollTarget,
    GridSelection
} from './grid.ts';
import { colName, cellRef, parseRange, parseRef } from './cell-ref.ts';
import type { CellRef } from './cell-ref.ts';
import { generalNum as formatGeneral } from './general.ts';
import type { Range } from './grid-core.ts';
import type { SlotCell } from './xlsx-provider.ts';

// ---------- public types ----------

type TimerHandle = ReturnType<typeof setTimeout>;

/** A background task started by runTask (cancellable, progress 0..1). */
export interface Task {
    cancelled: boolean;
    done: boolean;
    progress: number;
    cancel(): void;
}
/** What runTask drives: a generator yielding progress (0..1) and returning the result. */
export type TaskGen<T> = Iterator<number, T, undefined>;
export interface TaskOptions<T> {
    /** ms per slice (default 8) */
    budget?: number | undefined;
    onDone?: ((value: T) => void) | undefined;
    onError?: ((error: unknown) => void) | undefined;
    onProgress?: ((pct: number) => void) | undefined;
}

/** Running count / sum / min / max over a range of cells. */
export interface StatsAcc {
    cnt: number;
    nums: number;
    sum: number;
    min: number;
    max: number;
}
export type StatsGen = Generator<number, StatsAcc, undefined>;

/** Order-preserving sort keys per provider row: cat 0 number (hi/lo = uint64), 1 text (lo = index into strs), 2 blank. */
export interface SortKeys {
    cat: Uint8Array;
    hi: Uint32Array;
    lo: Uint32Array;
    strs: string[];
}
/** Distinct display values of a column; ids/values/counts are null when truncated. */
export interface ColumnIds {
    ids: Int32Array | null;
    values: string[] | null;
    counts: Int32Array | null;
    truncated: boolean;
}
export interface ColumnIdsFull extends ColumnIds {
    ids: Int32Array;
    values: string[];
    counts: Int32Array;
}
/** ColumnIds plus the display order of the distinct values (set by the filter popup scan). */
export interface ScannedColumn extends ColumnIds {
    order?: Int32Array | undefined;
}

/**
 * What the Shell needs from a provider: the Grid contract plus optional capabilities (resize persistence, theme,
 * fast sort/filter/stats paths). CsvProvider and XlsxProvider both satisfy it.
 */
export interface ShellProvider extends GridProvider {
    /** the Shell wraps colName once to append sort / filter markers; a provider's own colName must return a string */
    colName?(c: number): string;
    _shellColName?: boolean | undefined;
    /** true while the file is still loading (sort and filter are disabled) */
    readonly partial?: boolean | undefined;
    readonly dataVersion?: number | undefined;
    setColWidth?(c: number, w: number): void;
    setRowHeight?(r: number, h: number): void;
    setDark?(dark: boolean): void;
    forEachCell?(
        r1: number,
        c1: number,
        r2: number,
        c2: number,
        fn: (cell: SlotCell, r: number, c: number) => void
    ): void;
    columnIds?(c: number, maxDistinct?: number, raw?: boolean): Generator<number, ColumnIds, undefined>;
    sortKeyColumn?(c: number): Generator<number, SortKeys, undefined>;
    rangeStats?(
        acc: StatsAcc,
        rows: ArrayLike<number> | null | undefined,
        r1: number,
        r2: number,
        c1: number,
        c2: number
    ): StatsGen;
}

export type TextFilterOp = 'contains' | 'equals' | 'starts' | 'ends' | 'notcontains';
export interface TextFilter {
    kind: 'text';
    op: TextFilterOp;
    q: string;
}
export interface SetFilter {
    kind: 'set';
    set: Set<string>;
}
export type ColumnFilter = TextFilter | SetFilter;
export interface SortSpec {
    c: number;
    /** > 0 ascending, < 0 descending */
    dir: number;
}
/** The (never persisted) sort / filter view applied on top of the provider. */
export interface ShellView {
    sort: SortSpec | null;
    filters: Map<number, ColumnFilter>;
}

/** Source of an onEdit call: a grid edit or the formula bar. */
export type ShellEditSource = GridEditSource | 'formula';
export interface ShellOptions {
    readOnly?: boolean | undefined;
    /** show the sheet tab strip */
    tabs?: boolean | undefined;
    onEdit?: ((edits: GridEdit[], source: ShellEditSource) => void) | undefined;
    /** number of leading rows that are never sorted / filtered */
    headerRows?: (() => number) | undefined;
    onViewState?: ((state: ShellViewState) => void) | undefined;
    onLink?: ((e: GridLinkEvent) => void) | undefined;
}

/** Zoom / scroll / active cell (source row) as persisted by the pages. */
export interface ShellViewState {
    zoom: number;
    scroll: GridScroll;
    active: CellRef | null;
}
/** Shape accepted by applyViewState (a restored, possibly partial state). */
export interface ShellViewStateInput {
    zoom?: number | undefined;
    scroll?: GridScrollTarget | null | undefined;
    active?: CellRef | null | undefined;
}

interface FindBar {
    el: HTMLDivElement;
    input: HTMLInputElement;
    count: HTMLSpanElement;
    prev: HTMLButtonElement;
    next: HTMLButtonElement;
    mc: HTMLButtonElement;
    wc: HTMLButtonElement;
    close: HTMLButtonElement;
}
interface ColCacheEntry {
    ver: number;
    res: ColumnIds;
}
interface RadixPass {
    arr: ArrayLike<number>;
    sh: number;
    h: Uint32Array;
}

const root = globalThis;
const MAX_DISTINCT = 50000;
const SMALL_STATS = 20000;
const nowMs = (): number => (typeof performance !== 'undefined' ? performance.now() : Date.now());
const tick = (function (): (fn: () => void) => void {
    if (typeof window !== 'undefined' && typeof MessageChannel !== 'undefined') {
        const ch = new MessageChannel(),
            q: (() => void)[] = [];
        ch.port1.onmessage = () => {
            const f = q.shift();
            if (f) f();
        };
        return fn => {
            q.push(fn);
            ch.port2.postMessage(0);
        };
    }
    return fn => setTimeout(fn, 0);
})();

// Runs generator `gen` in slices of ~budget ms; a yielded number is the progress (0..1). First slice runs synchronously.
function runTask<T>(gen: TaskGen<T>, opts?: TaskOptions<T>): Task {
    const o: TaskOptions<T> = opts || {};
    const budget = o.budget || 8,
        t: Task = {
            cancelled: false,
            done: false,
            progress: 0,
            cancel: () => {
                if (!t.done) {
                    t.cancelled = true;
                    t.done = true;
                    try {
                        gen.return?.();
                    } catch (e) {
                        /* closing */
                    }
                }
            }
        };
    const step = () => {
        if (t.cancelled) return;
        const t0 = nowMs();
        try {
            for (;;) {
                const r = gen.next();
                if (r.done) {
                    t.done = true;
                    if (o.onDone) o.onDone(r.value);
                    return;
                }
                if (typeof r.value === 'number') t.progress = r.value;
                if (nowMs() - t0 >= budget) break;
            }
        } catch (e) {
            t.done = true;
            if (o.onError) o.onError(e);
            else if (root.console) root.console.error(e);
            return;
        }
        if (o.onProgress) o.onProgress(t.progress);
        tick(step);
    };
    step();
    return t;
}

// Stable bottom-up merge sort of an Int32Array of indices, sliceable.
function* mergeSortGen(
    a: Int32Array,
    less: (a: number, b: number) => boolean
): Generator<number, Int32Array, undefined> {
    const n = a.length,
        RUN = 16;
    let src: Int32Array = a,
        dst: Int32Array = new Int32Array(n),
        work = 0;
    for (let s = 0; s < n; s += RUN) {
        const e = Math.min(n, s + RUN);
        for (let i = s + 1; i < e; i++) {
            const v = src[i];
            let j = i - 1;
            while (j >= s && less(v, src[j])) {
                src[j + 1] = src[j];
                j--;
            }
            src[j + 1] = v;
        }
        work += RUN;
        if (work > 40000) {
            work = 0;
            yield 0;
        }
    }
    let passes = 0;
    for (let w = RUN; w < n; w *= 2) passes++;
    for (let w = RUN, ps = 0; w < n; w *= 2, ps++) {
        for (let s = 0; s < n; s += 2 * w) {
            const m = Math.min(n, s + w),
                e = Math.min(n, s + 2 * w);
            let i = s,
                j = m,
                k = s;
            while (i < m && j < e) dst[k++] = less(src[j], src[i]) ? src[j++] : src[i++];
            while (i < m) dst[k++] = src[i++];
            while (j < e) dst[k++] = src[j++];
            work += e - s;
            if (work > 40000) {
                work = 0;
                yield (ps + s / n) / Math.max(1, passes);
            }
        }
        const t = src;
        src = dst;
        dst = t;
    }
    if (src !== a) a.set(src);
    return a;
}

// LSD radix sort of source rows by (cat, hi, lo) keys indexed by row; stable, so ties keep ascending row order.
function* radixRows(
    rows: Int32Array,
    cat: Uint8Array,
    hi: Uint32Array,
    lo: Uint32Array
): Generator<number, Int32Array, undefined> {
    const n = rows.length,
        B = 1 << 16;
    const passes: RadixPass[] = [
        { arr: lo, sh: 0, h: new Uint32Array(B) },
        { arr: lo, sh: 16, h: new Uint32Array(B) },
        { arr: hi, sh: 0, h: new Uint32Array(B) },
        { arr: hi, sh: 16, h: new Uint32Array(B) },
        { arr: cat, sh: 0, h: new Uint32Array(256) }
    ];
    const h0 = passes[0].h,
        h1 = passes[1].h,
        h2 = passes[2].h,
        h3 = passes[3].h,
        h4 = passes[4].h;
    for (let i = 0; i < n; i++) {
        const r = rows[i],
            l = lo[r],
            g = hi[r];
        h0[l & 0xffff]++;
        h1[l >>> 16]++;
        h2[g & 0xffff]++;
        h3[g >>> 16]++;
        h4[cat[r]]++;
        if ((i & 0x3ffff) === 0x3ffff) yield 0;
    }
    let src = rows.slice(),
        dst = new Int32Array(n);
    let done = 0;
    for (const ps of passes) {
        const h = ps.h;
        let trivial = false,
            sum = 0;
        for (let k = 0; k < h.length; k++) {
            if (h[k] === n) trivial = true;
            const c = h[k];
            h[k] = sum;
            sum += c;
        }
        if (trivial) {
            done++;
            continue;
        }
        const arr = ps.arr,
            sh = ps.sh,
            mask = h.length - 1;
        for (let i = 0; i < n; i++) {
            const r = src[i];
            dst[h[(arr[r] >>> sh) & mask]++] = r;
            if ((i & 0x1ffff) === 0x1ffff) yield (done + i / n) / 5;
        }
        const t = src;
        src = dst;
        dst = t;
        done++;
        yield done / 5;
    }
    return src;
}

const NUM_RE = /^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/;
const f64 = new Float64Array(1),
    u32 = new Uint32Array(f64.buffer),
    LE = new Uint8Array(new Uint16Array([1]).buffer)[0] === 1;

function matchText(spec: TextFilter): (s: string) => boolean {
    const q = spec.q.toLowerCase();
    switch (spec.op) {
        case 'equals':
            return s => s.toLowerCase() === q;
        case 'starts':
            return s => s.toLowerCase().startsWith(q);
        case 'ends':
            return s => s.toLowerCase().endsWith(q);
        case 'notcontains':
            return s => s.toLowerCase().indexOf(q) < 0;
        default:
            return s => s.toLowerCase().indexOf(q) >= 0;
    }
}

const noMatch = (_s: string): boolean => false;

function textOp(v: string): TextFilterOp {
    switch (v) {
        case 'equals':
        case 'starts':
        case 'ends':
        case 'notcontains':
            return v;
        default:
            return 'contains';
    }
}

function hasIds<C extends ColumnIds>(col: C): col is C & ColumnIdsFull {
    return col.ids !== null && col.values !== null && col.counts !== null;
}

function h<K extends keyof HTMLElementTagNameMap>(
    tag: K,
    cls?: string | null,
    text?: string | null
): HTMLElementTagNameMap[K] {
    const e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text !== null && text !== undefined) e.textContent = text;
    return e;
}

function icon(path: string): SVGSVGElement {
    const s = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    s.setAttribute('viewBox', '0 0 16 16');
    s.setAttribute('width', '14');
    s.setAttribute('height', '14');
    const p = document.createElementNS('http://www.w3.org/2000/svg', 'path');
    p.setAttribute('d', path);
    p.setAttribute('fill', 'none');
    p.setAttribute('stroke', 'currentColor');
    p.setAttribute('stroke-width', '1.5');
    p.setAttribute('stroke-linecap', 'round');
    p.setAttribute('stroke-linejoin', 'round');
    s.appendChild(p);
    return s;
}

/** Source row of the active cell (falls back to the view row). */
const srcRow = (a: GridActiveCell): number => (a.sr !== undefined && a.sr !== null ? a.sr : a.r);

const fmtNum = (n: number): string => formatGeneral(Math.round(n * 1e9) / 1e9);

interface ShellDom {
    bar: HTMLDivElement;
    nameBox: HTMLInputElement;
    fx: HTMLSpanElement;
    formula: HTMLTextAreaElement;
    fxExpand: HTMLButtonElement;
    fxBox: HTMLDivElement;
    btnFind: HTMLButtonElement;
    btnClear: HTMLButtonElement;
    gridWrap: HTMLDivElement;
    gridHost: HTMLDivElement;
    findBar: FindBar;
    tabs: HTMLDivElement;
    status: HTMLDivElement;
    msg: HTMLSpanElement;
    stats: HTMLSpanElement;
    zoomBox: HTMLSpanElement;
    zOut: HTMLButtonElement;
    zVal: HTMLSpanElement;
    zIn: HTMLButtonElement;
}

class Shell {
    app: HTMLElement;
    o: ShellOptions;
    provider: ShellProvider | null;
    view: ShellView;
    loading: boolean;
    grid: Grid;
    // DOM chrome, created by _build() and assigned in the constructor
    bar: HTMLDivElement;
    nameBox: HTMLInputElement;
    fx: HTMLSpanElement;
    formula: HTMLTextAreaElement;
    fxExpand: HTMLButtonElement;
    fxBox: HTMLDivElement;
    btnFind: HTMLButtonElement;
    btnClear: HTMLButtonElement;
    gridWrap: HTMLDivElement;
    gridHost: HTMLDivElement;
    findBar: FindBar;
    tabs: HTMLDivElement;
    status: HTMLDivElement;
    msg: HTMLSpanElement;
    stats: HTMLSpanElement;
    zoomBox: HTMLSpanElement;
    zOut: HTMLButtonElement;
    zVal: HTMLSpanElement;
    zIn: HTMLButtonElement;
    /** view row -> source row of the active sort / filter, or null */
    _order: Int32Array | null | undefined;
    _tasks: { [name: string]: Task | undefined };
    _colCache: Map<string, ColCacheEntry>;
    _statKey: string | null;
    _statText0: string | undefined;
    _statT: TimerHandle | undefined;
    _findT: TimerHandle | 0 | undefined;
    _vt: TimerHandle | undefined;
    _fxBase: string | undefined;
    _popClose: (() => void) | null | undefined;

    constructor(app: HTMLElement, opts?: ShellOptions) {
        this.app = app;
        this.o = opts || {};
        this.provider = null;
        this.view = { sort: null, filters: new Map() };
        this.loading = false;
        this._tasks = {};
        this._colCache = new Map();
        this._statKey = null;
        const dom = this._build();
        this.bar = dom.bar;
        this.nameBox = dom.nameBox;
        this.fx = dom.fx;
        this.formula = dom.formula;
        this.fxExpand = dom.fxExpand;
        this.fxBox = dom.fxBox;
        this.btnFind = dom.btnFind;
        this.btnClear = dom.btnClear;
        this.gridWrap = dom.gridWrap;
        this.gridHost = dom.gridHost;
        this.findBar = dom.findBar;
        this.tabs = dom.tabs;
        this.status = dom.status;
        this.msg = dom.msg;
        this.stats = dom.stats;
        this.zoomBox = dom.zoomBox;
        this.zOut = dom.zOut;
        this.zVal = dom.zVal;
        this.zIn = dom.zIn;
        this.grid = new Grid(this.gridHost, { provider: null, readOnly: !!this.o.readOnly });
        this._wire();
    }

    _build(): ShellDom {
        const a = this.app;
        a.classList.add('ox-app');
        const bar = h('div', 'ox-bar');
        const nameBox = h('input', 'ox-name');
        nameBox.spellcheck = false;
        nameBox.setAttribute('aria-label', 'Name box');
        const fx = h('span', 'ox-fx', 'fx');
        const formula = h('textarea', 'ox-formula');
        formula.rows = 1;
        formula.wrap = 'off';
        formula.spellcheck = false;
        formula.setAttribute('aria-label', 'Formula bar');
        const fxExpand = h('button', 'ox-btn ox-fxexp', '\u22EF');
        fxExpand.title = 'Multi-line value: click to expand';
        fxExpand.style.display = 'none';
        const fxBox = h('div', 'ox-fxbox');
        fxBox.append(formula, fxExpand);
        const btnFind = h('button', 'ox-btn');
        btnFind.title = 'Find (Ctrl+F)';
        btnFind.appendChild(icon('M7 12A5 5 0 1 0 7 2a5 5 0 0 0 0 10zM11 11l3.5 3.5'));
        const btnClear = h('button', 'ox-btn ox-clear', 'Clear sort & filter');
        btnClear.style.display = 'none';
        bar.append(nameBox, fx, fxBox, btnClear, btnFind);

        const gridWrap = h('div', 'ox-grid');
        const gridHost = h('div');
        gridWrap.appendChild(gridHost);
        const findBar = this._buildFind();
        gridWrap.appendChild(findBar.el);

        const tabs = h('div', 'ox-tabs');
        if (!this.o.tabs) tabs.style.display = 'none';
        const status = h('div', 'ox-status');
        const msg = h('span', 'ox-msg', 'Ready');
        const stats = h('span', 'ox-stats');
        const zoomBox = h('span', 'ox-zoom');
        const zOut = h('button', 'ox-btn ox-z', '-');
        const zVal = h('span', 'ox-zval', '100%');
        const zIn = h('button', 'ox-btn ox-z', '+');
        zoomBox.append(zOut, zVal, zIn);
        status.append(msg, stats, zoomBox);
        a.append(bar, gridWrap, tabs, status);
        return {
            bar,
            nameBox,
            fx,
            formula,
            fxExpand,
            fxBox,
            btnFind,
            btnClear,
            gridWrap,
            gridHost,
            findBar,
            tabs,
            status,
            msg,
            stats,
            zoomBox,
            zOut,
            zVal,
            zIn
        };
    }

    _buildFind(): FindBar {
        const el = h('div', 'ox-find');
        el.style.display = 'none';
        const input = h('input', 'ox-find-in');
        input.placeholder = 'Find';
        input.spellcheck = false;
        const count = h('span', 'ox-find-count');
        const prev = h('button', 'ox-btn');
        prev.title = 'Previous (Shift+Enter)';
        prev.appendChild(icon('M3 10l5-5 5 5'));
        const next = h('button', 'ox-btn');
        next.title = 'Next (Enter)';
        next.appendChild(icon('M3 6l5 5 5-5'));
        const mc = h('button', 'ox-btn ox-tog', 'Aa');
        mc.title = 'Match case';
        const wc = h('button', 'ox-btn ox-tog', '[ ]');
        wc.title = 'Match entire cell';
        const close = h('button', 'ox-btn');
        close.title = 'Close (Esc)';
        close.appendChild(icon('M4 4l8 8M12 4l-8 8'));
        el.append(input, count, prev, next, mc, wc, close);
        return { el, input, count, prev, next, mc, wc, close };
    }

    _wire(): void {
        const g = this.grid;
        g.on('select', e => this._onSelect(e));
        g.on('edit', e => {
            this._colCache.clear();
            this._statKey = null;
            if (this.o.onEdit) this.o.onEdit(e.edits, e.source);
        });
        g.on('resizeCol', e => {
            if (this.provider && this.provider.setColWidth) this.provider.setColWidth(e.c, e.width);
            this._viewState();
        });
        g.on('resizeRow', e => {
            if (this.provider && this.provider.setRowHeight) this.provider.setRowHeight(e.r, e.height);
        });
        g.on('scroll', () => this._viewState());
        g.on('zoom', z => {
            this.zVal.textContent = Math.round((z && z.zoom ? z.zoom : g.getZoom()) * 100) + '%';
            this._viewState();
        });
        g.on('findrequest', () => this.openFind());
        g.on('gotorequest', () => {
            this.nameBox.focus();
            this.nameBox.select();
        });
        g.on('find', r => this._findCount(r));
        g.on('copyTooLarge', e =>
            this.setMessage(
                'Selection is too large to copy' + (e && e.cells ? ' (' + e.cells.toLocaleString() + ' cells)' : '')
            )
        );
        g.on('link', e => {
            if (this.o.onLink) this.o.onLink(e);
        });
        g.on('contextmenu', e => this._onContext(e));

        this.zOut.addEventListener('click', () => {
            g.setZoom(g.getZoom() / 1.15);
        });
        this.zIn.addEventListener('click', () => {
            g.setZoom(g.getZoom() * 1.15);
        });
        this.zVal.addEventListener('click', () => g.setZoom(1));
        this.btnFind.addEventListener('click', () => this.openFind());
        this.btnClear.addEventListener('click', () => this.clearView());

        this.nameBox.addEventListener('keydown', e => {
            if (e.key === 'Enter') {
                e.preventDefault();
                if (!this._goto(this.nameBox.value)) {
                    this.nameBox.classList.add('bad');
                    setTimeout(() => this.nameBox.classList.remove('bad'), 400);
                } else g.focus();
            } else if (e.key === 'Escape') {
                this._updateName();
                g.focus();
            }
        });
        this.formula.addEventListener('keydown', e => {
            if (e.key === 'Enter') {
                e.preventDefault();
                if (e.altKey || e.shiftKey) this._fxNewline();
                else this._commitFormula();
            } else if (e.key === 'Escape') {
                e.preventDefault();
                this._updateFormula();
                g.focus();
            }
        });
        this.formula.addEventListener('focus', () => {
            if (this._fxMulti()) this.fxBox.classList.add('open');
        });
        this.fxExpand.addEventListener('click', () => {
            this.fxBox.classList.toggle('open');
            this.formula.focus();
        });

        const f = this.findBar;
        f.input.addEventListener('input', () => {
            clearTimeout(this._findT);
            this._findT = setTimeout(() => {
                this._findT = 0;
                this._doFind(true);
            }, 160);
        });
        f.input.addEventListener('keydown', e => {
            if (e.key === 'Enter') {
                e.preventDefault();
                if (this._findT) {
                    clearTimeout(this._findT);
                    this._findT = 0;
                    this._doFind(true);
                    return;
                }
                const r = e.shiftKey ? g.findPrev() : g.findNext();
                this._findCount(r);
            } else if (e.key === 'Escape') {
                e.preventDefault();
                this.closeFind();
            }
        });
        f.next.addEventListener('click', () => this._findCount(g.findNext()));
        f.prev.addEventListener('click', () => this._findCount(g.findPrev()));
        f.close.addEventListener('click', () => this.closeFind());
        f.mc.addEventListener('click', () => {
            f.mc.classList.toggle('on');
            this._doFind(true);
        });
        f.wc.addEventListener('click', () => {
            f.wc.classList.toggle('on');
            this._doFind(true);
        });

        document.addEventListener('keydown', e => {
            const mod = e.metaKey || e.ctrlKey;
            if (mod && e.key.toLowerCase() === 'f' && !e.shiftKey) {
                e.preventDefault();
                this.openFind();
            }
        });
        const mo = new MutationObserver(() => {
            g.themeChanged();
            if (this.provider && this.provider.setDark) this.provider.setDark(this.isDark());
            g.refresh();
        });
        mo.observe(document.body, { attributes: true, attributeFilter: ['class', 'style'] });
        mo.observe(document.documentElement, { attributes: true, attributeFilter: ['style', 'class'] });
    }

    /** The provider; the view / stats pipeline only runs with one. */
    _need(): ShellProvider {
        const p = this.provider;
        if (!p) throw new Error('Shell has no provider');
        return p;
    }

    isDark(): boolean {
        return !document.body.classList.contains('vscode-light');
    }

    setProvider(p: ShellProvider): void {
        this._cancelAll();
        this.provider = p;
        this._order = null;
        this._colCache.clear();
        this._statKey = null;
        this._closePopup();
        this.view = { sort: null, filters: new Map() };
        if (!p._shellColName) {
            const orig = p.colName ? p.colName.bind(p) : null;
            p._shellColName = true;
            p.colName = (c: number): string => {
                let base = orig ? orig(c) : colName(c);
                const v = this.view;
                if (v.sort && v.sort.c === c) base += v.sort.dir > 0 ? ' ▲' : ' ▼';
                if (v.filters.has(c)) base += ' ▾*';
                return base;
            };
        }
        this.grid.setView(null);
        this.grid.setProvider(p);
        this._viewState();
        this._updateAll();
        this.btnClear.style.display = 'none';
    }

    setMessage(t: string): void {
        this.msg.textContent = t;
    }

    // ---- background tasks ----
    _start<T>(name: string, gen: TaskGen<T>, opts?: TaskOptions<T>): Task {
        this._cancel(name);
        const o: TaskOptions<T> = opts || {};
        const slot: { t?: Task } = {};
        const finish = () => {
            if (slot.t && this._tasks[name] === slot.t) delete this._tasks[name];
        };
        const t = (slot.t = runTask(
            gen,
            Object.assign({}, o, {
                onDone: (v: T) => {
                    finish();
                    if (o.onDone) o.onDone(v);
                },
                onError: (e: unknown) => {
                    finish();
                    if (o.onError) o.onError(e);
                    else if (root.console) root.console.error(e);
                }
            })
        ));
        if (!t.done) this._tasks[name] = t;
        return t;
    }
    _cancel(name: string): void {
        const t = this._tasks[name];
        if (t) {
            t.cancel();
            delete this._tasks[name];
        }
    }
    _cancelAll(): void {
        for (const k of Object.keys(this._tasks)) this._cancel(k);
        clearTimeout(this._statT);
        clearTimeout(this._findT);
        this._findT = 0;
    }
    busy(): boolean {
        return Object.keys(this._tasks).length > 0;
    }

    // ---- selection UI ----
    _onSelect(e: GridSelection): void {
        this._updateName();
        this._updateFormula();
        this._scheduleStats(e && e.dragging);
        this._viewState();
    }
    _updateAll(): void {
        this._colCache.clear();
        this._statKey = null;
        this._updateName();
        this._updateFormula();
        this._scheduleStats(false);
    }

    _updateName(): void {
        const s = this.grid.getSelection();
        if (!s || !s.active) return;
        const rg = s.ranges[s.ranges.length - 1];
        const single = rg.r1 === rg.r2 && rg.c1 === rg.c2;
        if (single) this.nameBox.value = cellRef(srcRow(s.active), s.active.c);
        else this.nameBox.value = rg.r2 - rg.r1 + 1 + 'R x ' + (rg.c2 - rg.c1 + 1) + 'C';
    }

    _updateFormula(): void {
        const s = this.grid.getSelection();
        if (!s || !s.active || !this.provider) {
            this.formula.value = '';
            return;
        }
        const sr = srcRow(s.active);
        this.formula.value = this.provider.getRaw(sr, s.active.c) || '';
        this._fxBase = this.formula.value;
        this.formula.readOnly = !!this.o.readOnly;
        const multi = this._fxMulti();
        this.fxExpand.style.display = multi ? '' : 'none';
        if (!multi) this.fxBox.classList.remove('open');
    }

    _fxMulti(): boolean {
        return this.formula.value.indexOf('\n') >= 0;
    }

    _fxNewline(): void {
        const f = this.formula;
        if (f.readOnly) return;
        const a = f.selectionStart,
            b = f.selectionEnd;
        f.value = f.value.slice(0, a) + '\n' + f.value.slice(b);
        f.selectionStart = f.selectionEnd = a + 1;
        this.fxExpand.style.display = '';
        this.fxBox.classList.add('open');
    }

    _closePopup(): void {
        if (this._popClose) this._popClose();
    }

    _commitFormula(): void {
        if (this.o.readOnly) return;
        const s = this.grid.getSelection();
        if (!s || !s.active) return;
        const sr = srcRow(s.active);
        // Untouched text is not re-committed (it would also round-trip multi-line cells through the input).
        if (this.formula.value !== this._fxBase && this.o.onEdit)
            this.o.onEdit([{ r: sr, c: s.active.c, text: this.formula.value }], 'formula');
        this.grid.refresh();
        this.grid.focus();
        this._updateFormula();
    }

    _goto(text: unknown): boolean {
        const t = String(text).trim();
        if (!t) return false;
        const ref = parseRef(t);
        const rg = parseRange(t) || (ref && { r1: ref.r, c1: ref.c, r2: ref.r, c2: ref.c });
        if (!rg) return false;
        this.grid.select(rg.r1, rg.c1, rg.r2, rg.c2);
        this.grid.scrollTo(rg.r1, rg.c1);
        return true;
    }

    _selCells(s: GridSelection): number {
        let n = 0;
        for (const rg of s.ranges) n += (rg.r2 - rg.r1 + 1) * (rg.c2 - rg.c1 + 1);
        return n;
    }

    _scheduleStats(dragging: boolean | undefined): void {
        clearTimeout(this._statT);
        const p = this.provider,
            s = this.grid.getSelection();
        if (!p || !s) {
            this._cancel('stats');
            this.stats.textContent = '';
            return;
        }
        const single = s.ranges.length === 1 && s.ranges[0].r1 === s.ranges[0].r2 && s.ranges[0].c1 === s.ranges[0].c2;
        if (single) {
            this._cancel('stats');
            this.stats.textContent = '';
            this._statKey = null;
            return;
        }
        if (dragging) return; // the grid sends a final select with dragging:false
        if (this._selCells(s) <= SMALL_STATS) {
            this._runStats(s);
            return;
        }
        this._statT = setTimeout(() => this._runStats(this.grid.getSelection()), 120);
    }

    _updateStats(): void {
        this._scheduleStats(false);
    }

    _statText(a: StatsAcc, pct?: number): string {
        const parts = [];
        if (a.nums) parts.push('Average: ' + fmtNum(a.sum / a.nums));
        parts.push('Count: ' + a.cnt);
        if (a.nums) {
            parts.push('Sum: ' + fmtNum(a.sum));
            parts.push('Min: ' + fmtNum(a.min));
            parts.push('Max: ' + fmtNum(a.max));
        }
        return (
            parts.join('    ') + (pct === undefined ? '' : '    (partial, calculating ' + Math.round(pct * 100) + '%)')
        );
    }

    _runStats(s: GridSelection): void {
        const p = this.provider;
        if (!p || !s) return;
        const key =
            s.ranges.map(g => g.r1 + ',' + g.c1 + ',' + g.r2 + ',' + g.c2).join(';') +
            '|' +
            (this._order ? this._order.length : -1) +
            '|' +
            p.rowCount;
        if (key === this._statKey && this._statText0 !== undefined && !this._tasks.stats) {
            this.stats.textContent = this._statText0;
            return;
        }
        const order = this._order,
            total = { cnt: 0, nums: 0, sum: 0, min: Infinity, max: -Infinity };
        const ranges = s.ranges.map(rg => ({
            r1: rg.r1,
            c1: rg.c1,
            r2: Math.min(rg.r2, (order ? order.length : p.rowCount) - 1),
            c2: Math.min(rg.c2, p.colCount - 1)
        }));
        const genericStats = (acc: StatsAcc, ord: ArrayLike<number> | null | undefined, rg: Range): StatsGen =>
            this._genericStats(acc, ord, rg);
        const cur: { acc: StatsAcc | null } = { acc: null };
        const live = (): StatsAcc => {
            const t = { cnt: total.cnt, nums: total.nums, sum: total.sum, min: total.min, max: total.max };
            if (cur.acc) {
                t.cnt += cur.acc.cnt;
                t.nums += cur.acc.nums;
                t.sum += cur.acc.sum;
                if (cur.acc.min < t.min) t.min = cur.acc.min;
                if (cur.acc.max > t.max) t.max = cur.acc.max;
            }
            return t;
        };
        const gen = function* (): Generator<number, void, undefined> {
            for (let i = 0; i < ranges.length; i++) {
                const rg = ranges[i];
                if (rg.r2 < rg.r1 || rg.c2 < rg.c1) continue;
                cur.acc = { cnt: 0, nums: 0, sum: 0, min: Infinity, max: -Infinity };
                const it = p.rangeStats
                    ? p.rangeStats(cur.acc, order, rg.r1, rg.r2, rg.c1, rg.c2)
                    : genericStats(cur.acc, order, rg);
                for (let r = it.next(); !r.done; r = it.next())
                    yield (i + (typeof r.value === 'number' ? r.value : 0)) / ranges.length;
                total.cnt += cur.acc.cnt;
                total.nums += cur.acc.nums;
                total.sum += cur.acc.sum;
                if (cur.acc.min < total.min) total.min = cur.acc.min;
                if (cur.acc.max > total.max) total.max = cur.acc.max;
                cur.acc = null;
            }
        };
        this._start('stats', gen(), {
            onProgress: pct => {
                this.stats.textContent = this._statText(live(), pct);
            },
            onDone: () => {
                this._statText0 = this._statText(total);
                this._statKey = key;
                this.stats.textContent = this._statText0;
            }
        });
    }

    // Provider-agnostic fallback (sliced): forEachCell per row slab, else getCell per cell.
    *_genericStats(a: StatsAcc, order: ArrayLike<number> | null | undefined, rg: Range): StatsGen {
        const p = this._need(),
            nc = rg.c2 - rg.c1 + 1,
            nr = rg.r2 - rg.r1 + 1;
        if (!order && p.forEachCell) {
            const slab = Math.max(1, Math.floor(100000 / nc));
            const cb = (cell: SlotCell): void => {
                if (cell.t === 'z' || cell.v === null || cell.v === '') return;
                a.cnt++;
                if (cell.t === 'n') {
                    a.nums++;
                    a.sum += cell.v;
                    if (cell.v < a.min) a.min = cell.v;
                    if (cell.v > a.max) a.max = cell.v;
                }
            };
            for (let r = rg.r1; r <= rg.r2; r += slab) {
                p.forEachCell(r, rg.c1, Math.min(rg.r2, r + slab - 1), rg.c2, cb);
                yield (r - rg.r1) / nr;
            }
            return a;
        }
        let work = 0;
        for (let r = rg.r1; r <= rg.r2; r++) {
            const sr = order ? order[r] : r;
            for (let c = rg.c1; c <= rg.c2; c++) {
                const cv = p.getCell(sr, c);
                if (!cv || cv.text === '') continue;
                a.cnt++;
                if (cv.kind === 'number') {
                    const n = Number(p.getRaw(sr, c));
                    if (n === n) {
                        a.nums++;
                        a.sum += n;
                        if (n < a.min) a.min = n;
                        if (n > a.max) a.max = n;
                    }
                }
            }
            work += nc;
            if (work > 30000) {
                work = 0;
                yield (r - rg.r1) / nr;
            }
        }
        return a;
    }

    // ---- find ----
    openFind(): void {
        const f = this.findBar;
        f.el.style.display = 'flex';
        const s = this.grid.getSelection();
        if (s && s.active && this.provider && !f.input.value) {
            const raw = this.provider.getRaw(srcRow(s.active), s.active.c);
            if (raw && raw.indexOf('\n') < 0 && raw.length < 80) f.input.value = raw;
        }
        f.input.focus();
        f.input.select();
        this._doFind(false);
    }
    closeFind(): void {
        this.findBar.el.style.display = 'none';
        this.grid.clearFind();
        this.findBar.count.textContent = '';
        this.grid.focus();
    }
    _doFind(_jump: boolean): void {
        const f = this.findBar,
            q = f.input.value;
        if (!q) {
            this.grid.clearFind();
            f.count.textContent = '';
            return;
        }
        this._findCount(
            this.grid.find(q, {
                matchCase: f.mc.classList.contains('on'),
                wholeCell: f.wc.classList.contains('on')
            })
        );
    }
    _findCount(r: GridFindState | null | undefined): void {
        if (!r) return;
        const more = r.done === false ? ' (searching...)' : r.truncated ? ' (limit reached)' : '';
        this.findBar.count.textContent = r.count
            ? r.index + ' of ' + r.count + more
            : r.done === false
              ? 'Searching...'
              : 'No results';
        this.findBar.count.classList.toggle('none', !r.count);
    }

    // ---- sort / filter (view only) ----
    _hdr(): number {
        return this.o.headerRows ? this.o.headerRows() : 0;
    }

    // Per-row ids into distinct values of column c (display text, or raw when raw=true); ids null if > MAX_DISTINCT values.
    *_columnIds(c: number, raw: boolean): Generator<number, ColumnIds, undefined> {
        const p = this._need(),
            hdr = this._hdr();
        if (p.columnIds) return yield* p.columnIds(c, MAX_DISTINCT, raw);
        const key = c + (raw ? 'r' : 'd'),
            ver = p.dataVersion === undefined ? -1 : p.dataVersion;
        const hit = this._colCache.get(key);
        if (hit && hit.ver === ver) return hit.res;
        const total = p.rowCount,
            ids = new Int32Array(total).fill(-1),
            map = new Map<string, number>(),
            values: string[] = [],
            cnt: number[] = [];
        let work = 0,
            truncated = false;
        for (let r = hdr; r < total; r++) {
            let t: string;
            if (raw) t = p.getRaw(r, c) || '';
            else {
                const cv = p.getCell(r, c);
                t = cv && cv.text !== undefined && cv.text !== null ? cv.text : '';
            }
            let k = map.get(t);
            if (k === undefined) {
                if (values.length >= MAX_DISTINCT) {
                    truncated = true;
                    break;
                }
                k = values.length;
                values.push(t);
                cnt.push(0);
                map.set(t, k);
            }
            ids[r] = k;
            cnt[k]++;
            if (++work > 20000) {
                work = 0;
                yield r / total;
            }
        }
        const res: ColumnIds = truncated
            ? { ids: null, values: null, counts: null, truncated: true }
            : { ids, values, counts: Int32Array.from(cnt), truncated: false };
        this._colCache.set(key, { ver, res });
        if (this._colCache.size > 6) {
            const oldest = this._colCache.keys().next().value;
            if (oldest !== undefined) this._colCache.delete(oldest);
        }
        return res;
    }

    // Row subset (Int32Array of source rows) passing one column filter; `rows` null = all body rows.
    *_applyFilter(rows: Int32Array | null, c: number, spec: ColumnFilter): Generator<number, Int32Array, undefined> {
        const p = this._need(),
            hdr = this._hdr(),
            total = p.rowCount;
        const isText = spec.kind === 'text';
        const col = yield* this._columnIds(c, isText);
        const n = rows ? rows.length : Math.max(0, total - hdr);
        const out = new Int32Array(n);
        let m = 0;
        if (hasIds(col)) {
            const test = isText ? matchText(spec) : noMatch,
                allowed = new Uint8Array(col.values.length);
            for (let i = 0; i < allowed.length; i++)
                allowed[i] = isText ? (test(col.values[i]) ? 1 : 0) : spec.set.has(col.values[i]) ? 1 : 0;
            const ids = col.ids;
            for (let i = 0; i < n; i++) {
                const r = rows ? rows[i] : hdr + i;
                if (allowed[ids[r]] === 1) out[m++] = r;
                if ((i & 0x3ffff) === 0x3ffff) yield i / n;
            }
        } else {
            const test = isText ? matchText(spec) : noMatch;
            for (let i = 0; i < n; i++) {
                const r = rows ? rows[i] : hdr + i;
                let ok: boolean;
                if (isText) ok = test(p.getRaw(r, c) || '');
                else {
                    const cv = p.getCell(r, c);
                    ok = spec.set.has(cv && cv.text !== undefined && cv.text !== null ? cv.text : '');
                }
                if (ok) out[m++] = r;
                if ((i & 0x7fff) === 0x7fff) yield i / n;
            }
        }
        return out.slice(0, m);
    }

    // Generic sort keys (same shape as CsvProvider.sortKeyColumn).
    *_sortKeys(c: number): Generator<number, SortKeys, undefined> {
        const p = this._need();
        if (p.sortKeyColumn) return yield* p.sortKeyColumn(c);
        const R = p.rowCount,
            cat = new Uint8Array(R).fill(2),
            hi = new Uint32Array(R),
            lo = new Uint32Array(R),
            strs: string[] = [];
        const hiI = LE ? 1 : 0,
            loI = LE ? 0 : 1;
        let work = 0;
        for (let r = 0; r < R; r++) {
            const raw = p.getRaw(r, c);
            if (raw) {
                const f = raw.charCodeAt(0);
                if (((f >= 48 && f <= 57) || f === 45 || f === 43 || f === 46) && raw.length < 64 && NUM_RE.test(raw)) {
                    const x = Number(raw);
                    f64[0] = x === 0 ? 0 : x;
                    let h = u32[hiI],
                        l = u32[loI];
                    if (h & 0x80000000) {
                        h = ~h;
                        l = ~l;
                    } else h |= 0x80000000;
                    cat[r] = 0;
                    hi[r] = h >>> 0;
                    lo[r] = l >>> 0;
                } else {
                    cat[r] = 1;
                    lo[r] = strs.length;
                    strs.push(raw.toLowerCase());
                }
            }
            if (++work > 20000) {
                work = 0;
                yield (0.6 * r) / R;
            }
        }
        return { cat, hi, lo, strs };
    }

    *_sortRows(rows: Int32Array, c: number, dir: number): Generator<number, Int32Array, undefined> {
        const keys = yield* this._sortKeys(c),
            cat = keys.cat,
            hi = keys.hi,
            lo = keys.lo,
            strs = keys.strs;
        if (strs.length) {
            const idx = new Int32Array(strs.length);
            for (let i = 0; i < idx.length; i++) idx[i] = i;
            yield* mergeSortGen(idx, (a, b) => strs[a] < strs[b]);
            const rank = new Int32Array(idx.length);
            let rk = 0;
            for (let i = 0; i < idx.length; i++) {
                if (i && strs[idx[i]] !== strs[idx[i - 1]]) rk = i;
                rank[idx[i]] = rk;
            }
            for (let i = 0; i < rows.length; i++) {
                const r = rows[i];
                if (cat[r] === 1) lo[r] = rank[lo[r]];
                if ((i & 0x3ffff) === 0x3ffff) yield 0.7;
            }
        }
        if (dir < 0) {
            for (let i = 0; i < rows.length; i++) {
                const r = rows[i];
                if (cat[r] !== 2) {
                    hi[r] = ~hi[r];
                    lo[r] = ~lo[r];
                }
                if ((i & 0x3ffff) === 0x3ffff) yield 0.75;
            }
        }
        return yield* radixRows(rows, cat, hi, lo);
    }

    *_viewGen(): Generator<number, Int32Array, undefined> {
        const p = this._need(),
            v = this.view,
            hdr = this._hdr(),
            total = p.rowCount;
        let rows: Int32Array | null = null;
        const filt = [...v.filters.entries()],
            steps = filt.length + (v.sort ? 1 : 0);
        let stepNo = 0;
        for (const [c, spec] of filt) {
            const it = this._applyFilter(rows, c, spec);
            for (let r = it.next(); ; r = it.next()) {
                if (r.done) {
                    rows = r.value;
                    break;
                }
                yield (stepNo + (r.value || 0)) / steps;
            }
            stepNo++;
        }
        if (!rows) {
            rows = new Int32Array(Math.max(0, total - hdr));
            for (let i = 0; i < rows.length; i++) rows[i] = hdr + i;
        }
        if (v.sort) {
            const it = this._sortRows(rows, v.sort.c, v.sort.dir);
            for (let r = it.next(); ; r = it.next()) {
                if (r.done) {
                    rows = r.value;
                    break;
                }
                yield (stepNo + (r.value || 0)) / steps;
            }
        }
        const out = new Int32Array(Math.min(hdr, total) + rows.length);
        for (let i = 0; i < hdr && i < total; i++) out[i] = i;
        out.set(rows, Math.min(hdr, total));
        return out;
    }

    recomputeView(): void {
        const p = this.provider;
        if (!p) return;
        this._cancel('view');
        const v = this.view;
        if (!v.sort && !v.filters.size) {
            this._order = null;
            this.grid.setView(null);
            this.btnClear.style.display = 'none';
            this._updateAll();
            this._status();
            return;
        }
        const label =
            v.sort && !v.filters.size ? 'Sorting' : v.filters.size && !v.sort ? 'Filtering' : 'Sorting and filtering';
        this.btnClear.style.display = '';
        this.setMessage(label + '...');
        this._start('view', this._viewGen(), {
            budget: 10,
            onProgress: pct => this.setMessage(label + ' ' + Math.round(pct * 100) + '%...'),
            onDone: order => {
                this._order = order;
                this.grid.setView({ order });
                this._status();
                this._updateAll();
            },
            onError: e => {
                this.setMessage('Sort/filter failed: ' + (e instanceof Error ? e.message : undefined));
                if (root.console) root.console.error(e);
            }
        });
    }

    _status(): void {
        const v = this.view,
            parts = [],
            hdr = this._hdr();
        if (v.sort) parts.push('Sorted by ' + colName(v.sort.c) + (v.sort.dir > 0 ? ' ascending' : ' descending'));
        if (v.filters.size)
            parts.push(
                'Filtered: ' +
                    (this._order ? Math.max(0, this._order.length - hdr) : 0) +
                    ' of ' +
                    Math.max(0, this._need().rowCount - hdr) +
                    ' rows'
            );
        this.setMessage(parts.length ? parts.join('  |  ') + '  (view only, file unchanged)' : 'Ready');
    }

    sortBy(c: number, dir: number): void {
        if (this._guardLoading()) return;
        this.view.sort = { c, dir };
        this.recomputeView();
    }
    clearView(): void {
        this.view = { sort: null, filters: new Map() };
        this.recomputeView();
    }

    _guardLoading(): boolean {
        if (!this.loading && !(this.provider && this.provider.partial)) return false;
        this.setMessage('Still loading: sort and filter are available when the file is fully loaded');
        return true;
    }

    filterColumn(c: number, x: number, y: number): void {
        if (this._guardLoading()) return;
        const cur = this.view.filters.get(c);
        this._closePopup();
        const pop = h('div', 'ox-pop');
        pop.append(h('div', 'ox-pop-title', 'Filter column ' + colName(c)));
        const mode = h('select', 'ox-pop-mode');
        [
            ['values', 'Choose values'],
            ['contains', 'Text contains'],
            ['equals', 'Text equals'],
            ['starts', 'Text starts with'],
            ['ends', 'Text ends with'],
            ['notcontains', 'Text does not contain']
        ].forEach(m => {
            const o = h('option', null, m[1]);
            o.value = m[0];
            mode.appendChild(o);
        });
        const search = h('input', 'ox-pop-search');
        search.spellcheck = false;
        const info = h('div', 'ox-pop-info', 'Scanning column...');
        const listBox = h('div', 'ox-pop-list'),
            spacer = h('div', 'ox-pop-spacer'),
            win = h('div', 'ox-pop-win');
        const all = h('label', 'ox-pop-row ox-pop-all');
        const allCb = h('input');
        allCb.type = 'checkbox';
        all.append(allCb, h('span', null, '(Select all)'));
        listBox.append(spacer);
        spacer.appendChild(win);
        const btns = h('div', 'ox-pop-btns');
        const ok = h('button', 'ox-btn ox-primary', 'Apply');
        const no = h('button', 'ox-btn', 'Cancel');
        btns.append(no, ok);
        pop.append(mode, search, info, all, listBox, btns);

        const ROW = 22,
            VIS = 12;
        let col: ScannedColumn | null = null,
            order: Int32Array | null | undefined = null,
            fl: Int32Array | null = null,
            sel: Uint8Array | null = null;
        const curText = cur && cur.kind === 'text' ? cur : null;
        mode.value = curText ? curText.op : 'values';
        search.value = curText ? curText.q : '';
        const isValues = () => mode.value === 'values';
        const render = (): void => {
            if (!col || !hasIds(col) || !fl || !sel) return;
            const cl = col,
                flv = fl,
                selv = sel;
            const top = Math.floor(listBox.scrollTop / ROW),
                n = flv.length;
            win.style.transform = 'translateY(' + top * ROW + 'px)';
            win.textContent = '';
            for (let i = top; i < Math.min(n, top + VIS + 2); i++) {
                const vi = flv[i],
                    l = h('label', 'ox-pop-row');
                const cb = h('input');
                cb.type = 'checkbox';
                cb.checked = selv[vi] === 1;
                cb.addEventListener('change', () => {
                    selv[vi] = cb.checked ? 1 : 0;
                    syncAll();
                });
                const v = cl.values[vi];
                l.append(cb, h('span', null, v === '' ? '(Blanks)' : v), h('em', null, String(cl.counts[vi])));
                win.appendChild(l);
            }
        };
        const syncAll = (): void => {
            if (!fl || !sel) return;
            let every = fl.length > 0;
            for (let i = 0; i < fl.length; i++)
                if (sel[fl[i]] !== 1) {
                    every = false;
                    break;
                }
            allCb.checked = every;
        };
        const refilter = (): void => {
            if (!col || !hasIds(col) || !order) return;
            const cl = col,
                ord = order;
            const q = search.value.toLowerCase();
            const keep: number[] = [];
            for (let i = 0; i < ord.length; i++) {
                const vi = ord[i];
                if (!q || cl.values[vi].toLowerCase().indexOf(q) >= 0) keep.push(vi);
            }
            fl = Int32Array.from(keep);
            spacer.style.height = fl.length * ROW + 'px';
            listBox.scrollTop = 0;
            info.textContent = fl.length.toLocaleString() + ' of ' + cl.values.length.toLocaleString() + ' values';
            syncAll();
            render();
        };
        const layout = (): void => {
            const vals = isValues() && col && hasIds(col);
            listBox.style.display = all.style.display = vals ? '' : 'none';
            search.placeholder = isValues() ? 'Search' : 'Text';
            if (col && !hasIds(col)) mode.options[0].disabled = true;
            if (!isValues())
                info.textContent =
                    col && !hasIds(col)
                        ? 'More than ' + MAX_DISTINCT.toLocaleString() + ' distinct values: filter by text'
                        : 'Rows whose value matches the text (case-insensitive)';
            else if (col && hasIds(col))
                info.textContent =
                    (fl ? fl.length : 0).toLocaleString() + ' of ' + col.values.length.toLocaleString() + ' values';
            if (vals) refilter();
        };
        listBox.addEventListener('scroll', render);
        allCb.addEventListener('change', () => {
            if (!fl || !sel) return;
            for (let i = 0; i < fl.length; i++) sel[fl[i]] = allCb.checked ? 1 : 0;
            render();
        });
        mode.addEventListener('change', () => {
            layout();
            search.focus();
        });
        search.addEventListener('input', () => {
            if (isValues() && col && hasIds(col)) refilter();
        });
        const close: () => void = () => {
            this._cancel('filterScan');
            pop.remove();
            document.removeEventListener('mousedown', off, true);
            if (this._popClose === close) this._popClose = null;
            this.grid.focus();
        };
        this._popClose = close;
        const off = (e: MouseEvent): void => {
            if (!(e.target instanceof Node) || !pop.contains(e.target)) close();
        };
        no.addEventListener('click', close);
        ok.addEventListener('click', () => {
            if (!isValues()) {
                const q = search.value;
                if (q === '') this.view.filters.delete(c);
                else this.view.filters.set(c, { kind: 'text', op: textOp(mode.value), q });
            } else if (col && hasIds(col) && sel) {
                let count = 0;
                for (let i = 0; i < sel.length; i++) count += sel[i];
                if (count === sel.length) this.view.filters.delete(c);
                else {
                    const set = new Set<string>();
                    for (let i = 0; i < sel.length; i++) if (sel[i]) set.add(col.values[i]);
                    this.view.filters.set(c, { kind: 'set', set });
                }
            } else return;
            close();
            this.recomputeView();
        });
        pop.addEventListener('keydown', e => {
            if (e.key === 'Escape') close();
            else if (e.key === 'Enter' && (e.target === search || e.target === mode || e.target === pop)) {
                e.preventDefault();
                ok.click();
            }
        });
        document.body.appendChild(pop);
        pop.style.left = Math.max(4, Math.min(x, innerWidth - pop.offsetWidth - 8)) + 'px';
        pop.style.top = Math.max(4, Math.min(y, innerHeight - pop.offsetHeight - 8)) + 'px';
        document.addEventListener('mousedown', off, true);
        layout();
        search.focus();
        const columnIds = (): Generator<number, ColumnIds, undefined> => this._columnIds(c, false);
        function* scan(): Generator<number, ScannedColumn, undefined> {
            const it = columnIds();
            let res: ScannedColumn;
            for (let r = it.next(); ; r = it.next()) {
                if (r.done) {
                    res = r.value;
                    break;
                }
                yield r.value;
            }
            if (hasIds(res)) {
                const nums: number[] = [],
                    strs: number[] = [];
                for (let i = 0; i < res.values.length; i++) {
                    const t = res.values[i];
                    (t !== '' && NUM_RE.test(t) ? nums : strs).push(i);
                }
                const vals = res.values,
                    coll = typeof Intl !== 'undefined' ? new Intl.Collator(undefined, { numeric: true }) : null;
                const ni = Int32Array.from(nums),
                    si = Int32Array.from(strs);
                yield* mergeSortGen(ni, (a, b) => Number(vals[a]) < Number(vals[b]));
                yield* mergeSortGen(
                    si,
                    coll ? (a, b) => coll.compare(vals[a], vals[b]) < 0 : (a, b) => vals[a] < vals[b]
                );
                const ord = new Int32Array(vals.length);
                let k = 0;
                const blank = si.length && vals[si[0]] === '' ? 1 : 0;
                if (blank) ord[k++] = si[0];
                for (let i = 0; i < ni.length; i++) ord[k++] = ni[i];
                for (let i = blank; i < si.length; i++) ord[k++] = si[i];
                res = Object.assign({}, res, { order: ord });
            }
            return res;
        }
        this._start('filterScan', scan(), {
            budget: 8,
            onProgress: pct => {
                info.textContent = 'Scanning column ' + Math.round(pct * 100) + '%...';
            },
            onDone: (res: ScannedColumn) => {
                col = res;
                if (hasIds(col)) {
                    order = col.order;
                    sel = new Uint8Array(col.values.length);
                    const set = cur && cur.kind === 'set' ? cur.set : null;
                    for (let i = 0; i < sel.length; i++) sel[i] = set ? (set.has(col.values[i]) ? 1 : 0) : 1;
                } else if (isValues()) mode.value = 'contains';
                layout();
            }
        });
    }

    _onContext(e: GridContextMenuEvent): void {
        if (e.zone !== 'colhdr' || !this.provider) return;
        e.handled = true;
        const c = e.c,
            v = this.view;
        const items: GridMenuItem[] = [
            ['Sort A to Z', '', () => this.sortBy(c, 1)],
            ['Sort Z to A', '', () => this.sortBy(c, -1)],
            ['Filter...', '', () => this.filterColumn(c, e.clientX, e.clientY)],
            ['Clear sort & filter', '', () => this.clearView(), !(v.sort || v.filters.size)],
            null,
            ['Autosize column', '', () => this.grid.autosizeColumn(c)]
        ];
        this.grid._showMenu(items, e.clientX, e.clientY);
    }

    // ---- persisted view state ----
    getViewState(): ShellViewState {
        const sel = this.grid.getSelection();
        return {
            zoom: this.grid.getZoom(),
            scroll: this.grid.getScroll(),
            active: sel && sel.active ? { r: srcRow(sel.active), c: sel.active.c } : null
        };
    }
    applyViewState(s: ShellViewStateInput | null | undefined): void {
        if (!s) return;
        if (s.zoom) this.grid.setZoom(s.zoom);
        if (s.active) this.grid.select(s.active.r, s.active.c);
        if (s.scroll) this.grid.setScroll(s.scroll);
    }
    _viewState(): void {
        if (!this.o.onViewState) return;
        clearTimeout(this._vt);
        this._vt = setTimeout(() => this.o.onViewState?.(this.getViewState()), 300);
    }
}

export { runTask, mergeSortGen, radixRows, Shell };
