import { formatCell, exportText, cellRef, parseInput } from './format.ts';
import type { FormatKind, ParsedInput } from './format.ts';
import type { ArrayRange, Cell, CellStyle, Merge, Sheet, SheetStore, Table, Workbook } from './xlsx-reader.ts';

const KEY = 16384;
const CACHE_MAX = 60000;

/** A formula cell whose result is not known yet (what the editor produces for `=...`): numeric tag, null value. */
export interface FormulaNumberCell {
    t: 'n';
    v: number | null;
    s: number;
    f?: string;
}
/** A cell as seen through the provider: a reader cell or an edited formula cell. */
export type SlotCell = Cell | FormulaNumberCell;

/** Extra, optional fields of a cell passed to / returned from the editing entry points. */
interface EditExtras {
    /** style index; defaults to the existing / default style of the position */
    s?: number | undefined;
    f?: string | undefined;
    /** number format to apply (derives a style) */
    fmt?: string | undefined;
    /** set by editFromText when the cell was emptied */
    cleared?: boolean | undefined;
}
type EditOf<C extends SlotCell> = C extends SlotCell ? Omit<C, 'f' | 's'> & EditExtras : never;
/** A cell to be written into the overlay (see setCell / editFromText). */
export type EditCell = EditOf<SlotCell>;
/** Result of editFromText that refuses the input. */
export interface EditError {
    error: string;
}

/** What the grid shows for a cell. */
export interface CellView {
    text: string;
    kind: FormatKind;
    style: CellStyle | null;
    color: string | undefined;
    link: boolean;
}

/** Payload of a pending (non-clearing) edit, the writer's `{t, v, f?, fmt?}` cell contract. */
export interface PendingCell {
    t: SlotCell['t'];
    v: SlotCell['v'];
    f: string | undefined;
    fmt?: string;
}
/** One edit for the writer; `cell` is null when the cell is cleared. */
export interface PendingEdit {
    sheet: number;
    r: number;
    c: number;
    cell: PendingCell | null;
}

/** Measures the pixel width of a text drawn with a style. */
export type MeasureFn = (text: string, style: CellStyle | null) => number;
/** Returns an error message for an invalid formula (given with leading '='), else null/undefined. */
export type ValidateFormulaFn = (formula: string) => string | null | undefined;
export interface XlsxProviderOptions {
    measure?: MeasureFn | null | undefined;
    validateFormula?: ValidateFormulaFn | null | undefined;
}

/** Callback of forEachCell. */
export type CellVisitor = (cell: SlotCell, r: number, c: number) => void;

/** Derived (number-format override) styles of one workbook, shared by all providers of it. */
interface DerivedStyles {
    map: Map<string, number>;
    info: Map<number, { base: number; fmt: string }>;
}
const derivedByWorkbook = new WeakMap<Workbook, DerivedStyles>();

function isEditError(x: EditCell | EditError): x is EditError {
    return 'error' in x && x.error !== undefined;
}

// copy of `cell` with style `s` (and the formula, when present)
function restyle(cell: EditCell, s: number): SlotCell {
    let out: SlotCell;
    switch (cell.t) {
        case 'n':
            out = { t: 'n', v: cell.v, s };
            break;
        case 's':
            out = { t: 's', v: cell.v, s };
            break;
        case 'b':
            out = { t: 'b', v: cell.v, s };
            break;
        case 'e':
            out = { t: 'e', v: cell.v, s };
            break;
        case 'd':
            out = { t: 'd', v: cell.v, s };
            break;
        case 'z':
            out = { t: 'z', v: cell.v, s };
            break;
    }
    if (cell.f !== undefined) out.f = cell.f;
    return out;
}

class XlsxProvider {
    wb: Workbook;
    sheet: Sheet;
    merges: Merge[];
    frozenRows: number;
    frozenCols: number;
    showGridLines: boolean;
    date1904: boolean;
    overlay: Map<number, SlotCell | null>;
    colOverride: Map<number, number>;
    _cache: Map<number, CellView>;
    _extraRows: number;
    _extraCols: number;
    _measure: MeasureFn | null;
    // supplied by the page once the editing module (formula engine + writer) has been loaded
    _validateFormula: ValidateFormulaFn | null;
    _extDirty: boolean;
    _zh: boolean;
    _shownEdit: Set<number>;
    _st: SheetStore | null;
    _blank: Map<number, CellView | null>;

    constructor(workbook: Workbook, sheet: Sheet, opts?: XlsxProviderOptions) {
        this.wb = workbook;
        this.sheet = sheet;
        this.merges = sheet.merges;
        this.frozenRows = sheet.frozenRows;
        this.frozenCols = sheet.frozenCols;
        this.showGridLines = sheet.showGridLines;
        this.date1904 = workbook.date1904;
        this.overlay = new Map();
        this.colOverride = new Map();
        this._cache = new Map();
        this._extraRows = 0;
        this._extraCols = 0;
        this._measure = (opts && opts.measure) || null;
        this._validateFormula = (opts && opts.validateFormula) || null;
        this._extDirty = false;
        this._zh = !!sheet.zeroHeightRows;
        this._shownEdit = new Set();
        this._st = sheet.store || null;
        this._blank = new Map();
    }

    // default styles can appear while a sheet is still streaming in
    get _hasDef(): boolean {
        const s = this.sheet;
        return !!((s.rowStyles && s.rowStyles.size) || (s.colStyles && s.colStyles.length));
    }
    get partial(): boolean {
        return !!this.sheet.partial;
    }
    // while streaming, <dimension> gives the scroll extent early
    get rowCount(): number {
        if (this._extDirty) this._recalcExtent();
        const s = this.sheet;
        return Math.max(s.rowCount, this._extraRows, s.partial && s.dimension ? s.dimension.r2 + 1 : 0);
    }
    get colCount(): number {
        if (this._extDirty) this._recalcExtent();
        const s = this.sheet;
        return Math.max(s.colCount, this._extraCols, s.partial && s.dimension ? s.dimension.c2 + 1 : 0);
    }

    _recalcExtent(): void {
        let rr = 0,
            cc = 0;
        this.overlay.forEach((cell, key) => {
            if (!cell) return;
            const r = Math.floor(key / KEY),
                c = key - r * KEY;
            if (r + 1 > rr) rr = r + 1;
            if (c + 1 > cc) cc = c + 1;
        });
        this._extraRows = rr;
        this._extraCols = cc;
        this._extDirty = false;
    }

    // default style index of a cell that does not exist (row style wins over column style)
    _defS(r: number, c: number): number {
        if (!this._hasDef) return 0;
        const rs = this.sheet.rowStyles.get(r);
        if (rs !== undefined) return rs;
        const cs = this.sheet.colStyles;
        for (let i = 0; i < cs.length; i++) if (c >= cs[i].c1 && c <= cs[i].c2) return cs[i].s;
        return 0;
    }

    _blankView(r: number, c: number): CellView | undefined {
        const s = this._defS(r, c);
        if (!s) return undefined;
        let v = this._blank.get(s);
        if (v === undefined) {
            const st = this.wb.styles[s];
            v =
                st && (st.bg || st.bl || st.br || st.bt || st.bb)
                    ? { text: '', kind: 'text', style: st, color: undefined, link: false }
                    : null;
            this._blank.set(s, v);
        }
        return v === null ? undefined : v;
    }

    // style index = `s` with number format `fmt`; registers a derived Style in wb.styles (shared per workbook)
    _derive(s: number, fmt: string): number {
        const wb = this.wb;
        let d = derivedByWorkbook.get(wb);
        if (d === undefined) {
            d = { map: new Map(), info: new Map() };
            derivedByWorkbook.set(wb, d);
        }
        const inf = d.info.get(s),
            base = inf ? inf.base : s,
            key = base + '|' + fmt;
        let idx = d.map.get(key);
        if (idx === undefined) {
            idx = wb.styles.length;
            const proto: Partial<CellStyle> = wb.styles[base] || wb.styles[0] || {};
            // the `{}` fallback (workbook without any style) yields a partial style, as before
            wb.styles.push(Object.assign({}, proto, { fmt }) as CellStyle);
            d.map.set(key, idx);
            d.info.set(idx, { base, fmt });
        }
        return idx;
    }
    _fmtOf(s: number): string | undefined {
        const d = derivedByWorkbook.get(this.wb),
            inf = d && d.info.get(s);
        return inf ? inf.fmt : undefined;
    }

    cellAt(r: number, c: number): SlotCell | undefined {
        if (this.overlay.size !== 0) {
            const o = this.overlay.get(r * KEY + c);
            if (o !== undefined) return o === null ? undefined : o;
        }
        return this._orig(r, c);
    }

    getCell(r: number, c: number): CellView | undefined {
        const key = r * KEY + c;
        let v = this._cache.get(key);
        if (v !== undefined) return v;
        const cell = this.cellAt(r, c);
        if (cell === undefined) return this._hasDef ? this._blankView(r, c) : undefined;
        const style = this.wb.styles[cell.s] || null;
        if (cell.t === 'z' && !style && !cell.f) return undefined;
        const f = formatCell(cell, style, this.date1904);
        v = {
            text: f.text,
            kind: f.kind,
            style: style,
            color: f.color,
            link: this.sheet.hyperlinks.size > 0 && this.sheet.hyperlinks.has(r + ',' + c)
        };
        if (this._cache.size >= CACHE_MAX) this._cache.clear();
        this._cache.set(key, v);
        return v;
    }

    getRaw(r: number, c: number): string {
        const cell = this.cellAt(r, c);
        if (cell === undefined) return '';
        return exportText(cell, this.wb.styles[cell.s] || null, this.date1904);
    }

    colWidth(c: number): number {
        const o = this.colOverride.get(c);
        if (o !== undefined) return o;
        const w = this.sheet.colWidths.get(c);
        return w !== undefined ? w : this.sheet.defaultColWidth;
    }

    rowHeight(r: number): number {
        const h = this.sheet.rowHeights.get(r);
        return h !== undefined ? h : this.sheet.defaultRowHeight;
    }

    isColHidden(c: number): boolean {
        return this.sheet.hiddenCols.has(c);
    }
    isRowHidden(r: number): boolean {
        if (this.sheet.hiddenRows.has(r)) return true;
        const shown = this.sheet.shownRows;
        return this._zh && !(shown !== null && shown.has(r)) && !this._shownEdit.has(r);
    }
    setColWidth(c: number, w: number): void {
        this.colOverride.set(c, w);
    }
    setRowHeight(r: number, h: number): void {
        this.sheet.rowHeights.set(r, h);
    }

    _same(o: SlotCell | undefined, cell: SlotCell): boolean {
        return o !== undefined && o.t === cell.t && o.v === cell.v && (o.f || '') === (cell.f || '') && o.s === cell.s;
    }

    // returns the previous overlay slot so callers can undo; an {error} result from editFromText is refused
    setCell(r: number, c: number, cell: EditCell | EditError | null | undefined): SlotCell | null | undefined {
        const key = r * KEY + c;
        const prev = this.overlay.has(key) ? this.overlay.get(key) : undefined;
        if (cell && isEditError(cell)) return prev;
        const orig = this._orig(r, c);
        if (cell === undefined || (cell === orig && orig !== undefined)) this.overlay.delete(key);
        else if (cell === null) {
            if (!orig || (orig.t === 'z' && !orig.f)) this.overlay.delete(key);
            else this.overlay.set(key, orig.s ? { t: 'z', v: null, s: orig.s } : null); // keep the style visible
        } else {
            const base = orig || prev;
            let st = cell.s !== undefined ? cell.s : base ? base.s : this._defS(r, c);
            if (cell.fmt) st = this._derive(st, cell.fmt);
            const withStyle = restyle(cell, st);
            if (orig !== undefined && this._same(orig, withStyle)) this.overlay.delete(key);
            else this.overlay.set(key, withStyle);
        }
        this._cache.delete(key);
        if (cell) {
            if (r + 1 > this._extraRows) this._extraRows = r + 1;
            if (c + 1 > this._extraCols) this._extraCols = c + 1;
            if (this._zh) this._shownEdit.add(r);
        }
        return prev;
    }

    restoreSlot(r: number, c: number, slot: SlotCell | null | undefined): void {
        const key = r * KEY + c;
        if (slot === undefined) this.overlay.delete(key);
        else this.overlay.set(key, slot);
        this._cache.delete(key);
        this._extDirty = true;
    }

    _orig(r: number, c: number): Cell | undefined {
        const st = this._st;
        if (st !== null) {
            const i = st.index(r, c);
            return i < 0 ? undefined : st.cellAt(i, r, c);
        }
        const row = this.sheet.rows[r];
        return row === undefined ? undefined : row[c];
    }

    _tableAt(r: number, c: number): Table | null {
        const ts = this.sheet.tables;
        if (ts) for (const t of ts) if (t.header && r === t.r1 && c >= t.c1 && c <= t.c2) return t;
        return null;
    }

    // array (CSE) formula ranges of this sheet: [{r1,c1,r2,c2,mr,mc}] (mr,mc = master cell); cheap, no cell scan
    arrayRanges(): ArrayRange[] {
        const a = this.sheet.arrays;
        return a ? a.map(x => ({ r1: x.r1, c1: x.c1, r2: x.r2, c2: x.c2, mr: x.mr, mc: x.mc })) : [];
    }

    // the array formula range covering (r,c) or null
    arrayAt(r: number, c: number): ArrayRange | null {
        const a = this.sheet.arrays;
        if (a)
            for (const x of a)
                if (r >= x.r1 && r <= x.r2 && c >= x.c1 && c <= x.c2)
                    return { r1: x.r1, c1: x.c1, r2: x.r2, c2: x.c2, mr: x.mr, mc: x.mc };
        return null;
    }

    // typed text -> edit cell {t,v,f?,fmt?}; keeps existing style. Returns {error} when the input cannot be saved.
    editFromText(
        r: number,
        c: number,
        text: string | number | boolean | null | undefined
    ): EditCell | EditError | undefined {
        const prev = this.cellAt(r, c);
        if (typeof text === 'string' && text.length > 32767)
            return { error: 'Text is longer than the 32767 character limit of a cell.' };
        const arrs = this.sheet.arrays;
        if (arrs && arrs.length) {
            for (const a of arrs) {
                if (
                    r < a.r1 ||
                    r > a.r2 ||
                    c < a.c1 ||
                    c > a.c2 ||
                    (r === a.mr && c === a.mc) ||
                    this.overlay.has(a.mr * KEY + a.mc)
                )
                    continue;
                return {
                    error:
                        'Cannot change part of an array formula (' +
                        cellRef(r, c) +
                        ' is inside ' +
                        cellRef(a.r1, a.c1) +
                        ':' +
                        cellRef(a.r2, a.c2) +
                        '). Edit ' +
                        cellRef(a.mr, a.mc) +
                        ' or clear the whole array.'
                };
            }
        }
        const style = prev ? this.wb.styles[prev.s] || null : null;
        const parsed = parseInput(text, prev, style, this.date1904);
        const tbl = this._tableAt(r, c);
        if (tbl) {
            if (parsed === null) return { error: 'A table header cell cannot be empty.' };
            if ((parsed.t === 'n' && parsed.f !== undefined) || parsed.t === 'e')
                return { error: 'A table header must be plain text, not a formula or error.' };
            const v = parsed.t === 's' ? parsed.v : String(text).trim();
            const lv = v.toLowerCase();
            for (let k = tbl.c1; k <= tbl.c2; k++) {
                if (k === c) continue;
                const o = this.cellAt(r, k);
                if (o && o.v !== null && o.v !== undefined && String(o.v).toLowerCase() === lv)
                    return { error: 'Table column names must be unique ("' + v + '" is already used).' };
            }
            return { t: 's', v, f: undefined, s: prev ? prev.s : this._defS(r, c) };
        }
        if (parsed === null)
            return prev === undefined || (prev.t === 'z' && !prev.f)
                ? undefined
                : { t: 'z', v: null, s: prev.s, cleared: true };
        if (parsed.t === 'n' && parsed.f !== undefined && this._validateFormula) {
            const err = this._validateFormula('=' + parsed.f);
            if (err) return { error: err };
        }
        const st = prev ? prev.s : this._defS(r, c);
        const out = this._fromParsed(parsed, st);
        if (parsed.t === 'n' && parsed.fmt) {
            out.fmt = parsed.fmt;
            out.s = this._derive(st, parsed.fmt);
        }
        return out;
    }

    _fromParsed(parsed: ParsedInput, s: number): EditCell {
        switch (parsed.t) {
            case 'n':
                return { t: 'n', v: parsed.v, f: parsed.f, s };
            case 's':
                return { t: 's', v: parsed.v, f: undefined, s };
            case 'b':
                return { t: 'b', v: parsed.v, f: undefined, s };
            case 'e':
                return { t: 'e', v: parsed.v, f: undefined, s };
        }
    }

    // writer edit cell contract: {t, v, f?, fmt?}; cleared cells are null
    pendingEdits(sheetIndex: number): PendingEdit[] {
        const out: PendingEdit[] = [];
        this.overlay.forEach((cell, key) => {
            const r = Math.floor(key / KEY),
                c = key - r * KEY;
            const orig = this._orig(r, c);
            if (cell === null || (cell.t === 'z' && !cell.f)) {
                if (orig === undefined || (orig.t === 'z' && !orig.f)) return;
                out.push({ sheet: sheetIndex, r: r, c: c, cell: null });
                return;
            }
            const fmt = this._fmtOf(cell.s);
            if (fmt === undefined && orig !== undefined && this._same(orig, cell)) return;
            const e: PendingCell = { t: cell.t, v: cell.v, f: cell.f };
            if (fmt !== undefined) e.fmt = fmt;
            out.push({ sheet: sheetIndex, r: r, c: c, cell: e });
        });
        return out;
    }

    get dirty(): boolean {
        return this.overlay.size > 0;
    }

    forEachCell(r1: number, c1: number, r2: number, c2: number, fn: CellVisitor): void {
        const st = this._st,
            ov = this.overlay;
        if (st !== null) {
            if (ov.size === 0) {
                st.each(r1, c1, r2, c2, fn);
                return;
            }
            st.each(r1, c1, r2, c2, (cell, r, c) => {
                const o = ov.get(r * KEY + c);
                if (o === undefined) fn(cell, r, c);
                else if (o) fn(o, r, c);
            });
            ov.forEach((cell, key) => {
                const r = Math.floor(key / KEY),
                    c = key - r * KEY;
                if (cell && r >= r1 && r <= r2 && c >= c1 && c <= c2 && st.index(r, c) < 0) fn(cell, r, c);
            });
            return;
        }
        const rows = this.sheet.rows;
        const lastR = Math.min(r2, rows.length - 1);
        for (let r = r1; r <= lastR; r++) {
            const row = rows[r];
            if (row === undefined) continue;
            const lastC = Math.min(c2, row.length - 1);
            for (let c = c1; c <= lastC; c++) {
                const cell = ov.size ? this.cellAt(r, c) : row[c];
                if (cell !== undefined) fn(cell, r, c);
            }
        }
        if (ov.size) {
            ov.forEach((cell, key) => {
                const r = Math.floor(key / KEY),
                    c = key - r * KEY;
                if (
                    r >= r1 &&
                    r <= r2 &&
                    c >= c1 &&
                    c <= c2 &&
                    (r >= rows.length || rows[r] === undefined || rows[r][c] === undefined) &&
                    cell
                )
                    fn(cell, r, c);
            });
        }
    }

    autosizeCol(c: number, measure?: MeasureFn | null): number {
        const m = measure || this._measure;
        if (!m) return this.colWidth(c);
        let max = 0,
            n = 0;
        const st = this._st,
            ov = this.overlay.size > 0;
        const end = st !== null ? Math.max(this.sheet.rowCount, this._extraRows) : this.sheet.rows.length;
        for (let r = 0; r < end && n < 3000; r++) {
            if (st !== null && !ov && st.index(r, c) < 0 && !this._hasDef) continue;
            const cv = this.getCell(r, c);
            if (cv === undefined || cv.text === '') continue;
            n++;
            const w = m(cv.text, cv.style);
            if (w > max) max = w;
        }
        const px = Math.max(48, Math.min(600, Math.ceil(max) + 16));
        this.colOverride.set(c, px);
        return px;
    }
}

export { XlsxProvider };
