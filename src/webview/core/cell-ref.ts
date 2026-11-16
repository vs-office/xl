// Spreadsheet cell addresses (A1 notation).
/** Zero-based cell coordinates. */
export interface CellRef {
    r: number;
    c: number;
}
/** Normalized (r1 <= r2, c1 <= c2) zero-based cell range. */
export interface RangeRef {
    r1: number;
    c1: number;
    r2: number;
    c2: number;
}

// ---------- addresses ----------
const MAXC = 16383,
    MAXR = 1048575;
const colCache: string[] = [];
export function colName(c: number): string {
    let s = colCache[c];
    if (s !== undefined) return s;
    if (!(c >= 0 && c <= MAXC) || c !== (c | 0)) return '';
    s = '';
    for (let n = c + 1; n > 0; n = ((n - 1) / 26) | 0) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
    return (colCache[c] = s);
}
export function colIndex(s: string) {
    const n = s.length;
    if (n < 1 || n > 3) return -1;
    let v = 0;
    for (let i = 0; i < n; i++) {
        let ch = s.charCodeAt(i);
        if (ch >= 97 && ch <= 122) ch -= 32;
        if (ch < 65 || ch > 90) return -1;
        v = v * 26 + (ch - 64);
    }
    return v - 1 <= MAXC ? v - 1 : -1;
}
export function cellRef(r: number, c: number) {
    return colName(c) + (r + 1);
}

// parses one endpoint into _pc/_pr (-1 when absent); returns false if invalid
let _pc = -1,
    _pr = -1;
function parsePart(s: string, a: number, b: number) {
    let i = a,
        ch: number;
    if (i < b && s.charCodeAt(i) === 36) i++;
    let c = 0,
        nl = 0;
    for (; i < b; i++) {
        ch = s.charCodeAt(i);
        if (ch >= 97 && ch <= 122) ch -= 32;
        if (ch < 65 || ch > 90) break;
        c = c * 26 + (ch - 64);
        if (++nl > 3) return false;
    }
    if (nl && c - 1 > MAXC) return false;
    if (i < b && s.charCodeAt(i) === 36) {
        if (!nl) return false;
        i++;
    }
    let r = 0,
        nd = 0;
    for (; i < b; i++) {
        ch = s.charCodeAt(i);
        if (ch < 48 || ch > 57) return false;
        r = r * 10 + (ch - 48);
        if (++nd > 7) return false;
    }
    if (nd && (r < 1 || r - 1 > MAXR)) return false;
    if (!nl && !nd) return false;
    _pc = nl ? c - 1 : -1;
    _pr = nd ? r - 1 : -1;
    return true;
}
export function parseRef(s: unknown): CellRef | null {
    if (typeof s !== 'string' || !parsePart(s, 0, s.length) || _pc < 0 || _pr < 0) return null;
    return { r: _pr, c: _pc };
}
export function parseRange(s: unknown): RangeRef | null {
    if (typeof s !== 'string') return null;
    const k = s.indexOf(':');
    let r1: number, c1: number, r2: number, c2: number;
    if (k < 0) {
        if (!parsePart(s, 0, s.length) || _pc < 0 || _pr < 0) return null;
        return { r1: _pr, c1: _pc, r2: _pr, c2: _pc };
    }
    if (!parsePart(s, 0, k)) return null;
    c1 = _pc;
    r1 = _pr;
    if (!parsePart(s, k + 1, s.length)) return null;
    c2 = _pc;
    r2 = _pr;
    if (c1 < 0 !== c2 < 0 || r1 < 0 !== r2 < 0 || (c1 < 0 && r1 < 0)) return null;
    if (r1 < 0) {
        r1 = 0;
        r2 = MAXR;
    }
    if (c1 < 0) {
        c1 = 0;
        c2 = MAXC;
    }
    return { r1: r1 < r2 ? r1 : r2, c1: c1 < c2 ? c1 : c2, r2: r1 < r2 ? r2 : r1, c2: c1 < c2 ? c2 : c1 };
}
