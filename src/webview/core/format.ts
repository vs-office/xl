import SSF from '../vendor/ssf.js';
import { colName, colIndex, cellRef, parseRef, parseRange } from './cell-ref.ts';
import { generalNum } from './general.ts';
export type { CellRef, RangeRef } from './cell-ref.ts';

// ---------- public types ----------
/** The fields of a grid cell this module reads. */
export interface FormatCell {
    t: string;
    v?: unknown;
    f?: string | null | undefined;
}
/** The fields of a cell style this module reads. */
export interface FormatStyle {
    fmt?: string | undefined;
}
export type FormatKind = 'text' | 'number' | 'date' | 'bool' | 'error';
export interface FormatResult {
    text: string;
    color: string | undefined;
    kind: FormatKind;
}
export interface ParsedString {
    t: 's';
    v: string;
}
export interface ParsedError {
    t: 'e';
    v: string;
}
export interface ParsedBool {
    t: 'b';
    v: boolean;
}
/** Numeric input; `f` is set (with v null) for formulas, `fmt` is an inferred number format. */
export interface ParsedNumber {
    t: 'n';
    v: number | null;
    f?: string;
    fmt?: string;
}
export type ParsedInput = ParsedString | ParsedError | ParsedBool | ParsedNumber;
/** Running numeric statistics over a selection. */
export interface Stats {
    count: number;
    sum: number;
    min: number;
    max: number;
    add(n: number): void;
    avg(): number;
}

// String(v) for cell values of unknown shape; null/undefined become ''.
function textOf(v: unknown): string {
    if (v === null || v === undefined) return '';
    if (typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean' || typeof v === 'bigint')
        return String(v);
    if (v instanceof Date) return String(v);
    return Object.prototype.toString.call(v);
}

// ---------- dates ----------
const MS_DAY = 86400000;
const UTC_1899_1230 = Date.UTC(1899, 11, 30),
    UTC_1904 = Date.UTC(1904, 0, 1);
const MAX_SERIAL_1900 = 2958465,
    MAX_SERIAL_1904 = 2957003;

function serialToDate(n: number, date1904?: boolean): Date {
    if (date1904) return new Date(UTC_1904 + Math.round(n * MS_DAY));
    // serials below 61 are shifted by the fictitious 1900-02-29 (serial 60 maps to Mar 1)
    const days = n < 61 ? n + 1 : n;
    return new Date(UTC_1899_1230 + Math.round(days * MS_DAY));
}
function dateToSerial(d: Date | number, date1904?: boolean): number {
    const t = d instanceof Date ? d.getTime() : +d;
    if (date1904) return (t - UTC_1904) / MS_DAY;
    const days = (t - UTC_1899_1230) / MS_DAY;
    return days < 61 ? days - 1 : days;
}
function isLeap(y: number) {
    return (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;
}
const DIM = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
// serial for a calendar date, or -1 if invalid (1900-02-29 accepted like Excel)
function ymdToSerial(y: number, m: number, d: number, date1904?: boolean): number {
    if (m < 1 || m > 12 || d < 1) return -1;
    if (!(d <= DIM[m - 1] || (m === 2 && d === 29 && (isLeap(y) || y === 1900)))) return -1;
    if (y === 1900 && m === 2 && d === 29) return date1904 ? -1 : 60;
    const days = Math.round((Date.UTC(y, m - 1, d) - (date1904 ? UTC_1904 : UTC_1899_1230)) / MS_DAY);
    if (y < 100) return -1;
    const s = date1904 ? days : days < 61 ? days - 1 : days;
    return s < (date1904 ? 0 : 1) || s > (date1904 ? MAX_SERIAL_1904 : MAX_SERIAL_1900) ? -1 : s;
}

const MONTHS = [
    'January',
    'February',
    'March',
    'April',
    'May',
    'June',
    'July',
    'August',
    'September',
    'October',
    'November',
    'December'
];
const MON3 = MONTHS.map(m => m.slice(0, 3));
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
const DAY3 = DAYS.map(d => d.slice(0, 3));
const P2: string[] = [];
for (let i = 0; i < 100; i++) P2.push(i < 10 ? '0' + i : '' + i);

// ---------- fixed-decimal fast path ----------
const POW10: number[] = [];
for (let i = 0; i <= 20; i++) POW10.push(Math.pow(10, i));
function exactRound(a: number, shift: number) {
    const es = a.toExponential(14),
        k = es.indexOf('e');
    return Math.round(Number(es.slice(0, k) + 'e' + (+es.slice(k + 1) + shift)));
}
function group(s: string) {
    const n = s.length;
    if (n <= 3) return s;
    let out = s.slice(0, n % 3 || 3);
    for (let i = n % 3 || 3; i < n; i += 3) out += ',' + s.slice(i, i + 3);
    return out;
}
const ZEROS = '0000000000000000000000';
function makeFixed(d: number, comma: boolean, pct: boolean): (v: number) => string | null {
    const sh = d + (pct ? 2 : 0),
        sc = POW10[sh];
    return function (v: number) {
        const a = v < 0 ? -v : v,
            x = a * sc;
        if (!(x < 1e15)) {
            if (!(x < Infinity)) return null;
            return fixedStr(bigDigits(x), d, comma, pct, v < 0);
        }
        let r = Math.round(x);
        const df = x - Math.floor(x) - 0.5,
            tol = x * 1e-15 + 1e-12;
        if (df < tol && df > -tol) r = exactRound(a, sh);
        return fixedStr('' + r, d, comma, pct, v < 0);
    };
}
// integer digits of a large double with 15 significant digits (like Excel), then zeros
function bigDigits(x: number) {
    const es = x.toExponential(14),
        k = es.indexOf('e');
    return es.charAt(0) + es.slice(2, k) + '0'.repeat(+es.slice(k + 1) - 14);
}
function fixedStr(s: string, d: number, comma: boolean, pct: boolean, neg: boolean) {
    if (d > 0) {
        if (s.length <= d) s = ZEROS.slice(0, d - s.length + 1) + s;
        const ip = s.slice(0, s.length - d);
        s = (comma ? group(ip) : ip) + '.' + s.slice(s.length - d);
    } else if (comma) s = group(s);
    if (pct) s += '%';
    return neg ? '-' + s : s;
}

// ---------- date compile ----------
const OP_LIT = 0,
    OP_Y2 = 1,
    OP_Y4 = 2,
    OP_MO1 = 3,
    OP_MO2 = 4,
    OP_MO3 = 5,
    OP_MO4 = 6,
    OP_MO5 = 7,
    OP_MI1 = 8,
    OP_MI2 = 9,
    OP_D1 = 10,
    OP_D2 = 11,
    OP_D3 = 12,
    OP_D4 = 13,
    OP_H1 = 14,
    OP_H2 = 15,
    OP_S1 = 16,
    OP_S2 = 17,
    OP_AMPM = 18,
    OP_AP = 19;
/** Compiled date/time renderer; null when the code needs the generic SSF engine. */
type DateFormatter = (v: number, d1904?: boolean) => string;
function compileDate(code: string): DateFormatter | null {
    const n = code.length;
    const ops: number[] = [],
        args: string[] = [];
    let i = 0,
        hasTime = false,
        ampm = false;
    const push = (o: number, a = ''): void => {
        ops.push(o);
        args.push(a);
    };
    while (i < n) {
        const ch = code.charCodeAt(i),
            lc = ch | 32;
        if (ch === 34) {
            const j = code.indexOf('"', i + 1);
            if (j < 0) return null;
            push(OP_LIT, code.slice(i + 1, j));
            i = j + 1;
        } else if (ch === 92) {
            if (i + 1 >= n) return null;
            push(OP_LIT, code.charAt(i + 1));
            i += 2;
        } else if (ch === 95) {
            push(OP_LIT, ' ');
            i += 2;
        } else if (lc === 121 || lc === 109 || lc === 100 || lc === 104 || lc === 115) {
            let j = i + 1;
            while (j < n && (code.charCodeAt(j) | 32) === lc) j++;
            const c = j - i;
            i = j;
            if (lc === 121) {
                if (c > 4) return null;
                push(c <= 2 ? OP_Y2 : OP_Y4);
            } else if (lc === 109) {
                if (c > 5) return null;
                push(-c);
            } else if (lc === 100) {
                if (c > 4) return null;
                push(OP_D1 + c - 1);
            } else if (lc === 104) {
                if (c > 2) return null;
                hasTime = true;
                push(OP_H1 + c - 1);
            } else {
                if (c > 2) return null;
                hasTime = true;
                push(OP_S1 + c - 1);
            }
        } else if (ch === 65 || ch === 97) {
            if (code.startsWith('AM/PM', i)) {
                push(OP_AMPM);
                ampm = true;
                hasTime = true;
                i += 5;
            } else if (code.startsWith('A/P', i)) {
                push(OP_AP);
                ampm = true;
                hasTime = true;
                i += 3;
            } else return null;
        } else if (ch === 46) {
            const nx = code.charCodeAt(i + 1);
            if (nx >= 48 && nx <= 57) return null;
            push(OP_LIT, '.');
            i++;
        } else if (
            ch === 32 ||
            ch === 45 ||
            ch === 47 ||
            ch === 58 ||
            ch === 44 ||
            ch === 40 ||
            ch === 41 ||
            ch === 36 ||
            ch === 43 ||
            ch === 33 ||
            ch === 38 ||
            ch === 126 ||
            ch === 123 ||
            ch === 125 ||
            ch === 60 ||
            ch === 61 ||
            ch === 62 ||
            ch === 94 ||
            ch === 39
        ) {
            push(OP_LIT, code.charAt(i));
            i++;
        } else return null;
    }
    // resolve m: minutes if previous significant token is h or next is s
    for (let k = 0; k < ops.length; k++) {
        if (ops[k] >= 0) continue;
        const c = -ops[k];
        let isMin = false;
        if (c <= 2) {
            for (let p = k - 1; p >= 0; p--) {
                const o = ops[p];
                if (o === OP_LIT) continue;
                isMin = o === OP_H1 || o === OP_H2;
                break;
            }
            if (!isMin)
                for (let p = k + 1; p < ops.length; p++) {
                    const o = ops[p];
                    if (o === OP_LIT) continue;
                    isMin = o === OP_S1 || o === OP_S2;
                    break;
                }
        }
        ops[k] = isMin ? OP_MI1 + c - 1 : OP_MO1 + c - 1;
    }
    for (let k = 0; k < ops.length; k++) if (ops[k] === OP_LIT && args[k] === undefined) args[k] = '';
    const N = ops.length;
    return function (v: number, d1904?: boolean): string {
        if (!(v >= 0)) return '########';
        const maxS = d1904 ? MAX_SERIAL_1904 : MAX_SERIAL_1900;
        let days = Math.floor(v),
            secs = 0;
        if (hasTime) {
            secs = Math.round((v - days) * 86400);
            if (secs >= 86400) {
                secs -= 86400;
                days++;
            }
        }
        if (days > maxS) return '########';
        let y: number, m: number, d: number, dow: number;
        if (d1904) {
            dow = (days + 5) % 7;
        } else {
            dow = (((days - 1) % 7) + 7) % 7;
        }
        if (!d1904 && days === 0) {
            y = 1900;
            m = 1;
            d = 0;
        } else if (!d1904 && days === 60) {
            y = 1900;
            m = 2;
            d = 29;
        } else {
            const u = d1904 ? days - 24107 : days < 60 ? days - 25568 : days - 25569;
            const z = u + 719468,
                era = Math.floor(z / 146097),
                doe = z - era * 146097;
            const yoe = Math.floor(
                (doe - Math.floor(doe / 1460) + Math.floor(doe / 36524) - Math.floor(doe / 146096)) / 365
            );
            const doy = doe - (365 * yoe + Math.floor(yoe / 4) - Math.floor(yoe / 100));
            const mp = Math.floor((5 * doy + 2) / 153);
            d = doy - Math.floor((153 * mp + 2) / 5) + 1;
            m = mp < 10 ? mp + 3 : mp - 9;
            y = yoe + era * 400 + (m <= 2 ? 1 : 0);
        }
        const H = (secs / 3600) | 0,
            Mi = ((secs / 60) | 0) % 60,
            S = secs % 60;
        let out = '';
        for (let k = 0; k < N; k++) {
            switch (ops[k]) {
                case OP_LIT:
                    out += args[k];
                    break;
                case OP_Y2:
                    out += P2[y % 100];
                    break;
                case OP_Y4:
                    out += y;
                    break;
                case OP_MO1:
                    out += m;
                    break;
                case OP_MO2:
                    out += P2[m];
                    break;
                case OP_MO3:
                    out += MON3[m - 1] || '';
                    break;
                case OP_MO4:
                    out += MONTHS[m - 1] || '';
                    break;
                case OP_MO5:
                    out += (MONTHS[m - 1] || '').charAt(0);
                    break;
                case OP_MI1:
                    out += Mi;
                    break;
                case OP_MI2:
                    out += P2[Mi];
                    break;
                case OP_D1:
                    out += d;
                    break;
                case OP_D2:
                    out += P2[d];
                    break;
                case OP_D3:
                    out += DAY3[dow];
                    break;
                case OP_D4:
                    out += DAYS[dow];
                    break;
                case OP_H1:
                    out += ampm ? H % 12 || 12 : H;
                    break;
                case OP_H2:
                    out += P2[ampm ? H % 12 || 12 : H];
                    break;
                case OP_S1:
                    out += S;
                    break;
                case OP_S2:
                    out += P2[S];
                    break;
                case OP_AMPM:
                    out += H < 12 ? 'AM' : 'PM';
                    break;
                case OP_AP:
                    out += H < 12 ? 'A' : 'P';
                    break;
            }
        }
        return out;
    };
}

// ---------- format preparation ----------
const SSF_OPTS = { date1904: false },
    SSF_OPTS_1904 = { date1904: true };
const NAMED: Record<string, string> = {
    black: '#000000',
    blue: '#0000ff',
    cyan: '#00ffff',
    green: '#00ff00',
    magenta: '#ff00ff',
    red: '#ff0000',
    white: '#ffffff',
    yellow: '#ffff00'
};
const PALETTE = [
    '#000000',
    '#ffffff',
    '#ff0000',
    '#00ff00',
    '#0000ff',
    '#ffff00',
    '#ff00ff',
    '#00ffff',
    '#800000',
    '#008000',
    '#000080',
    '#808000',
    '#800080',
    '#008080',
    '#c0c0c0',
    '#808080',
    '#9999ff',
    '#993366',
    '#ffffcc',
    '#ccffff',
    '#660066',
    '#ff8080',
    '#0066cc',
    '#ccccff',
    '#000080',
    '#ff00ff',
    '#ffff00',
    '#00ffff',
    '#800080',
    '#800000',
    '#008080',
    '#0000ff',
    '#00ccff',
    '#ccffff',
    '#ccffcc',
    '#ffff99',
    '#99ccff',
    '#ff99cc',
    '#cc99ff',
    '#ffcc99',
    '#3366ff',
    '#33cccc',
    '#99cc00',
    '#ffcc00',
    '#ff9900',
    '#ff6600',
    '#666699',
    '#969696',
    '#003366',
    '#339966',
    '#003300',
    '#333300',
    '#993300',
    '#993366',
    '#333399',
    '#333333'
];
const COLOR_RE = /\[(black|blue|cyan|green|magenta|red|white|yellow|color\s*(\d+))\]/i;
const COND_ANY = /\[[=<>]/;
const COND_RE = /\[(=|>[=]?|<[>=]?)(-?\d+(?:\.\d*)?)\]/;

// Normalizes a format code for SSF: drops [$-LCID] tokens (SSF turns [$-409] into "$"), removes spaces before % (re-added to the output, SSF rejects them), quotes bare non-ASCII chars, uppercases e+/e-.
const NORM_NEEDED = /\[\$-|[^\0-\x7f]|[ \\]+%|[eE][+-]|\\ /;
function pctPad(code: string) {
    const m = /((?:\\? )+)%/.exec(code.replace(/"[^"]*"/g, ''));
    return m ? m[1].replace(/\\/g, '') : '';
}
function normCode(code: string) {
    if (!NORM_NEEDED.test(code)) return code;
    let out = '',
        i = 0;
    const n = code.length;
    while (i < n) {
        const ch = code.charCodeAt(i);
        if (ch === 34) {
            const j = code.indexOf('"', i + 1);
            const e = j < 0 ? n : j + 1;
            out += code.slice(i, e);
            i = e;
        } else if (ch === 92) {
            out += code.charCodeAt(i + 1) === 32 && code.charCodeAt(i + 2) === 37 ? '' : code.slice(i, i + 2);
            i += 2;
        } else if (ch === 95) {
            out += code.slice(i, i + 2);
            i += 2;
        } else if (ch === 91) {
            const j = code.indexOf(']', i);
            const e = j < 0 ? n : j + 1,
                tok = code.slice(i, e);
            if (!/^\[\$-[^\]]*\]$/.test(tok)) out += tok;
            i = e;
        } else if (ch > 127) {
            out += '"' + code.charAt(i) + '"';
            i++;
        } else if (ch === 32) {
            let j = i;
            while (j < n && code.charCodeAt(j) === 32) j++;
            if (code.charCodeAt(j) !== 37) out += code.slice(i, j);
            i = j;
        } else if ((ch === 101 || ch === 69) && (code.charCodeAt(i + 1) === 43 || code.charCodeAt(i + 1) === 45)) {
            out += 'E';
            i++;
        } else {
            out += code.charAt(i);
            i++;
        }
    }
    return out;
}

function sectionColor(sec: string) {
    const m = COLOR_RE.exec(sec);
    if (!m) return undefined;
    if (m[2] !== undefined) {
        const i = +m[2];
        return i >= 1 && i <= 56 ? PALETTE[i - 1] : undefined;
    }
    return NAMED[m[1].toLowerCase()];
}
interface Cond {
    op: string;
    t: number;
}
function condOf(sec: string): Cond | null {
    const m = COND_RE.exec(sec);
    return m ? { op: m[1], t: parseFloat(m[2]) } : null;
}
function condHit(v: number, c: Cond | null): boolean {
    if (!c) return false;
    switch (c.op) {
        case '=':
            // loose on purpose: criteria may be text while the value is a number
            return v === c.t;
        case '>':
            return v > c.t;
        case '<':
            return v < c.t;
        case '<>':
            return v !== c.t;
        case '>=':
            return v >= c.t;
        case '<=':
            return v <= c.t;
    }
    return false;
}

const FIXED_RE = /^(#,##0|0)(?:\.(0{1,10}))?(%)?$/;
/** Parsed/compiled form of one number-format code, cached per code. */
interface PreparedFormat {
    code: string;
    isDate: boolean;
    general: boolean;
    fast: ((v: number, d1904?: boolean) => string | null) | null;
    ssfCode: string;
    colors: (string | undefined)[] | null;
    sec4: string[] | null;
    c1: Cond | null;
    c2: Cond | null;
    cond: boolean;
    textColor: string | undefined;
    hasAt: boolean;
    textOnly: boolean;
    pad: string;
    lowE: boolean;
    cache: Map<number | string, string>;
    cache1904: Map<number | string, string>;
    datePart: boolean;
    timePart: boolean;
    pct: boolean;
    genLit: boolean;
    broken: boolean;
}
const fmtCache = new Map<string, PreparedFormat>();
let lastFmt = 'General',
    lastPrep: PreparedFormat | null = null;
const CACHE_MAX = 4096;

function prepare(code: string): PreparedFormat {
    const cached = fmtCache.get(code);
    if (cached) return cached;
    const p: PreparedFormat = {
        code,
        isDate: false,
        general: code === 'General',
        fast: null,
        ssfCode: normCode(code),
        colors: null,
        sec4: null,
        c1: null,
        c2: null,
        cond: false,
        textColor: undefined,
        hasAt: false,
        textOnly: code === '@',
        pad: '',
        lowE: false,
        cache: new Map(),
        cache1904: new Map(),
        datePart: false,
        timePart: false,
        pct: false,
        genLit: false,
        broken: false
    };
    if (!p.general && !p.textOnly) {
        try {
            const sc = p.ssfCode;
            if (sc !== code) {
                p.pad = pctPad(code);
                p.lowE = /e[+-]/.test(code.replace(/"[^"]*"|\[[^\]]*\]/g, ''));
            }
            const secs = SSF._split(sc);
            if (secs.length > 4) throw new Error('sections');
            const m = FIXED_RE.exec(sc);
            if (m) {
                const f0 = makeFixed(m[2] ? m[2].length : 0, m[1] === '#,##0', !!m[3]);
                p.fast = f0;
                if (p.pad) {
                    const pad = p.pad;
                    p.fast = v => {
                        const r = f0(v);
                        return r === null ? null : r.replace('%', pad + '%');
                    };
                }
                p.pct = !!m[3];
            } else if (SSF.is_date(sc) && secs.length === 1) {
                p.isDate = true;
                p.fast = compileDate(sc);
            } else if (SSF.is_date(sc)) p.isDate = true;
            else if (/general/i.test(sc) && secs.length === 1 && !/^general$/i.test(sc)) p.genLit = true;
            else if (sc.indexOf('%') >= 0) p.pct = /^[^"]*%/.test(secs[0]);
            if (p.isDate) {
                const s = sc.replace(/\[(h+|m+|s+)\]/gi, 'h').replace(/"[^"]*"|\\.|_.|\[[^\]]*\]/g, '');
                p.timePart = /[hs]/i.test(s) || /AM\/PM|A\/P/i.test(s);
                p.datePart = /[yd]/i.test(s) || (/m/i.test(s) && !/[hs]/i.test(s));
            }
            if (/^general$/i.test(sc)) p.general = true;
            // sections for color/condition selection (mirrors SSF choose_fmt)
            if (sc.indexOf('[') >= 0 || secs.length > 1) {
                const last = secs[secs.length - 1],
                    lat = last.indexOf('@');
                p.hasAt = secs.some(s => s.indexOf('@') >= 0);
                let f: string[];
                switch (secs.length) {
                    case 1:
                        f = lat > -1 ? ['General', 'General', 'General', secs[0]] : [secs[0], secs[0], secs[0], '@'];
                        break;
                    case 2:
                        f = lat > -1 ? [secs[0], secs[0], secs[0], secs[1]] : [secs[0], secs[1], secs[0], '@'];
                        break;
                    case 3:
                        f = lat > -1 ? [secs[0], secs[1], secs[0], secs[2]] : [secs[0], secs[1], secs[2], '@'];
                        break;
                    default:
                        f = secs.slice();
                }
                p.sec4 = f;
                const cols = f.map(sectionColor);
                if (cols.some(c => c !== undefined)) p.colors = cols;
                if (sc.indexOf('[') >= 0 && (COND_ANY.test(f[0]) || COND_ANY.test(f[1]))) {
                    p.cond = true;
                    p.c1 = condOf(f[0]);
                    p.c2 = condOf(f[1]);
                }
                if (secs.length === 4 || lat > -1) p.textColor = sectionColor(last);
            } else p.hasAt = sc.indexOf('@') >= 0;
        } catch {
            p.broken = true;
            p.general = true;
            p.fast = null;
            p.isDate = false;
        }
    }
    if (fmtCache.size > 2000) fmtCache.clear();
    fmtCache.set(code, p);
    return p;
}
function getPrep(fmt: string): PreparedFormat {
    if (fmt === lastFmt && lastPrep !== null) return lastPrep;
    lastPrep = prepare(fmt);
    lastFmt = fmt;
    return lastPrep;
}
const GENERAL_PREP = prepare('General');

function pickIndex(p: PreparedFormat, v: number): number {
    if (p.cond) {
        if (condHit(v, p.c1)) return 0;
        if (condHit(v, p.c2)) return 1;
        return p.c1 && p.c2 ? 2 : 1;
    }
    return v > 0 ? 0 : v < 0 ? 1 : 2;
}

function ssfRun(p: PreparedFormat, v: number | string, d1904: boolean | undefined, code?: string): string | null {
    try {
        return SSF.format(code || p.ssfCode, v, d1904 ? SSF_OPTS_1904 : SSF_OPTS);
    } catch {
        return null;
    }
}
function slowNumber(p: PreparedFormat, v: number, d1904?: boolean): string {
    const cache = d1904 ? p.cache1904 : p.cache;
    let s: string | null | undefined = cache.get(v);
    if (s !== undefined) return s;
    if (p.genLit) {
        const sc = p.ssfCode;
        const out = ssfRun(p, Math.abs(v), d1904, sc.replace(/general/i, '"\u0001"'));
        s = out === null ? generalNum(v) : out.replace('\u0001', generalNum(v));
    } else {
        s = ssfRun(p, v, d1904);
        if (s !== null) {
            if (p.pad) s = s.replace('%', p.pad + '%');
            if (p.lowE) s = s.replace('E', 'e');
        }
        if (s === null) s = generalNum(v);
        else if (p.isDate && s === '') s = '########';
    }
    if (cache.size >= CACHE_MAX) cache.clear();
    cache.set(v, s);
    return s;
}
function numberText(p: PreparedFormat, v: number, d1904?: boolean): string {
    if (p.general) return generalNum(v);
    if (p.fast !== null) {
        const s = p.fast(v, d1904);
        if (s !== null) return s;
    }
    return slowNumber(p, v, d1904);
}

const MAXDATE_STR = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2})(?:\.\d+)?)?)?Z?$/;
function isoToSerial(s: string, d1904: boolean): number {
    const m = MAXDATE_STR.exec(s);
    if (!m) return NaN;
    const day = ymdToSerial(+m[1], +m[2], +m[3], d1904);
    if (day < 0) return NaN;
    return m[4] === undefined ? day : day + (+m[4] * 3600 + +m[5] * 60 + (m[6] ? +m[6] : 0)) / 86400;
}

function formatCell(cell: FormatCell | null | undefined, style?: FormatStyle | null, date1904?: boolean): FormatResult {
    if (cell === undefined || cell === null) return { text: '', color: undefined, kind: 'text' };
    const fmt = style ? style.fmt || 'General' : 'General';
    switch (cell.t) {
        case 'n': {
            const v = cell.v;
            if (typeof v !== 'number') {
                if (v === null || v === undefined) return { text: '', color: undefined, kind: 'number' };
                return { text: textOf(v), color: undefined, kind: 'text' };
            }
            if (v !== v || v === Infinity || v === -Infinity) return { text: '#NUM!', color: undefined, kind: 'error' };
            const p = fmt === 'General' ? GENERAL_PREP : getPrep(fmt);
            const text = numberText(p, v, date1904);
            let color: string | undefined;
            if (p.colors !== null) color = p.colors[pickIndex(p, v)];
            return { text, color, kind: p.isDate ? 'date' : 'number' };
        }
        case 's': {
            const v = cell.v;
            const s = textOf(v);
            if (fmt === 'General') return { text: s, color: undefined, kind: 'text' };
            const p = getPrep(fmt);
            if (!p.hasAt || p.textOnly) return { text: s, color: undefined, kind: 'text' };
            let t: string | null | undefined = p.cache.get('s' + s);
            if (t === undefined) {
                t = ssfRun(p, s, false);
                if (t === null) t = s;
                if (p.cache.size >= CACHE_MAX) p.cache.clear();
                p.cache.set('s' + s, t);
            }
            return { text: t, color: p.textColor, kind: 'text' };
        }
        case 'b':
            return { text: cell.v ? 'TRUE' : 'FALSE', color: undefined, kind: 'bool' };
        case 'e':
            return { text: textOf(cell.v), color: undefined, kind: 'error' };
        case 'd': {
            let v = cell.v;
            if (typeof v === 'string') v = isoToSerial(v, !!date1904);
            else if (v instanceof Date) v = dateToSerial(v, !!date1904);
            if (typeof v !== 'number' || v !== v) return { text: textOf(cell.v), color: undefined, kind: 'text' };
            let p = fmt === 'General' ? null : getPrep(fmt);
            if (p === null || !p.isDate) p = prepare('yyyy-mm-dd' + (v % 1 ? ' hh:mm:ss' : ''));
            const text = numberText(p, v, date1904);
            return { text, color: p.colors !== null ? p.colors[pickIndex(p, v)] : undefined, kind: 'date' };
        }
        default:
            return { text: '', color: undefined, kind: 'text' };
    }
}

// ---------- input parsing ----------
const ERRORS = new Set(['#N/A', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#NULL!', '#SPILL!', '#CALC!']);
const NUM_RE = /^[+-]?\$?(?:(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/;
const PAREN_RE = /^\(\$?(?:(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d*)?|\.\d+)\)$/;
const PCT_RE = /^[+-]?(?:(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d*)?|\.\d+)\s*%$/;
const TIME_RE = /^(\d{1,4}):(\d{2})(?::(\d{2})(?:\.\d+)?)?(?:\s*([AaPp])[Mm])?$/;
const ISO_RE = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T ]+(\d{1,2}):(\d{2})(?::(\d{2})(?:\.(\d+))?)?)?Z?$/;
const TAIL = '(?:\\s+(\\d{1,2}:\\d{2}(?::\\d{2})?(?:\\s*[AaPp][Mm])?))?';
const SLASH_RE = new RegExp('^(\\d{1,2})[/-](\\d{1,2})[/-](\\d{2,4})' + TAIL + '$');
const DOT_RE = new RegExp('^(\\d{1,2})\\.(\\d{1,2})\\.(\\d{2,4})' + TAIL + '$');
const MDY_RE = new RegExp('^([A-Za-z]{3,9})\\.?\\s+(\\d{1,2})(?:st|nd|rd|th)?,?\\s+(\\d{4})' + TAIL + '$');
const DMY_RE = new RegExp('^(\\d{1,2})[\\s-]([A-Za-z]{3,9})\\.?[\\s,-]+(\\d{2,4})' + TAIL + '$');
const MY_RE = new RegExp('^([A-Za-z]{3,9})\\.?[\\s-]+(\\d{4})$');

function monthOf(name: string): number {
    const l = name.toLowerCase();
    for (let i = 0; i < 12; i++) {
        const f = MONTHS[i].toLowerCase();
        if (l.length >= 3 && f.startsWith(l)) return i + 1;
    }
    return 0;
}
function fullYear(s: string): number {
    const y = +s;
    if (s.length === 4) return y;
    if (s.length === 3) return -1;
    return y < 30 ? 2000 + y : 1900 + y;
}
function timeFrac(hs: string, mis: string, ss: string | undefined, ap?: string): number {
    let h = +hs;
    const mi = +mis,
        s = ss ? +ss : 0;
    if (mi > 59 || s > 59) return -1;
    if (ap) {
        if (h < 1 || h > 12) return -1;
        h = (h % 12) + (ap === 'p' || ap === 'P' ? 12 : 0);
    }
    return (h * 3600 + mi * 60 + s) / 86400;
}
function tailFrac(t: string | undefined): number {
    if (!t) return 0;
    const m = TIME_RE.exec(t);
    return m ? timeFrac(m[1], m[2], m[3], m[4]) : -1;
}
function numLit(s: string) {
    let neg = false;
    if (s.charCodeAt(0) === 43) s = s.slice(1);
    else if (s.charCodeAt(0) === 45) {
        neg = true;
        s = s.slice(1);
    }
    if (s.charCodeAt(0) === 36) s = s.slice(1);
    if (s.indexOf(',') >= 0) s = s.replace(/,/g, '');
    const v = Number(s);
    return neg ? -v : v;
}
function dateInput(s: string, d1904: boolean, allowLoose: boolean): { v: number; fmt?: string } | null {
    let m = ISO_RE.exec(s);
    if (m) {
        const day = ymdToSerial(+m[1], +m[2], +m[3], d1904);
        if (day < 0) return null;
        if (m[4] === undefined) return { v: day, fmt: 'yyyy-mm-dd' };
        const f = timeFrac(m[4], m[5], m[6]);
        if (f < 0 || +m[4] > 23) return null;
        let v = day + f;
        if (m[7]) v += Number('0.' + m[7]) / 86400;
        return { v, fmt: 'yyyy-mm-dd hh:mm:ss' };
    }
    if (!allowLoose) {
        // General cells: only unambiguous full m/d/yyyy dates convert (Excel applies m/d/yyyy)
        m = SLASH_RE.exec(s);
        if (!m || m[4] !== undefined || m[3].length !== 4) return null;
        const day = ymdToSerial(+m[3], +m[1], +m[2], d1904);
        return day < 0 ? null : { v: day, fmt: 'm/d/yyyy' };
    }
    let y: number, mo: number, d: number, tail: string | undefined;
    if ((m = SLASH_RE.exec(s))) {
        mo = +m[1];
        d = +m[2];
        y = fullYear(m[3]);
        tail = m[4];
    } else if ((m = DOT_RE.exec(s))) {
        d = +m[1];
        mo = +m[2];
        y = fullYear(m[3]);
        tail = m[4];
    } else if ((m = MDY_RE.exec(s))) {
        mo = monthOf(m[1]);
        d = +m[2];
        y = +m[3];
        tail = m[4];
    } else if ((m = DMY_RE.exec(s))) {
        d = +m[1];
        mo = monthOf(m[2]);
        y = fullYear(m[3]);
        tail = m[4];
    } else if ((m = MY_RE.exec(s))) {
        mo = monthOf(m[1]);
        d = 1;
        y = +m[2];
        tail = undefined;
    } else return null;
    if (!mo || y < 0) return null;
    const day = ymdToSerial(y, mo, d, d1904);
    if (day < 0) return null;
    const t = tailFrac(tail);
    if (t < 0) return null;
    return { v: day + t };
}

function parseInput(
    input: string | number | boolean | null | undefined,
    _prevCell?: unknown,
    style?: FormatStyle | null,
    date1904?: boolean
): ParsedInput | null {
    if (input === '' || input === null || input === undefined) return null;
    const text = typeof input === 'string' ? input : String(input);
    const fmt = style && style.fmt ? style.fmt : 'General';
    const p = fmt === 'General' ? GENERAL_PREP : getPrep(fmt);
    const c0 = text.charCodeAt(0);
    if (p.textOnly) return { t: 's', v: text };
    if (c0 === 61 && text.length > 1) return { t: 'n', v: null, f: text.slice(1) };
    if (c0 === 39) return { t: 's', v: text.slice(1) };
    const s = text.trim();
    if (s === '') return { t: 's', v: text };
    const sc = s.charCodeAt(0),
        n = s.length;
    if (sc === 35) {
        const u = s.toUpperCase();
        if (ERRORS.has(u)) return { t: 'e', v: u };
        return { t: 's', v: text };
    }
    if (n === 4 || n === 5) {
        const u = s.toUpperCase();
        if (u === 'TRUE') return { t: 'b', v: true };
        if (u === 'FALSE') return { t: 'b', v: false };
    }
    // numbers
    if ((sc >= 48 && sc <= 57) || sc === 43 || sc === 45 || sc === 46 || sc === 36 || sc === 40) {
        if (n <= 400) {
            if (NUM_RE.test(s)) {
                const v = numLit(s);
                if (isFinite(v)) {
                    if (v === 0) return { t: 'n', v: 0 };
                    return { t: 'n', v };
                }
                return { t: 's', v: text };
            }
            if (sc === 40 && PAREN_RE.test(s)) {
                const v = numLit(s.slice(1, -1));
                return isFinite(v) ? { t: 'n', v: v === 0 ? 0 : -v } : { t: 's', v: text };
            }
            if (s.charCodeAt(n - 1) === 37 && PCT_RE.test(s)) {
                const body = s.slice(0, -1).trim();
                const clean = body.replace(/,/g, '');
                const v = Number(clean + 'e-2');
                if (isFinite(v)) {
                    const o: ParsedNumber = { t: 'n', v: v === 0 ? 0 : v };
                    if (fmt === 'General') o.fmt = /\.\d*[1-9]/.test(clean) ? '0.00%' : '0%';
                    return o;
                }
            }
        }
    }
    // dates and times
    if ((sc >= 48 && sc <= 57) || (sc >= 65 && sc <= 90) || (sc >= 97 && sc <= 122)) {
        if (n <= 64) {
            const tm = sc >= 48 && sc <= 57 ? TIME_RE.exec(s) : null;
            if (tm) {
                const f = timeFrac(tm[1], tm[2], tm[3], tm[4]);
                if (f >= 0) {
                    const o: ParsedNumber = { t: 'n', v: f };
                    if (!p.isDate)
                        o.fmt =
                            (+tm[1] >= 24 && !tm[4] ? '[h]:mm' : 'h:mm') +
                            (tm[3] ? ':ss' : '') +
                            (tm[4] ? ' AM/PM' : '');
                    return o;
                }
            } else {
                const r = dateInput(s, !!date1904, p.isDate);
                if (r) {
                    const o: ParsedNumber = { t: 'n', v: r.v };
                    if (!p.isDate && r.fmt) o.fmt = r.fmt;
                    return o;
                }
            }
        }
    }
    return { t: 's', v: text };
}

// ---------- export ----------
function pad2(n: number): string {
    return n < 10 ? '0' + n : '' + n;
}
function isoOf(v: number, d1904: boolean, p: PreparedFormat): string | null {
    if (!(v >= 0) || v > (d1904 ? MAX_SERIAL_1904 : MAX_SERIAL_1900)) return null;
    let days = Math.floor(v),
        secs = Math.round((v - days) * 86400);
    if (secs >= 86400) {
        secs -= 86400;
        days++;
    }
    const hh = Math.floor(secs / 3600),
        mm = Math.floor(secs / 60) % 60,
        ss = secs % 60;
    const tstr = pad2(hh) + ':' + pad2(mm) + ':' + pad2(ss);
    if (p.timePart && !p.datePart) return v < 1 ? tstr : hh + days * 24 + ':' + pad2(mm) + ':' + pad2(ss);
    if (!d1904 && days === 60) return '1900-02-29' + (secs ? ' ' + tstr : '');
    if (!d1904 && days === 0) return secs ? tstr : String(v);
    const d = serialToDate(days, d1904);
    const ds = d.getUTCFullYear() + '-' + pad2(d.getUTCMonth() + 1) + '-' + pad2(d.getUTCDate());
    return secs ? ds + ' ' + tstr : ds;
}
function looksTyped(s: string, _style: unknown, d1904?: boolean): boolean {
    const c0 = s.charCodeAt(0);
    if (c0 === 39 || (c0 === 61 && s.length > 1)) return true;
    const r = parseInput(s, null, { fmt: 'm/d/yyyy' }, d1904);
    return r !== null && r.t !== 's';
}
function exportText(cell: FormatCell | null | undefined, style?: FormatStyle | null, date1904?: boolean): string {
    if (!cell) return '';
    if (cell.f !== undefined && cell.f !== null && cell.f !== '') return '=' + cell.f;
    const v = cell.v;
    switch (cell.t) {
        case 'n': {
            if (typeof v !== 'number') return textOf(v);
            if (v !== v || v === Infinity || v === -Infinity) return '#NUM!';
            const fmt = style && style.fmt ? style.fmt : 'General';
            if (fmt !== 'General') {
                const p = getPrep(fmt);
                if (p.isDate) {
                    const s = isoOf(v, !!date1904, p);
                    if (s !== null) return s;
                } else if (p.pct && !p.broken) {
                    const m = Number((v * 100).toPrecision(15));
                    if (Math.abs(m) < 1e15) return String(m) + '%';
                }
            }
            if (v === 0) return '0';
            if (v % 1 !== 0) {
                const q = +v.toPrecision(15);
                if (q !== 0) return String(q);
            }
            return String(v);
        }
        case 's': {
            const s = textOf(v);
            if (s === '') return s;
            if (style && style.fmt === '@') return s;
            return looksTyped(s, style, date1904) ? "'" + s : s;
        }
        case 'b':
            return v ? 'TRUE' : 'FALSE';
        case 'e':
            return textOf(v);
        case 'd': {
            if (typeof v === 'string') return v;
            if (v instanceof Date) return v.toISOString().slice(0, 10);
            return textOf(v);
        }
        default:
            return '';
    }
}

// ---------- stats ----------
function createStats(): Stats {
    return {
        count: 0,
        sum: 0,
        min: Infinity,
        max: -Infinity,
        add(n: number) {
            this.count++;
            this.sum += n;
            if (n < this.min) this.min = n;
            if (n > this.max) this.max = n;
        },
        avg() {
            return this.count ? this.sum / this.count : NaN;
        }
    };
}
function isNumericCell(cell: FormatCell | null | undefined): boolean {
    return (
        !!cell &&
        cell.t === 'n' &&
        typeof cell.v === 'number' &&
        cell.v === cell.v &&
        cell.v !== Infinity &&
        cell.v !== -Infinity
    );
}

export {
    colName,
    colIndex,
    cellRef,
    parseRef,
    parseRange,
    serialToDate,
    dateToSerial,
    formatCell,
    parseInput,
    exportText,
    createStats,
    isNumericCell,
    generalNum as formatGeneral
};
export function isDateFormat(fmt: string) {
    return !!fmt && fmt !== 'General' && prepare(fmt).isDate;
}
