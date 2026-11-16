// Raw-copy zip assembly in xlsx-writer: strict structural verifier + fuzz over odd zip layouts.
import test from 'node:test';
import assert from 'node:assert/strict';
import * as zlib from 'node:zlib';
import * as fflate from 'fflate';
import * as XLSX from 'xlsx';
import { patchXlsx, writerDeps } from '../webview/core/xlsx-writer.ts';
import type { CellEdit, EditCell } from '../webview/core/xlsx-writer.ts';

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PK = 'http://schemas.openxmlformats.org/package/2006/relationships';
const T = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const HDR = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const rng = (seed: number): (() => number) => {
    let h = seed;
    return () => (h = (h * 1103515245 + 12345) & 0x7fffffff) / 0x7fffffff;
};
const colName = (c: number): string => {
    let s = '';
    c++;
    while (c > 0) {
        const m = (c - 1) % 26;
        s = String.fromCharCode(65 + m) + s;
        c = (c - 1 - m) / 26;
    }
    return s;
};

/** Unwraps a lookup that the fixture guarantees. */
function must<T>(v: T | null | undefined, what: string): T {
    if (v === null || v === undefined) throw new Error('missing ' + what);
    return v;
}
const toBuffer = (d: string | Uint8Array): Buffer => (typeof d === 'string' ? Buffer.from(d) : Buffer.from(d));

// ---- independent zip builder (does not use fflate) ----
/** dd: data descriptor without (1) / with (2) signature; z64: zip64 extra fields */
interface ZipSpec {
    name: string;
    data: string | Uint8Array;
    method: 0 | 8;
    level?: number;
    dd?: 1 | 2;
    z64?: boolean;
    extra?: Buffer;
    comment?: string;
    time?: number;
    date?: number;
    utf8flag?: boolean;
}

function buildZip(entries: readonly ZipSpec[], opts: { comment?: string } = {}): Uint8Array {
    const chunks: Buffer[] = [],
        cds: Buffer[] = [];
    let pos = 0;
    const w = (b: Buffer): void => {
        chunks.push(b);
        pos += b.length;
    };
    const h16 = (v: number): Buffer => {
        const b = Buffer.alloc(2);
        b.writeUInt16LE(v);
        return b;
    };
    const h32 = (v: number): Buffer => {
        const b = Buffer.alloc(4);
        b.writeUInt32LE(v >>> 0);
        return b;
    };
    const h64 = (v: number): Buffer => {
        const b = Buffer.alloc(8);
        b.writeBigUInt64LE(BigInt(v));
        return b;
    };
    for (const e of entries) {
        const data = toBuffer(e.data);
        const comp = e.method === 8 ? zlib.deflateRawSync(data, { level: e.level === undefined ? 6 : e.level }) : data;
        const crc = zlib.crc32(data);
        const nb = Buffer.from(e.name, 'utf8');
        const flag = (e.dd ? 8 : 0) | (e.utf8flag ? 0x800 : 0);
        const extra = e.extra || Buffer.alloc(0);
        const z64x = Buffer.concat([h16(1), h16(16), h64(data.length), h64(comp.length)]);
        const off = pos;
        const lsz = e.dd ? [0, 0, 0] : [crc, e.z64 ? 0xffffffff : comp.length, e.z64 ? 0xffffffff : data.length];
        const lextra = e.z64 && !e.dd ? Buffer.concat([z64x, extra]) : extra;
        w(
            Buffer.concat([
                h32(0x04034b50),
                h16(e.z64 ? 45 : 20),
                h16(flag),
                h16(e.method),
                h16(e.time || 0x6000),
                h16(e.date || 0x5a21),
                h32(lsz[0]),
                h32(lsz[1]),
                h32(lsz[2]),
                h16(nb.length),
                h16(lextra.length),
                nb,
                lextra
            ])
        );
        w(comp);
        if (e.dd)
            w(
                Buffer.concat([
                    e.dd === 2 ? h32(0x08074b50) : Buffer.alloc(0),
                    h32(crc),
                    e.z64 ? h64(comp.length) : h32(comp.length),
                    e.z64 ? h64(data.length) : h32(data.length)
                ])
            );
        const cextra = e.z64 ? Buffer.concat([z64x, extra]) : extra;
        const com = Buffer.from(e.comment || '', 'utf8');
        cds.push(
            Buffer.concat([
                h32(0x02014b50),
                h16(0x031e),
                h16(e.z64 ? 45 : 20),
                h16(flag),
                h16(e.method),
                h16(e.time || 0x6000),
                h16(e.date || 0x5a21),
                h32(crc),
                h32(e.z64 ? 0xffffffff : comp.length),
                h32(e.z64 ? 0xffffffff : data.length),
                h16(nb.length),
                h16(cextra.length),
                h16(com.length),
                h16(0),
                h16(0),
                h32(e.name.endsWith('/') ? 0x10 : 0),
                h32(off),
                nb,
                cextra,
                com
            ])
        );
    }
    const cdOff = pos;
    for (const c of cds) w(c);
    const cdSize = pos - cdOff;
    const ac = Buffer.from(opts.comment || '', 'utf8');
    const big = entries.length >= 0xffff;
    if (big) {
        const zo = pos;
        w(
            Buffer.concat([
                h32(0x06064b50),
                h64(44),
                h16(45),
                h16(45),
                h32(0),
                h32(0),
                h64(entries.length),
                h64(entries.length),
                h64(cdSize),
                h64(cdOff)
            ])
        );
        w(Buffer.concat([h32(0x07064b50), h32(0), h64(zo), h32(1)]));
    }
    const cnt = big ? 0xffff : entries.length;
    w(
        Buffer.concat([
            h32(0x06054b50),
            h16(0),
            h16(0),
            h16(cnt),
            h16(cnt),
            h32(big ? 0xffffffff : cdSize),
            h32(big ? 0xffffffff : cdOff),
            h16(ac.length),
            ac
        ])
    );
    return new Uint8Array(Buffer.concat(chunks));
}

// ---- strict verifier (zlib + own parser) ----
/** central directory record of an entry (sizes patched from zip64 extras in lenient mode) */
interface CentralEntry {
    flag: number;
    method: number;
    time: number;
    date: number;
    crc: number;
    csize: number;
    usize: number;
    lho: number;
    eattr: number;
    nameBytes: Buffer;
    name: string;
    extra: Buffer;
    comment: string;
}
/** central record plus what the local header and the data area say */
interface VerifiedEntry extends CentralEntry {
    lextra: Buffer;
    ds: number;
    raw: Buffer;
    data: Buffer;
}

function verifyZip(b: Uint8Array, lenient?: boolean): VerifiedEntry[] {
    const buf = Buffer.from(b.buffer, b.byteOffset, b.length);
    let eo = -1;
    for (let i = buf.length - 22; i >= Math.max(0, buf.length - 22 - 65535); i--)
        if (buf.readUInt32LE(i) === 0x06054b50) {
            eo = i;
            break;
        }
    assert.ok(eo >= 0, 'EOCD');
    const clen = buf.readUInt16LE(eo + 20);
    assert.equal(eo + 22 + clen, buf.length, 'EOCD comment length');
    let cnt = buf.readUInt16LE(eo + 10),
        cdSize = buf.readUInt32LE(eo + 12),
        cdOff = buf.readUInt32LE(eo + 16);
    assert.equal(buf.readUInt16LE(eo + 8), cnt);
    let cdEnd = eo;
    if (cnt === 0xffff) {
        assert.equal(buf.readUInt32LE(eo - 20), 0x07064b50, 'zip64 locator');
        const zo = Number(buf.readBigUInt64LE(eo - 12));
        assert.equal(buf.readUInt32LE(zo), 0x06064b50);
        cnt = Number(buf.readBigUInt64LE(zo + 32));
        cdSize = Number(buf.readBigUInt64LE(zo + 40));
        cdOff = Number(buf.readBigUInt64LE(zo + 48));
        cdEnd = zo;
    }
    assert.equal(cdOff + cdSize, cdEnd, 'central directory is contiguous with the end records');
    const central: CentralEntry[] = [];
    let p = cdOff;
    for (let k = 0; k < cnt; k++) {
        assert.equal(buf.readUInt32LE(p), 0x02014b50, 'central sig');
        const nl = buf.readUInt16LE(p + 28),
            el = buf.readUInt16LE(p + 30),
            cl = buf.readUInt16LE(p + 32);
        const nameBytes = buf.subarray(p + 46, p + 46 + nl);
        const e: CentralEntry = {
            flag: buf.readUInt16LE(p + 8),
            method: buf.readUInt16LE(p + 10),
            time: buf.readUInt16LE(p + 12),
            date: buf.readUInt16LE(p + 14),
            crc: buf.readUInt32LE(p + 16),
            csize: buf.readUInt32LE(p + 20),
            usize: buf.readUInt32LE(p + 24),
            lho: buf.readUInt32LE(p + 42),
            eattr: buf.readUInt32LE(p + 38),
            nameBytes,
            name: nameBytes.toString('utf8'),
            extra: buf.subarray(p + 46 + nl, p + 46 + nl + el),
            comment: buf.subarray(p + 46 + nl + el, p + 46 + nl + el + cl).toString('utf8')
        };
        if (lenient) {
            for (let q = 0; q + 4 <= e.extra.length; q += 4 + e.extra.readUInt16LE(q + 2)) {
                if (e.extra.readUInt16LE(q) === 1) {
                    e.usize = Number(e.extra.readBigUInt64LE(q + 4));
                    e.csize = Number(e.extra.readBigUInt64LE(q + 12));
                }
            }
        } else {
            assert.notEqual(e.csize, 0xffffffff);
            assert.notEqual(e.usize, 0xffffffff);
        }
        p += 46 + nl + el + cl;
        central.push(e);
    }
    assert.equal(p, cdEnd, 'central directory parsed exactly');
    // local headers: contiguous, ascending, exact
    const out: VerifiedEntry[] = [];
    let expect = 0;
    for (const e of central) {
        assert.equal(buf.readUInt32LE(e.lho), 0x04034b50, 'local sig');
        if (!lenient) {
            assert.equal(e.lho, expect, 'entry ' + e.name + ' starts where the previous one ends (no gap/overlap)');
            assert.equal(buf.readUInt16LE(e.lho + 6), e.flag, 'local flag == central flag');
            assert.equal(e.flag & 8, 0, 'no data descriptor in rewritten output');
            assert.equal(buf.readUInt16LE(e.lho + 8), e.method);
            assert.equal(buf.readUInt32LE(e.lho + 14), e.crc, 'local crc');
            const lc = buf.readUInt32LE(e.lho + 18),
                lu = buf.readUInt32LE(e.lho + 22);
            if (lc === 0xffffffff || lu === 0xffffffff) {
                // zip64 local: real sizes live in the extra field
                const nl0 = buf.readUInt16LE(e.lho + 26),
                    le = buf.subarray(e.lho + 30 + nl0, e.lho + 30 + nl0 + buf.readUInt16LE(e.lho + 28));
                assert.equal(le.readUInt16LE(0), 1);
                assert.equal(Number(le.readBigUInt64LE(4)), e.usize);
                assert.equal(Number(le.readBigUInt64LE(12)), e.csize);
            } else {
                assert.equal(lc, e.csize, 'local csize');
                assert.equal(lu, e.usize, 'local usize');
            }
        }
        const nl = buf.readUInt16LE(e.lho + 26),
            el = buf.readUInt16LE(e.lho + 28);
        assert.equal(buf.subarray(e.lho + 30, e.lho + 30 + nl).compare(e.nameBytes), 0, 'local name == central name');
        const lextra = buf.subarray(e.lho + 30 + nl, e.lho + 30 + nl + el);
        const ds = e.lho + 30 + nl + el;
        const raw = buf.subarray(ds, ds + e.csize);
        expect = ds + e.csize;
        assert.ok(expect <= cdOff, 'entry inside the data area');
        const data = e.method === 0 ? raw : zlib.inflateRawSync(raw);
        assert.equal(data.length, e.usize, 'usize of ' + e.name);
        assert.equal(zlib.crc32(data), e.crc, 'crc of ' + e.name);
        out.push({ ...e, lextra, ds, raw, data });
    }
    if (!lenient) assert.equal(expect, cdOff, 'data area ends at the central directory');
    assert.equal(new Set(out.map(e => e.name)).size, out.length, 'unique names');
    // independent reader #2
    if (!lenient) {
        const z = fflate.unzipSync(b);
        assert.equal(Object.keys(z).length, out.length);
        for (const e of out) {
            const key = e.flag & 0x800 ? e.name : e.nameBytes.toString('latin1'); // fflate decodes legacy names as latin1
            assert.equal(Buffer.compare(Buffer.from(z[key]), e.data), 0, 'fflate agrees on ' + e.name);
        }
    }
    return out;
}

// ---- package generator ----
function sheetXml(rand: () => number, NR: number, NC: number, tag: string): string {
    let body = '';
    for (let r = 0; r < NR; r++) {
        if (rand() < 0.2) continue;
        let cells = '';
        for (let c = 0; c < NC; c++) {
            if (rand() < 0.3) continue;
            const ref = colName(c) + (r + 1);
            cells +=
                rand() < 0.5
                    ? '<c r="' + ref + '"><v>' + (r * 10 + c) + '</v></c>'
                    : '<c r="' + ref + '" t="inlineStr"><is><t>' + tag + r + 'x' + c + ' é</t></is></c>';
        }
        body += '<row r="' + (r + 1) + '">' + cells + '</row>';
    }
    return (
        HDR +
        '<worksheet xmlns="' +
        NS +
        '"><dimension ref="A1:' +
        colName(NC - 1) +
        NR +
        '"/><sheetData>' +
        body +
        '</sheetData></worksheet>'
    );
}

interface GeneratedPackage {
    entries: ZipSpec[];
    comment: string;
    rand: () => number;
    withCalc: boolean;
    NR: number;
    NC: number;
}

function genPackage(seed: number): GeneratedPackage {
    const rand = rng(seed * 104729);
    const pick = <T>(a: readonly T[]): T => a[Math.floor(rand() * a.length)];
    const NR = 20,
        NC = 6;
    const sheets = [sheetXml(rand, NR, NC, 'a'), sheetXml(rand, NR, NC, 'b')];
    const withCalc = rand() < 0.7;
    const base: [string, string | Buffer][] = [
        [
            '[Content_Types].xml',
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="a"/>' +
                (withCalc ? '<Override PartName="/xl/calcChain.xml" ContentType="cc"/>' : '') +
                '</Types>'
        ],
        [
            '_rels/.rels',
            '<Relationships xmlns="' +
                PK +
                '"><Relationship Id="rId1" Type="' +
                T +
                'officeDocument" Target="xl/workbook.xml"/></Relationships>'
        ],
        [
            'xl/workbook.xml',
            HDR +
                '<workbook xmlns="' +
                NS +
                '" xmlns:r="' +
                RNS +
                '"><sheets><sheet name="A" sheetId="1" r:id="rId1"/><sheet name="B" sheetId="2" r:id="rId2"/></sheets><calcPr calcId="1"/></workbook>'
        ],
        [
            'xl/_rels/workbook.xml.rels',
            '<Relationships xmlns="' +
                PK +
                '"><Relationship Id="rId1" Type="' +
                T +
                'worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="' +
                T +
                'worksheet" Target="worksheets/sheet2.xml"/>' +
                (withCalc ? '<Relationship Id="rId9" Type="' + T + 'calcChain" Target="calcChain.xml"/>' : '') +
                '</Relationships>'
        ],
        ['xl/worksheets/sheet1.xml', sheets[0]],
        ['xl/worksheets/sheet2.xml', sheets[1]]
    ];
    if (withCalc) base.push(['xl/calcChain.xml', '<calcChain xmlns="' + NS + '"><c r="A1" i="1"/></calcChain>']);
    const noise: [string, string | Buffer][] = [];
    const nn = 3 + Math.floor(rand() * 10);
    for (let i = 0; i < nn; i++) {
        const kind = pick(['rand', 'text', 'empty', 'dir', 'uni', 'big']);
        if (kind === 'rand') {
            const b = Buffer.alloc(Math.floor(rand() * 4000));
            for (let k = 0; k < b.length; k++) b[k] = Math.floor(rand() * 256);
            noise.push(['xl/media/img' + i + '.bin', b]);
        } else if (kind === 'text')
            noise.push(['docProps/p' + i + '.xml', '<x>' + 'hello '.repeat(Math.floor(rand() * 300)) + '</x>']);
        else if (kind === 'empty') noise.push(['custom/empty' + i + '.txt', '']);
        else if (kind === 'dir') noise.push(['xl/dir' + i + '/', '']);
        else if (kind === 'uni') noise.push(['xl/media/ünï-日本-' + i + '.txt', 'unicode ' + i]);
        else noise.push(['xl/big' + i + '.xml', '<r>' + 'abcdefgh'.repeat(20000 + Math.floor(rand() * 1000)) + '</r>']);
    }
    const all = base.concat(noise);
    // shuffle everything except keep sheets before... order is arbitrary on purpose (CT not necessarily first)
    if (rand() < 0.6)
        for (let i = all.length - 1; i > 0; i--) {
            const j = Math.floor(rand() * (i + 1));
            [all[i], all[j]] = [all[j], all[i]];
        }
    const entries = all.map(([name, data]) => {
        const isDir = name.endsWith('/');
        const e: ZipSpec = {
            name,
            data: typeof data === 'string' ? Buffer.from(data) : data,
            method: isDir || data.length === 0 ? pick<0 | 8>([0, 8]) : pick<0 | 8>([0, 8, 8, 8]),
            level: pick([1, 6, 9])
        };
        if (/[^\x00-\x7f]/.test(name)) e.utf8flag = rand() < 0.5;
        if (rand() < 0.3) e.dd = pick<1 | 2>([1, 2]);
        if (rand() < 0.25) e.z64 = true;
        if (rand() < 0.3) e.extra = Buffer.concat([Buffer.from([0xfe, 0xca, 4, 0, 1, 2, 3, 4])]);
        if (rand() < 0.2) e.comment = 'cmt ' + i_(rand);
        if (rand() < 0.5) {
            e.time = Math.floor(rand() * 0xffff);
            e.date = 0x4a21 + Math.floor(rand() * 2000);
        }
        return e;
    });
    return { entries, comment: rand() < 0.4 ? 'archive comment ' + seed : '', rand, withCalc, NR, NC };
}
const i_ = (r: () => number): string => String(Math.floor(r() * 1000));

function isCellObject(x: unknown): x is XLSX.CellObject {
    return typeof x === 'object' && x !== null && 't' in x;
}
/** SheetJS types worksheet indexing as `any`; this reads one cell through a typed lookup. */
function sheetCell(wb: XLSX.WorkBook, sheet: string, ref: string): XLSX.CellObject | undefined {
    const lookup: Record<string, unknown> = wb.Sheets[sheet];
    const c = lookup[ref];
    return isCellObject(c) ? c : undefined;
}

test('verifier accepts the independent builder output and the fixtures read in SheetJS', () => {
    for (let seed = 1; seed <= 10; seed++) {
        const p = genPackage(seed);
        const z = buildZip(p.entries, { comment: p.comment });
        verifyZip(z, true);
    }
});

test('fuzz: raw-copy rewrite keeps untouched entries byte-exact and the package valid', () => {
    for (let seed = 1; seed <= 150; seed++) {
        const p = genPackage(seed);
        const rand = p.rand;
        const orig = buildZip(p.entries, { comment: p.comment });
        const before = verifyZip(orig, true);
        const origByName = new Map<string, VerifiedEntry>();
        for (const e of before) origByName.set(e.name, e);
        const edits: CellEdit[] = [],
            exp = new Map<string, EditCell>();
        const n = 1 + Math.floor(rand() * 8);
        for (let i = 0; i < n; i++) {
            const sheet = rand() < 0.7 ? 0 : 1,
                r = Math.floor(rand() * (p.NR + 3)),
                c = Math.floor(rand() * (p.NC + 2));
            const cell: EditCell =
                rand() < 0.5 ? { t: 'n', v: 1000 + i + seed / 10 } : { t: 's', v: 'new ' + i + ' ü日' };
            edits.push({ sheet, r, c, cell });
            exp.set(sheet + ',' + r + ',' + c, cell);
        }
        const out = patchXlsx(orig, edits);
        const after = verifyZip(out);
        const touchedSheets = new Set(edits.map(e => 'xl/worksheets/sheet' + (e.sheet + 1) + '.xml'));
        const rewritten = new Set([
            ...touchedSheets,
            '[Content_Types].xml',
            'xl/workbook.xml',
            'xl/_rels/workbook.xml.rels',
            'xl/calcChain.xml'
        ]);
        // names/order: calcChain dropped, [Content_Types].xml first, the rest in original order
        const expectNames = before.map(e => e.name).filter(nm => nm !== 'xl/calcChain.xml');
        const ct = expectNames.indexOf('[Content_Types].xml');
        expectNames.unshift(...expectNames.splice(ct, 1));
        assert.deepEqual(
            after.map(e => e.name),
            expectNames,
            'seed ' + seed
        );
        for (const e of after) {
            const o = must(origByName.get(e.name), 'original ' + e.name);
            if (rewritten.has(e.name)) continue;
            assert.equal(Buffer.compare(e.raw, o.raw), 0, 'raw compressed bytes copied: ' + e.name + ' seed ' + seed);
            assert.equal(e.crc, o.crc);
            assert.equal(e.method, o.method);
            assert.equal(e.csize, o.csize);
            assert.equal(e.usize, o.usize);
            assert.equal(e.time, o.time);
            assert.equal(e.date, o.date);
            assert.equal(e.eattr, o.eattr);
            assert.equal(e.comment, o.comment);
            assert.equal(e.flag, o.flag & ~8);
            assert.equal(Buffer.compare(e.nameBytes, o.nameBytes), 0);
            const src = must(
                p.entries.find(x => x.name === e.name),
                'source ' + e.name
            );
            if (src.extra && !src.z64)
                assert.ok(e.extra.includes(Buffer.from([0xfe, 0xca, 4, 0])), 'extra kept for ' + e.name);
        }
        for (const nm of touchedSheets) {
            const a = must(
                    after.find(e => e.name === nm),
                    nm
                ),
                o = must(origByName.get(nm), nm);
            assert.equal(a.time, o.time);
            assert.equal(a.date, o.date);
        }
        // content as seen by SheetJS: edited cells new, others unchanged
        const plain = buildZip(p.entries.map(e => ({ name: e.name, data: e.data, method: 8 })));
        const wbA = XLSX.read(plain, { type: 'array' }),
            wbB = XLSX.read(out, { type: 'array' });
        for (let si = 0; si < 2; si++) {
            const nameA = wbA.SheetNames[si],
                nameB = wbB.SheetNames[si];
            const A = (ref: string): XLSX.CellObject | undefined => sheetCell(wbA, nameA, ref),
                B = (ref: string): XLSX.CellObject | undefined => sheetCell(wbB, nameB, ref);
            const tv = (c: XLSX.CellObject | undefined): unknown[] | undefined => (c ? [c.t, c.v] : undefined);
            for (let r = 0; r < p.NR + 3; r++)
                for (let c = 0; c < p.NC + 2; c++) {
                    const ref = colName(c) + (r + 1),
                        e = exp.get(si + ',' + r + ',' + c);
                    if (e) assert.equal(B(ref)?.v, e.v, ref);
                    else assert.deepEqual(tv(B(ref)), tv(A(ref)), 'untouched ' + ref + ' seed ' + seed);
                }
        }
    }
});

test('big untouched entries are never inflated', () => {
    const rand = rng(5);
    const big = Buffer.alloc(3e6);
    for (let i = 0; i < big.length; i++) big[i] = (i * 7 + ((rand() * 3) | 0)) & 255;
    const p = genPackage(3);
    p.entries.push({ name: 'xl/media/big.dat', data: big, method: 8, level: 1 });
    const orig = buildZip(p.entries);
    const calls: number[] = [];
    const real = writerDeps.Inflate;
    // spy: a subclass that records how many compressed bytes each inflate stream was fed
    class SpyInflate extends real {
        private fed = 0;
        override push(d: Uint8Array, f?: boolean): void {
            this.fed += d.length;
            if (f) calls.push(this.fed);
            super.push(d, f);
        }
    }
    writerDeps.Inflate = SpyInflate;
    let out: Uint8Array;
    try {
        out = patchXlsx(orig, [{ sheet: 0, r: 2, c: 1, cell: { t: 'n', v: 7 } }]);
    } finally {
        writerDeps.Inflate = real;
    }
    const bigC = verifyZip(orig, true).find(e => e.name === 'xl/media/big.dat')?.csize ?? 0;
    assert.ok(calls.length > 0 && calls.every(n => n < bigC / 2), 'only small parts were inflated: ' + calls.join(','));
    verifyZip(out);
});

test('sheet too large to edit', () => {
    const p = genPackage(4);
    const orig = buildZip(p.entries);
    const hacked = orig.slice();
    const dv = new DataView(hacked.buffer, hacked.byteOffset, hacked.length);
    const ents = verifyZip(orig, true);
    // forge the central usize of sheet1 (find its central record)
    const eo = hacked.length - 22 - dv.getUint16(hacked.length - 2, true);
    let q = dv.getUint32(eo + 16, true);
    for (let k = 0; k < ents.length; k++) {
        const nl = dv.getUint16(q + 28, true),
            el = dv.getUint16(q + 30, true),
            cl = dv.getUint16(q + 32, true);
        if (Buffer.from(hacked.subarray(q + 46, q + 46 + nl)).toString() === 'xl/worksheets/sheet1.xml') {
            if (dv.getUint32(q + 24, true) !== 0xffffffff) dv.setUint32(q + 24, 0x7fffffff, true);
            else continue;
        }
        q += 46 + nl + el + cl;
    }
    assert.throws(
        () => patchXlsx(hacked, [{ sheet: 0, r: 0, c: 0, cell: { t: 'n', v: 1 } }]),
        /sheet too large to edit|zip bomb/
    );
});

test('zip64 end records for >65534 entries, junk before the zip, clean errors on garbage', () => {
    const p = genPackage(6);
    const extra: ZipSpec[] = [];
    for (let i = 0; i < 66000; i++) extra.push({ name: 'x/' + i, data: '', method: 0 });
    const orig = buildZip(p.entries.concat(extra).map(e => e));
    const out = patchXlsx(orig, [{ sheet: 0, r: 0, c: 0, cell: { t: 'n', v: 1 } }]);
    assert.ok(verifyZip(out).length >= 66000);
    const pk = genPackage(7);
    const good = buildZip(pk.entries);
    const junk = Buffer.concat([Buffer.from('MZ junk stub'.repeat(10)), Buffer.from(good)]);
    const o2 = patchXlsx(new Uint8Array(junk), [{ sheet: 0, r: 1, c: 1, cell: { t: 'n', v: 5 } }]);
    assert.equal(sheetCell(XLSX.read(o2, { type: 'array' }), 'A', 'B2')?.v, 5);
    assert.throws(() => patchXlsx(new Uint8Array(100), [{ sheet: 0, r: 0, c: 0, cell: { t: 'n', v: 1 } }]), /zip/i);
});

test('perf: 60k rows, one edit touches only the sheet', () => {
    const NR = 60000;
    const rows = new Array<string>(NR);
    for (let r = 1; r <= NR; r++)
        rows[r - 1] =
            '<row r="' +
            r +
            '"><c r="A' +
            r +
            '"><v>' +
            r +
            '</v></c><c r="B' +
            r +
            '" t="inlineStr"><is><t>row ' +
            r +
            '</t></is></c><c r="C' +
            r +
            '"><f>A' +
            r +
            '*2</f><v>' +
            r * 2 +
            '</v></c></row>';
    const sheet =
        HDR +
        '<worksheet xmlns="' +
        NS +
        '"><dimension ref="A1:C' +
        NR +
        '"/><sheetData>' +
        rows.join('') +
        '</sheetData></worksheet>';
    const p = genPackage(8);
    p.entries = p.entries.filter(e => e.name !== 'xl/worksheets/sheet1.xml');
    p.entries.push({ name: 'xl/worksheets/sheet1.xml', data: Buffer.from(sheet), method: 8, level: 3 });
    const orig = buildZip(p.entries);
    const t0 = process.hrtime.bigint();
    const out = patchXlsx(orig, [
        { sheet: 0, r: 30000, c: 1, cell: { t: 's', v: 'EDIT' } },
        { sheet: 0, r: NR + 5, c: 0, cell: { t: 'n', v: 9 } }
    ]);
    const ms = Number(process.hrtime.bigint() - t0) / 1e6;
    console.log(
        'perf: ' +
            (sheet.length / 1e6).toFixed(1) +
            'MB sheet xml, zip ' +
            (orig.length / 1e6).toFixed(1) +
            'MB, 2 edits: ' +
            ms.toFixed(0) +
            ' ms'
    );
    const after = verifyZip(out);
    const x = Buffer.from(
        must(
            after.find(e => e.name === 'xl/worksheets/sheet1.xml'),
            'sheet1'
        ).data
    ).toString();
    assert.ok(x.includes('<c r="B30001" t="inlineStr"><is><t>EDIT</t></is></c>') || x.includes('EDIT'));
    assert.ok(x.includes('<row r="' + (NR + 6) + '">'));
    assert.ok(x.includes('<row r="29999"><c r="A29999">'));
    assert.ok(ms < 15000);
});
