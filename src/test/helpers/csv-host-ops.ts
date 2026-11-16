import type { GridContextMenuEvent, GridMenuItem } from '../../webview/core/grid.ts';
import { defined, makeWebview, parseCsv } from './csv-host-sim.ts';
import type { FakeShell, WebviewHandle } from './csv-host-sim.ts';

export const tick = async (): Promise<void> => {
    for (let i = 0; i < 3; i++) {
        await new Promise(r => setImmediate(r));
    }
};
export const drain = async (w: WebviewHandle): Promise<void> => {
    await w.view['_queue'];
    await tick();
    await w.view['_queue'];
};
export function reparse(w: WebviewHandle) {
    const o = w.view.getOptions();
    return parseCsv(w.doc.text, {
        separator: o.separator,
        quote: '"',
        hasHeaders: o.hasHeaders,
        commentChar: '#',
        skipComments: o.skipComments
    });
}
export function state(w: WebviewHandle): string[][] {
    const p = defined(w.shell().provider, 'provider');
    const rows: string[][] = [];
    for (let r = 0; r < p.rowCount; r++) {
        const a: string[] = [];
        for (let c = 0; c < p.colCount; c++) a.push(p.getRaw(r, c));
        rows.push(a);
    }
    return rows;
}
export interface CheckResult {
    ok: boolean;
    msg: string;
}
export interface RunResult extends CheckResult {
    w: WebviewHandle;
}
export function check(name: string, w: WebviewHandle): CheckResult {
    // compare provider vs reparse
    const p = defined(w.shell().provider, 'provider');
    const pr = reparse(w);
    const exp = (pr.header ? [pr.header] : []).concat(pr.rows);
    const got = state(w);
    let ok = true;
    const msg: string[] = [];
    const norm = (a: string[][]): string[] =>
        a.map(r => {
            r = r.map(x => x.replace(/\r\n/g, '\n'));
            while (r.length < p.colCount) r.push('');
            return JSON.stringify(r);
        });
    const E = norm(exp),
        G = norm(got);
    if (E.length !== G.length || E.some((x, i) => x !== G[i])) {
        ok = false;
        msg.push('CELLS differ\n exp ' + E.join('|') + '\n got ' + G.join('|'));
    }
    // line bookkeeping
    for (let r = 0; r < p.rowCount; r++) {
        const sr = p.getSourceRange(r);
        const hh = !!pr.header;
        const el = hh && r === 0 ? pr.headerLine : pr.rowLine[r - (hh ? 1 : 0)];
        const es = hh && r === 0 ? pr.headerSpan : pr.rowSpan[r - (hh ? 1 : 0)];
        if (sr.start !== el || sr.end !== el + es) {
            ok = false;
            msg.push('range r' + r + ' prov ' + JSON.stringify(sr) + ' doc ' + el + '+' + es);
        }
    }
    return { ok, msg: name + ' text=' + JSON.stringify(w.doc.text) + '\n' + msg.join('\n') };
}
/** One scripted action against the simulated page. */
export type SyncStep = (w: WebviewHandle, sh: FakeShell) => void;
export type Step = SyncStep | ((w: WebviewHandle, sh: FakeShell) => Promise<void>);
export async function run(
    name: string,
    text: string,
    eol: number,
    steps: Step[],
    o?: Record<string, string | boolean>
): Promise<RunResult> {
    const w = makeWebview(text, eol, o);
    await drain(w);
    const sh = w.shell();
    try {
        for (const s of steps) {
            await s(w, sh);
            await drain(w);
        }
    } catch (e) {
        return { ok: false, msg: 'EXC ' + name + ' ' + (e instanceof Error ? e.stack : String(e)), w };
    }
    return { ...check(name, w), w };
}
export const edit =
    (r: number, c: number, t: string): SyncStep =>
    (_w, sh) =>
        sh.o.onEdit([{ r, c, text: t }], 'type');
export const editMany =
    (a: [number, number, string][]): SyncStep =>
    (_w, sh) =>
        sh.o.onEdit(
            a.map(([r, c, t]) => ({ r, c, text: t })),
            'paste'
        );

/** Opens the row-header context menu and returns the items the page offered. */
function openRowMenu(sh: FakeShell, clientX: number, clientY: number): readonly GridMenuItem[] {
    const g = sh.grid;
    const captured: { items?: readonly GridMenuItem[] } = {};
    g._showMenu = it => {
        captured.items = it;
    };
    const e: GridContextMenuEvent = { r: 0, c: 0, sr: 0, zone: 'rowhdr', clientX, clientY, handled: false };
    defined(g.contextmenu, 'contextmenu listener')(e);
    return defined(captured.items, 'context menu');
}
const menuAction = (items: readonly GridMenuItem[], label: RegExp): (() => void) =>
    defined(
        items.find(i => i && label.test(i[0])),
        'menu item ' + String(label)
    )[2];

export const del =
    (rows: number[]): SyncStep =>
    (_w, sh) => {
        // emulate context-menu delete
        sh.grid.getSelection = () => ({ ranges: rows.map(r => ({ r1: r, r2: r, c1: 0, c2: 0 })) });
        menuAction(openRowMenu(sh, 0, 0), /^Delete/)();
    };
export const add: SyncStep = (_w, sh) => {
    sh.grid.getSelection = () => ({ ranges: [{ r1: 1, r2: 1, c1: 0, c2: 0 }] });
    menuAction(openRowMenu(sh, 0, 0), /^Add/)();
};
