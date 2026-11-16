import test from 'node:test';
import assert from 'node:assert';
import { MergeIndex } from '../webview/core/grid-core.ts';
import type { ViewMap } from '../webview/core/grid-core.ts';
import { Grid } from '../webview/core/grid.ts';
import type { GridEdit, GridProvider } from '../webview/core/grid.ts';

const OX = { Grid };

test('Grid module loads headless and exposes API', () => {
    assert.equal(typeof OX.Grid, 'function');
    for (const m of [
        'setProvider',
        'refresh',
        'invalidate',
        'select',
        'getSelection',
        'scrollTo',
        'getScroll',
        'setScroll',
        'setZoom',
        'getZoom',
        'startEdit',
        'find',
        'findNext',
        'findPrev',
        'clearFind',
        'setView',
        'themeChanged',
        'focus',
        'destroy',
        'on',
        'off',
        'getColumnValues'
    ])
        assert.equal(typeof (OX.Grid.prototype as unknown as Record<string, unknown>)[m], 'function', m);
});

test('color parsing', () => {
    assert.deepEqual(OX.Grid._parseColor('#fff'), [255, 255, 255, 1]);
    assert.deepEqual(OX.Grid._parseColor('rgb(1, 2, 3)'), [1, 2, 3, 1]);
    assert.equal(OX.Grid._parseColor('nope'), null);
});

test('literalText providers: fill and internal paste do not shift =A2', () => {
    const mk = (literal: boolean) => {
        const raws = [['=A2'], ['']];
        const edits: GridEdit[] = [];
        // headless Grid: skip the DOM-building constructor and set only the fields the tested methods read
        const g = Object.create(OX.Grid.prototype) as Grid;
        const provider: Partial<GridProvider> = { literalText: literal, getRaw: (r, c) => (raws[r] || [])[c] };
        g.provider = provider as GridProvider; // only literalText/getRaw are read
        const view: Partial<ViewMap> = { toSource: (v: number) => v };
        g.view = view as ViewMap; // identity view: only toSource is read
        g.readOnly = false;
        g.active = { r: 2, c: 0 };
        g.mi = new MergeIndex();
        g.ranges = [{ r1: 0, c1: 0, r2: 0, c2: 0 }];
        g._clip = { raws: [['=A2']], text: '=A2', r: 0, c: 0 };
        g.emit = ((n: string, e?: { edits: GridEdit[] }) => {
            if (n === 'edit' && e) edits.push(...e.edits);
        }) as typeof g.emit; // collects 'edit' payloads only
        g._growPad = () => {};
        g._clampSel = () => {};
        g._selChanged = () => {};
        return { g, edits };
    };
    const cases: [boolean, string][] = [
        [true, '=A2'],
        [false, '=A4']
    ];
    for (const [literal, want] of cases) {
        const f = mk(literal);
        assert.equal(
            f.g._fillEdits({ r1: 0, c1: 0, r2: 0, c2: 0 }, 'down', 2, true)[1].text,
            want === '=A2' ? '=A2' : '=A4'
        );
        const p = mk(literal);
        p.g.ranges = [{ r1: 2, c1: 0, r2: 2, c2: 0 }];
        p.g.pasteText('=A2');
        assert.equal(p.edits[0].text, literal ? '=A2' : '=A4');
    }
});
