'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const vm = require('node:vm');
const webpack = require('webpack');
const configs = require('../../webpack.config.js');
const { build: buildStyles } = require('../../scripts/build-styles.js');

function run(config) {
    return new Promise((resolve, reject) => {
        webpack(config, (err, stats) => {
            if (err) return reject(err);
            if (stats.hasErrors()) return reject(new Error(stats.toString({ all: false, errors: true })));
            resolve();
        });
    });
}

test('webview bundles evaluate and expose the page API; the editing chunk stays out of the first-paint bundle', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xl-bundle-'));
    try {
        const webview = configs.find(c => c.name === 'webview');
        await run({ ...webview, output: { ...webview.output, path: dir } });
        assert.ok(buildStyles(dir) > 100);

        const read = f => fs.readFileSync(path.join(dir, f), 'utf8');
        for (const f of ['excel.js', 'csv.js', 'excel-edit.js']) assert.ok(read(f).length > 1000, f);

        for (const page of ['excel.js', 'csv.js']) {
            const ctx = vm.createContext({ console });
            ctx.self = ctx;
            vm.runInContext(read(page), ctx);
            const api = vm.runInContext('({ start: typeof xlPage.start, init: typeof xlPage.initPage })', ctx);
            assert.deepEqual(JSON.parse(JSON.stringify(api)), { start: 'function', init: 'function' }, page);
        }

        // editing pieces are a separate chunk: formula engine and writer are not in the pages, the chunk registers itself
        assert.ok(!/function patchXlsxSheets/.test(read('excel.js')));
        assert.ok(!/function patchXlsxSheets/.test(read('csv.js')));
        assert.ok(/function patchXlsxSheets/.test(read('excel-edit.js')));
        const ctx = vm.createContext({ console });
        ctx.self = ctx;
        vm.runInContext(read('excel.js'), ctx);
        vm.runInContext(read('excel-edit.js'), ctx);
        const chunks = vm.runInContext('Object.keys(self).filter(k => k.startsWith("webpackChunk"))', ctx);
        assert.equal(chunks.length, 1);
    } finally {
        fs.rmSync(dir, { recursive: true, force: true });
    }
});
