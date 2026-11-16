import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import vm from 'node:vm';
import webpack from 'webpack';
import type { Configuration } from 'webpack';

// The build scripts are CommonJS files; describe just the parts used here.
interface NamedConfig extends Configuration {
    name?: string;
}
interface StylesModule {
    build(dir: string): number;
}
const require = createRequire(import.meta.url);
const configs = require('../../webpack.config.js') as NamedConfig[];
const styles = require('../../scripts/build-styles.js') as StylesModule;

function run(config: Configuration): Promise<void> {
    return new Promise((resolve, reject) => {
        webpack(config, (err, stats) => {
            if (err) return reject(err);
            if (stats?.hasErrors()) return reject(new Error(stats.toString({ all: false, errors: true })));
            resolve();
        });
    });
}

test('webview bundles evaluate and expose the page API; the editing chunk stays out of the first-paint bundle', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xl-bundle-'));
    try {
        const webview = configs.find(c => c.name === 'webview');
        assert.ok(webview);
        await run({ ...webview, output: { ...webview.output, path: dir } });
        assert.ok(styles.build(dir) > 100);

        const read = (f: string): string => fs.readFileSync(path.join(dir, f), 'utf8');
        for (const f of ['excel.js', 'csv.js', 'excel-edit.js']) assert.ok(read(f).length > 1000, f);

        for (const page of ['excel.js', 'csv.js']) {
            const ctx = vm.createContext({ console }) as Record<string, unknown>;
            ctx.self = ctx;
            vm.runInContext(read(page), ctx);
            const api = vm.runInContext(
                '({ start: typeof xlPage.start, init: typeof xlPage.initPage })',
                ctx
            ) as unknown;
            assert.deepEqual(JSON.parse(JSON.stringify(api)), { start: 'function', init: 'function' }, page);
        }

        // editing pieces are a separate chunk: formula engine and writer are not in the pages, the chunk registers itself
        assert.ok(!/function patchXlsxSheets/.test(read('excel.js')));
        assert.ok(!/function patchXlsxSheets/.test(read('csv.js')));
        assert.ok(/function patchXlsxSheets/.test(read('excel-edit.js')));
        const ctx = vm.createContext({ console }) as Record<string, unknown>;
        ctx.self = ctx;
        vm.runInContext(read('excel.js'), ctx);
        vm.runInContext(read('excel-edit.js'), ctx);
        const chunks = vm.runInContext('Object.keys(self).filter(k => k.startsWith("webpackChunk"))', ctx) as string[];
        assert.equal(chunks.length, 1);
    } finally {
        fs.rmSync(dir, { recursive: true, force: true });
    }
});
