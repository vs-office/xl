'use strict';
// Minifies and concatenates the webview styles from src/webview into dist/webview/style.css.
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const SRC = path.join(ROOT, 'src', 'webview');
const STYLES = ['core/grid.css', 'styles/vscode.css'];

function minifyCss(css) {
    return css
        .replace(/\/\*[\s\S]*?\*\//g, '')
        .replace(/\s+/g, ' ')
        .replace(/\s*([{};:,>])\s*/g, '$1')
        .replace(/;}/g, '}')
        .trim();
}

function build(outDir = path.join(ROOT, 'dist', 'webview')) {
    fs.mkdirSync(outDir, { recursive: true });
    const css = STYLES.map(f => minifyCss(fs.readFileSync(path.join(SRC, f), 'utf8'))).join('\n');
    fs.writeFileSync(path.join(outDir, 'style.css'), css);
    return css.length;
}

module.exports = { build };

if (require.main === module) {
    console.log('style.css:', build());
}
