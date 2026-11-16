'use strict';
// Loads a webview module by path relative to src/webview (e.g. 'core/format'); Node strips the types natively.
const path = require('node:path');
const root = path.resolve(__dirname, '../../webview');

module.exports = rel => require(path.join(root, rel + '.ts'));
