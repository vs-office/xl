'use strict';

const path = require('path');
const webpack = require('webpack');

const hostConfig = {
    name: 'host',
    mode: 'none',
    entry: './src/extension.ts',
    output: {
        path: path.resolve(__dirname, 'dist'),
        filename: 'extension.js',
        libraryTarget: 'commonjs2',
        devtoolModuleFilenameTemplate: '../[resource-path]'
    },
    devtool: 'source-map',
    externals: {
        vscode: 'commonjs vscode'
    },
    module: {
        rules: [
            {
                test: /\.ts$/,
                exclude: /node_modules/,
                use: [{ loader: 'ts-loader' }]
            }
        ]
    },
    // One bundle for desktop and web: the webworker build only uses workspace.fs and has no Node requires.
    target: 'webworker',
    resolve: {
        mainFields: ['browser', 'module', 'main'],
        extensions: ['.ts', '.js'],
        fallback: {
            assert: require.resolve('assert'),
            path: require.resolve('path-browserify')
        }
    },
    plugins: [new webpack.ProvidePlugin({ process: 'process/browser' })],
    infrastructureLogging: {
        level: 'log'
    }
};

// Pages that run inside the webview. The inline script of each page calls xlPage.start(options, ignoreState).
// excel-edit.js is the lazily loaded chunk behind import('./edit.js') in excel.js.
const webviewConfig = {
    name: 'webview',
    mode: 'none',
    target: 'web',
    entry: {
        excel: './src/webview/excel.js',
        csv: './src/webview/csv.js'
    },
    output: {
        path: path.resolve(__dirname, 'dist', 'webview'),
        filename: '[name].js',
        chunkFilename: '[name].js',
        // excel.js sets __webpack_public_path__ from the host-provided script uri before loading the editing chunk
        publicPath: '',
        library: { name: 'xlPage', type: 'var' }
    },
    devtool: false,
    performance: { hints: false }
};

module.exports = [hostConfig, webviewConfig];
