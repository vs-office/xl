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
                exclude: [/node_modules/, path.resolve(__dirname, 'src', 'webview')],
                use: [
                    {
                        loader: 'ts-loader',
                        // tsconfig is a type-check-only config (noEmit, .ts import specifiers); the host bundle emits CommonJS
                        options: {
                            onlyCompileBundledFiles: true,
                            compilerOptions: { noEmit: false, allowImportingTsExtensions: false }
                        }
                    }
                ]
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
// excel-edit.js is the lazily loaded chunk behind import('./edit.ts') in excel.ts.
const webviewConfig = {
    name: 'webview',
    mode: 'none',
    target: 'web',
    entry: {
        excel: './src/webview/excel.ts',
        csv: './src/webview/csv.ts'
    },
    output: {
        path: path.resolve(__dirname, 'dist', 'webview'),
        filename: '[name].js',
        chunkFilename: '[name].js',
        // excel.ts sets __webpack_public_path__ from the host-provided script uri before loading the editing chunk
        publicPath: '',
        library: { name: 'xlPage', type: 'var' }
    },
    resolve: { extensions: ['.ts', '.js'] },
    module: {
        rules: [
            {
                // types are checked by tsc; webpack only strips them and keeps ES modules for code splitting
                test: /\.ts$/,
                include: path.resolve(__dirname, 'src', 'webview'),
                use: [
                    {
                        loader: 'ts-loader',
                        options: {
                            transpileOnly: true,
                            compilerOptions: { module: 'esnext', noEmit: false, allowImportingTsExtensions: false }
                        }
                    }
                ]
            }
        ]
    },
    devtool: false,
    performance: { hints: false }
};

module.exports = [hostConfig, webviewConfig];
