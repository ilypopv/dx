'use strict';

const path = require('path');
const webpack = require('webpack');

const tsRule = {
    test: /\.ts$/,
    exclude: /node_modules/,
    use: [{ loader: 'ts-loader' }]
};

module.exports = (env, argv) => {
    const production = argv.mode === 'production';
    const base = {
        mode: 'none',
        devtool: production ? false : 'source-map',
        performance: { hints: false },
        module: { rules: [tsRule] }
    };

    // One bundle for desktop and web: the webworker build only uses workspace.fs and has no Node requires.
    const hostConfig = {
        ...base,
        target: 'webworker',
        entry: './src/extension.ts',
        output: {
            path: path.resolve(__dirname, 'dist'),
            filename: 'extension.js',
            libraryTarget: 'commonjs2',
            clean: true,
            devtoolModuleFilenameTemplate: '../[resource-path]'
        },
        externals: {
            vscode: 'commonjs vscode'
        },
        resolve: {
            mainFields: ['browser', 'module', 'main'],
            extensions: ['.ts', '.js'],
            fallback: {
                assert: require.resolve('assert'),
                path: require.resolve('path-browserify')
            }
        },
        plugins: [new webpack.ProvidePlugin({ process: require.resolve('process/browser.js') })]
    };

    // The webview bundles are classic scripts. pdf.js and its worker are baked into out/pdf.js:
    // the worker source is inlined as a string and becomes a same-origin blob worker at runtime.
    // out/word.js carries the .docx and .doc readers and fflate.
    const webviewConfig = {
        ...base,
        target: 'web',
        entry: {
            pdf: './src/webview/pdf.ts',
            word: './src/webview/word.ts'
        },
        output: {
            path: path.resolve(__dirname, 'out'),
            filename: '[name].js',
            iife: true,
            clean: true,
            devtoolModuleFilenameTemplate: '../[resource-path]'
        },
        resolve: {
            mainFields: ['browser', 'module', 'main'],
            extensions: ['.ts', '.js', '.mjs'],
            // The UMD build of fflate probes for Node workers inside a try block.
            fallback: { worker_threads: false },
            alias: {
                'pdfjs-dist$': path.resolve(__dirname, 'node_modules/pdfjs-dist/build/pdf.min.mjs')
            }
        },
        module: {
            // pdf.js keeps a Node.js code path behind dynamic imports that never run in a webview.
            exprContextCritical: false,
            rules: [
                tsRule,
                {
                    test: /pdf\.worker\.min\.mjs$/,
                    type: 'asset/source'
                }
            ]
        }
    };

    return [hostConfig, webviewConfig];
};
