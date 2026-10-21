'use strict';

const js = require('@eslint/js');
const globals = require('globals');

const rules = {
    eqeqeq: ['warn', 'smart'],
    'no-unused-vars': ['warn', { args: 'none', caughtErrors: 'none' }],
    'no-empty': ['error', { allowEmptyCatch: true }],
    'no-constant-condition': ['error', { checkLoops: false }],
    'no-control-regex': 'off',
    'no-prototype-builtins': 'off',
    'no-useless-escape': 'warn',
    'no-var': 'off',
    'prefer-const': 'off'
};

module.exports = [
    {
        ignores: [
            'dist/**',
            'node_modules/**',
            'out/**',
            'media/vendor/**',
            'media/pdfjs/**',
            'docs/**',
            '**/*.vsix',
            'src/**/*.ts'
        ]
    },
    js.configs.recommended,
    {
        // Hand-written core modules: plain CommonJS scripts that also load under node for tests.
        files: ['src/core/**/*.js'],
        languageOptions: {
            ecmaVersion: 2022,
            sourceType: 'commonjs',
            globals: {
                ...globals.node,
                module: 'writable',
                require: 'readonly'
            }
        },
        rules
    },
    {
        files: ['test/**/*.js', 'webpack.config.js', 'eslint.config.js', 'scripts/**/*.js'],
        languageOptions: {
            ecmaVersion: 2022,
            sourceType: 'commonjs',
            globals: { ...globals.node, TextDecoder: 'readonly', TextEncoder: 'readonly' }
        },
        rules
    }
];
