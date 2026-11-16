import js from '@eslint/js';
import globals from 'globals';
import tseslint from 'typescript-eslint';
import { importX } from 'eslint-plugin-import-x';

// Legacy JS (webview + tests) keeps the old rule set until each file becomes TS; new ESLint 10 recommended rules are off here only.
const jsRules = {
    eqeqeq: ['error', 'smart'],
    'no-useless-assignment': 'off',
    'preserve-caught-error': 'off',
    'no-unused-vars': ['error', { args: 'none', caughtErrors: 'none' }],
    'no-empty': ['error', { allowEmptyCatch: true }],
    'no-constant-condition': ['error', { checkLoops: false }],
    'no-control-regex': 'off',
    'no-prototype-builtins': 'off',
    'no-useless-escape': 'error',
    'import-x/no-unresolved': 'off'
};

export default tseslint.config(
    {
        ignores: [
            'dist/**',
            'out/**',
            'node_modules/**',
            'docs/**',
            '.vscode-test/**',
            '**/*.d.ts',
            'src/webview/vendor/**'
        ]
    },
    js.configs.recommended,
    importX.flatConfigs.recommended,

    // Host extension and (soon) everything else: strict, type-aware, no `any` in any form.
    {
        files: ['src/**/*.ts'],
        extends: [...tseslint.configs.recommendedTypeChecked, importX.flatConfigs.typescript],
        languageOptions: { parserOptions: { projectService: true, tsconfigRootDir: import.meta.dirname } },
        rules: {
            '@typescript-eslint/no-explicit-any': 'error',
            '@typescript-eslint/no-unsafe-assignment': 'error',
            '@typescript-eslint/no-unsafe-member-access': 'error',
            '@typescript-eslint/no-unsafe-call': 'error',
            '@typescript-eslint/no-unsafe-return': 'error',
            '@typescript-eslint/no-unsafe-argument': 'error',
            '@typescript-eslint/no-non-null-assertion': 'error',
            '@typescript-eslint/ban-ts-comment': [
                'error',
                {
                    'ts-expect-error': 'allow-with-description',
                    'ts-ignore': true,
                    'ts-nocheck': true,
                    minimumDescriptionLength: 10
                }
            ],
            '@typescript-eslint/consistent-type-assertions': [
                'error',
                { assertionStyle: 'as', objectLiteralTypeAssertions: 'never' }
            ],
            '@typescript-eslint/consistent-type-imports': 'error',
            '@typescript-eslint/no-unused-vars': ['error', { argsIgnorePattern: '^_', caughtErrors: 'none' }],
            '@typescript-eslint/no-floating-promises': 'error',
            'no-restricted-syntax': [
                'error',
                {
                    selector: 'TSAsExpression > TSAsExpression[typeAnnotation.type="TSUnknownKeyword"]',
                    message: 'Double cast through unknown is banned: fix the types or narrow with a type guard.'
                }
            ],
            eqeqeq: ['error', 'always'],
            'no-console': 'error',
            'import-x/no-unresolved': 'off'
        }
    },

    // Webview and core modules: plain browser scripts that also load under node for tests (until they become TS modules).
    {
        files: ['src/webview/**/*.js'],
        languageOptions: {
            ecmaVersion: 2022,
            sourceType: 'script',
            globals: {
                ...globals.browser,
                module: 'writable',
                require: 'readonly',
                OX: 'writable',
                Buffer: 'readonly',
                fflate: 'readonly',
                SSF: 'readonly',
                acquireVsCodeApi: 'readonly',
                // defined by the inline script that the extension host injects into the page
                getOptions: 'readonly',
                ignoreState: 'readonly'
            }
        },
        rules: { ...jsRules, 'no-console': 'off' }
    },
    {
        // Top-level functions of the page scripts are called from the injected inline script.
        files: ['src/webview/csv.js', 'src/webview/excel.js'],
        rules: { 'no-unused-vars': ['error', { args: 'none', caughtErrors: 'none', vars: 'local' }] }
    },
    {
        files: ['src/test/**/*.js', 'scripts/**/*.js', 'webpack.config.js'],
        languageOptions: {
            ecmaVersion: 2022,
            sourceType: 'commonjs',
            globals: { ...globals.node, SSF: 'readonly', fflate: 'readonly' }
        },
        // Tests are rewritten in TS in phase 4; until then they keep the pre-migration leniency.
        rules: { ...jsRules, 'no-console': 'off', 'no-unused-vars': 'off', 'no-useless-escape': 'off', eqeqeq: 'off' }
    }
);
