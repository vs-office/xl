import js from '@eslint/js';
import globals from 'globals';
import tseslint from 'typescript-eslint';
import { importX } from 'eslint-plugin-import-x';

// Legacy JS (tests, scripts) keeps the old rule set until it becomes TS; new ESLint 10 recommended rules are off here only.
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

    // Host extension and webview: strict, type-aware, no `any` in any form.
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
            // control characters are matched on purpose (XML escaping, number format code checks), as in the JS config
            'no-control-regex': 'off',
            'no-console': 'error',
            'import-x/no-unresolved': 'off'
        }
    },

    {
        // the webview logs timings to the webview devtools console
        files: ['src/webview/**/*.ts'],
        rules: { 'no-console': 'off' }
    },

    {
        files: ['src/test/**/*.js', 'scripts/**/*.js', 'webpack.config.js'],
        languageOptions: {
            ecmaVersion: 2022,
            sourceType: 'commonjs',
            globals: { ...globals.node }
        },
        // Tests are rewritten in TS in phase 4; until then they keep the pre-migration leniency.
        rules: { ...jsRules, 'no-console': 'off', 'no-unused-vars': 'off', 'no-useless-escape': 'off', eqeqeq: 'off' }
    }
);
