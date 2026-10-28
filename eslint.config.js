// @ts-check
import js from '@eslint/js';
import tseslint from 'typescript-eslint';

/**
 * XSheet の ESLint 設定（flat config）。InterHub の設定に倣い、
 * 外部のコードを実行しないこと（01_requirements.md N1）をここでも確かめる。
 */
export default tseslint.config(
  {
    // 生成物・文書は対象外
    ignores: ['dist/**', 'node_modules/**', 'coverage/**', 'docs/**'],
  },
  js.configs.recommended,
  ...tseslint.configs.recommended,
  {
    // tools/ は Node で動かす補助スクリプト（変換器など）
    files: ['tools/**/*.mjs'],
    languageOptions: { globals: { console: 'readonly' } },
  },
  {
    languageOptions: {
      ecmaVersion: 2022,
      sourceType: 'module',
    },
    rules: {
      // N1：eval・new Function を使わない
      'no-eval': 'error',
      'no-implied-eval': 'error',
      'no-new-func': 'error',
      // 未使用引数は `_` 接頭辞で意図的な未使用を許す
      '@typescript-eslint/no-unused-vars': [
        'warn',
        { argsIgnorePattern: '^_', varsIgnorePattern: '^_' },
      ],
    },
  },
);
