// @ts-check
import { builtinModules } from 'node:module';
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
    // tools/ と fixtures/ の .mjs は Node で動かす補助スクリプト（変換器・期待値の抽出など）
    files: ['tools/**/*.mjs', 'fixtures/**/*.mjs'],
    languageOptions: { globals: { console: 'readonly' } },
  },
  {
    // 図形の画像の照合は、Playwright で開いたページの中（page.evaluate）でも動く
    files: ['fixtures/excel/compare-ref-geoms.mjs'],
    languageOptions: { globals: { window: 'readonly', document: 'readonly' } },
  },
  {
    // core・geom は Node 固有の API を使わない。jszip の型が `@types/node` を引き込み、
    // tsconfig.core.json（types: []）だけでは process・Buffer を止められないため、ここで止める（DOM は tsconfig.core.json が止める）
    files: ['src/core/**', 'src/geom/**'],
    rules: {
      'no-restricted-globals': [
        'error',
        ...['process', 'Buffer', 'global', 'require', 'module', '__dirname', '__filename'].map(
          (name) => ({ name, message: 'core・geom は Node 固有の API を使わない' }),
        ),
      ],
      'no-restricted-imports': [
        'error',
        {
          paths: builtinModules.map((name) => ({
            name,
            message: 'core・geom は Node 固有の API を使わない',
          })),
          patterns: [{ group: ['node:*'], message: 'core・geom は Node 固有の API を使わない' }],
        },
      ],
    },
  },
  {
    // render は core を公開の入口（index）と半公開の入口（internal）からだけ引く（02_architecture.md §2・M1-3）。
    // core の中の作りに直接つながると、公開の範囲が見えなくなり、M4 で core と render を別のチャンクにしたときに部品が重なる
    files: ['src/render/**'],
    rules: {
      'no-restricted-imports': [
        'error',
        {
          patterns: [
            {
              group: ['../core/*', '!../core/index', '!../core/internal'],
              message: 'render は core を ../core/index か ../core/internal から引く',
            },
          ],
        },
      ],
    },
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
