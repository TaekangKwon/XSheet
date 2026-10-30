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
      // render は geom を型だけ静的に引く（M2-3）。形の表（約 171 KB）は `shapes.ts` の `loadGeom` が動的 import で読む。
      // 動的 import（`import('../geom/index')`）はこの規則の対象外
      '@typescript-eslint/no-restricted-imports': [
        'error',
        {
          patterns: [
            {
              group: ['../geom', '../geom/*'],
              allowTypeImports: true,
              message: 'render は geom を型だけ引く（実体は shapes.ts の loadGeom で動的 import）',
            },
          ],
        },
      ],
    },
  },
  {
    // 入口の層（M6-2・ADR-0003 §2）は core・render の公開の入口の上に作る（写したファイルの中の作りに直接つながない）。
    // 文言の重ね方（resolveMessages）だけは render の messages から引く
    files: ['src/facade/**'],
    rules: {
      'no-restricted-imports': [
        'error',
        {
          patterns: [
            {
              group: ['../core/*', '!../core/index'],
              message: '入口の層は core を ../core/index から引く',
            },
            {
              group: ['../render/*', '!../render/index', '!../render/messages'],
              message: '入口の層は render を ../render/index（と ../render/messages）から引く',
            },
            {
              group: ['../geom', '../geom/*'],
              message: '入口の層は geom を引かない（形の表は render が遅れて読む）',
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
