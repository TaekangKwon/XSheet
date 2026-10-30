import { defineConfig } from 'vitest/config';

/**
 * XSheet のテスト設定。
 * core（tests/core/・tests/geom/ と直下）と render（tests/render/）は、どちらも node 環境で 1 回実行する。
 * render で DOM が要るファイルは、先頭の
 * `// @vitest-environment jsdom` で jsdom にする（InterHub と同じ。jsdom では fixtures を `new URL(…, import.meta.url)` で読めないため）。
 * 重さの計測（`*.perf.test.ts`）は、ほかのテストと同時に走ると時間がぶれて落ちるため、別の project にして
 * ほかが終わってから（`groupOrder`）1 ファイルずつ（`fileParallelism: false`）走らせる（M2-4）。
 */
const PERF = 'tests/**/*.perf.test.ts';

export default defineConfig({
  test: {
    watch: false,
    // 入口の層（facade/style.ts）が `style.css?inline` で CSS の文面を読む。既定では vitest が CSS を空にするため、style.css だけは中身を返す（M6-2）
    css: { include: [/src[\\/]style\.css/] },
    projects: [
      {
        extends: true,
        test: {
          name: 'core',
          include: ['tests/*.test.ts', 'tests/core/**/*.test.ts', 'tests/geom/**/*.test.ts'],
          exclude: [PERF],
          environment: 'node',
        },
      },
      {
        extends: true,
        test: {
          name: 'render',
          include: ['tests/render/**/*.test.ts', 'tests/facade/**/*.test.ts'],
          exclude: [PERF],
          environment: 'node',
          // 形の表（geom）は render から動的 import するため、部品のテストの前に読み込んでおく（M2-3）
          setupFiles: ['tests/render/setup.ts'],
        },
      },
      {
        extends: true,
        test: {
          name: 'perf',
          include: [PERF],
          environment: 'node',
          setupFiles: ['tests/render/setup.ts'],
          fileParallelism: false,
          sequence: { groupOrder: 1 },
        },
      },
    ],
  },
});
