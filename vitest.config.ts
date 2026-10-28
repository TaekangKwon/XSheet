import { defineConfig } from 'vitest/config';

/**
 * XSheet のテスト設定。
 * core（tests/core/・tests/geom/ と直下）と render（tests/render/）は、どちらも node 環境で 1 回実行する。
 * render で DOM が要るファイルは、先頭の
 * `// @vitest-environment jsdom` で jsdom にする（InterHub と同じ。jsdom では fixtures を `new URL(…, import.meta.url)` で読めないため）。
 */
export default defineConfig({
  test: {
    watch: false,
    projects: [
      {
        extends: true,
        test: {
          name: 'core',
          include: ['tests/*.test.ts', 'tests/core/**/*.test.ts', 'tests/geom/**/*.test.ts'],
          environment: 'node',
        },
      },
      {
        extends: true,
        test: {
          name: 'render',
          include: ['tests/render/**/*.test.ts'],
          environment: 'node',
        },
      },
    ],
  },
});
