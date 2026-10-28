import { defineConfig } from 'vitest/config';

/**
 * XSheet のテスト設定。
 * core（tests/core/ と直下）は node 環境、render（tests/render/）は jsdom 環境で 1 回実行する。
 */
export default defineConfig({
  test: {
    watch: false,
    projects: [
      {
        extends: true,
        test: {
          name: 'core',
          include: ['tests/*.test.ts', 'tests/core/**/*.test.ts'],
          environment: 'node',
        },
      },
      {
        extends: true,
        test: {
          name: 'render',
          include: ['tests/render/**/*.test.ts'],
          environment: 'jsdom',
        },
      },
    ],
  },
});
