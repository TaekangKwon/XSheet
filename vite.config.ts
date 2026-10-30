import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { defineConfig, type Plugin } from 'vite';

/**
 * ライブラリのビルド（M4-1・`npm run build`）。ESM だけを `dist/` に出す。
 * - 入口ごとに別のファイル（package.json の exports と同じ形）：`index`（全部・M6-2）・`core/index`・`core/internal`・`geom/index`・`render/index`。
 *   render は core を入口のファイルから引き（部品を重ねて持たない）、形の表（geom）は動的 import のまま別のチャンクで読む（N6）。
 * - 依存（jszip・ssf）は束ねない（使う側の依存として入る）。
 * - 圧縮しない（使う側のビルドが圧縮する）。大きさは `npm run size` で測る（M4-2）。
 * - 型（d.ts）は `tsc -p tsconfig.build.json` が出す。`style.css` と `NOTICE` はこの設定が `dist/` に写す。
 * テストの設定は vitest.config.ts、例のページの開発サーバーは examples/vite.config.ts。
 */
const root = fileURLToPath(new URL('.', import.meta.url));
const src = (path: string): string => fileURLToPath(new URL(`./src/${path}`, import.meta.url));

/** ファイルをそのまま dist に写す（JS から import しない CSS・著作権表示） */
function copyFiles(files: Record<string, string>): Plugin {
  return {
    name: 'xsheet-copy-files',
    apply: 'build',
    generateBundle() {
      for (const [fileName, from] of Object.entries(files)) {
        this.emitFile({ type: 'asset', fileName, source: readFileSync(from, 'utf8') });
      }
    },
  };
}

export default defineConfig({
  root,
  // 例のページ（examples/）の index.html を拾わない
  publicDir: false,
  plugins: [copyFiles({ 'style.css': src('style.css'), NOTICE: `${root}NOTICE` })],
  build: {
    outDir: 'dist',
    emptyOutDir: true,
    target: 'es2022',
    minify: false,
    sourcemap: false,
    copyPublicDir: false,
    lib: {
      entry: {
        index: src('index.ts'),
        'core/index': src('core/index.ts'),
        'core/internal': src('core/internal.ts'),
        'geom/index': src('geom/index.ts'),
        'render/index': src('render/index.ts'),
      },
      formats: ['es'],
    },
    rolldownOptions: {
      external: ['jszip', 'ssf'],
      output: {
        entryFileNames: '[name].js',
        chunkFileNames: 'chunks/[name]-[hash].js',
      },
    },
  },
});
