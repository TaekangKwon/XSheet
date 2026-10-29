import { fileURLToPath } from 'node:url';
import { defineConfig } from 'vite';

/**
 * 例のページの開発サーバー（`npm run examples` → http://localhost:5173/examples/）。
 * 根をリポジトリの根にして、`src/`（ライブラリ）と `fixtures/excel/`（サンプル）をそのまま配る。
 * 外へは公開しない（localhost だけで待ち受ける）。
 */
export default defineConfig({
  root: fileURLToPath(new URL('..', import.meta.url)),
  server: {
    host: 'localhost',
    port: 5173,
    strictPort: true,
    open: false,
  },
});
