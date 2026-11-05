// GitHub Pages のデモを作る（`npm run pages:build`・ADR-0005）。出力は pages/（Git の対象外）。
// - 例のページ（examples/index.html）を Vite でビルドし、pages/examples/ に置く。base は相対（./）にして、
//   サイトの場所（https://<owner>.github.io/XSheet/ など）に依らずに動くようにする。
// - 例のページのサンプル（examples/main.ts の SAMPLES。fixtures/excel の自作のファイル）を pages/fixtures/excel/ に写す。
//   例のページは ../fixtures/excel/<名前> を同じ出所から fetch する（開発サーバーと同じ相対の場所）。
// - 根の pages/index.html は、見本のブックを開いた例のページへ移るだけのページ（スクリプトなし）。
// ページの CSP は examples/index.html のまま（外の出所へは通信しない。選んだファイルはブラウザの中で読むだけ）。
import { copyFileSync, existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import process from 'node:process';
import { fileURLToPath } from 'node:url';
import { build } from 'vite';

const repo = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');
const out = path.join(repo, 'pages');

/** examples/main.ts の SAMPLES の file（`{ file: 'x.xlsx', ... }`）を拾う */
const main = readFileSync(path.join(repo, 'examples/main.ts'), 'utf8');
const samples = [...main.matchAll(/\{ file: '([^']+\.xlsx)'/g)].map((m) => m[1]);
if (samples.length === 0) {
  console.error('examples/main.ts の SAMPLES からサンプルを拾えない');
  process.exit(1);
}
const missing = samples.filter((f) => !existsSync(path.join(repo, 'fixtures/excel', f)));
if (missing.length > 0) {
  console.error(`fixtures/excel に無い：${missing.join(', ')}`);
  process.exit(1);
}

await build({
  configFile: path.join(repo, 'examples/vite.config.ts'),
  base: './',
  logLevel: 'warn',
  build: {
    outDir: out,
    emptyOutDir: true,
    rollupOptions: { input: path.join(repo, 'examples/index.html') },
  },
});

const fixtures = path.join(out, 'fixtures/excel');
mkdirSync(fixtures, { recursive: true });
for (const f of samples) copyFileSync(path.join(repo, 'fixtures/excel', f), path.join(fixtures, f));

writeFileSync(
  path.join(out, 'index.html'),
  `<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <meta http-equiv="Content-Security-Policy" content="default-src 'none'; base-uri 'none'; form-action 'none'" />
    <meta http-equiv="refresh" content="0; url=examples/?sample=showcase.xlsx" />
    <title>XSheet demo</title>
  </head>
  <body>
    <p><a href="examples/?sample=showcase.xlsx">Open the XSheet demo</a></p>
  </body>
</html>
`,
);

console.log(`Pages：OK（pages/・サンプル ${samples.length} 個）`);
