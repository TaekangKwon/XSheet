// 配る JS の大きさを測り、上限（N6）を超えたら失敗する（M4-2・`npm run size`）。
// 先に `npm run build` で dist/ を作っておくこと。
// - 入口ごとに、その入口から static に辿れるファイル（共有のチャンクを含む）の合計を測る。動的 import（形の表）は辿らない。
// - 測るのは 3 通り：dist のまま（圧縮なし）・それを gzip・Vite で圧縮（minify）してから gzip。使う側のビルドは圧縮するので、上限は最後の値で決める。
//   ライブラリモードの ES の出力は、圧縮しても空白とコメントの一部が残る（Vite の仕様）。使う側の圧縮より少し大きめに出る。
// - 依存（jszip・ssf）は含めない（使う側の依存）。
// - 単独版（M6-3・dist の根の xsheet*.js／mjs）は、依存を中に持ち圧縮済みのファイルをそのまま gzip して、別の上限で見る。
import { execFileSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync, statSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import process from 'node:process';
import { fileURLToPath } from 'node:url';
import { gzipSync } from 'node:zlib';
import { LAZY_PARTS } from './standalone/parts.mjs';

const repo = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

/**
 * 上限（B・圧縮して gzip）。M4-2 の実測（docs/02_architecture.md §8）に 1 割ほどの余裕を足した値。
 * 超えたら、大きくなった理由を確かめてから上限を見直す。
 */
const LIMITS = {
  'core/index': 31_000,
  'render/index': 52_000,
  'geom/index': 36_000,
};

/**
 * 単独版の上限（gzip・依存を含む）。M6-3 の実測（docs/02_architecture.md §8）に 1 割ほどの余裕を足した値。
 * xsheet.min.js は遅れて読む部品を中に持つ。xsheet.mjs は最初に読む分だけ（部品は別に測る）。
 */
const STANDALONE_LIMITS = {
  'xsheet.min.js': 122_000,
  'xsheet.mjs': 89_000,
  'xsheet-geom.mjs': 34_000,
};

/**
 * その入口から static import で辿れるファイル（入口を含む）。
 * `from "./x.js"` と副作用だけの `import "./x.js"` を拾い、動的 import（`import("./x.js")`）は拾わない。
 */
function closure(dir, entry) {
  const seen = new Set();
  const visit = (rel) => {
    if (seen.has(rel)) return;
    seen.add(rel);
    const text = readFileSync(path.join(dir, rel), 'utf8');
    for (const m of text.matchAll(/\b(?:from|import)\s*["'](\.{1,2}\/[^"']+)["']/g)) {
      visit(path.posix.normalize(path.posix.join(path.posix.dirname(rel), m[1])));
    }
  };
  visit(`${entry}.js`);
  return [...seen].sort();
}

function measure(dir, files) {
  const raw = files.reduce((n, f) => n + statSync(path.join(dir, f)).size, 0);
  // ファイルごとに gzip する（配信はファイル単位）
  const gzip = files.reduce(
    (n, f) => n + gzipSync(readFileSync(path.join(dir, f)), { level: 9 }).length,
    0,
  );
  return { raw, gzip };
}

const dist = path.join(repo, 'dist');
const min = mkdtempSync(path.join(tmpdir(), 'xsheet-size-'));
let failed = false;
try {
  execFileSync(
    process.execPath,
    [
      path.join(repo, 'node_modules/vite/bin/vite.js'),
      'build',
      '--minify',
      '--outDir',
      min,
      '--emptyOutDir',
      '--logLevel',
      'warn',
    ],
    { cwd: repo, stdio: 'inherit' },
  );
  // 拾い方が壊れていないこと：render は core の部品（共有のチャンク）を全部含み、形の表（geom の入口）は含まない
  for (const dir of [dist, min]) {
    const core = closure(dir, 'core/index').filter((f) => f.startsWith('chunks/'));
    const render = closure(dir, 'render/index');
    if (
      core.length === 0 ||
      core.some((f) => !render.includes(f)) ||
      render.includes('geom/index.js')
    ) {
      throw new Error(
        `import の辿り方が合わない（core：${core.join(', ')}／render：${render.join(', ')}）`,
      );
    }
  }
  const rows = [];
  for (const entry of Object.keys(LIMITS)) {
    const files = closure(dist, entry);
    const plain = measure(dist, files);
    const minified = measure(min, closure(min, entry));
    const limit = LIMITS[entry];
    const over = minified.gzip > limit;
    failed ||= over;
    rows.push({
      entry,
      files: files.length,
      raw: plain.raw,
      gzip: plain.gzip,
      minGzip: minified.gzip,
      limit,
      over,
    });
  }
  const fmt = (n) => n.toLocaleString('en-US');
  console.log('| 入口 | ファイル | dist | gzip | 圧縮＋gzip | 上限 |');
  console.log('| --- | ---: | ---: | ---: | ---: | ---: |');
  for (const r of rows) {
    console.log(
      `| ${r.entry} | ${r.files} | ${fmt(r.raw)} | ${fmt(r.gzip)} | ${fmt(r.minGzip)}${r.over ? ' ✗' : ''} | ${fmt(r.limit)} |`,
    );
  }
  // 単独版
  const standalone = Object.keys(STANDALONE_LIMITS);
  const missing = [
    `xsheet.min.js`,
    `xsheet.mjs`,
    ...LAZY_PARTS.map((p) => `xsheet-${p.name}.mjs`),
  ].filter((f) => !standalone.includes(f));
  if (missing.length > 0) throw new Error(`単独版の上限が無い：${missing.join(', ')}`);
  console.log('');
  console.log('| 単独版 | 圧縮済み | gzip | 上限 |');
  console.log('| --- | ---: | ---: | ---: |');
  for (const f of standalone) {
    const { raw, gzip } = measure(dist, [f]);
    const over = gzip > STANDALONE_LIMITS[f];
    failed ||= over;
    console.log(
      `| ${f} | ${fmt(raw)} | ${fmt(gzip)}${over ? ' ✗' : ''} | ${fmt(STANDALONE_LIMITS[f])} |`,
    );
  }
} finally {
  rmSync(min, { recursive: true, force: true });
}
if (failed) {
  console.error('上限を超えた入口がある（✗）');
  process.exit(1);
}
