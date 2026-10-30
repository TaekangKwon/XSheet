// 単独版のビルド（M6-3・ADR-0003 §1・`npm run build` の最後の段）。先に vite build（バンドラ向け）と tsc（d.ts）が dist/ を作っておくこと。
// 依存（jszip・ssf）を中に持ち、圧縮して dist/ の根に出す。バンドラ向けの出力は変えない。
// - xsheet.min.js：IIFE・大域 `XSheet`。遅れて読む部品も中に入れる（`<script>` 1 つで済むことがこの形の価値）。
// - xsheet.mjs：ESM。遅れて読む部品は同じフォルダの xsheet-<名前>.mjs に分け、使うときだけ import.meta.url からの相対で読む。
//   部品は部品ごとに別にビルドする（1 回のビルドで分けると、本体と部品が共有する小さな部品が 3 つ目のファイルになるため）。
// - xsheet.css：style.css と同じもの（view が自動で入れるので任意）。
// - xsheet.licenses.txt：中に持つ依存（jszip・ssf ほか）のライセンスの全文。
// - xsheet.d.ts：大域 `XSheet` と `declare module 'xsheet'` の型（dist の d.ts を 1 つにまとめる）。xsheet.d.mts は xsheet.mjs の型。
// ビルドの結果で、外への import（裸の名前）が無いこと・eval／new Function が無いこと（N1）を確かめ、違えば失敗する。
// jsdom で動くこと・型が通ることは `npm run check:standalone`、大きさは `npm run size`。外へは通信しない。
import { copyFileSync, existsSync, readdirSync, readFileSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import process from 'node:process';
import { fileURLToPath } from 'node:url';
import { build } from 'vite';
import { LAZY_PARTS, STANDALONE_FILES } from './parts.mjs';

const repo = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');
const dist = path.join(repo, 'dist');
const at = (p) => path.join(repo, p).split(path.sep).join('/');

/** 3 つのビルドに共通の設定（依存を中に持つ・圧縮する） */
const base = {
  configFile: false,
  root: repo,
  publicDir: false,
  logLevel: 'warn',
  resolve: {
    alias: [
      // jszip の browser 向けの dist（jszip.min.js）は setimmediate の `new Function` を中に持つ。ソース（lib）から束ね、
      // setimmediate は `new Function` の無い版に、Node のストリーム（stream）は空に替える（N1）
      { find: /^jszip$/, replacement: at('node_modules/jszip/lib/index.js') },
      { find: /^setimmediate$/, replacement: at('tools/standalone/set-immediate.mjs') },
      { find: /^stream$/, replacement: at('tools/standalone/empty.mjs') },
    ],
  },
};

/** 動的 import の行き先が遅れて読む部品なら、その部品 */
function lazyPartOf(resolvedId) {
  const id = (resolvedId ?? '').split(path.sep).join('/');
  return LAZY_PARTS.find((p) => id === at(p.module));
}

/** ESM の本体：遅れて読む部品の動的 import を、同じフォルダの xsheet-<名前>.mjs への外の import にする */
function lazyExternal() {
  return {
    name: 'xsheet-lazy-external',
    // Vite の解決（resolve）より先に見る
    enforce: 'pre',
    async resolveId(source, importer, options) {
      if (importer === undefined || options.isEntry) return null;
      const resolved = await this.resolve(source, importer, { ...options, skipSelf: true });
      const part = lazyPartOf(resolved?.id);
      return part ? { id: `./xsheet-${part.name}.mjs`, external: true } : null;
    },
  };
}

/** 1 つのビルドを走らせ、出力のチャンクを返す */
async function run(config) {
  const result = await build({ ...base, ...config });
  return (Array.isArray(result) ? result : [result]).flatMap((r) =>
    'output' in r ? r.output.filter((o) => o.type === 'chunk') : [],
  );
}

/** ライブラリモードの共通の build 設定 */
const lib = (entry, format, fileName, extra = {}) => ({
  build: {
    outDir: dist,
    emptyOutDir: false,
    target: 'es2022',
    minify: true,
    sourcemap: false,
    copyPublicDir: false,
    lib: { entry: at(entry), name: 'XSheet', formats: [format], fileName: () => fileName },
    // Vite のライブラリモードは es の出力の空白を圧縮しないため、出力の設定でも圧縮する
    rolldownOptions: { output: { codeSplitting: false, minify: true } },
    ...extra,
  },
});

if (!existsSync(path.join(dist, 'index.d.ts'))) {
  console.error('dist/index.d.ts が無い。npm run build（vite build と tsc の後）から実行する');
  process.exit(1);
}

const chunks = [
  ...(await run(lib('src/index.ts', 'iife', 'xsheet.min.js'))),
  ...(await run({ ...lib('src/index.ts', 'es', 'xsheet.mjs'), plugins: [lazyExternal()] })),
];
for (const part of LAZY_PARTS) {
  chunks.push(...(await run(lib(part.module, 'es', `xsheet-${part.name}.mjs`))));
}
copyFileSync(path.join(repo, 'src/style.css'), path.join(dist, 'xsheet.css'));

// 検査 1：出力は決まったファイルだけで、外への import は遅れて読む部品だけ（IIFE は import を持たない）
const allowed = new Set(LAZY_PARTS.map((p) => `./xsheet-${p.name}.mjs`));
const errors = [];
const names = chunks.map((c) => c.fileName).sort();
const expectedJs = STANDALONE_FILES.filter((f) => /\.m?js$/.test(f)).sort();
if (JSON.stringify(names) !== JSON.stringify(expectedJs)) {
  errors.push(`出力のファイルが違う：${names.join(', ')}（期待：${expectedJs.join(', ')}）`);
}
for (const c of chunks) {
  const imports = [...c.imports, ...c.dynamicImports].filter((i) => i !== c.fileName);
  const bad = c.fileName === 'xsheet.mjs' ? imports.filter((i) => !allowed.has(i)) : imports;
  if (bad.length > 0) errors.push(`${c.fileName} が外を import する：${bad.join(', ')}`);
}
// 検査 2：文字列からコードを作る書き方が無い（N1。束ねた依存も含めて文面で見る）。
// eval という名前（`(0,eval)`・`self.eval`・`["eval"]` を含む）・Function の呼び出し（`globalThis.Function(`・`.apply` など）・
// 文字列を渡す setTimeout／setInterval。実行時の止め（--disallow-code-generation-from-strings）は check.mjs が見る
const CODEGEN =
  /\beval\b|(?<![\w$])Function\s*(?:\(|\.(?:apply|call|bind)\b)|new\s+Function\b|\bset(?:Timeout|Interval)\s*\(\s*["'`]/;
for (const f of expectedJs) {
  const m = CODEGEN.exec(readFileSync(path.join(dist, f), 'utf8'));
  if (m) errors.push(`${f} に ${m[0]} がある`);
}
if (errors.length > 0) {
  console.error(errors.join('\n'));
  process.exit(1);
}

// 束ねた依存のライセンス：単独版は jszip・ssf などの中身を配るため、各パッケージのライセンスの全文を 1 つのファイルにまとめる
// （MIT・Apache-2.0 などは、配る物に著作権表示と許諾の文を含めることを求める）。束ねたかどうかはビルドの moduleIds で決める
const bundled = new Set();
for (const c of chunks) {
  for (const id of c.moduleIds) {
    const m = /node_modules\/((?:@[^/]+\/)?[^/]+)\//.exec(id.split(path.sep).join('/'));
    if (m) bundled.add(m[1]);
  }
}
const licenseText = [...bundled].sort().map((name) => {
  const dir = path.join(repo, 'node_modules', name);
  const { version, license } = JSON.parse(readFileSync(path.join(dir, 'package.json'), 'utf8'));
  const file = readdirSync(dir).find((f) => /^licen[cs]e/i.test(f));
  if (file === undefined) throw new Error(`${name} のライセンスの文が見つからない`);
  const text = readFileSync(path.join(dir, file), 'utf8').replace(/\r\n/g, '\n').trim();
  return `${name} ${version}（${license}）\n${'-'.repeat(60)}\n${text}\n`;
});
writeFileSync(
  path.join(dist, 'xsheet.licenses.txt'),
  `XSheet の単独版（xsheet.min.js・xsheet.mjs・xsheet-*.mjs）が中に持つ依存のライセンス（tools/standalone/build.mjs が作る）。\n` +
    `XSheet 自身の著作権表示は NOTICE。jszip は MIT と GPL-3.0-or-later の選択で、XSheet は MIT を選ぶ。\n\n${licenseText.join('\n')}`,
);

// 型：dist の d.ts（tsc の出力）を、index から import を辿って 1 つのファイルの ambient module にまとめる。
// 入口は `declare module 'xsheet'`、ほかは `declare module 'xsheet/_/<パス>'`。大域の `XSheet` はその入口の型
/** d.ts の相対の指定を dist の中のファイル（拡張子なし）にする */
function resolveDts(fromRel, spec) {
  const target = path.posix.normalize(path.posix.join(path.posix.dirname(fromRel), spec));
  for (const cand of [target, `${target}/index`]) {
    if (existsSync(path.join(dist, `${cand}.d.ts`))) return cand;
  }
  throw new Error(`d.ts の import を解けない：${fromRel} → ${spec}`);
}
const moduleName = (rel) => (rel === 'index' ? 'xsheet' : `xsheet/_/${rel}`);
const modules = [];
const seen = new Set();
const visit = (rel) => {
  if (seen.has(rel)) return;
  seen.add(rel);
  const text = readFileSync(path.join(dist, `${rel}.d.ts`), 'utf8').replace(/\r\n/g, '\n');
  const deps = [];
  const body = text
    .replace(/(\bfrom\s*|\bimport\s*\(\s*)(['"])(\.{1,2}\/[^'"]+)\2/g, (_, head, _q, spec) => {
      const to = resolveDts(rel, spec);
      deps.push(to);
      return `${head}'${moduleName(to)}'`;
    })
    // ambient の中では declare を重ねて書けない
    .replace(/^(export )?declare /gm, '$1')
    .replace(/^export \{\};\n/gm, '')
    .trim();
  modules.push(`declare module '${moduleName(rel)}' {\n${body.replace(/^(?=.)/gm, '  ')}\n}`);
  deps.forEach(visit);
};
visit('index');
const pkg = JSON.parse(readFileSync(path.join(repo, 'package.json'), 'utf8'));
writeFileSync(
  path.join(dist, 'xsheet.d.ts'),
  [
    `// XSheet ${pkg.version} 単独版（xsheet.min.js・xsheet.mjs）の型。tools/standalone/build.mjs が dist の d.ts から作る（手で直さない）。`,
    `// <script> で読むときは大域の XSheet、型は import type { Viewer } from 'xsheet' で使える。`,
    ...modules,
    `/** xsheet.min.js が作る大域の名前 */\ndeclare const XSheet: typeof import('xsheet');`,
    '',
  ].join('\n\n'),
);
writeFileSync(
  path.join(dist, 'xsheet.d.mts'),
  `// xsheet.mjs の型（中身は xsheet.d.ts の 'xsheet'）\n/// <reference path="./xsheet.d.ts" />\nexport * from 'xsheet';\n`,
);

console.log(
  `単独版：OK（${STANDALONE_FILES.join('・')}。外への import・eval／new Function なし。型は ${modules.length} モジュール）`,
);
