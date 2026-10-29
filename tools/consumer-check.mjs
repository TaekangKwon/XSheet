// 配る形（dist/）を、別のプロジェクトから import して型検査とビルドが通るか確かめる（M4-1・`npm run check:consumer`）。
// 先に `npm run build` で dist/ を作っておくこと。
// - 一時フォルダに使う側のプロジェクトを作り、node_modules/xsheet をこのリポジトリへのリンク（Windows は junction）にする。
//   package.json の exports（サブパス・types）を通って解決されることを見る（src/ を直接は見ない）。
// - 型検査は TypeScript（moduleResolution: bundler・strict・vite/client の型）、ビルドは Vite。どちらもこのリポジトリのものを使う。
// - ビルドの結果で、形の表（geom）が別のチャンク（動的 import）のままであること・style.css が CSS に入ることを見る。
// - 先に、npm pack で配る中身（exports の指す先・NOTICE）を見る（M4-2）。
// 外へは通信しない（依存はこのリポジトリの node_modules から解決する）。
import { execFileSync } from 'node:child_process';
import {
  existsSync,
  mkdirSync,
  mkdtempSync,
  readdirSync,
  readFileSync,
  rmdirSync,
  rmSync,
  symlinkSync,
  unlinkSync,
  writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import process from 'node:process';
import { fileURLToPath } from 'node:url';

const repo = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const keep = process.argv.includes('--keep');

for (const f of [
  'dist/core/index.js',
  'dist/core/index.d.ts',
  'dist/render/index.js',
  'dist/geom/index.js',
  'dist/style.css',
  'dist/NOTICE',
]) {
  if (!existsSync(path.join(repo, f))) {
    console.error(`${f} が無い。先に npm run build を実行する`);
    process.exit(1);
  }
}

// 配る中身（M4-2）：git 依存でも npm は pack して入れるため、package.json の files で入る物を npm pack --dry-run で見る。
// exports の指す先・NOTICE（根と dist の両方）が入り、テスト用の src/testing・ソースが入らないこと
{
  // npm の本体（npm-cli.js）を node で直接動かす（Windows の npm.cmd はシェルを通さないと動かないため）
  const npmCli =
    process.env.npm_execpath ??
    path.join(path.dirname(process.execPath), 'node_modules/npm/bin/npm-cli.js');
  const pack = JSON.parse(
    execFileSync(process.execPath, [npmCli, 'pack', '--dry-run', '--json'], {
      cwd: repo,
      encoding: 'utf8',
    }),
  )[0];
  const packed = new Set(pack.files.map((f) => f.path));
  const exportsMap = JSON.parse(readFileSync(path.join(repo, 'package.json'), 'utf8')).exports;
  const targets = Object.values(exportsMap).flatMap((v) =>
    typeof v === 'string' ? [v] : Object.values(v),
  );
  const missing = [...targets.map((t) => t.replace(/^\.\//, '')), 'NOTICE', 'dist/NOTICE'].filter(
    (f) => !packed.has(f),
  );
  const extra = [...packed].filter((f) => f.startsWith('src/') || f.includes('testing'));
  if (missing.length > 0 || extra.length > 0) {
    console.error(`配る中身が違う（無い：${missing.join(', ')}／余分：${extra.join(', ')}）`);
    process.exit(1);
  }
  if (
    readFileSync(path.join(repo, 'dist/NOTICE'), 'utf8') !==
    readFileSync(path.join(repo, 'NOTICE'), 'utf8')
  ) {
    console.error('dist/NOTICE が NOTICE と違う');
    process.exit(1);
  }
  console.log(
    `配る中身：OK（${pack.files.length} ファイル・${pack.unpackedSize.toLocaleString('en-US')} B）`,
  );
}

const dir = mkdtempSync(path.join(tmpdir(), 'xsheet-consumer-'));
const write = (rel, text) => {
  mkdirSync(path.dirname(path.join(dir, rel)), { recursive: true });
  writeFileSync(path.join(dir, rel), text);
};

write(
  'package.json',
  JSON.stringify({ name: 'xsheet-consumer', private: true, type: 'module' }, null, 2),
);
write(
  'tsconfig.json',
  JSON.stringify(
    {
      compilerOptions: {
        strict: true,
        target: 'ES2022',
        module: 'ESNext',
        moduleResolution: 'bundler',
        lib: ['ES2022', 'DOM', 'DOM.Iterable'],
        // 使う側の多くは Vite（*.css の import の型）。node_modules の型の探しかたは既定のまま
        types: ['vite/client'],
        typeRoots: [path.join(repo, 'node_modules')],
        noEmit: true,
        isolatedModules: true,
        verbatimModuleSyntax: true,
      },
      include: ['src'],
    },
    null,
    2,
  ),
);
write(
  'index.html',
  '<!doctype html><html><head><meta charset="utf-8"><script type="module" src="/src/main.ts"></script></head><body><div id="app"></div></body></html>\n',
);
// ふつうの使い方（core・render・style.css）。ビルドの入口
write(
  'src/main.ts',
  `import { DEFAULT_LIMITS, XlsxError, colName, openXlsx, type XlsxBook, type XlsxLimits } from 'xsheet';
import { MESSAGES_EN, renderSheet, renderWorkbook, type WorkbookView, type XSheetMessages } from 'xsheet/render';
import 'xsheet/style.css';

const limits: XlsxLimits = { ...DEFAULT_LIMITS };
const messages: Partial<XSheetMessages> = { noSheets: MESSAGES_EN.noSheets };
export async function show(el: HTMLElement, data: ArrayBuffer): Promise<WorkbookView> {
  return renderWorkbook(el, data, { limits, messages, nonce: 'n' });
}
export async function first(el: HTMLElement, data: ArrayBuffer): Promise<void> {
  let book: XlsxBook | undefined;
  try {
    book = await openXlsx(data);
    (await renderSheet(el, book, book.sheets[0]!.index)).destroy();
  } catch (e) {
    if (e instanceof XlsxError) console.log(e.kind);
  } finally {
    book?.close();
  }
}
// 呼ばない関数はビルドで消えるため、ファイルの選択から呼ぶ
const app = document.getElementById('app')!;
document.querySelector('input')?.addEventListener('change', async (e) => {
  const file = (e.target as HTMLInputElement).files?.[0];
  if (file === undefined) return;
  const data = await file.arrayBuffer();
  await first(app, data);
  (await show(app, data)).destroy();
  console.log(colName(1));
});
`,
);
// 残りの入口（geom・internal）は型検査だけ。ビルドの入口から static に引くと、形の表が別のチャンクにならないため
write(
  'src/other-entries.ts',
  `import { applyFillMode } from 'xsheet/internal';
import { presetNames, type GeomResult } from 'xsheet/geom';

export const checks: [unknown, number, GeomResult | null] = [applyFillMode, presetNames().length, null];
`,
);
mkdirSync(path.join(dir, 'node_modules'), { recursive: true });
const link = path.join(dir, 'node_modules/xsheet');
symlinkSync(repo, link, 'junction');

const run = (bin, args) =>
  execFileSync(process.execPath, [path.join(repo, 'node_modules', bin), ...args], {
    cwd: dir,
    stdio: 'inherit',
  });

try {
  console.log(`使う側のプロジェクト：${dir}`);
  run('typescript/bin/tsc', ['-p', 'tsconfig.json']);
  console.log('型検査：OK');
  run('vite/bin/vite.js', ['build', '--logLevel', 'warn']);
  const assets = path.join(dir, 'dist/assets');
  const files = readdirSync(assets);
  const js = files.filter((f) => f.endsWith('.js'));
  const css = files.filter((f) => f.endsWith('.css'));
  // 形の表は動的 import のまま別のチャンク（入口の JS に入らない）
  const withGeom = js.filter((f) =>
    readFileSync(path.join(assets, f), 'utf8').includes('"accentBorderCallout1"'),
  );
  const entry = js.find((f) => f.startsWith('index-'));
  if (js.length < 2 || withGeom.length !== 1 || withGeom[0] === entry) {
    throw new Error(`形の表が別のチャンクになっていない（${js.join(', ')}）`);
  }
  if (css.length !== 1 || !readFileSync(path.join(assets, css[0]), 'utf8').includes('.xs-sheet')) {
    throw new Error('style.css がビルドの CSS に入っていない');
  }
  console.log(`ビルド：OK（JS ${js.length} 個・形の表は ${withGeom[0]}・CSS ${css[0]}）`);
} catch (e) {
  console.error(e instanceof Error ? e.message : e);
  process.exitCode = 1;
} finally {
  if (keep) {
    console.log(`残した：${dir}`);
  } else {
    // リンクを先に外す（リンクの先＝このリポジトリを消さないため。消すのは一時フォルダの中だけ）
    try {
      unlinkSync(link);
    } catch {
      rmdirSync(link);
    }
    rmSync(dir, { recursive: true, force: true });
  }
}
