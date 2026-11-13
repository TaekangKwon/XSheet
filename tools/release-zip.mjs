// 単独版の zip を作る（M6-5・ADR-0003 §6・`npm run release:zip`）。GitHub Release に添付する物。先に `npm run build` で dist/ を作っておくこと。
// - 中身は xsheet-<版>/ の下に、単独版のファイル（tools/standalone/parts.mjs の一覧）・viewer.html・viewer2.html・NOTICE・LICENSE。
// - viewer.html は examples/standalone.html の写し。スクリプトの場所（../dist/xsheet.min.js）だけを同じフォルダ（xsheet.min.js）に直す。
// - viewer2.html は examples/standalone-viewer2.html（ADR-0007 の未決を試す確認用のビューア）の写し。xsheet.min.js と xsheet.css の場所を同じフォルダに直す。
// - 出力は release/xsheet-<版>-standalone.zip（Git の対象外）。zip は jszip（このライブラリの依存）で作る。外へは通信しない。
// タグ・Release を作るのは人（ユーザーの指示があるとき）。手順は docs/02_architecture.md §8。
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import process from 'node:process';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';
import { STANDALONE_FILES } from './standalone/parts.mjs';

const repo = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const { version } = JSON.parse(readFileSync(path.join(repo, 'package.json'), 'utf8'));
const folder = `xsheet-${version}`;

const missing = STANDALONE_FILES.filter((f) => !existsSync(path.join(repo, 'dist', f)));
if (missing.length > 0) {
  console.error(`dist に無い：${missing.join(', ')}。先に npm run build を実行する`);
  process.exit(1);
}

const SCRIPT = '<script src="../dist/xsheet.min.js"></script>';
const page = readFileSync(path.join(repo, 'examples/standalone.html'), 'utf8');
if (page.split(SCRIPT).length !== 2) {
  console.error(`examples/standalone.html に ${SCRIPT} が 1 つだけあること`);
  process.exit(1);
}

const SCRIPT2 = '<script src="../dist/xsheet.min.js"></script>';
const CSS2 = '<link rel="stylesheet" href="../dist/xsheet.css" nonce="xsheet-viewer" />';
const page2 = readFileSync(path.join(repo, 'examples/standalone-viewer2.html'), 'utf8');
if (page2.split(SCRIPT2).length !== 2 || page2.split(CSS2).length !== 2) {
  console.error(
    'examples/standalone-viewer2.html に xsheet.min.js の <script> と xsheet.css の <link> が 1 つずつあること',
  );
  process.exit(1);
}

const zip = new JSZip();
for (const f of STANDALONE_FILES)
  zip.file(`${folder}/${f}`, readFileSync(path.join(repo, 'dist', f)));
zip.file(`${folder}/viewer.html`, page.replace(SCRIPT, '<script src="xsheet.min.js"></script>'));
zip.file(
  `${folder}/viewer2.html`,
  page2
    .replace(SCRIPT2, '<script src="xsheet.min.js"></script>')
    .replace(CSS2, '<link rel="stylesheet" href="xsheet.css" nonce="xsheet-viewer" />'),
);
zip.file(`${folder}/NOTICE`, readFileSync(path.join(repo, 'NOTICE')));
zip.file(`${folder}/LICENSE`, readFileSync(path.join(repo, 'LICENSE')));

const out = path.join(repo, 'release', `${folder}-standalone.zip`);
mkdirSync(path.dirname(out), { recursive: true });
const bytes = await zip.generateAsync({
  type: 'uint8array',
  compression: 'DEFLATE',
  compressionOptions: { level: 9 },
});
writeFileSync(out, bytes);
console.log(
  `${path.relative(repo, out)}：${bytes.length.toLocaleString('en-US')} B（${Object.keys(zip.files).filter((n) => !n.endsWith('/')).length} ファイル）`,
);
