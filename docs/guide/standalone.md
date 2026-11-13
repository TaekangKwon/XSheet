# 単独版（`xsheet.min.js`・`xsheet.mjs`）

| 項目 | 値 |
| --- | --- |
| 文書 | XSheet 使い方の手引き |
| 更新日 | 2026-10-30 |
| 関連 | [README.md](README.md)・[../02_architecture.md](../02_architecture.md) §8・[ADR-0003](../adr/0003-universal-distribution.md) §1・§6 |

## TL;DR
- 単独版は依存（jszip・ssf）を中に持つ配布物。ビルドの道具も `node_modules` も要らない。
- `<script>` 1 つなら `xsheet.min.js`（大域の `XSheet`・`file://` でも動く）。バンドラ無しの `import` なら `xsheet.mjs`（http で配る）。
- 配り直すときは `LICENSE`・`NOTICE`・`xsheet.licenses.txt` を一緒に置く。

---

## 1. ファイル

`npm run build` の最後の段（`tools/standalone/build.mjs`）が `dist/` の根に出す。大きさは 2026-10-30 の実測。

| ファイル | 中身 | gzip |
| --- | --- | ---: |
| `xsheet.min.js` | IIFE。大域の `XSheet` を作る。依存と形の表・グラフの描画をすべて中に持つ・圧縮済み | 130,861 B |
| `xsheet.mjs` | ESM。依存を中に持つ・圧縮済み。形の表・グラフの描画は持たない | 85,931 B |
| `xsheet-geom.mjs` | 形の表（遅れて読む部品）。`xsheet.mjs` が図形を描くときだけ、同じフォルダから読む | 30,588 B |
| `xsheet-chart.mjs` | グラフの描画（遅れて読む部品）。`xsheet.mjs` がグラフを描くときだけ、同じフォルダから読む | 25,031 B |
| `xsheet.css` | 表示の CSS（`style.css` と同じ）。`view` が自動で入れるので、読まなくてよい | — |
| `xsheet.d.ts` | 型。大域の `XSheet` と `declare module 'xsheet'` | — |
| `xsheet.d.mts` | `xsheet.mjs` の型（中身は `xsheet.d.ts` を引くだけ） | — |
| `xsheet.licenses.txt` | 中に持つ依存のライセンスの全文 | — |

- 中身は `xsheet` の入口と同じ（`view`・`read`・`openXlsx`・`renderWorkbook`・`MESSAGES_*` など全部）。
- `eval`・`new Function` を使わない（ビルドが文面を検査する）。CSP は `script-src 'self'` で動く（[csp.md](csp.md)）。
- 遅れて読む部品の一覧は `tools/standalone/parts.mjs` の `LAZY_PARTS`。部品を足すと、`xsheet.mjs` の横の `xsheet-<名前>.mjs` が増える（`xsheet.min.js` は中に入れる）。

## 2. `<script>` で使う（`xsheet.min.js`）

```html
<script src="xsheet.min.js"></script>
<div id="viewer" style="height: 600px"></div>
<script>
  // 例：fetch で取った xlsx を出す（http のページで）
  XSheet.view('#viewer', fetch('book.xlsx'), { theme: 'auto' });
</script>
```

- ダブルクリック（`file://`）で開くページでも動く。ただし `file://` のページから `fetch` で別のファイルは読めない（ブラウザの決まり）。ファイルの選択（`<input type="file">`）かドロップで渡す。
- 出来上がりの例が [examples/standalone.html](../../examples/standalone.html)（ファイルを選ぶ・落とす・テーマと言語を替える）。このリポジトリでは先に `npm run build` を走らせる。
- 1 ファイルなので、形の表も最初に読む（図形の無いブックでも分割しない）。最初の読み込みを軽くしたいなら `xsheet.mjs` を使う。

## 3. バンドラ無しで `import` する（`xsheet.mjs`）

```html
<script type="module">
  import { view } from './xsheet.mjs';

  let viewer;
  document.querySelector('#file').addEventListener('change', async (e) => {
    viewer?.destroy();
    viewer = await view('#viewer', e.target.files[0]);
  });
</script>
```

- **http で配る**。ブラウザは `file://` のページでモジュールのスクリプトを読まない。このリポジトリでは `npm run build` の後に `npm run examples:static` を走らせ、http://localhost:5174/examples/esm.html を開く（[examples/esm.html](../../examples/esm.html)）。
- `xsheet-geom.mjs`・`xsheet-chart.mjs` を `xsheet.mjs` と**同じフォルダ**に置く。`xsheet.mjs` は図形・グラフのあるシートを開いたときだけ、自分の場所からの相対（`./xsheet-geom.mjs`・`./xsheet-chart.mjs`）で読む。置き忘れても表は出るが、図形は簡略化して通知を出し、グラフは種類名の破線枠のままになる。
- import map で名前を付けてもよい：

```html
<script type="importmap">
  { "imports": { "xsheet": "/vendor/xsheet/xsheet.mjs" } }
</script>
<script type="module">
  import { view } from 'xsheet';
</script>
```

- Deno・Node（20 以上）でも `import` できる。表示は DOM が要るので、そこでは `read` だけを使う。

## 4. 配る zip の中身

GitHub Release に添付する単独版の zip（[ADR-0003](../adr/0003-universal-distribution.md) §6・版を出す手順は [02](../02_architecture.md) §8）。

| ファイル | 内容 |
| --- | --- |
| `xsheet.min.js`・`xsheet.mjs`・`xsheet-geom.mjs`・`xsheet-chart.mjs`・`xsheet.css` | 上の 1 のとおり |
| `xsheet.d.ts`・`xsheet.d.mts` | 型 |
| `viewer.html` | 出来上がりのビューア（`examples/standalone.html` を `xsheet.min.js` と同じフォルダで動くようにしたもの）。ダブルクリックで開ける |
| `viewer2.html` | 確認用のビューア（`examples/standalone-viewer2.html`）。[ADR-0007](../adr/0007-moobius-viewer-api.md) の未決の 2 つ（表示に 0.5.0 の操作を足す・組み込みの日付の書式の既定を日本語版 Excel の表にする）を試す。検索・倍率・見出しの有無・クリックしたセルを試せる。ダブルクリックで開ける |
| `LICENSE` | XSheet のライセンス（MIT） |
| `NOTICE` | 形の表の元の Ecma の著作権表示 |
| `xsheet.licenses.txt` | 中に持つ依存のライセンスの全文 |

## 5. ライセンスの文

単独版は jszip・ssf などの中身をそのまま含む。これらのライセンス（MIT・Apache-2.0 など）は、配る物に著作権表示と許諾の文を含めることを求める。そのため単独版を自分のアプリと一緒に配るときは、`LICENSE`・`NOTICE`・`xsheet.licenses.txt` も一緒に置く。

`xsheet.licenses.txt` に入るもの（ビルドが束ねたパッケージから作る。2026-10-30 の時点）：

| パッケージ | 版 | ライセンス |
| --- | --- | --- |
| `immediate` | 3.0.6 | MIT |
| `jszip` | 3.10.2 | MIT OR GPL-3.0-or-later（MIT を選ぶ） |
| `lie` | 3.3.0 | MIT |
| `pako` | 1.0.11 | MIT AND Zlib |
| `ssf` | 0.11.2 | Apache-2.0 |

- バンドラ向けの出力（`xsheet` などの入口）は依存を束ねないので、このファイルは要らない（依存は使う側の `node_modules` に入る）。
- npm の `setimmediate` は `new Function` を含むため束ねず、同じ働きの小さな版（`tools/standalone/set-immediate.mjs`）に替えている。jszip の Node 用のストリームは空の部品に替えている。

## 6. TypeScript から使う

**`<script>` の大域の `XSheet`**：`xsheet.d.ts` をプロジェクトに置き、`tsconfig.json` の `include`（または `files`）に入れる。

```ts
// app.ts（xsheet.d.ts が型の対象に入っていること）
import type { Viewer } from 'xsheet'; // 型だけ。実行時には何も読まない

let viewer: Viewer | undefined;
async function show(file: File): Promise<void> {
  viewer?.destroy();
  viewer = await XSheet.view('#viewer', file, { lang: 'auto' });
}
```

**`xsheet.mjs` の import**：`xsheet.d.mts` と `xsheet.d.ts` を `xsheet.mjs` と同じフォルダに置く。`import { view } from './xsheet.mjs'` に型が付く。

- `xsheet.d.ts` は `declare module 'xsheet'` を持つ。npm（git 依存）の `xsheet` と同じプロジェクトで一緒に使わない（どちらか一方にする）。
- `npm run check:standalone` が、大域・`import type`・`./xsheet.mjs` の 3 通りの型検査を通す。

## 7. 自分でビルドする

```bash
npm ci
npm run build             # dist/ に単独版も出る
npm run check:standalone  # jsdom で XSheet.view が描く・素の Node で read が動く・型が通る
npm run size              # 単独版の大きさが上限の内
```
