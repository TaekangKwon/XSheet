# Web Worker で読む

| 項目 | 値 |
| --- | --- |
| 文書 | XSheet 使い方の手引き |
| 更新日 | 2026-10-30 |
| 関連 | [README.md](README.md)・[../02_architecture.md](../02_architecture.md) §2（core） |

## TL;DR
- 読み取り（`xsheet/core` の `openXlsx`）は DOM を使わないので、Worker の中でそのまま動く。
- 表示（`view`・`renderWorkbook`）は DOM が要るので main thread に置く。開いたブック（`XlsxBook`）は関数を持つので、スレッドの間で渡せない。
- Worker からは行の中身（`getRows` の結果。ただのデータ）を返す。`ArrayBuffer` は転送（transfer）で写さずに渡す。

---

## 1. 向いていること

| 場面 | Worker に置く |
| --- | --- |
| 全シートの文字を集める（検索・索引・変換） | 向く。main thread を止めずに済む |
| 大きなファイルの中身を確かめる（シートの一覧・行数） | 向く |
| 画面に出す | 置けない。`view` を main thread で呼ぶ（中で読み直す） |

表示と取り出しの両方をするなら、同じバイト列を 2 か所で読む。転送するとこちらの `ArrayBuffer` は空になるので、表示にも使うなら先に `buf.slice(0)` で写しておく。

## 2. バンドラ（Vite・webpack 5）での例

```ts
// xlsx.worker.ts
import { openXlsx, XlsxError } from 'xsheet/core';

self.onmessage = async (e: MessageEvent<ArrayBuffer>) => {
  try {
    const book = await openXlsx(e.data);
    const sheets = [];
    for (const info of book.sheets) {
      const sheet = await book.openSheet(info.index);
      const last = Math.min(sheet.rowCount, 1000);
      // cells：undefined＝空・null＝結合で隠れる
      const rows = sheet.getRows(1, last).map((row) => row.cells.map((c) => c?.text ?? ''));
      sheets.push({ name: info.name, rows });
    }
    book.close();
    self.postMessage({ ok: true, sheets });
  } catch (err) {
    // メッセージにセル値・シート名は入らない。種類で分ける
    self.postMessage({ ok: false, kind: err instanceof XlsxError ? err.kind : 'unknown' });
  }
};
```

```ts
// main thread
const worker = new Worker(new URL('./xlsx.worker.ts', import.meta.url), { type: 'module' });
worker.onmessage = (e) => console.log(e.data);

const buf = await file.arrayBuffer();
worker.postMessage(buf, [buf]); // 転送：写さずに渡す（buf はこちらでは使えなくなる）
```

- `xsheet/core` を使う。`xsheet`（全部）でも動くが、DOM を使わないことを約束しているのは `xsheet/core` だけ。
- `openXlsx` が受けるのは `ArrayBuffer` か `Uint8Array`。`File`・`Blob` のまま渡したいなら、Worker の中で `await blob.arrayBuffer()` してから渡すか、`xsheet` の `read` を使う。
- `getRows` の結果（`SheetRow`）・`SheetInfo`・`SheetLayout` は関数を持たないデータなので `postMessage` で返せる。`XlsxBook`・`XlsxSheet` は返せない。
- 上限は main thread と同じ（`openXlsx(data, { limits })`）。超えたら `XlsxError('too-large')`。

## 3. 単独版での例

`<script>` の単独版（`xsheet.min.js`）は、読み込んだ時点で DOM に触れない。classic の Worker で `importScripts` してよい。

```js
// worker.js（classic）
importScripts('xsheet.min.js');

self.onmessage = async (e) => {
  const book = await XSheet.openXlsx(e.data);
  self.postMessage(book.sheets.map((s) => s.name));
  book.close();
};
```

module の Worker なら `xsheet.mjs` を import する（`new Worker('worker.mjs', { type: 'module' })`）。

```js
// worker.mjs
import { read } from './xsheet.mjs';

self.onmessage = async (e) => {
  const book = await read(e.data);
  self.postMessage(book.sheets.length);
  book.close();
};
```

## 4. Node

Node（20 以上）でも同じ `xsheet/core` を使う（`worker_threads` の中でも同じ）。`fs` の `Buffer` は `Uint8Array` なので、そのまま渡せる。

```ts
import { readFile } from 'node:fs/promises';
import { openXlsx } from 'xsheet/core';

const book = await openXlsx(await readFile('book.xlsx'));
```
