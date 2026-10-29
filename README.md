# XSheet

.xlsx を Excel どおりに表示する JavaScript ライブラリ（読み取り専用）。npm のパッケージ名は `xsheet`。

- セルの値・表示形式・列幅と行高・結合・罫線・塗りと文字の装飾を、Excel 16 の表示に合わせて描く（寸法と図の位置は 1px 以内が目標）。
- 画像・図形（プリセットの形 176 種・自由な形）・線・グループ・図形の文字を、セルの上の層に描く。
- 100 万セルのシートでも見えている範囲だけを描く（スクロールで描き足す）。
- 外へ通信しない・`eval` を使わない・ファイルの中身は信頼しない入力として `textContent` でだけ DOM に入れる・読み取りの量に上限がある。

> 非公開の版（v0.1.0）。設計と決定は [docs/](docs/README.md)、変更の記録は [CHANGELOG.md](CHANGELOG.md)。

## 入れ方

非公開の間は GitHub のタグから git 依存で入れる（タグに `dist/` が入っている。ビルドのスクリプトは走らない。[ADR-0002](docs/adr/0002-build-and-distribution.md)）。

```bash
npm install github:TaekangKwon/XSheet#v0.1.0
```

依存は `jszip` と `ssf` だけ（一緒に入る）。

## 使い方

### ブック全体をタブ付きで表示する

```ts
import { renderWorkbook } from 'xsheet/render';
import 'xsheet/style.css';

const container = document.getElementById('viewer')!; // 高さが決まっている要素
const data = await file.arrayBuffer();                // .xlsx のバイト列（ArrayBuffer か Uint8Array）

const view = await renderWorkbook(container, data);
// …
view.destroy(); // リスナ・blob URL を放し、ブックを閉じ、表示を外す
```

- `container` は**高さが決まっていること**（`height: 100%` の祖先・flex の子など）。中のスクロール枠 `.xs-body` が、見えている範囲を決める窓になる。
- 開けないファイル（暗号化・壊れている・大きすぎる・対応しない形式）でも例外にせず、表示の中に文言を出す。
- `view.showSheet(index)` でシートを替えられる（タブを押したのと同じ。`index` は `SheetInfo.index`・1 始まり）。

### 1 シートだけ表示する

```ts
import { openXlsx } from 'xsheet';
import { renderSheet } from 'xsheet/render';
import 'xsheet/style.css';

const book = await openXlsx(data);
const first = book.sheets.find((s) => !s.hidden) ?? book.sheets[0];
const view = await renderSheet(container, book, first.index);
// …
view.destroy(); // ブックは閉じない
book.close();
```

### 読み取りだけ（DOM なし）

`xsheet`（core）は DOM も Node 固有の API も使わない。ブラウザ・Web Worker・Node 20 以上で動く。

```ts
import { openXlsx, XlsxError } from 'xsheet';

try {
  const book = await openXlsx(data, { limits: { maxSheetXmlBytes: 32 * 1024 * 1024 } });
  const sheet = await book.openSheet(book.sheets[0].index);
  const rows = sheet.getRows(1, 100); // 1〜100 行目（表示どおりの文字・書式・罫線の番号）
  book.close();
} catch (e) {
  if (e instanceof XlsxError) console.log(e.kind); // 'not-zip' | 'encrypted' | 'broken' | 'too-large' | 'unsupported'
}
```

API の全体は [docs/02_architecture.md](docs/02_architecture.md) §2（core）・§3（render）。

### 指定（`renderWorkbook`・`renderSheet`）

| 指定 | 既定 | 内容 |
| --- | --- | --- |
| `messages` | 日本語（`MESSAGES_JA`） | 文言。一部のキーだけ渡してもよい。英語は `MESSAGES_EN` を渡す |
| `drawings` | `true` | 図の層（画像・図形・線）を出すか。`false` なら図を読まない |
| `nonce` | なし | 書式の `<style>` 要素に付ける nonce（CSP の `style-src 'nonce-…'` 用） |
| `limits` | `DEFAULT_LIMITS` | 読み取りの上限（`renderWorkbook` だけ。`renderSheet` では `openXlsx` に渡す） |

## サブパス

| import | 中身 |
| --- | --- |
| `xsheet` | 読み取り（`openXlsx`・`XlsxError`・`DEFAULT_LIMITS`・寸法の換算など）と型 |
| `xsheet/render` | 表示（`renderWorkbook`・`renderSheet`・`MESSAGES_JA`・`MESSAGES_EN`） |
| `xsheet/style.css` | 表示の CSS（変数の既定値は明るい表示） |
| `xsheet/geom` | 図形の形の表と式の評価。render が図形を描くときに動的 import するので、ふつうは直接使わない |
| `xsheet/internal` | render が使う半公開の部品。版の互換を約束しない |

- ESM だけ。型は同梱（TypeScript は `moduleResolution: "bundler"` が前提）。
- 形の表（gzip 約 30 KB）は図形のあるシートを開いたときだけ読まれる（使う側のビルドでも別のチャンクになる）。大きさは [02](docs/02_architecture.md) §8。

## 見た目（CSS 変数）

`xsheet/style.css` は明るい表示の既定値を、詳細度 0（`:where(:root, :host)`）で持つ。使う側はどこで上書きしてもよい（Shadow DOM の中でも効く）。

```css
/* ダーク表示の例（値は docs/02_architecture.md §6 の「ダークの例」） */
:root[data-theme='dark'] {
  --xs-fg: #e6e6e6;
  --xs-bg: #1e1e1e;
  --xs-cell-bg: #0f1115;
  --xs-head-bg: #1f232c;
  --xs-grid: #232733;
  --xs-border: #3a3a3a;
  --xs-muted: #9ca3af;
  --xs-faint: #6b7386;
  --xs-accent: #60a5fa;
  --xs-warn: #fbbf24;
  --xs-warn-bg: rgba(251, 191, 36, 0.14);
  --xs-border-ink: #d4d4d4;
  --xs-backdrop: #ffffff;
  --xs-backdrop-display: block;
}
```

変数は 14 個（文字・面・セルの地・見出し・グリッド線・枠・薄い文字 2 つ・強調・通知 2 つ・墨色の罫線・図の下地 2 つ）。役割は [02](docs/02_architecture.md) §6。Excel の色（塗り・文字の色・図形の色）は変数ではなく、ファイルの色のまま描く。

## CSP

| 指定 | 理由 |
| --- | --- |
| `script-src 'self'` | ライブラリと形の表のチャンクを同じ出所から読む（`'unsafe-eval'` は要らない） |
| `style-src 'nonce-<値>'`（または `'unsafe-inline'`） | 書式の `<style>` 要素。`nonce` 指定で付く。`style.css` を `<link>` で読むなら、その `<link>` にも nonce を付けるか出所を足す |
| `img-src blob:` | 図の画像は blob URL で出す |

詳しくは [02](docs/02_architecture.md) §3「CSP」。`style-src` を nonce だけにした例が `examples/index.html`。

## 上限

悪意のあるファイルで止まらないよう、zip の展開量・部品の大きさ・件数・図形の式・画像の大きさに上限（25 項目）がある。超えたら `XlsxError('too-large')`（`renderWorkbook` は文言を出す）か、その図だけを簡略化して通知を出す。既定値は `DEFAULT_LIMITS`、一覧は [02](docs/02_architecture.md) §5。`limits` で一部だけ変えられる。

例外のメッセージにはセル値・シート名を入れない。

## 対応環境

- `xsheet`（core）：ES2022 のブラウザ・Web Worker・Node 20 以上。
- `xsheet/render`：Chrome・Edge・Firefox・Safari の最新 2 版。

## 開発

開発には Node 22.12 以上が要る。

| コマンド | 用途 |
| --- | --- |
| `npm run test` | テスト（vitest） |
| `npm run typecheck`・`npm run lint`・`npm run format:check` | 型・ESLint・Prettier の確認 |
| `npm run build` | `dist/` を作る（Vite のライブラリモード＋d.ts） |
| `npm run check:consumer` | 別のプロジェクトから `dist/` を import して型検査とビルドが通るか・配る中身 |
| `npm run size` | 入口ごとの大きさと上限 |
| `npm run examples` | 例のページ（http://localhost:5173/examples/）。ファイルの選択・明暗・文言の切り替え |

## ライセンス

非公開（`UNLICENSED`）。図形の形の表は ECMA-376 の `presetShapeDefinitions.xml` から作っており、Ecma International の著作権表示を [NOTICE](NOTICE) に置く（配布物に同梱する）。依存のライセンスは [02](docs/02_architecture.md) §8。
