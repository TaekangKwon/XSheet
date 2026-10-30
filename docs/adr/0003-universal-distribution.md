# 0003. 誰でも使える配り方と、InterHub の変更を取り込みやすい作り（M6）

> 版: 1.0 ・ 更新日: 2026-10-30 ・ 担当: kwon
> 関連文書: [0002](0002-build-and-distribution.md)（ビルドの形・D5） / [../02_architecture.md](../02_architecture.md) §2・§3・§8 / [../03_source-map.md](../03_source-map.md) / [../04_roadmap.md](../04_roadmap.md)（M6）

## TL;DR
- 今の XSheet は「Vite などのバンドラ＋ESM＋`moduleResolution: "bundler"`」の使う側でしか動かない。依存（jszip・ssf）を束ねず、どちらも CommonJS のため、`<script>` やバンドラ無しの ESM では読めない。使い方も「サブパス 2 つ・CSS の import・高さの決まった container・ダークは変数 14 個の上書き」と覚えることが多い。
- SheetJS（`xlsx.full.min.js`・`xlsx.mjs`）や PDF.js と同じく、**依存を束ねた単独の配布物**を足す：`xsheet.min.js`（`<script>` 1 つで大域の `XSheet`）と `xsheet.mjs`（バンドラ無しの `import`）。今のバンドラ向けの出力はそのまま残す。
- 使い方は **1 行**にする：`XSheet.view('#viewer', file)`。CSS は自動で入れ、高さが無い container でも見える既定の高さを付け、言語とテーマは `lang`・`theme` の語で選べるようにする（薄い入口の層 `src/facade/` を core・render の上に足すだけで、写したファイルは触らない）。
- InterHub は今もグラフ（ISSUE-101〜111）を足し続けている。取り込みを手作業の写しから**道具で差分を出す作業**に変える：対応表（`tools/sync/interhub.json`）と `npm run sync:status`・`sync:patch`（InterHub は読むだけ）。写したファイルの形を InterHub と揃えたまま保ち、XSheet 固有のものは別のファイルに置く。
- まず社内（リポジトリの権限がある人）向けに作る。社外への公開（リポジトリの公開・ライセンス D3・npm）は別に決める。

---

## ステータス

accepted（2026-10-30。「決めたこと」Q1〜Q4 はユーザーが提案のとおりに決めた。M6 の Issue は M6-1 から始める）

## コンテキスト

### 1. 今の使いにくさ（v0.1.0）

| # | 使いにくさ | 原因 |
| --- | --- | --- |
| P1 | `<script src>` で読めない。バンドラ無しの `<script type="module">`・CDN・import map でも読めない | ESM だけ（ADR-0002）。`dist` は `"jszip"`・`"ssf"` を裸の名前で import し、両方とも CommonJS（`main` が `./lib/index`・`./ssf`）なのでブラウザが直接読めない |
| P2 | 入れ方が限られる | 非公開の GitHub の git 依存だけ。リポジトリの権限が要る。zip を落として置くだけの使い方が無い |
| P3 | どこから何を import するかを覚える必要がある | `xsheet`（読み取り）と `xsheet/render`（表示）に分かれている。1 シートだけ出すには `openXlsx` → `renderSheet` → `destroy` → `close` の 4 手 |
| P4 | CSS を別に読まないと崩れる | `xsheet/style.css` の import か `<link>` が要る |
| P5 | 何も出ないことがある | container の高さが 0 だとスクロール枠が 0px になる（README に注意書きがあるだけ） |
| P6 | ダーク表示に手間がかかる | 変数 14 個を使う側が上書きする（値は README の例を写す） |
| P7 | 文言の切り替えが重い | `messages: MESSAGES_EN` のように物を import して渡す。既定は日本語だけ |
| P8 | 型の解決が限られる | d.ts の import に拡張子が無く、`node16`・`nodenext` では型を辿れない |

設計した本人も手順を見ないと使えない、というのが出発点（2026-10-30 ユーザー）。

### 2. 他のライブラリの配り方（2026-10-30 に公式の文書で確認）

| | SheetJS 0.20.3 | PDF.js |
| --- | --- | --- |
| `<script>` 1 つ | `dist/xlsx.full.min.js`（大域 `XLSX`）・軽い版 `xlsx.mini.min.js`・古いブラウザ向け `shim.min.js` | `pdf.min.mjs`＋`pdf.worker.min.mjs`（ESM）・古いブラウザ向け `legacy/` |
| バンドラ無しの ESM | `xlsx.mjs`（依存を中に持つ） | `pdf.mjs` |
| 出来上がりの画面 | なし | `web/viewer.html`（iframe で置ける） |
| 配り先 | 自前の CDN（`cdn.sheetjs.com`）が正。npm のレジストリの `xlsx` は 0.18.5 で止まっている。npm は CDN の tarball の URL で入れる | npm（`pdfjs-dist`）・CDN |

共通点は「依存を束ねた単独のファイル」と「最小の例が数行」。

### 3. InterHub の動き（読むだけで確認）

- 943e4f0 の後に 22 コミット。.xlsx のグラフ（ADR-0031・ISSUE-101〜111）が入った：core 側に `lib/xlsx/chart.ts`（831 行）・`types.ts`（+224）・`limits.ts`（+24）・`drawing.ts`・`format.ts`・`index.ts`、render 側に `lib/viewers/xlsx-chart-*.ts`（axis・layout・svg・thin・store など合わせて約 3,200 行）と `xlsx-image-layer.ts` の変更。照合の fixtures は JSON だけで約 26 万行。
- 2026-10-30 の時点でもグラフの作業は続いている（未コミットの変更あり）。
- xrag-Viewer にも同じ 943e4f0 から写した `packages/xlsx-view` がある（handoff の記録）。同じコードの写しが 3 つある。

## 決定

### 1. 配布物を 2 系統にする

| 系統 | ファイル | 誰のため | 中身 |
| --- | --- | --- | --- |
| A. 単独（新規） | `dist/xsheet.min.js` | `<script>` 1 つで使いたい人。ビルドの道具を持たない人・`file://` でダブルクリックで開くページ | IIFE・大域 `XSheet`・依存と形の表をすべて中に持つ・圧縮する |
| | `dist/xsheet.mjs`＋`dist/xsheet-geom.mjs` | バンドラ無しで `import` したい人（CDN・import map・Deno など） | ESM・依存を中に持つ・圧縮する。形の表は同じフォルダの `xsheet-geom.mjs` を `import.meta.url` からの相対で、図形を描くときだけ読む |
| | `dist/xsheet.css` | CSS を `<link>` で読みたい人・CSP で `<style>` を自動で足したくない人 | `style.css` と同じもの（下の 2 で自動で入るので任意） |
| | `dist/xsheet.d.ts` | 単独版を TypeScript から使う人 | 大域 `XSheet` の型 |
| B. バンドラ向け（今のまま） | `dist/core`・`render`・`geom`・`chunks`・`style.css` | Vite・webpack などで使う人 | ESM・依存は束ねない・圧縮しない（ADR-0002 のとおり） |

- 単独版の大きさは約 110 KB（gzip・実測は M6-3）。参考：使う側の Vite のビルドで最初に読む分が 78 KB、形の表が 30 KB（02 §8）。
- 今後入る重い部品（グラフの描画など）は、形の表と同じく**遅れて読む部品**として足す：B と `xsheet.mjs` では別のファイル、`xsheet.min.js` では中に入れる。ビルドの設定は遅れて読む部品の一覧（配列）から作り、部品を足すときは一覧に 1 行足すだけにする。
- `xsheet.min.js` が分割しない理由：`<script>` 1 つで済むことがこの系統の価値。ファイルを足すと、置き場所の相対パスで迷う。
- 単独版でも N1（`eval`・`new Function` を使わない）と CSP（`script-src 'self'`）は守る。束ねた jszip に `eval`・`new Function` が入っていないことをビルドの検査で見る（M6-3）。

### 2. 1 行で使える入口の層（`src/facade/`）

core と render の上に薄い層を足す。写したファイル（`src/core/`・`src/render/`）は変えない（§4 の規約）。

```html
<!-- 単独版 -->
<script src="xsheet.min.js"></script>
<div id="viewer"></div>
<script>
  input.onchange = () => XSheet.view('#viewer', input.files[0]);
</script>
```

```ts
// バンドラ
import { view } from 'xsheet';
const viewer = await view(container, file, { theme: 'auto' });
viewer.destroy();
```

| API | 内容 |
| --- | --- |
| `view(target, source, options?)` | ブック全体をタブ付きで出す（`renderWorkbook` の上に作る）。戻り値は `Viewer` |
| `read(source, options?)` | DOM なしで開く（`openXlsx` に同じ。受け取れる形を `view` と揃える） |
| `Viewer` | `sheets`（`SheetInfo[]`）・`showSheet(番号か名前)`・`destroy()` |

`target`・`source`：

- `target`：要素か CSS セレクタの文字列。
- `source`：`File`・`Blob`・`ArrayBuffer`・`Uint8Array` などの型付き配列・`Response`、またはそれらの Promise。
- URL の文字列は受けない。ライブラリは通信しない（CLAUDE.md の守ること）。使う側が `XSheet.view('#v', fetch(url))` と書けば、ライブラリは受け取った `Response` の中身を読むだけになる。

`options`（どれも省略できる）：

| 指定 | 既定 | 内容 |
| --- | --- | --- |
| `lang` | `'auto'` | `'ja'`・`'en'`・`'ko'`・`'auto'`（`<html lang>`、無ければ `navigator.language`。合うものが無ければ英語）。細かく変えたいときは今までどおり `messages` |
| `theme` | `'light'` | `'light'`・`'dark'`・`'auto'`（`prefers-color-scheme` に従う）。ダークの値は style.css に持つ（今は README の例） |
| `css` | `'auto'` | 表示の CSS を自動で入れる。`<link>`・import で読み込み済み（`--xs-fg` が定義済み）なら入れない。`'none'` なら入れない |
| `height` | `'auto'` | container の高さが 0 のときだけ、表示の根に既定の高さ（480px）を付けて `console.warn` を 1 回出す。数値・CSS の長さも渡せる |
| `drawings`・`nonce`・`limits`・`messages` | 今と同じ | `renderWorkbook` にそのまま渡す |

- `renderWorkbook`・`renderSheet`・`openXlsx` は残す。細かく作りたい人の道具で、`view` の下にあるもの。
- `MESSAGES_KO`（韓国語・23 キー）を足す。
- CSS の自動挿入は書式の `<style>` と同じく `nonce` を付ける。CSP の要件は今と変わらない（`style-src` の nonce か `'unsafe-inline'`）。

### 3. 入口（`package.json` の `exports`）

| import | 今 | 提案 |
| --- | --- | --- |
| `xsheet` | 読み取り（core） | **全部**：`view`・`read`＋core＋render の公開物（今の公開物はすべて残るので壊れない） |
| `xsheet/core` | なし | 読み取りだけ（DOM・Node 固有の API を使わないことを約束する入口。Worker・Node 向け） |
| `xsheet/render` | 表示 | 今のまま |
| `xsheet/style.css`・`xsheet/geom`・`xsheet/internal` | | 今のまま |

- 使う側が覚える名前を `xsheet` 1 つにする。バンドラは使わない部分を落とす（`sideEffects` は CSS だけ）。
- 条件：render のファイルは読み込んだ時点（モジュールの最上位）で DOM に触れない。今も触れていない（2026-10-30 に確認）。Node で `xsheet` を import するテストで守る。

### 4. InterHub の変更を取り込みやすくする

**4-1. 写しの規約**（写したファイルを InterHub の形のまま保つ）

1. InterHub から写したファイルは、ファイル名・関数の並び・中身を InterHub と揃えたまま保つ。整理や改名のためだけには変えない（差分が当たらなくなる）。
2. XSheet が決まった変換で直すのは次だけ：接頭辞（`ih-` → `xs-`）・文言（`t()` → `messages`）・import のパス・モジュール変数の状態をインスタンスへ（03 の 6 点）。これらは道具で機械的にかける。
3. それ以外の XSheet 固有の変更は、別のファイル（`src/facade/` など）に置くか、写したファイルの中では `// XSheet:` の印を付けて 1 か所にまとめる（3-way のマージで衝突したときに見分ける）。
4. 新しい機能を XSheet で先に作るときは、InterHub の構成に合う置き場所（core＝`lib/xlsx`・render＝`lib/viewers/xlsx-*`）に置く。

**4-2. 同期の道具**（開発用の Node のスクリプト。依存を足さない・通信しない・InterHub には書き込まない）

| ファイル・コマンド | 役割 |
| --- | --- |
| `tools/sync/interhub.json` | InterHub のパスと XSheet のパスの対応（core・render・テスト・fixtures）、除くもの（`spreadsheet.ts` の .xlsx 以外の部分など）、変換の規則、**取り込み済みのコミット**（今は `943e4f0`）。InterHub の場所は環境変数 `INTERHUB_DIR` で替えられる |
| `npm run sync:status` | 取り込み済みのコミットより後で、対応表のパスに触れた InterHub のコミットとファイルを一覧にする（新しいファイル・変わったファイル・行数・XSheet での行き先）。 |
| `npm run sync:patch -- <範囲>` | InterHub の差分（`git diff`）のパスを行き先へ置き換え、変換の規則をかけたパッチを一時フォルダに書く。当てるのは人（`git apply --3way`）。テストと fixtures も同じ扱い |

- 取り込んだら `interhub.json` の取り込み済みのコミットを進め、`progress.md` に記録する。03 の表は写した時点の記録として残し、今の対応は `interhub.json` を正とする。
- InterHub に未コミットの変更があっても読まない（コミット済みのものだけを取り込む）。

**4-3. 取り込む順序**

- グラフは InterHub で作業中（ISSUE-105〜111）なので、今写すと何度も写し直すことになる。**先に M6（配り方）を終え、InterHub のグラフが一段落してから同期の道具で取り込む**（M7 の最初の項目）。グラフの描画は §1 の遅れて読む部品として足す。
- 長い目では写す向きを逆にするのが一番の解決になる。InterHub（M5b）と xrag-Viewer（M5a）が xsheet を使うようになれば、新しい機能は XSheet で作り、使う側は版を上げるだけになる。M5 は保留中（2026-10-29 ユーザー決定）なので、この ADR では始めない。

### 5. 文書と例

- README の先頭を「3 行で始める」にする：`<script>`・バンドラ・Node（読み取りだけ）の 3 通り。細かい指定・CSP・上限は後ろか `docs/guide/` へ。
- `docs/guide/`：React・Vue での置き方（mount で `view`、unmount で `destroy`）、Worker で読む、CSP、ダーク表示の調整。
- 例：`examples/standalone.html`（`xsheet.min.js` を読むだけ・ビルド不要・`file://` でも開ける。ファイルを選ぶ・落とすと表示する画面で、そのまま出来上がりのビューアとして配る）・`examples/esm.html`（`xsheet.mjs` を import）。今の Vite の例も残す。

### 6. 配り先

| 段階 | 配り方 | 条件 |
| --- | --- | --- |
| 今（非公開） | タグに `dist/` を含める（ADR-0002）＋GitHub Release に単独版の zip（`xsheet.min.js`・`xsheet.mjs`・`xsheet-geom.mjs`・`xsheet.css`・`xsheet.d.ts`・`xsheet.d.mts`・`xsheet.licenses.txt`・`viewer.html`・`NOTICE`。M6-3 で `xsheet.d.mts`（`xsheet.mjs` の型）と、束ねた依存のライセンスの全文 `xsheet.licenses.txt` を足した）を添付 | リポジトリの権限がある人 |
| 公開後 | npm に公開（`provenance` 付き）。jsDelivr・unpkg からも自動で読める | D3（ライセンス）を決め、リポジトリを公開する |

## 決めたこと（2026-10-30 ユーザー。提案のとおり）

| # | こと | 決定 |
| --- | --- | --- |
| Q1 | 「誰でも」の範囲 | まず社内（リポジトリの権限がある人）で M6 を作る。公開（D3・npm）は別に決める |
| Q2 | `xsheet` の入口を「全部」にする（§3） | する（今の公開物は残るので壊れない） |
| Q3 | `view` の既定の言語 | `'auto'`（ブラウザの言語・合わなければ英語）。`renderWorkbook` の既定（日本語）は変えない |
| Q4 | 韓国語の文言を同梱する | する（`MESSAGES_KO`） |

## 却下した案

| 案 | 却下の理由 |
| --- | --- |
| バンドラ向けの出力を単独版だけに置き換える | バンドラの使う側で jszip・ssf を重ねて持つことになり、形の表の遅れた読み込みも使う側のビルドで扱いにくくなる。2 系統のほうが両方に向く |
| UMD（AMD・CommonJS・大域を 1 つで） | 今の使う側で AMD・`require` が要るところが無い。IIFE と ESM の 2 つで足りる |
| 依存を ESM の版に替え、import map で読ませる | ssf に ESM の版が無い。import map を書かせること自体が手間（P1 の解決にならない） |
| URL を受けて中で fetch する（PDF.js の `getDocument(url)` のように） | 通信しない約束（CLAUDE.md）を崩す。`fetch(url)` を渡せば 1 行で済む |
| Web Components（`<x-sheet>`） | 便利だが属性で src を受けると通信になる。`view` で足りるので M7 の候補にする |
| InterHub と XSheet を monorepo にまとめる | InterHub は別のリポジトリで、変更しない約束（CLAUDE.md）。D4 で monorepo にしないと決めている |
| 写したファイルを XSheet の好みで整理し直す | 以後の InterHub の差分が当たらなくなる（§4-1） |

## 結果

- 良い点：`<script>` 1 つ・`import` 1 つ・1 行で表示できる。InterHub の変更は「一覧を見る → パッチを当てる → テスト」で取り込める。重い部品を足しても、図形のないシートの読み込みは重くならない。
- 注意：ビルドと検査が増える（単独版の大きさの上限・裸の import が残っていないこと・`eval` の無いこと・jsdom で `XSheet.view` が動くこと）。版を出す手順（02 §8）に単独版の zip を足す。
- 注意：`xsheet` の入口を「全部」にすると、Node で import したときに render の部品も読み込まれる（実行はしない）。DOM を使わないことを約束するのは `xsheet/core` だけになる。
