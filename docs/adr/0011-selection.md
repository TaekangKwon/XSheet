# 0011. セル・行・列の選択、キーボードでの移動、コピー

> 版: 1.0 ・ 更新日: 2026-11-26 ・ 担当: kwon
> 関連文書: [0007](0007-moobius-viewer-api.md)（§4 位置の判定・§5 倍率・§7 見出しを隠す） / [0010](0010-scroll-element-and-print.md)（§2-1 遅れて読む部品） / [0003](0003-universal-distribution.md)（§4-1 写しの規約） / Issue [#55](https://github.com/TaekangKwon/XSheet/issues/55)・[#56](https://github.com/TaekangKwon/XSheet/issues/56)・[#57](https://github.com/TaekangKwon/XSheet/issues/57)

## TL;DR
- `selection: 'row' | 'cell'` を指定した表示だけ、クリック・キーボードで行・列・セルを 1 つ選べる。既定は無効で、今の利用者の見た目と動きは変わらない。
- 帯は重ねる層（`div.xs-sel`）1 つで描く。選んだ行番号・列記号の色を変える。`getSelection`・`setSelection`・`onSelectionChange` を足す。
- 矢印・PageUp/PageDown・Home・Ctrl+Home/End で選択を移し、表が付いて来る。Ctrl+C で選択をタブ区切りの文字にしてコピーする（`selectionText()` でも取れる）。
- 中身は遅れて読む部品（`selection` のチャンク・単独版は `xsheet-selection.mjs`）。render の大きさの上限を守るため（ADR-0010 版 1.3 と同じ考え）。

---

## ステータス

proposed（2026-11-26）。Issue #55〜#57 の案から、下の推奨の案で作った（M7-24〜M7-26）。**ユーザーの決定を待つ**。決定が推奨と違えば、作った物を直す。

| # | 決めること | 推奨（作った物） | 理由 |
| --- | --- | --- | --- |
| 1 | 選択の形 | 1 つの範囲（行の全体・列の全体・セル）＋今のセル。範囲の選択（Shift・ドラッグ）は入れない | 読むだけのビューアで行を見失わないことが目的（#55）。範囲の選択は後で足せる形にしておく |
| 2 | 既定と有効にする口 | 既定は無効。`selection: 'row'`（セルのクリックで行の全体）か `'cell'`。指定しない表示では `setSelection` も効かない | 今の利用者の見た目を変えない。部品を読まない表示に選択の状態を持たせない（§3） |
| 3 | セルのクリックで選ぶ物 | `'row'` と `'cell'` の 2 つ。十字（行と列の両方）は入れない | 十字は帯が 2 つになり、コピーの中身も決まらない |
| 4 | シートを替えたとき | 選択を外す（シートごとに覚えない）。`onSelectionChange(null, 'api')` を呼ぶ | ハイライトと同じ（ADR-0007 §7）。覚えるなら後で足せる |
| 5 | 部品の読み方 | 遅れて読む部品。`selection` を指定した表示は、最初のシートを出す前に読み終える | 選択だけで render が 66,530 になり上限 65,000 を超えた。却下：render の上限を上げる（ADR-0010 版 1.3 で同じ理由で却下した） |
| 6 | シートのタブを Ctrl+PageUp/PageDown で替えるか | 入れない | ブラウザのタブの切り替えとぶつかる |
| 7 | 支援技術への伝え方 | 表示の根の中の見えない `aria-live="polite"` の要素に、選択の番地（`C5`・`5:5`・`C:C`）を入れる | `aria-activedescendant` は描き直しで要素が消える。番地は言語に依らないので文言を増やさない |
| 8 | コピーの中身 | `text/plain` だけ。見えている文字（表示形式を当てた後）をタブと `\r\n` で区切る。`text/html` は入れない | Excel・メモ帳に貼って表の形に戻る。書式付きは後で足せる |
| 9 | 非表示の行・列のコピー | 入れない（見えている物だけ） | XSheet はフィルターを読まず、フィルターで隠れた行と手で隠した行を区別できない。Excel もフィルターで隠れた行は入れない |
| 10 | コピーの上限 | 1,000,000 セル。超えたらコピーせず、通知の帯に出す（文言 `copyTooLarge`） | 行・列の全体を選んで大きなシートで押しても固まらないように |

## コンテキスト

- 今の XSheet には選択もキーボードの扱いもクリップボードの扱いも無い。ブラウザの既定の動き（文字のドラッグでの選択とコピー・フォーカスのある要素の矢印でのスクロール）だけ。
- 表は見えている近くの行・列だけを描く（仮想化）。スクロールで `tbody` を作り直す（`sheet-view.ts`）。
- 罫線の層・図の層は、スクロールの中の `div.xs-wrap` に `position: absolute` で重ねている。
- 位置の判定（`cellAt`）と、セルが見える所までのスクロール（`scrollToCell`）は ADR-0007 §3・§4 で作った。
- render の大きさ（圧縮＋gzip）は 0.7.0 で 64,405、上限 65,000。
- 写しの規約（ADR-0003 §4-1）：`sheet-view.ts`・`style.css` などは InterHub の写し。処理は新しいファイルに置き、写したファイルには `// XSheet:` の印の付いた口だけを足す。

## 決定

### 1. 口（render・view2）

```ts
export type SelectionMode = 'row' | 'cell';
export type SelectionType = 'row' | 'column' | 'cell';
export type SelectionOrigin = 'user' | 'api';

export interface Selection {
  readonly type: SelectionType;
  /** 1 始まり・両端を含む。行の全体は列 1〜16,384、列の全体は行 1〜1,048,576。結合したセルは結合の範囲 */
  readonly range: CellRange;
  /** 今のセル（結合なら左上）。行・列の全体でも、キーボードで動かす起点として覚えている */
  readonly active: { readonly r: number; readonly c: number };
}

export type SelectionInput =
  | { type: 'row'; r: number; r2?: number }
  | { type: 'column'; c: number; c2?: number }
  | { type: 'cell'; r: number; c: number };

export interface RenderOptions {
  selection?: SelectionMode;
  onSelectionChange?: (selection: Selection | null, origin: SelectionOrigin) => void;
}

export interface SheetView {
  getSelection(): Selection | null;
  /** @throws RangeError 種類が正しくない・1 以上の整数でない・Excel の範囲を超える番号 */
  setSelection(selection: SelectionInput | null): void;
  /** 選択をコピーする文字（§5）。選択が無い・上限を超えると null */
  selectionText(): string | null;
}
```

- クリック：セル → `selection` の指定（`'row'` は行の全体、`'cell'` はセル）。行番号 → 行の全体、列記号 → 列の全体。左上の角・表の外・スクロールバーは何もしない。結合したセルは結合の範囲（行の全体なら結合の行の範囲）。
- 押してから離すまでに 4px より動いたクリック・左以外のボタン・文字を選んでいる（ドラッグ・ダブルクリック）ときは選ばない。文字のドラッグでのコピーを壊さない。
- `onSelectionChange` は選択が変わったときだけ呼ぶ（同じ物を選び直しても呼ばない）。`'user'` はクリック・キー、`'api'` は `setSelection`・シートの切り替え。
- `setSelection` はスクロールしない（見せるには `scrollToCell`）。行・列は表の外でもよい（帯は表の範囲で切る）。
- 印刷の描画（`renderForPrint`）には選択を出さない。

### 2. 見た目

- 帯は `div.xs-wrap` の中の `div.xs-sel` 1 つ（`position: absolute`・`z-index: 0`・`pointer-events: none`）。罫線の層の後・図の層の前に置き、DOM の順でセル・はみ出しの文字・ハイライト・罫線の上、図と sticky の見出し・行番号の下を通る。位置は寸法（倍率を掛けた `Geometry`）から出し、表を描き直すたびに合わせ直す（スクロールだけでは動かさない）。
- 行・列の全体は半透明の地（`--xs-sel-bg`）。セルは地を付けず枠（`--xs-sel-line`・2px）。Excel と同じ。
- 選んだ行番号・列記号は `is-selected`（`--xs-sel-head-bg`・`--xs-sel-head-fg`）。表を描くたびに、描いた見出しに付け直す。セルの選択では行番号と列記号の両方、行の全体では行番号だけ、列の全体では列記号だけ。見出しを隠す表示（`headers: false`）では帯だけ。
- ダークのテーマの値も持つ（`[data-xs-theme='dark']`・`auto`）。

### 3. 遅れて読む部品

- 中身（`selection.ts`：帯・クリック・キーボード・コピー）は、`selection` を指定した表示を作るときだけ動的 import で読む。口は `selection-gate.ts`。表示は最初のシートを出す前に読み終えるのを待つので、`renderSheet`・`renderWorkbook` の約束が解決した後は、`getSelection`・`setSelection` がすぐ効く。
- `selection.ts` は表示の部品を import しない（単独版の部品は部品ごとに別に束ねるので、import すると `xsheet-selection.mjs` に表示の部品が二重に入る）。列の記号の関数は口から受け取る。
- 読めなかったとき（単独版の ESM で `xsheet-selection.mjs` を置いていない等）は、選択の無い表示として動く。
- 大きさ：M7-24 の時点で render 65,162（口と表示の口だけ。選択を中に入れると 66,530）、`selection` のチャンク 2,181。キーボード・コピーを足した M7-26 の時点は §5 の後の計測。上限は M7-27 で決め直す。

### 4. キーボード（M7-25・#56）

- `selection` を指定した表示は、スクロール枠（`div.xs-body`）に `tabindex="0"` を付ける。クリックで選んだらフォーカスを移す（`preventScroll`）。フォーカスの枠はキーで動かしているときだけ（`:focus-visible`・`--xs-accent`）。指定しない表示には付けない。
- キー（Excel・エクスプローラーのプレビューに合わせる。行き先の計算は `selection-keys.ts`）：
  - ↑・↓・←・→：前・次の行・列。行の全体の選択では ←・→ が今のセルの列を、列の全体の選択では ↑・↓ が今のセルの行を動かす（選択の種類は変えない）。結合・複数の行の選択は範囲の端から動く。
  - PageUp・PageDown：表示枠のセルの部分の高さだけ。Home：行の先頭の列。Ctrl+Home：左上。Ctrl+End：値のある最後のセル（`rowCount`・`colCount`）。Mac の ⌘ も Ctrl と同じ。
  - 高さ 0（非表示）の行・幅 0 の列は飛ばす。表の範囲（データと枠を埋める分）の端より先へは動かない。
  - 選択が無いときの最初のキーは、見えている左上のセルを選ぶ。
- 選択を動かしたキー（端で動けないときも）だけ `preventDefault` する。Tab・Ctrl+F などページの操作、Shift・Alt の付いたキー、End だけ（Excel の End モード）は取らない。入力欄（`input`・`textarea`・`select`・`contenteditable`）からのキーは見ない。
- 動いた先が見えていなければ、最小の量だけスクロールする（中央に寄せない。Excel と同じ）。見出しの sticky の下に隠れない位置にする（`scrollToCell` の 1/3 の位置とは違う）。
- 支援技術：表示の根の中の見えない `span.xs-live`（`aria-live="polite"`）に選択の番地を入れる（`C5`・`B2:D4`・`5:5`・`C:D`）。
- 範囲の選択（Shift＋矢印）は入れない（ステータスの 1）。

### 5. コピー（M7-26・#57）

- フォーカスのあるスクロール枠の `copy` イベントで `event.clipboardData.setData('text/plain', …)` し、`preventDefault` する。同期で書けて、`navigator.clipboard` の許可（権限の確認・HTTPS）が要らない。外部へは送らない。
- 文字をドラッグ・ダブルクリックで選んでいる（`getSelection()` が畳まれていない）ときと、選択が無いときは何もせず、ブラウザの既定のコピーに任せる。
- 中身（`selection-copy.ts`。`selectionText()` も同じ文字を返す）：
  - `text/plain` だけ。見えている文字（`CellView.text`。表示形式を当てた後）を、タブと `\r\n` で区切る。最後の行にも `\r\n`（Excel のコピーと同じ）。
  - タブ・改行・`"` を含むセルは `"` で囲み、`"` は `""` にする（#57 の案のとおり。Excel に貼ると元の文字に戻る）。
  - 結合したセルは左上に文字、残りは空。結合の左上が選択より上にあれば空（読む区切りの 1 行前から読んで捨てる。`getRows` は区切りの最初の行に左上の文字を置くため）。
  - 非表示の行・列（高さ 0・幅 0）は入れない（ステータスの 9）。
  - 行の全体は値のある最後の列（`colCount`）まで、列の全体は値のある最後の行（`rowCount`）まで。セルの選択は範囲のまま。
- 上限：見える行 × 見える列が 1,000,000 セルを超えたらコピーせず（`preventDefault` はする）、通知の帯に `noticeCopyTooLarge` を出す（文言は 24 キー目。メッセージにセル値・シート名を含めない）。`selectionText()` は null。
- 右クリックのメニュー・`text/html`（書式付きの表）は作らない（ステータスの 8）。

## 却下した案

| 案 | 却下した理由 |
| --- | --- |
| `td`・`tr` にクラスを付けて塗る | 表は見えている近くだけを描き、スクロールで作り直すので、描き直しのたびに付け直しが要る。見えていない部分（行の右の方）は塗れない |
| render の上限を上げて、選択を render に入れる | 使う人の少ない機能のために、上限は実測に約 1 割の余裕という決まりを崩す（ADR-0010 版 1.3 と同じ） |
| `selection` を指定しない表示でも `setSelection` を効かせる | 部品を読まない表示にも選択の状態を持たせることになる。`setSelection` を呼んだときに読むと、`getSelection` がすぐ結果を返せない |

## 結果

- 良い点：今の利用者の見た目・動き・最初に読む大きさは変わらない（口の 300 バイトほどだけ）。帯は描き直しの影響を受けない。
- 注意：`selection` を指定した表示は、最初のシートを出すのに部品の読み込みを待つ（単独版の ESM では `xsheet-selection.mjs` を同じフォルダに置く）。
- 注意：範囲の選択（Shift＋矢印・ドラッグ）は無い。足すときは `Selection` の形（1 つの範囲＋今のセル）のまま、範囲を広げる。
