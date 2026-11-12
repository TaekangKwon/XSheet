# 0007. 検索・ハイライト・拡大縮小などの公開 API（Moobius の依頼 R1〜R7）

> 版: 1.0 ・ 更新日: 2026-11-12 ・ 担当: kwon
> 関連文書: [0003](0003-universal-distribution.md)（§4-1 写しの規約） / [0006](0006-npm-publish.md)（npm の版） / [01_requirements](../01_requirements.md)（M7 の候補） / [05_quality](../05_quality.md)（③ 実機の確認）

## TL;DR
- Moobius（書庫管理のファイルビューア）からの依頼 R1〜R7 を M7 として受け、公開 API を足す：検索 `findText`、`SheetView` の `setHighlights`・`focusHighlight`・`scrollToCell`・`cellAt`・`highlightAt`・`setZoom`・`zoom`・`showSheet`、`RenderOptions` の `zoom`・`headers`。
- 検索は依頼の案（同期・配列）と違い、**非同期で、中断と件数の上限を持つ**（`Promise<{ matches, truncated }>`）。依頼の案と違う点は Moobius に伝える。
- 拡大縮小は**寸法に倍率を掛けて描き直す**（CSS の `zoom`・`transform` は使わない）。見出しには倍率を掛けない。
- 曜日（`aaa`・`aaaa`）と、書式に直接書かれた和暦（`ggg`・`e` など）はいつも表示する。**組み込みの書式 27〜36・50〜58 は既定では西暦のままにし、読み取りの指定 `builtinDates: 'japanese-era'` で和暦にする**。
- 版：曜日を 0.4.1 で先に出し、残りは 0.5.0 にまとめる。

---

## ステータス

accepted（2026-11-12）。ユーザーが検討の推奨どおりに決めた：検索の形（§1・非同期・中断・件数の上限）、拡大縮小の方式と見出しに倍率を掛けないこと（§5）、和暦を読み取りの指定にすること（§6）、受け入れから外す項目（§8）、版の割り振り（§9）。§1 の形は依頼の案と違うので、型定義（§9）と一緒に Moobius に伝える。

## コンテキスト

- Moobius は ExcelJS で読んだ値を画面側で表にしており、日付が 9 時間ずれるなどの問題がある。XSheet の単独版に置き換えたいが、検索語の位置を見せる機能（検索・ハイライト・移動・クリックの判定・拡大縮小）は XSheet の外から作れないため、7 項目を依頼してきた（依頼資料・2026-11-12。API の名前と形は案で、XSheet の設計に合わせて変えてよいとしている）。
- 外から作れない理由：
  - 描き直しのたびに表を丸ごと作り直す（`sheet-view.ts`）。外から付けた印は残らない。
  - セルの文字は `textContent` で入れるだけで、列の目印が無い（`grid-table.ts`）。
  - 倍率の指定が無く、見える範囲の計算は `scrollTop`・`clientHeight` を実寸の px として使う。
- 作るための部品は表示の中にそろっている：`Geometry` の `rowTop`・`colLeft`・`rowAt`・`colAt`、描き直しのたびに通る `buildGridTable`。
- 表示は「見えている行・列と前後の余分だけを DOM にする」仮想スクロールだが、検索はシートのデータ、ハイライトは表示が持つ一致の一覧、移動先は寸法から出すので、まだ描いていない行・列も扱える。
- 制約は写しの規約（ADR-0003 §4-1）。`sheet-view.ts`・`grid-table.ts`・`grid.ts`・`format.ts`・`style.css` は InterHub の写しなので、処理は新しいファイルに置き、写したファイルには `// XSheet:` の印付きで呼び出しの口だけを足す。

## 決定

### 1. 検索（R1・core）

```ts
export function findText(sheet: XlsxSheet, query: string, options?: FindTextOptions): Promise<FindTextResult>;

export interface FindTextOptions {
  match?: 'partial' | 'cell';   // 既定 'partial'
  ignoreCase?: boolean;         // 既定 true
  ignoreNewlines?: boolean;     // 既定 true（\r\n・\n・\r を照合のときだけ除く）
  skipHidden?: boolean;         // 既定 true
  signal?: AbortSignal;         // 中断
  maxMatches?: number;          // 既定 10,000。上限 100,000（超える値は 100,000 にする）
}
export interface FindTextResult {
  matches: TextMatch[];         // 行 → 列 → 位置の順
  truncated: boolean;           // maxMatches で打ち切ったか
}
export interface TextMatch { r: number; c: number; start: number; end: number } // 1 始まり。start・end は CellView.text 上の UTF-16 の位置（end は含まない）
```

- 照合のルールは依頼のとおり：表示形式を当てた後の文字（`CellView.text`）で照らす。全角と半角は区別する。1 セルの複数の一致はすべて数える。結合は左上のセルだけ（`getRows` が隠れるセルを null で返す）。空の検索語は空の結果。図形・グラフの中の文字は対象外。
- 改行を除いた文字列の位置から元の位置へ戻す対応表を作り、`start`・`end` は元の `text` の位置で返す。
- 非表示の行・列は、描く側と同じく `createGeometry` の高さ・幅 0 で決める（非表示の印が無い高さ 0 の行も同じく除く）。
- 約 10ms 調べるごとに `setTimeout` で 1 回休む。区切りの長さは実機で測って直す。
- 中断されたら `AbortError`（`DOMException`）で reject する。
- 入口は `xsheet/core` と単独版の `XSheet.findText`。
- 1 シートずつ調べる。ブック全体は呼ぶ側がシートを 1 つずつ開く（`openSheet` は呼ぶたびに XML を展開し直す。開いたシートを使い回すかは呼ぶ側が決める）。

### 2. ハイライト（R2・render）

```ts
interface SheetView {
  setHighlights(matches: readonly TextMatch[] | null): void; // null・空配列で消す
}
```

- 一致の一覧は「行 → 列 → 範囲」の 2 段の表にして表示が持ち、描くたびにその行・列の一致を付け直す。
- セルの文字を一致の範囲で分け、**分けた全体を `span.xs-hits` 1 つで包む**。`.xs-cell`・`.xs-text` は `display: flex` なので、包まずに分けると各片が別の flex 項目になって片の端の空白が消える。
- 一致の片は `span.xs-hit`（`data-xs-hit` に一覧の番号）。色は CSS 変数 `--xs-hit-bg`（既定 `rgba(255, 255, 0, 0.4)`）・`--xs-hit-current-bg`（既定 `#ffa500`）。背景だけを付け、行の高さと列の幅を変えない。
- 自動の色の文字が塗りをまたいではみ出すとき（`background-clip: text` で描く）は、片に `--xs-fg` の文字の色を付けて、不透明なハイライトで文字が隠れないようにする。
- `###` で表示されるセル（数値が列幅に収まらない）は、一致に数え、セル全体を光らせる。

### 3. 今のヒットと移動（R3・render）

```ts
interface SheetView {
  focusHighlight(index: number, options?: { scroll?: boolean; smooth?: boolean }): void; // -1 で外す
  scrollToCell(r: number, c: number, options?: { smooth?: boolean }): void;
}
```

- 縦は、セルの下端から 1 行の高さを引いた位置（`rowTop(r + rowSpan)` − 行の高さ）が表示枠の上から約 1/3 に来るようにする。文字は下揃えなので、高い行・縦の結合でも文字が見える。
- 横は、列が見えていないときだけ動かす。はみ出す文字の中の一致は、一致の左端までの文字の幅を、画面の外の測るための要素（同じ書式のクラス・同じ文字の大きさ）で測ってから位置を決める。
- **スクロール先はスクロールする前にすべて計算し、`scrollTo` は 1 回だけ呼ぶ**。描いた後に測って直す方法は、描き直しが次のフレームになることと smooth の途中の描き直しのため成り立たない。
- 横は「一致が見える最小の移動」（一致を表示枠の右端に寄せる）にする。はみ出し元のセルから約 2 画面より遠い一致は、元のセルが描く範囲から外れて見えない。これは制約として残す。
- 今のヒットを替えるとき、相手が描いた範囲の中なら `data-xs-hit` で探してクラス（`is-current`）だけを替える。表を描き直さない。

### 4. 位置の判定（R4・render）

```ts
interface SheetView {
  cellAt(clientX: number, clientY: number): { r: number; c: number } | null; // 見出し・余白・表の外は null
  highlightAt(clientX: number, clientY: number): number;                  // 無ければ -1
}
```

- `cellAt` は `Geometry` の `rowAt`・`colAt` で出す（DOM を見ない）。結合は左上、はみ出す文字の上ははみ出し元のセルを返す。結合の範囲は描いた行の `CellView`（`rowSpan`・`colSpan`）から引く。
- `highlightAt` は `elementsFromPoint` で `.xs-hit` を拾い、`data-xs-hit` を返す。

### 5. 拡大縮小（R5・render）

```ts
type Zoom = number | { fit: 'width'; min?: number; max?: number };
interface SheetView {
  setZoom(zoom: Zoom): void;
  readonly zoom: number;        // fit のときは計算した値
}
interface RenderOptions {
  zoom?: Zoom;                  // 最初から倍率を指定して描く（ちらつきを防ぐ）
}
```

- **方式：寸法に倍率を掛けて描き直す。** `Geometry` を倍率で包む関数（`src/render/zoom.ts`）を作り、表・罫線の層・§3・§4 は包んだ `Geometry` を使う。線の幅（td の 1px の枠・罫線）には倍率を掛けないので、縮小しても 1px の罫線が消えない。
- 倍率は 0.5〜2.0。範囲の外は端の値にする。数でない値（`NaN` など）は `RangeError`。fit の `min`・`max` の既定は 0.5・2.0。
- 包んだ寸法の境目は整数にする：`top′(r) = round(top(r) × z)`、高さ＝`top′(r+1) − top′(r)`。小数のままだと罫線の層（SVG）と td の線が最大 1px ずれ、線がにじむ。境目を丸めるので合計の誤差はたまらない。
- **見出し（列見出しの高さ・行番号の列の幅・見出しの文字）には倍率を掛けない。** 見出しの高さは `grid.ts` の定数と `style.css` の `24px` に結び付いていて、写したファイルの直す所が増えるため。画面の座標は「倍率を掛けない見出し＋倍率を掛けたセル」なので、`rowAt`・`colAt` の逆算は、座標から見出しの寸法を先に引いてから倍率で割り、前後 1 行（列）を整数の境目と比べて合わせる。Excel（見出しも拡大縮小する）と見た目が違う点は残る。
- 文字の大きさは `calc(13px * var(--xs-zoom))`。書式ごとの大きさは `em` の比なので一緒に変わる。数値が収まるかの判定に使う最大桁幅（MDW）にも倍率を掛ける。
- 図の層は中を 100% の寸法で描き、層全体に `transform: scale()` を掛ける。図の出し入れの判定（`viewWindow`）は座標を倍率で割る。
- 倍率を替えるときは、替える前の左上のセル（行・列とセルの中の位置）を覚えておき、替えた後にそこへスクロールする。
- 幅に合わせる（fit）の幅は、データのある列と図の右端まで。表示枠の大きさの変化（今ある `ResizeObserver`）で合わせ直す。縦のスクロールバーの出入りで揺れないよう、スクロールバーの幅を引いた幅で計算し、変化が 1% 未満なら合わせ直さない。
- 使う側は、表示の外側に `transform: scale()` を掛けない（`getBoundingClientRect` とスクロールの値がずれて §4 が狂う）。

### 6. 曜日・和暦（R6・core）

- 日付の書式（`ssf.is_date` が真）のとき、ssf に渡す前に書式を小さな字句の分解で読み、`aaa`・`aaaa`（曜日）と `g`・`gg`・`ggg`・`e`・`ee`（元号・和暦の年）を、計算した文字を引用符で包んだ形に置き換える（`src/core/format-ja.ts`、`format.ts` には ssf を呼ぶ手前に口を 1 つ）。
  - 置き換えないもの：引用符の中、`\` の後、`[...]` の中、`AM/PM`・`A/P`。`;` で分かれた区画ごとに扱う。
  - `e`・`ee` は ssf が西暦の年と読むので、和暦が有効な区画（`g` があるか、`[$-411]`・`[$-ja-JP]` があるとき）だけ置き換える。
  - 曜日は `Date` を使わず、日付の通し番号から整数の計算で出す（1900 年基準と 1904 年基準を分ける。Excel の 1900 年のうるう年の誤り（通し番号 60）に合わせ、Excel の `WEEKDAY` と同じ結果にする）。時間帯のずれが入らない。
- 曜日と、書式に直接書かれた和暦は**いつも**表示する（今は `aaa`・`ggg` がそのまま文字で出ている不具合の修正）。
- **組み込みの書式 27〜36・50〜58 は既定では今のまま西暦**にし、読み取りの指定で和暦にする：

  ```ts
  interface OpenXlsxOptions { builtinDates?: 'default' | 'japanese-era' } // 既定 'default'
  interface ReadOptions { builtinDates?: 'default' | 'japanese-era' }     // XSheet.read・view でも使える
  ```

  - 既定を変えない理由：これらの番号の見た目は、ファイルではなく開く Excel の言語環境が決める。韓国語・中国語の環境のファイルも同じ番号を自分の言語の日付に使うので、既定を和暦にするとそれらが間違って見える。
  - 読み取りの指定にする理由：表示の文字（`CellView.text`）は読み取りのときに core が作る。描画の指定にすると、検索（§1）が見る文字と画面がずれる。
  - 使わない方法：ファイルから言語環境を判定する（xlsx に表示の言語環境の情報が無い）、`view` の `lang` を使い回す（`lang` は画面の文言の言語で、ファイルの言語環境ではない）。
  - 指定の表（番号ごとの書式）は、日本語版 Excel で各番号の見え方を確かめてから作る。
- 和暦を 1 次の範囲外とした InterHub の決定（InterHub ADR-0025・`format.ts` の冒頭）は、XSheet ではこの ADR で変える。
- 元年を「元」と出す指定（`[$-ja-JP-x-gannen]` と言われるもの）は、Excel で作ったファイルで確かめてから扱うかを決める。確かめるまでは「1」と出す。

### 7. 見出しを隠す・シートを替える（R7 と追加）

```ts
interface RenderOptions {
  headers?: boolean;            // 既定 true。false で行番号の列と列見出しを出さない
}
interface SheetView {
  showSheet(index: number): Promise<void>; // SheetInfo.index。中の SheetViewInstance.showSheet を出す
}
```

- `headers: false` は行番号の列の幅と見出しの高さを 0 として計算する。見出しの寸法を定数からインスタンスの値に替える（写したファイルに口が要る）。
- `showSheet` は依頼に無いが足す。今はシートを替えるたびに表示を作り直すしかなく、倍率とハイライトが毎回消える。
  - シートを替えたら、ハイライトと今のヒットは消す（一致の一覧はシートごと）。倍率は保つ。fit なら合わせ直す。
- §2〜§5・§7 の `SheetView` の追加は、`renderWorkbook` の `WorkbookView` も同じ形で持つ（中のシートの表示に渡す）。`view` の `Viewer` に足すかは別に決める。

### 8. 受け入れの確認から外す項目

依頼の受け入れ項目のうち、XSheet がまだ描かない表示に頼るものは今回の受け入れから外し、その表示を入れるとき（01_requirements の M7 の候補）に足す。

| 依頼の項目 | 外す理由 |
| --- | --- |
| 右寄せ・中央揃えのセルのはみ出しへのハイライト | 指定の揃えを描かない（揃えはセルの型で決まる） |
| 改行の前後 2 行にまたがるハイライト | 折り返しを描かない（改行は空白に見え、1 行で描く） |
| 書式の境目をまたぐハイライトの見え方 | 部分書式を描かない（ハイライトは 1 つ付くので確かめられる） |
| 大きなシートの先頭行・先頭列の固定 | ウィンドウ枠の固定を読まない。§3 の「上から 1/3」は固定の無い表示枠を基準にする |

### 9. 版と進め方

- 0.4.1：曜日（§6 の `aaa`・`aaaa`）。`main` から出すので、Unreleased にある表の見た目の変更（Issue #23。`renderSheet` の表の色も変わる）も入る。Moobius に伝える。
- 0.5.0：§1〜§5・§6 の和暦・§7。依頼のとおり R2〜R5 は分けずに 1 つの版にする。
- npm への公開は 0.6.0 以降に送る（ADR-0006 §4 の「0.5.0 の予定」は、公開の作業を再開するときに直す）。
- 型定義（`.d.ts`）を実装より先に作って Moobius に渡す（画面側の作業を並行できるように。§1 の形が依頼の案と違うことも、ここで伝える）。
- 作業は Issue に分けた（04 の「M7 Moobius の依頼」・GitHub #25〜#34）：曜日（M7-1）→ v0.4.1（M7-3）、型定義の先出し（M7-2）、検索 → ハイライト → 移動・判定 → 拡大縮小 → 見出しを隠す・`showSheet`（M7-4〜M7-7・M7-9）、和暦の指定（M7-8）、v0.5.0（M7-10）。
- 写したファイルへの口は 15 か所前後になる（ハイライトの文字の挿入、`buildGridTable` の引数、`Geometry` の包み、表示枠の大きさの割り算、MDW、図の層、見出しの寸法のインスタンス化など）。どれも `// XSheet:` の印を付け、処理は新しいファイル（`find-text.ts`・`format-ja.ts`・`highlight.ts`・`zoom.ts`）に置く。

### 10. 確かめ方

| 対象 | 方法 |
| --- | --- |
| §1（中断で `AbortError`・上限で `truncated`・依頼の受け入れのシートで依頼の案と同じ一致）・§2 の DOM（`span.xs-hits` 1 つ・クラス）・§6（曜日が `WEEKDAY` と同じ・字句の分解・指定の有無で 27〜36 の表示・`findText` と画面の文字が同じ）・§7 | vitest＋jsdom。データは `src/testing/make-xlsx.ts` の合成か、Excel で自作したファイル（依頼側のサンプルは使わない） |
| §3・§4・§5（スクロール先、倍率 50%・75%・150% で罫線の層の線と td の境目が合う、見出しの上を `cellAt` がセルにしない、倍率を替えても見ていたセルが残る、縮小時の描き直しの時間） | 実機のブラウザ（05_quality の ③） |

## 却下した案

| 案 | 却下の理由 |
| --- | --- |
| 検索を依頼の案どおり同期・配列だけにする | 大きなシートやありふれた文字（例：「a」）で画面が固まり、メモリも際限なく使う。単独版は Worker を使わずページのスレッドで動く。上限（`XlsxLimits`）を置く方針に合わない |
| 拡大縮小を CSS の `zoom` でする | 縮小すると 1px の罫線が消える（依頼側で確認済み）。スクロール・判定・図の 4 か所に換算が要る |
| 拡大縮小を `transform: scale()` でする | 線が薄くなる・文字がにじむおそれ。スクロールの大きさを別の要素で作る必要があり、換算も要る |
| 見出しにも倍率を掛ける | 写したファイルの直す所が増える（`grid.ts` の定数と `style.css` の `24px`）。Moobius は見出しを隠す見込み（§7） |
| 組み込みの書式 27〜36・50〜58 を既定で和暦にする | ほかの言語環境のファイルが間違って見える |
| ハイライトを CSS Custom Highlight API で塗る | `elementsFromPoint` で拾えず §4 の判定を別に作ることになる。対応ブラウザも確かめていない。`span` 1 つで包めば空白の問題は起きない |
| 移動の後に `span` の位置を測って直す | 描き直しは次のフレームで、smooth の途中にも描き直すため、測る時点が決まらない |

## 結果

- 良い点：Moobius は ExcelJS をやめて XSheet の単独版に置き換えられる。検索・ハイライト・拡大縮小は他の使う側も使える。倍率 1 では今と同じ道を通るので、既存の表示を壊しにくい。
- 注意：公開 API が増え、0.x の間も互換を守る範囲が広がる。
- 注意：写したファイルへの口が 15 か所前後に増え、InterHub を取り込むときに 3-way のマージで衝突しやすくなる。
- 注意：縮小すると同じ画面に入る行・列が増え、DOM が増える（50% で td がおよそ 2.7 倍という見積もり。測っていない）。遅ければ倍率 1 未満で横の余分を減らす。
- 注意：まだ確かめていないこと：`scrollend` と CSS Custom Highlight API の対応範囲、ブラウザの要素の大きさの上限（とても大きなシートを拡大したとき）、最小のフォントの大きさの設定の影響、日本語版 Excel での組み込みの書式 27〜36・50〜58 の見え方、元年の表示。
