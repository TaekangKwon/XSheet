# 0010. スクロール枠を出す口と、抜けの無い印刷の描画（InterX-Viewer の依頼 R15・R16）

> 版: 1.3 ・ 更新日: 2026-11-26 ・ 担当: kwon
> 関連文書: [0007](0007-moobius-viewer-api.md)（§5 倍率・§7 `SheetView` と `WorkbookView` の形をそろえる） / [0008](0008-align-wrap-spill.md)（§4 はみ出しの届く列） / [0003](0003-universal-distribution.md)（§4-1 写しの規約） / Issue [#53](https://github.com/TaekangKwon/XSheet/issues/53)・[#54](https://github.com/TaekangKwon/XSheet/issues/54)

## TL;DR
- R16（#53）：依頼の案（セルで位置を読む・移す・知らせる）は作らず、**スクロール枠の要素を `scrollElement` として出す**だけにする。InterX-Viewer の `scroll` は縦の割合（0〜1）で、割合の計算・`'user'` と `'api'` の区別・1 フレーム 1 回の知らせは InterX-Viewer の `createScrollSync` が要素 1 つから作れるため。
- R15（#54）：`renderForPrint(target, options)` を足す。今のシートのデータの範囲を、**XSheet が用紙の大きさのページに切り分けて**、仮想化せずに描く。ブラウザの改ページには任せない（幅のはみ出しは次のページに送られず切れるため）。今の表示（倍率・位置）には触らない。
- 印刷範囲（`_xlnm.Print_Area`）は core が今は読まないので、0.7.0 には入れず次の版に送る。0.7.0 はデータの範囲を印刷する。
- 版は 0.7.0（公開の型が増える）。InterX-Viewer の peer の範囲（`>=0.5.0 <0.7.0`）は InterX-Viewer の側で上げてもらう。

---

## ステータス

accepted（2026-11-26）。ユーザーが次の 6 点を推奨どおりに決めた。§1・§2 の形は依頼の案と違うので、#53・#54 にコメントで伝える。版 1.1（M7-20）は、実装で確かめたこと（§2-4 のテーマの当て方・§2-5 の計測・§2-6 の結合セルと空のページ）を書き足しただけで、決めたことは変えていない。版 1.2（M7-21）も、3 つのブラウザで確かめたこと（§2-4・§4）を書き足しただけ。版 1.3（M7-21）は、ユーザーの決定で印刷の処理を遅れて読む部品にした（§2-1）。

| # | 決めたこと | 理由 |
| --- | --- | --- |
| 1 | R16 は `scrollElement` を出すだけにする（§1） | 依頼元の能力は縦の割合だけで、セルの位置からは割合を出せない |
| 2 | 印刷のページへの切り分けは XSheet がする（§2） | ブラウザ任せでは右の列が切れ、「抜けなく」を満たせない |
| 3 | 印刷の目盛り線・見出し（列記号・行番号）は既定で出さない。指定で出せる（§2-4） | Excel の印刷の既定と同じ。画面とは見え方が変わる |
| 4 | 用紙の大きさの既定は A4 縦・Excel の「標準」の余白（659×978 px）（§2-1） | 使う側が何も指定しなくても 1 ページの大きさが決まる |
| 5 | 印刷するセルの数の上限の既定は 200,000（§2-5） | 仮の値。M7-20 の実機で DOM の重さを測って決め直す |
| 6 | 印刷範囲は 0.7.0 に入れず、次の版に送る（§3） | 依頼の完了の条件はデータの範囲で足りる。core の読み取りは別の危険がある |

## コンテキスト

- 依頼元：InterX-Viewer（https://github.com/TaekangKwon/InterX-Viewer）の xlsx プラグインは `renderWorkbook` で表を描く（`src/plugins/xlsx/engine.ts`）。InterX-Viewer の Issue #38 の #1（印刷・優先 高）と #8（xlsx のスクロールの位置・優先 低）を受けて、XSheet に #54・#53 として頼んできた（2026-11-26）。
- InterX-Viewer の `scroll` の能力（InterX-Viewer ADR-0017 §2・`src/core/types.ts` の `ScrollCapability`）：
  - 位置は縦の割合だけ（`scrollTop ÷ (scrollHeight − clientHeight)`）。横は持たない、画素は却下と ADR に書いてある。
  - `onChange(ratio, origin)` の `origin` は `'user' | 'api'`。割合が前に知らせた値と変わったときだけ、1 フレームに 1 回知らせる。
  - 共通の部品 `createScrollSync`（`src/ui/scroll-sync.ts`）が、スクロールする要素 1 つから上の全部を作る。`set` の後に来た scroll のイベントは、位置が 1px 以内なら `'user'` に数えない。
- InterX-Viewer の印刷の依頼（Issue #38 の #1）の条件：選んだシートが抜けずに印刷される。印刷が終わる・取り消すと元の表示（倍率・位置）に戻る。印刷の画面（`window.print()`）を出すのは InterX-Viewer。
- XSheet の今の作り：
  - スクロール枠は `div.xs-body`。`createSheetView` が 1 回だけ作り、シートを替えても同じ要素を使う（`sheet-view.ts`）。シートを替えると位置を左上に戻す。
  - 行も列も、見えている近くだけを描く（仮想化）。図の層（`createImageLayer`）も見える範囲の図だけを DOM に出す（`update(view)`）。
  - 表の大きさ（`gridExtent`）は「データ・図の右下・はみ出しの届く列・枠を埋める分」の大きい方。枠の大きさに 0 を渡すと、枠を埋める分が入らない。
  - 書式の CSS は `.xs-sheet[data-xs-fmt="n"] td.…` に限って当たる（`format-css.ts`）。罫線は `svg.xs-borders` 1 枚（`border-layer.ts`）。目盛り線は `td` の枠線（`style.css` の `--xs-border`）。
  - ダークのテーマは、表示の根かその祖先の `data-xs-theme` で CSS 変数を替える（`style.css`）。
  - core は `workbook.xml` の `<definedNames>` も、`<pageSetup>`・`<printOptions>`・改ページも読まない。
- 写しの規約（ADR-0003 §4-1）：`sheet-view.ts`・`grid-table.ts`・`grid.ts`・`style.css` は InterHub の写し。処理は新しいファイルに置き、写したファイルには `// XSheet:` の印の付いた呼び出しの口だけを足す。

## 決定

### 1. スクロール枠を出す（R16・render）

```ts
export interface SheetView {
  // …今の項目…
  /**
   * スクロール枠（仮想化の窓）。表示を作ってから `destroy` まで同じ要素。表が無い表示（開けないブック・シートが 0 枚）では null。
   * 位置を読む・`scrollTop` を替える・`scroll` のイベントを受けるのに使う。大きさ・スタイル・子要素は替えないこと。
   * シートの切り替え（左上に戻す）・倍率の替え直し（左上のセルを残す）でも位置が変わり、`scroll` のイベントが来る。
   * XSheet: ADR-0010 §1
   */
  readonly scrollElement: HTMLElement | null;
}
```

- `WorkbookView` は `SheetView` から受け継ぐ（ADR-0007 §7 と同じ）。確かめ用の `view2` の戻り値にも足す。公開の `view` の `Viewer` には足さない（`scrollToCell` も無いため。ADR-0007 §7 の未決と同じ扱い）。
- 外から `scrollTop` を替えると、利用者のスクロールと同じ道（scroll のイベント → 次のフレームで描き足す）で描く。XSheet 側の処理は増えない。
- InterX-Viewer は `createScrollSync(() => (view.scrollElement ? { element: view.scrollElement, events: view.scrollElement } : null), window)` で `scroll` を作れる（形は InterX-Viewer が決める）。
- シートを替えたときに左上へ戻すのは今のまま。InterX-Viewer はこれを `'user'` の割合 0 として受け取る。InterX-Viewer ADR-0017 §2 の「割合が変わったときだけ知らせる」により、2 つの表の間で知らせは往復しない。

### 2. 印刷の描画（R15・render）

#### 2-1. 口

```ts
export interface PrintOptions {
  /** 1 ページの中身の大きさ（CSS px。96px = 1in）。既定は A4 縦から Excel の「標準」の余白を引いた 659×978 */
  page?: { width: number; height: number };
  /** 'width' は全部の列が 1 ページの幅に収まるように縮める（縦だけページを分ける）。既定は縮めない */
  fit?: 'width';
  /** 列記号・行番号を各ページに出す。既定 false */
  headers?: boolean;
  /** 目盛り線（セルの境の薄い線）を出す。既定 false。罫線はいつも出す */
  gridlines?: boolean;
  /** 描くセルの数の上限（行 × 列。非表示の行・列は数えない）。既定 200000 */
  maxCells?: number;
  signal?: AbortSignal;
}

export interface PrintResult {
  /** 描いたページの数（表の無い表示・シートを開けていないときは 0） */
  readonly pages: number;
  /** 上限で行を打ち切った */
  readonly truncated: boolean;
  /** `target` に足した要素・blob URL を放す */
  destroy(): void;
}

export interface SheetView {
  /**
   * 今のシートを、印刷のために仮想化せずページに分けて `target` の中に描く。図・グラフ・画像を描き終えてから解決する。
   * 今の表示（倍率・位置・ハイライト）には触らない。印刷の画面は使う側が出し、終わったら `destroy` する。
   * @throws DOMException AbortError（中断したとき。足した物は消してから拒否する）
   * XSheet: ADR-0010 §2
   */
  renderForPrint(target: HTMLElement, options?: PrintOptions): Promise<PrintResult>;
}
```

- 描くのは表示中のシート（`activeSheet`）だけにする。依頼の `opts.sheet` は持たない（InterX-Viewer は選んだシートだけを印刷するため）。
- 描く要素の根は `div.xs-print`（中に `div.xs-print-page` をページの数だけ並べる）。`target` 自身には何も付けない。
- 版 1.3（M7-21・ユーザーの決定）：印刷の処理（`print.ts`）は遅れて読む部品にする。印刷の描画を render に入れると、render の大きさ（圧縮＋gzip）が 66,171 で上限 65,000 を超えた（M7-19 では 63,762）。印刷はたまにしか使わないので、グラフの描画（`chart-draw`）と同じく、`renderForPrint` を呼んだときだけ動的 import で読む（バンドラ向けは `print` のチャンク 3,110、単独版の ESM は `xsheet-print.mjs` 2,820）。`print.ts` は表示の部品（表・罫線・図の層など）を import せず、`print-gate.ts` の `PRINT_KIT` で受け取る（単独版の部品は部品ごとに別に束ねるので、import すると表示の部品が二重に入る）。却下した案：render の上限を上げる（上限は実測に約 1 割の余裕という決まりを、使う人の少ない機能のために崩す）。注意：単独版の ESM を使う側は `xsheet-print.mjs` も同じフォルダに置く。
- `@page`（用紙の大きさ・余白）は XSheet から入れない。使う側のページ全体に効いてしまうため。使う側は `@page` と `page` を同じ大きさにする。README に例（A4 縦・余白 0.75in／0.7in）を書く。

#### 2-2. ページへの切り分け

- 描く範囲：表示中のシートの `gridExtent` に、枠の大きさ 0 を渡した範囲（データ・図の右下・はみ出しの届く列。枠を埋める空のセルは入れない）。はみ出しの届く列は全部の行から出す（印刷では全部の行を読むため）。
- 倍率：いつも 100%（今の表示の倍率は使わない）。`fit: 'width'` のときだけ、`zoomGeometry`（ADR-0007 §5）で全部の列が `page.width` に収まる倍率を掛ける（1 より大きくはしない）。
- 切り分け：行は `page.height` に収まるところで、列は `page.width` に収まるところで、行・列の境で切る。順は Excel の既定と同じく「下へ、それから右へ」。非表示の行・列は印刷しない（高さ・幅 0 なので自然に飛ぶ）。1 行（1 列）だけでページを超えるときは、その行（列）だけのページにして、はみ出した分を切る。
- 各ページは、ページの範囲を窓にした固定の表示とする：表は `buildGridTable` を、その範囲の行と列で呼ぶ。罫線は範囲の分を `svg.xs-borders` 1 枚に描く。図の層は範囲を窓にして `update` する。ページは `overflow: hidden` で、ページをまたぐ図は両方のページに描いて端で切る。ページの間は `break-after: page`。
- ブラウザの改ページには任せない。幅のはみ出しは次のページに送られない。大きな SVG や絶対位置の図がページの境でどう切れるかもブラウザで違う。ページ 1 枚を 1 つの塊にすれば、どのブラウザでも同じ切れ方になる。

#### 2-3. 待つ物と中断

- グラフ（`XlsxBook.readChart` の読み込みは非同期。`chart-store.ts`）と画像（`img.decode()`）を描き終えてから解決する。描けないグラフは画面と同じ破線枠。
- `signal` は描き始める前・ページを 1 枚描くごと・待つ間に見る。中断したら、足した要素を外し blob URL を放してから `AbortError` で拒否する。
- `renderForPrint` の途中で表示の `destroy`・シートの切り替えが来たら、`AbortError` と同じに片付けて拒否する。

#### 2-4. 見た目

- `div.xs-print` に `print-color-adjust: exact`（と `-webkit-print-color-adjust: exact`）を付け、塗り・罫線の色が印刷で消えないようにする。
- テーマは印刷ではいつもライト：`.xs-print` の上で `--xs-*` をライトの値に定義し直す（祖先の `data-xs-theme="dark"` を受け継がない）。
  - 版 1.1：値を写して並べず、変数の既定値のブロックのセレクタを `:where(:root, :host, .xs-print)` にした（同じ値を 2 か所に持たない。詳細度は 0 のままで、`.xs-print` に直に当たるので祖先の値に勝つ）。
- 版 1.2（M7-21 で確かめた）：
  - 各ページに `contain: strict` を付ける。`fit: 'width'` で図の層を縮めると、層の中の図（100% の位置）を Chrome・Edge が印刷の幅に数え、文書全体を縮めて出していた（A4 横で 0.67 倍など。`overflow: hidden`・`contain: paint` では止まらない）。
  - 幅に合わせて縮めた印刷では、列記号も行番号と同じく倍率で小さくする（5px より小さくしない）。13px のままだと狭い列で切れる。
  - 使う側の注意（README）：画面の `color-scheme: dark` のまま印刷すると、Chrome・Edge は用紙の余白を暗く塗る（ページの背景は印刷で落ちる）。印刷では `color-scheme: light` にする。
- 書式の CSS は `div.xs-print` の中に専用の `<style>`（`nonce` を付ける）を置き、`.xs-print` に `data-xs-fmt` を付けて当てる。表示の `destroy` の後でも印刷の描画が崩れない。
- 目盛り線は既定で出さない（`gridlines: true` で出す）。見出しは既定で出さない（`headers: true` で、各ページに列記号と行番号を出す）。core は `printOptions` を読まないので、ファイルの指定には従わない。

#### 2-5. 大きいシート

- 描くセルの数が `maxCells` を超えたら、超える手前の行で打ち切り、`truncated: true` を返す（列は打ち切らない）。
- 200,000 は仮の値。100 列なら 2,000 行。M7-20 の実機で DOM の重さと印刷のプレビューの速さを測って決め直す。
- 版 1.1（M7-20 の計測）：Chrome（アプリの中のブラウザ）で 20,000 行 × 10 列の合成のシートを A4 縦で描いた。
  - 50,000 セル：105 ページ・要素 約 11.7 万・描画 0.4 秒。100,000 セル：209 ページ・約 23.4 万・0.7 秒。200,000 セル：417 ページ・約 46.8 万・1.2 秒（画面に出したときのレイアウトがさらに 4.3 秒）。
  - セルの数に比例して増えるので、既定の 200,000 は変えない。ブラウザの印刷のプレビューを作る時間は測っていない（M7-21 の実機で見る）。
  - ページは文書の外で組み、描き終えてから 1 回だけ `target` に取り付ける。取り付けたまま途中で譲ると、譲るたびに増えていく木の全体のスタイル・レイアウトを計算し直し、200,000 セルで 27 秒かかった。

#### 2-6. 制限（受け入れる）

- ページの境をまたぐはみ出しの文字・結合セルは、境で切れる。
- 版 1.1（M7-20 で確かめた）：境をまたぐ結合セルは、ページごとに範囲の中へ縮め、続きのページの最初の行にも左上の文字を出す（`getRows` が範囲より上の結合の左上を範囲の最初の行に置くため。仮想化の窓の端と同じ）。Excel は左上のページにだけ出すが、続きのページでも中身が分かるのでこのままにする。
- 何も描く物（値・書式・罫線・図）の無いページは出さない（Excel と同じ）。
- 改ページ（`rowBreaks`・`colBreaks`）・`fitToPage`・印刷タイトル（`_xlnm.Print_Titles`）・ヘッダーとフッターは扱わない。

### 3. 印刷範囲（次の版・core と render）

- core：`workbook.xml` の `<definedName name="_xlnm.Print_Area" localSheetId="n">` を読み、`SheetInfo` に `printArea?: readonly { r1: number; c1: number; r2: number; c2: number }[]`（1 始まり。型の名前は実装で決める）を足す。シート名の引用（`'Sheet 1'!$A$1:$H$50`）・カンマで並ぶ複数の範囲を読む。`#REF!` を含む物は無視する。範囲の数の上限を `XlsxLimits` に足す。
- render：印刷範囲があれば、§2-2 の範囲の代わりに使う。複数の範囲は範囲ごとにページを分ける（Excel と同じ）。
- 0.7.0 には入れない（ステータスの 6）。入れないあいだ、印刷範囲を決めたシートは Excel より広く（データの範囲全部）印刷される。抜けは無い。

### 4. 確かめ方

- 単体テスト（jsdom）：
  - `scrollElement`：シートを替えても同じ要素、表が無い表示では null。
  - ページへの切り分け：行の高さ・列の幅・非表示・1 ページを超える行・`fit: 'width'`。
  - `maxCells` の打ち切り、`AbortSignal`、`destroy` の後に何も残らないこと。
  - `.xs-print` のテーマとクラス。
  - テストのファイルは `make-xlsx.ts` で作る。
- 実機（examples）：
  - 2 つの表を左右に並べ、片方の `scrollElement` の割合をもう片方に写すページで、一緒に動くことを目で見る。
  - 縦にも横にもスクロールしないと見えないシート（図・グラフ・結合・はみ出し・ダークのテーマを含む）を、Chrome・Edge・Firefox の印刷のプレビューで見て、抜けが無いこと。
  - 版 1.2（M7-21）：印刷の画面は操作できないので、`examples/print.html` の自動の動き（URL の `auto=render`）で描き、Chrome・Edge は `--headless --print-to-pdf` と DevTools の `Page.printToPDF`（`preferCSSPageSize`）、Firefox は WebDriver BiDi の `browsingContext.print` で PDF にした。7 つの場合（A4 縦・横、幅に合わせる、見出し・目盛り線、ダークの画面、図・グラフの多いシート）で、3 つのブラウザとも PDF のページの数が描いたページの数と一致し、中身が切れずに出ることを見た。人が印刷の画面で見るのは InterX-Viewer の側の確認に任せる。

### 5. 進め方と版

| Issue | GitHub | 内容 | 依存 | 完了の条件 |
| --- | --- | --- | --- | --- |
| M7-19 | #53 | `scrollElement`（`SheetView`・`WorkbookView`・`view2`）・テスト・README 3 つ | — | テストが通る。並べた 2 つの表が一緒に動く |
| M7-20 | #54 | `renderForPrint`（切り分け・待つ・中断・印刷の CSS・テーマ）・テスト・README 3 つ | — | テストが通る。3 つのブラウザのプレビューで抜けが無い |
| M7-21 | — | 実機の確認・CHANGELOG・v0.7.0（版を出すのはユーザーの指示で） | M7-19・M7-20 | 依頼の完了の条件を満たす |
| （次の版） | 新しく作る | §3 の印刷範囲 | M7-21 | — |

- 順は M7-19 → M7-20 → M7-21。M7-19 は小さいので先に出す。
- 版は 0.7.0（`SheetView` に項目が増える）。InterX-Viewer の peer の範囲 `<0.7.0` は InterX-Viewer の側で上げてもらう。

## 却下した案

| 案 | 却下した理由 |
| --- | --- |
| 依頼の案：位置をセルで読む・移す・知らせる（`scrollPosition`・`setScrollPosition`・`onScroll`） | 依頼元の `scroll` は縦の割合で、セルの位置からは割合を出せない（表の全体の高さが要る）。作っても依頼元は使えない。セルで合わせたい使う側が出たら、今の `scrollAnchor`（`zoom.ts`）で足せる |
| XSheet が割合の口（`get`・`set`・`onChange`）を持つ | 依頼元の `createScrollSync` と同じ処理（割合・`'user'` と `'api'` の区別・1 フレームにまとめる）を二重に持つことになる |
| 表を 1 枚の大きな表として仮想化せずに描き、改ページはブラウザに任せる | 幅のはみ出しは次のページに送られず切れる。大きな SVG と絶対位置の図の切れ方がブラウザで違う |
| 今の表示の仮想化を一時的に止めて、そのまま印刷する | 今の表示の倍率・位置を動かすので、印刷の後に戻す必要がある（依頼の条件）。表示の DOM が大きくなる |
| 画像（canvas）や PDF を作って印刷する | 依存が増える。文字が画像になり、印刷の品質が落ちる |
| XSheet が `@page` を入れる | 使う側のページ全体（ほかの形式の印刷）に効く |

## 結果

- 良い点：R16 は XSheet の処理がほぼ増えない。R15 は今の表示に触らないので、印刷の後に元へ戻す処理が要らない。ページの切れ方がブラウザに依らない。
- 注意：印刷の見た目は画面と違う（目盛り線・見出しが既定で出ない、テーマはライト、倍率 100%）。
- 注意：`scrollElement` を出すと、「1 つの表示に 1 つのスクロール枠」が約束になる。変えるときは破壊的な変更になる。
- 注意：印刷範囲・改ページを読まないので、0.7.0 の印刷は Excel のページ分けとは一致しない。
