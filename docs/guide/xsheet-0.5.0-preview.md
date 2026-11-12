# 0.5.0 の型の先出し

| 項目 | 値 |
| --- | --- |
| 文書 | XSheet 使い方の手引き |
| 更新日 | 2026-11-12 |
| 関連 | [xsheet-0.5.0-preview.d.ts](xsheet-0.5.0-preview.d.ts)（型）・[ADR-0007](../adr/0007-moobius-viewer-api.md)（決定）・[standalone.md](standalone.md)（単独版） |

## TL;DR
- 0.5.0 で足す公開 API（検索・ハイライト・移動・位置の判定・拡大縮小・見出しを隠す・`showSheet`・和暦の指定）の型を、実装より先に [xsheet-0.5.0-preview.d.ts](xsheet-0.5.0-preview.d.ts) にまとめた。単独版 0.4.1 の `xsheet.d.ts` と同じフォルダに置くと、画面側のコードを 0.5.0 の形で書いて型検査できる。
- 依頼の案と違うのは 3 点：検索は**非同期で、中断と件数の上限を持つ**。シートを替える `showSheet` を足した。組み込みの日付の書式の和暦は**指定したときだけ**。
- 実装はまだ無い（0.4.1 で呼ぶと `undefined`）。0.5.0 を出したら、その版の `xsheet.d.ts` に置き換えてこのファイルは消す。

---

## 1. 置き方

単独版 0.4.1 の zip の `xsheet.d.ts` と同じフォルダに `xsheet-0.5.0-preview.d.ts` を置き、TypeScript に両方を読ませる（`tsconfig.json` の `include`・`files` か、`/// <reference path>`）。大域の `XSheet` と `import type { … } from 'xsheet'` の両方に 0.5.0 の追加が出る。

| 足す所 | 足すもの |
| --- | --- |
| `XSheet`（関数） | `findText` |
| 型 | `FindTextOptions`・`FindTextResult`・`TextMatch`・`Zoom`・`FocusHighlightOptions`・`ScrollToCellOptions`・`CellPosition`・`BuiltinDates` |
| `SheetView`・`WorkbookView` | `setHighlights`・`focusHighlight`・`scrollToCell`・`cellAt`・`highlightAt`・`setZoom`・`zoom`・`showSheet`（`WorkbookView` は前から持つ） |
| `RenderOptions`・`WorkbookOptions` | `zoom`・`headers` |
| `OpenXlsxOptions`・`ReadOptions`・`WorkbookOptions` | `builtinDates` |

このファイルは単独版の `xsheet.d.ts` の中の名前（`xsheet/_/…`）に重ねている。バンドラで `npm install` した XSheet の型には重ならない。

## 2. 依頼の案との違い

| 項目 | 依頼の案 | 0.5.0 | 理由 |
| --- | --- | --- | --- |
| 検索 R1 | 同期で配列を返す | `Promise<{ matches, truncated }>`。`signal` で中断（`AbortError`）、`maxMatches`（既定 10,000・上限 100,000）で打ち切る | 大きなシートやありふれた文字で画面が固まり、メモリも際限なく使うため。約 10ms ごとに休む |
| シートの切り替え | 無し | `SheetView.showSheet(index)` を足した。倍率は保ち、ハイライトは消す | 表示を作り直すと倍率とハイライトが毎回消えるため |
| 和暦 R6 | — | 書式に直接書いた和暦（`ggg`・`e`）と曜日はいつも出す。組み込みの書式 27〜36・50〜58 は `builtinDates: 'japanese-era'` のときだけ和暦 | これらの番号の見た目は開く Excel の言語環境で決まり、韓国語・中国語のファイルも同じ番号を使うため |
| 拡大縮小 R5 | — | 見出し（列見出し・行番号）には倍率を掛けない | 見出しの寸法が写したファイルの定数に結び付いているため。Excel と見た目が違う |

受け入れの確認から外す項目（右寄せのはみ出しのハイライト・折り返しをまたぐハイライトなど）は [ADR-0007](../adr/0007-moobius-viewer-api.md) §8 にある。XSheet がまだその表示を描かないため。

## 3. 使い方の例

`renderWorkbook` はブックを自分で開き、開いたブックを外に出さない。検索するシートは、同じバイト列を `openXlsx` で開いて取る。

```ts
const bytes = await file.arrayBuffer();
const view = await XSheet.renderWorkbook(el, bytes, { zoom: { fit: 'width' }, headers: false });
const book = await XSheet.openXlsx(bytes); // 検索用（閉じるのは呼ぶ側）

const ac = new AbortController();
const sheet = await book.openSheet(view.sheets[0].index);
const { matches, truncated } = await XSheet.findText(sheet, '見積', { signal: ac.signal });
view.setHighlights(matches);
view.focusHighlight(0); // 最初のヒットへスクロールして強調する

el.addEventListener('click', (e) => {
  const hit = view.highlightAt(e.clientX, e.clientY);
  if (hit >= 0) view.focusHighlight(hit, { scroll: false });
});
```

## 4. ADR-0007 に無く、型で決めた所

| 所 | 決めたこと |
| --- | --- |
| `focusHighlight` の `scroll`・`smooth` | 既定は `scroll: true`・`smooth: false` |
| `scrollToCell` の `smooth` | 既定 `false` |
| `RenderOptions.zoom` | 既定 1 |
| `focusHighlight` の番号 | `setHighlights` に渡した配列の番号（0 始まり） |
| `WorkbookOptions.builtinDates` | `renderWorkbook` もブックを自分で開くので足した（ADR-0007 §6 にも足した） |

## 5. まだ決まっていないこと

- `view` の戻り値 `Viewer`（と `ViewOptions` の `zoom`・`headers`）に同じ操作を足すか。決まるまでは、0.5.0 の操作を使うなら `renderWorkbook` を使う。
- 組み込みの書式 27〜36・50〜58 を和暦にしたときの書式（日本語版 Excel で確かめてから作る）。元年を「元」と出すか。
