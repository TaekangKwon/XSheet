import { describe, expect, it } from 'vitest';
import { createGeometry, openXlsx, type CellKind, type CellView } from '../../src/core/index';
import {
  borderXml,
  buildXlsx,
  fontXml,
  patternFillXml,
  type SheetSpec,
} from '../../src/testing/make-xlsx';
import { emptyLayout, type CellFormat, type SheetLayout } from '../../src/core/types';
import {
  applyMarginFormats,
  cellAlignClass,
  cellInnerHeight,
  clipRowToColumns,
  colSpacerWidths,
  collectVisibleRows,
  estimateNumberWidth,
  fitNumberText,
  gridExtent,
  nextVisibleRow,
  prevVisibleRow,
  rowNumWidth,
  spacerHeights,
  spanHeight,
  spillColumns,
  spillOverFill,
  spillWidth,
  visibleColRange,
  visibleRowRange,
  type GridRow,
} from '../../src/render/grid';

/**
 * `.xlsx` グリッドの純ロジック（ISSUE-84・ADR-0026 §決定 4）の単体テスト。
 * DOM を使わない部分（描画範囲・上下の余白・セル内側の高さ・はみ出しの span）だけを検証する。
 */

/** 既定のみの layout（既定行高 20px・既定列幅 64px）に行・列の指定を足す */
function layoutWith(patch: Partial<SheetLayout>): SheetLayout {
  return { ...emptyLayout(), ...patch };
}

/** セル 1 件（結合なし・既定は文字列） */
function cell(text: string, colSpan = 1, rowSpan = 1, kind: CellKind = 's'): CellView {
  return { text, kind, colSpan, rowSpan };
}

describe('rowNumWidth（行番号の列の幅）', () => {
  it('TD-84-1: 桁数に応じて広がる（最小 32px）', () => {
    expect(rowNumWidth(9)).toBe(32);
    expect(rowNumWidth(100)).toBe(40);
    expect(rowNumWidth(10000)).toBe(56);
    expect(rowNumWidth(1048576)).toBe(72);
  });

  it('TD-84-1: 0 や不正な値でも最小幅', () => {
    expect(rowNumWidth(0)).toBe(32);
    expect(rowNumWidth(Number.NaN)).toBe(32);
  });
});

describe('visibleRowRange（描く行の範囲）', () => {
  it('TD-84-2: 既定の行高（20px）でスクロール位置から見える行と前後の余分を出す', () => {
    const geo = createGeometry(layoutWith({}), 1000, 3);
    // scrollTop 400 → 行 21（20px × 20 行の手前）、高さ 200 → 行 31 まで見える
    const r = visibleRowRange(geo, 400, 200, 1000, 5);
    expect(r.top).toBe(21);
    expect(r.bottom).toBe(31);
    expect(r.first).toBe(16);
    expect(r.last).toBe(36);
  });

  it('TD-84-2: 先頭・末尾では範囲を行数の中に丸める', () => {
    const geo = createGeometry(layoutWith({}), 30, 3);
    const head = visibleRowRange(geo, 0, 100, 30, 40);
    expect(head.first).toBe(1);
    expect(head.last).toBe(30);
    const tail = visibleRowRange(geo, 99999, 100, 30, 5);
    expect(tail.top).toBe(30);
    expect(tail.last).toBe(30);
  });

  it('TD-84-2: 高さが可変（非表示・高い行）でも行番号を rowAt で引く', () => {
    const geo = createGeometry(
      layoutWith({
        rows: [
          { min: 1, max: 2, ht: null, hidden: true, custom: false },
          { min: 5, max: 5, ht: 60, hidden: false, custom: true },
        ],
      }),
      100,
      3,
    );
    // 行 1・2 は高さ 0。y=0 は表示行の先頭（行 3）
    expect(visibleRowRange(geo, 0, 10, 100, 0).top).toBe(3);
    // 行 3・4 = 40px のあと、行 5 は 80px（60pt）。y=45 は行 5 の中
    expect(visibleRowRange(geo, 45, 10, 100, 0).top).toBe(5);
  });

  it('TD-84-2: 高さ 0（jsdom 等）でも 1 行以上を出す', () => {
    const geo = createGeometry(layoutWith({}), 500, 3);
    const r = visibleRowRange(geo, 0, 0, 500, 40);
    expect(r.first).toBe(1);
    expect(r.last).toBe(41);
  });
});

/** 行 hiddenRows（1 始まりの配列）以外を非表示にした geometry。100 行 */
function filteredGeometry(visible: number[], rows = 100) {
  const hidden: SheetLayout['rows'][number][] = [];
  let prev = 1;
  for (const v of [...visible, rows + 1]) {
    if (v - 1 >= prev)
      hidden.push({ min: prev, max: v - 1, ht: null, hidden: true, custom: false });
    prev = v + 1;
  }
  return createGeometry(layoutWith({ rows: hidden }), rows, 3);
}

describe('nextVisibleRow・prevVisibleRow（見える行へ跳ぶ）', () => {
  const geo = filteredGeometry([2, 50, 51, 99]);

  it('TD-84-14: 非表示の行が続いても次の見える行へ（rowAt の二分探索）', () => {
    expect(nextVisibleRow(geo, 2, 100)).toBe(50);
    expect(nextVisibleRow(geo, 50, 100)).toBe(51);
    expect(nextVisibleRow(geo, 51, 100)).toBe(99);
    expect(nextVisibleRow(geo, 99, 100)).toBe(0); // 末尾の非表示だけが続く
    expect(nextVisibleRow(geo, 1, 100)).toBe(2); // 非表示の行からも
  });

  it('TD-84-14: 前の見える行へ', () => {
    expect(prevVisibleRow(geo, 99)).toBe(51);
    expect(prevVisibleRow(geo, 50)).toBe(2);
    expect(prevVisibleRow(geo, 2)).toBe(0); // 先頭の非表示だけが続く
  });
});

describe('visibleRowRange（見える行の数で前後の余分を数える）', () => {
  it('TD-84-14: 非表示の行が大半でも、前後の余分は見える行 overscan 行ぶん', () => {
    // 500 行に 1 行だけ見える（1・501・1001・…）10,000 行
    const visible: number[] = [];
    for (let r = 1; r <= 10000; r += 500) visible.push(r);
    const geo = filteredGeometry(visible, 10000);
    const y = geo.rowTop(2501); // 行 2501 が一番上
    const r = visibleRowRange(geo, y, 10, 10000, 3);
    expect(r.top).toBe(2501);
    expect(r.bottom).toBe(2501); // 10px 以内に次の見える行は無い
    expect(r.first).toBe(1001); // 見える行を 3 つ戻る（2001・1501・1001）
    expect(r.last).toBe(4001);
  });

  it('TD-84-14: 見える行が 1 つも無ければ 0', () => {
    const geo = filteredGeometry([], 10);
    expect(visibleRowRange(geo, 0, 100, 10, 5)).toEqual({ top: 0, bottom: 0, first: 0, last: 0 });
  });

  it('TD-84-14: 末尾だけ非表示でも、見える最後の行まで', () => {
    const geo = filteredGeometry([1, 2, 3], 10);
    const r = visibleRowRange(geo, 99999, 100, 10, 5);
    expect(r.top).toBe(3);
    expect(r.last).toBe(3);
  });
});

describe('spacerHeights（仮想化の上下の余白）', () => {
  it('TD-84-3: 上＝rowTop(first)、下＝totalHeight − rowTop(last + 1)', () => {
    const geo = createGeometry(layoutWith({}), 1000, 3);
    const s = spacerHeights(geo, 101, 140);
    expect(s.top).toBe(100 * 20);
    expect(s.bottom).toBe(860 * 20);
    expect(s.top + 40 * 20 + s.bottom).toBe(geo.totalHeight);
  });

  it('TD-84-3: 非表示の行・高さの違う行を含んでも、上＋描く行＋下＝総高さ', () => {
    const geo = createGeometry(
      layoutWith({
        rows: [
          { min: 3, max: 4, ht: null, hidden: true, custom: false },
          { min: 10, max: 10, ht: 30, hidden: false, custom: true },
        ],
      }),
      50,
      3,
    );
    const first = 6;
    const last = 20;
    const s = spacerHeights(geo, first, last);
    let drawn = 0;
    for (let r = first; r <= last; r++) drawn += geo.rowHeight(r);
    expect(s.top + drawn + s.bottom).toBe(geo.totalHeight);
  });

  it('TD-84-3: 全行を描くなら余白は 0', () => {
    const geo = createGeometry(layoutWith({}), 10, 3);
    expect(spacerHeights(geo, 1, 10)).toEqual({ top: 0, bottom: 0 });
  });
});

describe('cellInnerHeight・spanHeight（セルの内側の要素の高さ）', () => {
  it('TD-84-4: セルの高さから下罫線 1px を引く（1px のセルは 0・負にならない）', () => {
    expect(cellInnerHeight(20)).toBe(19);
    expect(cellInnerHeight(1)).toBe(0);
    expect(cellInnerHeight(0)).toBe(0);
  });

  it('TD-84-4: 結合の高さは、描く見える行 span 行ぶんの行高の合計', () => {
    const geo = createGeometry(
      layoutWith({ rows: [{ min: 2, max: 2, ht: 30, hidden: false, custom: true }] }),
      10,
      3,
    );
    const rows = [{ r: 1 }, { r: 2 }, { r: 5 }, { r: 6 }];
    expect(spanHeight(geo, rows, 0, 1)).toBe(20);
    expect(spanHeight(geo, rows, 0, 3)).toBe(20 + 40 + 20);
    expect(spanHeight(geo, rows, 2, 9)).toBe(40); // 行 5・6（描く行の終わりで止まる）
    expect(spanHeight(geo, rows, 1, 0)).toBe(40); // span 0 でも自分の行
  });
});

describe('spillColumns（文字列がはみ出して使える列数）', () => {
  it('TD-84-6: 右隣が空なら、値のあるセルの手前までの列数を返す（自分の列を含む）', () => {
    const cells = [cell('長い文字列'), undefined, undefined, cell('値'), undefined];
    expect(spillColumns(cells, 0)).toBe(3);
    expect(spillColumns(cells, 3)).toBe(2);
  });

  it('TD-84-6: 右隣に値があれば広げない（1）', () => {
    expect(spillColumns([cell('a'), cell('b')], 0)).toBe(1);
  });

  it('TD-84-6: 型の無い空文字のセルは空として扱う（数式の結果が "" の文字列セルは値あり＝次の TD）', () => {
    expect(spillColumns([cell('a'), { text: '', colSpan: 1, rowSpan: 1 }, undefined], 0)).toBe(3);
  });

  it('TD-84-18: 数式の結果が "" の文字列セル（kind s・text ""）は値のあるセルとして、はみ出しをそこで止める', () => {
    expect(spillColumns([cell('a'), undefined, cell(''), undefined], 0)).toBe(2);
    expect(spillColumns([cell('a'), cell('')], 0)).toBe(1);
  });

  it('TD-84-6: 結合で隠れるセル（null）の手前で止まる', () => {
    expect(spillColumns([cell('a'), undefined, null, undefined], 0)).toBe(2);
  });

  it('TD-84-6: 右隣の結合セル（span あり）は値のあるセルと同じく手前で止まる', () => {
    expect(spillColumns([cell('a'), undefined, cell('', 2, 1), null], 0)).toBe(2);
    expect(spillColumns([cell('a'), cell('', 1, 2)], 0)).toBe(1);
  });

  it('TD-84-18: はみ出すのは文字列（kind s）だけ。数値・真偽・エラー・日付・型の無いセルは自分のセルの幅で切る', () => {
    const empties = [undefined, undefined];
    for (const kind of ['n', 'b', 'e', 'd'] as const) {
      expect(spillColumns([cell('1,234', 1, 1, kind), ...empties], 0), kind).toBe(1);
    }
    expect(spillColumns([{ text: 'x', colSpan: 1, rowSpan: 1 }, ...empties], 0)).toBe(1);
    expect(spillColumns([cell('x'), ...empties], 0)).toBe(3);
  });

  it('TD-84-18: 数字だけの文字列・「▲1,234」・日付に見える文字列は、文字列なのではみ出す（表示文字からは推定しない）', () => {
    for (const text of ['12345678901234', '▲1,234', '2024/1/15', '00123']) {
      expect(spillColumns([cell(text), undefined, undefined], 0), text).toBe(3);
    }
    // 数値セルは、文字が文字列のように見えても切る
    expect(spillColumns([cell('abc', 1, 1, 'n'), undefined], 0)).toBe(1);
  });

  it('TD-84-6: 結合セル・空のセルははみ出さない（1）', () => {
    expect(spillColumns([cell('a', 2, 1), null, undefined], 0)).toBe(1);
    expect(spillColumns([cell('a', 1, 2), undefined], 0)).toBe(1);
    expect(spillColumns([undefined, undefined], 0)).toBe(1);
    expect(spillColumns([cell(''), undefined], 0)).toBe(1);
  });

  it('TD-84-6: 結合セル（merged）・表の列数より右の空の列（cols 指定）も扱う', () => {
    expect(spillColumns([{ text: 'a', colSpan: 1, rowSpan: 1, merged: true }, undefined], 0)).toBe(
      1,
    );
    // セルの配列が短くても、表の列数（足した空の列）まで空として数える
    expect(spillColumns([cell('a')], 0, 4)).toBe(4);
    expect(
      spillColumns([cell('a'), { text: '', colSpan: 1, rowSpan: 1, merged: true }], 0, 5),
    ).toBe(1);
  });

  it('TD-84-6: 列数は上限（1000）を超えない', () => {
    const cells: (CellView | null | undefined)[] = [cell('a')];
    for (let i = 0; i < 1500; i++) cells.push(undefined);
    expect(spillColumns(cells, 0)).toBe(1000);
  });
});

describe('spillWidth（はみ出す文字の入れ物の幅）', () => {
  /** 列 1〜3 は 64px、列 4 は非表示、列 5 は 100px、列 6 以降は 64px */
  const geo = createGeometry(
    layoutWith({
      cols: [
        { min: 4, max: 4, width: null, hidden: true },
        { min: 5, max: 5, width: 13.7, hidden: false },
      ],
    }),
    10,
    8,
  );

  it('TD-84-13: 自分の列から n 列ぶんの列幅の合計から、右端の罫線 1px を引く', () => {
    expect(geo.colWidth(1)).toBe(64);
    expect(spillWidth(geo, 0, 1)).toBe(63);
    expect(spillWidth(geo, 0, 3)).toBe(64 * 3 - 1);
    expect(spillWidth(geo, 1, 2)).toBe(64 * 2 - 1);
  });

  it('TD-84-13: 非表示の列（幅 0）は合計に入らない', () => {
    // 列 3（0 始まり 2）から 列 3・4（非表示）・5 の 3 列
    expect(spillWidth(geo, 2, 3)).toBe(64 + 0 + geo.colWidth(5) - 1);
  });

  it('TD-84-13: 自分の列が非表示なら 0（隠れたセルの文字を出さない）', () => {
    expect(spillWidth(geo, 3, 2)).toBe(0);
  });
});

/** 試験用の xlsx を開いて 1 枚目のシートを返す */
async function openSheet(spec: SheetSpec) {
  const bytes = await buildXlsx({ sheets: [spec], font: { name: 'Calibri', size: 11 } });
  const buf = bytes.buffer.slice(
    bytes.byteOffset,
    bytes.byteOffset + bytes.byteLength,
  ) as ArrayBuffer;
  const book = await openXlsx(buf);
  return book.openSheet(1);
}

/** collectVisibleRows の呼び出し（実際のシート）。span を { 行: [列ごとの rowSpan / null / 文字] } にして返す */
function summarize(rows: { r: number; cells: (CellView | null | undefined)[] }[]) {
  return rows.map((row) => ({
    r: row.r,
    cells: row.cells.map((c) =>
      c === null ? null : c === undefined ? undefined : `${c.text}:${c.rowSpan}`,
    ),
  }));
}

describe('collectVisibleRows（非表示の行を除き、結合を見える行へ寄せる）', () => {
  /** 縦の結合 A2:A5・列 B に値。行 3・4 を非表示 */
  const rowsAB: (string | null)[][] = [];
  for (let r = 1; r <= 8; r++) rowsAB.push([r === 2 ? 'M' : null, `b${r}`]);

  it('TD-84-15: 非表示の行は返さず、読むのは見える行だけ（隙間が大きいと別々に読む）', async () => {
    const sh = await openSheet({ name: 'S', rows: rowsAB });
    const geo = createGeometry(
      layoutWith({ rows: [{ min: 3, max: 6, ht: null, hidden: true, custom: false }] }),
      8,
      2,
    );
    const calls: [number, number][] = [];
    const rows = collectVisibleRows(
      geo,
      1,
      8,
      8,
      (a, b) => {
        calls.push([a, b]);
        return sh.getRows(a, b);
      },
      0,
    );
    expect(rows.map((r) => r.r)).toEqual([1, 2, 7, 8]);
    expect(calls).toEqual([
      [1, 2],
      [7, 8],
    ]);
  });

  it('TD-84-15: 隙間が小さければまとめて 1 回で読み、非表示の行を跨ぐ結合の rowSpan は見える行の数になる', async () => {
    const sh = await openSheet({ name: 'S', rows: rowsAB, merges: ['A2:A5'] });
    const geo = createGeometry(
      layoutWith({ rows: [{ min: 3, max: 4, ht: null, hidden: true, custom: false }] }),
      8,
      2,
    );
    const calls: [number, number][] = [];
    const rows = collectVisibleRows(geo, 1, 8, 8, (a, b) => {
      calls.push([a, b]);
      return sh.getRows(a, b);
    });
    expect(calls).toEqual([[1, 8]]);
    // 行 2 の A は行 2・5 の 2 行ぶん（行 3・4 は非表示）。行 5 の A は覆われる（null）
    expect(summarize(rows).map((r) => [r.r, r.cells[0]])).toEqual([
      [1, undefined],
      [2, 'M:2'],
      [5, null],
      [6, undefined],
      [7, undefined],
      [8, undefined],
    ]);
  });

  it('TD-84-16: 結合の先頭の行が非表示なら、結合内の最初の見える行へ移す（値と rowSpan を保つ）', async () => {
    const sh = await openSheet({ name: 'S', rows: rowsAB, merges: ['A2:A5'] });
    const geo = createGeometry(
      layoutWith({ rows: [{ min: 2, max: 3, ht: null, hidden: true, custom: false }] }),
      8,
      2,
    );
    const rows = collectVisibleRows(geo, 1, 8, 8, (a, b) => sh.getRows(a, b));
    // 結合 A2:A5 のうち見える行は 4・5。行 4 に値、rowSpan 2
    expect(summarize(rows).map((r) => [r.r, r.cells[0]])).toEqual([
      [1, undefined],
      [4, 'M:2'],
      [5, null],
      [6, undefined],
      [7, undefined],
      [8, undefined],
    ]);
    expect(rows[1]!.cells[0]).toMatchObject({ merged: true });
  });

  it('TD-84-16: 結合の全域が非表示ならセルごと消える', async () => {
    const sh = await openSheet({ name: 'S', rows: rowsAB, merges: ['A2:A3'] });
    const geo = createGeometry(
      layoutWith({ rows: [{ min: 2, max: 3, ht: null, hidden: true, custom: false }] }),
      8,
      2,
    );
    const rows = collectVisibleRows(geo, 1, 8, 8, (a, b) => sh.getRows(a, b));
    expect(rows.map((r) => r.r)).toEqual([1, 4, 5, 6, 7, 8]);
    expect(summarize(rows).every((r) => r.cells[0] !== 'M:1')).toBe(true);
  });

  it('TD-84-16: 見える行どうしの結合は rowSpan そのまま・結合の印（merged）が付く', async () => {
    const sh = await openSheet({ name: 'S', rows: rowsAB, merges: ['A2:A3'] });
    const geo = createGeometry(layoutWith({}), 8, 2);
    const rows = collectVisibleRows(geo, 1, 8, 8, (a, b) => sh.getRows(a, b));
    expect(summarize(rows)[1]!.cells[0]).toBe('M:2');
    expect(rows[1]!.cells[0]).toMatchObject({ merged: true });
  });

  it('TD-84-15: 描く範囲が空・見える行が無ければ空', () => {
    const geo = createGeometry(layoutWith({}), 8, 2);
    expect(collectVisibleRows(geo, 0, 0, 8, () => [])).toEqual([]);
    const hidden = filteredGeometry([], 8);
    expect(collectVisibleRows(hidden, 1, 8, 8, () => [])).toEqual([]);
  });
});

describe('gridExtent（表をスクロール枠いっぱいに広げる）', () => {
  it('TD-84-17: データが小さければ、枠の幅・高さを埋める列数・行数まで広げる', () => {
    // 幅 800px・行番号の列 32px → 768px。64px 列は 12 列で 768px（x=768 は 13 列目）
    const e = gridExtent(emptyLayout(), 3, 2, 800, 400);
    expect(e.cols).toBe(13);
    // 高さ 400px・20px 行 → y=400 は 21 行目
    expect(e.rows).toBe(21);
  });

  it('TD-84-17: データが枠より大きければデータのまま', () => {
    expect(gridExtent(emptyLayout(), 500, 40, 800, 400)).toEqual({ rows: 500, cols: 40 });
  });

  it('TD-84-17: 枠の大きさ 0（jsdom 等）ならデータのまま（最小 1×1）', () => {
    expect(gridExtent(emptyLayout(), 3, 2, 0, 0)).toEqual({ rows: 3, cols: 2 });
    expect(gridExtent(emptyLayout(), 0, 0, 0, 0)).toEqual({ rows: 1, cols: 1 });
  });

  it('TD-84-17: 最小の行数・列数を外から渡せる（画像のある範囲まで広げる用）', () => {
    expect(gridExtent(emptyLayout(), 3, 2, 0, 0, 50, 30)).toEqual({ rows: 50, cols: 30 });
  });

  it('TD-84-17: 上限は Excel の最大（1,048,576 行・16,384 列）', () => {
    expect(gridExtent(emptyLayout(), 1, 1, 0, 0, 9999999, 99999)).toEqual({
      rows: 1048576,
      cols: 16384,
    });
  });

  it('TD-84-17: 足す列の幅は geometry の既定（指定の列は指定の幅）。非表示の列は数えない', () => {
    const layout = layoutWith({
      cols: [
        { min: 1, max: 1, width: null, hidden: true },
        { min: 2, max: 2, width: 100, hidden: false },
      ],
    });
    const e = gridExtent(layout, 1, 1, 400, 100);
    const geo = createGeometry(layout, e.rows, e.cols);
    expect(geo.totalWidth + 32).toBeGreaterThanOrEqual(400);
    expect(createGeometry(layout, e.rows, e.cols - 1).totalWidth + 32).toBeLessThan(400);
  });
});

describe('estimateNumberWidth・fitNumberText（収まらない数値は ###）', () => {
  it('TD-84-20: 推定幅は 文字数 × MDW。小数点・桁区切り・コロン・スラッシュ・空白は MDW の半分', () => {
    expect(estimateNumberWidth('2024/1/15', 7)).toBe(7 * 7 + 2 * 3.5);
    expect(estimateNumberWidth('12345', 7)).toBe(35);
    expect(estimateNumberWidth('1,234.5', 7)).toBe(5 * 7 + 2 * 3.5);
    expect(estimateNumberWidth('-12.5%', 8)).toBe(5 * 8 + 4);
    expect(estimateNumberWidth('', 7)).toBe(0);
    // MDW が不正なら 7 とみなす
    expect(estimateNumberWidth('12', Number.NaN)).toBe(14);
  });

  it('TD-84-21: 収まる数値・日付は表示文字のまま（内側の幅＝セル幅 − 罫線 1px − 左右の余白 6px）', () => {
    // 64px 列：内側 57px。8 桁（56px）は収まる
    expect(fitNumberText('12345678', 'n', 64, 7)).toBe('12345678');
    expect(fitNumberText('2024/1/15', 'd', 64 + 20, 7)).toBe('2024/1/15');
    expect(fitNumberText('1,234.50', 'n', 64, 7)).toBe('1,234.50');
  });

  it('TD-84-21: 収まらない数値は # の繰り返し（内側の幅 ÷ MDW の切り捨て個数）', () => {
    // 64px 列：内側 57px → 8 個
    expect(fitNumberText('123456789', 'n', 64, 7)).toBe('########');
    expect(fitNumberText('1.23457E+13', 'n', 64, 7)).toBe('########');
    // 広い列：内側 140−7＝133px → 19 個
    expect(fitNumberText('1'.repeat(30), 'n', 140, 7)).toBe('#'.repeat(19));
  });

  it('TD-84-21: 幅が足りない列でも # は最低 1 個', () => {
    expect(fitNumberText('12345', 'n', 6, 7)).toBe('#');
    expect(fitNumberText('1', 'n', 3, 7)).toBe('#');
  });

  it('TD-84-21: 文字列・真偽・エラー・型なしは対象外（そのまま）', () => {
    for (const kind of ['s', 'b', 'e', undefined] as const) {
      expect(fitNumberText('123456789012345', kind, 64, 7), String(kind)).toBe('123456789012345');
    }
  });
});

describe('cellAlignClass（型ごとの既定の配置）', () => {
  it('TD-84-22: 数値・日付は右、真偽・エラーは中央、文字列・型なしは左（クラス無し）', () => {
    expect(cellAlignClass('n')).toBe('is-right');
    expect(cellAlignClass('d')).toBe('is-right');
    expect(cellAlignClass('b')).toBe('is-center');
    expect(cellAlignClass('e')).toBe('is-center');
    expect(cellAlignClass('s')).toBe('');
    expect(cellAlignClass(undefined)).toBe('');
  });
});

describe('列の仮想化（visibleColRange・colSpacerWidths・clipRowToColumns）', () => {
  /** 既定の列幅 64px の 16,384 列 */
  const wide = createGeometry(layoutWith({}), 100, 16384);

  it('TD-86-17: 描く列は可視範囲の左右 1 画面分（px）。遠い列でも描く列数は画面幅に比例する', () => {
    // 行番号の列 40px・枠の幅 1280px・scrollLeft 64,040 → A1 の x = 64,000（列 1,001 の左端）
    const r = visibleColRange(wide, 64040, 1280, 40, 16384);
    expect(r.left).toBe(1001);
    expect(r.right).toBe(1021); // x + 1280 = 65,280 → 列 1,021（65,280 は列 1,021 の左端）
    expect(r.first).toBe(1001 - 20);
    expect(r.last).toBe(1001 + 40);
    expect(r.last - r.first).toBeLessThan(70);
  });

  it('TD-86-17: 先頭では first が 1・末尾では last が列数に丸まる。列数 0 は全部 0', () => {
    expect(visibleColRange(wide, 0, 1280, 40, 16384).first).toBe(1);
    const geo = createGeometry(layoutWith({}), 10, 30);
    expect(visibleColRange(geo, 99999, 800, 40, 30).last).toBe(30);
    expect(visibleColRange(geo, 0, 800, 40, 0)).toEqual({ left: 0, right: 0, first: 0, last: 0 });
  });

  it('TD-86-17: 枠の幅が 0（レイアウト前）なら代わりの幅 1280px を使う', () => {
    const r = visibleColRange(wide, 0, 0, 40, 16384);
    expect(r.right).toBe(wide.colAt(1280 - 40));
    expect(r.last).toBe(wide.colAt(2 * 1280 - 40));
  });

  it('TD-86-17: 左右の余白の幅＋描く列の幅＝表の全幅（非表示の列を含んでも）', () => {
    const geo = createGeometry(
      layoutWith({ cols: [{ min: 5, max: 6, width: null, hidden: true }] }),
      10,
      3000,
    );
    const { left, right } = colSpacerWidths(geo, 100, 200);
    expect(left).toBe(geo.colLeft(100));
    expect(left + (geo.colLeft(201) - geo.colLeft(100)) + right).toBe(geo.totalWidth);
    expect(colSpacerWidths(geo, 1, 3000)).toEqual({ left: 0, right: 0 });
  });

  it('TD-86-17: 範囲だけを切り出し、cells[i] が列 first + i になる（範囲外・列数より右は undefined）', () => {
    const row = { r: 1, cells: [cell('a'), cell('b'), cell('c'), cell('d')] };
    expect(clipRowToColumns(row, 2, 3).cells.map((c) => c?.text)).toEqual(['b', 'c']);
    expect(clipRowToColumns(row, 4, 6).cells.map((c) => c?.text)).toEqual([
      'd',
      undefined,
      undefined,
    ]);
  });

  it('TD-86-17: 結合の起点が範囲の左外なら、起点を範囲の左端へ移し colSpan を範囲内へ縮める（結合のまま）', () => {
    // 列 2〜6 の結合（起点は列 2）。範囲は列 4〜8
    const row = {
      r: 1,
      cells: [cell('x'), cell('結合', 5, 1), null, null, null, null, cell('y'), undefined],
    };
    const clipped = clipRowToColumns(row, 4, 8);
    expect(clipped.cells[0]).toMatchObject({ text: '結合', colSpan: 3, rowSpan: 1, merged: true });
    expect(clipped.cells.slice(1, 3)).toEqual([null, null]);
    expect(clipped.cells[3]?.text).toBe('y');
  });

  it('TD-86-17: 起点が上の行にある（rowSpan で覆われる）null は、そのまま null（td を作らない）', () => {
    const row = { r: 2, cells: [cell('x'), null, null, cell('z')] };
    // 範囲 3〜4：左外の列 2 も null で、辿っても起点の無い行 → 補わない
    const clipped = clipRowToColumns(row, 3, 4);
    expect(clipped.cells[0]).toBeNull();
    expect(clipped.cells[1]?.text).toBe('z');
  });

  it('TD-86-17: 右端で切れる結合は範囲の右端までに縮める（結合のまま＝はみ出さない）', () => {
    const row = { r: 1, cells: [cell('a'), cell('結合', 5, 1), null, null, null, null] };
    const clipped = clipRowToColumns(row, 1, 3);
    expect(clipped.cells[1]).toMatchObject({ colSpan: 2, merged: true });
    // 範囲内に収まる結合・結合でないセルは触らない
    expect(clipRowToColumns(row, 1, 6).cells[1]).toMatchObject({ colSpan: 5 });
    expect(clipRowToColumns(row, 1, 6).cells[1]?.merged).toBeUndefined();
  });

  it('TD-86-17: はみ出しは描く範囲の右端で切る（右の列を空とみなして範囲の外へ広げない）', () => {
    const clipped = clipRowToColumns({ r: 1, cells: [cell('長い文字列')] }, 1, 3);
    expect(spillColumns(clipped.cells, 0, 3)).toBe(3);
  });
});

describe('collectVisibleRows の罫線の番号（ISSUE-98・GridRow.borders）', () => {
  it('TD-98-1: SheetRow.borders をそのまま持ち、非表示の行は罫線ごと返さない・罫線の無い行は borders を持たない', async () => {
    const bytes = await buildXlsx({
      sheets: [
        {
          name: 'S',
          rows: [
            [{ v: 'a', t: 'inlineStr', s: 1 }],
            [{ v: 'b', t: 'inlineStr', s: 1 }],
            [{ v: 'c', t: 'inlineStr' }],
          ],
        },
      ],
      xfs: [0, 0],
      borders: [
        borderXml({ left: {} }),
        borderXml({ bottom: { style: 'thin', color: 'auto="1"' } }),
      ],
      xfBorderIds: [0, 1],
    });
    const sh = await (await openXlsx(bytes)).openSheet(1);
    const geo = createGeometry(
      layoutWith({ rows: [{ min: 2, max: 2, ht: null, hidden: true, custom: false }] }),
      3,
      1,
    );
    const rows = collectVisibleRows(geo, 1, 3, 3, (a, b) => sh.getRows(a, b));
    expect(rows.map((r) => r.r)).toEqual([1, 3]);
    expect(rows[0]!.borders).toHaveLength(1);
    expect(sh.borders[rows[0]!.borders![0]!]?.bottom?.style).toBe('thin');
    expect(rows[1]!.borders).toBeUndefined();
  });
});

describe('fitNumberText の文字サイズの比（ISSUE-100・ADR-0030 §決定 10）', () => {
  // 列幅 64px（MDW 7）。内側の幅＝64−1−6＝57。「1,234」は 5 文字（桁区切りは半分）＝推定 31.5px
  it('TD-100-n: 比 1（既定）は今と同じ。比 2 で推定幅が倍（63px＞57px）になり ### に切り替わる', () => {
    expect(fitNumberText('1,234', 'n', 64, 7)).toBe('1,234');
    expect(fitNumberText('1,234', 'n', 64, 7, 1)).toBe('1,234');
    expect(fitNumberText('1,234', 'n', 64, 7, 2)).toBe('#'.repeat(4));
  });

  it('TD-100-n: 比 0.5 は小さい文字のため、今は ### になる幅でも収まる。文字列は比に関わらずそのまま', () => {
    expect(fitNumberText('1234567890', 'n', 64, 7)).toBe('#'.repeat(8));
    expect(fitNumberText('1234567890', 'n', 64, 7, 0.5)).toBe('1234567890');
    expect(fitNumberText('long text here', 's', 64, 7, 4)).toBe('long text here');
  });
});

describe('fitNumberText の # の個数（文字サイズの比で割る・ISSUE-100・ADR-0030 §決定 10）', () => {
  // 列幅 64px（MDW 7）。内側の幅 57px＝# 8 個ぶん（57÷7＝8.14）。大きい文字の # は比だけ幅を取る
  it('TD-100-n: 比が 1 より大きいとき、# の個数も比で割る（切り捨て）。比 1・既定は 8 個のまま', () => {
    const wide = '1234567890';
    expect(fitNumberText(wide, 'n', 64, 7)).toBe('#'.repeat(8));
    expect(fitNumberText(wide, 'n', 64, 7, 1)).toBe('#'.repeat(8));
    expect(fitNumberText(wide, 'n', 64, 7, 1.5)).toBe('#'.repeat(5));
    expect(fitNumberText(wide, 'n', 64, 7, 2)).toBe('#'.repeat(4));
    expect(fitNumberText(wide, 'n', 64, 7, 4)).toBe('#'.repeat(2));
  });

  it('TD-100-n: 最低 1 個（狭い列・大きい文字でも # が 0 個にならない）', () => {
    expect(fitNumberText('12345', 'n', 6, 7, 4)).toBe('#');
    expect(fitNumberText('12345', 'n', 20, 7, 4)).toBe('#');
  });

  it('TD-100-n: 比が 1 未満（小さい文字）は # を増やさない（今の個数のまま）', () => {
    expect(fitNumberText('1'.repeat(40), 'n', 64, 7, 0.5)).toBe('#'.repeat(8));
  });
});

describe('applyMarginFormats（表の範囲の外の余白に行・列の書式を当てる・ISSUE-100）', () => {
  /** 行 r の空の行（cells は表の列数ぶん undefined） */
  const row = (r: number, formats?: (number | undefined)[]): GridRow => ({
    r,
    cells: [],
    ...(formats ? { formats } : {}),
  });
  /** 列の書式：列 2＝10、列 4＝11、列 6＝-1（書式なし）、ほかは -1 */
  const colFormat = (c: number): number => (c === 2 ? 10 : c === 4 ? 11 : -1);
  const noRowFormat = (): number | undefined => undefined;
  /** 列 1〜6 の書式の番号（範囲より短い配列・穴は undefined で揃える） */
  const head = (r: GridRow): (number | undefined)[] =>
    Array.from({ length: 6 }, (_v, i) => r.formats?.[i]);
  /** 表の範囲は 3 行 × 3 列。描く列は 1〜6 */
  const apply = (
    rows: GridRow[],
    rowFormat: (r: number) => number | undefined = noRowFormat,
    col: (c: number) => number = colFormat,
    first = 1,
    last = 6,
  ): GridRow[] => applyMarginFormats(rows, first, last, 3, 3, rowFormat, col);

  it('TD-100-n: 表の行数より下の行（余白の行）は、描く全列に行の書式があればそれを、無ければ列の書式を当てる', () => {
    const out = apply([row(4), row(5)], (r) => (r === 5 ? 7 : undefined));
    // 行 4：行の書式なし → 列の書式（列 2＝10・列 4＝11）。行 5：行の書式 7 を全列へ
    expect(head(out[0]!)).toEqual([undefined, 10, undefined, 11, undefined, undefined]);
    expect(head(out[1]!)).toEqual([7, 7, 7, 7, 7, 7]);
  });

  it('TD-100-n: 表の中の行（r ≤ 行数）は、表の列数より右の余白の列だけに当てる（表の中の列は読み取りの結果のまま）', () => {
    const src = row(2, [5, undefined, 6]);
    const out = apply([src], (r) => (r === 2 ? 7 : undefined));
    // 列 1〜3 は読み取りの結果（5・undefined・6）のまま。列 4〜6 の余白に行の書式 7
    expect(head(out[0]!)).toEqual([5, undefined, 6, 7, 7, 7]);
    const out2 = apply([row(2, [5, undefined, 6])]);
    expect(head(out2[0]!)).toEqual([5, undefined, 6, 11, undefined, undefined]);
  });

  it('TD-100-n: 行の書式が書式なし（-1）なら、列の書式も当てない（行の書式が列の書式を打ち消す）', () => {
    const src = row(5);
    const out = apply([src], () => -1);
    expect(out[0]).toBe(src);
    expect(out[0]!.formats).toBeUndefined();
  });

  it('TD-100-n: 当てるものが無い行・余白に触れない行は、同じ行のまま返す。入力は変えない', () => {
    const inside = row(1, [1, 2, 3]); // 描く列が 1〜3 だけなら余白に触れない
    const none = row(5); // 余白の行だが行の書式も列の書式も無い
    const noCol = (): number => -1;
    const rows = [inside, none];
    const out = applyMarginFormats(rows, 1, 3, 3, 3, noRowFormat, colFormat);
    expect(out[0]).toBe(inside);
    expect(out[1]!.formats?.[1]).toBe(10); // 余白の行には列の書式（列 2）が当たる
    expect(applyMarginFormats([none], 1, 6, 3, 3, noRowFormat, noCol)[0]).toBe(none);
    const withCol = row(5);
    const outCol = applyMarginFormats([withCol], 1, 6, 3, 3, noRowFormat, colFormat);
    expect(outCol[0]).not.toBe(withCol);
    expect(withCol.formats).toBeUndefined(); // 入力の行は書き換えない
    const orig = [5, undefined, 6];
    const src = row(2, orig);
    apply([src]);
    expect(orig).toEqual([5, undefined, 6]);
  });

  it('TD-100-n: 描く範囲が遠い列（first が大きい）でも、描く列だけを見る。clipRowToColumns に渡すと同じ列に揃う', () => {
    const out = applyMarginFormats([row(4)], 3, 5, 3, 3, noRowFormat, colFormat);
    const clipped = clipRowToColumns(out[0]!, 3, 5);
    // 列 3・4・5 → 書式は [なし, 11, なし]
    expect(clipped.formats).toEqual([undefined, 11, undefined]);
  });

  it('TD-100-n: rowFormat は余白に触れる行にだけ 1 回、colFormat は描く余白の列ごとに 1 回（行が多くても再計算しない）', () => {
    let rowCalls = 0;
    let colCalls = 0;
    const rf = (r: number): number | undefined => {
      rowCalls++;
      return r === 3 ? 1 : undefined;
    };
    const cf = (c: number): number => {
      colCalls++;
      return colFormat(c);
    };
    applyMarginFormats([row(1), row(2), row(3), row(4), row(5)], 1, 6, 3, 3, rf, cf);
    // 行 1〜3 は表の中だが、描く列 1〜6 が表の列数 3 を超えるため余白（列 4〜6）に触れる。5 行とも rowFormat を 1 回
    expect(rowCalls).toBe(5);
    // 列の書式が要るのは行の書式の無い 4 行（行 3 以外）。余白の列は 4〜6、余白の行は全列 1〜6 ＝ 6 列ぶんだけ（各 1 回）
    expect(colCalls).toBe(6);
    rowCalls = 0;
    applyMarginFormats([row(1), row(2)], 1, 3, 3, 3, rf, cf);
    expect(rowCalls).toBe(0); // 表の中の行・列だけ → 余白に触れない
  });

  it('TD-100-n: 書式のブック（書式なし）や空の範囲では何もしない', () => {
    const rows = [row(4)];
    expect(applyMarginFormats(rows, 1, 6, 3, 3, noRowFormat, () => -1)[0]).toBe(rows[0]);
    expect(applyMarginFormats(rows, 5, 4, 3, 3, noRowFormat, colFormat)[0]).toBe(rows[0]);
    expect(applyMarginFormats([], 1, 6, 3, 3, noRowFormat, colFormat)).toEqual([]);
  });
});

describe('spillOverFill（はみ出しの塗りの判定・ISSUE-100・ADR-0030 §決定 9）', () => {
  const red: CellFormat = { fill: { rgb: 0xff0000, alpha: 1 } };
  const white: CellFormat = { fill: { rgb: 0xffffff, alpha: 1 } };
  const bold: CellFormat = { bold: true };
  const table = [red, white, bold];

  it('TD-100-n: はみ出す元のセルに白以外の塗りがあれば true', () => {
    expect(spillOverFill(table, [0, undefined, undefined], 0, 3)).toBe(true);
  });

  it('TD-100-n: 数えた列（元の右の空セル）のどれかに白以外の塗りがあれば true。数えていない列は見ない', () => {
    expect(spillOverFill(table, [undefined, undefined, 0], 0, 3)).toBe(true);
    expect(spillOverFill(table, [undefined, undefined, 0], 0, 2)).toBe(false);
  });

  it('TD-100-n: 白の塗り・塗りなしの装飾・書式なし・formats なしは塗りとして数えない', () => {
    expect(spillOverFill(table, [1, 1, 1], 0, 3)).toBe(false);
    expect(spillOverFill(table, [2, 2, undefined], 0, 3)).toBe(false);
    expect(spillOverFill(table, [undefined, undefined], 0, 2)).toBe(false);
    expect(spillOverFill(table, undefined, 0, 3)).toBe(false);
  });

  it('TD-100-n: 範囲外の番号（表に無い）は無視する。c が 0 でない位置からも数える', () => {
    expect(spillOverFill(table, [undefined, 99, 0], 1, 2)).toBe(true);
    expect(spillOverFill(table, [undefined, 99], 1, 1)).toBe(false);
  });
});

describe('GridRow.formats（ISSUE-100）', () => {
  /** xf 1＝赤の塗り。fonts[0] は既定（sz 11） */
  async function openFormatted() {
    const bytes = await buildXlsx({
      sheets: [
        {
          name: 'S',
          rows: [
            [{ v: 'a', t: 'inlineStr', s: 1 }],
            [{ v: 'b', t: 'inlineStr', s: 1 }],
            [{ v: 'c', t: 'inlineStr' }],
          ],
        },
      ],
      fonts: [fontXml('<sz val="11"/>')],
      fills: [patternFillXml('none'), patternFillXml('solid', 'rgb="FFFF0000"')],
      xfs: [0, 0],
      xfFontIds: [0, 0],
      xfFillIds: [0, 1],
    });
    return (await openXlsx(bytes)).openSheet(1);
  }

  it('TD-100-n: collectVisibleRows が formats を写す（非表示の行は返さない・書式の無い行は持たない）', async () => {
    const sh = await openFormatted();
    const geo = createGeometry(
      layoutWith({ rows: [{ min: 2, max: 2, ht: null, hidden: true, custom: false }] }),
      3,
      1,
    );
    const rows = collectVisibleRows(geo, 1, 3, 3, (a, b) => sh.getRows(a, b));
    expect(rows.map((r) => r.r)).toEqual([1, 3]);
    expect(rows[0]!.formats).toHaveLength(1);
    expect(sh.formats[rows[0]!.formats![0]!]?.fill?.rgb).toBe(0xff0000);
    expect(rows[1]!.formats).toBeUndefined();
  });

  it('結合の起点の行が非表示で、セルを最初に見える行へ移すとき、起点の書式も移す（移動先の自セルの書式は残さない）', async () => {
    // A2:A5 の結合。左上（A2）だけが赤の塗り。行 4（結合内の最初の見える行）の A4 は別の塗り（xf 2）を持つ
    const bytes = await buildXlsx({
      sheets: [
        {
          name: 'S',
          rows: [[null], [{ v: 'M', t: 'inlineStr', s: 1 }], [null], [{ s: 2 }], [null]],
          merges: ['A2:A5'],
        },
      ],
      fonts: [fontXml('<sz val="11"/>')],
      fills: [
        patternFillXml('none'),
        patternFillXml('solid', 'rgb="FFFF0000"'),
        patternFillXml('solid', 'rgb="FF0000FF"'),
      ],
      xfs: [0, 0, 0],
      xfFontIds: [0, 0, 0],
      xfFillIds: [0, 1, 2],
    });
    const sh = await (await openXlsx(bytes)).openSheet(1);
    const geo = createGeometry(
      layoutWith({ rows: [{ min: 2, max: 3, ht: null, hidden: true, custom: false }] }),
      5,
      1,
    );
    const rows = collectVisibleRows(geo, 1, 5, 5, (a, b) => sh.getRows(a, b));
    expect(rows.map((r) => r.r)).toEqual([1, 4, 5]);
    expect(rows[1]!.cells[0]).toMatchObject({ text: 'M', rowSpan: 2, merged: true });
    expect(sh.formats[rows[1]!.formats![0]!]?.fill?.rgb).toBe(0xff0000);
    // 起点に書式が無いときは、移動先の自セルの書式も残さない
    const plain = await buildXlsx({
      sheets: [
        {
          name: 'S',
          rows: [[null], [{ v: 'M', t: 'inlineStr' }], [null], [{ s: 1 }], [null]],
          merges: ['A2:A5'],
        },
      ],
      fonts: [fontXml('<sz val="11"/>')],
      fills: [patternFillXml('none'), patternFillXml('solid', 'rgb="FFFF0000"')],
      xfs: [0, 0],
      xfFontIds: [0, 0],
      xfFillIds: [0, 1],
    });
    const sh2 = await (await openXlsx(plain)).openSheet(1);
    const rows2 = collectVisibleRows(geo, 1, 5, 5, (a, b) => sh2.getRows(a, b));
    expect(rows2[1]!.cells[0]).toMatchObject({ text: 'M', merged: true });
    expect(rows2[1]!.formats?.[0]).toBeUndefined();
  });

  it('TD-100-n: clipRowToColumns が formats も描く列だけにする（列 first＋i の書式が formats[i]）', () => {
    const row = {
      r: 1,
      cells: [cell('a'), cell('b'), cell('c'), cell('d')],
      formats: [0, 1, undefined, 2],
    };
    const clipped = clipRowToColumns(row, 2, 4);
    expect(clipped.formats).toEqual([1, undefined, 2]);
    // formats の無い行は formats を持たない
    expect(clipRowToColumns({ r: 1, cells: [cell('a'), cell('b')] }, 1, 2).formats).toBeUndefined();
  });

  it('TD-100-n: 範囲の左外から続く結合は、起点のセルの書式を範囲の左端に持つ（結合は左上のセルの書式）', () => {
    const row = {
      r: 1,
      cells: [cell('m', 3), null, null, cell('z')],
      formats: [5, 6, 6, undefined],
    };
    const clipped = clipRowToColumns(row, 2, 4);
    expect(clipped.cells[0]).toMatchObject({ text: 'm', colSpan: 2 });
    expect(clipped.formats).toEqual([5, 6, undefined]);
  });
});
