/// <reference types="node" />
/**
 * 実際の Excel で作った基準ファイルによる、寸法（layout）→ px（geometry）の照合。
 * `fixtures/excel/ref-*.xlsx` は、Excel 16 を画面 125%（120dpi）の PC で操作して保存した実ファイル
 * （ISSUE-83 の U3・U4 の照合に使い、中身の図・図形・線・グラフは ISSUE-85 でも使う）。
 * 期待値は 96dpi 換算の px。列 B=20、C=幅 0.88 付近、D=非表示、E=9 付近、G・H=3.7 付近、
 * 行 3=30pt、行 5=非表示、行 7=0.75pt、行 20〜25=40.05pt を指定して保存してある。
 */
import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { createGeometry, type Geometry } from '../../src/core/geometry';
import { openXlsx, type XlsxSheet } from '../../src/core/index';

/** 基準ファイルを開いて、先頭シートと geometry を返す */
async function openRef(name: string): Promise<{ sheet: XlsxSheet; g: Geometry }> {
  const bytes = readFileSync(new URL(`../../fixtures/excel/${name}.xlsx`, import.meta.url));
  const sheet = await (await openXlsx(bytes)).openSheet(1);
  return { sheet, g: createGeometry(sheet.layout, sheet.rowCount, sheet.colCount) };
}

/** 列 A〜H の幅 */
const colWidths = (g: Geometry) => [1, 2, 3, 4, 5, 6, 7, 8].map((c) => g.colWidth(c));

describe('TD-83-9 Excel 実ファイル：Calibri 11（MDW 7）', () => {
  it('layout：MDW 7・フォントの既定の行高 20px・defaultRowHeight 14.4pt（customHeight 無し）・既定の列幅の指定なし', async () => {
    const { sheet } = await openRef('ref-calibri');
    expect(sheet.layout).toMatchObject({
      mdw: 7,
      fontRowPx: 20,
      defaultRowHeight: 14.4,
      customHeight: false,
      defaultColWidth: null,
      zeroHeight: false,
    });
    expect(sheet.rowCount).toBe(10000);
  });

  it('列幅：既定 64、B=145、C=6、D（非表示）=0、E=65、G・H=26', async () => {
    const { g } = await openRef('ref-calibri');
    expect(colWidths(g)).toEqual([64, 145, 6, 0, 65, 64, 26, 26]);
    expect(g.colLeft(9)).toBe(64 + 145 + 6 + 0 + 65 + 64 + 26 + 26);
  });

  it('行高：既定＝20（customHeight が無いので defaultRowHeight 14.4pt＝19 ではなくフォントの既定）、行 3（30pt）＝40、行 5（非表示）＝0、行 7（0.75pt）＝1、行 20〜25（40.05pt）＝53', async () => {
    const { g } = await openRef('ref-calibri');
    expect([1, 2, 3, 4, 5, 6, 7, 8].map((r) => g.rowHeight(r))).toEqual([
      20, 20, 40, 20, 0, 20, 1, 20,
    ]);
    expect([19, 20, 21, 22, 23, 24, 25, 26].map((r) => g.rowHeight(r))).toEqual([
      20, 53, 53, 53, 53, 53, 53, 20,
    ]);
  });

  it('1〜10,000 行目の rowTop が、既定の行高×個数＋差分の累積と一致する', async () => {
    const { g } = await openRef('ref-calibri');
    const heightOf = (r: number): number =>
      r === 3 ? 40 : r === 5 ? 0 : r === 7 ? 1 : r >= 20 && r <= 25 ? 53 : 20;
    let top = 0;
    for (let r = 1; r <= 10000; r++) {
      expect(g.rowTop(r)).toBe(top);
      top += heightOf(r);
    }
    // 検算：9,999 行 × 20 ＋ 行 3（+20）・行 5（-20）・行 7（-19）・行 20〜25（+33 × 6）
    expect(g.rowTop(10000)).toBe(9999 * 20 + 20 - 20 - 19 + 33 * 6);
    expect(g.totalHeight).toBe(top);
  });

  it('rowAt／colAt が rowTop／colLeft と整合する（非表示の行・列は飛ばす）', async () => {
    const { g } = await openRef('ref-calibri');
    expect(g.rowAt(g.rowTop(5))).toBe(6); // 行 5 は非表示
    expect(g.rowAt(g.rowTop(7))).toBe(7);
    expect(g.rowAt(g.rowTop(10000))).toBe(10000);
    expect(g.colAt(g.colLeft(4))).toBe(5); // 列 D は非表示
  });
});

describe('TD-83-10 Excel 実ファイル：游ゴシック・맑은 고딕（MDW 8）', () => {
  it('游ゴシック：MDW 8・既定の列 72・既定の行 25（フォントの既定。defaultRowHeight 18pt＝24 ではない）。列幅は MDW 8 の式で換算される', async () => {
    const { sheet, g } = await openRef('ref-yugothic');
    expect(sheet.layout).toMatchObject({ mdw: 8, fontRowPx: 25, defaultRowHeight: 18 });
    expect([g.colWidth(1), g.rowHeight(1)]).toEqual([72, 25]);
    expect(colWidths(g)).toEqual([72, 166, 7, 0, 73, 72, 30, 30]);
    expect([3, 5, 7, 20].map((r) => g.rowHeight(r))).toEqual([40, 0, 1, 53]);
  });

  it('맑은 고딕：MDW 8（韓国語版 Excel の既定 72px）・既定の行 22（フォントの既定。17.4pt＝23 ではない）', async () => {
    const { sheet, g } = await openRef('ref-malgun');
    expect(sheet.layout).toMatchObject({ mdw: 8, fontRowPx: 22, defaultRowHeight: 17.4 });
    expect([g.colWidth(1), g.rowHeight(1)]).toEqual([72, 22]);
    expect(colWidths(g)).toEqual([72, 166, 7, 0, 73, 72, 30, 30]);
  });
});

describe('TD-83-11 Excel 実ファイル：defaultColWidth（標準の幅 12）', () => {
  it('defaultColWidth＝12.77734375 は <col width> と同じ式で換算され、列 A も B も同じ px（89）', async () => {
    const { sheet, g } = await openRef('ref-calibri-stdw12');
    expect(sheet.layout).toMatchObject({ mdw: 7, defaultColWidth: 12.77734375 });
    expect(g.colWidth(1)).toBe(89);
    expect(g.colWidth(2)).toBe(89);
    expect(g.colWidth(50)).toBe(89);
  });
});
