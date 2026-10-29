/// <reference types="node" />
import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { createGeometry, openXlsx } from '../../src/core/index';
import { layoutDrawings } from '../../src/render/drawings';
import { MAX_GRID_COLS, MAX_GRID_ROWS } from '../../src/render/grid';
import { pathPoints } from '../../src/render/shapes';

/**
 * Excel 16 の COM で作った `ref-geoms.xlsx`（ISSUE-91）の全図形が、代替にならず形で描かれること（ISSUE-92・TD-92-11）。
 * 表に無い `prst`・評価の失敗・上限超えが 1 つも無いことの確認（形の読み取り → 配置 → 経路の一連）。
 */
describe('TD-92-11 ref-geoms.xlsx：全図形が形で描かれる', () => {
  it('代替の図形は 0。枠も塗りも文字も無い 2 図形（emptyDiamond・lineNone）だけが描かれず、残り 223 図形は形で描かれる', async () => {
    const book = await openXlsx(
      readFileSync(new URL('../../fixtures/excel/ref-geoms.xlsx', import.meta.url)),
    );
    const { drawings, notices } = await book.openDrawings(1);
    const sheet = await book.openSheet(1);
    const geo = createGeometry(sheet.layout, MAX_GRID_ROWS, MAX_GRID_COLS);
    expect(notices).toEqual([]);
    expect(drawings).toHaveLength(225);

    const boxes = layoutDrawings(geo, drawings);
    const notDrawn = boxes
      .filter((b) => b.shape?.kind !== 'drawn')
      .map((b) => drawings[b.index]!.name);
    expect(notDrawn).toEqual([]);
    // Excel でも見えない図形（枠線なし・塗りなし・文字なし）は描かない（ADR-0028 P4）
    const hidden = drawings.filter((_, i) => !boxes.some((b) => b.index === i)).map((d) => d.name);
    expect(hidden).toEqual(['emptyDiamond', 'lineNone']);
    expect(boxes).toHaveLength(223);

    for (const b of boxes) {
      const shape = b.shape!;
      expect(shape.faces.length, drawings[b.index]!.name).toBeGreaterThan(0);
      for (const f of shape.faces) {
        expect(f.d).toMatch(/^[MLCZ0-9 .-]+$/);
        for (const sp of f.subpaths) {
          for (const p of pathPoints(sp)) {
            expect(Number.isFinite(p.x) && Number.isFinite(p.y)).toBe(true);
          }
        }
      }
      expect(b.rect.x).toBeGreaterThanOrEqual(0);
      expect(b.rect.w).toBeGreaterThan(0);
      expect(b.rect.h).toBeGreaterThan(0);
    }
  });
});
