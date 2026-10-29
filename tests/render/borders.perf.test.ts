// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import type { BorderEdge, BorderStyle, CellBorder } from '../../src/core/index';
import { createBorderLayer } from '../../src/render/border-layer';
import { buildBorderPaths } from '../../src/render/borders';
import type { GridRow } from '../../src/render/grid';

/**
 * 罫線の重さの計測（ISSUE-98・ADR-0029 §決定 9）：全セルに罫線を引いた 10,000 セルの合成シートで、
 * 層の構築（辺の計算＋ DOM）の時間を測る。table 自身の作り直し（td 10,000 個）はこの層とは別に毎回かかるため、
 * 比較用に同じ規模の td の生成時間も測って出力する。
 * 許容は「再構築 1 回が 200ms（スクロールの 1 フレームの許容）を十分下回る」こと。jsdom は実ブラウザより遅いので余裕を見る。
 */

const ROWS = 100;
const COLS = 100;
const geo = {
  colWidth: () => 64,
  colLeft: (c: number) => Math.max(0, c - 1) * 64,
  rowTop: (r: number) => Math.max(0, r - 1) * 20,
};

const edge = (style: BorderStyle, rgb: number | null = null): BorderEdge => ({
  style,
  color: rgb === null ? null : { rgb, alpha: 1 },
});

/** thin・medium・dashed・double（と赤）を混ぜた 4 辺の罫線を、全セルに引く（行・列の添字で種類を変える） */
function synthetic(): { rows: GridRow[]; table: CellBorder[] } {
  const table: CellBorder[] = [];
  const styles: BorderStyle[] = ['thin', 'medium', 'dashed', 'double', 'dotted', 'thin'];
  for (let k = 0; k < styles.length; k++) {
    const e = edge(styles[k]!, k === 5 ? 0xff0000 : null);
    table.push({ left: e, right: e, top: e, bottom: e, diagonalUp: false, diagonalDown: false });
  }
  const rows: GridRow[] = [];
  for (let r = 1; r <= ROWS; r++) {
    const borders = Array.from(
      { length: COLS },
      (_, c) => (r * 7 + c * 3 + ((r * c) % 5)) % table.length,
    );
    rows.push({ r, cells: new Array(COLS).fill(undefined), borders });
  }
  return { rows, table };
}

describe('TD-98-13 重さ：10,000 セル（全セル罫線あり）', () => {
  it('層の構築（辺の計算＋ SVG）が 1 回 200ms を十分下回る。区間の数は上限内', () => {
    const { rows, table } = synthetic();
    const layer = createBorderLayer();
    // ウォームアップ（JIT）
    layer.render(buildBorderPaths(rows, table, geo, { first: 1, last: COLS }), 32, 24);

    const runs = 5;
    const times: number[] = [];
    let segments = 0;
    for (let i = 0; i < runs; i++) {
      const t0 = performance.now();
      const result = buildBorderPaths(rows, table, geo, { first: 1, last: COLS });
      const t1 = performance.now();
      layer.render(result, 32, 24);
      const t2 = performance.now();
      times.push(t2 - t0);
      segments = result.segments;
      if (i === runs - 1) {
        // 比較用：同じ規模の td 10,000 個の生成時間
        const t3 = performance.now();
        const tbody = document.createElement('tbody');
        for (let r = 0; r < ROWS; r++) {
          const tr = document.createElement('tr');
          for (let c = 0; c < COLS; c++) tr.append(document.createElement('td'));
          tbody.append(tr);
        }
        const t4 = performance.now();
        console.log(
          `[TD-98-13] 10,000 セル：辺の計算 ${(t1 - t0).toFixed(1)}ms ＋ SVG ${(t2 - t1).toFixed(1)}ms` +
            `（区間 ${result.segments}・path ${result.paths.length}）／ 比較：td 10,000 個の生成 ${(t4 - t3).toFixed(1)}ms`,
        );
      }
    }
    const worst = Math.max(...times);
    console.log(
      `[TD-98-13] 層の構築 ${runs} 回：${times.map((t) => t.toFixed(1)).join(' / ')} ms（最大 ${worst.toFixed(1)}ms）`,
    );
    expect(worst).toBeLessThan(100);
    expect(segments).toBeGreaterThan(0);
    expect(segments).toBeLessThanOrEqual(50_000);
    // 線種×色ごとにまとまる（path は十数本以内）
    expect(layer.element.querySelectorAll('path').length).toBeLessThan(40);
  });
});
