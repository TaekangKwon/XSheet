/// <reference types="node" />
import { readFileSync } from 'node:fs';
import { beforeAll, describe, expect, it } from 'vitest';
import { createGeometry, openXlsx, type SheetDrawing } from '../../src/core/index';
import { layoutDrawings, placeDrawing } from '../../src/render/drawings';

/**
 * 実際の Excel で作った基準ファイル `ref-shapes.xlsx` による、グループの子の位置の照合（ISSUE-89・ADR-0027 §決定 5）。
 * COM（`Shapes.GroupItems`）の Left・Top・Width・Height（pt。シートの左上基準）を px（96/72 倍）にして、
 * グループのアンカー（geometry から出した px）に子の座標系から写した矩形と 1px 以内で比べる。
 * Node 環境（jsdom は `new URL(…, import.meta.url)` を `fs` に渡せないため、DOM テストの `group.test.ts` と分けた）。
 */

describe('TD-89-8 ref-shapes.xlsx：グループの子の矩形が COM の値と 1px 以内で一致する', () => {
  let byName: Map<string, SheetDrawing>;
  let g: ReturnType<typeof createGeometry>;

  beforeAll(async () => {
    const book = await openXlsx(
      readFileSync(new URL('../../fixtures/excel/ref-shapes.xlsx', import.meta.url)),
    );
    byName = new Map((await book.openDrawings(1)).drawings.map((d) => [d.name, d]));
    const sheet = await book.openSheet(1);
    g = createGeometry(sheet.layout, 1_048_576, 16_384);
  });

  /** COM の pt（シートの左上基準）→ px（96/72 倍） */
  const px4 = (pt: number): number => (pt * 96) / 72;

  /** 子 1 件の COM の矩形（pt：Left, Top, Width, Height） */
  type Com = [number, number, number, number];

  /** グループ自身の矩形（px）：子の `anchor`（＝最も外側のグループのアンカー）だけを geometry から写したもの */
  function groupRect(name: string) {
    const d = byName.get(name);
    if (!d) throw new Error(`図が無い: ${name}`);
    return placeDrawing(g, { ...d, group: undefined })!;
  }

  /**
   * 名前の図の矩形（px）が COM と 1px 以内で一致する。
   * - 横（x・幅）：列幅は Excel の値のまま換算しているので、COM の pt × 96/72 と絶対値で比べる。
   * - 縦（y・高さ）：アプリの行高は Excel 16 の COM（画面 125%・既定 14.4pt）の行より高い（既定 20px）ため、シートの絶対位置は合わない
   *   （アンカーはセル基準で、図形は行に付いて動く）。グループの矩形に対する COM の比率で比べる。
   * @param group グループの COM の矩形（pt）。入れ子の子は最も外側のグループ
   */
  function expectCom(name: string, groupName: string, group: Com, com: Com): void {
    const d = byName.get(name);
    if (!d) throw new Error(`図が無い: ${name}`);
    const r = placeDrawing(g, d)!;
    const gr = groupRect(groupName);
    expect(r, name).not.toBeNull();
    expect(Math.abs(r.x - px4(com[0])), `${name} x ${r.x} vs ${px4(com[0])}`).toBeLessThanOrEqual(
      1,
    );
    expect(Math.abs(r.w - px4(com[2])), `${name} w ${r.w} vs ${px4(com[2])}`).toBeLessThanOrEqual(
      1,
    );
    const wantY = gr.y + ((com[1] - group[1]) / group[3]) * gr.h;
    const wantH = (com[3] / group[3]) * gr.h;
    expect(Math.abs(r.y - wantY), `${name} y ${r.y} vs ${wantY}`).toBeLessThanOrEqual(1);
    expect(Math.abs(r.h - wantH), `${name} h ${r.h} vs ${wantH}`).toBeLessThanOrEqual(1);
  }

  it('grpScaled（幅を 1.5 倍に拡大したグループ）：テキストボックス gTb・四角形 gRect・接続線 gConn', () => {
    const grp: Com = [20, 1605, 270, 50];
    expectCom('gTb', 'gTb', grp, [20, 1605, 90, 20]);
    expectCom('gRect', 'gTb', grp, [200, 1635, 90, 20]);
    expectCom('gConn', 'gTb', grp, [110, 1615, 90, 30]);
  });

  it('grpPlain（拡大なし）：hRect・hOval・hLine（高さ 0 の線）', () => {
    const grp: Com = [320, 1605, 140, 60];
    expectCom('hRect', 'hRect', grp, [320, 1605, 60, 30]);
    expectCom('hOval', 'hRect', grp, [400, 1625, 60, 30]);
    expectCom('hLine', 'hRect', grp, [320, 1665, 140, 0]);
  });

  it('grpOuter（入れ子）：grpInner の中の nRect1・nRect2、外側の直下の nRect3', () => {
    const grp: Com = [720, 1605, 110, 80];
    expectCom('nRect1', 'nRect1', grp, [720, 1605, 50, 30]);
    expectCom('nRect2', 'nRect1', grp, [780, 1605, 50, 30]);
    expectCom('nRect3', 'nRect1', grp, [720, 1655, 110, 30]);
  });

  it('grpRot（グループ自身を 30° 回転）：回転は無視して、子は回転前の位置・大きさに置く（COM の子は回転後の位置なので、座標系の値で照合する）', () => {
    // グループは 140×50pt。子 iRect(60×30pt)・iRect2(60×30pt) は、グループの左上と、そこから右へ 80pt・下へ 20pt の位置
    const grp: Com = [520, 1605, 140, 50];
    expectCom('iRect', 'iRect', grp, [520, 1605, 60, 30]);
    expectCom('iRect2', 'iRect', grp, [600, 1625, 60, 30]);
  });

  it('layoutDrawings：子 11 件すべてが配置される（線 gConn・hLine は外接矩形の SVG になる）', () => {
    const drawings = [...byName.values()];
    const boxes = layoutDrawings(g, drawings);
    const names = [
      'gTb',
      'gRect',
      'gConn',
      'hRect',
      'hOval',
      'hLine',
      'iRect',
      'iRect2',
      'nRect1',
      'nRect2',
      'nRect3',
    ];
    for (const n of names) {
      const i = drawings.findIndex((d) => d.name === n);
      const box = boxes.find((b) => b.index === i);
      expect(box, n).toBeDefined();
      if (n === 'gConn' || n === 'hLine') expect(box!.line, n).toBeDefined();
    }
  });
});
