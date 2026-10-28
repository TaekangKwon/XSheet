/// <reference types="node" />
/**
 * 実際の Excel で作った基準ファイルによる、図（drawing）の読み取りの照合（ISSUE-85）。
 * `fixtures/excel/ref-drawings.xlsx` は Excel 16 で作成・保存した実ファイル（すべて twoCellAnchor。Excel は oneCell／absolute を editAs で表す）。
 * 中身：png1（pic・editAs=oneCell・image1.png 40x30）・jpg1（pic・editAs=absolute・image2.jpg 64x48）・
 * svg1（a:blip r:embed＝image3.png の PNG 予備＋asvg:svgBlip r:embed＝image4.svg）・png2（pic・editAs=oneCell）・
 * group1（grpSp。中に図形 grpRectA・grpOvalB と画像 grpPic。ISSUE-89 から子 3 件に展開され、group1 自身は一覧に出ない）・text1（テキストボックス）・conn1（矢印付きの線 cxnSp。ISSUE-87 から種類 line で一覧に出る）・
 * pngRot（pic・a:srcRect l="16667"・rot="1800000"）。
 * `ref-calibri.xlsx` などは ISSUE-83 の寸法の照合で使った実ファイルで、drawing に pic1（画像）・rect1（図形 sp）・line1（線 cxnSp）・chart1（グラフ）が入っている。
 * 期待値は drawing XML の値（列・行は 0 始まり、ずれは EMU）。
 */
import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { openXlsx, type XlsxBook } from '../../src/core/index';

/** 基準ファイルを開く */
async function openRef(name: string): Promise<XlsxBook> {
  return openXlsx(readFileSync(new URL(`../../fixtures/excel/${name}.xlsx`, import.meta.url)));
}

describe('TD-85-12 Excel 実ファイル：ref-drawings.xlsx の図の一覧', () => {
  it('10 件（画像 6・図形 2・テキストボックス 1・線 1）。group1 は子 3 件（図形 2・画像 1）に展開され、自身は出ない。線（conn1）は line で出る。順序は drawing XML の順', async () => {
    const { drawings, notices } = await (await openRef('ref-drawings')).openDrawings(1);
    expect(notices).toEqual([]);
    expect(drawings.map((d) => [d.kind, d.name])).toEqual([
      ['image', 'png1'],
      ['image', 'jpg1'],
      ['image', 'svg1'],
      ['image', 'png2'],
      ['shape', 'grpRectA'],
      ['shape', 'grpOvalB'],
      ['image', 'grpPic'],
      ['textbox', 'text1'],
      ['line', 'conn1'],
      ['image', 'pngRot'],
    ]);
    expect(drawings.filter((d) => d.kind === 'image')).toHaveLength(6);
    expect(drawings.some((d) => d.kind === 'group')).toBe(false);
    // 線は図（画像など）の数に入らない
    expect(drawings.filter((d) => d.kind !== 'line')).toHaveLength(9);
  });

  it('TD-89-13 group1 の子：座標系（chOff・chExt が off・ext と同じ）と、子の xfrm（EMU）が drawing XML のとおり', async () => {
    const { drawings } = await (await openRef('ref-drawings')).openDrawings(1);
    const kids = ['grpRectA', 'grpOvalB', 'grpPic'].map((n) => drawings.find((d) => d.name === n)!);
    const levels = [{ ch: { x: 609600, y: 2011680, cx: 1727200, cy: 651510 }, own: null }];
    expect(kids.map((d) => d.group)).toEqual([
      { levels, xfrm: { x: 609600, y: 2011680, cx: 508000, cy: 254000 } },
      { levels, xfrm: { x: 1828800, y: 2011680, cx: 508000, cy: 254000 } },
      { levels, xfrm: { x: 609600, y: 2377440, cx: 381000, cy: 285750 } },
    ]);
    // 子のアンカーは group1 のアンカー（B12〜D15）
    for (const d of kids) {
      expect(d.anchor).toEqual({
        type: 'twoCell',
        from: { col: 1, colOff: 0, row: 11, rowOff: 0 },
        to: { col: 3, colOff: 508000, row: 14, rowOff: 102870 },
      });
    }
    // グループの外の図は group を持たない
    expect(drawings.filter((d) => d.group === undefined)).toHaveLength(7);
  });

  it('TD-89-13 group1 の中の画像 grpPic は、グループの外と同じ readImage で読める（PNG 40x30）', async () => {
    const book = await openRef('ref-drawings');
    const { drawings } = await book.openDrawings(1);
    const pic = drawings.find((d) => d.name === 'grpPic')!;
    expect(pic.kind).toBe('image');
    expect(pic.mediaPath).not.toBeNull();
    expect(await book.readImage(1, pic.mediaPath!)).toMatchObject({ ok: true, format: 'png' });
  });

  it('アンカー：すべて twoCell。from・to の列・行・ずれ（EMU）が drawing XML のとおり', async () => {
    const { drawings } = await (await openRef('ref-drawings')).openDrawings(1);
    const pt = (col: number, colOff: number, row: number, rowOff: number) => ({
      col,
      colOff,
      row,
      rowOff,
    });
    const group1 = { type: 'twoCell', from: pt(1, 0, 11, 0), to: pt(3, 508000, 14, 102870) };
    expect(drawings.map((d) => d.anchor)).toEqual([
      { type: 'twoCell', from: pt(1, 0, 1, 0), to: pt(1, 381000, 2, 102870) },
      { type: 'twoCell', from: pt(4, 76200, 1, 38100), to: pt(5, 76200, 3, 129540) },
      { type: 'twoCell', from: pt(1, 0, 5, 0), to: pt(1, 476250, 7, 110490) },
      { type: 'twoCell', from: pt(4, 0, 5, 0), to: pt(4, 381000, 6, 102870) },
      group1,
      group1,
      group1,
      { type: 'twoCell', from: pt(1, 0, 17, 0), to: pt(2, 406400, 18, 121920) },
      { type: 'twoCell', from: pt(1, 0, 21, 0), to: pt(3, 0, 23, 0) },
      { type: 'twoCell', from: pt(1, 59246, 25, 15875), to: pt(1, 376746, 26, 118745) },
    ]);
  });

  it('media のパス：svg1 は SVG（image4.svg）でなく PNG の予備（image3.png）。画像以外は null。回転・トリミングの pngRot も画像として出る', async () => {
    const { drawings } = await (await openRef('ref-drawings')).openDrawings(1);
    expect(drawings.map((d) => d.mediaPath)).toEqual([
      'xl/media/image1.png',
      'xl/media/image2.jpg',
      'xl/media/image3.png',
      'xl/media/image1.png',
      null,
      null,
      'xl/media/image1.png',
      null,
      null,
      'xl/media/image1.png',
    ]);
  });

  it('readImage：PNG 40x30・JPEG 64x48（先頭バイトで判定）・SVG の PNG 予備', async () => {
    const book = await openRef('ref-drawings');
    const { drawings } = await book.openDrawings(1);
    const read = async (name: string) => {
      const d = drawings.find((x) => x.name === name)!;
      const r = await book.readImage(1, d.mediaPath!);
      expect(r.ok).toBe(true);
      return r.ok ? { format: r.format, mime: r.mime, width: r.width, height: r.height } : null;
    };
    expect(await read('png1')).toEqual({ format: 'png', mime: 'image/png', width: 40, height: 30 });
    expect(await read('png2')).toEqual({ format: 'png', mime: 'image/png', width: 40, height: 30 });
    expect(await read('pngRot')).toEqual({
      format: 'png',
      mime: 'image/png',
      width: 40,
      height: 30,
    });
    expect(await read('jpg1')).toEqual({
      format: 'jpeg',
      mime: 'image/jpeg',
      width: 64,
      height: 48,
    });
    expect(await read('svg1')).toMatchObject({ format: 'png', width: 100, height: 100 });
  });

  it('SVG 本体（image4.svg）は、パスを直接渡しても描けない画像として返す', async () => {
    const r = await (await openRef('ref-drawings')).readImage(1, 'xl/media/image4.svg');
    expect(r).toEqual({ ok: false, reason: 'unsupported', label: null });
  });
});

describe('TD-85-13 Excel 実ファイル：ref-calibri・ref-yugothic・ref-malgun の図（画像・図形・線・グラフ）', () => {
  it('ref-calibri：一覧は 4 件（画像 pic1・図形 rect1・線 line1・グラフ chart1）。線 line1 は line で出る。アンカーは drawing XML のとおり', async () => {
    const { drawings, notices } = await (await openRef('ref-calibri')).openDrawings(1);
    expect(notices).toEqual([]);
    expect(drawings).toEqual([
      {
        kind: 'image',
        name: 'pic1',
        descr: '',
        mediaPath: 'xl/media/image1.png',
        anchor: {
          type: 'twoCell',
          from: { col: 2, colOff: 38100, row: 3, rowOff: 25400 },
          to: { col: 4, colOff: 358140, row: 5, rowOff: 128270 },
        },
      },
      {
        kind: 'shape',
        name: 'rect1',
        descr: '',
        mediaPath: null,
        rot: 0,
        flipH: false,
        flipV: false,
        hasOutline: true,
        hasFill: true,
        outline: { widthEmu: 12700, dash: 'solid' },
        // Excel 既定の図形：塗り accent1（#4472C4）・枠線 accent1＋shade（#172C51。ADR-0028 §決定 6.3 の例）
        fillColor: { rgb: 0x4472c4, alpha: 1 },
        outlineColor: { rgb: 0x172c51, alpha: 1 },
        geom: { type: 'preset', prst: 'rect', adj: [] },
        anchor: {
          type: 'twoCell',
          from: { col: 1, colOff: 0, row: 7, rowOff: 0 },
          to: { col: 1, colOff: 762000, row: 9, rowOff: 15240 },
        },
      },
      {
        kind: 'line',
        name: 'line1',
        descr: '',
        mediaPath: null,
        anchor: {
          type: 'twoCell',
          from: { col: 0, colOff: 0, row: 11, rowOff: 0 },
          to: { col: 2, colOff: 0, row: 13, rowOff: 0 },
        },
        rot: 0,
        flipH: false,
        flipV: false,
        line: {
          geom: 'line',
          adj: [50000, 50000, 50000],
          widthEmu: 6350,
          dash: 'solid',
          head: null,
          tail: null,
          color: { rgb: 0x4472c4, alpha: 1 },
        },
      },
      {
        kind: 'chart',
        name: 'chart1',
        descr: '',
        mediaPath: null,
        anchor: {
          type: 'twoCell',
          from: { col: 5, colOff: 0, row: 14, rowOff: 0 },
          to: { col: 9, colOff: 167640, row: 19, rowOff: 228600 },
        },
      },
    ]);
  });

  it('游ゴシック・맑은 고딕の基準ファイルも同じ構成（画像・図形・線・グラフの 4 件）で、画像は PNG 40x30', async () => {
    for (const name of ['ref-yugothic', 'ref-malgun']) {
      const book = await openRef(name);
      const { drawings } = await book.openDrawings(1);
      expect(drawings.map((d) => [d.kind, d.name])).toEqual([
        ['image', 'pic1'],
        ['shape', 'rect1'],
        ['line', 'line1'],
        ['chart', 'chart1'],
      ]);
      expect(drawings[0]!.anchor).toMatchObject({
        type: 'twoCell',
        from: { col: 2, colOff: 38100, row: 3, rowOff: 25400 },
      });
      const r = await book.readImage(1, drawings[0]!.mediaPath!);
      expect(r).toMatchObject({ ok: true, format: 'png', width: 40, height: 30 });
    }
  });

  it('図の無い基準ファイル（ref-calibri-stdw12）は空の一覧', async () => {
    expect(await (await openRef('ref-calibri-stdw12')).openDrawings(1)).toEqual({
      drawings: [],
      notices: [],
    });
  });

  it('表の読み取りは図の有無に影響されない（セルの行数・列数が従来どおり）', async () => {
    const book = await openRef('ref-drawings');
    await book.openDrawings(1);
    const sheet = await book.openSheet(1);
    expect(sheet.rowCount).toBeGreaterThan(0);
  });
});
