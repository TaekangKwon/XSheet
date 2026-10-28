/**
 * ISSUE-91：図形（`xdr:sp`。線でないもの）の形・回転と反転・枠線と塗りの有無・枠線の太さと線種の読み取り（ADR-0028 §決定 3・7）。
 * 文字の有無を問わず持たせる。自由図形（`a:custGeom`）は字句を確かめて読み、上限を超えたら形を持たせず `shape-simplified`（1 シート 1 回）。
 * 上限を超えても XlsxError にせず、表は必ず読める。
 */
import { describe, expect, it } from 'vitest';
import { evaluateGeom } from '../../src/geom/evaluate';
import { openXlsx, type XlsxLimits } from '../../src/core/index';
import {
  buildXlsx,
  custGeomXml,
  grpSpXml,
  lineShapeXml,
  shapeXml,
  twoCellAnchor,
  wsDr,
  type CustGeomSpec,
  type ShapeSpec,
} from '../../src/testing/make-xlsx';
import type { SheetDrawing } from '../../src/core/types';

const A1 = { col: 0, row: 0 };
const B2 = { col: 1, row: 1 };

/** 図の要素（XML）を並べた 1 シートのブックを開き、図と、表の読み取りの成否を返す */
async function open(contents: string[], limits?: Partial<XlsxLimits>) {
  const bytes = await buildXlsx({
    sheets: [
      {
        name: 'S',
        rows: [['cell']],
        drawing: { xml: wsDr(contents.map((c) => twoCellAnchor(A1, B2, c))), rels: [] },
      },
    ],
  });
  const book = await openXlsx(bytes, { limits });
  const found = await book.openDrawings(1);
  const sheet = await book.openSheet(1);
  return { ...found, cell: sheet.getRows(1, 1)[0]!.cells[0] };
}

/** 図形 1 つを読んで、その図を返す */
async function one(spec: Partial<ShapeSpec>, limits?: Partial<XlsxLimits>): Promise<SheetDrawing> {
  const { drawings } = await open([shapeXml({ name: 's', ...spec })], limits);
  expect(drawings).toHaveLength(1);
  return drawings[0]!;
}

const gd = (name: string, fmla: string): string => `<a:gd name="${name}" fmla="${fmla}"/>`;
const pt = (x: string | number, y: string | number): string => `<a:pt x="${x}" y="${y}"/>`;
/** 自由図形の最小の path（moveTo→lnTo×2→close） */
const triPath = (extra = ''): string =>
  `<a:path w="100" h="100"${extra}><a:moveTo>${pt(0, 0)}</a:moveTo><a:lnTo>${pt(100, 0)}</a:lnTo><a:lnTo>${pt(50, 100)}</a:lnTo><a:close/></a:path>`;
const cust = (c: CustGeomSpec): Partial<ShapeSpec> => ({ custGeom: custGeomXml(c) });
const SOLID = '<a:solidFill><a:srgbClr val="000000"/></a:solidFill>';
/** preset の形の調整値（preset でなければ null） */
const adjOf = (d: SheetDrawing) => (d.geom?.type === 'preset' ? d.geom.adj : null);

describe('TD-91-1 preset の形と調整値', () => {
  it('prstGeom prst と avLst の全名前の調整値（val N）を読む。図形の項目は文字の無い図形にも付く', async () => {
    const d = await one({
      prst: 'wedgeRectCallout',
      avLst: gd('adj1', 'val -20833') + gd('adj2', 'val 62500') + gd('radius', 'val 16667'),
    });
    expect(d.kind).toBe('shape');
    expect(d.geom).toEqual({
      type: 'preset',
      prst: 'wedgeRectCallout',
      adj: [
        ['adj1', -20833],
        ['adj2', 62500],
        ['radius', 16667],
      ],
    });
  });

  it('avLst が空・無い：調整値は空（既定は描画側の表）。prst が表に無い名前でもそのまま持たせる', async () => {
    expect((await one({ prst: 'rect' })).geom).toEqual({ type: 'preset', prst: 'rect', adj: [] });
    expect((await one({ prst: 'someFutureShape' })).geom).toEqual({
      type: 'preset',
      prst: 'someFutureShape',
      adj: [],
    });
  });

  it('val 以外の式・名前の字句の誤り・17 文字の名前は読まない。16 文字の名前は読む', async () => {
    const d = await one({
      prst: 'roundRect',
      avLst:
        gd('adj', '*/ w 1 2') + // 式は読まない
        gd('a b', 'val 1') + // 空白入りの名前
        gd('a-b', 'val 2') + // 英数字以外
        gd('', 'val 3') + // 空の名前
        gd('A'.repeat(17), 'val 4') + // 17 文字
        gd('B'.repeat(16), 'val 5') + // 16 文字は読む
        gd('c', 'val 1.5') + // 整数でない
        gd('d', 'val 1e5') + // 指数は不可
        gd('e', ' val  7 '), // 前後・間の空白は許す
    });
    expect(d.geom).toEqual({
      type: 'preset',
      prst: 'roundRect',
      adj: [
        ['B'.repeat(16), 5],
        ['e', 7],
      ],
    });
  });

  it('調整値は 16 個まで（17 個目以降は読まない）。値は ±2,147,483,647 に丸める', async () => {
    const names = Array.from({ length: 17 }, (_v, i) => `a${i + 1}`);
    const d = await one({
      prst: 'gear9',
      avLst: names.map((n, i) => gd(n, `val ${i + 1}`)).join(''),
    });
    const adj = adjOf(d);
    expect(adj).toHaveLength(16);
    expect(adj![15]).toEqual(['a16', 16]);
    const big = await one({
      prst: 'rect',
      avLst: gd('hi', 'val 99999999999') + gd('lo', 'val -99999999999'),
    });
    expect(adjOf(big)).toEqual([
      ['hi', 2147483647],
      ['lo', -2147483647],
    ]);
  });

  it('小さな上限（maxShapeAdjustments）でも同じ。通知は出ない', async () => {
    const { drawings, notices } = await open(
      [
        shapeXml({
          name: 's',
          prst: 'rect',
          avLst: gd('a', 'val 1') + gd('b', 'val 2') + gd('c', 'val 3'),
        }),
      ],
      { maxShapeAdjustments: 2 },
    );
    expect(adjOf(drawings[0]!)).toEqual([
      ['a', 1],
      ['b', 2],
    ]);
    expect(notices).toEqual([]);
  });

  it('prstGeom も custGeom も無い図形は形を持たない（それ以外の項目は持つ）', async () => {
    const xml = shapeXml({ name: 's' }).replace(/<a:prstGeom[\s\S]*?<\/a:prstGeom>/, '');
    const { drawings } = await open([xml]);
    expect(drawings[0]!.geom).toBeUndefined();
    expect(drawings[0]!.hasFill).toBe(true);
  });

  it('文字のある図形・テキストボックスも同じ形を持つ（種類は変わらない）', async () => {
    const a = await one({ prst: 'ellipse', paras: ['x'] });
    const b = await one({ prst: 'diamond', txBox: true, paras: ['y'] });
    expect([a.kind, b.kind]).toEqual(['shape', 'textbox']);
    expect(a.geom).toMatchObject({ type: 'preset', prst: 'ellipse' });
    expect(b.geom).toMatchObject({ type: 'preset', prst: 'diamond' });
    expect(a.text).toBeDefined();
  });

  it('線として扱うもの（cxnSp・線の preset の sp・cxnSp の custGeom）は geom・outline を持たない（線のまま）', async () => {
    const { drawings } = await open([
      lineShapeXml({ name: 'a' }),
      lineShapeXml({ name: 'b', asSp: true, prst: 'bentConnector3' }),
      lineShapeXml({ name: 'c' }).replace(
        /<a:prstGeom[\s\S]*?<\/a:prstGeom>/,
        custGeomXml({ pathLst: triPath() }),
      ),
    ]);
    expect(drawings.map((d) => d.kind)).toEqual(['line', 'line', 'line']);
    for (const d of drawings) {
      expect(d.geom).toBeUndefined();
      expect(d.outline).toBeUndefined();
      expect(d.line).toBeDefined();
    }
  });
});

describe('TD-91-2 文字の無い図形の回転・反転・枠線と塗りの有無', () => {
  it('rot・flipH・flipV・枠線と塗りの有無を持つ（文字の有無を問わない）', async () => {
    const d = await one({
      xfrm: 'rot="5400000" flipH="1"',
      box: [0, 0, 1000, 500],
      fill: '<a:noFill/>',
      ln: `<a:ln w="9525">${SOLID}</a:ln>`,
    });
    expect(d).toMatchObject({
      rot: 5400000,
      flipH: true,
      flipV: false,
      hasOutline: true,
      hasFill: false,
    });
    const e = await one({
      xfrm: 'rot="-60000" flipV="1"',
      box: [0, 0, 1000, 500],
      ln: '<a:ln><a:noFill/></a:ln>',
    });
    expect(e).toMatchObject({
      rot: 21540000,
      flipH: false,
      flipV: true,
      hasOutline: false,
      hasFill: true,
    });
  });

  it('xfrm の無い図形：回転 0・反転なし。xdr:style が無く a:ln・塗りも無ければ「あり」', async () => {
    const d = await one({});
    expect(d).toMatchObject({
      rot: 0,
      flipH: false,
      flipV: false,
      hasOutline: true,
      hasFill: true,
    });
  });

  it('lnRef idx="0"・fillRef idx="0" は枠線・塗りなし', async () => {
    const d = await one({ style: { ln: 0, fill: 0 } });
    expect(d).toMatchObject({ hasOutline: false, hasFill: false });
  });
});

describe('TD-93-11 回転前の大きさ（xfrmExt）', () => {
  it('回転のある図形は a:xfrm の a:ext（回転前の大きさ・EMU）を持つ（回転前の矩形の決め方に使う）', async () => {
    const d = await one({
      xfrm: 'rot="2700000" flipH="1"',
      box: [5715000, 13970000, 1143000, 508000],
    });
    expect(d.xfrmExt).toEqual({ cx: 1143000, cy: 508000 });
    expect(d.rot).toBe(2700000);
  });

  it('回転が無い図形・a:ext が 0 の図形は持たない（既存の図の形を増やさない）', async () => {
    expect((await one({ box: [0, 0, 1000, 500] })).xfrmExt).toBeUndefined();
    expect((await one({ xfrm: 'rot="2700000"', box: [0, 0, 0, 500] })).xfrmExt).toBeUndefined();
  });
});

describe('TD-91-3 枠線の太さと線種（線と同じ規則）', () => {
  it('a:ln w を EMU で持つ。a:prstDash val が線種', async () => {
    const d = await one({
      ln: `<a:ln w="38100">${SOLID}<a:prstDash val="lgDashDot"/></a:ln>`,
    });
    expect(d.outline).toEqual({ widthEmu: 38100, dash: 'lgDashDot' });
  });

  it('a:ln w が無ければ lnRef idx 1・2・3 → 6350・12700・19050。idx 0・範囲外・xdr:style 無しは 0', async () => {
    const w = async (style?: { ln: number; fill: number }) =>
      (await one({ style, ln: `<a:ln>${SOLID}</a:ln>` })).outline!.widthEmu;
    expect(await w({ ln: 1, fill: 1 })).toBe(6350);
    expect(await w({ ln: 2, fill: 1 })).toBe(12700);
    expect(await w({ ln: 3, fill: 1 })).toBe(19050);
    expect(await w({ ln: 0, fill: 1 })).toBe(0);
    expect(await w({ ln: 4, fill: 1 })).toBe(0);
    expect(await w(undefined)).toBe(0);
    // a:ln 自体が無く lnRef だけ
    expect((await one({ style: { ln: 2, fill: 1 } })).outline).toEqual({
      widthEmu: 12700,
      dash: 'solid',
    });
  });

  it('a:ln w が lnRef より優先。範囲外は 0〜20,116,800 に丸め、整数でなければ lnRef に戻る', async () => {
    const mk = (w: string) => `<a:ln w="${w}">${SOLID}</a:ln>`;
    expect((await one({ ln: mk('25400'), style: { ln: 3, fill: 1 } })).outline!.widthEmu).toBe(
      25400,
    );
    expect((await one({ ln: mk('99999999999') })).outline!.widthEmu).toBe(20116800);
    expect((await one({ ln: mk('-5') })).outline!.widthEmu).toBe(0);
    expect((await one({ ln: mk('abc'), style: { ln: 2, fill: 1 } })).outline!.widthEmu).toBe(12700);
  });

  it('未知の線種は solid。枠線が見えない図形（noFill）も太さと線種は持つ', async () => {
    const d = await one({ ln: '<a:ln w="12700"><a:noFill/><a:prstDash val="weird"/></a:ln>' });
    expect(d.hasOutline).toBe(false);
    expect(d.outline).toEqual({ widthEmu: 12700, dash: 'solid' });
  });

  it('線の太さ・線種の読み取りは今のまま（同じ規則）', async () => {
    const { drawings } = await open([
      lineShapeXml({ name: 'l', ln: `<a:ln w="19050">${SOLID}<a:prstDash val="sysDot"/></a:ln>` }),
    ]);
    expect(drawings[0]!.line).toMatchObject({ widthEmu: 19050, dash: 'sysDot' });
  });
});

describe('TD-91-4 custGeom の読み取り', () => {
  it('avLst・gdLst・rect・pathLst（6 命令・w・h・fill・stroke）を読む', async () => {
    const d = await one(
      cust({
        avLst: gd('adj', 'val 25000'),
        gdLst: gd('x1', '*/ w adj 100000') + gd('y1', '+- h 0 hd2') + gd('c', 'val 7'),
        rect: 'l="l" t="t" r="x1" b="y1"',
        pathLst:
          `<a:path w="200" h="100" fill="lighten" stroke="0">` +
          `<a:moveTo>${pt(0, 0)}</a:moveTo><a:lnTo>${pt('x1', 50)}</a:lnTo>` +
          `<a:arcTo wR="10" hR="20" stAng="0" swAng="5400000"/>` +
          `<a:quadBezTo>${pt(1, 2)}${pt(3, 4)}</a:quadBezTo>` +
          `<a:cubicBezTo>${pt(1, 2)}${pt(3, 4)}${pt('y1', -6)}</a:cubicBezTo>` +
          `<a:close/></a:path>` +
          `<a:path><a:moveTo>${pt(0, 0)}</a:moveTo><a:lnTo>${pt(5, 5)}</a:lnTo></a:path>`,
      }),
    );
    expect(d.geom).toEqual({
      type: 'custom',
      def: {
        av: [['adj', 25000]],
        gd: [
          ['x1', '*/', 'w', 'adj', 100000],
          ['y1', '+-', 'h', 0, 'hd2'],
          ['c', 'val', 7],
        ],
        rect: ['l', 't', 'x1', 'y1'],
        paths: [
          {
            w: 200,
            h: 100,
            fill: 'lighten',
            stroke: false,
            c: [
              ['M', 0, 0],
              ['L', 'x1', 50],
              ['A', 10, 20, 0, 5400000],
              ['Q', 1, 2, 3, 4],
              ['C', 1, 2, 3, 4, 'y1', -6],
              ['Z'],
            ],
          },
          {
            c: [
              ['M', 0, 0],
              ['L', 5, 5],
            ],
          },
        ],
      },
    });
  });

  it('Excel が書く形（整数の点・w・h）：avLst・gdLst・rect が空でも読める。rect が無ければ rect を持たない', async () => {
    const d = await one(cust({ pathLst: triPath() }));
    expect(d.geom).toEqual({
      type: 'custom',
      def: {
        av: [],
        gd: [],
        paths: [
          {
            w: 100,
            h: 100,
            c: [['M', 0, 0], ['L', 100, 0], ['L', 50, 100], ['Z']],
          },
        ],
      },
    });
  });

  it('path の fill は 6 種だけ（未知は省く）。stroke="false"・"0" は線なし、"1"・"true" は省く。w・h は正の整数だけ', async () => {
    const paths = async (attrs: string) => {
      const d = await one(
        cust({ pathLst: `<a:path ${attrs}><a:moveTo>${pt(0, 0)}</a:moveTo></a:path>` }),
      );
      return d.geom?.type === 'custom' ? d.geom.def.paths[0] : undefined;
    };
    expect(await paths('fill="none" stroke="false"')).toEqual({
      fill: 'none',
      stroke: false,
      c: [['M', 0, 0]],
    });
    expect(await paths('fill="darkenLess" stroke="true"')).toEqual({
      fill: 'darkenLess',
      c: [['M', 0, 0]],
    });
    expect(await paths('fill="weird" stroke="1"')).toEqual({ c: [['M', 0, 0]] });
    expect(await paths('w="0" h="-5"')).toEqual({ c: [['M', 0, 0]] });
  });

  it('prstGeom と custGeom が両方あれば prstGeom を採る', async () => {
    const xml = shapeXml({ name: 's', prst: 'ellipse' }).replace(
      '</a:prstGeom>',
      `</a:prstGeom>${custGeomXml({ pathLst: triPath() })}`,
    );
    const { drawings } = await open([xml]);
    expect(drawings[0]!.geom).toMatchObject({ type: 'preset', prst: 'ellipse' });
  });

  it('自由図形も回転・反転・枠線の項目を持つ', async () => {
    const d = await one({
      ...cust({ pathLst: triPath() }),
      xfrm: 'rot="900000" flipV="1"',
      box: [0, 0, 1000, 1000],
      ln: `<a:ln w="12700">${SOLID}</a:ln>`,
    });
    expect(d).toMatchObject({ rot: 900000, flipV: true, hasOutline: true });
    expect(d.outline).toEqual({ widthEmu: 12700, dash: 'solid' });
  });
});

describe('TD-91-5 custGeom の字句の誤り：形を持たせない（通知なし・他の項目は持つ）', () => {
  const bad = async (c: CustGeomSpec): Promise<void> => {
    const { drawings, notices, cell } = await open([shapeXml({ name: 's', ...cust(c) })]);
    expect(drawings).toHaveLength(1);
    expect(drawings[0]!.geom).toBeUndefined();
    expect(drawings[0]!.hasFill).toBe(true);
    expect(notices).toEqual([]);
    expect(cell).toMatchObject({ text: 'cell' });
  };
  const okPath = triPath();

  it('未知の演算子・引数の数の誤り（多い・少ない）', async () => {
    await bad({ gdLst: gd('x', 'foo w 1 2'), pathLst: okPath });
    await bad({ gdLst: gd('x', '*/ w 1'), pathLst: okPath });
    await bad({ gdLst: gd('x', '+- w 1 2 3'), pathLst: okPath });
    await bad({ gdLst: gd('x', 'sin 5'), pathLst: okPath });
    await bad({ gdLst: gd('x', 'val'), pathLst: okPath });
    await bad({ gdLst: gd('x', ''), pathLst: okPath });
  });

  it('不正な名前・引数の字句（空白入り・記号・17 文字以上・小数・指数・桁の多すぎる整数）', async () => {
    await bad({ gdLst: gd('a b', 'val 1'), pathLst: okPath });
    await bad({ gdLst: gd('x', '*/ w 1.5 2'), pathLst: okPath });
    await bad({ gdLst: gd('x', '*/ w 1e3 2'), pathLst: okPath });
    await bad({ gdLst: gd('x', '*/ w a-b 2'), pathLst: okPath });
    await bad({ gdLst: gd('x', `*/ w ${'n'.repeat(17)} 2`), pathLst: okPath });
    await bad({ gdLst: gd('x', '*/ w 1234567890123456 2'), pathLst: okPath });
    await bad({ gdLst: gd('x'.repeat(17), 'val 1'), pathLst: okPath });
  });

  it('avLst の式が val でない・rect の値が不正', async () => {
    await bad({ avLst: gd('adj', '*/ w 1 2'), pathLst: okPath });
    await bad({ avLst: gd('adj', 'val x'), pathLst: okPath });
    await bad({ rect: 'l="0" t="0" r="1.5" b="b"', pathLst: okPath });
    await bad({ rect: 'l="0" t="0" r="r"', pathLst: okPath });
  });

  it('path の誤り：点の不正・点の数の不一致・未知の命令・path の外の点・arcTo の引数の不正', async () => {
    const p = (inner: string): string => `<a:path>${inner}</a:path>`;
    await bad({ pathLst: p(`<a:moveTo>${pt('a b', 0)}</a:moveTo>`) });
    await bad({ pathLst: p(`<a:moveTo>${pt(1.5, 0)}</a:moveTo>`) });
    await bad({ pathLst: p(`<a:moveTo>${pt(0, 0)}${pt(1, 1)}</a:moveTo>`) });
    await bad({ pathLst: p('<a:moveTo/>') });
    await bad({ pathLst: p(`<a:quadBezTo>${pt(0, 0)}</a:quadBezTo>`) });
    await bad({ pathLst: p(`<a:cubicBezTo>${pt(0, 0)}${pt(1, 1)}</a:cubicBezTo>`) });
    await bad({ pathLst: p(`<a:wobble>${pt(0, 0)}</a:wobble>`) });
    await bad({ pathLst: `${pt(0, 0)}` });
    await bad({ pathLst: p('<a:arcTo wR="1" hR="x y" stAng="0" swAng="0"/>') });
    await bad({ pathLst: p('<a:arcTo wR="1" hR="2" stAng="0"/>') });
  });

  it('pathLst が無い・path が無い・custGeom が空（自己閉じ）', async () => {
    const swap = (inner: string): string =>
      shapeXml({ name: 's' }).replace(/<a:prstGeom[\s\S]*?<\/a:prstGeom>/, inner);
    const xmls = [
      shapeXml({ name: 's', ...cust({}) }),
      swap('<a:custGeom><a:avLst/></a:custGeom>'),
      swap('<a:custGeom/>'),
    ];
    for (const xml of xmls) {
      const { drawings, notices } = await open([xml]);
      expect(drawings[0]!.geom).toBeUndefined();
      expect(notices).toEqual([]);
    }
  });
});

describe('TD-91-6 上限：超えても XlsxError にならず、shape-simplified が 1 回だけ', () => {
  const guides = (n: number): string =>
    Array.from({ length: n }, (_v, i) => gd(`g${i}`, 'val 1')).join('');
  const cmds = (n: number): string =>
    `<a:path w="10" h="10"><a:moveTo>${pt(0, 0)}</a:moveTo>${`<a:lnTo>${pt(1, 1)}</a:lnTo>`.repeat(n - 1)}</a:path>`;

  it('1 図形の guide（avLst＋gdLst）が 512 個ちょうどなら読み、513 個なら形を持たせず通知', async () => {
    const half = (n: number): CustGeomSpec => ({
      avLst: Array.from({ length: 10 }, (_v, i) => gd(`a${i}`, `val ${i}`)).join(''),
      gdLst: guides(n - 10),
      pathLst: cmds(2),
    });
    const ok = await open([shapeXml({ name: 's', ...cust(half(512)) })]);
    expect(ok.drawings[0]!.geom).toMatchObject({ type: 'custom' });
    expect(ok.notices).toEqual([]);
    const over = await open([shapeXml({ name: 's', ...cust(half(513)) })]);
    expect(over.drawings).toHaveLength(1);
    expect(over.drawings[0]!.geom).toBeUndefined();
    expect(over.notices).toEqual(['shape-simplified']);
    expect(over.cell).toMatchObject({ text: 'cell' });
  });

  it('1 図形の path 命令（全 path の合計）が 4,096 個ちょうどなら読み、4,097 個なら形を持たせず通知', async () => {
    const ok = await open([shapeXml({ name: 's', ...cust({ pathLst: cmds(4000) + cmds(96) }) })]);
    expect(ok.drawings[0]!.geom).toMatchObject({ type: 'custom' });
    expect(ok.notices).toEqual([]);
    const over = await open([shapeXml({ name: 's', ...cust({ pathLst: cmds(4000) + cmds(97) }) })]);
    expect(over.drawings[0]!.geom).toBeUndefined();
    expect(over.notices).toEqual(['shape-simplified']);
    expect(over.cell).toMatchObject({ text: 'cell' });
  });

  it('1 シートの path 命令の合計：超えた図形と、以後の自由図形は形を持たせず通知 1 回。preset は影響を受けない', async () => {
    const mk = (name: string, n: number): string =>
      shapeXml({ name, ...cust({ pathLst: cmds(n) }) });
    const { drawings, notices } = await open(
      [mk('a', 6), mk('b', 6), shapeXml({ name: 'p', prst: 'ellipse' }), mk('c', 1), mk('d', 2)],
      { maxSheetGeomUnits: 10 },
    );
    const byName = Object.fromEntries(drawings.map((d) => [d.name, d.geom?.type ?? null]));
    expect(byName).toEqual({ a: 'custom', b: null, p: 'preset', c: null, d: null });
    expect(notices).toEqual(['shape-simplified']);
  });

  it('合計がちょうど上限なら全部読む。小さな上限（maxGeomGuides・maxGeomCommands）でも同じ', async () => {
    const mk = (name: string, n: number): string =>
      shapeXml({ name, ...cust({ pathLst: cmds(n) }) });
    const exact = await open([mk('a', 5), mk('b', 5)], { maxSheetGeomUnits: 10 });
    expect(exact.drawings.map((d) => d.geom?.type)).toEqual(['custom', 'custom']);
    expect(exact.notices).toEqual([]);
    const cap = await open([mk('a', 3), mk('b', 4)], { maxGeomCommands: 3 });
    expect(cap.drawings.map((d) => d.geom?.type)).toEqual(['custom', undefined]);
    expect(cap.notices).toEqual(['shape-simplified']);
    const g = await open(
      [shapeXml({ name: 's', ...cust({ gdLst: guides(3), pathLst: cmds(2) }) })],
      { maxGeomGuides: 2 },
    );
    expect(g.drawings[0]!.geom).toBeUndefined();
    expect(g.notices).toEqual(['shape-simplified']);
  });

  it('既定の上限の値（ADR-0028 §決定 7）', async () => {
    const lim = await import('../../src/core/limits');
    expect(lim.MAX_GEOM_GUIDES).toBe(512);
    expect(lim.MAX_GEOM_COMMANDS).toBe(4096);
    expect(lim.MAX_SHEET_GEOM_UNITS).toBe(200_000);
    expect(lim.MAX_SHAPE_ADJUSTMENTS).toBe(16);
    expect(lim.DEFAULT_LIMITS).toMatchObject({
      maxGeomGuides: 512,
      maxGeomCommands: 4096,
      maxSheetGeomUnits: 200_000,
      maxShapeAdjustments: 16,
    });
  });

  it('図の数の上限で出さなかった図形は、自由図形の命令を数えない', async () => {
    const mk = (name: string): string => shapeXml({ name, ...cust({ pathLst: cmds(6) }) });
    const { drawings, notices } = await open([mk('a'), mk('b'), mk('c')], {
      maxDrawings: 1,
      maxSheetGeomUnits: 10,
    });
    expect(drawings.map((d) => d.name)).toEqual(['a']);
    expect(notices).toEqual(['too-many']);
  });

  it('通知の順は too-many → text-truncated → shape-simplified', async () => {
    const mk = (name: string): string =>
      shapeXml({ name, paras: ['x'.repeat(20)], ...cust({ pathLst: cmds(6) }) });
    const { notices } = await open([mk('a'), mk('b'), mk('c')], {
      maxDrawings: 2,
      maxShapeTextChars: 5,
      maxGeomCommands: 3,
    });
    expect(notices).toEqual(['too-many', 'text-truncated', 'shape-simplified']);
  });
});

describe('TD-91-7 グループの子にも同じ読み取りが掛かる', () => {
  it('グループの中の図形が geom・outline・回転を持つ。自由図形の上限もシート合計に数える', async () => {
    const child = (name: string, spec: Partial<ShapeSpec>): string =>
      shapeXml({ name, box: [1000, 2000, 1000, 500], ...spec });
    const g = grpSpXml(
      'g',
      [
        child('c1', { prst: 'can', avLst: gd('adj', 'val 30000'), xfrm: 'rot="5400000"' }),
        child('c2', cust({ pathLst: triPath() })),
        child('c3', cust({ pathLst: triPath() })),
      ],
      6,
      { box: [1000, 2000, 4000, 3000] },
    );
    const { drawings, notices } = await open([g], { maxSheetGeomUnits: 5 });
    expect(drawings.map((d) => d.name)).toEqual(['c1', 'c2', 'c3']);
    expect(drawings[0]).toMatchObject({
      rot: 5400000,
      geom: { type: 'preset', prst: 'can', adj: [['adj', 30000]] },
      outline: { widthEmu: 0, dash: 'solid' },
    });
    expect(drawings[1]!.geom).toMatchObject({ type: 'custom' }); // 4 命令
    expect(drawings[2]!.geom).toBeUndefined(); // 合計 8 > 5
    expect(notices).toEqual(['shape-simplified']);
  });
});

describe('TD-91-9 読み取った自由図形は評価器（geom の evaluate.ts）でそのまま描ける', () => {
  /** 自由図形の定義を取る（custom でなければテストを失敗させる） */
  const defOf = (d: SheetDrawing) => {
    if (d.geom?.type !== 'custom') throw new Error('自由図形でない');
    return d.geom.def;
  };
  /** 経路の点がすべて有限 */
  const finite = (r: ReturnType<typeof evaluateGeom>): boolean =>
    r.ok &&
    r.paths.every((p) =>
      p.subpaths.every(
        (sp) =>
          Number.isFinite(sp.start.x) &&
          Number.isFinite(sp.start.y) &&
          sp.segs.every((s) =>
            (s.t === 'L' ? [s.x, s.y] : [s.x1, s.y1, s.x2, s.y2, s.x, s.y]).every(Number.isFinite),
          ),
      ),
    );

  it('guide・rect・6 命令（arcTo・quadBezTo を含む）の定義：失敗せず、有限の経路と文字領域を返す', async () => {
    const d = await one(
      cust({
        avLst: gd('adj', 'val 25000'),
        gdLst: gd('x1', '*/ w adj 100000') + gd('y1', '+- h 0 hd2') + gd('r1', '+/ x1 y1 2'),
        rect: 'l="l" t="t" r="x1" b="y1"',
        pathLst:
          `<a:path w="200" h="100">` +
          `<a:moveTo>${pt(0, 0)}</a:moveTo><a:lnTo>${pt('x1', 50)}</a:lnTo>` +
          `<a:arcTo wR="10" hR="20" stAng="0" swAng="5400000"/>` +
          `<a:quadBezTo>${pt(1, 2)}${pt(3, 4)}</a:quadBezTo>` +
          `<a:cubicBezTo>${pt(1, 2)}${pt(3, 4)}${pt('y1', 6)}</a:cubicBezTo>` +
          `<a:close/></a:path>`,
      }),
    );
    const r = evaluateGeom(defOf(d), undefined, 120, 60);
    expect(r.ok).toBe(true);
    expect(finite(r)).toBe(true);
    if (!r.ok) return;
    expect(r.paths).toHaveLength(1);
    expect(r.paths[0]!.subpaths[0]!.closed).toBe(true);
    expect(r.textRect.r - r.textRect.l).toBeGreaterThan(0);
  });

  it('調整値（avLst）を渡して上書きできる：読み取った av の名前で評価器が引ける', async () => {
    const d = await one(
      cust({
        avLst: gd('adj', 'val 25000'),
        gdLst: gd('x1', '*/ w adj 100000'),
        pathLst: `<a:path><a:moveTo>${pt(0, 0)}</a:moveTo><a:lnTo>${pt('x1', 'h')}</a:lnTo></a:path>`,
      }),
    );
    const def = defOf(d);
    const a = evaluateGeom(def, undefined, 100, 100);
    const b = evaluateGeom(def, new Map([['adj', 50000]]), 100, 100);
    if (!a.ok || !b.ok) throw new Error('評価に失敗');
    expect(a.paths[0]!.subpaths[0]!.segs[0]).toMatchObject({ x: 25 });
    expect(b.paths[0]!.subpaths[0]!.segs[0]).toMatchObject({ x: 50 });
  });

  it('点の数値・path の w・h が 15 桁の整数でも ±2,147,483,647 に丸める（評価器は失敗せず有限の経路を返す）', async () => {
    const d = await one(
      cust({
        avLst: gd('adj', 'val 999999999999999'),
        pathLst:
          `<a:path w="999999999999999" h="999999999999999"><a:moveTo>${pt('-999999999999999', 0)}</a:moveTo>` +
          `<a:lnTo>${pt('999999999999999', 5)}</a:lnTo></a:path>`,
      }),
    );
    expect(defOf(d)).toEqual({
      av: [['adj', 2147483647]],
      gd: [],
      paths: [
        {
          w: 2147483647,
          h: 2147483647,
          c: [
            ['M', -2147483647, 0],
            ['L', 2147483647, 5],
          ],
        },
      ],
    });
    expect(finite(evaluateGeom(defOf(d), undefined, 100, 100))).toBe(true);
    // guide の式の整数の引数も同じ
    const g = await one(
      cust({
        gdLst: gd('x', '*/ w 999999999999999 1'),
        pathLst: `<a:path><a:moveTo>${pt('x', 0)}</a:moveTo></a:path>`,
      }),
    );
    expect(defOf(g).gd).toEqual([['x', '*/', 'w', 2147483647, 1]]);
    expect(finite(evaluateGeom(defOf(g), undefined, 100, 100))).toBe(true);
  });
});

describe('TD-91-10 シート合計の上限は「命令の数＋guide の数」', () => {
  const guides = (n: number, p = 'g'): string =>
    Array.from({ length: n }, (_v, i) => gd(`${p}${i}`, 'val 1')).join('');
  const empty = '<a:path/>';

  it('命令 0 個・guide だけの自由図形を並べても、合計が上限を超えたところで止まり shape-simplified が 1 回', async () => {
    const mk = (name: string): string =>
      shapeXml({ name, ...cust({ gdLst: guides(4), pathLst: empty }) });
    // 1 図形 = guide 4＋空の path 1 = 5。上限 12 なら 2 つまで（10）、3 つ目（15）から形なし
    const { drawings, notices } = await open([mk('a'), mk('b'), mk('c'), mk('d')], {
      maxSheetGeomUnits: 12,
    });
    expect(drawings.map((d) => d.geom?.type ?? null)).toEqual(['custom', 'custom', null, null]);
    expect(notices).toEqual(['shape-simplified']);
  });

  it('avLst の guide も数える。guide と命令の合計がちょうど上限なら読む', async () => {
    const mk = (name: string): string =>
      shapeXml({
        name,
        ...cust({
          avLst: gd('a', 'val 1'),
          gdLst: guides(2),
          pathLst: `<a:path><a:moveTo>${pt(0, 0)}</a:moveTo><a:lnTo>${pt(1, 1)}</a:lnTo></a:path>`,
        }),
      });
    // 1 図形 = guide 3＋命令 2 = 5
    const exact = await open([mk('a'), mk('b')], { maxSheetGeomUnits: 10 });
    expect(exact.drawings.map((d) => d.geom?.type)).toEqual(['custom', 'custom']);
    expect(exact.notices).toEqual([]);
    const over = await open([mk('a'), mk('b')], { maxSheetGeomUnits: 9 });
    expect(over.drawings.map((d) => d.geom?.type)).toEqual(['custom', undefined]);
    expect(over.notices).toEqual(['shape-simplified']);
  });

  it('1 つの図形の guide だけで残りを超える場合も、形を持たせず通知（1 図形の上限 512 以内でも）', async () => {
    const { drawings, notices } = await open(
      [shapeXml({ name: 's', ...cust({ gdLst: guides(20), pathLst: empty }) })],
      { maxSheetGeomUnits: 10 },
    );
    expect(drawings[0]!.geom).toBeUndefined();
    expect(notices).toEqual(['shape-simplified']);
  });

  it('命令の無い path（空の path）も 1 つと数える：1 図形の上限・シートの上限どちらにも掛かる', async () => {
    const paths = (n: number): string => empty.repeat(n);
    const a = await open([shapeXml({ name: 's', ...cust({ pathLst: paths(4) }) })], {
      maxGeomCommands: 3,
    });
    expect(a.drawings[0]!.geom).toBeUndefined();
    expect(a.notices).toEqual(['shape-simplified']);
    const ok = await open([shapeXml({ name: 's', ...cust({ pathLst: paths(3) }) })], {
      maxGeomCommands: 3,
    });
    expect(ok.drawings[0]!.geom).toMatchObject({ type: 'custom' });
    const b = await open([shapeXml({ name: 's', ...cust({ pathLst: paths(6) }) })], {
      maxSheetGeomUnits: 5,
    });
    expect(b.drawings[0]!.geom).toBeUndefined();
    expect(b.notices).toEqual(['shape-simplified']);
  });
});
