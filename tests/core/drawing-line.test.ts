/**
 * 線・矢印と、図形の枠線・塗りの有無・回転・反転の読み取り（ISSUE-87・ADR-0027 §決定 1・3）。合成 xlsx で確かめる。
 */
import { describe, expect, it } from 'vitest';
import { openXlsx } from '../../src/core/index';
import {
  alternateContent,
  buildXlsx,
  graphicFrameXml,
  lineShapeXml,
  picXml,
  shapeXml,
  twoCellAnchor,
  wsDr,
  type LineShapeSpec,
  type ShapeSpec,
} from '../../src/testing/make-xlsx';
import type { DrawingLine, SheetDrawing } from '../../src/core/types';

const A1 = { col: 0, row: 0 };
const B2 = { col: 1, row: 1 };

/** XML の並び（図の要素）を 1 つずつ twoCellAnchor に入れて openDrawings の結果を返す */
async function readXml(contents: string[]): Promise<readonly SheetDrawing[]> {
  const bytes = await buildXlsx({
    sheets: [
      {
        name: 'S',
        rows: [['a']],
        drawing: { xml: wsDr(contents.map((c) => twoCellAnchor(A1, B2, c))), rels: [] },
      },
    ],
  });
  return (await (await openXlsx(bytes)).openDrawings(1)).drawings;
}

/** 線 1 本を読む（出なければ undefined） */
async function lineOf(spec: Omit<LineShapeSpec, 'name'>): Promise<DrawingLine | undefined> {
  const [d] = await readXml([lineShapeXml({ name: 'l', ...spec })]);
  return d?.line;
}

const SOLID = '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>';

describe('TD-87-10 線：cxnSp と線の preset の sp は line で出る', () => {
  it('cxnSp は prstGeom を問わず line。線の preset 10 種の sp も line（文字は読まない）。それ以外の sp は図形のまま', async () => {
    const presets = [
      'line',
      'straightConnector1',
      'bentConnector2',
      'bentConnector3',
      'bentConnector4',
      'bentConnector5',
      'curvedConnector2',
      'curvedConnector3',
      'curvedConnector4',
      'curvedConnector5',
    ];
    const ds = await readXml([
      lineShapeXml({ name: 'cxn-rect', prst: 'rect' }),
      ...presets.map((p) => shapeXml({ name: `sp-${p}`, prst: p, paras: ['文字'] })),
      shapeXml({ name: 'arrow', prst: 'rightArrow', paras: ['文字'] }),
    ]);
    expect(ds.map((d) => [d.kind, d.name])).toEqual([
      ['line', 'cxn-rect'],
      ...presets.map((p) => ['line', `sp-${p}`]),
      ['shape', 'arrow'],
    ]);
    // 線にした sp は文字・枠線・塗りの項目を持たない
    for (const d of ds.slice(1, 11)) {
      expect(d.text).toBeUndefined();
      expect(d.hasOutline).toBeUndefined();
      expect(d.hasFill).toBeUndefined();
    }
    expect(ds[11]!.text).toBeDefined();
  });

  it('形：線の preset はそのまま、それ以外（rect・未知の値・preset 無し）の cxnSp は直線（line）', async () => {
    const ds = await readXml([
      lineShapeXml({ name: 'a', prst: 'bentConnector3' }),
      lineShapeXml({ name: 'b', prst: 'rect' }),
      lineShapeXml({ name: 'c', prst: 'curvedConnector9' }),
      lineShapeXml({ name: 'd', asSp: true, prst: 'curvedConnector5' }),
    ]);
    expect(ds.map((d) => d.line!.geom)).toEqual([
      'bentConnector3',
      'line',
      'line',
      'curvedConnector5',
    ]);
  });

  it('既存の cxnSpXml（xdr:style も a:ln も無い）の線も出る（他のツールが書いたファイルで線を落とさない）', async () => {
    const ds = await readXml([
      '<xdr:cxnSp macro=""><xdr:nvCxnSpPr><xdr:cNvPr id="4" name="bare"/><xdr:cNvCxnSpPr/></xdr:nvCxnSpPr><xdr:spPr><a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom></xdr:spPr></xdr:cxnSp>',
    ]);
    expect(ds).toHaveLength(1);
    expect(ds[0]).toMatchObject({
      kind: 'line',
      line: { widthEmu: 0, dash: 'solid', head: null, tail: null },
    });
  });
});

describe('TD-87-11 足さない線：隠した線と見えない線', () => {
  const noFillLn = '<a:ln><a:noFill/></a:ln>';
  const cases: [string, Omit<LineShapeSpec, 'name'>, boolean][] = [
    ['cNvPr hidden', { hidden: true }, false],
    ['a:ln の a:noFill', { ln: noFillLn }, false],
    ['a:ln の a:noFill（lnRef idx=2 でも）', { ln: noFillLn, style: { ln: 2, fill: 0 } }, false],
    ['線の塗りの指定が無く lnRef idx=0', { style: { ln: 0, fill: 0 } }, false],
    ['a:ln に w だけで lnRef idx=0', { ln: '<a:ln w="9525"/>', style: { ln: 0, fill: 0 } }, false],
    ['xdr:style も a:ln も無い', {}, true],
    ['lnRef idx=2（a:ln 無し）', { style: { ln: 2, fill: 0 } }, true],
    [
      '線の塗りがあれば lnRef idx=0 でも出る',
      { ln: `<a:ln>${SOLID}</a:ln>`, style: { ln: 0, fill: 0 } },
      true,
    ],
    [
      'a:ln の塗りの指定が無く lnRef idx=2',
      { ln: '<a:ln w="9525"/>', style: { ln: 2, fill: 0 } },
      true,
    ],
    ['gradFill・pattFill の線', { ln: '<a:ln><a:gradFill/></a:ln>' }, true],
  ];
  for (const [label, spec, visible] of cases) {
    it(`${label} → ${visible ? '出る' : '出ない'}（cxnSp・線の sp とも）`, async () => {
      const ds = await readXml([
        lineShapeXml({ name: 'c', ...spec }),
        lineShapeXml({ name: 's', asSp: true, ...spec }),
      ]);
      expect(ds.map((d) => d.name)).toEqual(visible ? ['c', 's'] : []);
    });
  }

  it('見えない線は数えない（上限 1 でも、あとの見える線が出る）', async () => {
    const bytes = await buildXlsx({
      sheets: [
        {
          name: 'S',
          rows: [['a']],
          drawing: {
            xml: wsDr(
              [
                lineShapeXml({ name: 'hidden', hidden: true }),
                lineShapeXml({ name: 'invisible', ln: '<a:ln><a:noFill/></a:ln>' }),
                lineShapeXml({ name: 'visible' }),
              ].map((c) => twoCellAnchor(A1, B2, c)),
            ),
            rels: [],
          },
        },
      ],
    });
    const book = await openXlsx(bytes, { limits: { maxLines: 1 } });
    const { drawings, notices } = await book.openDrawings(1);
    expect(drawings.map((d) => d.name)).toEqual(['visible']);
    expect(notices).toEqual([]);
  });
});

describe('TD-87-12 線：調整値', () => {
  it('無ければ 50000。a:gd name="adjN" fmla="val N" だけを読む（val 以外の式・小数・adj の名前違いは読まない）', async () => {
    expect((await lineOf({}))!.adj).toEqual([50000, 50000, 50000]);
    const avLst =
      '<a:gd name="adj1" fmla="val 160000"/><a:gd name="adj2" fmla="*/ w 1 2"/><a:gd name="adj3" fmla="val -25000"/><a:gd name="adj4" fmla="val 7"/><a:gd name="adj" fmla="val 9"/>';
    expect((await lineOf({ prst: 'bentConnector5', avLst }))!.adj).toEqual([160000, 50000, -25000]);
    expect((await lineOf({ avLst: '<a:gd name="adj1" fmla="val 12.5"/>' }))!.adj[0]).toBe(50000);
  });

  it('Excel の U 字の値（1,800,000）や 99,999,999 は丸めない。巨大な桁・不正な値は既定値', async () => {
    const avLst =
      '<a:gd name="adj1" fmla="val 1800000"/><a:gd name="adj2" fmla="val -99999999"/><a:gd name="adj3" fmla="val 99999999999999999"/>';
    expect((await lineOf({ prst: 'bentConnector5', avLst }))!.adj).toEqual([
      1800000, -99999999, 50000,
    ]);
    expect(
      (await lineOf({
        avLst: '<a:gd name="adj1" fmla="val NaN"/><a:gd name="adj2" fmla="val "/>',
      }))!.adj,
    ).toEqual([50000, 50000, 50000]);
  });

  it('32bit 整数（±2,147,483,647）の境界：ちょうどは通り、超える値は上限に丸める', async () => {
    const at =
      '<a:gd name="adj1" fmla="val 2147483647"/><a:gd name="adj2" fmla="val -2147483647"/>';
    expect((await lineOf({ prst: 'bentConnector4', avLst: at }))!.adj).toEqual([
      2147483647, -2147483647, 50000,
    ]);
    const over =
      '<a:gd name="adj1" fmla="val 2147483648"/><a:gd name="adj2" fmla="val -2147483649"/><a:gd name="adj3" fmla="val 999999999999999"/>';
    expect((await lineOf({ prst: 'bentConnector5', avLst: over }))!.adj).toEqual([
      2147483647, -2147483647, 2147483647,
    ]);
  });

  it('prstGeom の外（custGeom の a:gd）は読まない', async () => {
    const [d] = await readXml([
      '<xdr:cxnSp macro=""><xdr:nvCxnSpPr><xdr:cNvPr id="4" name="c"/><xdr:cNvCxnSpPr/></xdr:nvCxnSpPr><xdr:spPr><a:custGeom><a:avLst><a:gd name="adj1" fmla="val 1"/></a:avLst></a:custGeom></xdr:spPr></xdr:cxnSp>',
    ]);
    expect(d!.line).toMatchObject({ geom: 'line', adj: [50000, 50000, 50000] });
  });
});

describe('TD-87-13 線：太さ（a:ln w・lnRef idx の既定）・線種', () => {
  it('a:ln w（EMU）。無ければ lnRef idx 1・2・3 → 6350・12700・19050。どちらも無ければ 0', async () => {
    expect((await lineOf({ ln: '<a:ln w="38100"/>', style: { ln: 3, fill: 0 } }))!.widthEmu).toBe(
      38100,
    );
    expect((await lineOf({ style: { ln: 1, fill: 0 } }))!.widthEmu).toBe(6350);
    expect((await lineOf({ style: { ln: 2, fill: 0 } }))!.widthEmu).toBe(12700);
    expect((await lineOf({ style: { ln: 3, fill: 0 } }))!.widthEmu).toBe(19050);
    expect((await lineOf({ style: { ln: 7, fill: 0 } }))!.widthEmu).toBe(0);
    expect((await lineOf({}))!.widthEmu).toBe(0);
    expect((await lineOf({ ln: '<a:ln w="x"/>', style: { ln: 2, fill: 0 } }))!.widthEmu).toBe(
      12700,
    );
  });

  it('w は 0〜20,116,800 に丸める', async () => {
    expect((await lineOf({ ln: '<a:ln w="-5"/>' }))!.widthEmu).toBe(0);
    expect((await lineOf({ ln: '<a:ln w="999999999"/>' }))!.widthEmu).toBe(20116800);
  });

  it('線種（a:prstDash val）：11 種はそのまま、無い・未知の値は solid', async () => {
    const dashes = [
      'solid',
      'dot',
      'dash',
      'lgDash',
      'dashDot',
      'lgDashDot',
      'lgDashDotDot',
      'sysDash',
      'sysDot',
      'sysDashDot',
      'sysDashDotDot',
    ];
    for (const v of dashes) {
      expect((await lineOf({ ln: `<a:ln><a:prstDash val="${v}"/></a:ln>` }))!.dash).toBe(v);
    }
    expect((await lineOf({ ln: '<a:ln><a:prstDash val="zigzag"/></a:ln>' }))!.dash).toBe('solid');
    expect((await lineOf({ ln: `<a:ln>${SOLID}</a:ln>` }))!.dash).toBe('solid');
    expect((await lineOf({}))!.dash).toBe('solid');
  });
});

describe('TD-87-14 線：矢じり（headEnd＝始点・tailEnd＝終点の type・w・len）', () => {
  const ln = (inner: string) => `<a:ln>${SOLID}${inner}</a:ln>`;

  it('type・w・len を読む。w・len が無い・未知の値は med', async () => {
    const l = await lineOf({
      ln: ln('<a:headEnd type="triangle" w="lg" len="sm"/><a:tailEnd type="stealth"/>'),
    });
    expect(l!.head).toEqual({ type: 'triangle', w: 'lg', len: 'sm' });
    expect(l!.tail).toEqual({ type: 'stealth', w: 'med', len: 'med' });
    const u = await lineOf({ ln: ln('<a:tailEnd type="oval" w="huge" len="x"/>') });
    expect(u!.tail).toEqual({ type: 'oval', w: 'med', len: 'med' });
  });

  it('5 種（triangle・stealth・arrow・diamond・oval）。未知の種類は triangle', async () => {
    for (const t of ['triangle', 'stealth', 'arrow', 'diamond', 'oval']) {
      expect((await lineOf({ ln: ln(`<a:tailEnd type="${t}"/>`) }))!.tail!.type).toBe(t);
    }
    expect((await lineOf({ ln: ln('<a:tailEnd type="foo"/>') }))!.tail!.type).toBe('triangle');
  });

  it('type="none"・type が無い（w だけ）・要素が無いは矢じりなし（null）', async () => {
    const l = await lineOf({ ln: ln('<a:headEnd type="none"/><a:tailEnd w="lg"/>') });
    expect(l!.head).toBeNull();
    expect(l!.tail).toBeNull();
    const none = await lineOf({ ln: ln('') });
    expect(none).toMatchObject({ head: null, tail: null });
  });
});

describe('TD-87-15 回転・反転（xfrm rot・flipH・flipV）', () => {
  it('線：rot（1/60000 度）・flipH・flipV。Excel の接続線のとおり rot="16200000" flipH="1"', async () => {
    const ds = await readXml([
      lineShapeXml({ name: 'a', xfrm: 'rot="16200000" flipH="1"' }),
      lineShapeXml({ name: 'b', xfrm: 'flipV="1"' }),
      lineShapeXml({ name: 'c', xfrm: 'rot="10800000" flipH="1" flipV="1"' }),
      lineShapeXml({ name: 'd', xfrm: 'flipH="true" flipV="false"' }),
      lineShapeXml({ name: 'e' }),
    ]);
    expect(ds.map((d) => [d.rot, d.flipH, d.flipV])).toEqual([
      [16200000, true, false],
      [0, false, true],
      [10800000, true, true],
      [0, true, false],
      [0, false, false],
    ]);
  });

  it('rot は 0〜21,599,999 に正規化する（負・1 周以上）。整数でなければ 0', async () => {
    const ds = await readXml(
      ['-5400000', '21600000', '23400000', '-43200000', 'abc', ''].map((rot, i) =>
        lineShapeXml({ name: `r${i}`, xfrm: `rot="${rot}"` }),
      ),
    );
    expect(ds.map((d) => d.rot)).toEqual([16200000, 0, 1800000, 0, 0, 0]);
  });

  it('図形は文字の有無を問わず rot・flipH・flipV を持つ（ISSUE-91）', async () => {
    const ds = await readXml([
      shapeXml({ name: 'text', paras: ['x'], xfrm: 'rot="5400000" flipV="1"' }),
      shapeXml({ name: 'notext', xfrm: 'rot="5400000"' }),
    ]);
    expect(ds[0]).toMatchObject({ rot: 5400000, flipH: false, flipV: true });
    expect(ds[1]).toMatchObject({ rot: 5400000, flipH: false, flipV: false });
  });
});

describe('TD-87-16 文字のある図形の枠線と塗りの有無（色は読まない）', () => {
  const NOFILL = '<a:noFill/>';
  const cases: [string, Omit<ShapeSpec, 'name'>, [boolean, boolean]][] = [
    ['どちらも指定なし（xdr:style も無い）', {}, [true, true]],
    [
      'spPr 直下の noFill と a:ln の noFill（ラベル）',
      { fill: NOFILL, ln: '<a:ln><a:noFill/></a:ln>' },
      [false, false],
    ],
    ['xdr:style の lnRef・fillRef が idx 2・1', { style: { ln: 2, fill: 1 } }, [true, true]],
    ['lnRef・fillRef が idx 0', { style: { ln: 0, fill: 0 } }, [false, false]],
    [
      '塗りの指定があれば fillRef idx=0 でも塗りあり',
      { fill: SOLID, style: { ln: 0, fill: 0 } },
      [false, true],
    ],
    [
      'a:ln の noFill は塗りの有無に影響しない（直下の塗りは solidFill）',
      { fill: SOLID, ln: '<a:ln><a:noFill/></a:ln>' },
      [false, true],
    ],
    [
      'a:ln の solidFill は塗りの有無に影響しない（直下が noFill）',
      { fill: NOFILL, ln: `<a:ln>${SOLID}</a:ln>` },
      [true, false],
    ],
    [
      'a:ln に塗りの指定が無く lnRef idx=0',
      { ln: '<a:ln w="9525"/>', style: { ln: 0, fill: 1 } },
      [false, true],
    ],
    [
      'a:ln に塗りの指定が無く lnRef idx=2',
      { ln: '<a:ln w="9525"/>', style: { ln: 2, fill: 0 } },
      [true, false],
    ],
  ];
  for (const [label, spec, [outline, fill]] of cases) {
    it(`${label} → 枠線${outline ? 'あり' : 'なし'}・塗り${fill ? 'あり' : 'なし'}`, async () => {
      const [d] = await readXml([shapeXml({ name: 's', paras: ['x'], ...spec })]);
      expect([d!.hasOutline, d!.hasFill]).toEqual([outline, fill]);
    });
  }

  it('Office 拡張（a:extLst の中の a14:hiddenFill・hiddenLine）の塗り・線は読まない', async () => {
    const ext =
      '<a:extLst><a:ext uri="{909E8E84-426E-40DD-AFC4-6F175D3DCCD1}"><a14:hiddenFill xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main">' +
      SOLID +
      '</a14:hiddenFill></a:ext></a:extLst>';
    const [d] = await readXml([
      shapeXml({ name: 's', paras: ['x'], fill: NOFILL, ln: '<a:ln><a:noFill/></a:ln>' + ext }),
    ]);
    expect([d!.hasOutline, d!.hasFill]).toEqual([false, false]);
  });
});

describe('TD-87-16b spPr の extLst で切る位置：spPr 直下の extLst だけ（a:ln・a:blip の中の extLst では切らない）', () => {
  const EXT =
    '<a:extLst><a:ext uri="{1}"><a14:hiddenFill xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main">' +
    SOLID +
    '</a14:hiddenFill></a:ext></a:extLst>';
  const INNER_EXT = '<a:extLst><a:ext uri="{2}"/></a:extLst>';
  const NOFILL = '<a:noFill/>';

  it('a:ln の中に extLst があっても a:ln は途中で切れない：a:ln の noFill が効き、後ろの要素も読める', async () => {
    const ln = '<a:ln w="9525"><a:noFill/>' + INNER_EXT + '</a:ln>';
    const [d] = await readXml([shapeXml({ name: 's', paras: ['x'], fill: SOLID, ln })]);
    expect([d!.hasOutline, d!.hasFill]).toEqual([false, true]);
  });

  it('a:ln の中の extLst の後ろに spPr 直下の extLst（hiddenFill）が続いても、後者から先は読まない', async () => {
    const ln = '<a:ln><a:noFill/>' + INNER_EXT + '</a:ln>' + EXT;
    const [d] = await readXml([shapeXml({ name: 's', paras: ['x'], fill: NOFILL, ln })]);
    expect([d!.hasOutline, d!.hasFill]).toEqual([false, false]);
  });

  it('blip の中の extLst が先にあっても、その後ろの a:ln を読む（線の太さ・矢じり）', async () => {
    const fill = '<a:blipFill><a:blip r:embed="rId1">' + INNER_EXT + '</a:blip></a:blipFill>';
    const ln =
      '<a:ln w="19050"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:tailEnd type="triangle"/>' +
      INNER_EXT +
      '</a:ln>';
    const line = await lineOf({ fill, ln } as Omit<LineShapeSpec, 'name'>);
    expect(line).toMatchObject({ widthEmu: 19050, tail: { type: 'triangle' } });
  });

  it('線：a:ln の中に extLst があっても、矢じり・線種を読む', async () => {
    const ln =
      '<a:ln w="12700"><a:prstDash val="dash"/><a:headEnd type="oval"/>' + INNER_EXT + '</a:ln>';
    expect(await lineOf({ ln })).toMatchObject({
      widthEmu: 12700,
      dash: 'dash',
      head: { type: 'oval' },
    });
  });
});

describe('TD-87-17 mc:AlternateContent：アンカーの中で Choice に graphicFrame があれば、uri を問わずその枝', () => {
  const CHARTEX = 'http://schemas.microsoft.com/office/drawing/2014/chartex';

  it('Choice が chartex の graphicFrame・Fallback が文字入りの図形：グラフになり、Fallback の文字は出ない', async () => {
    const ds = await readXml([
      alternateContent(
        graphicFrameXml('waterfall', CHARTEX),
        shapeXml({
          name: 'fb',
          paras: ['この図は、お使いのバージョンの Excel では利用できません。'],
        }),
      ),
    ]);
    expect(ds.map((d) => [d.kind, d.name])).toEqual([['chart', 'waterfall']]);
    expect(ds.some((d) => d.text !== undefined)).toBe(false);
  });

  it('chart の graphicFrame の Choice も、その枝を採る（Fallback の図形は出ない）', async () => {
    const ds = await readXml([
      alternateContent(graphicFrameXml('chart'), shapeXml({ name: 'fb', paras: ['x'] })),
    ]);
    expect(ds.map((d) => [d.kind, d.name])).toEqual([['chart', 'chart']]);
    expect(ds.some((d) => d.text !== undefined)).toBe(false);
  });

  it('スライサー（2010/slicer）・タイムライン（2012/timeslicer）：Fallback の説明文入りの図形を描かない', async () => {
    const ds = await readXml([
      alternateContent(
        graphicFrameXml('sl', 'http://schemas.microsoft.com/office/drawing/2010/slicer'),
        shapeXml({
          name: 'sl-fb',
          paras: ['このスライサーは Excel 2010 以降でサポートされています。'],
        }),
      ),
      alternateContent(
        graphicFrameXml('tl', 'http://schemas.microsoft.com/office/drawing/2012/timeslicer'),
        shapeXml({
          name: 'tl-fb',
          paras: ['このタイムラインは Excel 2013 以降でサポートされています。'],
        }),
      ),
    ]);
    expect(ds.map((d) => [d.kind, d.name])).toEqual([
      ['chart', 'sl'],
      ['chart', 'tl'],
    ]);
    expect(ds.some((d) => d.text !== undefined)).toBe(false);
  });

  it('graphicFrame が無い Choice は Fallback を採る。pic を含む枝は graphicFrame より優先', async () => {
    const ds = await readXml([
      alternateContent(shapeXml({ name: 'c1' }), shapeXml({ name: 'f1', paras: ['x'] })),
      alternateContent(graphicFrameXml('gf'), picXml({ name: 'fb-pic', embed: 'rId1' })),
    ]);
    expect(ds.map((d) => d.name)).toEqual(['f1', 'fb-pic']);
  });
});
