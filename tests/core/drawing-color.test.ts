/**
 * 図形・線・文字の色の読み取り（ISSUE-95・ADR-0028 §決定 6.1・6.4）。合成 xlsx で確かめる。
 * 出どころの順（直接の指定 → 参照の子の色 → 既定）、代表色、有無の規則が変わらないこと、グループの子にも掛かること。
 */
import { describe, expect, it } from 'vitest';
import { openXlsx } from '../../src/core/index';
import {
  buildThemeXml,
  buildXlsx,
  grpSpXml,
  lineShapeXml,
  shapeXml,
  twoCellAnchor,
  wsDr,
  type LineShapeSpec,
  type ShapeSpec,
} from '../../src/testing/make-xlsx';
import type { DrawingColor, SheetDrawing } from '../../src/core/types';

const A1 = { col: 0, row: 0 };
const B2 = { col: 1, row: 1 };

/** XML の並び（図の要素）を 1 つずつ twoCellAnchor に入れて openDrawings の結果を返す */
async function readXml(
  contents: string[],
  theme?: Record<string, string>,
): Promise<readonly SheetDrawing[]> {
  const bytes = await buildXlsx({
    sheets: [
      {
        name: 'S',
        rows: [['a']],
        drawing: { xml: wsDr(contents.map((c) => twoCellAnchor(A1, B2, c))), rels: [] },
      },
    ],
    ...(theme === undefined ? {} : { theme: { xml: buildThemeXml(theme) } }),
  });
  return (await (await openXlsx(bytes)).openDrawings(1)).drawings;
}

/** 図形 1 つを読む */
async function shapeOf(
  spec: Partial<ShapeSpec>,
  theme?: Record<string, string>,
): Promise<SheetDrawing> {
  const [d] = await readXml([shapeXml({ name: 's', ...spec })], theme);
  return d!;
}

/** 線 1 本を読む */
async function lineOf(
  spec: Partial<LineShapeSpec>,
  theme?: Record<string, string>,
): Promise<SheetDrawing> {
  const [d] = await readXml([lineShapeXml({ name: 'l', ...spec })], theme);
  return d!;
}

/** `#rrggbb` と不透明度（小数 2 桁）の短い表記にする（色が無ければ undefined） */
const hex = (c: DrawingColor | undefined): string | undefined =>
  c === undefined
    ? undefined
    : `#${c.rgb.toString(16).padStart(6, '0')}${c.alpha === 1 ? '' : `@${c.alpha.toFixed(2)}`}`;

const solid = (inner: string): string => `<a:solidFill>${inner}</a:solidFill>`;
const rgb = (v: string): string => `<a:srgbClr val="${v}"/>`;
const ln = (inner: string): string => `<a:ln w="9525">${inner}</a:ln>`;
const accent1 = '<a:schemeClr val="accent1"/>';
const textColorOf = (d: SheetDrawing): string | undefined =>
  hex(d.text?.paragraphs[0]?.runs[0]?.color);

describe('TD-95-1 Excel の既定の図形・テキストボックス・線の色', () => {
  it('既定の図形：塗り accent1・枠線 accent1＋shade 15000 は #172C51・文字 lt1', async () => {
    const d = await shapeOf({
      style: {
        ln: 2,
        fill: 1,
        lnClr: '<a:schemeClr val="accent1"><a:shade val="15000"/></a:schemeClr>',
        fontClr: '<a:schemeClr val="lt1"/>',
      },
      paras: ['文字'],
    });
    expect(hex(d.fillColor)).toBe('#4472c4');
    expect(hex(d.outlineColor)).toBe('#172c51');
    expect(textColorOf(d)).toBe('#ffffff');
  });

  it('テキストボックス：spPr の塗り lt1・枠線 lt1＋shade 50000 は #BCBCBC・文字 dk1', async () => {
    const d = await shapeOf({
      txBox: true,
      fill: solid('<a:schemeClr val="lt1"/>'),
      ln: ln(solid('<a:schemeClr val="lt1"><a:shade val="50000"/></a:schemeClr>')),
      style: {
        ln: 0,
        fill: 0,
        lnClr: accent1,
        fillClr: accent1,
        fontClr: '<a:schemeClr val="dk1"/>',
      },
      paras: ['文字'],
    });
    expect(hex(d.fillColor)).toBe('#ffffff');
    expect(hex(d.outlineColor)).toBe('#bcbcbc');
    expect(textColorOf(d)).toBe('#000000');
  });

  it('線：lnRef の子の色 accent1', async () => {
    const d = await lineOf({ style: { ln: 1, fill: 0 } });
    expect(d.kind).toBe('line');
    expect(hex(d.line!.color)).toBe('#4472c4');
  });

  it('テーマの色が違えば、その色になる（既定のテーマ以外を使う）', async () => {
    const theme = { dk1: rgb('111111'), lt1: rgb('EEEEEE'), accent1: rgb('FF8800') };
    const d = await shapeOf({ style: { ln: 1, fill: 1 }, paras: ['x'] }, theme);
    expect(hex(d.fillColor)).toBe('#ff8800');
    expect(hex(d.outlineColor)).toBe('#ff8800');
    expect(textColorOf(d)).toBe('#eeeeee');
    // 参照も直接の指定も無ければ、塗りは lt1・枠線と文字は tx1（＝dk1）
    const e = await shapeOf({ paras: ['x'] }, theme);
    expect([hex(e.fillColor), hex(e.outlineColor), textColorOf(e)]).toEqual([
      '#eeeeee',
      '#111111',
      '#111111',
    ]);
  });
});

describe('TD-95-2 塗りの出どころの順', () => {
  it('直接の指定が fillRef の子の色に勝つ。direct が無ければ fillRef の子の色（idx 1001 も同じ）', async () => {
    const clr = '<a:srgbClr val="00FF00"/>';
    expect(
      hex(
        (await shapeOf({ fill: solid(rgb('FF0000')), style: { ln: 0, fill: 1, fillClr: clr } }))
          .fillColor,
      ),
    ).toBe('#ff0000');
    expect(hex((await shapeOf({ style: { ln: 0, fill: 1, fillClr: clr } })).fillColor)).toBe(
      '#00ff00',
    );
    expect(hex((await shapeOf({ style: { ln: 0, fill: 1001, fillClr: clr } })).fillColor)).toBe(
      '#00ff00',
    );
  });

  it('解けない色（phClr・不正な値）はその出どころを飛ばして次へ進む', async () => {
    const ref = '<a:srgbClr val="00FF00"/>';
    // 直接の指定が解けない → fillRef の子
    const d1 = await shapeOf({
      fill: solid('<a:schemeClr val="phClr"/>'),
      style: { ln: 0, fill: 1, fillClr: ref },
    });
    expect(hex(d1.fillColor)).toBe('#00ff00');
    const d2 = await shapeOf({
      fill: solid(rgb('GGGGGG')),
      style: { ln: 0, fill: 1, fillClr: ref },
    });
    expect(hex(d2.fillColor)).toBe('#00ff00');
    // fillRef の子も解けない（phClr）→ 既定の lt1
    const d3 = await shapeOf({ style: { ln: 0, fill: 1, fillClr: '<a:schemeClr val="phClr"/>' } });
    expect(hex(d3.fillColor)).toBe('#ffffff');
    // 子が無い fillRef → lt1
    expect(hex((await shapeOf({ style: { ln: 0, fill: 1, fillClr: '' } })).fillColor)).toBe(
      '#ffffff',
    );
    // 色の要素の無い solidFill＋idx 0 の fillRef：塗りは「ある」（有無の規則）で、色は lt1
    const d4 = await shapeOf({ fill: '<a:solidFill/>', style: { ln: 0, fill: 0, fillClr: ref } });
    expect([d4.hasFill, hex(d4.fillColor)]).toEqual([true, '#ffffff']);
  });

  it('a:grpFill は未指定として扱う（fillRef の子の色 → 既定）', async () => {
    const d = await shapeOf({
      fill: '<a:grpFill/>',
      style: { ln: 0, fill: 1, fillClr: '<a:srgbClr val="123456"/>' },
    });
    expect(d.hasFill).toBe(true);
    expect(hex(d.fillColor)).toBe('#123456');
  });

  it('有無の規則は変わらない：a:noFill・idx 0 は色を持たない。塗りがあれば必ず持つ', async () => {
    const none = await shapeOf({ fill: '<a:noFill/>', style: { ln: 1, fill: 1 } });
    expect([none.hasFill, none.fillColor]).toEqual([false, undefined]);
    const idx0 = await shapeOf({ style: { ln: 1, fill: 0 } });
    expect([idx0.hasFill, idx0.fillColor]).toEqual([false, undefined]);
    const noStyle = await shapeOf({});
    expect([noStyle.hasFill, hex(noStyle.fillColor)]).toEqual([true, '#ffffff']);
  });

  it('色の変換（alpha・lumMod）も掛かる', async () => {
    const d = await shapeOf({
      fill: solid('<a:srgbClr val="FF0000"><a:alpha val="40000"/></a:srgbClr>'),
    });
    expect(hex(d.fillColor)).toBe('#ff0000@0.40');
    // accent1 ＋ lumMod 50000（HSL の明るさを半分）。U7 の照合までは現状の式のまま（#203864 か #1F3864）
    const e = await shapeOf({
      fill: solid('<a:schemeClr val="accent1"><a:lumMod val="50000"/></a:schemeClr>'),
    });
    expect(['#203864', '#1f3864']).toContain(hex(e.fillColor));
  });
});

describe('TD-95-3 枠線・線の出どころの順', () => {
  it('図形の枠線：a:ln の塗り → lnRef の子 → tx1。矢じりの有無・太さは変えない', async () => {
    const lnClr = '<a:srgbClr val="00FF00"/>';
    expect(
      hex(
        (await shapeOf({ ln: ln(solid(rgb('FF0000'))), style: { ln: 1, fill: 1, lnClr } }))
          .outlineColor,
      ),
    ).toBe('#ff0000');
    expect(hex((await shapeOf({ style: { ln: 1, fill: 1, lnClr } })).outlineColor)).toBe('#00ff00');
    // a:ln に塗りの指定が無く lnRef の子の色
    expect(
      hex(
        (await shapeOf({ ln: '<a:ln w="12700"/>', style: { ln: 2, fill: 1, lnClr } })).outlineColor,
      ),
    ).toBe('#00ff00');
    // 解けない → 次の出どころ → 既定の tx1
    const d = await shapeOf({
      ln: ln(solid('<a:schemeClr val="phClr"/>')),
      style: { ln: 1, fill: 1, lnClr: '<a:schemeClr val="phClr"/>' },
    });
    expect(hex(d.outlineColor)).toBe('#000000');
    const e = await shapeOf({
      ln: ln(solid('<a:schemeClr val="phClr"/>')),
      style: { ln: 1, fill: 1, lnClr },
    });
    expect(hex(e.outlineColor)).toBe('#00ff00');
  });

  it('枠線の有無は変わらない：a:ln の a:noFill・idx 0 は枠線の色を持たない', async () => {
    const a = await shapeOf({ ln: ln('<a:noFill/>'), style: { ln: 1, fill: 1 } });
    expect([a.hasOutline, a.outlineColor]).toEqual([false, undefined]);
    const b = await shapeOf({ style: { ln: 0, fill: 1 } });
    expect([b.hasOutline, b.outlineColor]).toEqual([false, undefined]);
    // 枠線が見えない図形の枠線の色は持たず、塗りの色は持つ
    expect(hex(a.fillColor)).toBe('#4472c4');
  });

  it('線（cxnSp・線の preset の sp）：a:ln の塗り → lnRef の子 → tx1。見えない線は出ない', async () => {
    const lnClr = '<a:srgbClr val="00FF00"/>';
    expect(
      hex(
        (await lineOf({ ln: ln(solid(rgb('FF0000'))), style: { ln: 1, fill: 0, lnClr } })).line!
          .color,
      ),
    ).toBe('#ff0000');
    expect(hex((await lineOf({ style: { ln: 1, fill: 0, lnClr } })).line!.color)).toBe('#00ff00');
    expect(hex((await lineOf({ style: { ln: 1, fill: 0, lnClr: '' } })).line!.color)).toBe(
      '#000000',
    );
    expect(hex((await lineOf({})).line!.color)).toBe('#000000');
    expect(hex((await lineOf({ asSp: true, ln: ln(solid(rgb('0000FF'))) })).line!.color)).toBe(
      '#0000ff',
    );
    const none = await readXml([lineShapeXml({ name: 'l', ln: ln('<a:noFill/>') })]);
    expect(none).toEqual([]);
  });

  it('線の太い指定・矢じり付きでも色は a:ln の塗りから読む（矢じりは線と同じ色＝別に持たない）', async () => {
    const d = await lineOf({
      ln: `<a:ln w="76200">${solid(rgb('336699'))}<a:tailEnd type="triangle"/></a:ln>`,
    });
    expect(hex(d.line!.color)).toBe('#336699');
    expect(d.line!.tail?.type).toBe('triangle');
  });
});

describe('TD-95-4 代表色（グラデーション・模様・画像）', () => {
  const gs = (pos: number, clr: string): string => `<a:gs pos="${pos}">${clr}</a:gs>`;
  const grad = (stops: string): string =>
    `<a:gradFill><a:gsLst>${stops}</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>`;
  const fillOf = async (fill: string): Promise<string | undefined> =>
    hex((await shapeOf({ fill })).fillColor);

  it('グラデーション：50% の位置の色（2 stop は線形 RGB の中間・1 stop はその色）', async () => {
    // 2 stop は線形 RGB で混ぜる（Excel の画素：黒→白の中央は #BCBCBC 付近）
    expect(await fillOf(grad(gs(0, rgb('000000')) + gs(100000, rgb('FFFFFF'))))).toBe('#bcbcbc');
    expect(await fillOf(grad(gs(0, rgb('FF0000')) + gs(100000, rgb('FF0000'))))).toBe('#ff0000');
    expect(await fillOf(grad(gs(30000, rgb('336699'))))).toBe('#336699');
  });

  it('グラデーション：3 stop は前後の stop の補間。pos の順が乱れていても pos の順に並べる', async () => {
    // 0:黒・50000:赤・100000:白 → 50% は赤そのもの
    const ordered = gs(0, rgb('000000')) + gs(50000, rgb('FF0000')) + gs(100000, rgb('FFFFFF'));
    expect(await fillOf(grad(ordered))).toBe('#ff0000');
    const shuffled = gs(100000, rgb('FFFFFF')) + gs(0, rgb('000000')) + gs(50000, rgb('FF0000'));
    expect(await fillOf(grad(shuffled))).toBe('#ff0000');
    // 0:黒・20000:黒・100000:白 → 50% は 20000〜100000 の (50-20)/80 = 0.375 → 0.375*255 ≒ 96
    const three = gs(0, rgb('000000')) + gs(20000, rgb('000000')) + gs(100000, rgb('FFFFFF'));
    expect(await fillOf(grad(three))).toBe('#606060');
    // 50% より前の stop だけ・後の stop だけなら、端の色
    expect(await fillOf(grad(gs(10000, rgb('FF0000')) + gs(20000, rgb('0000FF'))))).toBe('#0000ff');
    expect(await fillOf(grad(gs(60000, rgb('FF0000')) + gs(80000, rgb('0000FF'))))).toBe('#ff0000');
  });

  it('グラデーション：不透明度も補間する。解けない stop は飛ばす。stop が無い・全部解けなければ fillRef 以降へ', async () => {
    const half =
      gs(0, '<a:srgbClr val="FF0000"><a:alpha val="100000"/></a:srgbClr>') +
      gs(100000, '<a:srgbClr val="FF0000"><a:alpha val="0"/></a:srgbClr>');
    expect(await fillOf(grad(half))).toBe('#ff0000@0.50');
    expect(
      await fillOf(grad(gs(0, '<a:schemeClr val="phClr"/>') + gs(100000, rgb('00FF00')))),
    ).toBe('#00ff00');
    expect(await fillOf(grad(''))).toBe('#ffffff');
    expect(await fillOf(grad(gs(0, '<a:schemeClr val="phClr"/>')))).toBe('#ffffff');
  });

  it('グラデーション：stop は 16 個まで（17 個目以降は読まない）', async () => {
    const first16 = Array.from({ length: 16 }, () => gs(0, rgb('FF0000'))).join('');
    // 17 個目（白・100000）を読むと 50% が赤から動く。読まなければ赤のまま
    expect(await fillOf(grad(first16 + gs(100000, rgb('FFFFFF'))))).toBe('#ff0000');
    expect(
      await fillOf(
        grad(first16.slice(0, -1 * gs(0, rgb('FF0000')).length) + gs(100000, rgb('FFFFFF'))),
      ),
    ).toBe('#ff8080');
  });

  const patt = (prst: string, inner: string): string =>
    `<a:pattFill prst="${prst}">${inner}</a:pattFill>`;
  const fg = (c: string): string => `<a:fgClr>${c}</a:fgClr>`;
  const bg = (c: string): string => `<a:bgClr>${c}</a:bgClr>`;

  it('模様：前景色と背景色を、pct の割合（その他は 50%）で混ぜる', async () => {
    const both = fg(rgb('000000')) + bg(rgb('FFFFFF'));
    expect(await fillOf(patt('pct25', both))).toBe('#bfbfbf'); // 前景 25%
    expect(await fillOf(patt('pct75', both))).toBe('#404040');
    expect(await fillOf(patt('pct5', both))).toBe('#f2f2f2');
    expect(await fillOf(patt('horz', both))).toBe('#808080');
    expect(await fillOf(patt('pct15', both))).toBe('#808080'); // 名前の無い割合はその他
  });

  it('模様：片方しか無ければその色。両方無い・解けなければ fillRef 以降へ', async () => {
    expect(await fillOf(patt('pct25', fg(rgb('FF0000'))))).toBe('#ff0000');
    expect(await fillOf(patt('pct25', bg(rgb('00FF00'))))).toBe('#00ff00');
    expect(await fillOf(patt('pct25', ''))).toBe('#ffffff');
    expect(await fillOf(patt('pct25', fg('<a:schemeClr val="phClr"/>')))).toBe('#ffffff');
  });

  it('画像の塗り：画像は読まず #D9D9D9（不透明）', async () => {
    expect(
      await fillOf(
        '<a:blipFill><a:blip r:embed="rId9"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>',
      ),
    ).toBe('#d9d9d9');
  });

  it('線の a:gradFill・a:pattFill も同じ規則', async () => {
    const g = await shapeOf({ ln: ln(grad(gs(0, rgb('000000')) + gs(100000, rgb('FFFFFF')))) });
    expect(hex(g.outlineColor)).toBe('#bcbcbc');
    const p = await lineOf({ ln: ln(patt('pct25', fg(rgb('000000')) + bg(rgb('FFFFFF')))) });
    expect(hex(p.line!.color)).toBe('#bfbfbf');
  });
});

describe('TD-95-5 文字の色', () => {
  const run = (rPrInner: string | undefined, text = 'あ') => ({
    text,
    ...(rPrInner === undefined ? {} : { rPrInner }),
  });
  const colorsOf = (d: SheetDrawing): (string | undefined)[] =>
    d.text!.paragraphs.flatMap((p) => p.runs.map((r) => hex(r.color)));

  it('ランの a:rPr の塗り → fontRef の子の色 → tx1（ラン単位で違う色にできる）', async () => {
    const d = await shapeOf({
      style: { ln: 0, fill: 0, fontClr: '<a:srgbClr val="00FF00"/>' },
      paras: [
        { runs: [run(solid(rgb('FF0000'))), run(undefined), run('<a:latin typeface="Arial"/>')] },
      ],
    });
    expect(colorsOf(d)).toEqual(['#ff0000', '#00ff00', '#00ff00']);
    // xdr:style が無い・fontRef に子が無い → tx1
    expect(colorsOf(await shapeOf({ paras: [{ runs: [run(undefined)] }] }))).toEqual(['#000000']);
    expect(
      colorsOf(
        await shapeOf({
          style: { ln: 0, fill: 0, fontClr: '' },
          paras: [{ runs: [run(undefined)] }],
        }),
      ),
    ).toEqual(['#000000']);
  });

  it('a:noFill は透明（不透明度 0）。a:gradFill は代表色。解けない色は fontRef へ進む', async () => {
    const gradFill =
      '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="000000"/></a:gs><a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs></a:gsLst></a:gradFill>';
    const d = await shapeOf({
      style: { ln: 0, fill: 0, fontClr: '<a:srgbClr val="00FF00"/>' },
      paras: [
        {
          runs: [
            run('<a:noFill/>'),
            run(gradFill),
            run(solid('<a:schemeClr val="phClr"/>')),
            run(solid('<a:srgbClr val="12"/>')),
          ],
        },
      ],
    });
    expect(d.text!.paragraphs[0]!.runs[0]!.color!.alpha).toBe(0);
    expect(colorsOf(d).slice(1)).toEqual(['#bcbcbc', '#00ff00', '#00ff00']);
  });

  it('a:fld のランも同じ。a:br にも色を持つ', async () => {
    const d = await shapeOf({
      paras: [
        {
          runs: [
            { text: 'セル', fld: true, rPrInner: solid(rgb('0000FF')) },
            { br: true },
            run(solid(rgb('FF00FF'))),
          ],
        },
      ],
    });
    expect(colorsOf(d)).toEqual(['#0000ff', '#000000', '#ff00ff']);
  });

  it('a:rPr の中の a:ln（文字の輪郭）・a:uFill（下線の色）・a:highlight の色は文字の色にしない', async () => {
    const inner =
      `<a:ln w="9525">${solid(rgb('FF0000'))}</a:ln>` +
      `<a:highlight>${rgb('FFFF00')}</a:highlight>` +
      `<a:uFill>${solid(rgb('00FFFF'))}</a:uFill>`;
    const d = await shapeOf({ paras: [{ runs: [run(inner), run(inner + solid(rgb('112233')))] }] });
    expect(colorsOf(d)).toEqual(['#000000', '#112233']);
  });

  it('文字のある図形でも塗り・枠線の色は文字と別に読む', async () => {
    const d = await shapeOf({
      fill: solid(rgb('FFFF00')),
      ln: ln(solid(rgb('FF0000'))),
      paras: [{ runs: [run(solid(rgb('0000FF')))] }],
    });
    expect([hex(d.fillColor), hex(d.outlineColor), textColorOf(d)]).toEqual([
      '#ffff00',
      '#ff0000',
      '#0000ff',
    ]);
  });
});

describe('TD-95-6 グループの子・文字の無い図形・形の種類', () => {
  it('グループの子の図形・線・文字にも同じ規則で色が付く', async () => {
    const box: [number, number, number, number] = [0, 0, 100, 100];
    const group = grpSpXml(
      'g',
      [
        shapeXml({
          name: 'c1',
          box,
          fill: solid(rgb('FF0000')),
          ln: ln(solid(rgb('00FF00'))),
          paras: [{ runs: [{ text: 'a', rPrInner: solid(rgb('0000FF')) }] }],
        }),
        lineShapeXml({ name: 'c2', box, ln: ln(solid(rgb('112233'))) }),
        shapeXml({ name: 'c3', id: 9, box, style: { ln: 1, fill: 1 } }),
      ],
      6,
      { box: [0, 0, 100, 100] },
    );
    const ds = await readXml([group]);
    expect(ds.map((d) => d.name)).toEqual(['c1', 'c2', 'c3']);
    expect([hex(ds[0]!.fillColor), hex(ds[0]!.outlineColor), textColorOf(ds[0]!)]).toEqual([
      '#ff0000',
      '#00ff00',
      '#0000ff',
    ]);
    expect(hex(ds[1]!.line!.color)).toBe('#112233');
    expect([hex(ds[2]!.fillColor), hex(ds[2]!.outlineColor)]).toEqual(['#4472c4', '#4472c4']);
  });

  it('文字の無い図形・線として扱わない sp すべて・自由図形・画像/グラフは、図形だけが色を持つ', async () => {
    const d = await shapeOf({ prst: 'ellipse', fill: solid(rgb('FF0000')) });
    expect(d.text).toBeUndefined();
    expect(hex(d.fillColor)).toBe('#ff0000');
    const pic = await readXml([
      '<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="c"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"/></a:graphic></xdr:graphicFrame>',
    ]);
    expect([pic[0]!.fillColor, pic[0]!.outlineColor, pic[0]!.line]).toEqual([
      undefined,
      undefined,
      undefined,
    ]);
  });

  it('色の値は整数と不透明度だけ（文字列を持たない）', async () => {
    const d = await shapeOf({
      fill: solid('<a:prstClr val="red"/>'),
      ln: ln(solid(rgb('0A0B0C'))),
      paras: ['x'],
    });
    for (const c of [d.fillColor!, d.outlineColor!, d.text!.paragraphs[0]!.runs[0]!.color!]) {
      expect(Object.keys(c).sort()).toEqual(['alpha', 'rgb']);
      expect(Number.isInteger(c.rgb)).toBe(true);
    }
  });
});

describe('TD-95-7 非信頼の入力', () => {
  it('色の要素の細工（閉じの無い塗り・巨大な変換の列・壊れた rPr）でも例外にならず、数 ms で返る', async () => {
    const many = '<a:lumMod val="1"/>'.repeat(5000);
    const unclosed = '<a:gradFill><a:gsLst>' + '<a:gs pos="1">'.repeat(3000);
    const d = await shapeOf({
      fill: unclosed,
      ln: ln(solid(`<a:srgbClr val="FF0000">${many}</a:srgbClr>`)),
      paras: [{ runs: [{ text: 'x', rPrInner: '<a:solidFill>'.repeat(3000) }] }],
    });
    expect(d.kind).toBe('shape');
    expect(d.outlineColor).toBeDefined();
    expect(d.text!.paragraphs[0]!.runs[0]!.color).toBeDefined();
  });
});

describe('TD-95-8 参照の idx が数値でない・欠けているとき', () => {
  it('子の色があれば子の色を使う（fillRef・lnRef・fontRef）。子が無ければ既定。idx 0 は今までどおり見ない', async () => {
    const clr = '<a:srgbClr val="00FF00"/>';
    const xml = shapeXml({
      name: 'x',
      style: { ln: 1, fill: 1, lnClr: clr, fillClr: clr, fontClr: clr },
      paras: ['a'],
    });
    for (const broken of [' idx="x"', ' idx=""', '']) {
      const bad = xml.replace(/(<a:(?:lnRef|fillRef)) idx="1"/g, `$1${broken}`);
      expect(bad).not.toBe(xml);
      const [d] = await readXml([bad]);
      expect([hex(d!.fillColor), hex(d!.outlineColor), textColorOf(d!)]).toEqual([
        '#00ff00',
        '#00ff00',
        '#00ff00',
      ]);
    }
    const noChild = shapeXml({
      name: 'y',
      style: { ln: 1, fill: 1, lnClr: '', fillClr: '' },
      paras: ['a'],
    }).replace(/ idx="1"/g, '');
    const [e] = await readXml([noChild]);
    expect([hex(e!.fillColor), hex(e!.outlineColor)]).toEqual(['#ffffff', '#000000']);
    const idx0 = await shapeOf({
      style: { ln: 0, fill: 0, lnClr: clr, fillClr: clr },
      ln: ln(solid('<a:schemeClr val="phClr"/>')),
      fill: solid('<a:schemeClr val="phClr"/>'),
    });
    expect([hex(idx0.fillColor), hex(idx0.outlineColor)]).toEqual(['#ffffff', '#000000']);
  });
});
