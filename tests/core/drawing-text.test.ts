/**
 * 図形の文字の読み取り（ISSUE-87・ADR-0027 §決定 1）。合成 xlsx（`make-xlsx.ts`）で、段落・ラン・改行・フィールド・書式・本文の枠・箇条書き・実体参照を確かめる。
 */
import { describe, expect, it } from 'vitest';
import { openXlsx } from '../../src/core/index';
import {
  buildXlsx,
  grpSpXml,
  shapeXml,
  spXml,
  twoCellAnchor,
  wsDr,
  type ParaSpec,
  type ShapeSpec,
} from '../../src/testing/make-xlsx';
import type { DrawingText, SheetDrawing } from '../../src/core/types';

const A1 = { col: 0, row: 0 };
const B2 = { col: 1, row: 1 };

/** 図形 1 つ（twoCellAnchor）の openDrawings の結果を返す */
async function readShapes(specs: ShapeSpec[]): Promise<readonly SheetDrawing[]> {
  const anchors = specs.map((s) => twoCellAnchor(A1, B2, shapeXml(s)));
  const bytes = await buildXlsx({
    sheets: [{ name: 'S', rows: [['a']], drawing: { xml: wsDr(anchors), rels: [] } }],
  });
  return (await (await openXlsx(bytes)).openDrawings(1)).drawings;
}

/** 図形 1 つの文字を返す（無ければ失敗） */
async function textOf(spec: Omit<ShapeSpec, 'name'>): Promise<DrawingText> {
  const [d] = await readShapes([{ name: 's', ...spec }]);
  expect(d!.text).toBeDefined();
  return d!.text!;
}

/** 段落ごとの本文（改行は \n） */
const flat = (t: DrawingText): string[] =>
  t.paragraphs.map((p) => p.runs.map((r) => (r.br ? '\n' : r.text)).join(''));

/** 既定の書式のラン */
const run = (text: string, over: object = {}) => ({
  br: false,
  text,
  sizePt: 11,
  bold: false,
  italic: false,
  underline: false,
  strike: false,
  color: { rgb: 0, alpha: 1 }, // 既定（xdr:style も a:rPr の塗りも無い）はテーマの tx1＝黒
  ...over,
});

describe('TD-87-1 文字：段落・ラン・a:br・a:fld（TxLink）を出現順に読む', () => {
  it('段落ごとに、ランと改行とフィールドを出現順に並べる（空の段落も 1 段落）', async () => {
    const t = await textOf({
      paras: ['one', { runs: ['a', { br: true }, 'b', { text: 'c', fld: true }] }, { runs: [] }],
    });
    expect(flat(t)).toEqual(['one', 'a\nbc', '']);
    expect(t.paragraphs[1]!.runs.map((r) => r.br)).toEqual([false, true, false, false]);
    expect(t.paragraphs[0]!.runs[0]).toEqual(run('one'));
  });

  it('セル参照の文字（sp textlink）は a:fld type="TxLink" の a:t で拾う', async () => {
    const t = await textOf({
      textlink: '$A$1',
      paras: [{ runs: [{ text: 'セルの値', fld: true }] }],
    });
    expect(flat(t)).toEqual(['セルの値']);
  });

  it('テキストボックス（txBox）も図形も文字を読む。種類は今までどおり', async () => {
    const ds = await readShapes([
      { name: 'tb', txBox: true, paras: ['t'] },
      { name: 'sh', paras: ['s'] },
    ]);
    expect(ds.map((d) => [d.kind, d.name, d.text ? flat(d.text) : null])).toEqual([
      ['textbox', 'tb', ['t']],
      ['shape', 'sh', ['s']],
    ]);
  });

  it('文字のある図形は、回転・反転・枠線と塗りの有無も持つ（文字の無い図形は持たない）', async () => {
    const [d] = await readShapes([{ name: 's', paras: ['x'], style: { ln: 2, fill: 1 } }]);
    expect(d).toMatchObject({
      rot: 0,
      flipH: false,
      flipV: false,
      hasOutline: true,
      hasFill: true,
    });
  });

  it('グループの中の図形は読まない（今までどおり 1 つの「グループ」。文字は持たない）', async () => {
    const group = grpSpXml('g', [shapeXml({ name: 'inner', paras: ['中の文字'] })]);
    const bytes = await buildXlsx({
      sheets: [
        {
          name: 'S',
          rows: [['a']],
          drawing: { xml: wsDr([twoCellAnchor(A1, B2, group)]), rels: [] },
        },
      ],
    });
    const { drawings } = await (await openXlsx(bytes)).openDrawings(1);
    expect(drawings).toHaveLength(1);
    expect(drawings[0]).toMatchObject({ kind: 'group', name: 'g' });
    expect(drawings[0]!.text).toBeUndefined();
  });
});

describe('TD-87-2 文字の無い図形（ランの文字が無い）', () => {
  it('a:endParaRPr だけ・空の a:t・改行だけ・txBody なしは、text を持たない（図形の項目は持つ。ISSUE-91）', async () => {
    const ds = await readShapes([
      { name: 'endOnly', paras: [{ runs: [], endSz: 1800 }] },
      { name: 'emptyT', paras: [{ runs: [''] }] },
      { name: 'brOnly', paras: [{ runs: [{ br: true }] }] },
      { name: 'noBody' },
    ]);
    for (const d of ds) {
      expect(d.kind).toBe('shape');
      expect(Object.keys(d).sort()).toEqual([
        'anchor',
        'descr',
        'fillColor',
        'flipH',
        'flipV',
        'geom',
        'hasFill',
        'hasOutline',
        'kind',
        'mediaPath',
        'name',
        'outline',
        'outlineColor',
        'rot',
      ]);
    }
  });

  it('1 つでも文字のランがあれば文字のある図形（ほかの段落が空でも）', async () => {
    const t = await textOf({ paras: [{ runs: [] }, 'x', { runs: [{ br: true }] }] });
    expect(flat(t)).toEqual(['', 'x', '\n']);
  });
});

describe('TD-87-3 文字の書式：大きさ（既定・丸め）・太字・斜体・下線・取り消し線', () => {
  it('a:rPr sz（1/100pt）を pt にする。無ければ 11・1〜400 に丸める・小数も持つ', async () => {
    const t = await textOf({
      paras: [
        {
          runs: [
            { text: 'a', sz: 1800 },
            { text: 'b' },
            { text: 'c', sz: 5 },
            { text: 'd', sz: 99999 },
            { text: 'e', sz: 1050 },
            { text: 'f', sz: 'x' as never },
          ],
        },
      ],
    });
    expect(t.paragraphs[0]!.runs.map((r) => r.sizePt)).toEqual([18, 11, 1, 400, 10.5, 11]);
  });

  it('太字・斜体・下線（none 以外）・取り消し線（noStrike 以外）', async () => {
    const t = await textOf({
      paras: [
        {
          runs: [
            { text: 'a', b: true, i: true, u: 'sng', strike: 'sngStrike' },
            { text: 'b', b: false, i: false, u: 'none', strike: 'noStrike' },
            { text: 'c' },
            { text: 'd', u: 'dbl', strike: 'dblStrike' },
          ],
        },
      ],
    });
    const r = t.paragraphs[0]!.runs;
    expect(r[0]).toEqual(run('a', { bold: true, italic: true, underline: true, strike: true }));
    expect(r[1]).toEqual(run('b'));
    expect(r[2]).toEqual(run('c'));
    expect(r[3]).toEqual(run('d', { underline: true, strike: true }));
  });

  it('空の段落の大きさは a:endParaRPr sz（無ければ 11）。a:br の大きさは a:br の a:rPr', async () => {
    const t = await textOf({
      paras: [{ runs: [], endSz: 2400 }, { runs: [] }, { runs: ['x', { br: true, sz: 1400 }] }],
    });
    expect(t.paragraphs.map((p) => p.endSizePt)).toEqual([24, 11, 11]);
    expect(t.paragraphs[2]!.runs[1]).toEqual(run('', { br: true, sizePt: 14 }));
  });

  it('読まないもの（強調表示・フォント名・上付き・行間・字下げ・ハイパーリンク）は結果に出ない。色（a:rPr の塗り）と、ハイパーリンクの文字は読む', async () => {
    const raw =
      '<a:pPr marL="342900" indent="-342900" algn="ctr"><a:lnSpc><a:spcPct val="150000"/></a:lnSpc><a:spcBef><a:spcPts val="600"/></a:spcBef></a:pPr>' +
      '<a:r><a:rPr lang="ja-JP" sz="1200" baseline="30000"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:highlight><a:srgbClr val="FFFF00"/></a:highlight><a:latin typeface="Arial"/><a:ea typeface="MS Gothic"/><a:hlinkClick r:id="rId9"/></a:rPr><a:t>link</a:t></a:r>';
    const t = await textOf({ paras: [{ raw } as ParaSpec] });
    expect(t.paragraphs[0]!.runs).toEqual([
      run('link', { sizePt: 12, color: { rgb: 0xff0000, alpha: 1 } }),
    ]);
    expect(t.paragraphs[0]!.align).toBe('ctr');
    expect(Object.keys(t.paragraphs[0]!).sort()).toEqual(['align', 'bullet', 'endSizePt', 'runs']);
    expect(Object.keys(t).sort()).toEqual([
      'anchor',
      'fontScale',
      'insets',
      'noWrap',
      'paragraphs',
      'vert',
    ]);
  });
});

describe('TD-87-4 段落の揃えと本文の枠（縦位置・余白・折り返し・縮小・縦書き）', () => {
  it('段落の揃え：l・ctr・r・just・dist。無い・未知の値は l', async () => {
    const t = await textOf({
      paras: ['l', 'ctr', 'r', 'just', 'dist', 'xyz', 'none'].map((algn, i) => ({
        runs: [String(i)],
        ...(i === 0 ? {} : { algn }),
      })),
    });
    expect(t.paragraphs.map((p) => p.align)).toEqual(['l', 'ctr', 'r', 'just', 'dist', 'l', 'l']);
  });

  it('既定値：縦位置 t・余白 91440/45720/91440/45720・折り返す・縮めない・横書き', async () => {
    const t = await textOf({ paras: ['x'] });
    expect(t).toMatchObject({
      anchor: 't',
      insets: { l: 91440, t: 45720, r: 91440, b: 45720 },
      noWrap: false,
      fontScale: 1,
      vert: 'horz',
    });
  });

  it('bodyPr の値：anchor・lIns/tIns/rIns/bIns・wrap="none"・vert・normAutofit fontScale', async () => {
    const t = await textOf({
      paras: ['x'],
      bodyPr: 'anchor="ctr" lIns="0" tIns="10" rIns="20" bIns="30" wrap="none" vert="eaVert"',
      bodyPrInner: '<a:normAutofit fontScale="92500" lnSpcReduction="10000"/>',
    });
    expect(t).toMatchObject({
      anchor: 'ctr',
      insets: { l: 0, t: 10, r: 20, b: 30 },
      noWrap: true,
      fontScale: 0.925,
      vert: 'eaVert',
    });
    expect((await textOf({ paras: ['x'], bodyPr: 'anchor="b" wrap="square"' })).anchor).toBe('b');
    expect((await textOf({ paras: ['x'], bodyPr: 'anchor="b" wrap="square"' })).noWrap).toBe(false);
  });

  it('縦書き（vert）の 7 つの値はそのまま持つ', async () => {
    const verts = [
      'horz',
      'vert',
      'vert270',
      'wordArtVert',
      'eaVert',
      'mongolianVert',
      'wordArtVertRtl',
    ];
    const ds = await readShapes(
      verts.map((v, i) => ({ name: `s${i}`, paras: ['x'], bodyPr: `vert="${v}"` })),
    );
    expect(ds.map((d) => d.text!.vert)).toEqual(verts);
  });

  it('未知・不正な値は既定値に直す：anchor（t・ctr・b 以外は t）・vert（horz）・余白（既定。負は 0）・fontScale（0.01〜1 に丸める）', async () => {
    const t = await textOf({
      paras: ['x'],
      bodyPr: 'anchor="just" vert="xyz" lIns="abc" tIns="-5" rIns="" bIns="99999999999999999"',
      bodyPrInner: '<a:normAutofit fontScale="0"/>',
    });
    expect(t).toMatchObject({
      anchor: 't',
      vert: 'horz',
      insets: { l: 91440, t: 0, r: 91440, b: 45720 },
      fontScale: 0.01,
    });
    const big = await textOf({ paras: ['x'], bodyPrInner: '<a:normAutofit fontScale="250000"/>' });
    expect(big.fontScale).toBe(1);
    const none = await textOf({ paras: ['x'], bodyPrInner: '<a:noAutofit/>' });
    expect(none.fontScale).toBe(1);
  });
});

describe('TD-87-5 箇条書き：buChar・buAutoNum（番号・startAt・種類の切り替わり）', () => {
  it('buChar はその文字。buAutoNum は同じ種類が続く段落を 1 から数える。種類が変わる・箇条書きの無い段落をはさむと数え直す', async () => {
    const num = (type: string, startAt?: number) => ({ autoNum: type, startAt });
    const t = await textOf({
      paras: [
        { runs: ['a'], bullet: { char: '•' } },
        { runs: ['b'], bullet: num('arabicPeriod') },
        { runs: ['c'], bullet: num('arabicPeriod') },
        { runs: ['d'], bullet: num('alphaLcPeriod') },
        { runs: ['e'] },
        { runs: ['f'], bullet: num('arabicPeriod', 5) },
        { runs: ['g'], bullet: num('arabicPeriod', 5) },
        { runs: ['h'], bullet: num('arabicPeriod') },
        { runs: ['i'], bullet: 'none' },
        { runs: ['j'], bullet: num('arabicPeriod') },
      ],
    });
    // g は startAt が同じ 5 なので続けて 6。h は startAt なし＝続き（7）。none ではさむと j は 1 に戻る
    expect(t.paragraphs.map((p) => p.bullet)).toEqual([
      '•',
      '1.',
      '2.',
      '1.',
      null,
      '5.',
      '6.',
      '7.',
      null,
      '1.',
    ]);
  });

  it('startAt が途中で変わったら、そこから数え直す。同じ段落の buChar は buAutoNum があれば使わない', async () => {
    const t = await textOf({
      paras: [
        { runs: ['x'], bullet: { autoNum: 'arabicPeriod' } },
        { runs: ['y'], bullet: { autoNum: 'arabicPeriod', startAt: 3 } },
        { runs: ['z'], bullet: { autoNum: 'arabicPeriod', startAt: 3 } },
      ],
    });
    expect(t.paragraphs.map((p) => p.bullet)).toEqual(['1.', '3.', '4.']);
    const both = await textOf({
      paras: [
        {
          raw: '<a:pPr><a:buAutoNum type="arabicPeriod"/><a:buChar char="x"/></a:pPr><a:r><a:t>a</a:t></a:r>',
        },
      ],
    });
    expect(both.paragraphs[0]!.bullet).toBe('1.');
  });

  it('buChar の記号は 1 文字（サロゲートペアを割らない）。空・無しは null', async () => {
    const t = await textOf({
      paras: [
        { runs: ['a'], bullet: { char: '●●' } },
        { runs: ['b'], bullet: { char: '😀' } },
        { runs: ['c'], bullet: { char: '' } },
        { runs: ['d'] },
      ],
    });
    expect(t.paragraphs.map((p) => p.bullet)).toEqual(['●', '😀', null, null]);
  });
});

describe('TD-87-6 本文：実体参照の復元・制御文字の除去・文字の位置は保つ', () => {
  it('a:t の実体参照（&amp; &lt; &gt; &quot; &apos; 数値参照）を戻す', async () => {
    const t = await textOf({
      paras: [{ runs: [{ rawText: 'a&amp;b &lt;c&gt; &#65;&#x42; q&quot;&apos;' }] }],
    });
    expect(flat(t)).toEqual([`a&b <c> AB q"'`]);
  });

  it('タブと改行以外の制御文字（C0・DEL・C1）を除く。空白は詰めない', async () => {
    const t = await textOf({
      paras: [{ runs: [{ rawText: ' x&#1;y&#9;z&#x7F;w\nv&#x85;u&#xB; ' }] }],
    });
    expect(flat(t)).toEqual([' xy\tzw\nvu ']);
  });

  it('制御文字だけの本文は、文字の無い図形になる', async () => {
    const [d] = await readShapes([{ name: 's', paras: [{ runs: [{ rawText: '&#1;&#2;' }] }] }]);
    expect(d!.text).toBeUndefined();
  });
});

describe('TD-87-7 既存の図形の読み取りに影響しない', () => {
  it('spXml（txBody なし）の図形は text を持たない。名前・説明・種類・アンカーに図形の項目（ISSUE-91）', async () => {
    const bytes = await buildXlsx({
      sheets: [
        {
          name: 'S',
          rows: [['a']],
          drawing: { xml: wsDr([twoCellAnchor(A1, B2, spXml('r', true))]), rels: [] },
        },
      ],
    });
    const { drawings } = await (await openXlsx(bytes)).openDrawings(1);
    expect(drawings).toEqual([
      {
        anchor: expect.anything(),
        kind: 'textbox',
        name: 'r',
        descr: '',
        mediaPath: null,
        rot: 0,
        flipH: false,
        flipV: false,
        hasOutline: true,
        hasFill: true,
        outline: { widthEmu: 0, dash: 'solid' },
        fillColor: { rgb: 0xffffff, alpha: 1 },
        outlineColor: { rgb: 0, alpha: 1 },
        geom: { type: 'preset', prst: 'rect', adj: [] },
      },
    ]);
  });
});
