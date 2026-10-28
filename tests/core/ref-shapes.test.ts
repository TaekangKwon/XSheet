/// <reference types="node" />
/**
 * 実際の Excel で作った基準ファイルによる、図形の文字と線・矢印の読み取りの照合（ISSUE-87・ADR-0026 §決定 1・3・U1）。
 * `fixtures/excel/ref-shapes.xlsx` は Excel 16 の COM（PowerShell・画面 125%）で作成・保存した実ファイル。
 * 1 シートに最上位 133 図形（線 67・テキストボックス 20・図形 40・グラフ 2・グループ 4）が入り、
 * うち `shapeHidden`（非表示）と `lnInvisible`（線なし）は一覧に出ない＝最上位は 131 件。
 * グループ 4 件は ISSUE-89 で中の子を展開する（グループ自身は一覧に出ない）：最上位 127 件＋子 11 件＝138 件を返す。
 * 期待値は 2 系統：
 * - COM の値（`Rotation`・`HorizontalFlip`＝-1 が反転・`VerticalFlip`・`Weight`＝pt・`DashStyle`・
 *   `*ArrowheadStyle/Length/Width`・`Adjustments`・`TextFrame2` の余白（pt）・向き・揃え・箇条書き）
 * - Excel が書いた drawing XML の値（`a:ln`・`a:prstGeom`・`a:bodyPr` の属性など）
 * 基準の JSON は作業用の場所にあり消えうるため、必要な期待値はこのファイルに直接書いた。
 */
import { readFileSync } from 'node:fs';
import { beforeAll, describe, expect, it } from 'vitest';
import { openXlsx, type SheetDrawing, type SheetDrawings } from '../../src/core/index';

/** 基準ファイルを開いて 1 枚目のシートの図を読む */
async function readRef(): Promise<SheetDrawings> {
  const book = await openXlsx(
    readFileSync(new URL('../../fixtures/excel/ref-shapes.xlsx', import.meta.url)),
  );
  return book.openDrawings(1);
}

/** 空白区切りの名前を配列にする（長い一覧を読みやすく書くため） */
const names = (s: string): string[] => s.split(/\s+/).filter((x) => x !== '');

let result: SheetDrawings;
let byName: Map<string, SheetDrawing>;

beforeAll(async () => {
  result = await readRef();
  byName = new Map(result.drawings.map((d) => [d.name, d]));
});

/** 名前で図を 1 つ取る（無ければテストを失敗させる） */
function get(name: string): SheetDrawing {
  const d = byName.get(name);
  if (!d) throw new Error(`図が無い: ${name}`);
  return d;
}

/** 文字の段落を [揃え, 箇条書き, 本文（改行は \n）] に畳んだもの */
function paras(name: string): [string, string | null, string][] {
  return get(name).text!.paragraphs.map((p) => [
    p.align,
    p.bullet,
    p.runs.map((r) => (r.br ? '\n' : r.text)).join(''),
  ]);
}

/** 線の名前（全 66 本。drawing XML の順） */
const LINE_NAMES = names(`
  lnPlain lnFlipV lnFlipH lnFlipHV lnRot90 lnRot180 lnRot270 lnFlipVRot90 lnFlipHRot180 lnHoriz
  lnVert lnRot45 bent3 bent3FlipV bent3FlipH bent3Adj bent3AdjNeg bent3AdjBig curve curveFlipV
  curveFlipH curveAdj curveRot90 arrow_tri arrow_open arrow_stealth arrow_diamond arrow_oval
  arrowMed_tri arrowMed_open arrowMed_stealth arrowMed_diamond arrowMed_oval arrowShortWide
  arrowLongNarrow arrowThick dash_solid dash_squareDot dash_roundDot dash_dash dash_dashDot
  dash_dashDotDot dash_longDash dash_longDashDot wThin wDefault w3pt w6pt w12pt linkElbow
  linkStraight linkCurve uRightRight uLeftLeft uTopTop uBottomBottom vElbow bentRightTop
  bentTopBottom bentLeftRight bentBack bent4 curved4 bent5 curved5 curved2
`);

/** テキストボックス（`cNvSpPr txBox`）の名前（全 20 個） */
const TEXTBOX_NAMES = names(`
  txtMulti txtStyle txtRuns txtJust txtDist label emptyTextbox textboxBordered noWrapTb
  autofitShrink autofitResize bullets numbered vertUp vertDown vertFarEast vertStack vertHorzRotFE
  rotTb45 srcTb
`);

/** 図形（txBox でない `xdr:sp`）の名前（全 39 個） */
const SHAPE_NAMES = names(`
  outer inner inner2 noLineRect noFillRect defaultRect emptyRect roundRect ellipse diamond
  rightArrowBlock wrapNone insets linkedText rot90 rot30 rot270 rot180 flipHText flipVText dstRect
  uA uB uC uD vTop vBottom wA wB pA_bent4 pB_bent4 pA_curved4 pB_curved4 pA_bent5 pB_bent5
  pA_curved5 pB_curved5 pA_curved2 pB_curved2
`);

/** 文字の無い図形（`a:endParaRPr` だけ。text を持たない）の名前（全 20 個） */
const NO_TEXT_NAMES = names(`
  emptyRect emptyTextbox uA uB uC uD vTop vBottom wA wB pA_bent4 pB_bent4 pA_curved4 pB_curved4
  pA_bent5 pB_bent5 pA_curved5 pB_curved5 pA_curved2 pB_curved2
`);

/** 文字のある図形・テキストボックスの名前（全 39 個） */
const TEXT_NAMES = [...TEXTBOX_NAMES, ...SHAPE_NAMES].filter((n) => !NO_TEXT_NAMES.includes(n));

/** グループの子（ISSUE-89 で展開。XML の順）：gConn・hLine は線、gTb は文字入りのテキストボックス、残りは図形 */
const GROUP_CHILD_NAMES = names(`
  gTb gRect gConn hRect hOval hLine iRect iRect2 nRect1 nRect2 nRect3
`);

describe('TD-87-24 ref-shapes.xlsx：全図形の種類と件数', () => {
  it('138 件（線 68・テキストボックス 21・図形 47・グラフ 2）。グループ自身は出ず、子 11 件に展開。通知は無い。順序は drawing XML の順', () => {
    expect(result.notices).toEqual([]);
    expect(result.drawings).toHaveLength(138);
    const kindOf = (k: string): string[] =>
      result.drawings.filter((d) => d.kind === k).map((d) => d.name);
    expect(kindOf('line')).toEqual([...LINE_NAMES, 'gConn', 'hLine']);
    expect(kindOf('line')).toHaveLength(68);
    expect(kindOf('textbox').sort()).toEqual([...TEXTBOX_NAMES, 'gTb'].sort());
    expect(kindOf('shape').sort()).toEqual(
      [...SHAPE_NAMES, ...names('gRect hRect hOval iRect iRect2 nRect1 nRect2 nRect3')].sort(),
    );
    expect(kindOf('chart')).toEqual(['chartBar', 'chartexWaterfall']);
    expect(kindOf('group')).toEqual([]);
    expect(kindOf('image')).toEqual([]);
    expect(kindOf('smartart')).toEqual([]);
    // 重なり順：線 → 図形（文字・接続元／先）→ グラフ → グループの子。drawing XML の順のまま
    expect(result.drawings[0]!.name).toBe('lnPlain');
    expect(result.drawings.slice(-13).map((d) => d.name)).toEqual([
      'chartBar',
      'chartexWaterfall',
      ...GROUP_CHILD_NAMES,
    ]);
  });

  it('非表示の図形（shapeHidden・cNvPr hidden="1"）と線なしの線（lnInvisible・a:noFill）は出ない', () => {
    expect(byName.has('shapeHidden')).toBe(false);
    expect(byName.has('lnInvisible')).toBe(false);
    // 出た 131 件の中に、隠した図形の文字（"hidden"）も無い
    const texts = result.drawings.flatMap((d) =>
      (d.text?.paragraphs ?? []).flatMap((p) => p.runs.map((r) => r.text)),
    );
    expect(texts).not.toContain('hidden');
  });

  it('アンカー：線・グラフ・グループの子は twoCell。from・to の列・行・ずれ（EMU）が drawing XML のとおり（子のアンカーは最も外側のグループのもの）', () => {
    const pt = (col: number, colOff: number, row: number, rowOff: number) => ({
      col,
      colOff,
      row,
      rowOff,
    });
    const anchorOf = (n: string) => get(n).anchor;
    expect(anchorOf('lnPlain')).toEqual({
      type: 'twoCell',
      from: pt(0, 254000, 8, 60960),
      to: pt(2, 50800, 10, 76200),
    });
    expect(anchorOf('chartBar')).toEqual({
      type: 'twoCell',
      from: pt(0, 254000, 96, 33020),
      to: pt(6, 152400, 108, 124460),
    });
    expect(anchorOf('chartexWaterfall')).toEqual({
      type: 'twoCell',
      from: pt(6, 533400, 96, 33020),
      to: pt(13, 76200, 108, 124460),
    });
    const grpScaled = {
      type: 'twoCell',
      from: pt(0, 254000, 111, 83820),
      to: pt(6, 25400, 114, 170180),
    };
    const grpPlain = {
      type: 'twoCell',
      from: pt(6, 406400, 111, 83820),
      to: pt(9, 355600, 115, 114300),
    };
    const grpRot = {
      type: 'twoCell',
      from: pt(10, 508000, 111, 83820),
      to: pt(13, 457200, 114, 170180),
    };
    const grpOuter = {
      type: 'twoCell',
      from: pt(15, 0, 111, 83820),
      to: pt(17, 177800, 117, 2540),
    };
    for (const n of ['gTb', 'gRect', 'gConn']) expect(anchorOf(n), n).toEqual(grpScaled);
    for (const n of ['hRect', 'hOval', 'hLine']) expect(anchorOf(n), n).toEqual(grpPlain);
    for (const n of ['iRect', 'iRect2']) expect(anchorOf(n), n).toEqual(grpRot);
    for (const n of ['nRect1', 'nRect2', 'nRect3']) expect(anchorOf(n), n).toEqual(grpOuter);
  });
});

describe('TD-87-25 ref-shapes.xlsx：文字（段落・ラン・改行・書式・箇条書き・参照）', () => {
  it('文字の無い図形・テキストボックス（emptyRect・emptyTextbox・a:endParaRPr だけの 18 図形）は text を持たない', () => {
    expect(NO_TEXT_NAMES).toHaveLength(20);
    for (const n of NO_TEXT_NAMES) {
      const d = get(n);
      expect(d.text, n).toBeUndefined();
      // 文字が無い図形も回転・反転・枠線と塗りの有無・形・枠線の太さと線種を持つ（ISSUE-91）
      expect(d.rot, n).toBeDefined();
      expect(d.hasOutline, n).toBeDefined();
      expect(d.hasFill, n).toBeDefined();
      expect(d.geom, n).toBeDefined();
      expect(d.outline, n).toBeDefined();
    }
    // 文字のある図形は 39 個で、全部 text を持つ
    expect(TEXT_NAMES).toHaveLength(39);
    for (const n of TEXT_NAMES) expect(get(n).text, n).toBeDefined();
  });

  it('txtMulti：複数段落・a:br・空の段落・中央揃え・縦位置 中央・太字 14pt（COM の Text は "Title\\nline2\\vbrk\\n\\nlast"）', () => {
    expect(paras('txtMulti')).toEqual([
      ['ctr', null, 'Title'],
      ['l', null, 'line2\nbrk'],
      ['l', null, ''],
      ['l', null, 'last'],
    ]);
    const t = get('txtMulti').text!;
    expect(t.anchor).toBe('ctr'); // COM VerticalAnchor=3（中央）
    // 2 つ目の段落は a:br を挟んだ 3 つの要素（ラン・改行・ラン）
    expect(t.paragraphs[1]!.runs.map((r) => r.br)).toEqual([false, true, false]);
    expect(t.paragraphs[2]!.runs).toEqual([]); // 空の段落
    const title = t.paragraphs[0]!.runs[0]!;
    expect([title.sizePt, title.bold, title.italic]).toEqual([14, true, false]);
    expect(t.paragraphs[1]!.runs.map((r) => r.sizePt)).toEqual([11, 11, 11]);
  });

  it('txtStyle：斜体・下線・取り消し線・8pt（2 つのラン）・24pt。右揃え・縦位置 下（COM Alignment=3・VerticalAnchor=4）', () => {
    const t = get('txtStyle').text!;
    expect(t.anchor).toBe('b');
    expect(paras('txtStyle')).toEqual([
      ['r', null, 'italic'],
      ['r', null, 'under'],
      ['r', null, 'strike'],
      ['r', null, '日本語 8pt'],
      ['r', null, 'big24'],
    ]);
    const flags = (i: number) => {
      const r = t.paragraphs[i]!.runs[0]!;
      return [r.italic, r.underline, r.strike, r.bold, r.sizePt];
    };
    expect(flags(0)).toEqual([true, false, false, false, 11]);
    expect(flags(1)).toEqual([false, true, false, false, 11]);
    expect(flags(2)).toEqual([false, false, true, false, 11]);
    // "日本語 " と "8pt" は言語（ja-JP／en-US）が違うので 2 つのラン。どちらも 8pt
    expect(t.paragraphs[3]!.runs.map((r) => [r.text, r.sizePt])).toEqual([
      ['日本語 ', 8],
      ['8pt', 8],
    ]);
    expect(t.paragraphs[4]!.runs[0]!.sizePt).toBe(24);
    // 最後の段落の段落記号（a:endParaRPr sz=2400）。ほかの段落は既定の 11
    expect(t.paragraphs.map((p) => p.endSizePt)).toEqual([11, 11, 11, 11, 24]);
  });

  it('txtRuns：1 段落に 3 つのラン（"AAA" 太字 11pt・" bbb " 11pt・"CCC" 20pt）。段落記号は 20pt', () => {
    const p = get('txtRuns').text!.paragraphs;
    expect(p).toHaveLength(1);
    expect(p[0]!.runs.map((r) => [r.text, r.sizePt, r.bold])).toEqual([
      ['AAA', 11, true],
      [' bbb ', 11, false],
      ['CCC', 20, false],
    ]);
    expect(p[0]!.endSizePt).toBe(20);
  });

  it('txtJust・txtDist：両端揃え（COM Alignment=4）・均等割り付け（5）。txtJust は余白 0', () => {
    expect(paras('txtJust')).toEqual([['just', null, 'justify justify justify justify justify']]);
    expect(get('txtJust').text!.insets).toEqual({ l: 0, t: 0, r: 0, b: 0 });
    expect(paras('txtDist')).toEqual([['dist', null, 'dist']]);
  });

  it('bullets・numbered：buChar は "•"、buAutoNum は startAt=3 から 3.・4.・5.。通常の段落を挟むと付かない', () => {
    expect(paras('bullets')).toEqual([
      ['l', '•', 'one'],
      ['l', '•', 'two'],
      ['l', '•', 'three'],
    ]);
    expect(paras('numbered')).toEqual([
      ['l', '3.', 'first'],
      ['l', '4.', 'second'],
      ['l', '5.', 'third'],
      ['l', null, 'plain'],
    ]);
  });

  it('linkedText：セル参照の文字（a:fld type="TxLink"）の表示文字が本文になる', () => {
    expect(paras('linkedText')).toEqual([['l', null, 'linked text']]);
  });

  it('入れ子の四角形・日本語・既定の文字（outer・inner・inner2・default 系）は 11pt・左揃え・縦位置 上', () => {
    expect(paras('outer')).toEqual([['l', null, 'システム']]);
    expect(paras('inner')).toEqual([['l', null, 'サーバ']]);
    expect(paras('inner2')).toEqual([['l', null, 'DB']]);
    expect(paras('label')).toEqual([['l', null, 'HTTPS']]);
    expect(paras('srcTb')).toEqual([['l', null, 'client']]);
    expect(paras('dstRect')).toEqual([['l', null, 'API']]);
    expect(paras('textboxBordered')).toEqual([['l', null, 'tb border']]);
    for (const n of ['outer', 'inner', 'inner2', 'defaultRect', 'ellipse', 'diamond']) {
      const t = get(n).text!;
      expect(t.anchor, n).toBe('t');
      expect(t.paragraphs[0]!.runs[0]!.sizePt, n).toBe(11);
      expect(t.fontScale, n).toBe(1);
    }
  });

  it('縦書き 5 種：COM Orientation 2〜6 ⇔ vert270・vert・eaVert・mongolianVert・wordArtVertRtl。横書きは horz', () => {
    const vertOf = (n: string) => get(n).text!.vert;
    expect(vertOf('vertUp')).toBe('vert270');
    expect(vertOf('vertDown')).toBe('vert');
    expect(vertOf('vertFarEast')).toBe('eaVert');
    expect(vertOf('vertStack')).toBe('mongolianVert');
    expect(vertOf('vertHorzRotFE')).toBe('wordArtVertRtl');
    for (const n of ['vertUp', 'vertDown', 'vertFarEast', 'vertStack', 'vertHorzRotFE']) {
      // 5 つとも日本語と英字の 2 つのラン（"ABC"・"縦書き"）
      expect(paras(n), n).toEqual([['l', null, 'ABC縦書き']]);
    }
    for (const n of ['txtMulti', 'outer', 'wrapNone', 'insets', 'rot30']) {
      expect(vertOf(n), n).toBe('horz');
    }
  });

  it('折り返し：wrapNone・noWrapTb は noWrap（COM WordWrap=0）。ほかは折り返す', () => {
    const noWrap = TEXT_NAMES.filter((n) => get(n).text!.noWrap);
    expect(noWrap.sort()).toEqual(['noWrapTb', 'wrapNone']);
    expect(paras('wrapNone')).toEqual([['l', null, 'a long text without wrap']]);
    expect(paras('noWrapTb')).toEqual([['l', null, 'a long text without wrap']]);
  });

  it('自動調整：autofitShrink（normAutofit・fontScale なし）は縮小の倍率 1。autofitResize（spAutoFit）も 1', () => {
    expect(get('autofitShrink').text!.fontScale).toBe(1);
    expect(paras('autofitShrink')).toEqual([['l', null, 'shrink this long text please']]);
    expect(get('autofitResize').text!.fontScale).toBe(1);
    expect(paras('autofitResize')).toEqual([['l', null, 'resize shape']]);
  });

  it('内側の余白（EMU）：COM MarginLeft/Top/Right/Bottom（pt）×12700 と一致。既定は 7.2・3.6pt＝91440・45720', () => {
    const insetsOf = (n: string) => get(n).text!.insets;
    // COM [20, 10, 5, 2] pt
    expect(insetsOf('insets')).toEqual({ l: 254000, t: 127000, r: 63500, b: 25400 });
    // COM [0, 0, 7.2, 3.6] pt（lIns・tIns だけ 0 を指定。rIns・bIns は既定）
    expect(insetsOf('roundRect')).toEqual({ l: 0, t: 0, r: 91440, b: 45720 });
    // COM [0, 0, 0, 0] pt
    expect(insetsOf('txtJust')).toEqual({ l: 0, t: 0, r: 0, b: 0 });
    // 既定（COM 7.2・3.6・7.2・3.6）
    for (const n of TEXT_NAMES.filter((x) => !['insets', 'roundRect', 'txtJust'].includes(x))) {
      expect(insetsOf(n), n).toEqual({ l: 91440, t: 45720, r: 91440, b: 45720 });
    }
  });

  it('縦位置：中央 = txtMulti、下 = txtStyle・roundRect、そのほかは上（COM VerticalAnchor 3・4・1）', () => {
    const anchors = TEXT_NAMES.filter((n) => get(n).text!.anchor !== 't').sort();
    expect(anchors).toEqual(['roundRect', 'txtMulti', 'txtStyle']);
    expect(get('roundRect').text!.anchor).toBe('b');
    expect(paras('roundRect')).toEqual([['r', null, 'R']]);
  });

  it('Excel が書いた a:rPr には sz がある。lstStyle は空なので、段落・ランの大きさは 11 か指定値で、丸めは起きない', () => {
    const sizes = new Set<number>();
    for (const n of TEXT_NAMES) {
      for (const p of get(n).text!.paragraphs) {
        sizes.add(p.endSizePt);
        for (const r of p.runs) sizes.add(r.sizePt);
      }
    }
    expect([...sizes].sort((a, b) => a - b)).toEqual([8, 11, 14, 20, 24]);
  });
});

describe('TD-87-26 ref-shapes.xlsx：枠線と塗りの有無', () => {
  it('label（線なし・塗りなし）・noLineRect（線なし）・noFillRect（塗りなし）・defaultRect・textboxBordered', () => {
    const of = (n: string) => [get(n).hasOutline, get(n).hasFill];
    expect(of('label')).toEqual([false, false]); // COM Line.Visible=0・FillVisible=0
    expect(of('noLineRect')).toEqual([false, true]); // Line.Visible=0・FillVisible=-1
    expect(of('noFillRect')).toEqual([true, false]); // Line.Visible=-1・FillVisible=0
    expect(of('defaultRect')).toEqual([true, true]);
    expect(of('textboxBordered')).toEqual([true, true]);
  });

  it('上記以外の文字のある図形は、すべて枠線あり・塗りあり（COM Line.Visible・FillVisible とも -1）', () => {
    const special = ['label', 'noLineRect', 'noFillRect'];
    for (const n of TEXT_NAMES.filter((x) => !special.includes(x))) {
      expect([get(n).hasOutline, get(n).hasFill], n).toEqual([true, true]);
    }
  });
});

describe('TD-87-27 ref-shapes.xlsx：回転と反転（U1）が COM の Rotation・HorizontalFlip・VerticalFlip と一致', () => {
  /** [名前, COM Rotation（度）, COM HorizontalFlip=-1, COM VerticalFlip=-1]。表に無い文字付き図形・線は 0・反転なし */
  const ROT_TABLE: [string, number, boolean, boolean][] = [
    ['lnFlipV', 0, false, true],
    ['lnFlipH', 0, true, false],
    ['lnFlipHV', 0, true, true],
    ['lnRot90', 90, false, false],
    ['lnRot180', 180, false, false],
    ['lnRot270', 270, false, false],
    ['lnFlipVRot90', 90, false, true],
    ['lnFlipHRot180', 180, true, false],
    ['lnRot45', 45, false, false],
    ['bent3FlipV', 0, false, true],
    ['bent3FlipH', 0, true, false],
    ['curveFlipV', 0, false, true],
    ['curveFlipH', 0, true, false],
    ['curveRot90', 90, false, false],
    ['rot90', 90, false, false],
    ['rot30', 30, false, false],
    ['rot270', 270, false, false],
    ['rot180', 180, false, false],
    ['rotTb45', 45, false, false],
    ['flipHText', 0, true, false],
    ['flipVText', 0, false, true],
    // 接続した接続線（Excel が接続先の向きに合わせて rot と flip を書く）
    ['linkElbow', 270, true, false],
    ['linkCurve', 270, true, false],
    ['uLeftLeft', 180, false, true],
    ['uTopTop', 90, true, true],
    ['uBottomBottom', 270, true, false],
    ['vElbow', 270, true, false],
    ['bentTopBottom', 270, false, true],
    ['bentLeftRight', 180, true, false],
    ['bentBack', 0, true, false],
    ['bent4', 270, true, false],
    ['curved4', 270, true, false],
    ['bent5', 270, true, false],
    ['curved5', 270, true, false],
    ['curved2', 270, true, false],
  ];

  it('回転・反転のある 35 個：rot（1/60000 度）= COM Rotation×60000・flipH／flipV = COM の Flip が -1', () => {
    expect(ROT_TABLE).toHaveLength(35);
    for (const [n, deg, h, v] of ROT_TABLE) {
      const d = get(n);
      expect([d.rot, d.flipH, d.flipV], n).toEqual([deg * 60000, h, v]);
    }
  });

  it('上の表に無い 70 個（線と文字のある図形）は、回転 0・反転なし（rot は 0 で持つ）', () => {
    const tabled = new Set(ROT_TABLE.map((r) => r[0]));
    const rest = [...LINE_NAMES, ...TEXT_NAMES].filter((n) => !tabled.has(n));
    expect(rest).toHaveLength(70);
    for (const n of rest) {
      const d = get(n);
      expect([d.rot, d.flipH, d.flipV], n).toEqual([0, false, false]);
    }
  });

  it('線の 8 つ（反転 3・回転 3・反転＋回転 2）：上下・左右・両方の反転と 90°・180°・270° の組み合わせが区別できる', () => {
    const v = (n: string) => [get(n).rot! / 60000, get(n).flipH, get(n).flipV];
    expect(v('lnFlipV')).toEqual([0, false, true]);
    expect(v('lnFlipH')).toEqual([0, true, false]);
    expect(v('lnFlipHV')).toEqual([0, true, true]);
    expect(v('lnFlipVRot90')).toEqual([90, false, true]);
    expect(v('lnFlipHRot180')).toEqual([180, true, false]);
    expect(v('lnRot90')).toEqual([90, false, false]);
    expect(v('lnRot180')).toEqual([180, false, false]);
    expect(v('lnRot270')).toEqual([270, false, false]);
  });
});

describe('TD-87-28 ref-shapes.xlsx：線の形（geom）・調整値（adj）・太さ', () => {
  it('geom：line・straightConnector1・bentConnector2〜5・curvedConnector2〜5（Excel が書いた prst のとおり）', () => {
    const geomOf = (n: string) => get(n).line!.geom;
    // 直線（cxnSp の prst="line"）。lnPlain・lnFlip* などは straightConnector1
    for (const n of names(
      'lnRot90 lnRot180 lnRot270 lnHoriz lnVert lnRot45 dash_solid dash_squareDot dash_roundDot dash_dash dash_dashDot dash_dashDotDot dash_longDash dash_longDashDot wThin wDefault w3pt w6pt w12pt',
    )) {
      expect(geomOf(n), n).toBe('line');
    }
    for (const n of names(
      'lnPlain lnFlipV lnFlipH lnFlipHV lnFlipVRot90 lnFlipHRot180 arrow_tri arrow_open arrow_stealth arrow_diamond arrow_oval arrowMed_tri arrowMed_open arrowMed_stealth arrowMed_diamond arrowMed_oval arrowShortWide arrowLongNarrow arrowThick linkStraight',
    )) {
      expect(geomOf(n), n).toBe('straightConnector1');
    }
    for (const n of names(
      'bent3 bent3FlipV bent3FlipH bent3Adj bent3AdjNeg bent3AdjBig uRightRight uLeftLeft uTopTop uBottomBottom vElbow bentTopBottom',
    )) {
      expect(geomOf(n), n).toBe('bentConnector3');
    }
    for (const n of ['linkElbow', 'bentRightTop']) expect(geomOf(n), n).toBe('bentConnector2');
    expect(geomOf('bent4')).toBe('bentConnector4');
    for (const n of ['bentLeftRight', 'bentBack', 'bent5']) {
      expect(geomOf(n), n).toBe('bentConnector5');
    }
    for (const n of names('curve curveFlipV curveFlipH curveAdj curveRot90 linkCurve')) {
      expect(geomOf(n), n).toBe('curvedConnector3');
    }
    expect(geomOf('curved2')).toBe('curvedConnector2');
    expect(geomOf('curved4')).toBe('curvedConnector4');
    expect(geomOf('curved5')).toBe('curvedConnector5');
  });

  it('adj：a:avLst の adj1〜adj3（COM Adjustments×100000 と一致）。無ければ 50000', () => {
    const adjOf = (n: string) => get(n).line!.adj;
    expect(adjOf('bent3Adj')).toEqual([30000, 50000, 50000]); // COM 0.3
    expect(adjOf('bent3AdjNeg')).toEqual([-40000, 50000, 50000]); // COM -0.4
    expect(adjOf('bent3AdjBig')).toEqual([160000, 50000, 50000]); // COM 1.6
    expect(adjOf('curveAdj')).toEqual([20000, 50000, 50000]); // COM 0.2
    expect(adjOf('uRightRight')).toEqual([145000, 50000, 50000]); // COM 1.45
    expect(adjOf('linkCurve')).toEqual([50000, 50000, 50000]); // 明示の val 50000
    expect(adjOf('vElbow')).toEqual([50000, 50000, 50000]);
    expect(adjOf('bentLeftRight')).toEqual([-90000, 50000, 190000]); // COM -0.9・0.5・1.9
    expect(adjOf('bentBack')).toEqual([-90000, 50000, 190000]);
    expect(adjOf('bent4')).toEqual([-18947, 65000, 50000]); // COM -0.18947・0.65
    expect(adjOf('curved4')).toEqual([-18947, 65000, 50000]);
    expect(adjOf('bent5')).toEqual([-16364, 50000, 116364]); // COM -0.16364・0.5・1.16364
    expect(adjOf('curved5')).toEqual([-16364, 50000, 116364]);
    // adj の無い線は、COM が 0.5 を返すもの（bent3 系・curve 系）も含めて 50000
    for (const n of names(
      'bent3 bent3FlipV bent3FlipH curve curveFlipV curveFlipH curveRot90 bentTopBottom',
    )) {
      expect(adjOf(n), n).toEqual([50000, 50000, 50000]);
    }
  });

  it('U 字（uLeftLeft・uTopTop・uBottomBottom）：端の揃った形で Excel が書く val 1800000（COM 18.0）をそのまま読む（丸めない）', () => {
    // adj は 32bit 整数の範囲に収めるだけ（値を有限に保つため）。幅 12,700 EMU の線にこの値を掛けて引くので、丸めると形が変わる
    for (const n of ['uLeftLeft', 'uTopTop', 'uBottomBottom']) {
      expect(get(n).line!.adj, n).toEqual([1800000, 50000, 50000]);
    }
  });

  it('太さ（EMU）：a:ln w があればその値（COM Weight×12700）。無い線は xdr:style lnRef idx=1 → 6350（0.5pt）', () => {
    const widthOf = (n: string) => get(n).line!.widthEmu;
    expect(widthOf('wThin')).toBe(3175); // COM 0.25pt
    expect(widthOf('w3pt')).toBe(38100);
    expect(widthOf('w6pt')).toBe(76200);
    expect(widthOf('w12pt')).toBe(152400);
    expect(widthOf('arrowThick')).toBe(63500); // COM 5pt
    // wDefault は a:ln 自体が無く、lnRef idx=1 の 0.5pt（COM Weight=0.5）
    expect(widthOf('wDefault')).toBe(6350);
    for (const n of names(
      'dash_solid dash_squareDot dash_roundDot dash_dash dash_dashDot dash_dashDotDot dash_longDash dash_longDashDot',
    )) {
      expect(widthOf(n), n).toBe(25400); // COM 2pt
    }
    const explicit = new Set(
      names(
        'wThin w3pt w6pt w12pt arrowThick dash_solid dash_squareDot dash_roundDot dash_dash dash_dashDot dash_dashDotDot dash_longDash dash_longDashDot',
      ),
    );
    // 残りの 53 本は a:ln w を持たず、lnRef idx=1 の 6350（COM Weight=0.5）
    const rest = LINE_NAMES.filter((n) => !explicit.has(n));
    expect(rest).toHaveLength(53);
    for (const n of rest) expect(widthOf(n), n).toBe(6350);
  });
});

describe('TD-87-29 ref-shapes.xlsx：矢じり（種類・大きさ）と線種', () => {
  type End = { type: string; w: string; len: string } | null;
  const end = (type: string, w: string, len: string): End => ({ type, w, len });
  const med = (type: string): End => end(type, 'med', 'med');

  it('矢じり 5 種×大小：始点は小（sm・sm）、終点は大（lg・lg）。COM の Style 2〜6 ⇔ triangle・arrow・stealth・diamond・oval', () => {
    const cases: [string, string][] = [
      ['arrow_tri', 'triangle'], // COM Style=2
      ['arrow_open', 'arrow'], // 3（開いた矢じり）
      ['arrow_stealth', 'stealth'], // 4
      ['arrow_diamond', 'diamond'], // 5
      ['arrow_oval', 'oval'], // 6
    ];
    for (const [n, type] of cases) {
      const l = get(n).line!;
      expect(l.head, n).toEqual(end(type, 'sm', 'sm')); // COM BeginArrowheadLength=1・Width=1
      expect(l.tail, n).toEqual(end(type, 'lg', 'lg')); // COM EndArrowheadLength=3・Width=3
    }
  });

  it('終点だけ中サイズの矢じり 5 種（arrowMed_*）：始点は無し（COM BeginArrowheadStyle=1）', () => {
    const cases: [string, string][] = [
      ['arrowMed_tri', 'triangle'],
      ['arrowMed_open', 'arrow'],
      ['arrowMed_stealth', 'stealth'],
      ['arrowMed_diamond', 'diamond'],
      ['arrowMed_oval', 'oval'],
    ];
    for (const [n, type] of cases) {
      const l = get(n).line!;
      expect(l.head, n).toBeNull();
      expect(l.tail, n).toEqual(med(type));
    }
  });

  it('大きさの組み合わせ：arrowShortWide は幅 lg・長さ sm（COM Length=1・Width=3）、arrowLongNarrow は幅 sm・長さ lg', () => {
    expect(get('arrowShortWide').line!.tail).toEqual(end('triangle', 'lg', 'sm'));
    expect(get('arrowLongNarrow').line!.tail).toEqual(end('triangle', 'sm', 'lg'));
    expect(get('arrowShortWide').line!.head).toBeNull();
    expect(get('arrowLongNarrow').line!.head).toBeNull();
    expect(get('arrowThick').line!.tail).toEqual(med('triangle'));
  });

  it('curve 系は終点が stealth（COM EndArrowheadStyle=4）。接続線・カギ線・曲線の多くは終点が三角（COM 2）で始点は無し', () => {
    for (const n of names('curve curveFlipV curveFlipH curveAdj')) {
      expect(get(n).line!.tail, n).toEqual(med('stealth'));
      expect(get(n).line!.head, n).toBeNull();
    }
    for (const n of names(
      'bent3 bent3FlipV bent3FlipH bent3Adj bent3AdjNeg bent3AdjBig linkElbow linkStraight linkCurve uRightRight uLeftLeft uTopTop uBottomBottom vElbow bentRightTop bentTopBottom bentLeftRight bentBack bent4 curved4 bent5 curved5 curved2',
    )) {
      expect(get(n).line!.tail, n).toEqual(med('triangle'));
      expect(get(n).line!.head, n).toBeNull();
    }
  });

  it('矢じりの無い線（a:headEnd・a:tailEnd が無い。COM Style=1）は head・tail とも null', () => {
    const none = names(`
      lnPlain lnFlipV lnFlipH lnFlipHV lnRot90 lnRot180 lnRot270 lnFlipVRot90 lnFlipHRot180 lnHoriz
      lnVert lnRot45 curveRot90 dash_solid dash_squareDot dash_roundDot dash_dash dash_dashDot
      dash_dashDotDot dash_longDash dash_longDashDot wThin wDefault w3pt w6pt w12pt
    `);
    expect(none).toHaveLength(26);
    for (const n of none) {
      expect(get(n).line!.head, n).toBeNull();
      expect(get(n).line!.tail, n).toBeNull();
    }
    // 残り 40 本は、どれかの端に矢じりがある
    const withArrow = LINE_NAMES.filter((n) => get(n).line!.head || get(n).line!.tail);
    expect(withArrow).toHaveLength(66 - 26);
  });

  it('線種 8 種：COM DashStyle 1〜8 ⇔ solid・dot・dot・dash・dashDot・sysDashDotDot・lgDash・lgDashDot（cap の違いは読まない）', () => {
    const dashOf = (n: string) => get(n).line!.dash;
    expect(dashOf('dash_solid')).toBe('solid'); // 1
    expect(dashOf('dash_squareDot')).toBe('dot'); // 2（cap="sq"）
    expect(dashOf('dash_roundDot')).toBe('dot'); // 3（cap="rnd"）
    expect(dashOf('dash_dash')).toBe('dash'); // 4
    expect(dashOf('dash_dashDot')).toBe('dashDot'); // 5
    expect(dashOf('dash_dashDotDot')).toBe('sysDashDotDot'); // 6（Excel は sysDashDotDot を書く）
    expect(dashOf('dash_longDash')).toBe('lgDash'); // 7
    expect(dashOf('dash_longDashDot')).toBe('lgDashDot'); // 8
    expect(dashOf('uRightRight')).toBe('dash'); // 接続線の破線（COM DashStyle=4）
    const dashed = new Set(
      names(
        'dash_squareDot dash_roundDot dash_dash dash_dashDot dash_dashDotDot dash_longDash dash_longDashDot uRightRight',
      ),
    );
    for (const n of LINE_NAMES.filter((x) => !dashed.has(x))) expect(dashOf(n), n).toBe('solid');
  });
});

describe('TD-87-30 ref-shapes.xlsx：グラフ・グループ（ISSUE-89 で子に展開）', () => {
  it('chartBar は通常のグラフ（chart）。名前・アンカー以外の情報は持たない（mediaPath は null）', () => {
    const d = get('chartBar');
    expect(d.kind).toBe('chart');
    expect(d.mediaPath).toBeNull();
    expect(d.text).toBeUndefined();
    expect(d.line).toBeUndefined();
  });

  it('chartexWaterfall（mc:AlternateContent の Choice の graphicFrame）は「グラフ」になり、Fallback の sp（"この図は…"）は出ない', () => {
    const d = get('chartexWaterfall');
    expect(d.kind).toBe('chart');
    expect(d.text).toBeUndefined();
    // Fallback の文字も、名前が空の図も出ない（name="" の図は 0 個）
    expect(result.drawings.filter((x) => x.name === '')).toEqual([]);
    const allText = result.drawings.flatMap((x) =>
      (x.text?.paragraphs ?? []).flatMap((p) => p.runs.map((r) => r.text)),
    );
    expect(allText.some((t) => t.includes('この図は'))).toBe(false);
    // グラフは 2 つだけ（Fallback の sp が shape に紛れていない）
    expect(result.drawings.filter((x) => x.kind === 'chart')).toHaveLength(2);
  });

  it('グループ 4 件は展開され、グループ自身（grpScaled・grpPlain・grpRot・grpOuter・grpInner）は一覧にも件数にも出ない', () => {
    for (const n of ['grpScaled', 'grpPlain', 'grpRot', 'grpOuter', 'grpInner']) {
      expect(byName.has(n), n).toBe(false);
    }
    expect(result.drawings.filter((x) => x.kind === 'group')).toEqual([]);
    for (const n of GROUP_CHILD_NAMES) expect(byName.has(n), n).toBe(true);
  });

  it('子の種類：gTb はテキストボックス（文字 gtext）、gConn・hLine は終点に矢じりの線、残りは図形（文字あり）', () => {
    expect(get('gTb')).toMatchObject({ kind: 'textbox' });
    expect(paras('gTb')).toEqual([['l', null, 'gtext']]);
    expect(get('gRect').kind).toBe('shape');
    for (const n of ['gConn', 'hLine']) {
      const d = get(n);
      expect(d.kind, n).toBe('line');
      expect(d.line?.tail, n).not.toBeNull();
      expect(d.line?.head, n).toBeNull();
    }
    expect(paras('nRect3')).toEqual([['l', null, 'n3']]);
  });
});

/** 矩形（EMU：x, y, cx, cy） */
const box = (x: number, y: number, cx: number, cy: number) => ({ x, y, cx, cy });

describe('TD-89-12 ref-shapes.xlsx：グループの子の座標系（drawing XML の値。EMU）', () => {
  it('grpScaled：ext の幅 3429000 ÷ chExt の幅 2286000 ＝ 1.5 倍に拡大。子の xfrm は子の座標系のまま', () => {
    const levels = [{ ch: box(254000, 20383500, 2286000, 635000), own: null }];
    expect(get('gTb').group).toEqual({ levels, xfrm: box(254000, 20383500, 762000, 254000) });
    expect(get('gRect').group).toEqual({ levels, xfrm: box(1778000, 20764500, 762000, 254000) });
    expect(get('gConn').group).toEqual({ levels, xfrm: box(1016000, 20510500, 762000, 381000) });
  });

  it('grpPlain：chOff・chExt が off・ext と同じ（拡大なし）。線 hLine の高さは 0', () => {
    const levels = [{ ch: box(4064000, 20383500, 1778000, 762000), own: null }];
    expect(get('hRect').group).toEqual({ levels, xfrm: box(4064000, 20383500, 762000, 381000) });
    expect(get('hOval').group).toEqual({ levels, xfrm: box(5080000, 20637500, 762000, 381000) });
    expect(get('hLine').group).toEqual({ levels, xfrm: box(4064000, 21145500, 1778000, 0) });
  });

  it('grpRot：グループ自身の回転（rot=1800000）は読まない。子は回転前の xfrm のまま、子自身の回転も無い', () => {
    const levels = [{ ch: box(6604000, 20383500, 1778000, 635000), own: null }];
    expect(get('iRect').group).toEqual({ levels, xfrm: box(6604000, 20383500, 762000, 381000) });
    expect(get('iRect2').group).toEqual({ levels, xfrm: box(7620000, 20637500, 762000, 381000) });
    expect(get('iRect').rot).toBe(0);
  });

  it('grpOuter：入れ子（grpInner の中の nRect1・nRect2、外側の直下の nRect3）。内側の own は外側の座標系での位置', () => {
    const outer = { ch: box(9144000, 20383500, 1397000, 1016000), own: null };
    const inner = {
      ch: box(9144000, 20383500, 1397000, 381000),
      own: box(9144000, 20383500, 1397000, 381000),
    };
    expect(get('nRect1').group).toEqual({
      levels: [outer, inner],
      xfrm: box(9144000, 20383500, 635000, 381000),
    });
    expect(get('nRect2').group).toEqual({
      levels: [outer, inner],
      xfrm: box(9906000, 20383500, 635000, 381000),
    });
    expect(get('nRect3').group).toEqual({
      levels: [outer],
      xfrm: box(9144000, 21018500, 1397000, 381000),
    });
  });

  it('グループの外の図は group を持たない', () => {
    for (const n of ['lnPlain', 'chartBar', 'txtMulti', 'outer']) {
      expect(get(n).group, n).toBeUndefined();
    }
  });
});
