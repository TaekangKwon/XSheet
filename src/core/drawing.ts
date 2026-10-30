// 正: InterHub docs/plan/adr/0026-xlsx-image-overlay.md
/**
 * シートの図（画像・グラフ・図形）の読み取り：`<drawing r:id>` → シートの rels → drawingN.xml → drawing の rels → media。
 * DOMParser に依存せず、正規表現と `indexOf` で線形時間に走査する。
 * 開きタグは `[^<>]*>`（`>` が無くても次の `<` で打ち切る）、閉じタグは `indexOf`（`findElement`）か、入れ子を数える一度きりの走査（`findBalanced`）で探す。
 * 閉じの無い要素や `>` の無い開きタグを大量に並べた細工でも、時間は XML の大きさに比例する。
 * 外部取得をしない：`TargetMode="External"` の関係と `a:blip` の `r:link` は辿らない。
 * 図形・テキストボックスの文字（`drawing-text.ts`）と、線・矢印（`xdr:cxnSp`・線の preset の `xdr:sp`。種類 `line`）の形・回転・反転・枠線と塗りの有無（`drawing-shape.ts`）も読む（ADR-0027）。
 * 線でない図形の形（`a:prstGeom` の調整値・`a:custGeom` の定義。`drawing-geom.ts`）と枠線の太さ・線種も、文字の有無を問わず読む（ADR-0028）。
 * 図形の塗り・枠線、線、文字のランの色（出どころの順・代表色）は `drawing-color.ts` で解き、ブックのテーマ（`loadSheetDrawings` の `theme`）を使う（ISSUE-95）。
 * 線は図とは別に数える。画像のトリミング（`srcRect`）・回転・反転は読まない（第 1 段）。
 * グループ（`xdr:grpSp`）は中の子（`sp`・`cxnSp`・`pic`・`graphicFrame`・入れ子のグループ）を展開して一覧に足し、グループ自身は足さない。
 * 子はトップレベルと同じ `addDrawing` で読み、グループの子の座標系（`SheetDrawing.group`）を持たせる。写像（px）は表示側が行う（ISSUE-89）。
 * グラフは `graphicData` の `c:chart r:id` を同じ rels で解いて `chartPath` に持つ（読むのは `XlsxBook.readChart`。ISSUE-101）。
 * 図の名前・画像のパスはログに出さない（console を使わない）。
 */
import { readFillColor, readOutlineColor } from './drawing-color';
import { readGroupFrame, readGroupHead, readXfrm, type GroupHead } from './drawing-group';
import { readCustGeom, readPresetAdjustments } from './drawing-geom';
import {
  isLinePreset,
  readLine,
  readShapeProps,
  readStroke,
  type ShapeProps,
} from './drawing-shape';
import { readShapeText } from './drawing-text';
import { MAX_COLS, MAX_DRAWING_TEXT, MAX_ROWS, type XlsxLimits } from './limits';
import type { XlsxPackage } from './open';
import { DEFAULT_THEME, type Theme } from './theme';
import {
  XlsxError,
  type AnchorPoint,
  type DrawingAnchor,
  type DrawingGroup,
  type DrawingKind,
  type DrawingNotice,
  type GroupLevel,
  type GroupRect,
  type SheetDrawing,
  type SheetDrawings,
  type ShapeGeom,
} from './types';
import { baseOf, dirOf, parseRels, resolveTarget, type Relationship } from './workbook';
import { type Attrs, NS, attrByLocalName, findElement, findOpenTag, parseAttrs } from './xml';

/** EMU（ずれ・大きさ）の上限（2^31-1）。負・これを超える値は丸める */
const MAX_EMU = 2_147_483_647;

/** mc:AlternateContent を辿る深さの上限（入れ子の細工で再帰が深くならないように） */
const MAX_AC_DEPTH = 4;

/** 図の無い結果（共有しない：呼び出し側が書き換えても他に響かないよう毎回作る） */
const empty = (): SheetDrawings => ({ drawings: [], notices: [] });

/** アンカー（位置の入れ物）の要素名 */
const ANCHOR_RE = new RegExp(
  String.raw`<(${NS})(twoCellAnchor|oneCellAnchor|absoluteAnchor|AlternateContent)\b[^<>]*>`,
  'g',
);

/** アンカーの中の図の要素（最初に現れるものがその図） */
const CONTENT_RE = new RegExp(
  String.raw`<(${NS})(sp|grpSp|graphicFrame|cxnSp|pic|contentPart|AlternateContent)\b[^<>]*>`,
);

/** グループの直下の子（図の要素。入れ子のグループ・`mc:AlternateContent` を含む）の開きタグ */
const CHILD_RE = new RegExp(
  String.raw`<(${NS})(sp|grpSp|graphicFrame|cxnSp|pic|AlternateContent)\b[^<>]*>`,
  'g',
);

/** `xdr:pic` の開きタグ */
const PIC_RE = new RegExp(String.raw`<${NS}pic\b`);

/**
 * シートの XML（`<sheetData>` の後ろの部分）から `<drawing r:id>` の関係の Id を取り出す。無ければ null。
 * `legacyDrawing`（コメントの吹き出し・フォームコントロール）は別の要素なので拾わない。
 */
export function findDrawingRelId(sheetXml: string): string | null {
  // <drawing> は <sheetData> より後ろにある。巨大な sheetData を走査しないよう、最後の sheetData の閉じ以降だけを見る
  const tail = sheetXml.lastIndexOf('sheetData>');
  const el = findElement(sheetXml, 'drawing', tail < 0 ? 0 : tail);
  if (el === null) return null;
  return attrByLocalName(parseAttrs(el.open), 'id') ?? null;
}

/** 走査中に持ち回るもの */
interface Ctx {
  pkg: XlsxPackage;
  /** drawing のフォルダ（末尾に / ）。rels の Target の基準 */
  dir: string;
  rels: Map<string, Relationship>;
  limits: XlsxLimits;
  out: SheetDrawing[];
  /** 出した図（線を除く）の数 */
  figures: number;
  /** 出した線の数 */
  lines: number;
  /** 図の数が上限を超えた（以後の図は出さない。線の走査は続ける） */
  figuresFull: boolean;
  /** 線の数が上限を超えた（以後の線は出さない。図の走査は続ける） */
  linesFull: boolean;
  /** 図形の文字の合計（1 シート。文字数にラン・段落の数も足す。上限で以後の図形は文字を持たせない） */
  textUnits: number;
  /** 図形の文字を切った・捨てた（1 シートに 1 回だけ通知する） */
  textTruncated: boolean;
  /** ブックのテーマの 12 色（図形・線・文字の色の解決に使う。ブックで 1 回だけ読んだもの） */
  theme: Theme;
  /** 読んだ自由図形の path 命令と guide の数の合計（1 シート） */
  geomUnits: number;
  /** 自由図形の命令と guide の合計が上限を超えた（以後の自由図形は形を持たせない） */
  geomFull: boolean;
  /** 自由図形を上限で簡略化した（形を持たせなかった。1 シートに 1 回だけ通知する） */
  shapeSimplified: boolean;
}

/**
 * `<name>…</name>` を、同じ名前の入れ子を数えて線形時間で切り出す（`xml.ts` の `findElement` は最初の閉じで切るため、入れ子がありうる要素用）。
 * `from` は開きタグの先頭の位置。閉じが無いときは body を null・end を xml の長さにして走査を打ち切らせる。
 */
function findBalanced(
  xml: string,
  name: string,
  from: number,
): { body: string | null; end: number } | null {
  const re = new RegExp(String.raw`<(/?)(${NS})${name}\b[^<>]*>`, 'g');
  re.lastIndex = from;
  const first = re.exec(xml);
  if (first === null) return null;
  if (first[0].endsWith('/>')) return { body: null, end: re.lastIndex };
  const bodyStart = re.lastIndex;
  let depth = 1;
  for (let m = re.exec(xml); m !== null; m = re.exec(xml)) {
    if (m[1] === '/') {
      depth--;
      if (depth === 0) return { body: xml.slice(bodyStart, m.index), end: re.lastIndex };
    } else if (!m[0].endsWith('/>')) {
      depth++;
    }
  }
  return { body: null, end: xml.length };
}

/**
 * `mc:AlternateContent` の中身から、採る枝を決める：`xdr:pic` を含む枝（Choice・Fallback を問わず）、
 * 無ければ（図の要素を包むアンカーの中の AlternateContent のとき）`graphicFrame` を持つ Choice（`graphicData uri` は問わない。
 * 新しい種類のグラフ・スライサー・タイムラインの Fallback にある「お使いのバージョンでは利用できません」という説明文入りの図形を出さない）、
 * 無ければ Fallback、Fallback も無ければ最初の Choice。採る枝が無ければ null。
 * @param inAnchor アンカーの中の AlternateContent か（アンカー自身を枝に持つ外側のものは false。スライサーの Fallback のアンカーを今のまま採る）
 */
function pickBranch(body: string, inAnchor: boolean): string | null {
  const branches: string[] = [];
  let fallback: string | null = null;
  // この AlternateContent の直下の Choice・Fallback だけを順に拾う（入れ子の AlternateContent の枝は、外側の枝ごと飛ばす）
  const re = new RegExp(String.raw`<${NS}(Choice|Fallback)\b[^<>]*>`, 'g');
  for (let pos = 0; ;) {
    re.lastIndex = pos;
    const m = re.exec(body);
    if (m === null) break;
    const el = findBalanced(body, m[1]!, m.index);
    if (el === null) break;
    if (el.body !== null) {
      if (m[1] === 'Choice') branches.push(el.body);
      else fallback ??= el.body;
    }
    pos = el.end;
  }
  const all = fallback === null ? branches : [...branches, fallback];
  const pic = all.find((b) => PIC_RE.test(b));
  if (pic !== undefined) return pic;
  if (inAnchor) {
    const frame = branches.find((b) => hasGraphicFrame(b));
    if (frame !== undefined) return frame;
  }
  return fallback ?? branches[0] ?? null;
}

/** 枝が `graphicFrame`（グラフ・新しい種類のグラフ・スライサー・タイムラインなど。`graphicData uri` は問わない）を含むか */
function hasGraphicFrame(branch: string): boolean {
  return findOpenTag(branch, 'graphicFrame') !== null;
}

/** 整数の文字（前後の空白可）を数にする。整数でない・桁が多すぎるなら null */
function toInt(s: string | undefined): number | null {
  if (s === undefined) return null;
  const t = s.trim();
  if (!/^-?\d{1,15}$/.test(t)) return null;
  return Number(t);
}

/** EMU を 0〜MAX_EMU に丸める（負は 0） */
const clampEmu = (n: number): number => Math.min(Math.max(n, 0), MAX_EMU);

/** `<xdr:from>` / `<xdr:to>` の中身から 1 点を読む。番号が無い・不正・範囲外なら null（ずれが無ければ 0） */
function readPoint(head: string, tag: 'from' | 'to'): AnchorPoint | null {
  const body = findElement(head, tag)?.body;
  if (body === null || body === undefined) return null;
  const num = (name: string): string | undefined => findElement(body, name)?.body?.trim();
  const col = toInt(num('col'));
  const row = toInt(num('row'));
  if (col === null || row === null || col < 0 || col >= MAX_COLS || row < 0 || row >= MAX_ROWS)
    return null;
  const off = (name: string): number | null => {
    const s = num(name);
    return s === undefined ? 0 : (toInt(s) ?? null);
  };
  const colOff = off('colOff');
  const rowOff = off('rowOff');
  if (colOff === null || rowOff === null) return null;
  return { col, colOff: clampEmu(colOff), row, rowOff: clampEmu(rowOff) };
}

/** `<xdr:ext cx cy>` / `<xdr:pos x y>` の 2 つの整数の属性を読む。無い・不正なら null */
function readPair(head: string, tag: 'ext' | 'pos', a: string, b: string): [number, number] | null {
  const open = findElement(head, tag)?.open;
  if (open === undefined) return null;
  const attrs = parseAttrs(open);
  const x = toInt(attrs[a]);
  const y = toInt(attrs[b]);
  return x === null || y === null ? null : [clampEmu(x), clampEmu(y)];
}

/** アンカーの種類と、図の要素より前の部分（from・to・pos・ext）から位置を読む。読めなければ null */
function readAnchor(kind: string, head: string): DrawingAnchor | null {
  if (kind === 'twoCellAnchor') {
    const from = readPoint(head, 'from');
    const to = readPoint(head, 'to');
    return from && to ? { type: 'twoCell', from, to } : null;
  }
  if (kind === 'oneCellAnchor') {
    const from = readPoint(head, 'from');
    const ext = readPair(head, 'ext', 'cx', 'cy');
    return from && ext ? { type: 'oneCell', from, cx: ext[0], cy: ext[1] } : null;
  }
  const pos = readPair(head, 'pos', 'x', 'y');
  const ext = readPair(head, 'ext', 'cx', 'cy');
  return pos && ext ? { type: 'absolute', x: pos[0], y: pos[1], cx: ext[0], cy: ext[1] } : null;
}

/** 画像の参照（`a:blip r:embed`）を zip のパスにする。外部・r:embed 無し・zip に無いものは null（`r:link` は読まない） */
function resolveMedia(ctx: Ctx, blip: Attrs | null): string | null {
  const id = blip === null ? undefined : attrByLocalName(blip, 'embed');
  if (id === undefined || id === '') return null;
  const rel = ctx.rels.get(id);
  if (rel === undefined || rel.external) return null;
  const path = resolveTarget(ctx.dir, rel.target);
  return ctx.pkg.has(path) ? path : null;
}

/**
 * グラフの参照（`graphicFrame` の `a:graphicData` の子 `c:chart r:id`）を zip のパスにする（ADR-0031 §決定 2）。
 * 名前空間の接頭辞は問わず局所名 `chart` で引く。`uri` が通常のグラフ（`…/chart`）でないもの（chartex など）・`r:id` が無い・
 * 外部（`TargetMode="External"`）・zip に無いものは null（表示側は破線枠のまま）。
 * @param at `graphicFrame` の開きタグの位置
 */
function resolveChart(ctx: Ctx, body: string, at: number): string | null {
  const frame = findElement(body, 'graphicFrame', at)?.body;
  if (frame === null || frame === undefined) return null;
  const data = findElement(frame, 'graphicData');
  if (data === null || data.body === null) return null;
  if (!(parseAttrs(data.open).uri ?? '').endsWith('/chart')) return null;
  const tag = findOpenTag(data.body, 'chart');
  const id = tag === null ? undefined : attrByLocalName(parseAttrs(tag), 'id');
  if (id === undefined || id === '') return null;
  const rel = ctx.rels.get(id);
  if (rel === undefined || rel.external) return null;
  const path = resolveTarget(ctx.dir, rel.target);
  return ctx.pkg.has(path) ? path : null;
}

/**
 * 図形（`xdr:sp`）・線（`xdr:cxnSp`）の中身を切り出し、見た目（回転・反転・枠線と塗り・形）を読む。
 * `at` は要素の開きタグの位置。閉じが無いなどで中身が取れなければ null。
 */
function readShape(
  body: string,
  at: number,
  name: 'sp' | 'cxnSp',
): { spBody: string; props: ShapeProps } | null {
  const spBody = findElement(body, name, at)?.body ?? null;
  return spBody === null ? null : { spBody, props: readShapeProps(spBody) };
}

/**
 * アンカーの中の図の要素 1 つを一覧に足す。`at` は図の要素の開きタグの位置、`name` はその要素名。
 * 隠した図（`cNvPr hidden`）と、見えない線は足さない。
 * 線・矢印（`xdr:cxnSp`・線の preset の `xdr:sp`）は種類 `line` で、図とは別の上限（`maxLines`）で数える。
 * 数が上限を超えたら、その種類だけ足さず（`figuresFull`・`linesFull`）、もう一方は足し続ける。
 * 図形（`textbox`・`shape`）は、文字の有無を問わず回転・反転・枠線と塗りの有無・枠線の太さと線種・形を持つ。文字は文字のある図形だけが持つ。
 */
function addDrawing(
  ctx: Ctx,
  anchor: DrawingAnchor,
  body: string,
  at: number,
  name: string,
  depth: number,
  levels: readonly GroupLevel[] | null = null,
): void {
  if (name === 'AlternateContent') {
    const acBody = findBalanced(body, 'AlternateContent', at)?.body ?? null;
    const branch = acBody === null || depth >= MAX_AC_DEPTH ? null : pickBranch(acBody, true);
    if (branch === null) return;
    const m = CONTENT_RE.exec(branch);
    if (m !== null) addDrawing(ctx, anchor, branch, m.index, m[2]!, depth + 1, levels);
    return;
  }
  // グループは自分の頭（nvGrpSpPr・grpSpPr）だけを見る。`cNvPr` の無いグループで、子の `cNvPr`（隠した子など）を自分のものにしない
  const grpBody = name === 'grpSp' ? (findBalanced(body, 'grpSp', at)?.body ?? null) : null;
  const head = grpBody === null ? null : readGroupHead(grpBody);
  const cNvPr = head !== null ? head.cNvPr : findOpenTag(body, 'cNvPr', at);
  const attrs = cNvPr === null ? null : parseAttrs(cNvPr);
  if (attrs?.hidden === '1' || attrs?.hidden === 'true') return;
  // グループは中の子を展開する（グループ自身は足さない）。展開できないグループだけ、下で「グループ」の破線枠になる
  if (grpBody !== null && head !== null && expandGroup(ctx, anchor, grpBody, head, depth, levels)) {
    return;
  }

  let kind: DrawingKind;
  let mediaPath: string | null = null;
  let chartPath: string | null = null;
  let shape: { spBody: string; props: ShapeProps } | null = null;
  if (name === 'pic') {
    kind = 'image';
    const blip = findOpenTag(body, 'blip', at);
    mediaPath = resolveMedia(ctx, blip === null ? null : parseAttrs(blip));
  } else if (name === 'graphicFrame') {
    const uri = parseAttrs(findOpenTag(body, 'graphicData', at) ?? '').uri ?? '';
    kind = uri.endsWith('/diagram') ? 'smartart' : 'chart';
    if (kind === 'chart') chartPath = resolveChart(ctx, body, at);
  } else if (name === 'grpSp') {
    kind = 'group';
  } else if (name === 'cxnSp') {
    shape = readShape(body, at, 'cxnSp');
    if (shape === null) return;
    kind = 'line';
  } else if (name === 'sp') {
    shape = readShape(body, at, 'sp');
    if (shape !== null && isLinePreset(shape.props.prst)) {
      kind = 'line';
    } else {
      const tx = parseAttrs(findOpenTag(body, 'cNvSpPr', at) ?? '').txBox;
      kind = tx === '1' || tx === 'true' ? 'textbox' : 'shape';
    }
  } else {
    kind = 'shape';
  }

  // グループの中の図：グループの子の座標系での位置（`a:xfrm`）が要る。`xfrm` が読めない図は位置が決まらないので、黙って足さない
  // （仕様：入れ子の中でも同じ。兄弟の図は足す。トップレベルの図はアンカーが位置なので `xfrm` を要さない）
  let group: DrawingGroup | undefined;
  if (levels !== null) {
    const xfrm = childXfrm(body, at, name, shape, head);
    if (xfrm === null) return;
    group = { levels, xfrm };
  }

  const drawing: SheetDrawing = {
    anchor,
    kind,
    name: (attrs?.name ?? '').slice(0, MAX_DRAWING_TEXT),
    descr: (attrs?.descr ?? '').slice(0, MAX_DRAWING_TEXT),
    mediaPath,
  };
  if (chartPath !== null) drawing.chartPath = chartPath;
  if (group !== undefined) drawing.group = group;

  if (kind === 'line') {
    const props = shape!.props;
    // 見えない線（a:ln の noFill・lnRef idx="0"）は足さない（数えない）
    if (!props.hasOutline) return;
    if (ctx.lines >= ctx.limits.maxLines) {
      ctx.linesFull = true;
      return;
    }
    ctx.lines++;
    drawing.rot = props.rot;
    drawing.flipH = props.flipH;
    drawing.flipV = props.flipV;
    drawing.line = readLine(props);
    drawing.line.color = readOutlineColor(props, ctx.theme);
    ctx.out.push(drawing);
    return;
  }

  if (ctx.figures >= ctx.limits.maxDrawings) {
    ctx.figuresFull = true;
    return;
  }
  ctx.figures++;
  if (shape !== null) {
    // 図形の見た目（回転・反転・枠線と塗りの有無・枠線の太さと線種・形）。文字の有無を問わず持たせる（ADR-0028 §決定 3）
    const props = shape.props;
    drawing.rot = props.rot;
    drawing.flipH = props.flipH;
    drawing.flipV = props.flipV;
    // 回転のあるトップレベルの図形だけ、`a:xfrm` の `a:ext`（回転前の大きさ）を持たせる（回転前の矩形の決め方に使う。ISSUE-93）
    if (group === undefined && props.rot !== 0) {
      const xf = readXfrm(shape.spBody);
      if (xf !== null && xf.cx > 0 && xf.cy > 0) drawing.xfrmExt = { cx: xf.cx, cy: xf.cy };
    }
    drawing.hasOutline = props.hasOutline;
    drawing.hasFill = props.hasFill;
    drawing.outline = readStroke(props);
    // 色は「ある」と決まった塗り・枠線にだけ持たせる（有無の規則は ADR-0027 のまま）
    if (props.hasFill) drawing.fillColor = readFillColor(props, ctx.theme);
    if (props.hasOutline) drawing.outlineColor = readOutlineColor(props, ctx.theme);
    const geom = readShapeGeom(ctx, props);
    if (geom !== undefined) drawing.geom = geom;
    // 文字（上限は 1 シートの合計の残りでも切る。合計には文字数のほかラン・段落も 1 つずつ数える。尽きた後の図形は文字を持たせない）
    const room = ctx.limits.maxSheetTextChars - ctx.textUnits;
    const read = readShapeText(
      shape.spBody,
      {
        maxChars: ctx.limits.maxShapeTextChars,
        maxUnits: Math.max(room, 0),
        maxParagraphs: ctx.limits.maxShapeParagraphs,
        maxRuns: ctx.limits.maxShapeRuns,
      },
      ctx.theme,
    );
    if (read.truncated) ctx.textTruncated = true;
    if (read.text !== null) {
      ctx.textUnits += read.units;
      drawing.text = read.text;
    }
  }
  ctx.out.push(drawing);
}

/**
 * 図形の形（`a:prstGeom`＋調整値、または `a:custGeom` の定義）を読む。形が無い・自由図形の字句の誤り・上限超えは undefined。
 * 自由図形の上限（1 図形の guide・命令、1 シートの命令の合計）を超えたら `shape-simplified` を立て、シート合計を超えたら以後の自由図形は読まない。
 * 図形 1 つ（`sp`）の読み取りで、トップレベルの図形もグループの子も同じ `addDrawing` を通る。
 */
function readShapeGeom(ctx: Ctx, props: ShapeProps): ShapeGeom | undefined {
  if (props.prst !== null && props.prst !== '') {
    return {
      type: 'preset',
      prst: props.prst,
      adj: readPresetAdjustments(props.geomBody, ctx.limits.maxShapeAdjustments),
    };
  }
  if (props.custGeom === null) return undefined;
  if (ctx.geomFull) {
    ctx.shapeSimplified = true;
    return undefined;
  }
  const read = readCustGeom(props.custGeom.body, {
    maxGuides: ctx.limits.maxGeomGuides,
    maxCommands: ctx.limits.maxGeomCommands,
    sheetRoom: Math.max(ctx.limits.maxSheetGeomUnits - ctx.geomUnits, 0),
  });
  if (read.ok) {
    ctx.geomUnits += read.units;
    return { type: 'custom', def: read.def };
  }
  if (read.reason !== 'invalid') {
    ctx.shapeSimplified = true;
    if (read.reason === 'sheet-limit') ctx.geomFull = true;
  }
  return undefined;
}

/**
 * グループの中の図の `a:xfrm`（グループの子の座標系での位置と大きさ）。`at` は図の要素の開きタグの位置。
 * `sp`・`cxnSp` は `readShape` が切り出した中身から、`pic`・`graphicFrame` はその要素の中身から、グループは自分の `grpSpPr`（`head`）から読む。読めなければ null。
 */
function childXfrm(
  body: string,
  at: number,
  name: string,
  shape: { spBody: string } | null,
  head: GroupHead | null,
): GroupRect | null {
  if (name === 'sp' || name === 'cxnSp') return shape === null ? null : readXfrm(shape.spBody);
  if (name === 'grpSp') {
    const pr = head?.pr?.body;
    return pr === null || pr === undefined ? null : readXfrm(pr);
  }
  if (name === 'pic' || name === 'graphicFrame') {
    const el = findElement(body, name, at)?.body;
    return el === null || el === undefined ? null : readXfrm(el);
  }
  return null;
}

/**
 * グループの中身を展開して、子を一覧に足す。展開したら true。
 * 展開できないとき（深さが上限を超える・座標系が読めない・`chExt` が 0）は false で、何も足さない（呼び出し側が破線枠にする）。
 * 子は `levels`（外側のグループから順の座標系）に、このグループの座標系を足して `addDrawing` へ渡す。
 * @param grp グループの中身（開きタグと閉じタグの間）
 * @param head グループ自身の cNvPr・grpSpPr（`readGroupHead`）
 * @param levels このグループを囲むグループの座標系（トップレベルのグループは null）
 */
function expandGroup(
  ctx: Ctx,
  anchor: DrawingAnchor,
  grp: string,
  head: GroupHead,
  depth: number,
  levels: readonly GroupLevel[] | null,
): boolean {
  if ((levels?.length ?? 0) >= ctx.limits.maxGroupDepth) return false;
  const frame = readGroupFrame(head.pr, levels !== null);
  if (frame === null) return false;
  scanGroupChildren(ctx, anchor, grp, frame.end, depth, [
    ...(levels ?? []),
    { ch: frame.ch, own: frame.own },
  ]);
  return true;
}

/**
 * グループの中身（`xml`）の `from` 以降にある直下の子を順に `addDrawing` へ渡す。
 * 子の要素 1 つ分だけを切り出して渡す（図の読み取りが、後ろの兄弟の要素まで読まないように）。
 * 子の終わりを探したあと、その中身は走査し直さない（入れ子のグループは `expandGroup` が自分の中身だけを走査する）。
 */
function scanGroupChildren(
  ctx: Ctx,
  anchor: DrawingAnchor,
  xml: string,
  from: number,
  depth: number,
  levels: readonly GroupLevel[],
): void {
  let pos = from;
  while (!(ctx.figuresFull && ctx.linesFull)) {
    CHILD_RE.lastIndex = pos;
    const m = CHILD_RE.exec(xml);
    if (m === null) return;
    const name = m[2]!;
    const el =
      name === 'grpSp' || name === 'AlternateContent'
        ? findBalanced(xml, name, m.index)
        : findElement(xml, name, m.index);
    if (el === null) return;
    pos = el.end;
    if (el.body === null) continue;
    addDrawing(ctx, anchor, xml.slice(m.index, el.end), 0, name, depth, levels);
  }
}

/** drawing XML（またはその一部）のアンカーを順に読み、図を足す。トップレベルの `mc:AlternateContent` は採る枝の中を読む */
function scanAnchors(ctx: Ctx, xml: string, depth: number): void {
  let pos = 0;
  // 図・線の片方が上限に達しても走査は続ける（もう一方が足りる）。両方が上限を超えたら、以後に足せるものは無い
  while (!(ctx.figuresFull && ctx.linesFull)) {
    ANCHOR_RE.lastIndex = pos;
    const m = ANCHOR_RE.exec(xml);
    if (m === null) return;
    const name = m[2]!;
    if (name === 'AlternateContent') {
      const ac = findBalanced(xml, name, m.index);
      if (ac === null) return;
      pos = ac.end;
      if (ac.body !== null && depth < MAX_AC_DEPTH) {
        const branch = pickBranch(ac.body, false);
        if (branch !== null) scanAnchors(ctx, branch, depth + 1);
      }
      continue;
    }
    const el = findElement(xml, name, m.index);
    if (el === null) return;
    pos = el.end;
    if (el.body === null) continue;
    const content = CONTENT_RE.exec(el.body);
    if (content === null) continue;
    const anchor = readAnchor(name, el.body.slice(0, content.index));
    if (anchor === null) continue;
    addDrawing(ctx, anchor, el.body, content.index, content[2]!, depth);
  }
}

/**
 * シートの図を読む。`relId` はシートの `<drawing r:id>`。
 * 上限超え・読み取りの失敗は例外にせず、通知（`notices`）で返す。図の無いシート・zip に解決できない参照は空の一覧。
 * @param sheetPath シート XML の zip の中のパス（rels の場所と Target の基準）
 */
export async function loadSheetDrawings(
  pkg: XlsxPackage,
  sheetPath: string,
  relId: string,
  limits: XlsxLimits,
  theme: Theme = DEFAULT_THEME,
): Promise<SheetDrawings> {
  const small = limits.maxSmallPartBytes;
  try {
    const sheetRels = await pkg.readText(
      `${dirOf(sheetPath)}_rels/${baseOf(sheetPath)}.rels`,
      small,
    );
    if (sheetRels === null) return empty();
    const rel = parseRels(sheetRels).get(relId);
    if (rel === undefined || rel.external) return empty();
    const drawingPath = resolveTarget(dirOf(sheetPath), rel.target);
    if (!pkg.has(drawingPath)) return empty();
    const xml = await pkg.readText(drawingPath, small);
    if (xml === null) return empty();
    const relsXml = await pkg.readText(
      `${dirOf(drawingPath)}_rels/${baseOf(drawingPath)}.rels`,
      small,
    );
    const ctx: Ctx = {
      pkg,
      dir: dirOf(drawingPath),
      rels: relsXml === null ? new Map() : parseRels(relsXml),
      limits,
      out: [],
      figures: 0,
      lines: 0,
      figuresFull: false,
      linesFull: false,
      textUnits: 0,
      textTruncated: false,
      theme,
      geomUnits: 0,
      geomFull: false,
      shapeSimplified: false,
    };
    scanAnchors(ctx, xml, 0);
    const notices: DrawingNotice[] = [];
    if (ctx.figuresFull || ctx.linesFull) notices.push('too-many');
    if (ctx.textTruncated) notices.push('text-truncated');
    if (ctx.shapeSimplified) notices.push('shape-simplified');
    return { drawings: ctx.out, notices };
  } catch (e) {
    const notice: DrawingNotice =
      e instanceof XlsxError && e.kind === 'too-large' ? 'part-too-large' : 'broken';
    return { drawings: [], notices: [notice] };
  }
}
