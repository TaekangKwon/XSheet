// 正: InterHub docs/plan/adr/0025-xlsx-own-reader.md
/**
 * xlsx の読み取り部品の公開の入口（画面に依存しない。Node の vitest でも動く）。
 * jszip で zip を開き、workbook・共有文字列・styles・シートの XML を自前で読み、
 * 数値の表示形式だけ `ssf` に任せる。シートは 1 枚ずつ開き、見えている行だけを展開する。
 * セルの値・シート名はログに出さない（console を使わない）。
 */
import { parseBorders, type BorderTable } from './borders';
import { parseCellFormats, stylesUseTheme, type FormatTable } from './cell-format';
import { parseChart } from './chart';
import { findDrawingRelId, loadSheetDrawings } from './drawing';
import { createFormatter, type CellFormatter } from './format';
import { createImageBudget, readImageFrom, type ImageBudget } from './image';
import { DEFAULT_LIMITS, type XlsxLimits } from './limits';
import { checkSignature, openPackage, toBytes, type XlsxPackage } from './open';
import { createEmptySheet, createSheet, type ReleasableSheet } from './sheet';
import { parseSharedStrings, type SharedStrings } from './shared-strings';
import { parseStyles } from './styles';
import { DEFAULT_THEME, loadTheme, type Theme } from './theme';
import {
  XlsxError,
  type ChartResult,
  type ImageResult,
  type SheetDrawings,
  type SheetInfo,
  type XlsxBook,
  type XlsxSheet,
} from './types';
import { loadWorkbook } from './workbook';

export { XlsxError } from './types';
// 公開の型は、公開の関数・型から辿れるものを全部出す（図の型も openDrawings の戻り値から辿れる）
export type {
  AnchorPoint,
  ArrowSize,
  ArrowType,
  BorderEdge,
  BorderStyle,
  CellBorder,
  CellFormat,
  CellKind,
  CellView,
  ChartAxis,
  ChartBox,
  ChartFailReason,
  ChartLegend,
  ChartLine,
  ChartMarker,
  ChartMarkerSymbol,
  ChartModel,
  ChartPaint,
  ChartResult,
  ChartSeries,
  ChartTextStyle,
  ChartTickMark,
  ChartTitle,
  ColLayout,
  CustGeomDef,
  CustGeomPath,
  DrawingAnchor,
  DrawingColor,
  DrawingGroup,
  DrawingKind,
  DrawingLine,
  DrawingNotice,
  DrawingText,
  DrawingTextParagraph,
  DrawingTextRun,
  GeomArg,
  GeomPathFill,
  GroupLevel,
  GroupRect,
  ImageFailReason,
  ImageFormat,
  ImageResult,
  LineDash,
  LineEnd,
  LineGeom,
  ManualLayout,
  RowLayout,
  ShapeGeom,
  ShapeOutline,
  SheetDrawing,
  SheetDrawings,
  SheetInfo,
  SheetLayout,
  SheetRow,
  TextAlign,
  TextAnchor,
  TextVert,
  XlsxBook,
  XlsxErrorKind,
  XlsxSheet,
} from './types';
export type { Rgba } from './color';
export { createGeometry, colWidthToPx, defaultColWidthToPx, ptToPx } from './geometry';
export { fontDefaultRowPx, lookupMdw } from './font-metrics';
export type { Geometry } from './geometry';
export type { XlsxLimits } from './limits';
export { DEFAULT_LIMITS } from './limits';
export { colName } from './xml';
export { isPaperFill } from './cell-format';
export { findText } from './find-text';
export type { FindTextOptions, FindTextResult, TextMatch } from './find-text';

/** styles.xml から読むもの：表示形式の整形器と、寸法の換算に使う最大桁幅 MDW・フォントの既定の行高・罫線の表・書式（塗り・文字の装飾）の表 */
interface StyleInfo {
  formatter: CellFormatter;
  mdw: number;
  fontRowPx: number | null;
  borderTable: BorderTable;
  formatTable: FormatTable;
}

/** 1 シートのグラフの点の合計の数え方（readChart の呼び出しをまたいで持つ） */
interface ChartBudget {
  /** これまでに読めたグラフの点の合計 */
  used: number;
  /** 合計の上限を一度でも超えた（以後、新しいグラフはすべて too-large） */
  exhausted: boolean;
  /** すでに数えたグラフのパス（同じグラフを何度読んでも 1 回だけ数える） */
  counted: Set<string>;
}

/** 空の ChartBudget */
const createChartBudget = (): ChartBudget => ({ used: 0, exhausted: false, counted: new Set() });

/** openXlsx の追加の指定 */
export interface OpenXlsxOptions {
  /** 上限を替える（既定は DEFAULT_LIMITS。試験で小さい値を渡す用） */
  limits?: Partial<XlsxLimits>;
}

/**
 * xlsx を開く（シートの中身は openSheet で開くときに初めて展開する）。
 * @throws XlsxError zip でない（not-zip）・暗号化（encrypted）・壊れている（broken）・大きすぎる（too-large）・読めない形式（unsupported）
 */
export async function openXlsx(
  data: ArrayBuffer | Uint8Array,
  options: OpenXlsxOptions = {},
): Promise<XlsxBook> {
  const limits: XlsxLimits = { ...DEFAULT_LIMITS, ...options.limits };
  const bytes = toBytes(data);
  checkSignature(bytes);
  const pkg = await openPackage(bytes, limits);
  let workbook;
  try {
    workbook = await loadWorkbook(pkg, limits);
  } catch (e) {
    pkg.dispose();
    throw e;
  }

  let closed = false;
  // 開いたシートは弱い参照で持つ（画面が捨てたシートの XML を、ここで抱え続けないため）
  const opened = new Set<WeakRef<ReleasableSheet>>();
  let stringsPromise: Promise<SharedStrings> | null = null;
  let formatterPromise: Promise<StyleInfo> | null = null;
  // テーマの 12 色（ブックで 1 回だけ読む。図形へ渡すのは ISSUE-95）
  let themePromise: Promise<Theme> | null = null;
  // 図：シートの <drawing r:id>（openSheet で読んだ XML から覚える。null＝図なし）。画像の合計の上限はシートごとに数える
  const drawingRelIds = new Map<number, string | null>();
  const imageBudgets = new Map<number, ImageBudget>();
  /** readImage を 1 つずつ順に走らせる（合計の上限を、呼んだ順に確実に数えるため） */
  let imageChain: Promise<unknown> = Promise.resolve();
  // グラフ：シートごとに、読んだグラフの点の合計を数える（上限は maxSheetChartPoints。画像の合計と同じく呼んだ順に数え、超えたら以後の新しいグラフは too-large）
  const chartBudgets = new Map<number, ChartBudget>();
  /** readChart を 1 つずつ順に走らせる（合計の上限を、呼んだ順に確実に数えるため） */
  let chartChain: Promise<unknown> = Promise.resolve();

  const loadStrings = (p: XlsxPackage): Promise<SharedStrings> => {
    stringsPromise ??= (async () => {
      const path = workbook.sharedStringsPath;
      const xml =
        path !== null && p.has(path) ? await p.readText(path, limits.maxSharedStringsBytes) : null;
      return parseSharedStrings(xml);
    })();
    return stringsPromise;
  };
  /** テーマの 12 色（xl/theme/theme1.xml。無い・読めないときは既定の 12 色。ブックで 1 回だけ読む） */
  const loadThemeOnce = (p: XlsxPackage): Promise<Theme> => {
    themePromise ??= loadTheme(p, workbook.themePath, limits);
    return themePromise;
  };

  /** 表示形式の整形器・寸法の換算に使う最大桁幅 MDW・罫線の表・書式の表（いずれも styles.xml から。ブックで 1 回だけ読む） */
  const loadFormatter = (p: XlsxPackage): Promise<StyleInfo> => {
    formatterPromise ??= (async () => {
      const path = workbook.stylesPath;
      const xml =
        path !== null && p.has(path) ? await p.readText(path, limits.maxSmallPartBytes) : null;
      const styles = parseStyles(xml);
      // 罫線・塗り・文字の色はテーマで解く（テーマはブックで 1 回だけ読む。罫線・書式の表も styles.xml と同じくブックで 1 回）。
      // <fonts>・<fills>・<borders> が theme を使わないブックでは、テーマの読み取りを増やさない
      const theme = stylesUseTheme(xml) ? await loadThemeOnce(p) : DEFAULT_THEME;
      return {
        formatter: createFormatter(styles, workbook.date1904),
        mdw: styles.mdw,
        fontRowPx: styles.fontRowPx,
        borderTable: parseBorders(xml, theme, limits),
        formatTable: parseCellFormats(xml, theme, limits),
      };
    })();
    return formatterPromise;
  };

  /**
   * グラフの部品を読んで `ChartModel` にする（失敗は例外にせず理由つき）。部品は `maxSmallPartBytes` で抑え、テーマ（ブックで 1 回）と `date1904` を渡す。
   * シートの点の合計（系列数 × 項目の数）は、まだ数えていないグラフだけ数える（同じグラフの再読み込みは数え直さない）。
   * 合計を使い切ったら以後そのシートの新しいグラフは読まずに `too-large`（一度超えたら戻らない）。
   */
  const readChartFrom = async (path: string, budget: ChartBudget): Promise<ChartResult> => {
    if (!pkg.has(path)) return { ok: false, reason: 'broken' };
    const charged = !budget.counted.has(path);
    if (charged && budget.exhausted) return { ok: false, reason: 'too-large' };
    let xml: string | null;
    try {
      xml = await pkg.readText(path, limits.maxSmallPartBytes);
    } catch (e) {
      if (closed) throw new XlsxError('broken', 'closed');
      return {
        ok: false,
        reason: e instanceof XlsxError && e.kind === 'too-large' ? 'too-large' : 'broken',
      };
    }
    if (xml === null) return { ok: false, reason: 'broken' };
    const theme = await loadThemeOnce(pkg);
    const parsed = parseChart(xml, theme, limits, workbook.date1904);
    if (!parsed.ok) return parsed;
    // 描画側が既定の系列色（テーマの accent）を当てられるよう、テーマも返す（ISSUE-103）
    const r: ChartResult = { ...parsed, theme };
    if (!charged) return r;
    // await をまたいで同じパスが数え済みになっていたら数え直さない（破棄された表示の読み込みと新しい表示の読み込みが同じパスを重ねて数えない）
    if (budget.counted.has(path)) return r;
    const points = parsed.chart.series.length * parsed.chart.categories.length;
    if (budget.used + points > limits.maxSheetChartPoints) {
      budget.exhausted = true;
      return { ok: false, reason: 'too-large' };
    }
    budget.used += points;
    budget.counted.add(path);
    return r;
  };

  const sheets: readonly SheetInfo[] = workbook.sheets.map((s) => ({
    index: s.index,
    name: s.name,
    hidden: s.hidden,
  }));

  return {
    sheets,
    async openSheet(index: number): Promise<XlsxSheet> {
      if (closed) throw new XlsxError('broken', 'closed');
      const info = workbook.sheets[index - 1];
      if (!Number.isInteger(index) || info === undefined)
        throw new RangeError('シートの番号が範囲外');
      if (!info.isWorksheet) return createEmptySheet();
      if (info.path === null) throw new XlsxError('broken', 'sheet path');
      const xml = await pkg.readText(info.path, limits.maxSheetXmlBytes);
      if (xml === null) throw new XlsxError('broken', 'sheet part');
      drawingRelIds.set(index, findDrawingRelId(xml));
      const [strings, { formatter, mdw, fontRowPx, borderTable, formatTable }] = await Promise.all([
        loadStrings(pkg),
        loadFormatter(pkg),
      ]);
      if (closed) throw new XlsxError('broken', 'closed');
      const sheet = createSheet(xml, {
        strings,
        formatter,
        date1904: workbook.date1904,
        limits,
        mdw,
        fontRowPx,
        borderTable,
        formatTable,
      });
      for (const ref of opened) if (ref.deref() === undefined) opened.delete(ref);
      opened.add(new WeakRef(sheet));
      return sheet;
    },
    async openDrawings(index: number): Promise<SheetDrawings> {
      if (closed) throw new XlsxError('broken', 'closed');
      const info = workbook.sheets[index - 1];
      if (!Number.isInteger(index) || info === undefined)
        throw new RangeError('シートの番号が範囲外');
      if (!info.isWorksheet || info.path === null) return { drawings: [], notices: [] };
      let relId = drawingRelIds.get(index);
      if (relId === undefined) {
        // openSheet より先に呼ばれた：シートの XML から <drawing> の指定だけを読む
        try {
          const xml = await pkg.readText(info.path, limits.maxSheetXmlBytes);
          relId = xml === null ? null : findDrawingRelId(xml);
        } catch (e) {
          if (closed) throw new XlsxError('broken', 'closed');
          const tooLarge = e instanceof XlsxError && e.kind === 'too-large';
          return { drawings: [], notices: [tooLarge ? 'part-too-large' : 'broken'] };
        }
        drawingRelIds.set(index, relId);
      }
      if (relId === null) return { drawings: [], notices: [] };
      // 図形の色の解決（ISSUE-95）で使うテーマ（読めなくても既定の 12 色になり、表示は止まらない）
      const theme = await loadThemeOnce(pkg);
      const result = await loadSheetDrawings(pkg, info.path, relId, limits, theme);
      if (closed) throw new XlsxError('broken', 'closed');
      return result;
    },
    readImage(index: number, path: string): Promise<ImageResult> {
      if (closed) return Promise.reject(new XlsxError('broken', 'closed'));
      if (!Number.isInteger(index) || workbook.sheets[index - 1] === undefined)
        return Promise.reject(new RangeError('シートの番号が範囲外'));
      // 合計の上限は、呼び出し側が指定したシートの分で数える（複数のシートが同じ画像を共有しても、数え先が入れ替わらない）
      const budget = imageBudgets.get(index) ?? createImageBudget();
      imageBudgets.set(index, budget);
      const run = imageChain.then(async () => {
        const r = await readImageFrom(pkg, path, limits, budget);
        if (closed) throw new XlsxError('broken', 'closed');
        return r;
      });
      imageChain = run.catch(() => undefined);
      return run;
    },
    readChart(index: number, path: string): Promise<ChartResult> {
      if (closed) return Promise.reject(new XlsxError('broken', 'closed'));
      if (!Number.isInteger(index) || workbook.sheets[index - 1] === undefined)
        return Promise.reject(new RangeError('シートの番号が範囲外'));
      const budget = chartBudgets.get(index) ?? createChartBudget();
      chartBudgets.set(index, budget);
      const run = chartChain.then(async () => {
        const r = await readChartFrom(path, budget);
        if (closed) throw new XlsxError('broken', 'closed');
        return r;
      });
      chartChain = run.catch(() => undefined);
      return run;
    },
    close() {
      closed = true;
      for (const ref of opened) ref.deref()?.release();
      opened.clear();
      stringsPromise = null;
      formatterPromise = null;
      themePromise = null;
      drawingRelIds.clear();
      imageBudgets.clear();
      chartBudgets.clear();
      pkg.dispose();
    },
  };
}
