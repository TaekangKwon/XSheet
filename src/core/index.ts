// 正: InterHub docs/plan/adr/0025-xlsx-own-reader.md
/**
 * xlsx の読み取り部品の公開の入口（画面に依存しない。Node の vitest でも動く）。
 * jszip で zip を開き、workbook・共有文字列・styles・シートの XML を自前で読み、
 * 数値の表示形式だけ `ssf` に任せる。シートは 1 枚ずつ開き、見えている行だけを展開する。
 * セルの値・シート名はログに出さない（console を使わない）。
 */
import { parseBorders, type BorderTable } from './borders';
import { parseCellFormats, stylesUseTheme, type FormatTable } from './cell-format';
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
  type ImageResult,
  type SheetDrawings,
  type SheetInfo,
  type XlsxBook,
  type XlsxSheet,
} from './types';
import { loadWorkbook } from './workbook';

export { XlsxError } from './types';
export type {
  AnchorPoint,
  BorderEdge,
  BorderStyle,
  CellBorder,
  CellFormat,
  CellKind,
  CellView,
  ColLayout,
  DrawingAnchor,
  DrawingKind,
  DrawingNotice,
  ImageFailReason,
  ImageFormat,
  ImageResult,
  RowLayout,
  SheetDrawing,
  SheetDrawings,
  SheetInfo,
  SheetLayout,
  SheetRow,
  XlsxBook,
  XlsxErrorKind,
  XlsxSheet,
} from './types';
export { createGeometry, colWidthToPx, defaultColWidthToPx, ptToPx } from './geometry';
export { fontDefaultRowPx, lookupMdw } from './font-metrics';
export type { Geometry } from './geometry';
export type { XlsxLimits } from './limits';
export { DEFAULT_LIMITS } from './limits';
export { colName } from './xml';
export { isPaperFill } from './cell-format';

/** styles.xml から読むもの：表示形式の整形器と、寸法の換算に使う最大桁幅 MDW・フォントの既定の行高・罫線の表・書式（塗り・文字の装飾）の表 */
interface StyleInfo {
  formatter: CellFormatter;
  mdw: number;
  fontRowPx: number | null;
  borderTable: BorderTable;
  formatTable: FormatTable;
}

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
    close() {
      closed = true;
      for (const ref of opened) ref.deref()?.release();
      opened.clear();
      stringsPromise = null;
      formatterPromise = null;
      themePromise = null;
      drawingRelIds.clear();
      imageBudgets.clear();
      pkg.dispose();
    },
  };
}
