// XSheet 0.3.0 単独版（xsheet.min.js・xsheet.mjs）の型。tools/standalone/build.mjs が dist の d.ts から作る（手で直さない）。

// <script> で読むときは大域の XSheet、型は import type { Viewer } from 'xsheet' で使える。

declare module 'xsheet' {
  /**
   * パッケージの入口（`xsheet`）。入口の層（`view`・`read`）＋core（読み取り）＋render（表示）の公開物を全部出す（M6-2）。
   * 読み込んだ時点では DOM に触れない（Node で import してよい）。DOM・Node 固有の API を使わないことを約束する入口は `xsheet/core`。
   * 使わない部分はバンドラが落とす（package.json の sideEffects は CSS だけ）。
   */
  export * from 'xsheet/_/core/index';
  export * from 'xsheet/_/render/index';
  export * from 'xsheet/_/facade/index';
}

declare module 'xsheet/_/core/index' {
  import { type XlsxLimits } from 'xsheet/_/core/limits';
  import { type XlsxBook } from 'xsheet/_/core/types';
  export { XlsxError } from 'xsheet/_/core/types';
  export type { AnchorPoint, ArrowSize, ArrowType, BorderEdge, BorderStyle, CellBorder, CellFormat, CellKind, CellView, ChartAxis, ChartBox, ChartFailReason, ChartLegend, ChartLine, ChartMarker, ChartMarkerSymbol, ChartModel, ChartPaint, ChartResult, ChartSeries, ChartTextStyle, ChartTickMark, ChartTitle, ColLayout, CustGeomDef, CustGeomPath, DrawingAnchor, DrawingColor, DrawingGroup, DrawingKind, DrawingLine, DrawingNotice, DrawingText, DrawingTextParagraph, DrawingTextRun, GeomArg, GeomPathFill, GroupLevel, GroupRect, ImageFailReason, ImageFormat, ImageResult, LineDash, LineEnd, LineGeom, ManualLayout, RowLayout, ShapeGeom, ShapeOutline, SheetDrawing, SheetDrawings, SheetInfo, SheetLayout, SheetRow, TextAlign, TextAnchor, TextVert, XlsxBook, XlsxErrorKind, XlsxSheet, } from 'xsheet/_/core/types';
  export type { Rgba } from 'xsheet/_/core/color';
  export { createGeometry, colWidthToPx, defaultColWidthToPx, ptToPx } from 'xsheet/_/core/geometry';
  export { fontDefaultRowPx, lookupMdw } from 'xsheet/_/core/font-metrics';
  export type { Geometry } from 'xsheet/_/core/geometry';
  export type { XlsxLimits } from 'xsheet/_/core/limits';
  export { DEFAULT_LIMITS } from 'xsheet/_/core/limits';
  export { colName } from 'xsheet/_/core/xml';
  export { isPaperFill } from 'xsheet/_/core/cell-format';
  /** openXlsx の追加の指定 */
  export interface OpenXlsxOptions {
      /** 上限を替える（既定は DEFAULT_LIMITS。試験で小さい値を渡す用） */
      limits?: Partial<XlsxLimits>;
  }
  /**
   * xlsx を開く（シートの中身は openSheet で開くときに初めて展開する）。
   * @throws XlsxError zip でない（not-zip）・暗号化（encrypted）・壊れている（broken）・大きすぎる（too-large）・読めない形式（unsupported）
   */
  export function openXlsx(data: ArrayBuffer | Uint8Array, options?: OpenXlsxOptions): Promise<XlsxBook>;
}

declare module 'xsheet/_/core/limits' {
  /**
   * xlsx の読み取りの上限（定数の置き場は本ファイルの 1 か所だけ。図形の guide・path 命令の上限だけは geom の `defs.ts`）。
   * 読み込むファイルは信頼できない入力として扱う（Excel の仕様上限と XML の大きさの見積もりに基づく値）。
   * zip の展開（zip 爆弾）・巨大な添字・大量の結合セルで、画面が固まったりメモリが尽きたりしないようにする。
   * 超えたら `XlsxError('too-large')`。
   */
  import { MAX_GEOM_COMMANDS, MAX_GEOM_GUIDES } from 'xsheet/_/geom/defs';
  /** 上限の組。テストでは小さい値を渡して上限超えを確かめる */
  export interface XlsxLimits {
      /** 1 シートの XML を展開した後の大きさ（バイト） */
      maxSheetXmlBytes: number;
      /** sharedStrings.xml を展開した後の大きさ（バイト） */
      maxSharedStringsBytes: number;
      /** workbook.xml・styles.xml・各 rels など小さな部品 1 つの展開後の大きさ（バイト） */
      maxSmallPartBytes: number;
      /** zip の中のファイル数 */
      maxEntries: number;
      /** シートの数 */
      maxSheets: number;
      /** 1 シートの結合セルの数 */
      maxMerges: number;
      /** 1 シートの図（線・矢印を除く）の数。超えた分は出さない（ADR-0026 §決定 7・ADR-0027 §決定 6） */
      maxDrawings: number;
      /** 1 シートの線・矢印（`line`）の数。図とは別に数え、超えた分は出さない（ADR-0027 §決定 6） */
      maxLines: number;
      /** 図形 1 つの文字数。超えた分は切って末尾に「…」を付ける（ADR-0027 §決定 6） */
      maxShapeTextChars: number;
      /** 図形 1 つの段落（`a:p`）の数。超えた分は捨てる */
      maxShapeParagraphs: number;
      /** 図形 1 つのラン（`a:r`・`a:br`・`a:fld`）の数。超えた分は捨てる */
      maxShapeRuns: number;
      /** 1 シートの図形の文字の合計（文字数に、ラン・段落の数を 1 つずつ足す）。超えたら以後の図形は文字を持たせない */
      maxSheetTextChars: number;
      /** グループの入れ子の深さ（トップレベルのグループを 1 と数える）。これより深いグループは展開せず破線枠にする（ADR-0027 §決定 6） */
      maxGroupDepth: number;
      /** 画像 1 枚の展開後の大きさ（バイト）。超えた画像は枠にする */
      maxImageBytes: number;
      /** 1 シートの画像の展開後の合計（バイト）。超えたら以後の画像を枠にする */
      maxSheetImageBytes: number;
      /** 画像の 1 辺（px）。超えた画像は枠にする */
      maxImageSidePx: number;
      /** 画像 1 枚の画素数（幅×高さ）。超えた画像は枠にする */
      maxImagePixels: number;
      /** 自由図形（`a:custGeom`）1 つの guide の数（`avLst`＋`gdLst`）。超えたら形を持たせず `shape-simplified`（ADR-0028 §決定 7） */
      maxGeomGuides: number;
      /** 自由図形 1 つの path 命令の数（全 path の合計）。超えたら形を持たせず `shape-simplified` */
      maxGeomCommands: number;
      /** 1 シートの自由図形の「path 命令の数＋guide（`avLst`＋`gdLst`）の数」の合計。超えたら、以後の自由図形は形を持たせず `shape-simplified` */
      maxSheetGeomUnits: number;
      /** 図形 1 つの調整値（`prstGeom` の `avLst`）の数。超えた分は読まない */
      maxShapeAdjustments: number;
      /** styles.xml の罫線（`<border>`）の数。超えた位置の罫線を参照するセルは罫線なし（ADR-0029 §決定 9） */
      maxBorders: number;
      /** styles.xml のフォント（`<font>`）の数。超えた位置のフォントを参照するセルは装飾なし（ADR-0030 §決定 11） */
      maxFonts: number;
      /** styles.xml の塗り（`<fill>`）の数。超えた位置の塗りを参照するセルは塗りなし（ADR-0030 §決定 11） */
      maxFills: number;
      /** 重複を除いた「見える書式」（塗り・文字の装飾）の数。超えた書式のセルは書式なし（ADR-0030 §決定 11） */
      maxCellFormats: number;
      /** グラフ 1 つの系列数。超えたグラフは `too-large`（破線枠。ADR-0031 §決定 9） */
      maxChartSeries: number;
      /** グラフ 1 つの全系列の点の合計（系列数 × 項目の数）。超えたグラフは `too-large` */
      maxChartPoints: number;
      /** グラフ 1 つの系列名・項目・タイトルの文字数の合計。超えたグラフは `too-large` */
      maxChartTextChars: number;
      /** 1 シートで読んだグラフの点の合計。超えたら、以後そのシートの新しいグラフは `too-large` */
      maxSheetChartPoints: number;
  }
  /**
   * 1 シートの XML の上限 = 160 MiB。
   * セル 1 つの XML は `<c r="AB1234" s="12" t="s"><v>12345</v></c>` で 40〜80 バイト。
   * 200 万セル級のシートでも 80 バイト × 200 万 = 160 MB に収まる見積もり（強制するセル数の上限ではない）。
   * これより大きいと JS の文字列（2 バイト文字を含むと 2 倍）でブラウザのメモリを圧迫するので読まない。
   */
  export const MAX_SHEET_XML_BYTES: number;
  /** 共有文字列の上限 = 64 MiB（同じ文字は 1 回しか入らないので、シート本体より小さく見積もる） */
  export const MAX_SHARED_STRINGS_BYTES: number;
  /** 小さな部品の上限 = 16 MiB（styles.xml は書式が数万あっても数 MB。それ以上は異常） */
  export const MAX_SMALL_PART_BYTES: number;
  /** zip の中のファイル数の上限（Excel のブックは数十個。画像が多くても数千） */
  export const MAX_ENTRIES = 10000;
  /** シート数の上限（実データは 26 枚。Excel も実用上 1,000 枚で足りる） */
  export const MAX_SHEETS = 1000;
  /** 1 シートの結合セル数の上限（200 万セルのシートで全セルが 4 セル結合でも 50 万） */
  export const MAX_MERGES = 500000;
  /** 1 シートの図の数の上限（線・矢印は数えない。設計書の図は数十個。細工したファイルで DOM が膨らまない水準） */
  export const MAX_DRAWINGS = 1000;
  /** 1 シートの線・矢印の数の上限（図とは別に数える。線の多い構成図で画像・図形が押し出されないように） */
  export const MAX_LINES = 1000;
  /** 図形 1 つの文字数の上限（長い文章を入れた図形でも収まる水準。超えた分は切る） */
  export const MAX_SHAPE_TEXT_CHARS = 10000;
  /** 図形 1 つの段落数の上限 */
  export const MAX_SHAPE_PARAGRAPHS = 500;
  /** 図形 1 つのラン（`a:br` を含む）の数の上限 */
  export const MAX_SHAPE_RUNS = 2000;
  /** 1 シートの図形の文字の合計の上限（文字数＋ラン＋段落。図形 1,000 個×1,000 文字の水準。細工したファイルで DOM が膨らまない） */
  export const MAX_SHEET_TEXT_CHARS = 1000000;
  /** グループの入れ子の深さの上限（走査は XML の大きさ×この値で抑えられる。設計書の構成図は 2〜3 段） */
  export const MAX_GROUP_DEPTH = 8;
  /** 画像 1 枚の上限 = 20 MiB（画面写しは数百 KB〜数 MB） */
  export const MAX_IMAGE_BYTES: number;
  /** 1 シートの画像の合計の上限 = 128 MiB（blob にして持つ量の上限。細工した zip 爆弾でメモリが尽きない水準） */
  export const MAX_SHEET_IMAGE_BYTES: number;
  /** 画像の 1 辺の上限 = 16,384 px */
  export const MAX_IMAGE_SIDE_PX = 16384;
  /** 画像 1 枚の画素数の上限 = 64 MP（64 × 1,048,576。8,192 × 8,192 まで） */
  export const MAX_IMAGE_PIXELS: number;
  /** 図の名前・説明（`cNvPr` の name・descr）の文字数の上限。超えた分は切る */
  export const MAX_DRAWING_TEXT = 1000;
  /** 1 色の変換（`a:lumMod` などの子）の数の上限。超えた分は読まない（ADR-0028 §決定 7。Excel の画面で付けるのは数個） */
  export const MAX_COLOR_TRANSFORMS = 16;
  /** グラデーションの stop（`a:gs`）の数の上限。超えた分は読まない（Excel の画面で作れるのは 10 まで。使うのは ISSUE-95） */
  export const MAX_GRADIENT_STOPS = 16;
  /**
   * 1 図形の guide の数の上限（`MAX_GEOM_GUIDES`）と path 命令の数の上限（`MAX_GEOM_COMMANDS`）。超えたら形を描かず代替。
   * 評価器も同じ値で守るため、定義は geom の `defs.ts` に置き、ここでは引いて出し直す（定数の置き場の唯一の例外）。
   */
  export { MAX_GEOM_COMMANDS, MAX_GEOM_GUIDES };
  /**
   * 1 シートの自由図形の「path 命令の数＋guide の数」の合計の上限（ADR-0028 §決定 7。guide も数える＝命令 0 個で guide だけの自由図形を並べて保持量が膨らむのを防ぐ）。
   * 超えたら以後の自由図形は形を持たせず代替
   */
  export const MAX_SHEET_GEOM_UNITS = 200000;
  /** 図形 1 つの調整値（`avLst`）の数の上限（preset の最大は 8。ADR-0028 §決定 7）。超えた分は読まない */
  export const MAX_SHAPE_ADJUSTMENTS = 16;
  /** styles.xml の罫線（`<border>`）の数の上限（Excel のセルの書式の上限 64,000 を丸めた値。ADR-0029 §決定 9）。超えた分は読まない */
  export const MAX_BORDERS = 65536;
  /** styles.xml のフォント（`<font>`）の数の上限（`MAX_BORDERS` と同じ考え方。ADR-0030 §決定 11）。超えた分は読まない */
  export const MAX_FONTS = 65536;
  /** styles.xml の塗り（`<fill>`）の数の上限（`MAX_BORDERS` と同じ考え方。ADR-0030 §決定 11）。超えた分は読まない */
  export const MAX_FILLS = 65536;
  /** 重複を除いた見える書式の数の上限（＝生成 CSS の規則の数。ADR-0030 §決定 11）。超えた書式のセルは書式なし */
  export const MAX_CELL_FORMATS = 4096;
  /** グラフ 1 つの系列数の上限（Excel のグラフ 1 つの系列数の上限。ADR-0031 §決定 9） */
  export const MAX_CHART_SERIES = 255;
  /** グラフ 1 つの全系列の点の合計の上限（系列 10・点 1,000 のグラフの 10 倍の水準） */
  export const MAX_CHART_POINTS = 100000;
  /** グラフ 1 つの系列名・項目・タイトルの文字数の合計の上限 */
  export const MAX_CHART_TEXT_CHARS = 200000;
  /** 1 シートで読んだグラフの点の合計の上限（グラフ 10 個分。超えたら以後の新しいグラフは破線枠） */
  export const MAX_SHEET_CHART_POINTS = 1000000;
  /** 既定の上限の組 */
  export const DEFAULT_LIMITS: Readonly<XlsxLimits>;
  /** 書式コードの長さの上限（文字数）。非信頼の styles.xml から来る書式を ssf へ渡す前の ReDoS 対策 */
  export const MAX_FORMAT_CODE_LENGTH = 255;
  /** Excel の行数の上限（1,048,576）。これを超える行番号は不正として読まない */
  export const MAX_ROWS = 1048576;
  /** Excel の列数の上限（XFD = 16,384）。これを超える列は不正として読まない */
  export const MAX_COLS = 16384;
  /** 1 つのセルの文字数の上限（Excel の上限 32,767）。超えた分は切る */
  export const MAX_CELL_TEXT = 32767;
  /** 列幅（`<col width>`・`defaultColWidth`・`baseColWidth`、文字数）の上限。Excel の上限 255。超えた値は丸める */
  export const MAX_COL_WIDTH = 255;
  /** 行高（`<row ht>`・`defaultRowHeight`、pt）の上限。Excel の上限 409.5。超えた値は丸める */
  export const MAX_ROW_PT = 409.5;
}

declare module 'xsheet/_/geom/defs' {
  /**
   * geom が持つ共通の定義（演算子の表・上限・経路の型）。geom は他の層に依存しないので、ここに置いて core・render から引く。
   * InterHub では評価器（viewers/xlsx-geom.ts）がこれらを core（`lib/xlsx/drawing-geom.ts`・`limits.ts`）と
   * 描画側（`viewers/xlsx-shapes.ts`）から import していた（層の逆転）。
   * - `GUIDE_ARITY`：core の `drawing-geom.ts`（自由図形の字句の検査）が引く（M1-2）。
   * - `MAX_GEOM_GUIDES`・`MAX_GEOM_COMMANDS`：core の `limits.ts` が引いて `DEFAULT_LIMITS` に入れる（M1-2）。
   * - `Pt`・`PathSeg`：render の `xlsx-shapes.ts` が引く（M2-1）。
   * 形の表（`preset-geom.json`）は読まない（core から引いても表がバンドルに入らないように、評価器と分ける）。
   */
  /** 式の演算子ごとの引数の数（DrawingML の guide の 17 種。読み取り側の字句の検査と評価器で同じ表を使う） */
  export const GUIDE_ARITY: ReadonlyMap<string, number>;
  /** 1 図形の guide の数の上限（自由図形の `avLst`＋`gdLst`。preset の最大は gear9 の 244。ADR-0028 §決定 7）。超えたら形を描かず代替 */
  export const MAX_GEOM_GUIDES = 512;
  /** 1 図形の path 命令の数の上限（全 path の合計。ADR-0028 §決定 7）。超えたら形を描かず代替 */
  export const MAX_GEOM_COMMANDS = 4096;
  /** 点（px） */
  export interface Pt {
      x: number;
      y: number;
  }
  /** 経路の区間：直線（`L`）か 3 次ベジェ（`C`。制御点 1・2 と終点） */
  export type PathSeg = {
      t: 'L';
      x: number;
      y: number;
  } | {
      t: 'C';
      x1: number;
      y1: number;
      x2: number;
      y2: number;
      x: number;
      y: number;
  };
}

declare module 'xsheet/_/core/types' {
  /** xlsx の読み取り部品の公開の型（画面に依存しない） */
  import type { Rgba } from 'xsheet/_/core/color';
  import type { Theme } from 'xsheet/_/core/theme';
  /** 読めない理由の種類 */
  export type XlsxErrorKind = 'not-zip' | 'encrypted' | 'broken' | 'too-large' | 'unsupported';
  /**
   * xlsx を読めないときの例外。種類は `kind` で決まる。
   * メッセージには、原本の中身（セルの値・シート名）を入れない。
   */
  export class XlsxError extends Error {
      readonly kind: XlsxErrorKind;
      constructor(kind: XlsxErrorKind, detail?: string);
  }
  /** シート 1 枚の情報 */
  export interface SheetInfo {
      /** 1 始まり・workbook の順（非表示も数える）（シートの識別番号） */
      index: number;
      name: string;
      /** state が hidden または veryHidden */
      hidden: boolean;
  }
  /**
   * セルの型（`<c t>` 由来）。
   * - `n`：数値（t 無し・`n`。書式で日付や通貨の表示になる数値も `n`）
   * - `s`：文字列（`s`・`str`・`inlineStr`）
   * - `b`：真偽、`e`：エラー、`d`：日付（`t="d"`・ISO 8601）
   */
  export type CellKind = 'n' | 's' | 'b' | 'e' | 'd';
  /** 1 つのセルの描き方。結合で隠れるセルは null（CellView ではなく行の配列に入る） */
  export interface CellView {
      /** 表示形式を当てた後の文字（Excel で見える文字）。文字列が空（数式の結果が "" など）のセルは "" */
      text: string;
      /**
       * セルの型。リーダーは必ず付ける（結合セルの左上を除く）。
       * 文字列で `text` が "" のセルは、数式の結果が空文字など「値のあるセル」を表す（Excel は右隣への文字のはみ出しを止める）。
       * 書式だけのセル・値の無いセルは CellView にならない（undefined）。
       */
      kind?: CellKind;
      colSpan: number;
      rowSpan: number;
  }
  /** セルの罫線の線種（SpreadsheetML の `ST_BorderStyle` から `none` を除いた 13 種。ADR-0029 §決定 3） */
  export type BorderStyle = 'thin' | 'hair' | 'dotted' | 'dashed' | 'dashDot' | 'dashDotDot' | 'medium' | 'mediumDashed' | 'mediumDashDot' | 'mediumDashDotDot' | 'slantDashDot' | 'thick' | 'double';
  /** 罫線の 1 辺：線種と色。`color` が null なら自動（描く側が墨色にする） */
  export interface BorderEdge {
      style: BorderStyle;
      color: Rgba | null;
  }
  /**
   * セル 1 つの罫線（ADR-0029 §決定 2）。線の無い辺は undefined。
   * `diagonalUp`（左下→右上）・`diagonalDown`（左上→右下）は斜線の向きで、線種・色は `diagonal` の辺。
   */
  export interface CellBorder {
      left?: BorderEdge;
      right?: BorderEdge;
      top?: BorderEdge;
      bottom?: BorderEdge;
      diagonal?: BorderEdge;
      diagonalUp: boolean;
      diagonalDown: boolean;
  }
  /**
   * セルの書式（塗り・文字の装飾。ADR-0030 §決定 2）。見える書式だけを持つ（見えない xf は「書式なし」＝番号を付けない）。
   * 色は RGB の整数と不透明度（塗りの透明度は持たず、常に不透明）。ファイルの文字列は持たない。
   * `SheetRow.formats` の値・`XlsxSheet.formats` の位置が指す。
   */
  export interface CellFormat {
      /** 塗りの色（模様・グラデーションは代表色の 1 色）。無ければ塗りなし */
      fill?: Rgba;
      /** 文字の色。無ければ自動（不透明な `#000000` も自動に寄せてある） */
      color?: Rgba;
      bold?: true;
      italic?: true;
      /** 取り消し線 */
      strike?: true;
      /** 下線（`singleAccounting` は `single`・`doubleAccounting` は `double`） */
      underline?: 'single' | 'double';
      /** 文字サイズ ÷ 既定のフォント（先頭の `<font>`）のサイズ。0.5〜4・0.01 単位。1 なら持たない */
      sizeRatio?: number;
  }
  /** 展開した 1 行 */
  export interface SheetRow {
      /** 1 始まりの行番号 */
      r: number;
      /** 長さ colCount。undefined＝空のセル、null＝結合で隠れる */
      cells: (CellView | null | undefined)[];
      /**
       * 罫線の番号（長さは `cells` と同じ）。値は `XlsxSheet.borders` の位置、undefined＝罫線なし。
       * 値の無い（書式だけの）セル・結合で隠れるセルも含む。罫線のあるセルが 1 つも無い行は持たない。
       */
      borders?: (number | undefined)[];
      /**
       * 書式の番号（長さは `cells` と同じ）。値は `XlsxSheet.formats` の位置、undefined＝書式なし。
       * 値の無い（書式だけの）セル・結合で隠れるセルも含む。セルの要素の無い位置には、`customFormat="1"` の行の書式、
       * 無ければ列の書式（`<col style>`）が当たる（セルの要素がある位置は、その `s` だけ。ADR-0030 §決定 7）。
       * 書式のあるセルが 1 つも無い行は持たない。
       */
      formats?: (number | undefined)[];
  }
  /** `<cols>` の `<col min max width hidden>` 1 件（`min`〜`max` は 1 始まり・両端を含み、展開せずに持つ） */
  export interface ColLayout {
      min: number;
      max: number;
      /** 列幅（文字数。`<col width>` の値そのまま）。無ければ null＝既定の列幅 */
      width: number | null;
      hidden: boolean;
      /** 列の書式（`<col style>`＝cellXfs の位置）。無い・不正なら持たない（ADR-0030 §決定 7。`getRows` が書式の番号に解く） */
      xf?: number;
  }
  /**
   * `<row>` の寸法の指定。同じ指定が続く行は 1 件にまとめる（ランレングス。`min`〜`max` は 1 始まり・両端を含む）。
   * 全行に同じ `ht` があっても 1 件になる。
   */
  export interface RowLayout {
      min: number;
      max: number;
      /** 行高（pt）。無ければ null＝シートの既定の行高 */
      ht: number | null;
      hidden: boolean;
      /** `customHeight="1"`（`ht` があるときだけ true。Excel はこれが無い `ht` を自動調整の値として扱う） */
      custom: boolean;
  }
  /**
   * シートの寸法（Excel の列幅・行高・非表示）。px への換算は `geometry.ts` の `createGeometry` が行う。
   * 列・行の指定は「指定のある番号だけ」の疎な配列（全要素の配列は作らない）。
   */
  export interface SheetLayout {
      /** 列の指定（`<cols>` の出現順。重なりの扱いは geometry 側） */
      cols: readonly ColLayout[];
      /** `<sheetFormatPr defaultColWidth>`（文字数・余白を含む値）。無ければ null */
      defaultColWidth: number | null;
      /** `<sheetFormatPr baseColWidth>`（余白を含まない文字数）。無ければ 8 */
      baseColWidth: number;
      /** `<sheetFormatPr defaultRowHeight>`（pt）。無ければ 15 */
      defaultRowHeight: number;
      /** `<sheetFormatPr customHeight="1">`。true なら `defaultRowHeight` がそのまま既定の行高 */
      customHeight: boolean;
      /** `<sheetFormatPr zeroHeight>`。true なら、`<row>` の無い行は既定で非表示（高さ 0） */
      zeroHeight: boolean;
      /** 行の指定（行番号の昇順・重複なし・まとめ済み）。`zeroHeight` のときは `<row>` のあるすべての行を持つ */
      rows: readonly RowLayout[];
      /** 最大桁幅 MDW（px）。styles.xml の先頭の font から決める。無ければ 7 */
      mdw: number;
      /** 先頭の font の既定の行高（px・96dpi）。実測表に無い・font が無ければ null（`defaultRowHeight` を使う） */
      fontRowPx: number | null;
  }
  /** 既定の最大桁幅 MDW（px。Calibri 11 相当） */
  export const DEFAULT_MDW = 7;
  /** 寸法の指定が何も無いシートの layout（既定値のみ） */
  export function emptyLayout(mdw?: number, fontRowPx?: number | null): SheetLayout;
  /** 開いた 1 シート */
  export interface XlsxSheet {
      /** 表示する行数（値のある最後の行と結合の最後の行。空のシートは 0） */
      readonly rowCount: number;
      /** 表示する列数（同じく最後の列） */
      readonly colCount: number;
      /** 列幅・行高・非表示の寸法（Excel の値。px への換算は `createGeometry`） */
      readonly layout: SheetLayout;
      /** 罫線の表（`SheetRow.borders` の値が指す。ブックで 1 つを全シートが共有する。罫線の無いブックは空） */
      readonly borders: readonly CellBorder[];
      /** 書式の表（`SheetRow.formats` の値が指す。ブックで 1 つを全シートが共有する。見える書式の無いブックは空） */
      readonly formats: readonly CellFormat[];
      /**
       * 行 r（1 始まり）の「行の書式」（`<row customFormat="1" s>`）の書式の番号。表の範囲（`rowCount`）の外の行も引く。
       * `customFormat` が無い・`<row>` が無い・見える書式の無いブックは undefined、書式なしの xf は -1
       * （書式なしの行の書式は、列の書式も打ち消す）。`getRows` の `formats` は表の範囲の中だけなので、余白の位置の塗りに使う。
       */
      rowFormat(r: number): number | undefined;
      /**
       * 列 c（1 始まり）の「列の書式」（`<col style>`。重なった指定は文書で先のもの）の書式の番号。表の範囲（`colCount`）の外の列も引く。
       * 指定の無い列・書式なしの xf・見える書式の無いブックは -1。
       */
      colFormat(c: number): number;
      /**
       * 行 first〜last（1 始まり・両端を含む）を展開して返す。行の無い番号も空の行として返す。
       * 結合の左上が範囲より上にあるときは、範囲の最初の行に左上の文字を置き、rowSpan を範囲内に縮める
       * （範囲の上端で切れても表が崩れないように）。展開した行は保持しない。
       */
      getRows(first: number, last: number): SheetRow[];
  }
  /** アンカーの 1 点（セルの番号とセル内のずれ）。番号は 0 始まり（drawing の値そのまま。A1 のセルが col 0・row 0） */
  export interface AnchorPoint {
      col: number;
      /** セルの左端からのずれ（EMU。914400 EMU＝1 インチ） */
      colOff: number;
      row: number;
      /** セルの上端からのずれ（EMU） */
      rowOff: number;
  }
  /**
   * 図の位置（アンカー）。`editAs`（行・列を動かした時の振る舞い）は読まない（保存時の値が既にその結果）。
   * - twoCell：`from`〜`to`
   * - oneCell：`from` を起点に `cx`×`cy`（EMU）の大きさ
   * - absolute：シートの左上からの絶対位置 `x`・`y`（EMU）に `cx`×`cy`
   */
  export type DrawingAnchor = {
      type: 'twoCell';
      from: AnchorPoint;
      to: AnchorPoint;
  } | {
      type: 'oneCell';
      from: AnchorPoint;
      cx: number;
      cy: number;
  } | {
      type: 'absolute';
      x: number;
      y: number;
      cx: number;
      cy: number;
  };
  /**
   * 図の種類。
   * - image：画像（`xdr:pic`）。描ける形式かは `readImage` が決める
   * - chart：グラフ（`graphicFrame`。SmartArt 以外の `graphicFrame` もここに入れる）
   * - smartart：SmartArt（`graphicFrame` の `graphicData` が diagram）
   * - shape：図形（`xdr:sp`）／ textbox：テキストボックス（`cNvSpPr txBox`）／ group：グループ（`xdr:grpSp`）
   * - line：線・矢印（`xdr:cxnSp`、または線の preset の `xdr:sp`）。図とは別に数える（ADR-0027 §決定 3・6）
   */
  export type DrawingKind = 'image' | 'chart' | 'smartart' | 'shape' | 'textbox' | 'group' | 'line';
  /** 文字の縦位置（`a:bodyPr anchor`。`t`・`ctr`・`b` 以外は `t`） */
  export type TextAnchor = 't' | 'ctr' | 'b';
  /** 段落の揃え（`a:pPr algn`。この 5 つ以外は `l`） */
  export type TextAlign = 'l' | 'ctr' | 'r' | 'just' | 'dist';
  /** 文字の向き（`a:bodyPr vert`。この 7 つ以外は `horz`＝横書き） */
  export type TextVert = 'horz' | 'vert' | 'vert270' | 'wordArtVert' | 'eaVert' | 'mongolianVert' | 'wordArtVertRtl';
  /**
   * 図形・線・文字の色（ISSUE-95。ADR-0028 §決定 6）。RGB の整数（0〜0xFFFFFF）と不透明度（0〜1）だけを持つ（ファイルの文字列は持たない）。
   * グラデーション・模様・画像の塗りは代表色の 1 色にしてある。
   */
  export interface DrawingColor {
      /** RGB（`0xRRGGBB`） */
      rgb: number;
      /** 不透明度（0＝透明〜1＝不透明） */
      alpha: number;
  }
  /** 段落の中の 1 つ（文字のラン、または `a:br` の改行） */
  export interface DrawingTextRun {
      /** true なら `a:br`（改行）。`text` は空 */
      br: boolean;
      /** 本文（`a:t`。実体参照を戻し、タブと改行以外の制御文字を除いたもの。改行のときは空） */
      text: string;
      /** 文字の大きさ（pt）。`a:rPr sz`（無ければ 11）を 1〜400 に丸めた値 */
      sizePt: number;
      bold: boolean;
      italic: boolean;
      /** 下線（`u` が `none` 以外） */
      underline: boolean;
      /** 取り消し線（`strike` が `noStrike` 以外） */
      strike: boolean;
      /**
       * 文字の色（ISSUE-95）。`a:rPr` の `a:solidFill`・`a:gradFill`（代表色）→ 図形の `a:fontRef` の子の色 → テーマの `tx1`。
       * `a:noFill` は透明（不透明度 0）。リーダーは常に持たせる（型は既存の組み立て側を壊さないよう省略可）
       */
      color?: DrawingColor;
  }
  /** 図形の文字の 1 段落 */
  export interface DrawingTextParagraph {
      /** ランと改行を出現順に並べたもの（空の段落は空配列） */
      runs: DrawingTextRun[];
      align: TextAlign;
      /**
       * 段落の先頭に足す箇条書きの文字（無ければ null）。`buChar` はその文字、`buAutoNum` は同じ種類が続く段落を
       * 1（`startAt` があればその値）から数えた「1.」の形に、lib 側で整形済み
       */
      bullet: string | null;
      /** 段落記号の文字の大きさ（pt。`a:endParaRPr sz`・無ければ 11）。空の段落の行の高さに使う */
      endSizePt: number;
  }
  /** 図形の文字（`xdr:txBody`）。大きさは pt、余白は EMU、列挙値はそのまま（px への換算は表示側） */
  export interface DrawingText {
      /** 段落（上限 500。文字が 1 つも無い図形は text 自体を持たない） */
      paragraphs: DrawingTextParagraph[];
      /** 縦位置（`a:bodyPr anchor`） */
      anchor: TextAnchor;
      /** 内側の余白（EMU。`lIns`・`tIns`・`rIns`・`bIns`。既定 91440・45720・91440・45720） */
      insets: {
          l: number;
          t: number;
          r: number;
          b: number;
      };
      /** 折り返さない（`a:bodyPr wrap="none"`） */
      noWrap: boolean;
      /** 縮小の倍率（`a:normAutofit fontScale`÷100000。縮めないときは 1。0.01〜1 に丸める） */
      fontScale: number;
      /** 文字の向き（`a:bodyPr vert`） */
      vert: TextVert;
  }
  /** 線の形（`a:prstGeom prst`。この 10 通り以外・`custGeom` は `line`＝直線で代替） */
  export type LineGeom = 'line' | 'straightConnector1' | 'bentConnector2' | 'bentConnector3' | 'bentConnector4' | 'bentConnector5' | 'curvedConnector2' | 'curvedConnector3' | 'curvedConnector4' | 'curvedConnector5';
  /** 線種（`a:prstDash val`。この 11 通り以外は `solid`） */
  export type LineDash = 'solid' | 'dot' | 'dash' | 'lgDash' | 'dashDot' | 'lgDashDot' | 'lgDashDotDot' | 'sysDash' | 'sysDot' | 'sysDashDot' | 'sysDashDotDot';
  /** 矢じりの種類（`type="none"`・無い場合は矢じりなしなので、この型の値を持たず null）。未知の値は `triangle` */
  export type ArrowType = 'triangle' | 'stealth' | 'diamond' | 'oval' | 'arrow';
  /** 矢じりの大きさ（`w`・`len`。無い・未知の値は `med`） */
  export type ArrowSize = 'sm' | 'med' | 'lg';
  /** 線の端の矢じり（`a:headEnd`＝始点・`a:tailEnd`＝終点） */
  export interface LineEnd {
      type: ArrowType;
      w: ArrowSize;
      len: ArrowSize;
  }
  /** 線・矢印の形と見た目（色は `color`。矢じりも同じ色） */
  export interface DrawingLine {
      geom: LineGeom;
      /** 調整値 adj1〜adj3（`a:gd name="adjN" fmla="val N"` の N。1/100000 の割合。無ければ 50000。-1,000,000〜1,000,000 に丸める） */
      adj: [number, number, number];
      /** 太さ（EMU）。`a:ln w`、無ければ `lnRef idx` 1・2・3 → 6350・12700・19050、どちらも無ければ 0（表示側が 1px にする） */
      widthEmu: number;
      dash: LineDash;
      /** 始点の矢じり（なしは null） */
      head: LineEnd | null;
      /** 終点の矢じり（なしは null） */
      tail: LineEnd | null;
      /**
       * 線の色（ISSUE-95）。`a:ln` の `a:solidFill`・`a:gradFill`・`a:pattFill`（代表色）→ `a:lnRef` の子の色 → テーマの `tx1`。
       * リーダーは常に持たせる（型は既存の組み立て側を壊さないよう省略可）
       */
      color?: DrawingColor;
  }
  /** 形の式・命令の引数：整数、または組み込み変数・guide の名前（英数字 16 文字まで）。名前は `Map` のキーにだけ使う */
  export type GeomArg = number | string;
  /** 自由図形の path の塗りの濃淡の種類（`a:path fill`。`none`＝塗らない） */
  export type GeomPathFill = 'none' | 'norm' | 'lighten' | 'lightenLess' | 'darken' | 'darkenLess';
  /** 自由図形の path 1 つ（`a:path`）。命令は `['M', x, y]`・`['L', x, y]`・`['A', wR, hR, stAng, swAng]`・`['Q', x1, y1, x, y]`・`['C', x1, y1, x2, y2, x, y]`・`['Z']` */
  export interface CustGeomPath {
      /** 座標系の幅・高さ（`a:path w`・`h`。無い・0 なら形の座標系＝EMU のまま） */
      w?: number;
      h?: number;
      /** 塗りの濃淡の種類（`a:path fill`。無ければ省く＝`norm`） */
      fill?: GeomPathFill;
      /** 線を引くか（`a:path stroke`。無ければ省く＝引く） */
      stroke?: boolean;
      /** 命令の列（`moveTo`・`lnTo`・`arcTo`・`quadBezTo`・`cubicBezTo`・`close`） */
      c: ReadonlyArray<readonly [op: string, ...args: GeomArg[]]>;
  }
  /**
   * 自由図形の定義（`a:custGeom`）。式の評価器（geom の `evaluate.ts` の `GeomDef`）にそのまま渡せる形。
   * 式は演算子の表（17 種）と引数の数、引数は整数か名前の字句を確かめ済み。
   */
  export interface CustGeomDef {
      /** 調整値の既定（`a:avLst` の `a:gd name fmla="val N"`）：`[名前, 値]` */
      av: ReadonlyArray<readonly [string, number]>;
      /** guide の列（`a:gdLst`）：`[名前, 演算子, 引数…]` */
      gd: ReadonlyArray<readonly [name: string, op: string, ...args: GeomArg[]]>;
      /** 文字領域（`a:rect` の l・t・r・b。無ければ省く＝矩形全体） */
      rect?: readonly GeomArg[];
      /** path の列（`a:pathLst`） */
      paths: readonly CustGeomPath[];
  }
  /**
   * 図形（`xdr:sp`。線として扱わないもの）の形。
   * - preset：`a:prstGeom prst`＋調整値（`a:avLst` の `val N` を名前ごとに。16 個まで。無ければ空）。表に無い名前もそのまま持つ（判定は描画側）
   * - custom：自由図形（`a:custGeom`）の定義。字句の誤り・上限超えのときは持たせない（形なし）
   */
  export type ShapeGeom = {
      type: 'preset';
      prst: string;
      adj: ReadonlyArray<readonly [string, number]>;
  } | {
      type: 'custom';
      def: CustGeomDef;
  };
  /** 図形の枠線の太さと線種（線と同じ規則。色は読まない） */
  export interface ShapeOutline {
      /** 太さ（EMU）。`a:ln w`、無ければ `lnRef idx` 1・2・3 → 6350・12700・19050、どちらも無ければ 0（表示側が 1px にする） */
      widthEmu: number;
      /** 線種（`a:prstDash val`。無い・未知は `solid`） */
      dash: LineDash;
  }
  /** グループの子の座標系の矩形（EMU。位置は負でもよい。大きさ `cx`・`cy` は 0 以上） */
  export interface GroupRect {
      x: number;
      y: number;
      cx: number;
      cy: number;
  }
  /** グループ 1 段の座標系（外側のグループから順に並べる） */
  export interface GroupLevel {
      /** このグループの子が使う座標系（`a:chOff`・`a:chExt`。`cx`・`cy` は 1 以上） */
      ch: GroupRect;
      /**
       * このグループ自身の位置と大きさ（`a:off`・`a:ext`。親のグループの子の座標系）。
       * 最も外側のグループは null（図のアンカーの矩形がその位置になる）
       */
      own: GroupRect | null;
  }
  /**
   * グループの中の図の位置（ADR-0027 §決定 5）。`SheetDrawing.anchor` は最も外側のグループのアンカーで、
   * 表示側がそこ（px）から `levels` を外側から順にたどり、`xfrm` を写して矩形にする。
   */
  export interface DrawingGroup {
      /** 外側のグループから順の座標系（1 段以上。深さの上限は `maxGroupDepth`） */
      levels: readonly GroupLevel[];
      /** 図自身の `a:xfrm`（`a:off`・`a:ext`。最も内側のグループの子の座標系。回転前の矩形そのもの） */
      xfrm: GroupRect;
  }
  /** シートの図 1 つ（一覧は drawing XML の順＝後のものが上に重なる） */
  export interface SheetDrawing {
      anchor: DrawingAnchor;
      kind: DrawingKind;
      /** `cNvPr` の name（無ければ空） */
      name: string;
      /** `cNvPr` の descr＝代替テキスト（無ければ空） */
      descr: string;
      /**
       * kind が image のとき、画像の zip 内のパス（`readImage` に渡す）。
       * 外部リンク（`TargetMode="External"`・`r:link`）・zip に無いもの・`r:embed` の無いものは null（辿らない。表示側は「表示できない画像」の枠にする）。
       * SVG 併記（`asvg:svgBlip`）は SVG を使わず、PNG の予備（`a:blip r:embed`）のパスを持つ。image 以外は null
       */
      mediaPath: string | null;
      /**
       * kind が chart のとき、グラフの部品（`xl/charts/chartN.xml`）の zip 内のパス（`readChart` に渡す。ISSUE-101・ADR-0031 §決定 2）。
       * `graphicData` の `c:chart r:id` を drawing の rels で解いたもの。外部リンク・zip に無いもの・chartex・`r:id` の無いものは持たない（表示側は破線枠のまま）
       */
      chartPath?: string;
      /** 文字のある図形（kind が shape・textbox で、ランの文字が 1 つ以上ある）の文字。無い図形・それ以外の kind は持たない */
      text?: DrawingText;
      /** 回転（`a:xfrm rot`。1/60000 度・0〜21,599,999 に正規化）。図形（`xdr:sp`。文字の有無を問わない）と線が持つ */
      rot?: number;
      /**
       * `a:xfrm` の `a:ext`（回転前の大きさ。EMU）。回転のある（`rot` が 0 でない）トップレベルの図形だけが持つ（`a:ext` が読めない・0 以下なら持たない）。
       * Excel のアンカーは、回転が 45° 付近でも回転後の外接矩形ではなく、回転前の大きさ（縦横を入れ替えた場合もある）で書かれるため、入れ替えの判定に使う（ISSUE-93）
       */
      xfrmExt?: {
          cx: number;
          cy: number;
      };
      /** 左右反転（`a:xfrm flipH`）。図形と線が持つ */
      flipH?: boolean;
      /** 上下反転（`a:xfrm flipV`）。図形と線が持つ */
      flipV?: boolean;
      /** Excel で枠線が見える（色は読まず有無だけ）。図形（`xdr:sp`。文字の有無を問わない）が持つ */
      hasOutline?: boolean;
      /** Excel で塗りがある（色は読まず有無だけ）。図形が持つ */
      hasFill?: boolean;
      /** 図形の形（`a:prstGeom`・`a:custGeom`）。図形が持つ。`prstGeom` も読める `custGeom` も無い図形・自由図形の字句の誤り・上限超えは持たない */
      geom?: ShapeGeom;
      /** 図形の枠線の太さと線種。図形が持つ（枠線が見えるかは `hasOutline`） */
      outline?: ShapeOutline;
      /**
       * 塗りの色（ISSUE-95）。塗りがある（`hasFill`）図形だけが持つ。`spPr` 直下の塗り（代表色）→ `a:fillRef` の子の色 → テーマの `lt1`。
       * 塗り・枠線の「有無」の規則は変えない（ADR-0027）
       */
      fillColor?: DrawingColor;
      /** 枠線の色（ISSUE-95）。枠線が見える（`hasOutline`）図形だけが持つ。`a:ln` の塗り（代表色）→ `a:lnRef` の子の色 → テーマの `tx1` */
      outlineColor?: DrawingColor;
      /** kind が line のとき、線の形と見た目 */
      line?: DrawingLine;
      /**
       * グループの中の図（展開した子）が持つ、グループの子の座標系での位置。無ければトップレベルの図で、`anchor` がそのまま矩形。
       * 持つ図の `anchor` は最も外側のグループのもの。入れ子の中で展開できなかったグループ（深さの上限超え・`chExt` が 0）も、これを持つ破線枠（kind が group）として出る
       */
      group?: DrawingGroup;
  }
  /**
   * 図を読めなかった／全部は出せなかった理由（表示側がバナーにする。表は必ず出る）。
   * - part-too-large：drawing XML・rels が展開後の上限（16MiB）を超えた＝図を出さない
   * - broken：図の部品が読めなかった（zip の壊れ）＝図を出さない
   * - too-many：図が上限（1,000）、または線・矢印が上限（1,000。図とは別に数える）を超えた＝先頭の分だけ出す
   * - text-truncated：図形の文字を上限で切った・捨てた（1 シートに 1 回）
   * - shape-simplified：自由図形（`a:custGeom`）が上限（1 図形の guide・path 命令、1 シートの path 命令の合計）を超え、形を持たせなかった（1 シートに 1 回。ISSUE-91。バナーにするのは ISSUE-92）
   */
  export type DrawingNotice = 'part-too-large' | 'broken' | 'too-many' | 'text-truncated' | 'shape-simplified';
  /** `openDrawings` の結果 */
  export interface SheetDrawings {
      /** 図の一覧（図の無いシートは空） */
      drawings: readonly SheetDrawing[];
      /** バナーにする理由（重複なし。空なら問題なし） */
      notices: readonly DrawingNotice[];
  }
  /** 描ける画像の形式（先頭バイトで判定） */
  export type ImageFormat = 'png' | 'jpeg' | 'gif' | 'bmp' | 'webp';
  /**
   * `readImage` が画像を出せなかった理由。
   * - unsupported：描けない形式（EMF・WMF・TIFF・SVG・不正・途中で切れた・空）。`label` に分かれば形式名（EMF など）
   * - too-large：1 枚の上限（20MiB）超え（シートの合計の残りが少なくても、先にこちらを判定する）
   * - over-budget：1 シートの画像の合計（128MiB）超え。一度超えたら、以後そのシートの新しい画像はすべてこれ（表示側はバナー）
   * - oversize-px：1 辺 16,384px・64MP 超え
   * - missing：zip に無い
   * - broken：展開に失敗した
   */
  export type ImageFailReason = 'unsupported' | 'too-large' | 'over-budget' | 'oversize-px' | 'missing' | 'broken';
  /** `readImage` の結果。失敗しても例外にしない（表示側は枠にする） */
  export type ImageResult = {
      ok: true;
      format: ImageFormat;
      /** 判定した形式の MIME（Blob の type に使う。拡張子・Content_Types は信じない） */
      mime: string;
      /** px（ヘッダの値） */
      width: number;
      height: number;
      /** 画像のバイト列（Blob にして `<img>` の src にだけ使う） */
      bytes: Uint8Array;
  } | {
      ok: false;
      reason: ImageFailReason;
      label: string | null;
  };
  /**
   * グラフを出せなかった理由（`readChart`・`parseChart` が返す。表示側はどの理由でも破線枠。ADR-0031 §決定 1・9）。
   * - unsupported：G1 に無い種類・構造（面・散布図・3D・組み合わせ・第 2 軸・対数軸など）
   * - too-large：系列数・点の合計・文字の合計・シートの点の合計の上限超え、または部品が展開後の上限超え
   * - broken：XML が壊れている（`c:chartSpace`・`c:chart`・`c:plotArea` が読めない）・部品が zip に無い・展開に失敗した
   * - no-data：描ける系列が 0（保存値の無い系列を除いた結果）
   */
  export type ChartFailReason = 'unsupported' | 'too-large' | 'broken' | 'no-data';
  /** `readChart`・`parseChart` の結果。失敗しても例外にしない */
  export type ChartResult = {
      ok: true;
      chart: ChartModel;
      /** ブックのテーマの 12 色（`readChart` が付ける。既定の系列色を描画側で当てるため。`parseChart` 単体は付けない。ISSUE-103） */
      theme?: Theme;
  } | {
      ok: false;
      reason: ChartFailReason;
  };
  /** グラフの塗り：`none` は `a:noFill`（描かない）、`solid` は色（グラデーション・模様は代表色）。明示が無い所は持たない＝既定の色は描画側 */
  export type ChartPaint = {
      kind: 'none';
  } | {
      kind: 'solid';
      color: Rgba;
  };
  /** グラフの線（`a:ln`）。持つのは明示のあった項目だけ */
  export interface ChartLine {
      /** 線の塗り（`a:noFill` なら `none`） */
      paint?: ChartPaint;
      /** 太さ（EMU。`a:ln w`。0〜20,116,800 に丸める） */
      widthEmu?: number;
      /** 線種（`a:prstDash`。未知の値は `solid`） */
      dash?: LineDash;
  }
  /** 面（`c:spPr`）の塗りと枠線 */
  export interface ChartBox {
      fill?: ChartPaint;
      line?: ChartLine;
  }
  /** 文字の見た目（`a:defRPr`・`a:rPr`・`a:bodyPr`）。明示のあった項目だけ持つ */
  export interface ChartTextStyle {
      /** 文字の大きさ（pt。`sz` ÷ 100 を 1〜400 に丸めた値） */
      sizePt?: number;
      bold?: boolean;
      /** 文字の色（`a:noFill` は不透明度 0） */
      color?: Rgba;
      /** 文字の回転（度。`a:bodyPr rot` ÷ 60000。±90 を超える値＝Excel の「自動」を表す特別値は持たない） */
      rotDeg?: number;
  }
  /** `c:manualLayout`（明示の値だけ持つ。モードの既定は `factor`・対象の既定は `outer`＝描画側が決める） */
  export interface ManualLayout {
      target?: 'inner' | 'outer';
      xMode?: 'edge' | 'factor';
      yMode?: 'edge' | 'factor';
      wMode?: 'edge' | 'factor';
      hMode?: 'edge' | 'factor';
      x?: number;
      y?: number;
      w?: number;
      h?: number;
  }
  /**
   * タイトル（グラフ・軸）。文字（`lines`）は `c:rich` の段落（`a:br` でも分ける）か `c:strRef` の保存値。
   * 文字の無いタイトル（`lines` が空）は Excel の自動のタイトルで、文字は描画側が決める
   */
  export interface ChartTitle {
      lines: readonly string[];
      text: ChartTextStyle;
      /** プロットと重ねる（`c:overlay`） */
      overlay: boolean;
      layout?: ManualLayout;
  }
  /** 凡例 */
  export interface ChartLegend {
      /** 位置（`c:legendPos`。無ければ `r`） */
      pos: 'r' | 'l' | 't' | 'b' | 'tr';
      overlay: boolean;
      /** 出さない項目の位置（`c:legendEntry` の `c:delete` が真のものの `c:idx`） */
      deleted: readonly number[];
      text: ChartTextStyle;
      layout?: ManualLayout;
  }
  /** 目盛の印の語（`c:majorTickMark`・`c:minorTickMark`） */
  export type ChartTickMark = 'none' | 'in' | 'out' | 'cross';
  /** 軸（項目軸 `c:catAx`・`c:dateAx`、数値軸 `c:valAx`）。持つのは明示のあった項目だけ */
  export interface ChartAxis {
      kind: 'cat' | 'date' | 'val';
      /** 軸を消す（`c:delete`） */
      deleted: boolean;
      pos?: 'b' | 'l' | 'r' | 't';
      /** 範囲の明示（`c:scaling` の `c:min`・`c:max`） */
      min?: number;
      max?: number;
      /** 向き（`c:orientation`。`maxMin` は反転。無ければ `minMax`） */
      orientation: 'minMax' | 'maxMin';
      majorUnit?: number;
      minorUnit?: number;
      crosses?: 'autoZero' | 'min' | 'max';
      crossesAt?: number;
      crossBetween?: 'between' | 'midCat';
      /** 主・補助の目盛線（持つ＝線がある。`c:spPr` の線の明示だけ中に持つ） */
      majorGridlines?: ChartLine;
      minorGridlines?: ChartLine;
      majorTickMark?: ChartTickMark;
      minorTickMark?: ChartTickMark;
      tickLblPos?: 'nextTo' | 'low' | 'high' | 'none';
      /** 項目の文字を何個おきに出すか（`c:tickLblSkip`。1 以上） */
      tickLblSkip?: number;
      /** 目盛の文字の表示形式（`c:numFmt`。書式コードは長すぎれば General） */
      numFmt?: {
          code: string;
          sourceLinked: boolean;
      };
      title?: ChartTitle;
      /** 軸の線（`c:spPr` の `a:ln`） */
      line?: ChartLine;
      /** 目盛の文字（`c:txPr`） */
      text: ChartTextStyle;
  }
  /** 折れ線のマーカーの形（`c:symbol`。未知の値は持たない） */
  export type ChartMarkerSymbol = 'auto' | 'circle' | 'dash' | 'diamond' | 'dot' | 'none' | 'picture' | 'plus' | 'square' | 'star' | 'triangle' | 'x';
  /** 系列のマーカー（`c:marker`）。持つのは明示のあった項目だけ */
  export interface ChartMarker {
      symbol?: ChartMarkerSymbol;
      /** 大きさ（pt。`c:size` を 2〜72 に丸めた値） */
      size?: number;
      fill?: ChartPaint;
      line?: ChartLine;
  }
  /** グラフの系列 1 つ */
  export interface ChartSeries {
      /**
       * `c:idx`（無い・負・整数でなければ出現の位置）。`c:order` で並べた後・保存値の無い系列を除いた後も保たれる。
       * 既定の色は配列の位置でなくこの `idx` で決まる前提（ISSUE-102）。凡例の `c:legendEntry` の `c:idx`（`ChartLegend.deleted`）との対応は要確認（ISSUE-104）
       */
      idx: number;
      /** 系列名（`c:tx` の保存値。無ければ空。凡例の「Series n」は描画側） */
      name: string;
      /** 値（長さは `ChartModel.categories` と同じ。数値にならない値・無い点は null） */
      values: readonly (number | null)[];
      /** 系列の塗り・線（`c:spPr`） */
      fill?: ChartPaint;
      line?: ChartLine;
      marker?: ChartMarker;
      smooth?: boolean;
      /** 点ごとの塗りの上書き（`c:dPt`。点の位置 → 塗り） */
      points?: ReadonlyMap<number, ChartPaint>;
      /** `c:numCache` の書式コード（軸の `sourceLinked` で使う。長すぎれば General） */
      formatCode?: string;
  }
  /**
   * 読み取ったグラフ（G1：縦棒・横棒・折れ線・円・ドーナツ。ADR-0031 §決定 2）。
   * 値は `c:numCache`／`c:strCache` の保存値だけ（数式は計算しない）。明示が無い所は持たず、既定の色・範囲・配置は描画側が決める。
   */
  export interface ChartModel {
      type: 'bar' | 'line' | 'pie' | 'doughnut';
      /** 棒の向き（`c:barDir`。棒だけ持つ。縦棒 `col`・横棒 `bar`） */
      barDir?: 'col' | 'bar';
      /** 並べ方。棒は `clustered`・`stacked`・`percentStacked`、折れ線は `standard`・`stacked`・`percentStacked`、円・ドーナツは `standard` */
      grouping: 'clustered' | 'stacked' | 'percentStacked' | 'standard';
      /** `c:varyColors`（要素が無ければ true。Excel 16 の COM で確認・ADR-0031 C3。円・ドーナツで「常に点ごと」にするかは描画側） */
      varyColors: boolean;
      /** 項目の文字（数値・日付は書式を当てた後。項目の無いグラフは `1, 2, 3…`） */
      categories: readonly string[];
      /** 系列（`c:order` の順。保存値の無い系列は除く） */
      series: readonly ChartSeries[];
      /** 項目軸（`c:catAx`・`c:dateAx`）と数値軸（`c:valAx`）。円・ドーナツは持たない */
      catAxis?: ChartAxis;
      valAxis?: ChartAxis;
      title?: ChartTitle;
      /** `c:autoTitleDeleted`（自動のタイトルを出さない） */
      autoTitleDeleted: boolean;
      legend?: ChartLegend;
      /** プロットの `c:layout` の `c:manualLayout` */
      plotLayout?: ManualLayout;
      /** グラフの面（`c:chartSpace` の `c:spPr`）・プロットの面（`c:plotArea` の `c:spPr`） */
      chartArea: ChartBox;
      plotArea: ChartBox;
      /** `c:chartSpace` の `c:txPr`（既定の文字） */
      text: ChartTextStyle;
      /** `c:style`（`mc:AlternateContent` の `mc:Fallback` の中も見る。1〜48） */
      style?: number;
      /**
       * `c14:style`（`mc:AlternateContent` の `mc:Choice` の中の `c14:style` の `val`。整数）。要素が無い・読めなければ持たない。
       * 描画は値でなく「持つか」だけを見る（持つ＝Excel が COM の旧 API などで作った `c:spPr` 無しのグラフ。目盛線・軸線 #898989・タイトル 18pt。ISSUE-109・ADR-0031 §決定 4）。
       * `style` とは別の項目で、`style` の読み方（`mc:Fallback`）は変えない
       */
      c14Style?: number;
      /** `c:roundedCorners` の `val`（`val` が無ければ真・未知の値も真）。要素が無ければ持たない（描画側の既定は角丸＝真。ISSUE-109・ADR-0031 §決定 3） */
      roundedCorners?: boolean;
      /** 棒：``c:gapWidth`（0〜500）・`c:overlap`（-100〜100） */
      gapWidth?: number;
      overlap?: number;
      /** 円・ドーナツ：`c:firstSliceAng`（度。0〜360）・`c:holeSize`（%。ドーナツだけ） */
      firstSliceAng?: number;
      holeSize?: number;
      /** 折れ線のグループの `c:marker`（真ならマーカーあり。系列の `marker.symbol` が優先）。折れ線だけ持つ */
      lineMarkers?: boolean;
      /** 空の点の扱い（`c:dispBlanksAs`。要素が無ければ `gap`） */
      dispBlanksAs: 'gap' | 'zero' | 'span';
  }
  /** 開いた xlsx */
  export interface XlsxBook {
      readonly sheets: readonly SheetInfo[];
      /** シートを 1 枚開く（index は SheetInfo.index）。シートの XML と行の位置の索引だけを持つ */
      openSheet(index: number): Promise<XlsxSheet>;
      /**
       * シートの図（画像・グラフ・図形）の一覧を返す（index は SheetInfo.index。表示側が求めた時に読む）。
       * 上限超え・解析失敗でも例外にせず、`notices` で知らせる（表は openSheet で必ず出せる）。外部リンクは辿らない。
       * @throws RangeError index が範囲外／ XlsxError(broken) 閉じた後
       */
      openDrawings(index: number): Promise<SheetDrawings>;
      /**
       * 画像（`openDrawings(index)` が返した mediaPath）を読み、先頭バイトで形式を判定してヘッダの縦横を返す。
       * 出せないときは例外にせず `ok: false`（理由つき）を返す。1 シートの画像の合計の上限は、指定した `index` のシートの分で数える
       * （複数のシートが同じ画像を共有しても、数え先は呼び出し側が決める）。
       * @throws RangeError index が範囲外／ XlsxError(broken) 閉じた後
       */
      readImage(index: number, path: string): Promise<ImageResult>;
      /**
       * グラフ（`openDrawings(index)` が返した chartPath）を読み、G1 のグラフのモデルを返す（ADR-0031 §決定 2）。
       * 描けないとき（G1 外・壊れている・上限超え・描ける系列が無い）は例外にせず `ok: false`（理由つき）。
       * 1 シートのグラフの点の合計の上限は、指定した `index` のシートの分で呼んだ順に数える（`readImage` と同じ）。
       * @throws RangeError index が範囲外／ XlsxError(broken) 閉じた後
       */
      readChart(index: number, path: string): Promise<ChartResult>;
      /** 展開済みのデータを放す。以後 openSheet は例外、開いたシートの getRows は空 */
      close(): void;
  }
}

declare module 'xsheet/_/core/color' {
  import { parseHex6, sysClrValue, type Theme } from 'xsheet/_/core/theme';
  export { parseHex6, sysClrValue };
  /** 色：RGB の整数（0〜0xFFFFFF）と不透明度（0〜1） */
  export interface Rgba {
      rgb: number;
      alpha: number;
  }
  /** 見つけた色の要素 */
  export interface ColorElement {
      /** 要素名（接頭辞を除く。例 `srgbClr`） */
      name: string;
      /** 開きタグ全体（`parseAttrs` に渡せる） */
      open: string;
      /** 中身（子の変換）。自己閉じ・閉じタグの無いときは null */
      body: string | null;
  }
  /**
   * sRGB の成分（0〜1）と線形 RGB を行き来する式の指定。null は sRGB の区分式（IEC 61966-2-1）。
   * 数値なら単純なべき乗（γ 2.2・2.3 など）。`tint`・`shade`・`scrgbClr` の変換が使う（U7）。
   */
  export const LINEAR_GAMMA: number | null;
  /** sRGB の成分（0〜1）→ 線形 RGB（0〜1）。範囲外は収め、NaN は 0 */
  export function toLinear(c: number, gamma?: number | null): number;
  /** 線形 RGB（0〜1）→ sRGB の成分（0〜1）。範囲外は収め、NaN は 0 */
  export function fromLinear(l: number, gamma?: number | null): number;
  /**
   * path の `fill` の濃淡：`shade`（黒へ寄せる）か `tint`（白へ寄せる）の係数（0〜1）。
   * Excel の画素から決めた値：darken＝0.6・lighten＝0.6 は成分ごとの比がちょうど 0.6、darkenLess・lightenLess は 0.802〜0.809 に揃うので 0.805。
   * 濃淡は `applyFillMode` が **sRGB の値のまま**掛ける（線形 RGB にすると Excel の面より暗く／明るくなりすぎる）。
   */
  export const FILL_MODE_FACTORS: ReadonlyMap<string, {
      kind: 'shade' | 'tint';
      factor: number;
  }>;
  /** プリセット色の数（別名を含む。試験用） */
  export const PRESET_COLOR_COUNT: number;
  /** `prstClr` の名前から RGB の整数を引く。未知の名前は null */
  export function lookupPresetColor(name: string): number | null;
  /**
   * `xml`（例 `<a:solidFill>` の中身）で最初に見つかる色の要素を切り出す。無ければ null。
   * 閉じタグは `indexOf` で探す（閉じが無ければ body は null＝子の変換を読まない）。
   */
  export function findColorElement(xml: string): ColorElement | null;
  /**
   * 基本の色に、子の変換を出現順に掛ける（ADR-0028 §決定 6.3）。
   * 読む子は最初の `MAX_COLOR_TRANSFORMS` 個まで（無視する要素・不正な値も 1 個と数える）。
   * 各段の結果は 0〜1 に収める。値が整数でない変換は飛ばす。無視する変換（`comp`・`inv`・`gray`・`gamma`・`invGamma`・`red`/`green`/`blue` の系統）は色を変えない。
   */
  export function applyTransforms(base: Rgba, body: string | null): Rgba;
  /**
   * 切り出した色の要素 1 つを解く（基本の色 → 子の変換）。解けなければ null（例外にしない）。
   * 解けないときに次の出どころへ進むのは呼び出し側（ISSUE-95）。
   */
  export function resolveColorElement(el: ColorElement, theme: Theme): Rgba | null;
  /**
   * `xml`（`<a:solidFill>` の中身など）の最初の色の要素を解く。色の要素が無い・解けなければ null。
   * 色は RGB の整数と不透明度で返す（ファイルの文字列は返さない）。
   */
  export function resolveColor(xml: string, theme: Theme): Rgba | null;
  /**
   * path の `fill`（`norm`・`darken`・`darkenLess`・`lighten`・`lightenLess`）を色に掛ける。
   * 係数は `FILL_MODE_FACTORS`（U8）。未知の名前・`norm` は色のまま。不透明度は変えない。
   * 掛けるのは sRGB の値のまま（shade＝`c × k`・tint＝`1 − (1 − c) × k`）。変換の `shade`・`tint`（線形 RGB）とは別の式で、Excel の面の画素と合う方を採った。
   */
  export function applyFillMode(color: Rgba, mode: string): Rgba;
  /**
   * 2 色を **線形 RGB** で割合を混ぜる（`t`＝0 で a、1 で b。0〜1 に収め、NaN は 0）。不透明度は sRGB と同じ線形補間。
   * Excel の 2 stop のグラデーションの中央の色がこの混色と合う（ISSUE-95・U7。3 stop 以上は sRGB の混色 `mixColors` と合う）。
   */
  export function mixColorsLinear(a: Rgba, b: Rgba, t: number): Rgba;
  /**
   * 2 色を割合で混ぜる（`t`＝0 で a、1 で b。0〜1 に収め、NaN は 0）。各成分は sRGB の値の線形補間で、不透明度も混ぜる。
   * グラデーションの補間・模様の混色に使う（適用は ISSUE-95）。
   */
  export function mixColors(a: Rgba, b: Rgba, t: number): Rgba;
}

declare module 'xsheet/_/core/theme' {
  /**
   * ブックのテーマの 12 色（`xl/theme/theme1.xml` の `a:clrScheme`）の読み取り。
   * テーマの書式（`a:fmtScheme`）・フォント（`a:fontScheme`）は読まない。
   * テーマが無い・読めない・上限を超える・色が欠けているときは、Office 2013〜2022 の既定の 12 色で補う
   * （`XlsxError` にもバナーにもしない）。テーマはブックに 1 つで、styles.xml と同じくブックで 1 回だけ読む（index.ts）。
   */
  import type { XlsxLimits } from 'xsheet/_/core/limits';
  import type { XlsxPackage } from 'xsheet/_/core/open';
  import { type Attrs } from 'xsheet/_/core/xml';
  /** テーマの色の位置（`a:clrScheme` の子の要素名）。この順で 12 色 */
  export const THEME_SLOTS: readonly ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
  export type ThemeSlot = (typeof THEME_SLOTS)[number];
  /** テーマの 12 色（位置 → RGB の整数）。`Map` なので位置の名前が `__proto__` などでも汚染されない。常に 12 色そろう */
  export type Theme = ReadonlyMap<ThemeSlot, number>;
  /** 既定のテーマ（共有する。変更しない） */
  export const DEFAULT_THEME: Theme;
  /** 16 進 6 桁（`[0-9A-Fa-f]{6}`）だけを整数にする。それ以外・未指定は null */
  export function parseHex6(s: string | undefined): number | null;
  /** `sysClr` の色：`lastClr`（16 進 6 桁）。無い・不正なら `windowText`＝黒・`window`＝白、その他は null */
  export function sysClrValue(attrs: Attrs): number | null;
  /**
   * `theme1.xml` の `a:clrScheme` から 12 色を読む。欠けた色・不正な色はその色だけ既定で補う。
   * xml が null・`clrScheme` が無い・閉じが無いときは既定の 12 色（新しい `Map`）。
   */
  export function parseTheme(xml: string | null): Theme;
  /**
   * テーマを読む。`path` が無い・部品が zip に無い・上限（小さな部品）を超える・読めないときは既定の 12 色。
   * 画面を閉じた後の読み取り（`broken`・`closed`）だけは呼び出し側へ投げる（index.ts が閉じたかを確かめる）。
   */
  export function loadTheme(pkg: XlsxPackage, path: string | null, limits: XlsxLimits): Promise<Theme>;
}

declare module 'xsheet/_/core/open' {
  import type { XlsxLimits } from 'xsheet/_/core/limits';
  /** 開いた zip（部品の読み出し口） */
  export interface XlsxPackage {
      /** zip の中のファイルのパス（先頭に / を付けない） */
      has(path: string): boolean;
      /** 部品を文字（UTF-8）で読む。無ければ null。展開後の大きさが maxBytes を超えたら too-large */
      readText(path: string, maxBytes: number): Promise<string | null>;
      /** 部品をバイト列で読む（画像など）。無ければ null。展開しながら数え、maxBytes を超えたらそこで止めて too-large */
      readBytes(path: string, maxBytes: number): Promise<Uint8Array | null>;
      /** 読み込んだ zip を放す */
      dispose(): void;
  }
  /** Uint8Array へそろえる */
  export function toBytes(data: ArrayBuffer | Uint8Array): Uint8Array;
  /**
   * 先頭の印を確かめる。
   * - `PK\x03\x04` 以外は zip でない。
   * - OLE（`D0 CF 11 E0`）は、パスワード付きの xlsx（暗号化）か古い形式。どちらにしても読めないので encrypted。
   */
  export function checkSignature(bytes: Uint8Array): void;
  /** zip を開く。壊れていれば broken、ファイル数が多すぎれば too-large */
  export function openPackage(bytes: Uint8Array, limits: XlsxLimits): Promise<XlsxPackage>;
}

declare module 'xsheet/_/core/xml' {
  /** 要素名の前に付いてよい名前空間の接頭辞（正規表現の断片。例 `x:`） */
  export const NS: string;
  /** XML の実体参照（`&amp;` `&#x41;` `&#65;`）を戻す。属性値と、`_xHHHH_` を使わない値に使う */
  export function decodeEntities(s: string): string;
  /**
   * 文字列セルの本文を戻す。実体参照のあと、Excel の `_xHHHH_`（制御文字などの退避）を戻す。
   * `_x005F_` は `_` を表す（本物の `_x0041_` という文字列を守るため）。
   */
  export function decodeText(s: string): string;
  /** `findElement` の結果 */
  export interface ElementSpan {
      /** 開きタグ全体（属性を含む。`parseAttrs` に渡せる） */
      open: string;
      /** 中身。自己閉じ（`<cols/>`）・閉じタグの無いときは null */
      body: string | null;
      /** 要素の終わりの次の位置（次の検索の開始位置）。閉じタグの無いときは xml の長さ（検索を打ち切る） */
      end: number;
  }
  /**
   * `from` 以降で最初の `<name ...>…</name>`（名前空間の接頭辞可）を、線形時間で切り出す。
   * 開きタグを探し、閉じタグは `indexOf` で探す（`[\s\S]*?` の正規表現は、閉じが無いと入れ子の走査で二乗になる）。
   */
  export function findElement(xml: string, name: string, from?: number): ElementSpan | null;
  /**
   * `<name …>…</name>`（名前空間の接頭辞可）を出現順に callback へ渡す。同じ名前の入れ子は扱わない（`c`・`si`・`t` など入れ子にならない要素用）。
   * 開きタグは `[^<>]*>`、閉じは `indexOf` で探すので、`>` や閉じの無い要素を大量に並べても時間は線形。
   * 自己閉じ（`<c/>`）は body を null で渡す。閉じタグの無い要素に当たったら、以後に閉じが無いので打ち切る。
   * 正規表現は名前ごとに 1 度だけ作る（セル 1 つごとに作らない）。callback の中で同じ `xml` を再び走査してよい。
   */
  export function forEachElement(xml: string, name: string, cb: (open: string, body: string | null) => void): void;
  /**
   * `from` 以降で最初の `<name …>` の開きタグだけを返す（閉じは探さない。名前空間の接頭辞可。無ければ null）。
   * 開きタグは `[^<>]*>` で打ち切るので、`>` の無いタグを大量に並べても時間は線形。正規表現は名前ごとに 1 度だけ作る。
   */
  export function findOpenTag(xml: string, name: string, from?: number): string | null;
  /** 属性の一覧。キーに `__proto__` などが来てもプロトタイプを汚さないよう、プロトタイプの無いオブジェクトで返す */
  export type Attrs = Record<string, string | undefined>;
  /**
   * タグ（`<c r="A1" s="2">` の属性の部分）から属性を取り出す。値は実体参照を戻す。
   * 名前の先頭は後読みで「前が名前の文字でない位置」に限る（長い名前の途中から照合し直して二乗時間になるのを防ぐ。線形時間）。
   */
  export function parseAttrs(tag: string): Attrs;
  /** 名前空間の接頭辞を除いた属性の値（`r:id` と `rel:id` を同じに扱う）。`id` を探すとき `sheetId` は拾わない */
  export function attrByLocalName(attrs: Attrs, local: string): string | undefined;
  /** 列番号（1 始まり）を列の記号にする。1→A、27→AA。1 未満・整数でなければ空文字 */
  export function colName(c: number): string;
  /** 番地 `AB12` の位置（1 始まり）。範囲外・不正なら null */
  export function parseCellRef(ref: string): {
      r: number;
      c: number;
  } | null;
  /** 列だけの番地（`AB`）。セルの r 属性（`AB12`）から列だけ欲しいとき用。範囲外・不正なら 0 */
  export function colOfRef(ref: string): number;
  /** 開きタグの名前の直後の文字として許す文字（`<c` の後が `o`（`<cols`）などでないことの確認用） */
  export function isTagDelimiter(ch: string | undefined): boolean;
  /** Strict 形式（ISO 29500 Strict）の名前空間の目印。1 次では読まない（unsupported） */
  export const STRICT_NS = "purl.oclc.org/ooxml";
  /** ルート要素の宣言（先頭の 2,000 文字）に Strict の名前空間があるか */
  export function looksStrict(xml: string): boolean;
}

declare module 'xsheet/_/core/geometry' {
  import type { SheetLayout } from 'xsheet/_/core/types';
  /** 行高 pt → px。`round(pt × 96 / 72)`（15pt＝20px、18.75pt＝25px） */
  export function ptToPx(pt: number): number;
  /**
   * 列幅（`<col width>` の文字数・余白を含む値）→ px。
   * ECMA-376 Part 1 の `col` 要素の換算式 `trunc(((256·w + trunc(128/MDW)) / 256) · MDW)`。
   * @param w 列幅（文字数）
   * @param mdw 最大桁幅 MDW（px）
   */
  export function colWidthToPx(w: number, mdw: number): number;
  /**
   * 既定の列幅（`defaultColWidth` が無いとき）→ px。`ceil((baseColWidth·MDW + 余白) / 8) · 8`
   * （MDW 7 で 64px、MDW 8 で 72px）。余白は `2·ceil(MDW/4) + 1`（MDW 5〜8 では 5px）。
   * 根拠: Excel 16 の実ファイルの `<col width>` の端数から余白を逆算（2026-10-27）。
   */
  export function defaultColWidthToPx(baseColWidth: number, mdw: number): number;
  /** 換算した座標（px・すべて整数）。番号は 1 始まり */
  export interface Geometry {
      /** 列 c の左端の x（c ≤ 1 は 0）。cols より右の列も既定の幅で延ばして返す */
      colLeft(c: number): number;
      /** 列 c の幅（非表示は 0。c < 1 は 0） */
      colWidth(c: number): number;
      /** 行 r の上端の y（r ≤ 1 は 0）。rows より下の行も既定の高さで延ばして返す */
      rowTop(r: number): number;
      /** 行 r の高さ（非表示は 0。r < 1 は 0） */
      rowHeight(r: number): number;
      /**
       * y を含む行（1〜rows）。負・NaN は先頭、下端以降は最後の行へ丸める。rows が 0 なら 0。
       * 非表示の行（高さ 0）は、同じ y の次の表示行より前になるため返らない（末尾だけが非表示のときを除く）。
       */
      rowAt(y: number): number;
      /** x を含む列（1〜cols）。丸め方は `rowAt` と同じ */
      colAt(x: number): number;
      /** 表の幅（1〜cols の列幅の合計） */
      readonly totalWidth: number;
      /** 表の高さ（1〜rows の行高の合計） */
      readonly totalHeight: number;
  }
  /**
   * layout から座標（px）の計算器を作る。
   * - 列幅は `<col width>` を `colWidthToPx` で換算し、指定の無い列は既定の列幅（`defaultColWidth` があれば同じ式で換算、
   *   無ければ `defaultColWidthToPx`）。非表示の列は 0。
   * - 行高：シートの既定の行高は、`customHeight="1"` なら `defaultRowHeight`（pt）の換算値、
   *   そうでなければフォントの既定の行高（`fontRowPx`。実測表）を優先し、無ければ `defaultRowHeight` の換算値。
   *   （Excel は `customHeight` の無い `defaultRowHeight`・`<row ht>` を、その PC のフォントから計算した行高で置き換えるため）
   *   行は、`customHeight="1"` の `ht` は `ptToPx(ht)`。`customHeight` の無い `ht` は、ファイルの
   *   `defaultRowHeight` と ±0.01 で等しければシートの既定、違えば保存時の自動調整値として `ptToPx(ht)` とする。
   *   `ht` が無ければシートの既定。非表示の行は 0。
   *   `zeroHeight` のときは、`<row>` の無い行を既定で高さ 0 とする（`<row>` のある行は表示）。
   * - 列幅・行高は Excel の上限（255 文字・409.5pt）に丸める。
   * @param layout `XlsxSheet.layout`
   * @param rows 表の行数（`XlsxSheet.rowCount`）。`totalHeight`・`rowAt` の範囲に使う
   * @param cols 表の列数（`XlsxSheet.colCount`）。`totalWidth`・`colAt` の範囲に使う
   */
  export function createGeometry(layout: SheetLayout, rows: number, cols: number): Geometry;
}

declare module 'xsheet/_/core/font-metrics' {
  /**
   * フォントごとの寸法の実測表：最大桁幅 MDW（px）と、既定の行高（px）。
   * 出所: 2026-10-27・Windows GDI（TextRenderer）96dpi で数字 0〜9 の最大幅を実測。
   * Calibri 11＝7・游ゴシック 11＝8・맑은 고딕 11＝8・MS Pゴシック 11＝8 は、既知の Excel の値と一致する。
   * 既定の行高は Excel 16 の COM で確認した値（96dpi 換算）。
   */
  /**
   * フォント名を表の名前にそろえる。NFKC 正規化（全角英字を半角に）・小文字・空白除去のあと、現地語名の別名を引く。
   * 表に無い名前は正規化しただけの名前を返す。
   */
  export function canonicalFontName(name: string): string;
  /** サイズの上限（pt）。Excel のフォントサイズの上限 */
  export const MAX_FONT_SIZE = 409;
  /**
   * フォントの名前・サイズ（pt）→ 最大桁幅 MDW（px・96dpi）。
   * 実測表のサイズの間は、近い 2 点から線形補間して四捨五入する。表の範囲外は、端の点に比例させる。
   * 表に無いフォントは Calibri の行を使う。サイズが不正（数でない・0 以下）なら null。
   */
  export function lookupMdw(name: string, size: number): number | null;
  /**
   * フォントの既定の行高（px・96dpi）。実測した（名前・サイズ）の組だけを持ち、無ければ null。
   * Calibri 11＝20、游ゴシック 11＝25、맑은 고딕 11＝22、MS Pゴシック 11＝18、Arial 10＝17。
   */
  export function fontDefaultRowPx(name: string, size: number): number | null;
}

declare module 'xsheet/_/core/cell-format' {
  /**
   * styles.xml の書式の読み取り：`<fills>`（塗り）・`<fonts>`（文字の装飾）と、cellXfs の `fontId`・`fillId`。
   * ブックで 1 回だけ読み、セルの `s` 属性（cellXfs の位置）→ 書式の番号 → 塗りと文字の装飾、と引けるようにする（`FormatTable`）。
   * - 見える書式（塗り・太字・斜体・下線・取り消し線・自動でない色・比が 1 でない文字サイズのどれか）が無い xf は「書式なし」にまとめる。
   * - 同じ見た目の書式は 1 つにまとめる（キーは数値だけ。ファイルの文字列は持たない）。まとめた後の数は `maxCellFormats` まで。
   * - 塗りは 1 色に畳む：`solid`＝`fgColor`、ほかの模様＝前景と背景の混色、グラデーション＝50% の位置の色。
   * - `applyFill`・`applyFont` は見ない（F5）。cellStyleXfs は読まない。フォント名は読まない。
   * - 色は SpreadsheetML の `<color>`（cell-color.ts）をテーマで解いてから持つ。解けない値は例外にせず既定の色・書式なしに倒す。
   * - 走査は線形（開きタグは `[^<>]*>`、閉じは `indexOf`）。
   */
  import { type Rgba } from 'xsheet/_/core/color';
  import { type XlsxLimits } from 'xsheet/_/core/limits';
  import type { Theme } from 'xsheet/_/core/theme';
  import type { CellFormat } from 'xsheet/_/core/types';
  /** 書式の表 */
  export interface FormatTable {
      /** 書式の番号 → 書式（見える書式だけ・重複なし。`SheetRow.formats` の値・`XlsxSheet.formats` の位置） */
      readonly formats: readonly CellFormat[];
      /** cellXfs の位置（セルの `s`）→ 書式の番号。書式なし・不正・上限の外は -1 */
      readonly xfFormat: Int32Array;
  }
  /** 書式の無い表（styles.xml が無い・壊れているとき。共有する・変更しない） */
  export const EMPTY_FORMAT_TABLE: FormatTable;
  /**
   * styles.xml から書式の表を作る（ブックで 1 回）。
   * @param xml styles.xml（無ければ null＝書式なし）
   * @param theme ブックのテーマ（色の `theme` を解く）
   * @param limits `maxFonts`・`maxFills` を超えた要素は読まない（それを参照するセルは装飾なし・塗りなし）。
   *   重複を除いた書式の数が `maxCellFormats` を超えた書式のセルは書式なし
   */
  export function parseCellFormats(xml: string | null, theme: Theme, limits: XlsxLimits): FormatTable;
  /**
   * 「紙」の塗りか：解決後の塗りが不透明な `#FFFFFF`（`rgb`・`theme`・`indexed` など出どころは問わない）。
   * 紙は塗りなしと同じに描く（ADR-0030 1.1）ため、表の範囲を広げる判定（sheet.ts）と描き分け（viewer）の両方がこれを使う。
   * `#FEFEFE`・`#F2F2F2`・半透明の白は紙ではない（alpha < 1）。
   */
  export function isPaperFill(fill: Rgba | undefined): boolean;
  /**
   * 書式の色に `theme` の指定があるか（`<fonts>`・`<fills>`・`<borders>` の中だけを見る）。無ければテーマを読まなくても
   * 色は決まるので、呼び出し側（index.ts）はテーマの読み取りを省ける（図の無いブックでテーマの読み取りを増やさない。ISSUE-94）。
   */
  export function stylesUseTheme(xml: string | null): boolean;
}

declare module 'xsheet/_/render/index' {
  /**
   * render の入口（`xsheet/render`）。DOM の表・罫線・図の層。
   * InterHub `viewers/xlsx-*.ts`・`spreadsheet.ts` の .xlsx 部分から作る。`renderSheet` は M2-3、`renderWorkbook` は M2-4。
   */
  export { MESSAGES_EN, MESSAGES_JA, MESSAGES_KO, type XSheetMessages } from 'xsheet/_/render/messages';
  export { renderSheet, type RenderOptions, type SheetView } from 'xsheet/_/render/render-sheet';
  export { renderWorkbook, type WorkbookOptions, type WorkbookView } from 'xsheet/_/render/render-workbook';
}

declare module 'xsheet/_/render/messages' {
  /**
   * render の文言（`RenderOptions.messages`）。InterHub の `t('spreadsheet.…')` の代わり。
   * 既定は日本語。英語・韓国語も同梱し、使う側は一部のキーだけを渡して差し替えられる。
   * 差し込み（`{detail}`・`{format}`）は render が行う。文言は `textContent`・`title` にだけ入れる。
   */
  /** 文言の表（23 キー。元は InterHub `locales/ja.ts` の `spreadsheet.*` と `data.noData`） */
  export interface XSheetMessages {
      /** 図形の形の表（geom）を読み込めない（図の通知。形のある図形は代替で描く）。InterHub では xlsx の読み込み失敗の文言だった */
      libLoadFailed: string;
      /** ファイルを解析できない（`{detail}` は例外の文言。セル値・シート名は入らない） */
      parseFailed: string;
      /** シートが 1 枚も無い */
      noSheets: string;
      /** シートにセルも図も無い */
      noData: string;
      encrypted: string;
      tooLarge: string;
      unsupported: string;
      /** 非表示のシートのタブの補足（`renderWorkbook`） */
      hiddenSheet: string;
      /** タブを前・次へ送るボタン（`renderWorkbook`） */
      tabsPrev: string;
      tabsNext: string;
      /** 図の破線枠の種類名 */
      chart: string;
      smartart: string;
      shape: string;
      textbox: string;
      group: string;
      /** 描けない形式の画像（`{format}` は EMF などの形式名） */
      imageFormat: string;
      /** 表示できない画像 */
      imageBroken: string;
      /** 図の通知（`SheetDrawings.notices` と画像の予算超え） */
      noticeTooLarge: string;
      noticeBroken: string;
      noticeTooMany: string;
      noticeTextTruncated: string;
      noticeShapeSimplified: string;
      noticeOverBudget: string;
  }
  /** 日本語（既定） */
  export const MESSAGES_JA: Readonly<XSheetMessages>;
  /** 英語 */
  export const MESSAGES_EN: Readonly<XSheetMessages>;
  /** 韓国語（M6-2・ADR-0003 Q4） */
  export const MESSAGES_KO: Readonly<XSheetMessages>;
  /**
   * 使う文言の表を作る：既定（日本語）に、渡された分だけを上書きする。
   * 文字列でない値・表に無いキーは無視する（使う側の誤りで表示が壊れないように）。
   */
  export function resolveMessages(messages?: Partial<XSheetMessages>, base?: Readonly<XSheetMessages>): Readonly<XSheetMessages>;
  /** `{name}` を差し込む。値の無い `{name}` はそのまま残す（InterHub の `t()` と同じ） */
  export function formatMessage(template: string, params?: Readonly<Record<string, string | number>>): string;
}

declare module 'xsheet/_/render/render-sheet' {
  /**
   * 公開の入口 `renderSheet`（M2-3）：開いたブックの 1 シートを `container` の中に表示する。
   * 表示は呼び出しごとのインスタンスで、`destroy()` でリスナ・blob URL・シートを放す（ブックは閉じない）。
   * 中身は `sheet-view.ts`（InterHub `spreadsheet.ts` の .xlsx の部分）。タブ付きのブック全体は `renderWorkbook`（M2-4）。
   */
  import type { XlsxBook } from 'xsheet/_/core/index';
  import { type XSheetMessages } from 'xsheet/_/render/messages';
  /** 表示の指定 */
  export interface RenderOptions {
      /** 文言（一部のキーだけでよい。既定は日本語の `MESSAGES_JA`） */
      messages?: Partial<XSheetMessages>;
      /** 図の層（画像・図形・線）を出すか（既定 true）。false なら図を読まず、図の通知も出さない */
      drawings?: boolean;
      /** 書式の `<style>` に付ける nonce（CSP の `style-src 'nonce-…'` 用） */
      nonce?: string;
  }
  /** 表示 1 つ */
  export interface SheetView {
      /** リスナ・ResizeObserver・blob URL・シートを放し、表示を `container` から外す。何度呼んでもよい。ブックは閉じない */
      destroy(): void;
  }
  /**
   * ブックの 1 シートを `container` の中に表示する。
   * `container` は高さが決まっていること（中のスクロール枠 `.xs-body` が仮想化の窓になる）。
   * シートを開けない（壊れている等）ときは、例外にせず枠の中にメッセージを出す。
   *
   * @param container 表示を置く要素（子として `div.xs-sheet` を足す。`container` 自身には何も付けない）
   * @param book `openXlsx` で開いたブック（閉じるのは呼び出し側）
   * @param sheetIndex シートの番号（`SheetInfo.index`。1 始まり・非表示のシートも数える）
   * @throws RangeError ブックに無い番号（DOM には何も足さない）
   */
  export function renderSheet(container: HTMLElement, book: XlsxBook, sheetIndex: number, options?: RenderOptions): Promise<SheetView>;
}

declare module 'xsheet/_/render/render-workbook' {
  /**
   * 公開の入口 `renderWorkbook`（M2-4）：xlsx のバイト列を開き、シートタブ付きで `container` の中に表示する。
   * InterHub `spreadsheet.ts` の `renderXlsxView` のうち、ブックを開く・タブ・最初のシートの選択・エラー表示の部分。
   * - 表示は 1 冊につき 1 つのシートの表示（`createSheetView`）を使い回し、タブでシートを替える
   *   （書式の `<style>` はブックで 1 回だけ作り、タブを替えても作り直さない）。
   * - 開けない（暗号化・壊れている・大きすぎる・対応しない形式）ときは例外にせず、種類ごとの文言を `div.xs-banner` に出す。
   * - `destroy()` でタブ列・シートの表示（blob URL・リスナ）を放し、ブックを閉じ、根を外す。
   * シート名はユーザーのファイル内容（信頼できない入力）なので、`textContent` と `title` にだけ入れる。
   */
  import { type SheetInfo, type XlsxLimits } from 'xsheet/_/core/index';
  import type { RenderOptions, SheetView } from 'xsheet/_/render/render-sheet';
  /** `renderWorkbook` の指定 */
  export interface WorkbookOptions extends RenderOptions {
      /** 読み取りの上限を替える（`openXlsx` の `limits`。既定は `DEFAULT_LIMITS`） */
      limits?: Partial<XlsxLimits>;
  }
  /** タブ付きのブックの表示 */
  export interface WorkbookView extends SheetView {
      /**
       * ブックのシートの一覧（タブの並び。開けなかったブックでは空）。
       * XSheet: `view` の `Viewer.sheets`・名前でのシートの切り替えのために足した（M6-2）
       */
      readonly sheets: readonly SheetInfo[];
      /**
       * シートを替える（タブを押したのと同じ）。表示中のシートなら何もしない。
       * @param index シートの番号（`SheetInfo.index`。1 始まり・非表示のシートも数える）
       * @throws RangeError ブックに無い番号（開けなかったブック・シートの無いブックでは常に）
       */
      showSheet(index: number): Promise<void>;
  }
  /**
   * xlsx を開き、シートタブ付きで `container` の中に表示する。
   * 最初に出すシートは、最初の表示（非表示でない）シート。全部非表示なら先頭。
   * `container` は高さが決まっていること（タブ列の下のスクロール枠 `.xs-body` が仮想化の窓になる）。
   *
   * @param container 表示を置く要素（子として `div.xs-workbook` を足す。`container` 自身には何も付けない）
   * @param data xlsx のバイト列
   */
  export function renderWorkbook(container: HTMLElement, data: ArrayBuffer | Uint8Array, options?: WorkbookOptions): Promise<WorkbookView>;
}

declare module 'xsheet/_/facade/index' {
  /**
   * 入口の層（M6-2）。`view`（1 行で表示）・`read`（DOM なしで開く）。
   * core・render の上の薄い層で、InterHub から写したファイルとは別に置く（ADR-0003 §4-1 の 3）。
   */
  export type { XSheetSource } from 'xsheet/_/facade/source';
  export type { XSheetTheme } from 'xsheet/_/facade/style';
  export { read, view, type ReadOptions, type Viewer, type ViewOptions, type XSheetLang, } from 'xsheet/_/facade/view';
}

declare module 'xsheet/_/facade/source' {
  /**
   * `view`・`read` が受け取る xlsx の入れ物をバイト列にする（M6-2）。
   * URL の文字列は受けない（ライブラリは通信しない）。使う側が `fetch(url)` を渡せば、受け取った `Response` の中身を読むだけ。
   * 別の realm（iframe・jsdom）の物でも動くよう、`instanceof` ではなく形で見分ける。
   */
  /** `view`・`read` が受け取る xlsx の入れ物 */
  export type XSheetSource = Blob | ArrayBuffer | ArrayBufferView | Response | PromiseLike<Blob | ArrayBuffer | ArrayBufferView | Response>;
  /**
   * 入れ物を Uint8Array にする。
   * @throws TypeError 受けない形（URL の文字列など）・失敗した `Response`（`ok` が false）
   */
  export function toSourceBytes(source: XSheetSource): Promise<Uint8Array>;
}

declare module 'xsheet/_/facade/style' {
  /** テーマ。`auto` は `prefers-color-scheme` に従う */
  export type XSheetTheme = 'light' | 'dark' | 'auto';
  /** container の高さが 0 のときに付ける高さ */
  export const DEFAULT_HEIGHT = "480px";
  /** 表示の CSS が無ければ入れる（読み込み済み・入れ済みなら何もしない） */
  export function ensureStyle(container: HTMLElement, nonce?: string): void;
  /** テーマを表示の根に付ける（`light` は何も付けず、使う側の `--xs-*` の上書きがそのまま効く） */
  export function applyTheme(root: HTMLElement, theme: XSheetTheme | undefined): void;
  /**
   * 表示の根の高さを決める。
   * - 数値は px、文字列は CSS の長さとしてそのまま付ける。
   * - `auto`（既定）は container の高さが 0 のときだけ {@link DEFAULT_HEIGHT} を付ける（それ以外は style.css の 100%）。
   *   container が描かれていないときは、描かれたときに決める。
   * @returns 後始末（描かれるのを待つ ResizeObserver を外す）
   */
  export function applyHeight(container: HTMLElement, root: HTMLElement, height: number | string | undefined): () => void;
}

declare module 'xsheet/_/facade/view' {
  /**
   * 1 行で使える入口 `view`・`read`（M6-2）。core（`openXlsx`）と render（`renderWorkbook`）の上の薄い層で、
   * 写したファイル（`src/core/`・`src/render/`）の作りには立ち入らない。
   * - `view`：xlsx をタブ付きで出す。表示の根 `div.xs-viewer` を container の子に足し、その中に `renderWorkbook` の表示を置く。
   *   言語・テーマ・CSS・高さを語で選べる。開けないファイルは今までどおり例外にせず、帯（`div.xs-banner`）に文言を出す。
   * - `read`：DOM なしで開く（`openXlsx` と同じ。受け取れる形を `view` と揃える）。
   */
  import { type SheetInfo, type XlsxBook, type XlsxLimits } from 'xsheet/_/core/index';
  import { type XSheetMessages } from 'xsheet/_/render/index';
  import { type XSheetSource } from 'xsheet/_/facade/source';
  import { type XSheetTheme } from 'xsheet/_/facade/style';
  /** 文言の言語。`auto` は `<html lang>`、無ければブラウザの言語。合うものが無ければ英語 */
  export type XSheetLang = 'ja' | 'en' | 'ko' | 'auto';
  /** `read` の指定 */
  export interface ReadOptions {
      /** 読み取りの上限を替える（既定は `DEFAULT_LIMITS`） */
      limits?: Partial<XlsxLimits>;
  }
  /** `view` の指定（どれも省略できる） */
  export interface ViewOptions extends ReadOptions {
      /** 文言の言語（既定 `auto`） */
      lang?: XSheetLang;
      /** 文言を一部だけ替える（`lang` で選んだ表の上に重ねる） */
      messages?: Partial<XSheetMessages>;
      /** テーマ（既定 `light`）。`auto` は `prefers-color-scheme` に従う */
      theme?: XSheetTheme;
      /** 表示の CSS を自動で入れるか（既定 `auto`＝読み込み済みでなければ入れる。`none` は入れない） */
      css?: 'auto' | 'none';
      /** 表示の高さ（既定 `auto`＝container の高さが 0 のときだけ 480px）。数値は px、文字列は CSS の長さ */
      height?: number | string;
      /** 図の層（画像・図形・線）を出すか（既定 true） */
      drawings?: boolean;
      /** 自動で入れる `<style>`・書式の `<style>` に付ける nonce（CSP の `style-src 'nonce-…'` 用） */
      nonce?: string;
  }
  /** `view` の戻り値 */
  export interface Viewer {
      /** シートの一覧（タブの並び。非表示のシートも入る。開けなかったファイルでは空） */
      readonly sheets: readonly SheetInfo[];
      /**
       * シートを替える。
       * @param sheet シートの番号（`SheetInfo.index`・1 始まり）か名前
       * @throws RangeError ブックに無い番号・名前
       */
      showSheet(sheet: number | string): Promise<void>;
      /** 表示を外し、ブック・リスナ・blob URL を放す。何度呼んでもよい */
      destroy(): void;
  }
  /**
   * xlsx をタブ付きで表示する。
   * 同じ container に続けて呼ぶと、前の表示を外してから出す（読み込み中の前の呼び出しは、表示の無い Viewer を返す）。
   * @param target 表示を置く要素か CSS セレクタ（子として `div.xs-viewer` を足す）
   * @param source File・Blob・ArrayBuffer・型付き配列・Response、またはその Promise（`fetch(url)` をそのまま渡せる）
   * @throws TypeError 要素が見つからない・受けない形の source・失敗した Response
   */
  export function view(target: HTMLElement | string, source: XSheetSource, options?: ViewOptions): Promise<Viewer>;
  /**
   * xlsx を DOM なしで開く（`openXlsx` と同じ。閉じるのは呼び出し側の `close()`）。
   * @throws TypeError 受けない形の source・失敗した Response
   * @throws XlsxError 開けないファイル（`kind` で種類が分かる）
   */
  export function read(source: XSheetSource, options?: ReadOptions): Promise<XlsxBook>;
}

/** xsheet.min.js が作る大域の名前 */
declare const XSheet: typeof import('xsheet');

