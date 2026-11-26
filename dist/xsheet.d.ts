// XSheet 0.8.0 単独版（xsheet.min.js・xsheet.mjs）の型。tools/standalone/build.mjs が dist の d.ts から作る（手で直さない）。

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
  import type { BuiltinDates } from 'xsheet/_/core/format-ja';
  import { type XlsxLimits } from 'xsheet/_/core/limits';
  import { type XlsxBook } from 'xsheet/_/core/types';
  export { XlsxError } from 'xsheet/_/core/types';
  export type { AnchorPoint, ArrowSize, ArrowType, BorderEdge, BorderStyle, CellBorder, CellFormat, CellKind, CellRange, CellView, ChartAxis, ChartBox, ChartFailReason, ChartLegend, ChartLine, ChartMarker, ChartMarkerSymbol, ChartModel, ChartPaint, ChartResult, ChartSeries, ChartTextStyle, ChartTickMark, ChartTitle, ColLayout, CustGeomDef, CustGeomPath, DrawingAnchor, DrawingColor, DrawingGroup, DrawingKind, DrawingLine, DrawingNotice, DrawingText, DrawingTextParagraph, DrawingTextRun, GeomArg, GeomPathFill, GroupLevel, GroupRect, ImageFailReason, ImageFormat, ImageResult, LineDash, LineEnd, LineGeom, ManualLayout, RowLayout, RunFont, ShapeGeom, ShapeOutline, SheetDrawing, SheetDrawings, SheetInfo, SheetLayout, SheetRow, TextAlign, TextAnchor, TextRun, TextVert, XlsxBook, XlsxErrorKind, XlsxSheet, } from 'xsheet/_/core/types';
  export type { Rgba } from 'xsheet/_/core/color';
  export { createGeometry, colWidthToPx, defaultColWidthToPx, ptToPx } from 'xsheet/_/core/geometry';
  export { fontDefaultRowPx, lookupMdw } from 'xsheet/_/core/font-metrics';
  export type { Geometry } from 'xsheet/_/core/geometry';
  export type { XlsxLimits } from 'xsheet/_/core/limits';
  export { DEFAULT_LIMITS } from 'xsheet/_/core/limits';
  export { colName } from 'xsheet/_/core/xml';
  export { isPaperFill } from 'xsheet/_/core/cell-format';
  export { findText } from 'xsheet/_/core/find-text';
  export type { FindTextOptions, FindTextResult, TextMatch } from 'xsheet/_/core/find-text';
  export type { BuiltinDates } from 'xsheet/_/core/format-ja';
  /** openXlsx の追加の指定 */
  export interface OpenXlsxOptions {
      /** 上限を替える（既定は DEFAULT_LIMITS。試験で小さい値を渡す用） */
      limits?: Partial<XlsxLimits>;
      /**
       * 組み込みの日付の書式 27〜36・50〜58 の見せ方（既定 `'default'`＝西暦の表のまま）。
       * `'japanese-era'` で日本語版 Excel と同じ表（28 は `令和8年11月11日`、27 は `R8.11.11` など）にする。
       * 書式に直接書かれた和暦（`ggg`・`e` など）と曜日（`aaa`・`aaaa`）は、指定しなくてもいつも出す。
       * XSheet: ADR-0007 §6（M7-8）
       */
      builtinDates?: BuiltinDates;
  }
  /**
   * xlsx を開く（シートの中身は openSheet で開くときに初めて展開する）。
   * @throws XlsxError zip でない（not-zip）・暗号化（encrypted）・壊れている（broken）・大きすぎる（too-large）・読めない形式（unsupported）
   */
  export function openXlsx(data: ArrayBuffer | Uint8Array, options?: OpenXlsxOptions): Promise<XlsxBook>;
}

declare module 'xsheet/_/core/format-ja' {
  /**
   * 書式コードを小さな字句に分け、ssf が知らない日本語の字句（曜日・元号・和暦の年）を、計算した文字を引用符で包んだ形に置き換える。
   * 置き換えるのは、引用符の中・`\` の後・`_`／`*` の後・`[...]` の中・`AM/PM`・`A/P` 以外にある `a` の並びだけ。
   * 曜日は `Date` を使わず、日付の通し番号から整数の計算で出す（時間帯のずれが入らない）。
   * 和暦（`g`・`e`・`r`）は M7-8。正解は Excel 16（ja-JP）で作った `fixtures/excel/ref-weekday.xlsx`・`ref-era.xlsx`。
   */
  /** 字句の種類 */
  export type FormatTokenKind = 
  /** `"..."` の文字（text は引用符を含む元の形） */
  'quoted'
  /** `\x`・`_x`・`*x`（text は記号を含む 2 文字） */
   | 'escaped'
  /** `[...]`（色・条件・`$-411` など。text は角かっこを含む） */
   | 'bracket'
  /** `AM/PM`・`A/P`（大文字小文字を問わない） */
   | 'ampm'
  /** 曜日 `aaa`・`aaaa`（text は元の並び） */
   | 'weekday'
  /** 元号 `g`・`gg`・`ggg`（大文字も。text は元の並び） */
   | 'era'
  /** 和暦の年 `e`・`ee`（大文字も） */
   | 'eraYear'
  /** 和暦の `r`（＝`ee`）・`rr`（＝`gggee`）（大文字も） */
   | 'eraR'
  /** `General`（大文字小文字を問わない。中の e・G を字句にしない） */
   | 'general'
  /** 上のどれでもない文字（1 文字ずつ） */
   | 'char';
  export interface FormatToken {
      kind: FormatTokenKind;
      text: string;
  }
  /** `;` で分けた区画。tokens は区画の字句（`;` 自体は含まない） */
  export type FormatSection = readonly FormatToken[];
  /**
   * 書式コードを `;` の区画ごとの字句に分ける。
   * 閉じていない引用符・角かっこは末尾までを 1 つの字句にする（壊れた書式でも止まらない）。
   * 1 文字ずつ前へ進むだけなので、長さに比例した時間で終わる。
   */
  export function tokenizeFormat(code: string): FormatSection[];
  /**
   * Excel の日付の通し番号から曜日（0＝日曜〜6＝土曜）を出す。Excel の `WEEKDAY(値)-1` と同じ。
   * 1900 年基準：Excel は 1900-02-29（通し番号 60）を実在の日として数えるので、通し番号 1（1900-01-01）を日曜とする
   * 連続した数え方になる（61 以降は実際の暦と合う）。1904 年基準：通し番号 0 が 1904-01-01（金曜）。
   * 時刻の部分（小数）は切り捨てる。
   */
  export function weekdayOfSerial(serial: number, date1904: boolean): number;
  /**
   * 区画の和暦の扱い（ref-era.xlsx で確かめた）。
   * - `ja`：和暦で出す（言語の指定が日本語、または指定が無く元号 `g`・`rr` がある）
   * - `gannen`：`ja` に加え、2 文字以上の元号の後の年が 1 なら「元」（`[$-ja-JP-x-gannen]`）
   * - `other`：日本語でない言語の指定（`[$-409]` など）。元号は空、年（`e`）は ssf の西暦のまま
   * - `none`：和暦の字句が無い、または指定が無く `e` だけ（Excel は開く言語環境で変わるので西暦のまま。ADR-0007 §6）
   */
  type EraMode = 'ja' | 'gannen' | 'other' | 'none';
  /** 曜日・和暦の字句を含む書式の準備（区画ごとの字句と和暦の扱い）。含まなければ null */
  export interface JaFormat {
      sections: FormatSection[];
      /** 区画ごとの和暦の扱い（sections と同じ長さ） */
      eras: EraMode[];
  }
  /** 書式コードに曜日・和暦の字句があれば字句に分けたものを返す。無ければ null（そのまま ssf に渡せる） */
  export function prepareJaFormat(code: string): JaFormat | null;
  /**
   * 曜日の字句を、値から計算した曜日を引用符で包んだ形に置き換えた書式コードを返す。
   * 区画の選び方は ssf に任せるため、全部の区画を同じ曜日で置き換える。負の値の区画（第 2 区画）は ssf が
   * 絶対値で日付を出すので、曜日も絶対値から出す。値が有限でなければ何もしない。
   * 日付の範囲の外（Excel が #### にする値）で、使う区画に曜日があれば、セル全体を空にする書式 `""` を返す
   * （曜日の無い日付の書式も ssf が空にするのにそろえる）。
   */
  export function applyJaFormat(prepared: JaFormat, value: number, date1904: boolean): string;
  /** 通し番号の元号と和暦の年 */
  export interface EraOfSerial {
      /** 元号の名前（`令和`） */
      name: string;
      /** 略称（`令`） */
      short: string;
      /** 記号（`R`） */
      letter: string;
      /** 和暦の年（1 始まり） */
      year: number;
  }
  /**
   * Excel の日付の通し番号（0 以上。時刻の部分は切り捨て）から元号と和暦の年を出す。
   * 1900 年基準の 0〜60（1900-01-00〜Excel の 1900-02-29）は 1900 年、61 以降は 1899-12-30 からの日数で年を出す
   * （UTC で数えるので時間帯に依らない）。1904 年基準は 1462 を足して 1900 年基準にする。
   * Excel の範囲（1900 年〜）では明治 33 年から。
   */
  export function eraOfSerial(serial: number, date1904: boolean): EraOfSerial;
  /** ssf の使う部分（`format.ts` の SsfApi と同じ形） */
  interface SsfLike {
      format(fmt: string | number, val: unknown, opts?: object): string;
      is_date(fmt: string): boolean;
  }
  /**
   * ssf を包み、曜日・和暦の字句を含む書式コードだけ、ssf に渡す前に置き換える。
   * `is_date` は置き換える字句があれば真にする（`aaa`・`[$-411]g` だけの書式は ssf では日付にならないため）。
   * 字句の分解は書式コードごとに 1 回だけ（上限付きで覚える）。
   */
  export function withJaTokens<T extends SsfLike>(inner: T): SsfLike;
  /**
   * 組み込みの書式 27〜36・50〜58 の、日本語版 Excel での書式コード
   * （ref-era.xlsx の builtin のシートで、Excel 16（ja-JP）の NumberFormat を読んで確かめた）。
   * `builtinDates: 'japanese-era'` のときだけ使う（既定は `format.ts` の表のまま。ADR-0007 §6）。
   */
  export const JA_BUILTIN_DATES: ReadonlyMap<number, string>;
  /**
   * 組み込みの日付の書式 27〜36・50〜58 の見せ方。
   * `'japanese-era'`＝日本語版 Excel と同じ（和暦を含む）。`'default'`（既定）＝今までの西暦の表。
   */
  export type BuiltinDates = 'default' | 'japanese-era';
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
      /** 1 シートの印刷範囲（`_xlnm.Print_Area` のカンマで並ぶ範囲）の数。超えた分は読まない */
      maxPrintAreas: number;
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
  /** 1 シートの印刷範囲の数の上限（範囲ごとにページを分けるので、ページの数が増えすぎないようにする） */
  export const MAX_PRINT_AREAS = 1000;
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
      /**
       * 印刷範囲（`_xlnm.Print_Area`。定義の順）。無いシートは持たない。
       * 列だけの範囲（`$A:$C`）は行が 1〜1,048,576、行だけの範囲（`$1:$5`）は列が 1〜16,384 になる
       */
      printArea?: readonly CellRange[];
  }
  /** セルの範囲（1 始まり・両端を含む。`r1 <= r2`・`c1 <= c2`） */
  export interface CellRange {
      readonly r1: number;
      readonly c1: number;
      readonly r2: number;
      readonly c2: number;
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
      /**
       * リッチテキストの区切り（共有文字列・inlineStr の `<r>`）。`text` を連結すると `CellView.text`。
       * `<rPr>` のある区切りが 1 つも無い文字列は持たない
       */
      runs?: readonly TextRun[];
  }
  /**
   * リッチテキストの区切りの書式（`<rPr>`）。区切りはセルの書式を継がず、これだけで描く
   * （持たない項目は既定＝太字なし・自動の色・既定の大きさ）。フォント名・上付き・下付きは持たない
   */
  export type RunFont = Pick<CellFormat, 'color' | 'bold' | 'italic' | 'strike' | 'underline' | 'sizeRatio'>;
  /** リッチテキストの 1 つの区切り（XSheet: ADR-0009 §1） */
  export interface TextRun {
      /** 区切りの文字（空にならない） */
      text: string;
      /** `<rPr>` のある区切りの書式。無ければセルの書式のまま */
      font?: RunFont;
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
      /**
       * 横の配置。標準（`general`）は持たない（型で決まる：数値・日付は右、真偽・エラーは中央、文字列は左）。
       * `distributed`・`fill`・`justify` は `left`、`centerContinuous` は `center` として持つ
       */
      hAlign?: 'left' | 'center' | 'right';
      /** 縦の配置。下（`bottom`・既定）は持たない。`justify`・`distributed` は `top` として持つ */
      vAlign?: 'top' | 'center';
      /** 折り返して全体を表示する（`wrapText`）。列幅（結合なら結合の幅）で折り返し、文字の中の改行も残す */
      wrap?: true;
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
   * - XSheet: 配置と折り返し（`<alignment>`）は cell-align.ts で読んで足す（ADR-0008 §1）。
   * - 色は SpreadsheetML の `<color>`（cell-color.ts）をテーマで解いてから持つ。解けない値は例外にせず既定の色・書式なしに倒す。
   * - 走査は線形（開きタグは `[^<>]*>`、閉じは `indexOf`）。
   */
  import { type Rgba } from 'xsheet/_/core/color';
  import type { RunFontReader } from 'xsheet/_/core/rich-text';
  import { type XlsxLimits } from 'xsheet/_/core/limits';
  import type { Theme } from 'xsheet/_/core/theme';
  import type { CellFormat } from 'xsheet/_/core/types';
  /** 書式の表 */
  export interface FormatTable {
      /** 書式の番号 → 書式（見える書式だけ・重複なし。`SheetRow.formats` の値・`XlsxSheet.formats` の位置） */
      readonly formats: readonly CellFormat[];
      /** cellXfs の位置（セルの `s`）→ 書式の番号。書式なし・不正・上限の外は -1 */
      readonly xfFormat: Int32Array;
      /** `<rPr>` の中身 → 区切りの書式（`<font>` と同じ読み方）。書式の表が空のときは無い（区切りの書式を読まない） */
      readonly readRunFont?: RunFontReader;
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

declare module 'xsheet/_/core/rich-text' {
  /**
   * リッチテキストの区切り（`<si>`・`<is>` の `<r>`）と、その書式（`<rPr>`）の読み取り（M7-16）。
   * - `<rPr>` の中身は `<font>` と同じ読み方（`cell-format.ts` が渡す読み手）。
   * - 区切りの文字を連結したものが平文と違えば、区切りを持たない（平文で描く）。
   * - 走査は線形（`forEachElement`・`findElement`。閉じは indexOf）。DOM 非依存。
   */
  import type { RunFont, TextRun } from 'xsheet/_/core/types';
  /** `<rPr>` の中身 → 区切りの書式 */
  export type RunFontReader = (rPrBody: string | null) => RunFont;
  /** 区切りの書式を持ちうる断片か（`<rPr` が無ければ読まない） */
  export function mayHaveRuns(inner: string): boolean;
  /**
   * `<si>`・`<is>` の中身から区切りを読む。
   * @param inner 開きタグと閉じタグの間の文字列
   * @param text セルの文字（平文。`MAX_CELL_TEXT` で切った後）。区切りも同じ長さで切る
   * @param readFont `<rPr>` の読み手
   * @returns `<rPr>` のある区切りが無い・連結が平文と合わないときは undefined
   */
  export function readRuns(inner: string, text: string, readFont: RunFontReader): TextRun[] | undefined;
}

declare module 'xsheet/_/core/find-text' {
  import type { XlsxSheet } from 'xsheet/_/core/types';
  /** `findText` の指定 */
  export interface FindTextOptions {
      /** `'partial'`＝文字の一部に一致・`'cell'`＝セルの文字全体が一致（既定 `'partial'`） */
      match?: 'partial' | 'cell';
      /** 英字などの大文字と小文字を区別しない（既定 true）。全角と半角はいつも区別する */
      ignoreCase?: boolean;
      /** セルの中の改行（`\r\n`・`\n`・`\r`）を照合のときだけ除く（既定 true）。位置は元の文字で返す */
      ignoreNewlines?: boolean;
      /** 非表示の行・列（高さ・幅 0 の行・列も）を調べない（既定 true） */
      skipHidden?: boolean;
      /** 中断（`AbortController` の `signal`）。中断されたら `AbortError`（`DOMException`）で reject する */
      signal?: {
          readonly aborted: boolean;
      };
      /** 返す一致の数の上限（既定 10,000。100,000 より大きい値は 100,000 にする） */
      maxMatches?: number;
  }
  /** 1 つの一致 */
  export interface TextMatch {
      /** 行（1 始まり） */
      r: number;
      /** 列（1 始まり） */
      c: number;
      /** 一致の始まり（`CellView.text` 上の UTF-16 の位置） */
      start: number;
      /** 一致の終わり（含まない） */
      end: number;
  }
  /** `findText` の結果 */
  export interface FindTextResult {
      /** 一致（行 → 列 → セルの中の位置の順） */
      matches: TextMatch[];
      /** `maxMatches` で打ち切ったか（上限を超える一致があった） */
      truncated: boolean;
  }
  /**
   * シートの中の文字を探す。表示形式を当てた後の文字（`CellView.text`）で照らす。
   * - 1 セルの中の重ならない一致をすべて返す。結合したセルは左上のセルだけ。空の検索語は空の結果。
   * - 図形・グラフの中の文字は探さない。
   * - 約 10ms ごとに休むので、大きなシートでも画面は固まらない。
   * - ブック全体を探すときは、呼ぶ側がシートを 1 つずつ `openSheet` で開いて呼ぶ。
   * @param sheet `XlsxBook.openSheet` で開いたシート
   * @param query 検索語
   * @throws DOMException `signal` で中断された（`name` は `'AbortError'`）
   */
  export function findText(sheet: XlsxSheet, query: string, options?: FindTextOptions): Promise<FindTextResult>;
}

declare module 'xsheet/_/render/index' {
  /**
   * render の入口（`xsheet/render`）。DOM の表・罫線・図の層。
   * InterHub `viewers/xlsx-*.ts`・`spreadsheet.ts` の .xlsx 部分から作る。`renderSheet` は M2-3、`renderWorkbook` は M2-4。
   */
  export { MESSAGES_EN, MESSAGES_JA, MESSAGES_KO, type XSheetMessages } from 'xsheet/_/render/messages';
  export { renderSheet, type CellPosition, type FocusHighlightOptions, type RenderOptions, type ScrollToCellOptions, type SheetView, } from 'xsheet/_/render/render-sheet';
  export { renderWorkbook, type WorkbookOptions, type WorkbookView } from 'xsheet/_/render/render-workbook';
  export type { Zoom } from 'xsheet/_/render/zoom';
  export type { PrintOptions, PrintResult } from 'xsheet/_/render/print';
  export type { Selection, SelectionInput, SelectionMode, SelectionOrigin, SelectionType, } from 'xsheet/_/render/selection';
}

declare module 'xsheet/_/render/messages' {
  /**
   * render の文言（`RenderOptions.messages`）。InterHub の `t('spreadsheet.…')` の代わり。
   * 既定は日本語。英語・韓国語も同梱し、使う側は一部のキーだけを渡して差し替えられる。
   * 差し込み（`{detail}`・`{format}`）は render が行う。文言は `textContent`・`title` にだけ入れる。
   */
  /** 文言の表（24 キー。元は InterHub `locales/ja.ts` の `spreadsheet.*` と `data.noData`） */
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
      /** XSheet: 選択のコピーのセルの数が上限を超えた（ADR-0011 §5） */
      noticeCopyTooLarge: string;
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
  import type { TextMatch, XlsxBook } from 'xsheet/_/core/index';
  import { type XSheetMessages } from 'xsheet/_/render/messages';
  import type { PrintOptions, PrintResult } from 'xsheet/_/render/print';
  import type { Selection, SelectionInput, SelectionMode, SelectionOrigin } from 'xsheet/_/render/selection';
  import type { Zoom } from 'xsheet/_/render/zoom';
  /** 表示の指定 */
  export interface RenderOptions {
      /** 文言（一部のキーだけでよい。既定は日本語の `MESSAGES_JA`） */
      messages?: Partial<XSheetMessages>;
      /** 図の層（画像・図形・線）を出すか（既定 true）。false なら図を読まず、図の通知も出さない */
      drawings?: boolean;
      /** 書式の `<style>` に付ける nonce（CSP の `style-src 'nonce-…'` 用） */
      nonce?: string;
      /**
       * 最初の倍率（既定 1）。最初から指定すると、1 で描いてから替えるちらつきが無い。
       * XSheet: ADR-0007 §5（M7-7）
       * @throws RangeError（`renderSheet`・`renderWorkbook` が投げる）数でない倍率・正しくない指定
       */
      zoom?: Zoom;
      /**
       * 行番号の列と列見出しを出すか（既定 true）。false なら出さず、セルを表示枠の左上から描く。
       * XSheet: ADR-0007 §7（M7-9）
       */
      headers?: boolean;
      /**
       * セル・行番号・列記号のクリックで選べるようにする（既定は無効＝クリックでは選ばない。`setSelection` はいつでも効く）。
       * `'row'` はセルのクリックでその行の全体（エクスプローラーのプレビューと同じ）、`'cell'` はそのセルを選ぶ。
       * 行番号・列記号のクリックは、どちらでも行・列の全体。
       * XSheet: ADR-0011 §1（M7-24）
       */
      selection?: SelectionMode;
      /**
       * 選択が変わったときに呼ぶ（同じ選択では呼ばない）。`origin` は `'user'`（クリック・キー）か `'api'`（`setSelection`・シートの切り替え）。
       * XSheet: ADR-0011 §1（M7-24）
       */
      onSelectionChange?: (selection: Selection | null, origin: SelectionOrigin) => void;
  }
  /** `focusHighlight` の指定（XSheet: ADR-0007 §3） */
  export interface FocusHighlightOptions {
      /** そのヒットが見える所までスクロールする（既定 true） */
      scroll?: boolean;
      /** なめらかにスクロールする（既定 false） */
      smooth?: boolean;
  }
  /** `scrollToCell` の指定（XSheet: ADR-0007 §3） */
  export interface ScrollToCellOptions {
      /** なめらかにスクロールする（既定 false） */
      smooth?: boolean;
  }
  /** 画面の位置にあるセル（1 始まり） */
  export interface CellPosition {
      r: number;
      c: number;
  }
  /** 表示 1 つ */
  export interface SheetView {
      /**
       * 検索の一致（`findText` の `matches`）を光らせる。`null`・空の配列で消す。
       * スクロールで描き直しても残り、シートを替えると消える。色は CSS 変数 `--xs-hit-bg`・`--xs-hit-current-bg`。
       * 一致の片は `span.xs-hit`（`data-xs-hit` に配列の番号）。収まらず `###` で出す数値のセルは、セル全体を光らせる。
       * XSheet: ADR-0007 §2（M7-5）
       */
      setHighlights(matches: readonly TextMatch[] | null): void;
      /**
       * 今のヒットを替え、そのヒットが見える所までスクロールする。描いている範囲の中なら表を描き直さず、印（`is-current`）だけを替える。
       * スクロール先は縦が表示枠の上から約 1/3、横はヒットが見えていないときだけ見える所まで（はみ出す文字の中のヒットにも届く）。
       * @param index `setHighlights` に渡した配列の番号（0 始まり）。-1 で外す
       * @throws RangeError 整数でない・-1 より小さい・配列の長さ以上の番号
       * XSheet: ADR-0007 §3（M7-6）
       */
      focusHighlight(index: number, options?: FocusHighlightOptions): void;
      /**
       * セルが見える所までスクロールする（縦は表示枠の上から約 1/3。横は見えていないときだけ動かす）。
       * @param r 行（1 始まり）
       * @param c 列（1 始まり）
       * @throws RangeError 1 以上の整数でない番号
       */
      scrollToCell(r: number, c: number, options?: ScrollToCellOptions): void;
      /**
       * 画面の座標（`MouseEvent.clientX`・`clientY`）にあるセル。見出し・スクロールバー・表の外は null。
       * 結合したセルは左上、はみ出した文字の範囲（文字の届かない右の空セルを含む）ははみ出し元のセルを返す。
       * XSheet: ADR-0007 §4（M7-6）
       */
      cellAt(clientX: number, clientY: number): CellPosition | null;
      /** 画面の座標にあるハイライトの番号（`setHighlights` の配列の番号）。無ければ -1 */
      highlightAt(clientX: number, clientY: number): number;
      /**
       * 倍率を替える。数は 0.5〜2.0（範囲の外は端の値）。`{ fit: 'width' }` はデータのある列と図の右端までを表示枠の幅に合わせ、
       * 枠の大きさが変わると合わせ直す（`min`・`max` の既定は 0.5・2.0）。セルの寸法と文字に倍率を掛けて描き直し、
       * 左上に見ていたセルを左上に残す。見出し（列記号・行番号の列の幅）には倍率を掛けない。線の太さは変わらない。
       * 表示の外側に `transform: scale()` を掛けないこと（`cellAt`・`highlightAt` の位置がずれる）。
       * @throws RangeError 数でない倍率（`NaN` など）・`{ fit: 'width' }` でない指定
       * XSheet: ADR-0007 §5（M7-7）
       */
      setZoom(zoom: Zoom): void;
      /** 今の倍率（fit のときは計算した値） */
      readonly zoom: number;
      /**
       * シートを替える（表示を作り直さない）。倍率は保ち（幅に合わせるなら合わせ直す）、ハイライトと今のヒットは消す。
       * 表示中のシートなら何もしない。開けないシートは、例外にせず枠の中にメッセージを出す。
       * @param index シートの番号（`SheetInfo.index`。1 始まり・非表示のシートも数える）
       * @throws RangeError ブックに無い番号
       * XSheet: ADR-0007 §7（M7-9）
       */
      showSheet(index: number): Promise<void>;
      /**
       * 選んでいるシートの番号（`SheetInfo.index`）。開けなかったシートでもその番号。表の無い表示（開けなかったブック）では null。
       * `showSheet` を呼ぶとすぐに替わる（表示し終えるのを待たない）。
       * XSheet: ADR-0007 §7（版 1.1）
       */
      readonly activeSheet: number | null;
      /**
       * スクロール枠（仮想化の窓 `div.xs-body`）。表示を作ってから `destroy` まで同じ要素で、表の無い表示（開けなかったブック・
       * シートが 0 枚）と `destroy` の後は null。位置を読む・`scrollTop`・`scrollLeft` を替える・`scroll` のイベントを受けるのに使う。
       * 大きさ・スタイル・子要素は替えないこと。シートを替えると左上に戻り、倍率を替えると見ていた左上のセルに合わせて動く
       * （どちらも `scroll` のイベントが来る）。
       * XSheet: ADR-0010 §1（M7-19）
       */
      readonly scrollElement: HTMLElement | null;
      /**
       * 表示中のシートを、印刷のために仮想化せずページに分けて `target` の中に描く（`div.xs-print` を 1 つ足す。`target` 自身には何も付けない）。
       * 範囲はデータ・図・はみ出しの届く列まで。ページは「下へ、それから右へ」の順で、各ページは `page` の大きさの固定の窓（境で切る）。
       * 画像・グラフを描き終えてから解決する。今の表示（倍率・位置・ハイライト）には触らない。
       * 印刷の画面（`window.print()`）は使う側が出し、`@page` の大きさを `page` と同じにする。終わったら結果の `destroy` を呼ぶ。
       * 見た目はいつもライトのテーマ。目盛り線と見出しは既定で出さない。表の無い表示・何も描く物の無いシートは 0 ページ。
       * @throws RangeError 正しくない指定
       * @throws DOMException AbortError（`signal`・表示の `destroy`・シートの切り替えで中断した。足した物は片付けてから拒否する）
       * XSheet: ADR-0010 §2（M7-20）
       */
      renderForPrint(target: HTMLElement, options?: PrintOptions): Promise<PrintResult>;
      /**
       * 今の選択（無ければ null）。
       * XSheet: ADR-0011 §1（M7-24）
       */
      getSelection(): Selection | null;
      /**
       * 選択を替える（null で外す）。行・列は表の外でもよい（帯は表の範囲で切る）。セルは結合なら結合の範囲に広げる。
       * スクロールはしない（見せるには `scrollToCell`）。シートを替えると外れる。
       * @throws RangeError 種類が正しくない・1 以上の整数でない・Excel の範囲（1,048,576 行・16,384 列）を超える番号
       * XSheet: ADR-0011 §1（M7-24）
       */
      setSelection(selection: SelectionInput | null): void;
      /**
       * 選択をコピーする文字（Ctrl+C でクリップボードに入る物と同じ）。見えている文字（表示形式を当てた後）をタブと `
    ` で区切る。
       * タブ・改行・`"` を含むセルは `"` で囲む。結合したセルは左上に文字、残りは空。非表示の行・列は入れない。
       * 行・列の全体は値のある範囲までに切る。選択が無い・100 万セルを超えると null。
       * XSheet: ADR-0011 §5（M7-26）
       */
      selectionText(): string | null;
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

declare module 'xsheet/_/render/print' {
  /**
   * 印刷の描画（`renderForPrint`・XSheet の追加。M7-20）。
   * 表示中のシートのデータの範囲を、用紙の中身の大きさのページに切り分けて、仮想化せずに描く。今の表示（倍率・位置）には触らない。
   * - ページは「下へ、それから右へ」の順（Excel の既定）。各ページは、その行・列の範囲を窓にした固定の表示：
   *   表（`buildGridTable`）・罫線の層（`svg.xs-borders` 1 枚）・図の層（範囲で切る入れ物の中）を重ね、ページの端で切る。
   *   ブラウザの改ページには任せない（幅のはみ出しは次のページへ送られず切れるため）。
   * - 範囲は、データ・図の右下・はみ出しの届く列（全部の行から）。枠を埋める空のセルは入れない。何も描く物の無いページは出さない。
   *   ファイルに印刷範囲（`_xlnm.Print_Area`）があれば、その範囲を定義の順に、範囲ごとにページを分けて描く（§3・M7-23）。
   * - 倍率はいつも 100%。`fit: 'width'` のときだけ、全部の列が 1 ページの幅に収まるまで縮める。
   * - 描くセルの数が上限を超えたら、超える手前の行で打ち切る（列は打ち切らない）。
   * - 画像・グラフを描き終えてから解決する。中断（`signal`・表示の destroy・シートの切り替え）では、足した物を片付けてから AbortError。
   * ユーザーのファイル内容は `textContent` だけで入れる（表・図の層の部品と同じ）。例外のメッセージにセル値・シート名を含めない。
   */
  import type { CellRange, colName, createGeometry, Geometry, SheetDrawing, XlsxBook, XlsxSheet } from 'xsheet/_/core/index';
  import type { createBorderLayer } from 'xsheet/_/render/border-layer';
  import type { buildBorderPaths } from 'xsheet/_/render/borders';
  import type { drawingExtent, DrawingBox, rectsIntersect } from 'xsheet/_/render/drawings';
  import type { buildFormatCss } from 'xsheet/_/render/format-css';
  import type { applyMarginFormats, collectVisibleRows, rowNumWidth } from 'xsheet/_/render/grid';
  import type { buildGridTable } from 'xsheet/_/render/grid-table';
  import type { createImageLayer } from 'xsheet/_/render/image-layer';
  import type { XSheetMessages } from 'xsheet/_/render/messages';
  import type { spillReachCols } from 'xsheet/_/render/spill-extent';
  import type { zoomGeometry } from 'xsheet/_/render/zoom';
  /**
   * 印刷の描画が使う、表示と共有の部品（print-gate.ts の `PRINT_KIT`）。
   * このファイルは遅れて読む部品なので、表示の部品は import せずに受け取る（単独版の xsheet-print.mjs に表示の部品を二重に入れないため。ADR-0010 §2）。
   */
  export interface PrintKit {
      colName: typeof colName;
      createGeometry: typeof createGeometry;
      createBorderLayer: typeof createBorderLayer;
      buildBorderPaths: typeof buildBorderPaths;
      drawingExtent: typeof drawingExtent;
      rectsIntersect: typeof rectsIntersect;
      buildFormatCss: typeof buildFormatCss;
      applyMarginFormats: typeof applyMarginFormats;
      collectVisibleRows: typeof collectVisibleRows;
      rowNumWidth: typeof rowNumWidth;
      buildGridTable: typeof buildGridTable;
      createImageLayer: typeof createImageLayer;
      spillReachCols: typeof spillReachCols;
      zoomGeometry: typeof zoomGeometry;
      /** 見出し行の高さ（grid.ts の `GRID_HEADER_HEIGHT`） */
      GRID_HEADER_HEIGHT: number;
      /** Excel の最大の行数・列数（grid.ts の `MAX_GRID_ROWS`・`MAX_GRID_COLS`） */
      MAX_GRID_ROWS: number;
      MAX_GRID_COLS: number;
  }
  /** 印刷の指定（ADR-0010 §2-1） */
  export interface PrintOptions {
      /**
       * 1 ページの中身の大きさ（CSS px。96px = 1in）。既定は A4 縦（210×297mm）から Excel の「標準」の余白
       * （左右 0.7in・上下 0.75in）を引いた 659×978。使う側の `@page` と同じ大きさにすること。
       */
      page?: {
          width: number;
          height: number;
      };
      /** `'width'` は全部の列が 1 ページの幅に収まるように縮める（縦だけページを分ける）。既定は縮めない（100%） */
      fit?: 'width';
      /** 列記号・行番号を各ページに出す（既定 false） */
      headers?: boolean;
      /** 目盛り線（セルの境の薄い線）を出す（既定 false）。罫線はいつも出す */
      gridlines?: boolean;
      /** 描くセルの数の上限（行 × 列。非表示の行・列は数えない）。既定 200000 */
      maxCells?: number;
      /**
       * ファイルの印刷範囲（`_xlnm.Print_Area`）を使わず、データの範囲を印刷する（Excel の「印刷範囲を無視」）。
       * 既定 false（印刷範囲があれば、その範囲だけを範囲ごとにページを分けて描く）
       */
      ignorePrintArea?: boolean;
      /** 中断する（描いている間・画像を待つ間に見る） */
      signal?: AbortSignal;
  }
  /** 印刷の描画の結果（ADR-0010 §2-1） */
  export interface PrintResult {
      /** 描いたページの数（表の無い表示・シートを開けていない・描く物の無いシートは 0） */
      readonly pages: number;
      /** 上限（`maxCells`）で行を打ち切った */
      readonly truncated: boolean;
      /** `target` に足した要素を外し、blob URL を放す。何度呼んでもよい */
      destroy(): void;
  }
  /** 既定の 1 ページの中身の大きさ（A4 縦・Excel の「標準」の余白。CSS px） */
  export const PRINT_PAGE_DEFAULT: Readonly<{
      width: number;
      height: number;
  }>;
  /** 描くセルの数の上限の既定（仮の値。ADR-0010 §2-5） */
  export const PRINT_MAX_CELLS_DEFAULT = 200000;
  /** 解いた印刷の指定 */
  export interface ResolvedPrintOptions {
      page: {
          width: number;
          height: number;
      };
      fit: boolean;
      headers: boolean;
      gridlines: boolean;
      maxCells: number;
      ignorePrintArea: boolean;
  }
  /**
   * 印刷の指定を解く。
   * @throws RangeError 用紙の大きさが正の有限の数でない・`fit` が `'width'` でない・`maxCells` が 1 以上の数でない
   */
  export function resolvePrintOptions(options?: PrintOptions): ResolvedPrintOptions;
  /** 1 ページの範囲（行・列は 1 始まり・両端を含む） */
  export interface PrintPage {
      r1: number;
      r2: number;
      c1: number;
      c2: number;
  }
  /** 1 つの範囲のページの切れ目 */
  export interface PrintAreaBands {
      /** 行のページの切れ目（上から） */
      rowBands: readonly [number, number][];
      /** 列のページの切れ目（左から） */
      colBands: readonly [number, number][];
  }
  /** 印刷の割り付け（DOM を使わない計算の結果） */
  export interface PrintPlan {
      /** 描く最後の行（全部の範囲で一番下） */
      rows: number;
      /** 描く最後の列（全部の範囲で一番右） */
      cols: number;
      /** 倍率（`fit: 'width'` のときだけ 1 未満） */
      zoom: number;
      truncated: boolean;
      /** 範囲ごとの切れ目（印刷範囲の順。印刷範囲の無いシートはデータの範囲 1 つ）。範囲ごとにページを分ける */
      areas: readonly PrintAreaBands[];
  }
  /** 割り付けに使うシートの部分 */
  export type PrintSheet = Pick<XlsxSheet, 'layout' | 'rowCount' | 'colCount' | 'getRows' | 'formats'>;
  /**
   * 番号 start〜end を、長さ limit に収まる区切りに分ける（ページの切れ目）。切れ目は行・列の境。
   * 1 つだけで limit を超える番号は、その番号だけの区切りにする。長さ 0 の番号（非表示）だけの区切りは返さない。
   *
   * @param edge 番号 i の始まりの位置（`rowTop`・`colLeft`）
   * @param at 位置を含む番号（`rowAt`・`colAt`）
   */
  export function splitBands(start: number, end: number, edge: (i: number) => number, at: (v: number) => number, limit: number): [number, number][];
  /**
   * 印刷の範囲・倍率・ページの切れ目を決める（ADR-0010 §2-2・§2-5・§3）。
   * - 印刷範囲が無ければ、行は データ・図の最下行まで、列は データ・図の最右列・はみ出しの届く列（見える全部の行から）まで。
   * - 印刷範囲があれば、その範囲を定義の順に使う（範囲の外へのはみ出しは広げない）。列だけ（`$A:$C`）の範囲は
   *   データ・図の最下行まで、行だけ（`$1:$5`）の範囲はデータ・図の最右列・はみ出しの届く列までに縮める。
   * - 見える行を上から数え、全部の範囲を通して見える行 × 見える列が `maxCells` を超える手前で打ち切る
   *   （最初の 1 行は必ず描く。打ち切った後の範囲は描かない）。
   */
  export function planPrint(kit: PrintKit, sheet: PrintSheet, boxes: readonly DrawingBox[], opts: ResolvedPrintOptions, printArea?: readonly CellRange[]): PrintPlan;
  /**
   * 寸法を 1 ページの窓にする：範囲の外の行・列は高さ・幅 0、位置は範囲の左上を 0 とする。
   * 表・罫線の部品にそのまま渡すと、範囲の外の余白を作らずに範囲だけを描く。
   */
  export function windowGeometry(g: Geometry, p: PrintPage): Geometry;
  /** 印刷の描画に使う、表示中のシートの状態（sheet-view.ts が渡す） */
  export interface PrintSource {
      /** `sheets` は印刷範囲（`SheetInfo.printArea`）を引くのに使う */
      book: Pick<XlsxBook, 'readImage' | 'readChart' | 'sheets'>;
      /** シートの番号（`SheetInfo.index`） */
      index: number;
      sheet: XlsxSheet;
      /** 図（図を出さない指定なら空） */
      drawings: readonly SheetDrawing[];
      /** 配置済みの図（A1 の左上基準・100% の px） */
      boxes: readonly DrawingBox[];
      messages: Readonly<XSheetMessages>;
      nonce?: string;
  }
  /**
   * 表示中のシートを、印刷のためにページに分けて `target` の中に描く（ADR-0010 §2）。
   * @throws RangeError 正しくない指定
   * @throws DOMException AbortError（中断。足した物は片付けてから拒否する）
   */
  export function renderPrint(kit: PrintKit, target: HTMLElement, src: PrintSource | null, options?: PrintOptions): Promise<PrintResult>;
}

declare module 'xsheet/_/render/border-layer' {
  /**
   * `.xlsx` の罫線層（ISSUE-98）：`svg.xs-borders` 1 枚の中身を、{@link buildBorderPaths} の結果から作る。
   * - `createElementNS` だけで作る（`innerHTML` は使わない）。`d`・`stroke-width`・`stroke-dasharray` は数値から組み立てた文字列だけ。
   * - 墨色（自動・黒）は class `is-ink` だけ（色は CSS のテーマ変数 `--xs-border-ink`。テーマの切替で描き直さない）。
   *   ほかの明示色は整数から組み立てた `#rrggbb` を CSSOM（`style.setProperty`）で入れる（CSP のため style 属性の文字列は組み立てない）。
   * - 二重線は、セルの面の色の隙間（`is-gap`・幅 3）の上に 1px の線 2 本。隙間は全 path の線より先（下）に描く。
   * - 位置と大きさは描いた線の外接矩形だけ（A1 の左上を原点とする座標を viewBox で保つ）。線が無ければ中身を空にする。
   */
  import type { BorderPaths } from 'xsheet/_/render/borders';
  /** 罫線層 */
  export interface BorderLayer {
      /** wrapper に置く svg（常に 1 枚。中身だけ入れ替わる） */
      readonly element: SVGSVGElement;
      /**
       * 結果で中身を作り直す。線が無ければ空にする。
       * @param result {@link buildBorderPaths} の結果（座標は A1 の左上が原点）
       * @param originX A1 の左上の x（行番号の列の幅）
       * @param originY A1 の左上の y（見出し行の高さ）
       */
      render(result: BorderPaths, originX: number, originY: number): void;
      /** 中身を空にする（シートの切替・後始末） */
      clear(): void;
  }
  /** 罫線層を作る（svg は常に 1 枚で、中身だけを `render` で入れ替える） */
  export function createBorderLayer(): BorderLayer;
}

declare module 'xsheet/_/render/borders' {
  /**
   * `.xlsx` のセルの罫線を、SVG の path にするための純関数（DOM 非依存・ISSUE-98）。
   * - 描いた行（`GridRow.borders`）と罫線の表（`XlsxSheet.borders`）から、すべての横の辺・縦の辺（と斜線）について
   *   「隣り合う 2 つの指定」から 1 つの線を決め（太さの順位 → 線種 → 左・上のセル）、結合セルの内側の辺は描かず、
   *   同じ境目の上で同じ線が続く区間を 1 本にまとめて、線種×色ごとの path の `d` にする。
   * - 線種の幅・パターン・太さの順位は {@link BORDER_STYLES}（定数 1 か所）。Excel の画像との照合（E1〜E5）で直すのはここだけ。
   * - 座標は A1 の左上を原点とする px（geometry の `colLeft`・`rowTop` と同じ座標）。1px の線はグリッド線の画素に重ねる。
   */
  import type { BorderEdge, BorderStyle, CellBorder, Geometry } from 'xsheet/_/core/index';
  import type { GridRow } from 'xsheet/_/render/grid';
  /** 1 回の描画で作る区間（線 1 本・斜線 1 本）の数の上限。超えた分は描かない（ADR-0029 §決定 7） */
  export const MAX_BORDER_SEGMENTS = 50000;
  /** 線種 1 つの描き方 */
  export interface BorderStyleSpec {
      /** 線の幅（px）。二重線は隙間を含む全体の幅（1px の線 2 本＋ 1px の隙間） */
      width: number;
      /** `stroke-dasharray` の数値（px）。実線は null */
      dash: readonly number[] | null;
      /** 太さの順位（大きいほど強い。辺の優先の 1 番目の基準） */
      weight: number;
      /** 線種の分類（同じ順位の中の優先。大きいほど強い） */
      kind: number;
      /** 二重線（隙間＋ 1px の線 2 本で描く） */
      double?: boolean;
  }
  /**
   * 13 線種の幅・パターン・優先の表（ADR-0029 §決定 3。暫定値で、E2 の照合で直す定数 1 か所）。
   * パターンの数値は px。`slantDashDot` の斜めの切り口は再現しない。
   */
  export const BORDER_STYLES: Readonly<Record<BorderStyle, BorderStyleSpec>>;
  /**
   * 同じ辺の 2 つの指定から、描く 1 つを選ぶ（ADR-0029 §決定 4）。
   * 片方だけならそれ。両方あれば、太さの順位が大きい方 → 同じなら線種の分類（double ＞ 実線 ＞ 破線系 ＞ dotted ＞ hair）
   * → それでも同じなら左（上）のセルの指定 `a`。線種が表に無い指定は無いものとして扱う。
   *
   * @param a 左（上）のセルの、この辺（右・下）の指定
   * @param b 右（下）のセルの、この辺（左・上）の指定
   */
  export function pickBorderEdge(a: BorderEdge | undefined, b: BorderEdge | undefined): BorderEdge | undefined;
  /** 1 本の path にまとめた線（線種×色×斜線かどうかごと） */
  export interface BorderPath {
      /** 線種 */
      style: BorderStyle;
      /** 墨色（自動・黒）。色はテーマ変数で決める（`color` は null） */
      ink: boolean;
      /** 明示色（`#rrggbb`）。墨色なら null */
      color: string | null;
      /** 明示色の不透明度（0〜1。1 なら不透明） */
      alpha: number;
      /** `stroke-width`（px）。二重線は隙間の幅（3） */
      width: number;
      /** `stroke-dasharray`。実線・二重線は null */
      dasharray: string | null;
      /** 斜線の path か（縦横の線は画素にそろえて描くが、斜線はそろえない） */
      diagonal: boolean;
      /** path の `d`。二重線は中心線（隙間の描画に使う） */
      d: string;
      /** 二重線だけ：1px の線 2 本の `d`（中心線から 1px ずつ離した 2 本） */
      lineD?: string;
  }
  /** 描いた線の外接矩形（A1 の左上を 0 とする px。線の幅を含む） */
  export interface BorderBounds {
      left: number;
      top: number;
      right: number;
      bottom: number;
  }
  /** {@link buildBorderPaths} の結果 */
  export interface BorderPaths {
      paths: BorderPath[];
      /** 線が 1 本も無ければ null */
      bounds: BorderBounds | null;
      /** 作った区間の数 */
      segments: number;
      /** 区間の上限で打ち切ったか */
      truncated: boolean;
  }
  /**
   * 描いた範囲の罫線を、線種×色ごとの path にする。
   *
   * - 範囲の左端の縦の辺は左隣の列、右端の縦の辺は右隣の列のセル（範囲外でも `rows` は全列を持つ）の指定も見る。
   *   上端の横の辺は下のセルの上の辺だけ、下端の横の辺は上のセルの下の辺だけで決める。
   * - 結合セルは `GridRow` の左上のセルの `colSpan`・`rowSpan`（行の畳みで縮んだ見える行の数）から範囲を復元し、
   *   内側の辺は描かない。外周は面する構成セルの辺から取る。斜線は左上のセルの指定で結合範囲の対角に引く。
   * - 幅 0 の列・行の辺は描かない（`rows` は見える行だけ）。非表示の列・行をはさむ辺は、前後の見える列・行の指定で決める。
   * - 同じ境目の上で同じ線が続く区間は 1 本（`M x1 y H x2`・`M x y1 V y2`）にまとめる。区間の数が `maxSegments` を超えたら
   *   以後は描かない（`truncated`）。
   *
   * @param rows 描く行（見える行だけ・行番号の昇順。全列を持つ）
   * @param table 罫線の表（`XlsxSheet.borders`）
   * @param geo 寸法（`colLeft`・`colWidth`・`rowTop`）
   * @param range 描く列の範囲（1 始まり・両端を含む）
   * @param maxSegments 区間の数の上限（既定 {@link MAX_BORDER_SEGMENTS}）
   */
  export function buildBorderPaths(rows: readonly GridRow[], table: readonly CellBorder[], geo: Pick<Geometry, 'colLeft' | 'colWidth' | 'rowTop'>, range: {
      first: number;
      last: number;
  }, maxSegments?: number): BorderPaths;
}

declare module 'xsheet/_/render/grid' {
  /**
   * `.xlsx` グリッド（`spreadsheet.ts`）の純ロジック（DOM 非依存・ISSUE-84）。
   * 表を Excel の寸法（`lib/xlsx/geometry.ts`）で固定して描くための計算だけを切り出す。
   * - 描く行の範囲・仮想化の上下の余白・セルの内側の高さは、すべて geometry から出す（DOM の実測はしない）。
   * - 非表示の行（高さ 0）は DOM に出さない。描く範囲と前後の余分は「見える行の数」で数え、結合の rowSpan は
   *   描く見える行の数へ縮める（ADR-0026 §決定 5 の予備案。列の畳みは `<col>` の `visibility: collapse` のまま）。
   * - 表の範囲は、データの行・列より広げて、スクロール枠を埋める（Excel のシートのように空セルが続く）。
   * - 文字列のはみ出し（右隣が空なら、その列幅の合計まで文字を出す）は、使える列数と幅の純関数にする。
   */
  import { type CellFormat, type CellKind, type CellView, type SheetLayout, type SheetRow, type Geometry } from 'xsheet/_/core/index';
  /** 列記号の見出し行の高さ（px）。`style.css` の `.xs-grid thead th` の `height` と同じ値にする */
  export const GRID_HEADER_HEIGHT = 24;
  /** Excel の最大行数・最大列数（表を画面いっぱいに広げるときの上限） */
  export const MAX_GRID_ROWS = 1048576;
  export const MAX_GRID_COLS = 16384;
  /** 隣り合う描く行の間に非表示の行が何行までなら、まとめて 1 回で読むか（結合の途中を正しく扱うため）。これを超える隙間は別々に読む */
  export const GAP_READ_MAX = 200;
  /** セルの内側の要素の左右の余白の合計（px）。`style.css` の `.xs-cell` の左右 padding（3px ずつ）と同じ値 */
  export const CELL_PAD_X = 6;
  /**
   * 行番号の列の幅（px）を行数の桁数から出す。1 桁 8px＋左右の余白 16px、最小 32px。
   * 列幅は固定レイアウトなので、最大の行番号（`rowCount`）が収まる幅を先に決める。
   *
   * @param rowCount シートの総行数
   */
  export function rowNumWidth(rowCount: number): number;
  /** 描く行の範囲（番号はすべて 1 始まり・両端を含む。見える行（高さ > 0）だけを数える。見える行が無いときはすべて 0） */
  export interface RowRange {
      /** 見えている最初の行 */
      top: number;
      /** 見えている最後の行 */
      bottom: number;
      /** 描く最初の行（`top` の手前の余分を含む。見える行） */
      first: number;
      /** 描く最後の行（`bottom` の先の余分を含む。見える行） */
      last: number;
  }
  /**
   * 行 r の次の見える行（高さ > 0）。無ければ 0。非表示の行が続いても `rowAt`（二分探索）で跳ぶ。
   * r の下端の y を含む行は、同じ y に並ぶ非表示の行より後ろの、次の見える行になる（geometry の `rowAt` の仕様）。
   *
   * @param total 行数の上限
   */
  export function nextVisibleRow(geo: Pick<Geometry, 'rowTop' | 'rowHeight' | 'rowAt'>, r: number, total: number): number;
  /** 行 r の前の見える行（高さ > 0）。無ければ 0 */
  export function prevVisibleRow(geo: Pick<Geometry, 'rowTop' | 'rowHeight' | 'rowAt'>, r: number): number;
  /**
   * スクロール位置から、見えている行と、描く行（前後に「見える行」で overscan 行ぶん足した範囲）を出す。
   * 行の高さは可変（Excel の行高・非表示）なので、`rowAt`（geometry の二分探索）で引く。
   * 非表示の行は数えない（フィルターで大半が畳まれていても、描く見える行は一定数）。
   *
   * @param geo `createGeometry` の結果
   * @param scrollTop スクロール枠の `scrollTop`（表の最初の行の上端を 0 とみなす。見出し行は sticky で重なるため無視してよい）
   * @param clientHeight スクロール枠の表示高さ（px）
   * @param total 行数
   * @param overscan 見えている範囲の前後に足す見える行の数
   */
  export function visibleRowRange(geo: Pick<Geometry, 'rowAt' | 'rowTop' | 'rowHeight'>, scrollTop: number, clientHeight: number, total: number, overscan: number): RowRange;
  /**
   * 仮想化で描かない上下の行ぶんの高さ（px）。上＝`rowTop(first)`、下＝`totalHeight − rowTop(last + 1)`。
   * 上の余白＋描く行の高さ＋下の余白が、いつも `totalHeight` に一致する。
   */
  export function spacerHeights(geo: Pick<Geometry, 'rowTop' | 'totalHeight'>, first: number, last: number): {
      top: number;
      bottom: number;
  };
  /** 枠の幅が分からない（0＝まだレイアウト前）とき、描く列の範囲に使う代わりの幅（px） */
  export const COL_FALLBACK_VIEW_WIDTH = 1280;
  /** 描く列の範囲（番号はすべて 1 始まり・両端を含む。`RowRange` の列版） */
  export interface ColRange {
      /** 見えている最初の列 */
      left: number;
      /** 見えている最後の列 */
      right: number;
      /** 描く最初の列（`left` の手前に 1 画面分の余白を含む） */
      first: number;
      /** 描く最後の列（`right` の先に 1 画面分の余白を含む） */
      last: number;
  }
  /**
   * スクロール位置から、見えている列と、描く列（左右に 1 画面分の余白を足した範囲）を出す（列の仮想化）。
   * 列幅は可変（Excel の列幅・非表示）なので、`colAt`（geometry の二分探索）で引く。範囲は px で決めるため、
   * 列が極端に狭くても広くても描く列数は画面幅に比例する（遠い列に図や値があっても、16,384 列を描かない）。
   *
   * @param geo `createGeometry` の結果（列数は表の列数）
   * @param scrollLeft スクロール枠の `scrollLeft`
   * @param clientWidth スクロール枠の表示幅（px）。0 以下ならレイアウト前とみなして代わりの幅を使う
   * @param originX 表の A1 の左上の x（行番号の列の幅。スクロール枠の内容の左端から）
   * @param total 列数
   */
  export function visibleColRange(geo: Pick<Geometry, 'colAt'>, scrollLeft: number, clientWidth: number, originX: number, total: number): ColRange;
  /**
   * 描かない左右の列ぶんの幅（px）。左＝`colLeft(first)`、右＝`totalWidth − colLeft(last + 1)`。
   * 左の余白＋描く列の幅＋右の余白が、いつも `totalWidth` に一致する（`spacerHeights` の列版）。
   */
  export function colSpacerWidths(geo: Pick<Geometry, 'colLeft' | 'totalWidth'>, first: number, last: number): {
      left: number;
      right: number;
  };
  /**
   * 行のセルを、描く列 first〜last（1 始まり）だけにする。返す行の `cells[i]` は列 `first + i`。
   * - 描く範囲の左端が結合セルの途中（結合の起点が範囲の左外）なら、起点のセルを範囲の左端へ移し、`colSpan` を範囲内へ縮める。
   *   起点が上の行にある（`rowSpan` で覆われる）セルは、これまでどおり null のまま（td を作らない）。
   * - 右端で切れる結合セルは、範囲の右端までに `colSpan` を縮める。縮めたセルは結合のまま扱う（はみ出さない）。
   *
   * @param row 全列を持つ行
   * @param first 描く最初の列
   * @param last 描く最後の列
   */
  export function clipRowToColumns(row: GridRow, first: number, last: number): GridRow;
  /**
   * 表の範囲（データの行・列）の外にある余白の位置へ、行の書式 → 列の書式の順で書式の番号を当てる（ADR-0030 §決定 7 の余白への拡張）。
   * 読み取り（`getRows`）は表の範囲の中だけを解くため、余白の行（`r > dataRows`）と余白の列（`列 > dataCols`）は、
   * Excel で行・列全体に続く塗りがここで抜ける。表の範囲は広げない（描く余白の位置に番号を足すだけ）。
   * - 行の書式（`rowFormat`）があればそれ（-1＝書式なしは何も当てず、列の書式も打ち消す）。無ければ列の書式（`colFormat`・-1 は無し）。
   * - 描く列 first〜last のうち、余白の位置だけを見る。表の中の位置（読み取りが解いた番号）は変えない。
   * - 余白に触れる行にだけ `rowFormat` を 1 回呼ぶ。`colFormat` は列ごとに 1 回（結果を覚える）。入力の行は書き換えず、
   *   足した行だけ新しい `GridRow` にする（足すものが無い行は同じ行をそのまま返す）。
   *
   * @param rows 描く行（`collectVisibleRows` の結果）
   * @param first 描く最初の列（1 始まり）
   * @param last 描く最後の列
   * @param dataRows 表のデータの行数（`XlsxSheet.rowCount`）
   * @param dataCols 表のデータの列数（`XlsxSheet.colCount`）
   * @param rowFormat 行 r の行の書式の番号（`XlsxSheet.rowFormat`）
   * @param colFormat 列 c（1 始まり）の列の書式の番号（`XlsxSheet.colFormat`）
   */
  export function applyMarginFormats(rows: readonly GridRow[], first: number, last: number, dataRows: number, dataCols: number, rowFormat: (r: number) => number | undefined, colFormat: (c: number) => number): GridRow[];
  /**
   * セルの内側の要素の高さ（px）。セル（td）の高さ（行高、結合は描く行の高さの合計）から下罫線 1px を引く。0 を下回らない。
   */
  export function cellInnerHeight(cellHeight: number): number;
  /**
   * 描く行のうち、i 番目の行から span 行ぶん（結合した、描く見える行の数）の高さの合計（px）。
   *
   * @param rows 描く行（見える行だけ。上から順）
   */
  export function spanHeight(geo: Pick<Geometry, 'rowHeight'>, rows: readonly {
      r: number;
  }[], i: number, span: number): number;
  /** 描くセル。`CellView` に、結合セルかどうかを足したもの */
  export interface GridCell extends CellView {
      /** 結合セル。行の畳みで `rowSpan` が 1 になっても結合のまま扱う（はみ出さない） */
      merged?: boolean;
  }
  /** 描く行（`SheetRow` の上位互換。見える行だけ） */
  export interface GridRow {
      /** 1 始まりの行番号 */
      r: number;
      /** 長さは列数以下。範囲より右・範囲外の列は `undefined`（空のセル）、null＝結合で隠れる */
      cells: (GridCell | null | undefined)[];
      /**
       * 罫線の番号（`SheetRow.borders` を写したもの。長さは `cells` と同じ・値は `XlsxSheet.borders` の位置）。
       * 罫線のあるセルが無い行は持たない。結合で隠れるセル（`cells` が null）にも番号が入る。
       */
      borders?: (number | undefined)[];
      /**
       * 書式の番号（`SheetRow.formats` を写したもの。長さは `cells` と同じ・値は `XlsxSheet.formats` の位置）。
       * 書式のあるセルが無い行は持たない。結合で隠れるセル（`cells` が null）にも番号が入る。
       */
      formats?: (number | undefined)[];
  }
  /**
   * 文字列がはみ出して使える列数（自分の列を含む）を決める。Excel の「文字列は右隣が空ならはみ出し、
   * 値のあるセルの手前で切る」を表す。td は 1 セル 1 個のまま（罫線を残す）で、文字だけをはみ出させる。
   * - 文字列のとき（`kind === 's'` で非空、結合の無いセル）だけ右隣の空セルを数える。型で判定する
   *   （数字だけの文字列・「▲1,234」は文字列なのではみ出す。数値・真偽・エラー・日付、型の無いセルは自分の列で切る）。
   * - 結合セル・空のセルは 1（自分の列だけ）。
   * - 値のあるセル（数式の結果が "" の文字列セルを含む）・結合で隠れるセル（null）・結合セルの手前で止める。HTML の `colspan` 上限と揃えて 1000 列まで。
   *
   * @param cells 行のセル（`GridRow.cells`）
   * @param c 対象の列（0 始まり）
   * @param cols 表の列数。`cells` より右の列（足した空の列）も空のセルとして数える
   * @returns 使える列数（1 以上）。1 ならはみ出さない
   */
  export function spillColumns(cells: readonly (GridCell | null | undefined)[], c: number, cols?: number): number;
  /**
   * はみ出しに「紙でない塗り」が絡むか（ADR-0030 §決定 9）。はみ出す元のセル（列 c）と、`spillColumns` が数えた右の列のどれかに
   * 白以外の塗りがあれば true（入れ物へ `is-over-fill` を付け、文字の下に面の色を敷かない）。白の塗りは塗りなしとして数える。
   * 追加の走査はせず、数えた列の `formats` を見るだけ。
   *
   * @param table 書式の表（`XlsxSheet.formats`）
   * @param formats 行の書式の番号（`GridRow.formats`。描く列に揃えたもの。無ければ false）
   * @param c はみ出す元の列（0 始まり）
   * @param columns `spillColumns` の結果（自分の列を含む列数）
   */
  export function spillOverFill(table: readonly CellFormat[], formats: readonly (number | undefined)[] | undefined, c: number, columns: number): boolean;
  /**
   * はみ出す文字の入れ物の幅（px）。自分の列から `columns` 列ぶんの列幅の合計から、右端の罫線 1px を引く。
   * 非表示の列（幅 0）は合計に入らない。自分の列が非表示（幅 0）なら 0（隠れたセルの文字を出さない）。
   *
   * @param geo `createGeometry` の結果
   * @param c 対象の列（0 始まり）
   * @param columns `spillColumns` の結果（自分の列を含む列数）
   */
  export function spillWidth(geo: Pick<Geometry, 'colLeft' | 'colWidth'>, c: number, columns: number): number;
  /**
   * 表を画面いっぱいに広げたときの行数・列数。Excel のシートのように、データの先にも空のセル
   * （既定の幅と高さ。geometry は範囲外の番号も既定の寸法で延ばす）が続くようにする。
   * 行数＝max(データの行数, `minRows`, スクロール枠の高さを埋める行数)、列数も同じ。上限は Excel の最大（1,048,576 行・16,384 列）。
   * 後続の画像（図のある範囲まで表を広げる）は `minRows`・`minCols` に図の最下行・最右列を渡す。
   *
   * @param layout `XlsxSheet.layout`
   * @param dataRows データの行数（`XlsxSheet.rowCount`）
   * @param dataCols データの列数（`XlsxSheet.colCount`）
   * @param viewWidth スクロール枠の表示幅（px。`clientWidth`）
   * @param viewHeight スクロール枠の表示高さ（px。`clientHeight`）
   * @param minRows 表示する行数の最小（既定 0）
   * @param minCols 表示する列数の最小（既定 0）
   */
  export function gridExtent(layout: SheetLayout, dataRows: number, dataCols: number, viewWidth: number, viewHeight: number, minRows?: number, minCols?: number): {
      rows: number;
      cols: number;
  };
  /**
   * 行 first〜last のうち、見える行（高さ > 0）だけを DOM 用の行にして返す。
   * 非表示の行は返さない。結合セル（`rowSpan` > 1）は、描く見える行の数へ縮める：
   * - 先頭の行が非表示の結合は、結合内の最初に見える行へセルを移す（値はその行の td に出る）。
   * - 非表示の行を跨ぐ結合の `rowSpan` は、見える行の数にする。
   * 隣り合う見える行の間の非表示の行が `gapMax` 行までなら、まとめて 1 回で読む（結合の途中を正しく扱うため）。
   * これを超える隙間（フィルターで大半を畳んだシートなど）は別々に読み、隙間の向こうへ続く結合は、
   * 次の見える行に（同じ文字の）結合セルが現れる。
   *
   * @param geo `createGeometry` の結果
   * @param first 描く最初の行
   * @param last 描く最後の行
   * @param total 行数
   * @param readRows 行 a〜b（連続）を読む関数（`XlsxSheet.getRows` と同じ形・範囲外の行は呼び出し側で空の行にする）
   * @param gapMax まとめて読む隙間の上限
   */
  export function collectVisibleRows(geo: Pick<Geometry, 'rowTop' | 'rowHeight' | 'rowAt'>, first: number, last: number, total: number, readRows: (a: number, b: number) => SheetRow[], gapMax?: number): GridRow[];
  /**
   * 数値の表示文字の推定幅（px）。DOM は測らず「文字数 × MDW」で見積もる。
   * 数字・符号・E・％・通貨記号などは MDW、小数点・桁区切り・コロン・スラッシュ（日付）・空白は MDW の半分とする。
   *
   * @param text 表示文字
   * @param mdw 最大桁幅 MDW（`SheetLayout.mdw`）
   */
  export function estimateNumberWidth(text: string, mdw: number): number;
  /**
   * 収まらない数値を `###` にする（Excel と同じ）。数値・日付（`kind` が `n`・`d`）の表示文字の推定幅が
   * セルの内側の幅（セル幅 − 右罫線 1px − 左右の余白）を超えたら、`#` の繰り返し（内側の幅 ÷ MDW ÷ 文字サイズの比（1 以上）の切り捨て個数、最低 1）を返す。
   * 収まるときは表示文字のまま。文字列・真偽・エラーは対象外（そのまま）。
   *
   * @param text 表示文字
   * @param kind セルの型
   * @param cellWidth セルの幅（px。結合は結合範囲の列幅の合計）
   * @param mdw 最大桁幅 MDW
   * @param sizeRatio 文字サイズの比（`CellFormat.sizeRatio`。推定幅に掛ける。既定 1。太字は掛けない・ADR-0030 §決定 10）
   */
  export function fitNumberText(text: string, kind: CellKind | undefined, cellWidth: number, mdw: number, sizeRatio?: number): string;
  /**
   * セルの型ごとの既定の配置（Excel の「標準」。セルの配置指定は読まない）。
   * 数値・日付は右、真偽・エラーは中央、文字列・型なしは左（クラス無し）。
   */
  export function cellAlignClass(kind: CellKind | undefined): 'is-right' | 'is-center' | '';
}

declare module 'xsheet/_/render/drawings' {
  /**
   * `.xlsx` の図（画像・図形など）をセルの上に重ねるための純ロジック（DOM 非依存・ISSUE-86）。
   * - アンカー（`lib/xlsx` の `DrawingAnchor`）を、A1 の左上を原点とする px の矩形にする（geometry から出す）。
   * - 図の最下行・最右列（表をそこまで広げる）・スクロール枠の可視範囲との交差判定を出す。
   * 位置は列幅・行高と同じ geometry から計算するため、表のセルと 1px もずれない。
   * グループの中の図（ISSUE-89）は、最も外側のグループのアンカーの矩形へ、グループの子の座標系（`chOff`・`chExt`）から線形に写して矩形にする。
   */
  import type { AnchorPoint, DrawingAnchor, SheetDrawing, Geometry, DrawingGroup } from 'xsheet/_/core/index';
  import { type LineRender, type ShapeDrawn } from 'xsheet/_/render/shapes';
  /** 1px あたりの EMU（Office の英語メートル単位。px＝EMU / 9525） */
  export const EMU_PER_PX = 9525;
  /** A1 の左上を原点とする矩形（px・整数） */
  export interface PxRect {
      x: number;
      y: number;
      w: number;
      h: number;
  }
  /**
   * 配置済みの図（`index` は `SheetDrawing[]` の位置。重ね順はこの昇順＝drawing XML の順）。
   * 線（kind が line）の `rect` は、アンカーの矩形ではなく `line.bbox`（矢じり・太さ・U 字の張り出しを含む外接矩形）。
   * 形で描く図形（`shape`）の `rect` は、経路の外接矩形とアンカーの矩形を合わせた矩形（`shape.rect`。ISSUE-92）。
   */
  export interface DrawingBox {
      index: number;
      rect: PxRect;
      /** 線（kind が line）の描画結果。描画側が再計算しない。線以外は持たない */
      line?: LineRender;
      /** 形で描く図形（kind が shape・textbox で、形を評価できたもの）の描画結果。形を評価できない図形（代替）・それ以外は持たない */
      shape?: ShapeDrawn;
  }
  /** アンカーの矩形の計算に使う geometry の部分（テストで差し替えやすいよう必要なものだけ） */
  export type AnchorGeometry = Pick<Geometry, 'colLeft' | 'colWidth' | 'rowTop' | 'rowHeight'>;
  /** 図の配置に使う geometry の部分（`AnchorGeometry` ＋ 線をシートの範囲に丸めるための全幅・全高） */
  export type DrawingGeometry = AnchorGeometry & Pick<Geometry, 'totalWidth' | 'totalHeight'>;
  /**
   * アンカーの点の x（px）：その列の左端＋ずれ。ずれが列の幅を超える時は列の右端で止める
   * （列幅の僅差で隣のセルへ飛び出さない。ADR-0026 §決定 8）。
   * geometry の番号は 1 始まり、アンカーの `col` は 0 始まり。
   */
  export function anchorPointX(geo: AnchorGeometry, p: AnchorPoint): number;
  /** アンカーの点の y（px）：その行の上端＋ずれ（行の高さを超える時は行の下端で止める） */
  export function anchorPointY(geo: AnchorGeometry, p: AnchorPoint): number;
  /**
   * グループの中の図の矩形（px・整数）を出す（ADR-0027 §決定 5）。
   * `anchor` は最も外側のグループのアンカーの矩形。グループの子の座標系（`chOff`・`chExt`）から親の矩形へ、外側から順に線形に写す：
   * `x = 親.x + (子.off.x − chOff.x) × 親.w / chExt.cx`（y・幅・高さも同じ）。入れ子のグループは、内側のグループ自身の `off`・`ext` を外側の座標系から写して、その矩形を次の親にする。
   * グループ自身の回転・反転は無視する。図の `xfrm` は回転前の矩形そのものなので、縦横の入れ替えはしない。
   * 端は四捨五入して幅・高さは端の差で出す。矩形は A1 の左・上へはみ出しても切らない（切ると線の傾き・文字の折り返しが変わる。はみ出しは
   * 線の `clampPath` とレイヤー（`overflow: hidden`）が隠す）。位置は ±MAX_PX、大きさは MAX_PX までで止める。座標系が壊れている・有限でない値になるときは null。
   */
  export function groupChildRect(anchor: PxRect, group: DrawingGroup): PxRect | null;
  /**
   * 図の矩形（px）。トップレベルの図はアンカーの矩形、グループの中の図は最も外側のグループのアンカーから写した矩形
   * （幅・高さは 0 や負のままで、大きさの判定はしない。判定は `layoutDrawings` が行う）。
   */
  export function placeDrawing(geo: AnchorGeometry, drawing: SheetDrawing): PxRect | null;
  /**
   * アンカーを A1 の左上基準の矩形（px）にする。大きさが 0 以下の図（非表示の行・列だけにかかる・逆向きの twoCell・
   * 大きさ 0 の ext）は null（出さない）。
   * - twoCell：`from`〜`to` の差。
   * - oneCell：`from` を起点に `cx`×`cy`。
   * - absolute：シートの左上からの `x`・`y` に `cx`×`cy`。
   */
  export function anchorRect(geo: AnchorGeometry, anchor: DrawingAnchor): PxRect | null;
  /**
   * 図の一覧を配置する。`index` を元の位置のまま持つので重ね順が保たれる。
   * - グループの中の図（`group` を持つ）：矩形は最も外側のグループのアンカーから写した値（`placeDrawing`）。`xfrm` は回転前の矩形なので、線の描画では縦横を入れ替えない。
   * - 図形（kind が shape・textbox で形を持つもの。ISSUE-92）：大きさが 0 以下の図は除く。形を評価できれば `rect` は経路の外接矩形とアンカーの矩形を合わせた矩形で、
   *   `shape` に描画結果を持つ。枠も塗りも文字も無い図形（Excel でも見えない）は除く。形を評価できない図形は、次の「線以外」と同じ（代替の描き方にする）。
   * - 線以外：大きさが 0 以下の図は除く。`rect` はアンカーの矩形（グループの中の図は写した矩形）。
   * - 線（kind が line）：アンカーの幅か高さが 0 でも残す（水平・垂直の線）。幅・高さが両方 0 以下のときだけ除く
   *   （負の値＝逆向きの twoCell は 0 とみなす）。`rect` は `renderLine` の外接矩形（シートの範囲に丸め済み）で、
   *   `line` に描画結果を持つ。外接矩形が空（シートの範囲の外へ追い出された）なら除く。
   *   `SheetDrawing.line` が無い線（壊れたデータ）は既定の直線として扱う。
   */
  export function layoutDrawings(geo: DrawingGeometry, drawings: readonly SheetDrawing[]): DrawingBox[];
  /**
   * 図を収めるのに必要な表の行数・列数（最下行・最右列。1 始まり）。図が無ければ 0。
   * `gridExtent` の `minRows`・`minCols` に渡して、図のある範囲まで表を広げる。
   *
   * @param geo 行数・列数の上限を Excel の最大にした geometry（`rowAt`・`colAt` が図の端を含む行・列へ届くように）
   */
  export function drawingExtent(geo: Pick<Geometry, 'rowAt' | 'colAt'>, boxes: readonly DrawingBox[]): {
      rows: number;
      cols: number;
  };
  /**
   * スクロール枠の可視範囲（A1 の左上基準・px）に、上下左右 1 画面分の余白を足した範囲。
   * 表の原点（A1 の左上）はスクロール枠の内容の (`originX`, `originY`)＝行番号の列の幅・見出し行の高さ。
   * 枠の大きさが 0（まだ DOM に付いていない）なら幅・高さ 0 の範囲を返す（何も交差しない）。
   */
  export function viewWindow(scrollLeft: number, scrollTop: number, clientWidth: number, clientHeight: number, originX: number, originY: number): PxRect;
  /** 2 つの矩形が交わるか（接するだけは交わらない。幅・高さ 0 の矩形は何とも交わらない） */
  export function rectsIntersect(a: PxRect, b: PxRect): boolean;
  /** 範囲 `win` と交わる図だけを、元の順（重ね順）で返す。図は最大 1,000 個なので毎回の線形走査でよい */
  export function pickVisible(boxes: readonly DrawingBox[], win: PxRect): DrawingBox[];
}

declare module 'xsheet/_/render/shapes' {
  /**
   * `.xlsx` の線・矢印を SVG で描くための純ロジック（DOM 非依存・ISSUE-88）。
   * 流れ：アンカーの矩形 → 回転前の矩形 → 形の経路（prst・調整値）→ 反転 → 中心で時計回りに回転
   * → シートの範囲に丸める → 矢じり（端の向き）・線の端の縮め・外接矩形 → SVG 用の `d` 文字列。
   * 座標はすべて px（A1 の左上を原点、y は下向き）。`<marker>` は使わず、矢じりは多角形・楕円として計算する。
   * 描画（`createElementNS`）は後続の `image-layer.ts` が行い、ここは数値と数値だけの文字列を返す。
   * 図形の形（ISSUE-92・ADR-0028）も同じ写像（反転 → 回転 → 丸め）で扱う：`renderShape` が、形の評価器（geom の `evaluate.ts`）の経路を
   * 部分経路の列（閉じ＝`Z`・path ごとの塗り・線・面の濃淡）・外接矩形・文字領域にする。
   */
  import type { SheetDrawing, DrawingLine, LineDash, LineEnd } from 'xsheet/_/core/index';
  import type { PxRect } from 'xsheet/_/render/drawings';
  import type { PathSeg, Pt } from 'xsheet/_/geom/index';
  /**
   * 形の表と評価器（geom）を読む（M2-3）。表が大きいため render からは動的 import にし、形のある図形を描くシートを開く時だけ読む。
   * 何度呼んでも読むのは 1 回（約束を共有する）。読めなければ例外（呼び出し側は代替で描き、通知を出す）。
   */
  export function loadGeom(): Promise<void>;
  /** geom を読み込み済みか（未読なら `renderShape` は形のある図形も代替を返す） */
  export function isGeomLoaded(): boolean;
  /** 点（px）と経路の区間（直線 `L`・3 次ベジェ `C`）。定義は geom の `defs.ts`（評価器と同じ型） */
  export type { PathSeg, Pt };
  /**
   * 経路（部分経路）：始点と、そこから続く区間の列。
   * `closed` が true なら閉じる（`Z`。図形の経路）。線の経路は持たない（開いたまま）。
   */
  export interface ShapePath {
      start: Pt;
      segs: PathSeg[];
      closed?: boolean;
  }
  /** シートの範囲（0〜全幅・0〜全高。geometry の `totalWidth`・`totalHeight`） */
  export interface SheetBounds {
      totalWidth: number;
      totalHeight: number;
  }
  /** 回転・反転（`SheetDrawing` の `rot`［1/60000 度］・`flipH`・`flipV`。無ければ 0・なし） */
  export interface Xfrm {
      rot?: number;
      flipH?: boolean;
      flipV?: boolean;
  }
  /** 矢じり 1 つの形（多角形か楕円。座標系は呼び出し元が決める） */
  export type ArrowShape = {
      kind: 'polygon';
      /** どちらの端か（`head`＝経路の始点、`tail`＝終点） */
      end: 'head' | 'tail';
      points: Pt[];
      /** 塗る（`arrow`＝開いた V だけ false。線で描く） */
      filled: boolean;
      /** 閉じる（`Z`。開いた V だけ false） */
      closed: boolean;
  } | {
      kind: 'ellipse';
      end: 'head' | 'tail';
      cx: number;
      cy: number;
      rx: number;
      ry: number;
      /** 長軸（rx 側）の向き（度。時計回り＝y 下向きの座標系） */
      angle: number;
  };
  /** 1 本の線の描画結果（`renderLine` の戻り値） */
  export interface LineRender {
      /** 外接矩形（シートの原点基準・px・シートの範囲に丸め済み）。`<svg>` の位置・可視判定・`drawingExtent` に使う */
      bbox: PxRect;
      /** `<path d>`：外接矩形の左上を原点とした座標。線の端は矢じりの分だけ縮めてある */
      d: string;
      /** 線の太さ（px。1〜16） */
      strokeWidth: number;
      /** `stroke-dasharray` の数値（空なら実線） */
      dash: number[];
      /** 矢じり（始点・終点の順。無いものは出さない）。座標は外接矩形の左上を原点とする */
      heads: ArrowShape[];
      /** 経路（シートの原点基準・丸め済み・線の端は縮める前）。テスト・後続の判定用 */
      path: ShapePath;
  }
  /** 図形の枠線の miterlimit（SVG の既定値。`stroke-miterlimit` に明示し、外接矩形の余白もこれで決める。ISSUE-92） */
  export const SHAPE_MITER_LIMIT = 4;
  /**
   * アンカーの矩形から「回転前の矩形」を出す。図の回転が 45° 以上 135° 未満、または 225° 以上 315° 未満なら、
   * アンカー（回転後の外接）は縦横が入れ替わっているため、幅と高さを入れ替える。中心はアンカーと同じ（ADR-0027 §決定 4）。
   *
   * 45°・135°・225°・315° ちょうどの境界は、Excel のアンカーが縦横を入れ替えたか入れ替えていないかが回転の値だけでは決まらない（45° ちょうどの
   * 図形は入れ替えていない）。`ext`（`a:xfrm` の `a:ext`＝回転前の大きさ。px）があれば、アンカーの幅・高さを `ext` に近い向きへ合わせて決める
   * （入れ替えた方が `ext` に近ければ入れ替える。差が 1px 以内＝正方形に近いなど決められないときは、回転の値で決める。ISSUE-93）。
   *
   * @param rect アンカーの矩形（px。幅・高さは 0 でもよい）
   * @param rot 回転（1/60000 度。無ければ 0）
   * @param ext 回転前の大きさ（px。無ければ回転の値だけで決める）
   */
  export function preRotationRect(rect: PxRect, rot: number | undefined, ext?: {
      w: number;
      h: number;
  }): PxRect;
  /** 図形の回転前の大きさ（`xfrmExt`。EMU）を px にする。持たない・読めないなら undefined（`preRotationRect` の `ext` に渡す） */
  export function xfrmExtPx(drawing: Pick<SheetDrawing, 'xfrmExt'>): {
      w: number;
      h: number;
  } | undefined;
  /**
   * 形（`prst`）と調整値から、回転前の矩形（幅 w・高さ h・左上が原点）の中の経路を作る。式は ECMA-376 の preset 定義に合わせる。
   * - `line`・`straightConnector1`・未知の値：(0,0)→(w,h) の直線。
   * - `bentConnector2〜5`：折れ線。`curvedConnector4・5` は同じ調整値の `bentConnector4・5` の折れ線で代替する。
   * - `curvedConnector2・3`：3 次ベジェ。
   * 調整値は割合のまま使い（範囲外も張り出しとして許す）、丸めは回転後の点に対して行う（`clampPath`）。
   *
   * @param adj 調整値 adj1〜adj3（1/100000。無い要素は 50000）
   */
  export function shapePath(geom: string, adj: readonly (number | undefined)[] | undefined, w: number, h: number): ShapePath;
  /**
   * 矩形の中の経路を、反転（`flipH`：x→w-x、`flipV`：y→h-y）してから、矩形の中心で時計回りに `rot` 回転し、
   * シートの座標（px）にする。画面座標（y 下向き）では `x' = cx + dx·cos − dy·sin`、`y' = cy + dx·sin + dy·cos`。
   *
   * @param path 矩形の左上を原点とする経路（`shapePath` の結果）
   * @param box 回転前の矩形（`preRotationRect` の結果）
   */
  export function transformPath(path: ShapePath, box: PxRect, xfrm: Xfrm): ShapePath;
  /**
   * 経路の全点（制御点を含む）をシートの範囲（0〜全幅・0〜全高）に丸める。
   * 細工した極端な調整値で、座標・`<svg>` の大きさ・表の範囲が巨大にならないための安全弁（ADR-0027 §決定 4）。
   * 範囲の中の点は変えない（Excel の U 字線の張り出しは縮まない）。
   */
  export function clampPath(path: ShapePath, sheet: SheetBounds): ShapePath;
  /** 経路の全点（始点・制御点・終点）を並べる */
  export function pathPoints(path: ShapePath): Pt[];
  /** 線の太さ（px）：EMU / 9525 を 1〜16 に丸める。0・未指定・有限でない値は 1px */
  export function strokeWidthPx(widthEmu: number | undefined): number;
  /**
   * 線種を `stroke-dasharray` の数値にする。`solid` は空（実線）、`dot`・`sysDot` は点線 [w, 2w]、
   * それ以外は破線 [4w, 3w]（w は線の太さ px）。
   */
  export function dashArray(dash: LineDash | string | undefined, strokeWidth: number): number[];
  /** 矢じりの大きさ（px）：max(線の太さ, 2px) × 係数（`sm` 2・`med` 3・`lg` 5。未知は `med`）。上限 48px */
  export function arrowDimension(strokeWidth: number, size: string | undefined): number;
  /**
   * 端の接線の向き（単位ベクトル）：経路の外へ向かう向き。
   * `tail`＝終点で、一つ前の（終点と違う）点から終点への向き。`head`＝始点で、次の（始点と違う）点から始点への向き。
   * 直線・折れ線は端の線分、ベジェは端に隣る制御点→端点。長さ 0 なら更に遡る。全点が同じなら (1, 0)。
   */
  export function endDirection(path: ShapePath, end: 'head' | 'tail'): Pt;
  /**
   * 矢じり 1 つの形を、端点 `tip` と外向きの単位ベクトル `dir` から作る（座標はシートの原点基準）。
   * - `triangle`（と未知の値）：塗った三角。先端が端点。
   * - `stealth`：後ろが切れ込んだ塗った四角形。先端が端点。
   * - `arrow`：開いた V（線だけ）。先端が端点。
   * - `diamond`・`oval`：塗ったひし形・楕円。中心が端点。
   *
   * @param strokeWidth 線の太さ（px）。大きさ＝max(太さ, 2px)×係数（上限 48px）
   */
  export function arrowShape(end: 'head' | 'tail', spec: LineEnd, tip: Pt, dir: Pt, strokeWidth: number): ArrowShape;
  /**
   * 線の端を縮めた経路を返す（塗った三角・stealth の下に線が突き出ないため）。
   * 始点側は始点と最初の区間の制御点、終点側は終点と最後の区間の制御点を、端の接線に沿って動かす。
   * 縮める長さは、その端の区間の長さ（弦）を超えない。区間が 1 つで両端を縮めるときは、合計が区間の長さに収まるよう按分する。
   *
   * @param headShrink 始点側を縮める長さ（px）
   * @param tailShrink 終点側を縮める長さ（px）
   */
  export function shrinkPath(path: ShapePath, headShrink: number, tailShrink: number): ShapePath;
  /**
   * 経路を `<path d>` の文字列にする（数値だけから組み立てる。NaN・Infinity は出さない）。閉じた経路は末尾に `Z`。
   *
   * @param originX 座標の原点の x（この値を引く。外接矩形の左上を原点にするときに渡す）
   * @param originY 座標の原点の y
   */
  export function pathData(path: ShapePath, originX?: number, originY?: number): string;
  /** 部分経路の列を 1 つの `<path d>` にする（図形の path 1 つ。部分経路を空白でつなぐ。空なら空文字） */
  export function subpathsData(paths: readonly ShapePath[], originX?: number, originY?: number): string;
  /** 多角形の矢じりを `<path d>` の文字列にする（閉じる図形は `Z`）。座標はそのまま（原点の移動は `renderLine` が済ませる） */
  export function arrowPolygonData(a: Extract<ArrowShape, {
      kind: 'polygon';
  }>): string;
  /**
   * 1 本の線を、SVG に描ける数値にまとめる。
   * アンカーの矩形 → 回転前の矩形 → 経路 → 反転 → 回転 → シートの範囲へ丸め → 矢じり・縮め・外接矩形。
   * 外接矩形は、丸めた経路の点（制御点を含む）と矢じりの範囲の外接矩形に、線の太さの半分と矢じりの大きさの半分
   * （大きい方）を四方に足し、シートの範囲に丸めたもの。水平・垂直の線（幅か高さが 0）でも太さ分の大きさを持つ。
   *
   * @param rect アンカーの矩形（px。回転後の外接。幅か高さが 0 でもよい）。`preRotated` が true なら回転前の矩形そのもの
   * @param line 線の形と見た目
   * @param xfrm 回転・反転
   * @param sheet シートの範囲
   * @param preRotated `rect` が回転前の矩形か（グループの中の線。`xfrm` の矩形そのもので、縦横を入れ替えない。ISSUE-89）
   */
  export function renderLine(rect: PxRect, line: DrawingLine, xfrm: Xfrm, sheet: SheetBounds, preRotated?: boolean): LineRender;
  /** 面の濃淡の種類（path の `fill` が `darken`・`darkenLess`・`lighten`・`lightenLess`。`norm`・`none` は膜なし） */
  export type ShapeShade = 'darken' | 'darkenLess' | 'lighten' | 'lightenLess';
  /** 図形の path 1 つ分の描画結果（`<path>` 1 つ） */
  export interface ShapeFace {
      /** `<path d>`：外接矩形（`ShapeDrawn.rect`）の左上を原点とした座標。部分経路は空白でつなぎ、閉じた部分経路は `Z` */
      d: string;
      /** 塗る（Excel で塗りがあり、path の `fill` が `none` でない） */
      fill: boolean;
      /** 線を引く（Excel で枠線が見え、path の `stroke` が false でない） */
      stroke: boolean;
      /** 面の濃淡（塗るときだけ。無ければ null） */
      shade: ShapeShade | null;
      /** 部分経路（シートの原点基準・丸め済み）。テスト・後続の判定用 */
      subpaths: ShapePath[];
  }
  /** ダーク表示の下地の、縁取りの片側の幅（px。ADR-0028 §決定 6.7）。下地の svg はこの分だけ広げ、縁取りの太さは線の太さ＋この 2 倍 */
  export const BACKDROP_PAD = 3;
  /** 下地（ダーク表示の白）の path 1 つ：`fill`＝形の塗りとして白く塗る、`stroke`＝線の縁取り（線の太さ＋片側 3px の白い線） */
  export interface ShapeBackdropPath {
      /** `<path d>`：面と同じ座標（外接矩形の左上が原点） */
      d: string;
      fill: boolean;
      stroke: boolean;
  }
  /** 形で描く図形の結果 */
  export interface ShapeDrawn {
      kind: 'drawn';
      /** 外接矩形（シートの原点基準・px・整数）：経路の外接矩形（線の太さの半分を足し、丸めた点から出す）とアンカーの矩形を合わせたもの。可視判定・`drawingExtent`・DOM の位置に使う */
      rect: PxRect;
      /** 回転前の矩形（シートの原点基準。文字の枠を置く。回転は DOM が中心で掛ける） */
      box: PxRect;
      /** 文字領域（回転前の矩形の左上が原点。反転済み。文字そのものは反転しない） */
      textRect: PxRect;
      /** 描く path（何も塗らず線も引かない path は含めない。文字だけの図形は空） */
      faces: ShapeFace[];
      /**
       * ダーク表示の下地の path（ISSUE-96。ADR-0028 §決定 6.7）。塗りのある図形・塗りが無く文字のある図形は形の塗りの path（`fill` が `none` でない path）、
       * 枠線は線の縁取り。塗りも文字も無い図形は枠線の縁取りだけ。塗りの path が無く文字がある形は、回転前の矩形を回した形
       */
      backdrop: ShapeBackdropPath[];
      /** 枠線の太さ（px。1〜16） */
      strokeWidth: number;
      /** 枠線の線種（`stroke-dasharray` の数値。空なら実線） */
      dash: number[];
  }
  /**
   * 図形 1 つの描き方。
   * - `drawn`：形の経路で描く（`ShapeDrawn`）。
   * - `hidden`：枠も塗りも文字も無く、Excel でも見えない図形。描かない（図の数には入れる）。
   * - `fallback`：形を評価できない（`geom` が無い・表に無い `prst`・評価の失敗・上限超え）。ADR-0027 の描き方（矩形の文字枠・種類名の破線枠）にする。
   */
  export type ShapeRender = ShapeDrawn | {
      kind: 'hidden';
  } | {
      kind: 'fallback';
  };
  /** 図形が表示する文字を持つか（ランの文字が 1 文字以上ある。改行だけ・空のランは含めない） */
  export function shapeHasText(drawing: SheetDrawing): boolean;
  /** 塗りがあり、その色が見える（不透明度 0 超。色が無い＝暫定の色は見える）。下地の塗りの条件（ISSUE-96） */
  export function filledShown(drawing: SheetDrawing): boolean;
  /** 見える文字がある（文字があり、不透明度 0 超のランが 1 つ以上ある。`a:noFill` の文字は見えない）。下地の塗りの条件（ISSUE-96） */
  export function textShown(drawing: SheetDrawing): boolean;
  /**
   * 図形 1 つを、SVG に描ける数値にまとめる（ADR-0028 §決定 2・4・5・8）。
   * アンカーの矩形 → 回転前の矩形 → 形の経路（`evaluateGeom`）→ 反転 → 中心で回転 → シートの範囲に丸め → 外接矩形。
   * 外接矩形は、丸めた経路の点（制御点を含む）の外接矩形に、枠線があれば太さの半分を足して整数に広げ、アンカーの矩形と合わせたもの。
   * 枠も塗りも無く文字だけの図形は経路を作らず、外接矩形はアンカーの矩形。
   * 文字領域は形の `rect`（回転前の矩形の中）を、反転に合わせて矩形の中で鏡写しにしたもの（`flipH` なら x を w − r へ）。
   *
   * @param rect アンカーの矩形（px・整数。回転後の外接）。`preRotated` が true なら回転前の矩形そのもの
   * @param drawing 図形（`geom`・`outline`・`hasOutline`・`hasFill`・`rot`・`flipH`・`flipV`・`text`）
   * @param sheet シートの範囲
   * @param preRotated `rect` が回転前の矩形か（グループの中の図形。縦横を入れ替えない。ISSUE-89）
   */
  export function renderShape(rect: PxRect, drawing: SheetDrawing, sheet: SheetBounds, preRotated?: boolean): ShapeRender;
}

declare module 'xsheet/_/geom/index' {
  /**
   * geom の入口（`xsheet/geom`）。図形の形の表と式の評価。他の層に依存しない。
   * 形の表（`preset-geom.json`・約 171 KB）を読むので、render からは動的 import する（M2-3）。
   * core は表を読まない `defs.ts` だけを引く。
   */
  export { GUIDE_ARITY, MAX_GEOM_COMMANDS, MAX_GEOM_GUIDES, type PathSeg, type Pt } from 'xsheet/_/geom/defs';
  export { evaluateGeom, evaluateGuides, getPresetGeom, presetNames, type GeomArg, type GeomCmd, type GeomDef, type GeomFailure, type GeomFill, type GeomGuide, type GeomPath, type GeomPathDef, type GeomResult, type GeomSubpath, type GeomTextRect, type GuidesResult, } from 'xsheet/_/geom/evaluate';
}

declare module 'xsheet/_/geom/evaluate' {
  /**
   * `.xlsx` の図形の「形」の定義（DrawingML の guide の式と path 命令）を評価する純ロジック（DOM 非依存・InterHub ISSUE-90）。
   * 流れ：形の定義（表の preset か `custGeom`）＋調整値＋回転前の矩形の幅・高さ（px）
   * → 組み込み変数 → `avLst`（既定を調整値で上書き）→ `gdLst` を上から 1 回ずつ → path 命令を経路（直線・3 次ベジェ）へ。
   * 式は自前の解釈器で評価し、`eval`・`new Function` は使わない（docs/01_requirements.md N1）。
   * 名前（guide・調整値・`prst`）は `Map` のキーにだけ使い、普通のオブジェクトのキーにしない。
   * 反転・回転・シート範囲への丸め・`d` の組み立ては描画側（ISSUE-92・`xlsx-shapes.ts` の写像）が行う。ここは回転前の矩形の中の経路まで。
   * 形の表は生成物 `preset-geom.json`（変換器 `tools/geom/build-preset-geom.mjs`・出典はリポジトリ直下の NOTICE）。
   * コメント中の ISSUE-N・ADR-N は InterHub の番号。
   */
  import { MAX_GEOM_COMMANDS, MAX_GEOM_GUIDES, type PathSeg, type Pt } from 'xsheet/_/geom/defs';
  export { MAX_GEOM_COMMANDS, MAX_GEOM_GUIDES };
  /** path の塗りの濃淡の種類（`none`＝塗らない） */
  export type GeomFill = 'none' | 'norm' | 'lighten' | 'lightenLess' | 'darken' | 'darkenLess';
  /** 式・命令の引数：整数、または組み込み変数・guide の名前 */
  export type GeomArg = number | string;
  /** guide 1 つ：`[名前, 演算子, 引数…]`（例 `['x1', '+-', 'w', 'adj1', 0]`） */
  export type GeomGuide = readonly [name: string, op: string, ...args: GeomArg[]];
  /** path 命令 1 つ：`['M', x, y]`・`['L', x, y]`・`['A', wR, hR, stAng, swAng]`・`['Q', x1, y1, x, y]`・`['C', x1, y1, x2, y2, x, y]`・`['Z']` */
  export type GeomCmd = readonly [op: string, ...args: GeomArg[]];
  /** path 1 つの定義 */
  export interface GeomPathDef {
      /** 座標系の幅・高さ（無い・0 なら形の座標系＝EMU のまま） */
      w?: number;
      h?: number;
      /** 塗りの濃淡の種類（無ければ `norm`） */
      fill?: GeomFill;
      /** 線を引くか（無ければ true） */
      stroke?: boolean;
      /** 命令の列 */
      c: readonly GeomCmd[];
  }
  /** 形の定義（表の preset も `custGeom` もこの形に揃える） */
  export interface GeomDef {
      /** 調整値の既定（`avLst`）：`[名前, 値]` */
      av?: ReadonlyArray<readonly [string, number]>;
      /** guide の列（`gdLst`）。名前は「それより前」に定義したものだけ引ける */
      gd?: readonly GeomGuide[];
      /** 文字領域（l・t・r・b の引数。無ければ矩形全体） */
      rect?: readonly GeomArg[];
      /** path の列 */
      paths: readonly GeomPathDef[];
      /**
       * 評価した経路・文字領域を矩形の中で上下反転する（y → h − y）。原本に定義の無い `upArrow` を `downArrow` から作るために使う（`getPresetGeom`）。
       * ファイルの `custGeom` は持たない
       */
      flipY?: boolean;
  }
  /** 評価の失敗の理由（代替の描き方は ISSUE-92）。`name`＝未知の名前・前方参照、`op`＝未知の演算子、`args`＝引数の数の誤り、`cmd`＝未知の path 命令、`limit`＝上限超え */
  export type GeomFailure = 'name' | 'op' | 'args' | 'cmd' | 'limit';
  /** 部分経路：始点・区間の列・閉じるか（区間は `xlsx-shapes.ts` の直線 `L`・3 次ベジェ `C` と同じ表し方） */
  export interface GeomSubpath {
      start: Pt;
      segs: PathSeg[];
      closed: boolean;
  }
  /** path 1 つの経路（px・回転前の矩形の左上が原点・y は下向き） */
  export interface GeomPath {
      fill: GeomFill;
      stroke: boolean;
      subpaths: GeomSubpath[];
  }
  /** 文字領域（px・回転前の矩形の左上が原点） */
  export interface GeomTextRect {
      l: number;
      t: number;
      r: number;
      b: number;
  }
  /** 評価の結果：成功なら経路と文字領域、失敗なら理由（例外にはしない） */
  export type GeomResult = {
      ok: true;
      paths: GeomPath[];
      textRect: GeomTextRect;
  } | {
      ok: false;
      reason: GeomFailure;
  };
  /** guide の評価の結果（値は EMU） */
  export type GuidesResult = {
      ok: true;
      values: Map<string, number>;
  } | {
      ok: false;
      reason: GeomFailure;
  };
  /**
   * guide を評価して名前 → 値（EMU）の表にする。
   * 順序は 組み込み変数 → `av`（`adjust` の同じ名前の値で上書き。`av` に無い名前は無視）→ `gd` を上から 1 回ずつ。
   * 前方参照・未知の名前・未知の演算子・引数の数の誤り・上限超えは例外にせず失敗として返す。
   *
   * @param def 形の定義
   * @param adjust ファイルの調整値（名前 → 値）。無ければ既定のまま
   * @param w 回転前の矩形の幅（px）
   * @param h 回転前の矩形の高さ（px）
   */
  export function evaluateGuides(def: GeomDef, adjust: ReadonlyMap<string, number> | undefined, w: number, h: number): GuidesResult;
  /**
   * 形の定義を評価して、回転前の矩形の中の経路と文字領域を返す（座標は px・矩形の左上が原点）。
   * 失敗（未知の名前・演算子・命令、引数の数の誤り、上限超え）は例外にせず `{ ok: false, reason }` で返す。
   * 経路は直線と 3 次ベジェだけ（`quadBezTo` は誤差なしで変換・`arcTo` は 90° 以下に分けて近似。誤差は半径の 0.03% 未満）。
   * 文字領域は `rect` の l・t・r・b を評価した値。`rect` が無い・幅か高さが 0 以下なら矩形全体。
   * 反転した `rect`（r ≦ l など）は、調整値の極端な値で式が裏返ったものとみなして矩形全体へ戻す
   * （幅 0 以下の領域には文字が並ばず、図形ごと文字が消えるより、図形の矩形に並べる方が読める）。
   *
   * @param def 形の定義（表の preset か `custGeom`）
   * @param adjust ファイルの調整値（名前 → 値）
   * @param w 回転前の矩形の幅（px）
   * @param h 回転前の矩形の高さ（px）
   */
  export function evaluateGeom(def: GeomDef, adjust: ReadonlyMap<string, number> | undefined, w: number, h: number): GeomResult;
  /** preset 名（`prst`）から形の定義を引く。表に無い（線の 10 種・未知の名前）なら undefined。`upArrow` は原本に無いので `downArrow` の上下反転（`flipY`）で返す */
  export function getPresetGeom(prst: string): GeomDef | undefined;
  /** 表にある preset 名の一覧（原本の出現順。別名の `upArrow` は含めない） */
  export function presetNames(): string[];
}

declare module 'xsheet/_/render/format-css' {
  /**
   * `.xlsx` のセルの書式（塗り・文字の装飾）の CSS を作る純関数（ISSUE-100・DOM 非依存）。
   * - 書式ごとに規則 1 行：`.xs-sheet[data-xs-fmt="<連番>"] td.xs-xf-<番号> { … }`。
   *   連番の属性で絞るため、既存の `.xs-grid td`（グリッド線）より詳細度が強く、同じ文書の別の表示と混ざらない。
   * - 文字列は固定の語・整数・整数から組んだ `#rrggbb`・小数（文字サイズの比）・固定のテーマ変数の参照
   *   （白の塗りの `var(--xs-cell-bg)` だけ）で組む。ファイルの文字列は入れない。
   * - ダーク表示の描き分けは規則に焼き込む（テーマ変数は増やさず、切替で描き直さない）：
   *   白以外の塗り＝Excel の色の面と黒の文字（明示色はそのまま）、塗りなし・白の塗り（紙）＝画面の面と文字色（明示色だけ書く）。
   */
  import { type CellFormat } from 'xsheet/_/core/index';
  /**
   * 書式の番号 → td に足すクラス名（`xs-xf-<番号>`）。
   * @param index `XlsxSheet.formats` の位置
   */
  export function formatClassName(index: number): string;
  /**
   * 「紙でない」塗りがあるか。はみ出す文字の下の面（グリッド線隠し）を敷くかどうかの判定に使う（ADR-0030 §決定 9）。
   * 白の塗り・塗りなし・書式なし（undefined）は false。
   */
  export function hasOverFill(format: CellFormat | undefined): boolean;
  /**
   * はみ出した「自動」の文字色を、列ごとの区間で塗り分けるグラデーション（`linear-gradient`）を作る（ADR-0030 §決定 9）。
   * 塗りの上の自動色は黒、塗りなし・紙の上は画面の文字色で、はみ出した文字が塗りの状態の違う列をまたぐと、
   * 単色ではダーク表示でどちらかの面の上に読めない色が残る（黄塗りの黒文字が暗い面へ・明るい文字が薄青の面へ）ため、列の面に合わせる。
   * - 呼び出し側は `background-image` にこれを入れ、`background-clip: text; color: transparent` を足す（span の左端＝元の左端）。
   * - 元の文字が明示色（`CellFormat.color`。黒は自動に寄せ済み）なら null（Excel と同じく単色）。
   * - 区間は元の列から右へ。隣り合う同じ色は 1 区間にまとめ、幅 0（非表示の列）は数えない。区間が 1 つなら null（塗り分け不要）。
   * - 文字列は固定の語・整数の px・`#000000`・`var(--xs-fg)` だけ（ファイルの文字列は入らない）。
   *
   * @param table 書式の表（`XlsxSheet.formats`）
   * @param formats 行の書式の番号（`GridRow.formats`。描く列に揃えたもの。無ければ全列「塗りなし」）
   * @param c はみ出す元の列（0 始まり。`formats` の添字）
   * @param widths 元の列から右へ数えた列幅（px）。`spillColumns` の列数ぶん。不正な値は 0 とみなす
   * @returns グラデーションの文字列。塗り分けが要らなければ null
   */
  export function spillInkGradient(table: readonly CellFormat[], formats: readonly (number | undefined)[] | undefined, c: number, widths: readonly number[]): string | null;
  /**
   * 書式の表から CSS の文字列を作る（ブックを開いた時に 1 回）。
   * 書式の番号は `formats` の位置のまま（宣言が 1 つも出ない書式は規則を作らないが、後ろの番号はずれない）。
   *
   * @param formats `XlsxSheet.formats`（ブックで共有する書式の表）
   * @param seq 表示ごとの連番（`.xs-sheet` の `data-xs-fmt`）。整数でなければ 0
   * @returns 規則を改行でつないだ文字列。書式が無ければ空の文字列
   */
  export function buildFormatCss(formats: readonly CellFormat[], seq: number): string;
}

declare module 'xsheet/_/render/grid-table' {
  /**
   * `.xlsx` グリッドの table を組む（InterHub `spreadsheet.ts` の `buildGridTable` とその部品・M2-3）。
   * 寸法・範囲の計算は DOM 非依存の `grid.ts`、書式の CSS は `format-css.ts` にあり、ここは DOM を組むだけ。
   * ユーザーのファイル内容（セル値）は信頼できない入力として扱い、`innerHTML` は使わず
   * `document.createElement` + `textContent`（元の値は `title`）だけで入れる。
   */
  import type { CellFormat, Geometry } from 'xsheet/_/core/index';
  import { type GridHits } from 'xsheet/_/render/highlight';
  import { type ColRange, type GridRow } from 'xsheet/_/render/grid';
  /**
   * グリッドの table を Excel の寸法（geometry）で固定して組む（計測テスト `format.perf.test.ts` のため export）：
   * `colgroup`（列幅）・見出し行（列記号）・上の余白・行（Excel の行番号＋セル）・下の余白。
   * - 列幅は `<col>` の `width`、行の高さは `<tr>` の `height` とセルの内側の要素の高さで固定する（`table-layout: fixed`）。
   * - 非表示の列は幅 0 とし、`<col>` に `xs-collapsed`（`visibility: collapse`）を付けて畳む（ADR-0026 §決定 5）。
   *   非表示の行は `rows` に含まれない（呼び出し側で除く）。結合セルの `rowSpan` は描く行の数。
   * - セル値は `textContent` のみで入れる。結合で隠れるセル（null）は td を作らない。td は 1 セル 1 個（罫線を残す）。
   *   文字列のはみ出しは、内側の要素を右の空セルの列幅の合計まで広げ、文字の範囲だけ背景色で下の罫線を覆う。
   *   セルの幅（結合を含む）が 0 なら内側の要素を作らない（畳んだ列の文字を出さない）。
   * - 配置は型ごとの既定（数値・日付は右、真偽・エラーは中央、文字列は左）。収まらない数値・日付は `###` にし、元の値を title に入れる
   *   （大きな文字のセルは推定幅に文字サイズの比を掛ける）。
   * - 書式（塗り・文字の装飾）のあるセルの td には、className に `xs-xf-<番号>` を 1 語足すだけ（規則は表示の根の `<style>`。style 属性は書かない）。
   *   塗り（白以外）が絡むはみ出しの入れ物には `is-over-fill` を付ける（ADR-0030 §決定 9）。
   * - 列も仮想化する：描く列は `range`（可視範囲の左右 1 画面分）だけ。描かない左右の列は、幅 `colLeft` の差の
   *   余白の列（`<col>` 1 本＋各行に空の td 1 個）で置き換え、表の全幅は geometry の totalWidth のまま保つ。
   *   範囲の端で切れる結合セルは `clipRowToColumns` が `colSpan` を縮める。
   *
   * @param colName 列番号（1 始まり）→ 列記号
   * @param cols 表の列数（データの列数より広いことがある）
   * @param rows 描く行（見える行だけ。行番号の昇順）
   * @param geo 寸法の換算結果（`createGeometry`）
   * @param rowCount 表の行数（データの行数より広いことがある）
   * @param mdw 最大桁幅 MDW（`SheetLayout.mdw`。収まらない数値を `###` にする推定幅に使う）
   * @param range 描く列の範囲（`first`〜`last`・1 始まり）。表の全列を描くときは 1〜`cols`
   * @param formats 書式の表（`XlsxSheet.formats`。`rows[i].formats` の値が指す）
   * @param hits 検索の一致のハイライト（XSheet: ADR-0007 §2。無ければ null）
   * @param headers 行番号の列と列見出しを出すか（XSheet: ADR-0007 §7。false なら列見出しの行を作らず、行番号の列を幅 0 で畳む）
   */
  export function buildGridTable(colName: (c: number) => string, cols: number, rows: GridRow[], geo: Geometry, rowCount: number, mdw: number, range: Pick<ColRange, 'first' | 'last'>, formats: readonly CellFormat[], hits?: GridHits | null, headers?: boolean): HTMLTableElement;
}

declare module 'xsheet/_/render/highlight' {
  /**
   * 検索の一致のハイライト（M7-5・Moobius の依頼 R2）。
   * - 一致の一覧（`findText` の `matches`）を「行 → 列 → 範囲」の 2 段の表にして表示が持ち、
   *   表を描くたびに描く行・列の一致だけを付け直す（仮想スクロールで描き直しても残る）。
   * - セルの文字を一致の範囲で分け、**分けた全体を `span.xs-hits` 1 つで包む**。`.xs-cell`・`.xs-text` は flex なので、
   *   包まずに分けると各片が別の flex 項目になり、片の端の空白が消える。
   * - 一致の片は `span.xs-hit`（`data-xs-hit` に一覧の番号）。今のヒットは `is-current`。色は CSS 変数
   *   `--xs-hit-bg`・`--xs-hit-current-bg` で、背景だけを付ける（行の高さ・列の幅を変えない）。
   * - 自動の色の文字が塗りをまたぐはみ出し（`background-clip: text` で描く）は、片に `is-ink`（`--xs-fg` の文字の色）を付ける。
   *   不透明なハイライトが、親の背景で描く文字を隠さないようにするため。
   * - `###` で出すセル（数値が列幅に収まらない）は、表示の文字と一致の位置が合わないので、セル全体を光らせる。
   * 文字は `textContent`（テキストノード）だけで入れる（セル値は信頼できない入力）。
   */
  import type { TextMatch } from 'xsheet/_/core/index';
  /** 1 セルの中の一致 */
  export interface CellHit {
      /** `CellView.text` 上の位置（UTF-16） */
      start: number;
      /** 終わり（含まない） */
      end: number;
      /** 一致の一覧（`setHighlights` に渡した配列）の番号 */
      index: number;
  }
  /** 表を描くときに引くハイライト */
  export interface GridHits {
      /** 今のヒットの番号（無ければ -1）。`focusHighlight` が替える（M7-6） */
      current: number;
      /** 一致の一覧の数（`setHighlights` に渡した配列の長さ） */
      readonly count: number;
      /** 番号 index の一致（番号の正しくない一致・範囲外は undefined） */
      match(index: number): TextMatch | undefined;
      /** 行 r・列 c（1 始まり）の一致（無ければ undefined） */
      at(r: number, c: number): readonly CellHit[] | undefined;
  }
  /**
   * 一致の一覧から、表を描くときに引く表を作る。null・空の配列は null（ハイライトなし）。
   * 番号の正しくない一致（整数でない行・列、範囲の逆転）は飛ばす。同じセルの一致は始まりの順に並べる。
   */
  export function createGridHits(matches: readonly TextMatch[] | null | undefined, current?: number): GridHits | null;
  /**
   * 描いた表の中の片の `is-current` を、今のヒットに合わせて付け替える（表を描き直さない。ADR-0007 §3）。
   * @returns 今のヒットの片が描いた表の中にあったか
   */
  export function markCurrent(table: ParentNode, current: number): boolean;
  /** 片の番号（`data-xs-hit`）。片でなければ -1 */
  export function hitIndexOf(el: Element | null): number;
  /**
   * セルの入れ物（`div.xs-cell`）の文字に一致を付ける。`grid-table.ts` が入れ物を作った後に呼ぶ。
   * @param box セルの入れ物（文字は直下のテキストか、はみ出しなら `span.xs-text` の中）
   * @param text セルの文字（`CellView.text`）
   * @param shown 表示した文字（収まらない数値は `###`）
   * @param hits このセルの一致（無ければ何もしない）
   * @param current 今のヒットの番号
   * @param slice 区切りで分けるノードの作り手（リッチテキスト。ADR-0009 §2）。無ければ平文のテキストノード
   */
  export function decorateCell(box: HTMLElement, text: string, shown: string, hits: readonly CellHit[] | undefined, current: number, slice?: (from: number, to: number) => Node[]): void;
}

declare module 'xsheet/_/render/image-layer' {
  /**
   * `.xlsx` グリッドの上に重ねる画像レイヤー（ISSUE-86）。
   * - 構造：`position: absolute` のレイヤー要素（原点＝A1 の左上）の中に、図ごとの要素を `left/top/width/height`（px）で置く。
   *   表とは独立で、スクロールは枠ごと動くため位置の再計算は要らない。可視範囲（上下左右 1 画面分）と交わる図だけを DOM に出し入れする。
   * - 画像は `<img>`（`decoding="async"`・`alt` は説明か名前）。画像以外・描けない画像・壊れた画像は種類名入りの破線枠。
   *   種類名は文言（`messages`）を `textContent` で入れる（図の名前・パスは画面にも console にも出さない。名前は `alt` のみ）。
   * - 図形・テキストボックスは、形を評価できれば形どおりに描く（ISSUE-92）：`div.xs-figure`（外接矩形）＞ `<svg>`（path 1 つにつき `<path>` 1 つ。枠線・塗り・面の濃淡はクラスと CSS）
   *   ＋ `div.xs-shape-box`（回転前の矩形・回転・図形の矩形で切る）＞ `div.xs-shape-text`（文字領域）。枠も塗りも文字も無い図形は描かない（配置の時点で除く）。
   * - 形を評価できない図形（代替）で文字のあるものは種類名ではなく文字を描く（ISSUE-88。段落＝`div`・ラン＝`span`・改行＝`br`、すべて `textContent`）。
   *   枠線・塗りの有無はクラス、回転・縦書きは列挙値の対応表から `style` に入れる（ファイルの文字列は style に入れない）。
   * - 線・矢印（ISSUE-88）は 1 本につき `<svg>` 1 つ（外接矩形の位置）。`<marker>`・`innerHTML` は使わず、矢じりは `<path>`・`<ellipse>`。
   * - 色（ISSUE-96）：lib の整数 RGB・不透明度を検証して `#rrggbb`・数値にし、要素の style（CSSOM）で入れる（図形の fill・stroke、線・矢じり、文字の color、代替の文字枠の背景・枠線）。
   *   面の濃淡は lib の `applyFillMode`。色が無い（省略）ときは style を入れず、暫定のテーマ変数（CSS）のまま。
   * - ダーク表示の下地（ISSUE-96）：図形・線・代替の文字枠は、図の要素の直前に独立した `svg.xs-backdrop`（`z-index: -1`・片側 3px 広げる）を兄弟として置く。
   *   形の塗りの path と線の縁取り（太さ＋6px・実線・丸）だけを持ち、色は CSS のテーマ変数（ライトは透明）で決める＝style に色を入れず、テーマの切り替えで描き直さない。
   *   レイヤーが `isolation: isolate` で重ね順の文脈を作り、下地は全図の下・セルの上に描かれる。図と一緒に出し入れする。
   * - グラフ（ISSUE-103・ADR-0031）：kind が chart で `chartPath` があり `readChart` を渡されたら、図が初めて見える範囲に入った時に読み（{@link createChartStore}。1 つずつ順に）、
   *   `layoutChart` → `buildChartSvg` で `<svg>` にして破線枠の代わりに入れる。`<svg>` は図の位置ごとにシート単位でキャッシュし、範囲から外れたら DOM から外すだけ（作り直さない）。
   *   グラフの面が塗りなし（場面の `backdrop`）なら、ダーク表示の下地も図形と同じ兄弟の `svg.xs-backdrop`（図の直前・図の矩形と同じ位置と大きさ）にする＝グラフより下の図形・画像を隠さない。
   *   下地は `<svg>` と一緒にキャッシュし、図と一緒に出し入れする（`<svg>` の中には置かない）。
   *   読み込み中・失敗（unsupported・too-large・broken・no-data・場面が null・例外）は破線枠「グラフ」のまま（P8）。シートの差し替え・dispose でキャッシュと `<svg>` を捨て、後から返った結果は足さない。
   * - blob URL は {@link createImageStore}（シート 1 枚ぶん）が持つ。シートの差し替え・{@link ImageLayer.dispose} で revoke する。
   * 位置の計算は DOM 非依存の `drawings.ts`、blob の管理は `image-store.ts` に切り出している。
   */
  import type { ChartResult, ImageResult, SheetDrawing } from 'xsheet/_/core/index';
  import { type DrawingBox, type PxRect } from 'xsheet/_/render/drawings';
  import { type ImageEntry } from 'xsheet/_/render/image-store';
  import { type XSheetMessages } from 'xsheet/_/render/messages';
  /** 画像レイヤーに渡すシートの図 */
  export interface LayerSheet {
      /** 図の一覧（drawing XML の順） */
      drawings: readonly SheetDrawing[];
      /** 配置済みの図（`layoutDrawings` の結果。`index` は `drawings` の位置） */
      boxes: readonly DrawingBox[];
      /** 画像を読む（`XlsxBook.readImage` をシート番号込みで束ねたもの） */
      readImage: (path: string) => Promise<ImageResult>;
      /**
       * グラフを読む（`XlsxBook.readChart` をシート番号込みで束ねたもの）。省略すると、グラフは破線枠のまま。
       * 描けたら破線枠の代わりに `<svg>`（ISSUE-103）。どの失敗も破線枠のまま。
       */
      readChart?: (path: string) => Promise<ChartResult>;
      /** 画像の合計の上限を超えたとき（バナー用）。1 シートにつき 1 回 */
      onOverBudget?: () => void;
  }
  /** 画像レイヤー */
  export interface ImageLayer {
      /** スクロール枠の中の wrapper に取り付けるレイヤー要素 */
      readonly element: HTMLElement;
      /** 描くシートを差し替える（前のシートの blob URL はすべて revoke し、DOM の図も外す） */
      setSheet(sheet: LayerSheet | null): void;
      /** 原点（A1 の左上）の位置とレイヤーの大きさ（px。wrapper の左上基準） */
      place(left: number, top: number, width: number, height: number): void;
      /** 可視範囲（A1 基準・余白込み。`viewWindow`）と交わる図だけを DOM に出し、外れた図を外す */
      update(view: PxRect): void;
      /** すべての blob URL を revoke し、DOM の図を外す。以後 `setSheet`・`update` は何もしない */
      dispose(): void;
      /** XSheet: 今までに始めた画像・グラフの読み込みが、図へ反映し終わるまで待つ（印刷の描画が使う。ADR-0010 §2-3） */
      settled(): Promise<void>;
  }
  /**
   * 破線枠に入れる種類名。画像（`kind: 'image'`）は、描けない形式で形式名が分かれば「画像（EMF）」、それ以外は「画像（表示できない）」。
   *
   * @param entry 画像の読み込み失敗（画像以外・読み込み前は不要）
   * @param messages 文言の表（既定は日本語）
   */
  export function frameLabel(drawing: SheetDrawing, entry?: ImageEntry, messages?: Readonly<XSheetMessages>): string;
  /**
   * 文字を描く図形か（kind が shape・textbox で、表示する文字が 1 文字以上ある。改行だけ・空のランは含めない）。
   * 該当しない図形は種類名入りの破線枠になる。
   */
  export function hasShapeText(drawing: SheetDrawing): boolean;
  /**
   * 画像レイヤーを作る。
   *
   * @param deps テスト用に差し替える依存（blob URL の作成・解放）。通常は省略（`URL.createObjectURL`・`revokeObjectURL`）
   */
  export function createImageLayer(deps?: {
      createUrl?: (blob: Blob) => string;
      revokeUrl?: (url: string) => void;
      /** 破線枠の文言（既定は日本語） */
      messages?: Readonly<XSheetMessages>;
  }): ImageLayer;
}

declare module 'xsheet/_/render/image-store' {
  /**
   * `.xlsx` の画像の blob URL を、シート 1 枚ぶん管理する（DOM 非依存・ISSUE-86）。
   * - media のパス → 読み込み結果の `Map`。同じ画像を複数のアンカーで共有し、`readImage` は 1 パスにつき 1 回だけ呼ぶ。
   * - 判定済みの MIME を明示した Blob から URL を作る（`<img>` の `src` にだけ使う前提。拡張子・Content_Types は信じない）。
   * - 読み込み（展開）は 1 本ずつ待ち行列で行い、`read` を呼ぶ直前に破棄済みかを確かめる。`dispose()` で未開始の読み込みは捨てる
   *   （シートを切り替えた後も前のシートの画像の展開が走り続けて、新しいシートの画像が待たされるのを防ぐ）。
   * - `dispose()`（タブの切替・ビューの破棄）ですべて revoke する。破棄後に読み込みが終わっても URL は作らない
   *   （読み込み中にタブやファイルが切り替わった場合の古い結果の抑止。`spreadsheet.ts` の `openSeq` と同じ目的を、シート単位の破棄で担う）。
   * 画像のバイト・パスは console に出さない。
   */
  import type { ImageFailReason, ImageResult } from 'xsheet/_/core/index';
  /** 読み込み結果。成功は blob URL、失敗は理由（表示側が破線枠にする）。`label` は EMF など形式名（分かれば） */
  export type ImageEntry = {
      ok: true;
      url: string;
  } | {
      ok: false;
      reason: ImageFailReason;
      label: string | null;
  };
  /** 画像ストアの外部依存（テストで差し替える） */
  export interface ImageStoreDeps {
      /** 画像を読む（`XlsxBook.readImage` をシート番号込みで束ねたもの） */
      read: (path: string) => Promise<ImageResult>;
      /** Blob から URL を作る（既定 `URL.createObjectURL`） */
      createUrl?: (blob: Blob) => string;
      /** URL を解放する（既定 `URL.revokeObjectURL`） */
      revokeUrl?: (url: string) => void;
      /** 1 シートの画像の合計の上限を超えた（`over-budget`）ことを初めて知ったとき 1 回だけ呼ぶ（バナー用） */
      onOverBudget?: () => void;
  }
  /** シート 1 枚の画像ストア */
  export interface ImageStore {
      /** 読み込み済みの結果（未読み込み・読み込み中は undefined） */
      peek(path: string): ImageEntry | undefined;
      /**
       * 画像を読む。読み込み済み・読み込み中のパスは共有する。破棄済み（または読み込み中に破棄された）なら null。
       * 例外にはしない（読めなければ `reason: 'broken'` の失敗）。
       */
      load(path: string): Promise<ImageEntry | null>;
      /** 作った URL をすべて revoke し、以後の読み込み結果を捨てる */
      dispose(): void;
      /** 保持している blob URL の数（テスト用） */
      readonly urlCount: number;
  }
  /** 画像ストアを作る */
  export function createImageStore(deps: ImageStoreDeps): ImageStore;
}

declare module 'xsheet/_/render/spill-extent' {
  /**
   * 最後の列からはみ出す文字の分だけ、表の列数を広げる（M7-11・Moobius の依頼 R8）。
   * - 行ごとに「値のある最後のセル」が文字列ではみ出せるときだけ、`列の左端 + 文字の幅の見積もり` を含む列を数える。
   * - 文字の幅は DOM で測らず、文字数から多めに見積もる（多めに出た分は空の列が増えるだけで、文字は切れない）。
   * - 調べるのは描いた行と移動先の行だけ（シートを開くときに全行を読まない）。広げた列数は表示（`sheet-view.ts`）が持つ。
   * DOM 非依存の純関数。
   */
  import type { CellFormat, Geometry, TextRun } from 'xsheet/_/core/index';
  import type { GridRow } from 'xsheet/_/render/grid';
  /**
   * 文字の幅の見積もり（px・倍率 1）。多めに出す。
   * @param text 表示する文字（1 行。改行は空白として数える）
   * @param format セルの書式（文字の大きさの比・太字・斜体）
   * @param runs リッチテキストの区切り（`<rPr>` のある区切りはその書式で数える。ADR-0009 §3）
   */
  export function estimateCellTextWidth(text: string, format?: CellFormat, runs?: readonly TextRun[]): number;
  /**
   * 行の文字がはみ出して届く列（1 始まり）の最大。はみ出す文字の無い行だけなら 0。
   * 行の「値のある最後のセル」だけを見る（それより左のセルは、右に値があるので最後の列を越えない）。
   * 右寄せ・折り返しのセルは右へはみ出さない。中央はセルの中心から文字の幅の半分だけ右へ届く（ADR-0008 §3）。
   *
   * @param geo 倍率 1 の寸法（表の範囲より右の列も既定の幅で延びるもの）
   * @param rows 調べる行（`GridRow`。`cells[i]` は列 i + 1）
   * @param formats 書式の表（`XlsxSheet.formats`）
   */
  export function spillReachCols(geo: Pick<Geometry, 'colLeft' | 'colAt' | 'colWidth'>, rows: readonly Pick<GridRow, 'cells' | 'formats'>[], formats: readonly CellFormat[]): number;
}

declare module 'xsheet/_/render/zoom' {
  /**
   * 拡大縮小（M7-7・Moobius の依頼 R5）。**寸法に倍率を掛けて描き直す**（CSS の `zoom`・`transform` で表を縮めない）。
   * - `Geometry` を倍率で包む。境目を整数に丸める（`top′(r) = round(top(r) × z)`・高さは境目の差）ので、
   *   罫線の層（SVG）と td の線がずれず、誤差もたまらない。線の幅（td の 1px の枠・罫線）には倍率を掛けない。
   * - 見出し（列見出しの高さ・行番号の列の幅・見出しの文字）には倍率を掛けない（`sheet-view.ts` が原点を倍率なしで渡す）。
   * - 文字の大きさは CSS 変数 `--xs-zoom`（`style.css`）。図の層は中を 100% で描き、層に `transform: scale()` を掛ける。
   * - 幅に合わせる（fit）は、データのある列と図の右端までを表示枠の幅に合わせる。
   */
  import type { Geometry } from 'xsheet/_/core/index';
  import type { DrawingBox, PxRect } from 'xsheet/_/render/drawings';
  /** 倍率。数は 0.5〜2.0、`{ fit: 'width' }` はデータのある列と図の右端までを表示枠の幅に合わせる */
  export type Zoom = number | {
      fit: 'width';
      min?: number;
      max?: number;
  };
  /** 倍率の下限・上限 */
  export const ZOOM_MIN = 0.5;
  export const ZOOM_MAX = 2;
  /** 解決した倍率の指定（数は範囲に丸めた値。fit は範囲に丸めた下限・上限） */
  export type ZoomSpec = {
      fixed: number;
  } | {
      fit: {
          min: number;
          max: number;
      };
  };
  /**
   * 倍率の指定を確かめて解決する。数は 0.5〜2.0 に丸める。fit の `min`・`max` は数でなければ既定（0.5・2.0）、
   * 逆転していれば入れ替える。
   * @throws RangeError 数でない値（`NaN`・`Infinity` を含む）・`{ fit: 'width' }` でないもの
   */
  export function resolveZoom(zoom: Zoom): ZoomSpec;
  /**
   * 寸法を倍率で包む。境目（行の上端・列の左端）を整数に丸め、高さ・幅は境目の差にする。
   * `rowAt`・`colAt` は倍率で割って元の寸法で引いた後、前後の行（列）を丸めた境目と比べて合わせる。
   * 倍率 1 なら元の寸法をそのまま返す（今までと同じ道を通る）。
   */
  export function zoomGeometry(g: Geometry, z: number): Geometry;
  /** 図の出し入れの判定の範囲（倍率を掛けた座標）を、図の層の中の座標（100%）に戻す */
  export function unzoomRect(rect: PxRect, z: number): PxRect;
  /**
   * 幅に合わせるときの中身の幅（100% の px）：データのある列の右端と、図の右端の大きいほう。
   * @param dataWidth データのある列（1〜`colCount`）の幅の合計
   */
  export function contentWidth(dataWidth: number, boxes: readonly DrawingBox[]): number;
  /**
   * 幅に合わせる倍率。合わせ直さなくてよい（中身・枠の幅が無い、今の倍率との差が 1% 未満）ときは null。
   * @param available セルに使える表示枠の幅（px。行番号の列とスクロールバーを除く）
   * @param content 中身の幅（100% の px）
   * @param current 今の倍率
   */
  export function fitZoom(available: number, content: number, range: {
      min: number;
      max: number;
  }, current: number): number | null;
  /** 表示枠の左上にあるセルと、そのセルの中の位置の割合（倍率を替えても同じセルを左上に残すため） */
  export interface ScrollAnchor {
      r: number;
      fy: number;
      c: number;
      fx: number;
  }
  /** 今のスクロール位置（A1 の左上からの px）の左上のセル */
  export function scrollAnchor(g: Geometry, scrollTop: number, scrollLeft: number): ScrollAnchor;
  /** 左上のセルを、新しい寸法でのスクロール位置にする */
  export function anchorScroll(g: Geometry, a: ScrollAnchor): {
      top: number;
      left: number;
  };
  export function scrollbarWidth(): number;
}

declare module 'xsheet/_/render/selection' {
  /**
   * セル・行・列の選択（XSheet の追加。M7-24・Issue #55）。
   * - 選択は 1 つの範囲（行の全体・列の全体・セル）と、今のセル（キーボードで動かす起点）。
   * - 帯はスクロール枠の中の `div.xs-sel`（重ねる層）1 つで描く。表は見えている近くだけを描き直すので、`td` に印を付けると
   *   描き直しのたびに付け直しが要り、見えていない部分は塗れないため。位置は寸法（`Geometry`）から出す。
   *   重なりは罫線の層の後・図の層の前（セル・はみ出しの文字・ハイライト・罫線の上、図・sticky の見出しの下）。
   * - 選んだ行の行番号・列の列記号は、表を描くたびに `is-selected` を付け直す（描いた行・列だけ）。
   * - クリック：セル → 指定（`'row'` は行の全体、`'cell'` はセル）、行番号 → 行の全体、列記号 → 列の全体。
   *   結合したセルは結合の範囲。動いた距離が小さいクリックだけを選択にし、文字をドラッグで選んでいるときは何もしない。
   * - キーボード（M7-25）：スクロール枠に `tabindex="0"` を付け、クリックで選んだらフォーカスを移す。行き先は selection-keys.ts。
   *   選択を動かしたキーだけ既定の動きを止め、見えない所へ動いたら最小の量だけスクロールする。
   *   選択の番地（`C5`・`5:5`・`C:C`）を見えない `aria-live` の要素に入れて、支援技術に伝える。
   * - コピー（M7-26）：フォーカスのあるスクロール枠の `copy` イベントで、選択をタブ区切りの文字にして書く（selection-copy.ts）。
   *   文字を選んでいるときはブラウザの既定のコピーに任せる。上限を超えたらコピーせず、通知の帯に出す。
   * DOM の要素は作るだけで、ユーザーのファイル内容は入れない。
   */
  import type { CellRange, colName, createGeometry, Geometry, XlsxSheet } from 'xsheet/_/core/index';
  /** セルをクリックしたときに選ぶ物（`RenderOptions.selection`） */
  export type SelectionMode = 'row' | 'cell';
  /** 選択の種類 */
  export type SelectionType = 'row' | 'column' | 'cell';
  /** 選択の変わった元（`'user'` はクリック・キー、`'api'` は `setSelection`・シートの切り替え） */
  export type SelectionOrigin = 'user' | 'api';
  /** 今の選択（XSheet: ADR-0011 §1） */
  export interface Selection {
      readonly type: SelectionType;
      /**
       * 選んだ範囲（1 始まり・両端を含む）。行の全体は列 1〜16,384、列の全体は行 1〜1,048,576。
       * 結合したセルを選ぶと結合の範囲（行の全体なら結合の行の範囲）
       */
      readonly range: CellRange;
      /** 今のセル（結合なら左上）。行・列の全体を選んでいても、キーボードで動かす起点として列・行を覚えている */
      readonly active: {
          readonly r: number;
          readonly c: number;
      };
  }
  /** `setSelection` に渡す選択（行・列の範囲は省くと 1 つ） */
  export type SelectionInput = {
      type: 'row';
      r: number;
      r2?: number;
  } | {
      type: 'column';
      c: number;
      c2?: number;
  } | {
      type: 'cell';
      r: number;
      c: number;
  };
  /** 選択が見る表示の状態（`sheet-view.ts` が渡す。どれも呼んだ時点の値） */
  export interface SelectionHost {
      /** 表示の根（`div.xs-sheet`。読み上げの要素を置く） */
      readonly root: HTMLElement;
      /** スクロール枠（`div.xs-body`。クリック・キーを受ける） */
      readonly body: HTMLElement;
      /** 帯を置く入れ物（`div.xs-wrap`）と、帯をその前に置く要素（図の層） */
      readonly wrapper: HTMLElement;
      readonly before: Element;
      sheet(): XlsxSheet | null;
      geometry(): Geometry | null;
      /** 行番号の列の幅・列見出しの高さ */
      origin(): {
          x: number;
          y: number;
      };
      /** 表の行数・列数（帯はこの範囲で切る） */
      extent(): {
          rows: number;
          cols: number;
      };
      table(): HTMLTableElement | null;
      /** 今のスクロール位置で描く（描いた範囲に収まっていれば何もしない） */
      draw(): void;
      /** 選択が変わった（同じ選択では呼ばない） */
      changed(selection: Selection | null, origin: SelectionOrigin): void;
      /** コピーするセルの数が上限を超えた（通知の帯に出す） */
      copyTooLarge(): void;
      /** 列の記号（列記号の見出しを探す。selection-gate.ts が渡す） */
      readonly colName: typeof colName;
      /** シートの寸法（コピーで非表示の行・列を見る。selection-gate.ts が渡す） */
      readonly createGeometry: typeof createGeometry;
  }
  /** {@link createSelection} の戻り値 */
  export interface SelectionController {
      get(): Selection | null;
      /**
       * 選択を替える（null で外す）。範囲は表の外でもよい（帯は表の範囲で切る）。
       * @throws RangeError 種類が正しくない・1 以上の整数でない・Excel の範囲を超える番号
       */
      set(input: SelectionInput | null, origin?: SelectionOrigin): void;
      /** 表を描き直した・寸法が変わった後に、帯の位置と見出しの印を合わせ直す */
      layout(): void;
      /** 選択をコピーする文字（タブ区切り・`
    `）。選択が無い・上限を超えると null */
      text(): string | null;
      /** 今の選択を、知らせずに外す前に知らせる（シートの切り替え）。選択が無ければ何もしない */
      clear(): void;
      dispose(): void;
  }
  /**
   * 行 r・列 c を含む結合の範囲（無ければそのセル 1 つ）。シートの行を上へ辿って、結合の左上を探す。
   * `getRows` は範囲より上で始まる結合の左上を範囲の最初の行に置くので、`MERGE_LOOKBACK` 行より上で始まる結合は、そこからの範囲になる。
   */
  export function mergeRangeAt(sheet: Pick<XlsxSheet, 'rowCount' | 'getRows'>, r: number, c: number): CellRange;
  /**
   * 選択を作る。`mode` が無ければクリックでは選ばない（`set` だけが効く）。
   * @param mode セルをクリックしたときに選ぶ物
   */
  export function createSelection(host: SelectionHost, mode: SelectionMode | undefined): SelectionController;
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
  import { type BuiltinDates, type SheetInfo, type XlsxLimits } from 'xsheet/_/core/index';
  import type { RenderOptions, SheetView } from 'xsheet/_/render/render-sheet';
  /** `renderWorkbook` の指定 */
  export interface WorkbookOptions extends RenderOptions {
      /** 読み取りの上限を替える（`openXlsx` の `limits`。既定は `DEFAULT_LIMITS`） */
      limits?: Partial<XlsxLimits>;
      /**
       * 組み込みの日付の書式 27〜36・50〜58 の見せ方（`OpenXlsxOptions.builtinDates` と同じ。`renderWorkbook` はブックを自分で開くため）。
       * XSheet: ADR-0007 §6（M7-8）
       */
      builtinDates?: BuiltinDates;
      /**
       * シートタブの列を置く場所（既定 `'top'`＝表の上。`'bottom'` は Excel と同じく表の下）。
       * `'bottom'` ではタブ列を DOM でもシートの表示の後ろに置き、根に `xs-tabs-bottom` を付ける。
       * 知らない値は `'top'` と同じ。
       * XSheet: タブの位置を選べるように足した（Issue #17）
       */
      tabPosition?: 'top' | 'bottom';
      /**
       * シートが替わり（タブ・`showSheet`）、そのシートを表示し終えた後に呼ぶ。最初のシートの表示・表示中のシートの
       * `showSheet`・後の切り替えに追い越された切り替え・`destroy` の後では呼ばない。
       * 例外は `showSheet` を失敗にせず、`setTimeout` で投げ直す。
       * @param index 表示したシートの番号（`SheetInfo.index`）
       * XSheet: ADR-0007 §7（版 1.1）
       */
      onSheetChange?: (index: number) => void;
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
   * タブ列は既定で表の上、`tabPosition: 'bottom'` なら表の下。
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
  import { type BuiltinDates, type SheetInfo, type XlsxBook, type XlsxLimits } from 'xsheet/_/core/index';
  import { type XSheetMessages } from 'xsheet/_/render/index';
  import { type XSheetSource } from 'xsheet/_/facade/source';
  import { type XSheetTheme } from 'xsheet/_/facade/style';
  /** 文言の言語。`auto` は `<html lang>`、無ければブラウザの言語。合うものが無ければ英語 */
  export type XSheetLang = 'ja' | 'en' | 'ko' | 'auto';
  /** `read` の指定 */
  export interface ReadOptions {
      /** 読み取りの上限を替える（既定は `DEFAULT_LIMITS`） */
      limits?: Partial<XlsxLimits>;
      /** 組み込みの日付の書式 27〜36・50〜58 の見せ方（`OpenXlsxOptions.builtinDates` と同じ。ADR-0007 §6） */
      builtinDates?: BuiltinDates;
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
      /** シートタブの列を置く場所（既定 `'top'`＝表の上。`'bottom'` は表の下） */
      tabPosition?: 'top' | 'bottom';
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

