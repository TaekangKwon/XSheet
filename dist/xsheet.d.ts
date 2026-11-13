// XSheet 0.5.0 単独版（xsheet.min.js・xsheet.mjs）の型。tools/standalone/build.mjs が dist の d.ts から作る（手で直さない）。

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
  export type { AnchorPoint, ArrowSize, ArrowType, BorderEdge, BorderStyle, CellBorder, CellFormat, CellKind, CellView, ChartAxis, ChartBox, ChartFailReason, ChartLegend, ChartLine, ChartMarker, ChartMarkerSymbol, ChartModel, ChartPaint, ChartResult, ChartSeries, ChartTextStyle, ChartTickMark, ChartTitle, ColLayout, CustGeomDef, CustGeomPath, DrawingAnchor, DrawingColor, DrawingGroup, DrawingKind, DrawingLine, DrawingNotice, DrawingText, DrawingTextParagraph, DrawingTextRun, GeomArg, GeomPathFill, GroupLevel, GroupRect, ImageFailReason, ImageFormat, ImageResult, LineDash, LineEnd, LineGeom, ManualLayout, RowLayout, ShapeGeom, ShapeOutline, SheetDrawing, SheetDrawings, SheetInfo, SheetLayout, SheetRow, TextAlign, TextAnchor, TextVert, XlsxBook, XlsxErrorKind, XlsxSheet, } from 'xsheet/_/core/types';
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
  import type { TextMatch, XlsxBook } from 'xsheet/_/core/index';
  import { type XSheetMessages } from 'xsheet/_/render/messages';
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

