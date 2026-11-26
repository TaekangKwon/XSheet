/**
 * xlsx の読み取りの上限（定数の置き場は本ファイルの 1 か所だけ。図形の guide・path 命令の上限だけは geom の `defs.ts`）。
 * 読み込むファイルは信頼できない入力として扱う（Excel の仕様上限と XML の大きさの見積もりに基づく値）。
 * zip の展開（zip 爆弾）・巨大な添字・大量の結合セルで、画面が固まったりメモリが尽きたりしないようにする。
 * 超えたら `XlsxError('too-large')`。
 */
import { MAX_GEOM_COMMANDS, MAX_GEOM_GUIDES } from '../geom/defs';
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
export declare const MAX_SHEET_XML_BYTES: number;
/** 共有文字列の上限 = 64 MiB（同じ文字は 1 回しか入らないので、シート本体より小さく見積もる） */
export declare const MAX_SHARED_STRINGS_BYTES: number;
/** 小さな部品の上限 = 16 MiB（styles.xml は書式が数万あっても数 MB。それ以上は異常） */
export declare const MAX_SMALL_PART_BYTES: number;
/** zip の中のファイル数の上限（Excel のブックは数十個。画像が多くても数千） */
export declare const MAX_ENTRIES = 10000;
/** シート数の上限（実データは 26 枚。Excel も実用上 1,000 枚で足りる） */
export declare const MAX_SHEETS = 1000;
/** 1 シートの結合セル数の上限（200 万セルのシートで全セルが 4 セル結合でも 50 万） */
export declare const MAX_MERGES = 500000;
/** 1 シートの図の数の上限（線・矢印は数えない。設計書の図は数十個。細工したファイルで DOM が膨らまない水準） */
export declare const MAX_DRAWINGS = 1000;
/** 1 シートの線・矢印の数の上限（図とは別に数える。線の多い構成図で画像・図形が押し出されないように） */
export declare const MAX_LINES = 1000;
/** 図形 1 つの文字数の上限（長い文章を入れた図形でも収まる水準。超えた分は切る） */
export declare const MAX_SHAPE_TEXT_CHARS = 10000;
/** 図形 1 つの段落数の上限 */
export declare const MAX_SHAPE_PARAGRAPHS = 500;
/** 図形 1 つのラン（`a:br` を含む）の数の上限 */
export declare const MAX_SHAPE_RUNS = 2000;
/** 1 シートの図形の文字の合計の上限（文字数＋ラン＋段落。図形 1,000 個×1,000 文字の水準。細工したファイルで DOM が膨らまない） */
export declare const MAX_SHEET_TEXT_CHARS = 1000000;
/** グループの入れ子の深さの上限（走査は XML の大きさ×この値で抑えられる。設計書の構成図は 2〜3 段） */
export declare const MAX_GROUP_DEPTH = 8;
/** 画像 1 枚の上限 = 20 MiB（画面写しは数百 KB〜数 MB） */
export declare const MAX_IMAGE_BYTES: number;
/** 1 シートの画像の合計の上限 = 128 MiB（blob にして持つ量の上限。細工した zip 爆弾でメモリが尽きない水準） */
export declare const MAX_SHEET_IMAGE_BYTES: number;
/** 画像の 1 辺の上限 = 16,384 px */
export declare const MAX_IMAGE_SIDE_PX = 16384;
/** 画像 1 枚の画素数の上限 = 64 MP（64 × 1,048,576。8,192 × 8,192 まで） */
export declare const MAX_IMAGE_PIXELS: number;
/** 図の名前・説明（`cNvPr` の name・descr）の文字数の上限。超えた分は切る */
export declare const MAX_DRAWING_TEXT = 1000;
/** 1 色の変換（`a:lumMod` などの子）の数の上限。超えた分は読まない（ADR-0028 §決定 7。Excel の画面で付けるのは数個） */
export declare const MAX_COLOR_TRANSFORMS = 16;
/** グラデーションの stop（`a:gs`）の数の上限。超えた分は読まない（Excel の画面で作れるのは 10 まで。使うのは ISSUE-95） */
export declare const MAX_GRADIENT_STOPS = 16;
/**
 * 1 図形の guide の数の上限（`MAX_GEOM_GUIDES`）と path 命令の数の上限（`MAX_GEOM_COMMANDS`）。超えたら形を描かず代替。
 * 評価器も同じ値で守るため、定義は geom の `defs.ts` に置き、ここでは引いて出し直す（定数の置き場の唯一の例外）。
 */
export { MAX_GEOM_COMMANDS, MAX_GEOM_GUIDES };
/**
 * 1 シートの自由図形の「path 命令の数＋guide の数」の合計の上限（ADR-0028 §決定 7。guide も数える＝命令 0 個で guide だけの自由図形を並べて保持量が膨らむのを防ぐ）。
 * 超えたら以後の自由図形は形を持たせず代替
 */
export declare const MAX_SHEET_GEOM_UNITS = 200000;
/** 図形 1 つの調整値（`avLst`）の数の上限（preset の最大は 8。ADR-0028 §決定 7）。超えた分は読まない */
export declare const MAX_SHAPE_ADJUSTMENTS = 16;
/** styles.xml の罫線（`<border>`）の数の上限（Excel のセルの書式の上限 64,000 を丸めた値。ADR-0029 §決定 9）。超えた分は読まない */
export declare const MAX_BORDERS = 65536;
/** styles.xml のフォント（`<font>`）の数の上限（`MAX_BORDERS` と同じ考え方。ADR-0030 §決定 11）。超えた分は読まない */
export declare const MAX_FONTS = 65536;
/** styles.xml の塗り（`<fill>`）の数の上限（`MAX_BORDERS` と同じ考え方。ADR-0030 §決定 11）。超えた分は読まない */
export declare const MAX_FILLS = 65536;
/** 重複を除いた見える書式の数の上限（＝生成 CSS の規則の数。ADR-0030 §決定 11）。超えた書式のセルは書式なし */
export declare const MAX_CELL_FORMATS = 4096;
/** グラフ 1 つの系列数の上限（Excel のグラフ 1 つの系列数の上限。ADR-0031 §決定 9） */
export declare const MAX_CHART_SERIES = 255;
/** グラフ 1 つの全系列の点の合計の上限（系列 10・点 1,000 のグラフの 10 倍の水準） */
export declare const MAX_CHART_POINTS = 100000;
/** グラフ 1 つの系列名・項目・タイトルの文字数の合計の上限 */
export declare const MAX_CHART_TEXT_CHARS = 200000;
/** 1 シートで読んだグラフの点の合計の上限（グラフ 10 個分。超えたら以後の新しいグラフは破線枠） */
export declare const MAX_SHEET_CHART_POINTS = 1000000;
/** 既定の上限の組 */
export declare const DEFAULT_LIMITS: Readonly<XlsxLimits>;
/** 書式コードの長さの上限（文字数）。非信頼の styles.xml から来る書式を ssf へ渡す前の ReDoS 対策 */
export declare const MAX_FORMAT_CODE_LENGTH = 255;
/** Excel の行数の上限（1,048,576）。これを超える行番号は不正として読まない */
export declare const MAX_ROWS = 1048576;
/** Excel の列数の上限（XFD = 16,384）。これを超える列は不正として読まない */
export declare const MAX_COLS = 16384;
/** 1 つのセルの文字数の上限（Excel の上限 32,767）。超えた分は切る */
export declare const MAX_CELL_TEXT = 32767;
/** 列幅（`<col width>`・`defaultColWidth`・`baseColWidth`、文字数）の上限。Excel の上限 255。超えた値は丸める */
export declare const MAX_COL_WIDTH = 255;
/** 行高（`<row ht>`・`defaultRowHeight`、pt）の上限。Excel の上限 409.5。超えた値は丸める */
export declare const MAX_ROW_PT = 409.5;
