/** xlsx の読み取り部品の公開の型（画面に依存しない） */
import type { Rgba } from './color';
import type { Theme } from './theme';
/** 読めない理由の種類 */
export type XlsxErrorKind = 'not-zip' | 'encrypted' | 'broken' | 'too-large' | 'unsupported';
/**
 * xlsx を読めないときの例外。種類は `kind` で決まる。
 * メッセージには、原本の中身（セルの値・シート名）を入れない。
 */
export declare class XlsxError extends Error {
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
export declare const DEFAULT_MDW = 7;
/** 寸法の指定が何も無いシートの layout（既定値のみ） */
export declare function emptyLayout(mdw?: number, fontRowPx?: number | null): SheetLayout;
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
