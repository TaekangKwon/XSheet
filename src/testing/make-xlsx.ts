/**
 * 試験用の xlsx を jszip で組む部品（中身はすべて架空。実在の文書は使わない）。
 * 読み取り部品（src/core）の試験で、ふつうの形と「実験が想定していなかった形」
 * （名前空間の接頭辞付き XML・`r` 属性の省略・`t="d"`・1904 年起点など）を作り分けるために使う。
 */
import JSZip from 'jszip';

const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const STRICT_MAIN = 'http://purl.oclc.org/ooxml/spreadsheetml/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/** セルの型（OOXML の `t` 属性）。省略は数値 */
export type CellType = 'n' | 's' | 'str' | 'b' | 'e' | 'd' | 'inlineStr';

/** セルの細かい指定 */
export interface CellObj {
  t?: CellType;
  /** 値。文字列は XML エスケープして入れる（inlineStr の本文もこれ） */
  v?: string | number;
  /** セルの書式（cellXfs の位置） */
  s?: number;
  /** 数式（`v` は保存済みの値） */
  f?: string;
  /** 番地を明示する。null で r 属性を省く（このセルだけ） */
  r?: string | null;
  /** タグへそのまま足す属性（例 ` __proto__="x"`） */
  attrs?: string;
}

/**
 * セルの指定。位置は rows の添字（0 行目 0 列目＝A1）。
 * - 文字列：inlineStr
 * - 数値：数値セル
 * - 真偽：t="b"
 * - null：セルを作らない（空の位置）
 */
export type CellSpec = string | number | boolean | null | CellObj;

/** `<cols>` の `<col>` 1 件の指定 */
export interface ColSpec {
  min: number;
  max: number;
  /** 列幅（文字数）。省略すると width 属性を付けない */
  width?: number;
  hidden?: boolean;
  /** `style`（cellXfs の位置。列全体の書式） */
  style?: number;
}

/** `<sheetFormatPr>` の指定（指定した項目だけ属性にする） */
export interface FormatPrSpec {
  defaultColWidth?: number;
  baseColWidth?: number;
  /** pt */
  defaultRowHeight?: number;
  zeroHeight?: boolean;
  /** `customHeight="1"`（defaultRowHeight をそのまま既定の行高にする） */
  customHeight?: boolean;
}

/** drawing の rels の 1 件 */
export interface DrawingRelSpec {
  id: string;
  /** Target（drawing の位置からの相対。`../media/image1.png`） */
  target: string;
  /** Type の最後の語（既定 `image`） */
  type?: string;
  /** `TargetMode="External"` を付ける */
  external?: boolean;
}

/** シートの drawing の指定（画像・図形の試験用。アンカーの XML は下の `wsDr` ほかで組む） */
export interface DrawingSpec {
  /** drawing の XML 全体（`wsDr` で包んだもの） */
  xml: string;
  /** drawing の rels（Id → Target） */
  rels?: DrawingRelSpec[];
  /** zip の中のファイル名（既定 `drawing<シート番号>.xml`） */
  file?: string;
  /** drawing の本体を zip に入れない（シートの rels だけが指す＝zip に解決できない Target の試験） */
  omitPart?: boolean;
}

/** シート 1 枚の指定 */
export interface SheetSpec {
  name: string;
  state?: 'hidden' | 'veryHidden';
  /** 行ごとのセル。行番号は 1 から */
  rows?: CellSpec[][];
  /** 結合セル（`A1:C2`） */
  merges?: string[];
  /** 列の寸法（`<cols>`） */
  cols?: ColSpec[];
  /** `<sheetFormatPr>` */
  formatPr?: FormatPrSpec;
  /** 行の寸法（行番号 → `ht`（pt）・`hidden`・`custom`＝customHeight を付けるか、既定は付ける）。rows の行に付く（rows に無い行番号は `<row>` を追加しない） */
  rowAttrs?: Record<
    number,
    {
      ht?: number;
      hidden?: boolean;
      custom?: boolean;
      /** 行の書式（`s`＝cellXfs の位置）。効くのは `customFormat` が真のとき */
      s?: number;
      /** `customFormat="1"` を付ける */
      customFormat?: boolean;
    }
  >;
  /** `<sheetData>` の中身を直接指定する（rows より優先） */
  sheetDataXml?: string;
  /** シートの XML 全体を直接指定する（最優先） */
  xml?: string;
  /** zip の中のファイル名（既定 `sheet<番号>.xml`。workbook の順とファイル名の順をずらす試験用） */
  file?: string;
  /** グラフシート（chartsheet）として rels に載せる */
  chartsheet?: boolean;
  /** drawing（`<drawing r:id>`・シートの rels・drawingN.xml・その rels を足す）。画像本体は XlsxSpec.extraFiles に `xl/media/…` で入れる */
  drawing?: DrawingSpec;
}

/** テーマ（`xl/theme/<file>`）の指定 */
export interface ThemeSpec {
  /** theme1.xml の中身（`buildThemeXml` で組める） */
  xml: string;
  /** `xl/theme/` の中のファイル名（既定 `theme1.xml`） */
  file?: string;
  /** workbook.xml.rels に theme の関係を載せない（部品だけが zip にある） */
  omitRel?: boolean;
  /** 部品を zip に入れない（rels だけが指す） */
  omitPart?: boolean;
}

/** xlsx 全体の指定 */
export interface XlsxSpec {
  sheets: SheetSpec[];
  /** 共有文字列。各要素は `<si>` の中身の XML 断片（例 `<t>abc</t>`） */
  sharedStrings?: string[];
  /** 独自の表示形式（numFmtId → 書式コード） */
  numFmts?: Record<number, string>;
  /** cellXfs。各要素は numFmtId（位置がセルの s） */
  xfs?: number[];
  /** 罫線（`<borders>`）。各要素は `<border>` 要素の全体（`borderXml` で組める）。省くと `<borders>` を入れない */
  borders?: string[];
  /** cellXfs の `borderId`（xfs と同じ位置。省略した位置は 0） */
  xfBorderIds?: number[];
  /** フォント（`<fonts>`）。各要素は `<font>` 要素の全体（`fontXml` で組める。先頭が既定のフォント）。`font` より優先。省くと `font` に従う */
  fonts?: string[];
  /** 塗り（`<fills>`）。各要素は `<fill>` 要素の全体（`patternFillXml`・`gradientFillXml` で組める）。省くと `<fills>` を入れない */
  fills?: string[];
  /** cellXfs の `fontId`（xfs と同じ位置。省略した位置は 0） */
  xfFontIds?: number[];
  /** cellXfs の `fillId`（xfs と同じ位置。省略した位置は 0） */
  xfFillIds?: number[];
  /** `<colors><indexedColors>` の `rgb`（ARGB 8 桁の 16 進。`indexed` 0 から順に上書きする）。省くと入れない */
  indexedColors?: string[];
  /** styles.xml の先頭の font（最大桁幅 MDW の算出元）。省略すると `<fonts>` を入れない */
  font?: { name: string; size: number };
  date1904?: boolean;
  /** spreadsheetml の要素に付ける名前空間の接頭辞（例 `x` → `<x:c>`） */
  prefix?: string;
  /** すべてのセルの r 属性を省く */
  omitRef?: boolean;
  /** Strict 形式の名前空間にする */
  strict?: boolean;
  /** styles.xml を入れない */
  noStyles?: boolean;
  /** sharedStrings.xml を入れない（sharedStrings が空でも入れない） */
  noSharedStrings?: boolean;
  /** テーマ（`xl/theme/theme1.xml`・workbook.xml.rels の theme の関係）。省くとテーマ無し */
  theme?: ThemeSpec;
  /** グラフの部品（`xl/charts/<キー>` → XML。`chartXml` で組める）。`[Content_Types].xml` に Override も足す。省くと何も足さない */
  charts?: Record<string, string>;
  /** 追加・上書きのファイル */
  extraFiles?: Record<string, string | Uint8Array>;
}

/** XML の文字をエスケープする */
export const esc = (s: string | number): string =>
  String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');

/** 列番号（1 始まり）→ 記号 */
export function colLetters(c: number): string {
  let s = '';
  for (let n = c; n > 0; n = Math.floor((n - 1) / 26))
    s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

/** rows の指定から `<sheetData>` の中身を作る */
function buildSheetData(
  rows: CellSpec[][],
  q: (name: string) => string,
  omitRef: boolean,
  rowAttrs: SheetSpec['rowAttrs'] = {},
): string {
  return rows
    .map((row, ri) => {
      const r = ri + 1;
      const cells = row
        .map((spec, ci) => {
          if (spec === null) return '';
          const o: CellObj =
            typeof spec === 'string'
              ? { t: 'inlineStr', v: spec }
              : typeof spec === 'number'
                ? { v: spec }
                : typeof spec === 'boolean'
                  ? { t: 'b', v: spec ? 1 : 0 }
                  : spec;
          const ref = o.r === undefined ? `${colLetters(ci + 1)}${r}` : o.r;
          const attrs =
            (ref !== null && !omitRef ? ` r="${ref}"` : '') +
            (o.s ? ` s="${o.s}"` : '') +
            (o.t && o.t !== 'n' ? ` t="${o.t}"` : '') +
            (o.attrs ?? '');
          const f = o.f !== undefined ? `<${q('f')}>${esc(o.f)}</${q('f')}>` : '';
          if (o.t === 'inlineStr') {
            return `<${q('c')}${attrs}><${q('is')}><${q('t')} xml:space="preserve">${esc(o.v ?? '')}</${q('t')}></${q('is')}></${q('c')}>`;
          }
          if (o.v === undefined) return `<${q('c')}${attrs}/>`;
          const v = o.t === 'str' || o.t === 'e' || o.t === 'd' ? esc(o.v) : String(o.v);
          return `<${q('c')}${attrs}>${f}<${q('v')}>${v}</${q('v')}></${q('c')}>`;
        })
        .join('');
      const ra = rowAttrs[r];
      const rowExtra =
        (ra?.ht !== undefined
          ? ` ht="${ra.ht}"${ra.custom === false ? '' : ' customHeight="1"'}`
          : '') +
        (ra?.hidden ? ' hidden="1"' : '') +
        (ra?.s !== undefined ? ` s="${ra.s}"` : '') +
        (ra?.customFormat ? ' customFormat="1"' : '');
      return `<${q('row')} r="${r}"${rowExtra}>${cells}</${q('row')}>`;
    })
    .join('');
}

/** `<sheetFormatPr>`（指定が無ければ空文字） */
function formatPrXml(p: FormatPrSpec | undefined, q: (name: string) => string): string {
  if (p === undefined) return '';
  const attrs =
    (p.baseColWidth !== undefined ? ` baseColWidth="${p.baseColWidth}"` : '') +
    (p.defaultColWidth !== undefined ? ` defaultColWidth="${p.defaultColWidth}"` : '') +
    (p.defaultRowHeight !== undefined ? ` defaultRowHeight="${p.defaultRowHeight}"` : '') +
    (p.customHeight ? ' customHeight="1"' : '') +
    (p.zeroHeight ? ' zeroHeight="1"' : '');
  return `<${q('sheetFormatPr')}${attrs}/>`;
}

/** `<cols>`（指定が無ければ空文字） */
function colsXml(cols: ColSpec[] | undefined, q: (name: string) => string): string {
  if (!cols?.length) return '';
  const items = cols
    .map(
      (c) =>
        `<${q('col')} min="${c.min}" max="${c.max}"${c.width !== undefined ? ` width="${c.width}" customWidth="1"` : ''}${c.hidden ? ' hidden="1"' : ''}${c.style !== undefined ? ` style="${c.style}"` : ''}/>`,
    )
    .join('');
  return `<${q('cols')}>${items}</${q('cols')}>`;
}

/** 試験用の xlsx（zip のバイト列）を組む */
export async function buildXlsx(spec: XlsxSpec): Promise<Uint8Array> {
  const prefix = spec.prefix;
  const q = (name: string) => (prefix ? `${prefix}:${name}` : name);
  const ns = spec.strict ? STRICT_MAIN : MAIN;
  const nsDecl = prefix ? `xmlns:${prefix}="${ns}"` : `xmlns="${ns}"`;
  const z = new JSZip();

  const fileOf = (s: SheetSpec, i: number) => s.file ?? `sheet${i + 1}.xml`;
  const chartOverrides = Object.keys(spec.charts ?? {})
    .map(
      (name) =>
        `<Override PartName="/xl/charts/${esc(name)}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`,
    )
    .join('');
  // グラフを持つ spec だけ、Excel が開けるよう主要パーツの Override も足す（Excel は workbook.xml が
  // application/xml のままだと「ファイル形式が正しくありません」で拒否する。既存の合成ファイルは変えない）
  const SS = 'application/vnd.openxmlformats-officedocument.spreadsheetml';
  const coreOverrides = chartOverrides
    ? `<Override PartName="/xl/workbook.xml" ContentType="${SS}.sheet.main+xml"/>` +
      spec.sheets
        .map((s, i) =>
          s.chartsheet
            ? `<Override PartName="/xl/chartsheets/${esc(fileOf(s, i))}" ContentType="${SS}.chartsheet+xml"/>`
            : `<Override PartName="/xl/worksheets/${esc(fileOf(s, i))}" ContentType="${SS}.worksheet+xml"/>`,
        )
        .join('') +
      spec.sheets
        .map((s, i) =>
          s.drawing && !s.drawing.omitPart
            ? `<Override PartName="/xl/drawings/${esc(s.drawing.file ?? `drawing${i + 1}.xml`)}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`
            : '',
        )
        .join('') +
      (spec.noStyles
        ? ''
        : `<Override PartName="/xl/styles.xml" ContentType="${SS}.styles+xml"/>`) +
      (!spec.noSharedStrings && spec.sharedStrings !== undefined
        ? `<Override PartName="/xl/sharedStrings.xml" ContentType="${SS}.sharedStrings+xml"/>`
        : '') +
      (spec.theme && !spec.theme.omitPart
        ? `<Override PartName="/xl/theme/${esc(spec.theme.file ?? 'theme1.xml')}" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>`
        : '')
    : '';
  z.file(
    '[Content_Types].xml',
    `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${coreOverrides}${chartOverrides}</Types>`,
  );
  z.file(
    '_rels/.rels',
    `${HEAD}<Relationships xmlns="${PKG_REL}"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
  );

  const sheetEls = spec.sheets
    .map(
      (s, i) =>
        `<${q('sheet')} name="${esc(s.name)}" sheetId="${10 + i}"${s.state ? ` state="${s.state}"` : ''} r:id="rId${i + 1}"/>`,
    )
    .join('');
  z.file(
    'xl/workbook.xml',
    `${HEAD}<${q('workbook')} ${nsDecl} xmlns:r="${REL}"><${q('workbookPr')}${spec.date1904 ? ' date1904="1"' : ''}/><${q('sheets')}>${sheetEls}</${q('sheets')}></${q('workbook')}>`,
  );

  const hasSst = !spec.noSharedStrings && spec.sharedStrings !== undefined;
  const rels =
    spec.sheets
      .map((s, i) => {
        const kind = s.chartsheet ? 'chartsheet' : 'worksheet';
        const dir = s.chartsheet ? 'chartsheets' : 'worksheets';
        return `<Relationship Id="rId${i + 1}" Type="${REL}/${kind}" Target="${dir}/${fileOf(s, i)}"/>`;
      })
      .join('') +
    (spec.noStyles ? '' : `<Relationship Id="rId90" Type="${REL}/styles" Target="styles.xml"/>`) +
    (hasSst
      ? `<Relationship Id="rId91" Type="${REL}/sharedStrings" Target="sharedStrings.xml"/>`
      : '') +
    (spec.theme && !spec.theme.omitRel
      ? `<Relationship Id="rId92" Type="${REL}/theme" Target="theme/${spec.theme.file ?? 'theme1.xml'}"/>`
      : '');
  z.file(
    'xl/_rels/workbook.xml.rels',
    `${HEAD}<Relationships xmlns="${PKG_REL}">${rels}</Relationships>`,
  );

  spec.sheets.forEach((s, i) => {
    const dir = s.chartsheet ? 'chartsheets' : 'worksheets';
    let xml = s.xml;
    if (xml === undefined) {
      const data =
        s.sheetDataXml ?? buildSheetData(s.rows ?? [], q, spec.omitRef ?? false, s.rowAttrs);
      const merges = s.merges?.length
        ? `<${q('mergeCells')} count="${s.merges.length}">${s.merges.map((m) => `<${q('mergeCell')} ref="${m}"/>`).join('')}</${q('mergeCells')}>`
        : '';
      const rNs = s.drawing ? ` xmlns:r="${REL}"` : '';
      const drawingRef = s.drawing ? `<${q('drawing')} r:id="rId1"/>` : '';
      xml = `${HEAD}<${q('worksheet')} ${nsDecl}${rNs}>${formatPrXml(s.formatPr, q)}${colsXml(s.cols, q)}<${q('sheetData')}>${data}</${q('sheetData')}>${merges}${drawingRef}</${q('worksheet')}>`;
    }
    z.file(`xl/${dir}/${fileOf(s, i)}`, xml);
    if (s.drawing) {
      const dFile = s.drawing.file ?? `drawing${i + 1}.xml`;
      z.file(
        `xl/${dir}/_rels/${fileOf(s, i)}.rels`,
        `${HEAD}<Relationships xmlns="${PKG_REL}"><Relationship Id="rId1" Type="${REL}/drawing" Target="../drawings/${dFile}"/></Relationships>`,
      );
      if (!s.drawing.omitPart) z.file(`xl/drawings/${dFile}`, s.drawing.xml);
      if (s.drawing.rels?.length) {
        const items = s.drawing.rels
          .map(
            (r) =>
              `<Relationship Id="${r.id}" Type="${REL}/${r.type ?? 'image'}" Target="${esc(r.target)}"${r.external ? ' TargetMode="External"' : ''}/>`,
          )
          .join('');
        z.file(
          `xl/drawings/_rels/${dFile}.rels`,
          `${HEAD}<Relationships xmlns="${PKG_REL}">${items}</Relationships>`,
        );
      }
    }
  });

  if (spec.theme && !spec.theme.omitPart)
    z.file(`xl/theme/${spec.theme.file ?? 'theme1.xml'}`, spec.theme.xml);

  if (hasSst) {
    const items = spec.sharedStrings!;
    z.file(
      'xl/sharedStrings.xml',
      `${HEAD}<${q('sst')} ${nsDecl} count="${items.length}" uniqueCount="${items.length}">${items.map((si) => `<${q('si')}>${si}</${q('si')}>`).join('')}</${q('sst')}>`,
    );
  }

  if (!spec.noStyles) {
    const fmts = Object.entries(spec.numFmts ?? {});
    const numFmts = fmts.length
      ? `<${q('numFmts')} count="${fmts.length}">${fmts.map(([id, code]) => `<${q('numFmt')} numFmtId="${id}" formatCode="${esc(code)}"/>`).join('')}</${q('numFmts')}>`
      : '';
    const fonts = spec.fonts
      ? `<${q('fonts')} count="${spec.fonts.length}">${spec.fonts.join('')}</${q('fonts')}>`
      : spec.font
        ? `<${q('fonts')} count="1"><${q('font')}><${q('sz')} val="${spec.font.size}"/><${q('name')} val="${esc(spec.font.name)}"/></${q('font')}></${q('fonts')}>`
        : '';
    const fills = spec.fills
      ? `<${q('fills')} count="${spec.fills.length}">${spec.fills.join('')}</${q('fills')}>`
      : '';
    const xfs = spec.xfs ?? [0];
    const borders = spec.borders
      ? `<${q('borders')} count="${spec.borders.length}">${spec.borders.join('')}</${q('borders')}>`
      : '';
    const colors = spec.indexedColors
      ? `<${q('colors')}><${q('indexedColors')}>${spec.indexedColors.map((c) => `<${q('rgbColor')} rgb="${c}"/>`).join('')}</${q('indexedColors')}></${q('colors')}>`
      : '';
    // cellStyleXfs にも <xf> を置く（cellXfs の数え方が混ざらないことの確認）
    z.file(
      'xl/styles.xml',
      `${HEAD}<${q('styleSheet')} ${nsDecl}>${numFmts}${fonts}${fills}${borders}<${q('cellStyleXfs')} count="1"><${q('xf')} numFmtId="49" fontId="0" fillId="0" borderId="0"/></${q('cellStyleXfs')}><${q('cellXfs')} count="${xfs.length}">${xfs.map((id, i) => `<${q('xf')} numFmtId="${id}" fontId="${spec.xfFontIds?.[i] ?? 0}" fillId="${spec.xfFillIds?.[i] ?? 0}" borderId="${spec.xfBorderIds?.[i] ?? 0}" xfId="0"><${q('alignment')} vertical="top"/></${q('xf')}>`).join('')}</${q('cellXfs')}>${colors}</${q('styleSheet')}>`,
    );
  }

  for (const [name, xml] of Object.entries(spec.charts ?? {})) z.file(`xl/charts/${name}`, xml);
  for (const [path, content] of Object.entries(spec.extraFiles ?? {})) z.file(path, content);
  return z.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

/** 罫線の 1 辺の指定（線種と、`<color>` の属性の文字列。例 `{ style: 'thin', color: 'auto="1"' }`。color 省略は `<color>` を付けない） */
export interface BorderEdgeSpec {
  style?: string;
  color?: string;
}

/**
 * `<border>` 要素を組む（`XlsxSpec.borders` の 1 件）。
 * edges のキーは辺の要素名（left・right・top・bottom・diagonal・start・end）。attrs は `<border>` へ足す属性（例 ` diagonalUp="1"`）。
 */
export function borderXml(edges: Partial<Record<string, BorderEdgeSpec>> = {}, attrs = ''): string {
  const inner = Object.entries(edges)
    .map(([name, e]) => {
      const style = e?.style !== undefined ? ` style="${e.style}"` : '';
      return e?.color !== undefined
        ? `<${name}${style}><color ${e.color}/></${name}>`
        : `<${name}${style}/>`;
    })
    .join('');
  return `<border${attrs}>${inner}</border>`;
}

/** `<font>` 要素を組む（`XlsxSpec.fonts` の 1 件）。inner は中身の XML（例 `<b/><sz val="11"/><color rgb="FFFF0000"/>`） */
export function fontXml(inner = ''): string {
  return `<font>${inner}</font>`;
}

/**
 * `<fill>`（`patternFill`）を組む（`XlsxSpec.fills` の 1 件）。
 * fg・bg は `<fgColor>`・`<bgColor>` の属性の文字列（例 `rgb="FFFF0000"`。省くと要素を付けない）。patternType 省略は属性を付けない
 */
export function patternFillXml(patternType?: string, fg?: string, bg?: string): string {
  const type = patternType !== undefined ? ` patternType="${patternType}"` : '';
  const inner =
    (fg !== undefined ? `<fgColor ${fg}/>` : '') + (bg !== undefined ? `<bgColor ${bg}/>` : '');
  return inner === ''
    ? `<fill><patternFill${type}/></fill>`
    : `<fill><patternFill${type}>${inner}</patternFill></fill>`;
}

/**
 * `<fill>`（`gradientFill`）を組む。stops は `[position（0〜1）, <color> の属性の文字列]` の並び。attrs は `<gradientFill>` へ足す属性（例 ` type="path"`）
 */
export function gradientFillXml(stops: [number, string][], attrs = ''): string {
  const inner = stops.map(([p, c]) => `<stop position="${p}"><color ${c}/></stop>`).join('');
  return `<fill><gradientFill${attrs}>${inner}</gradientFill></fill>`;
}

/**
 * theme1.xml を組む（`a:clrScheme` の位置ごとの中身 XML を渡す。例 `{ accent1: '<a:srgbClr val="112233"/>' }`）。
 * 渡した位置だけを出す。後ろに `a:fmtScheme`（`phClr`）も置く。
 */
export function buildThemeXml(slots: Record<string, string>): string {
  const colors = Object.entries(slots)
    .map(([name, inner]) => `<a:${name}>${inner}</a:${name}>`)
    .join('');
  return `${HEAD}<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T"><a:themeElements><a:clrScheme name="T">${colors}</a:clrScheme><a:fmtScheme name="T"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst></a:fmtScheme></a:themeElements></a:theme>`;
}

/** 任意のファイルの並びから zip を組む（部品が欠けた・Strict などの試験用） */
export async function buildZip(files: Record<string, string | Uint8Array>): Promise<Uint8Array> {
  const z = new JSZip();
  for (const [path, content] of Object.entries(files)) z.file(path, content);
  return z.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

/** 文字列を UTF-8 のバイト列にする */
export const utf8 = (s: string): Uint8Array => new TextEncoder().encode(s);

// ---- drawing（図）の XML の部品。試験で必要な要素だけを、Excel が書く形に近づけて組む ----

/** drawing のアンカーの 1 点（番号は 0 始まり） */
export interface PointSpec {
  col: number;
  colOff?: number;
  row: number;
  rowOff?: number;
}

const XDR_NS = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const C_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const CX_NS = 'http://schemas.microsoft.com/office/drawing/2014/chartex';

/** `<xdr:from>` / `<xdr:to>` */
function pointXml(tag: 'from' | 'to', p: PointSpec): string {
  return `<xdr:${tag}><xdr:col>${p.col}</xdr:col><xdr:colOff>${p.colOff ?? 0}</xdr:colOff><xdr:row>${p.row}</xdr:row><xdr:rowOff>${p.rowOff ?? 0}</xdr:rowOff></xdr:${tag}>`;
}

/** drawing XML 全体（`xdr:wsDr`）。アンカーの XML を並べて包む */
export function wsDr(anchors: string[]): string {
  return `${HEAD}<xdr:wsDr xmlns:xdr="${XDR_NS}" xmlns:a="${A_NS}" xmlns:mc="${MC_NS}" xmlns:r="${REL}">${anchors.join('')}</xdr:wsDr>`;
}

/** twoCellAnchor（`editAs` は任意）。content は図の要素（`picXml` など） */
export function twoCellAnchor(
  from: PointSpec,
  to: PointSpec,
  content: string,
  editAs?: string,
): string {
  return `<xdr:twoCellAnchor${editAs ? ` editAs="${editAs}"` : ''}>${pointXml('from', from)}${pointXml('to', to)}${content}<xdr:clientData/></xdr:twoCellAnchor>`;
}

/** oneCellAnchor（`from`＋`ext`） */
export function oneCellAnchor(from: PointSpec, cx: number, cy: number, content: string): string {
  return `<xdr:oneCellAnchor>${pointXml('from', from)}<xdr:ext cx="${cx}" cy="${cy}"/>${content}<xdr:clientData/></xdr:oneCellAnchor>`;
}

/** absoluteAnchor（`pos`＋`ext`） */
export function absoluteAnchor(
  x: number,
  y: number,
  cx: number,
  cy: number,
  content: string,
): string {
  return `<xdr:absoluteAnchor><xdr:pos x="${x}" y="${y}"/><xdr:ext cx="${cx}" cy="${cy}"/>${content}<xdr:clientData/></xdr:absoluteAnchor>`;
}

/** 図の共通の `cNvPr` */
function cNvPr(id: number, name: string, descr?: string, hidden?: boolean): string {
  return `<xdr:cNvPr id="${id}" name="${esc(name)}"${descr !== undefined ? ` descr="${esc(descr)}"` : ''}${hidden ? ' hidden="1"' : ''}/>`;
}

/** `a:off`・`a:ext` の値（x, y, cx, cy。EMU） */
export type Box = [number, number, number, number];

/** `a:off`・`a:ext`（`a:xfrm` の中身。省くと 0, 0, 100, 100） */
function offExt(box: Box | undefined): string {
  const [x, y, cx, cy] = box ?? [0, 0, 100, 100];
  return `<a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/>`;
}

/** 画像（`xdr:pic`）の指定 */
export interface PicSpec {
  name: string;
  id?: number;
  descr?: string;
  hidden?: boolean;
  /** `a:blip r:embed`（省くと embed 無し） */
  embed?: string;
  /** `a:blip r:link`（外部の画像） */
  link?: string;
  /** `asvg:svgBlip r:embed`（SVG 併記） */
  svgEmbed?: string;
  /** `a:srcRect`（トリミング）の属性（例 `l="16667"`） */
  srcRect?: string;
  /** `a:xfrm` の属性（例 `rot="1800000"`） */
  xfrmAttrs?: string;
  /** `a:off`・`a:ext`（x, y, cx, cy。EMU。グループの子の座標系）。省くと 0, 0, 100, 100 */
  box?: Box;
}

/** 画像の要素 */
export function picXml(p: PicSpec): string {
  const svg = p.svgEmbed
    ? `<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="${p.svgEmbed}"/></a:ext></a:extLst>`
    : '';
  const attrs = (p.embed ? ` r:embed="${p.embed}"` : '') + (p.link ? ` r:link="${p.link}"` : '');
  const src = p.srcRect ? `<a:srcRect ${p.srcRect}/>` : '';
  const xfrm = p.xfrmAttrs ? ` ${p.xfrmAttrs}` : '';
  return `<xdr:pic><xdr:nvPicPr>${cNvPr(p.id ?? 2, p.name, p.descr, p.hidden)}<xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip${attrs}>${svg}</a:blip>${src}<a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm${xfrm}>${offExt(p.box)}</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>`;
}

/** 図形（`xdr:sp`）。`txBox` ならテキストボックス */
export function spXml(name: string, txBox = false, id = 3): string {
  return `<xdr:sp macro="" textlink=""><xdr:nvSpPr>${cNvPr(id, name)}<xdr:cNvSpPr${txBox ? ' txBox="1"' : ''}/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:sp>`;
}

/** 線・矢印（`xdr:cxnSp`） */
export function cxnSpXml(name: string, id = 4): string {
  return `<xdr:cxnSp macro=""><xdr:nvCxnSpPr>${cNvPr(id, name)}<xdr:cNvCxnSpPr/></xdr:nvCxnSpPr><xdr:spPr><a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom></xdr:spPr></xdr:cxnSp>`;
}

/**
 * グラフ・SmartArt（`xdr:graphicFrame`）。`uri` は `graphicData` の uri（既定はグラフ）。
 * `chartRid` を渡すと `graphicData` の子に `c:chart r:id` を出す（省くと自己閉じの `graphicData`＝今までと同じ出力）。
 * `chartTag` は子の要素名（既定 `c:chart`。chartex の試験では `cx:chart`）
 */
export function graphicFrameXml(
  name: string,
  uri = 'http://schemas.openxmlformats.org/drawingml/2006/chart',
  id = 5,
  box: Box = [0, 0, 0, 0],
  chartRid?: string,
  chartTag = 'c:chart',
): string {
  const prefix = chartTag.split(':')[0]!;
  const ns = prefix === 'cx' ? CX_NS : C_NS;
  const data =
    chartRid === undefined
      ? `<a:graphicData uri="${uri}"/>`
      : `<a:graphicData uri="${uri}"><${chartTag} xmlns:${prefix}="${ns}" r:id="${chartRid}"/></a:graphicData>`;
  return `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr>${cNvPr(id, name)}<xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm>${offExt(box)}</xdr:xfrm><a:graphic>${data}</a:graphic></xdr:graphicFrame>`;
}

/** グループの座標系（`xdr:grpSpPr` の `a:xfrm`） */
export interface GroupFrameSpec {
  /** `a:off`（x, y）・`a:ext`（cx, cy）：親の座標系での位置と大きさ */
  box: Box;
  /** `a:chOff`（x, y）・`a:chExt`（cx, cy）：子の座標系。省くと `box` と同じ */
  ch?: Box;
  /** `a:xfrm` の属性（例 `rot="1800000"`） */
  attrs?: string;
}

/** グループ（`xdr:grpSp`）。children は中の図の要素。frame を省くと `xdr:grpSpPr` は空（座標系の無いグループ） */
export function grpSpXml(name: string, children: string[], id = 6, frame?: GroupFrameSpec): string {
  let pr = '<xdr:grpSpPr/>';
  if (frame) {
    const [x, y, cx, cy] = frame.box;
    const [cox, coy, cex, cey] = frame.ch ?? frame.box;
    pr = `<xdr:grpSpPr><a:xfrm${frame.attrs ? ` ${frame.attrs}` : ''}><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/><a:chOff x="${cox}" y="${coy}"/><a:chExt cx="${cex}" cy="${cey}"/></a:xfrm></xdr:grpSpPr>`;
  }
  return `<xdr:grpSp><xdr:nvGrpSpPr>${cNvPr(id, name)}<xdr:cNvGrpSpPr/></xdr:nvGrpSpPr>${pr}${children.join('')}</xdr:grpSp>`;
}

/** `mc:AlternateContent`（Choice が 1 つと Fallback が 1 つ） */
export function alternateContent(choice: string, fallback: string): string {
  return `<mc:AlternateContent><mc:Choice Requires="a14">${choice}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`;
}

// ---- 図形の文字・線（ISSUE-87）。Excel が書く形に近づけた txBody・xfrm・a:ln・xdr:style の合成 ----

/** 文字のラン 1 つの指定（`text` は実体参照をエスケープしない生の文字列＝`esc` する） */
export interface RunSpec {
  text?: string;
  /** `a:rPr sz`（1/100pt） */
  sz?: number;
  b?: boolean;
  i?: boolean;
  /** `a:rPr u`（例 `sng`・`none`） */
  u?: string;
  /** `a:rPr strike`（例 `sngStrike`・`noStrike`） */
  strike?: string;
  /** `a:rPr` の中身の XML をそのまま入れる（例 `<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>`。ISSUE-95） */
  rPrInner?: string;
  /** true なら `a:br`（改行。`sz` は `a:br` の `a:rPr` に付く） */
  br?: boolean;
  /** true なら `a:fld type="TxLink"`（セル参照の文字） */
  fld?: boolean;
  /** `a:t` の中身をそのまま入れる（エスケープしない。実体参照・制御文字の試験用） */
  rawText?: string;
}

/** 段落 1 つの指定 */
export interface ParaSpec {
  /** 文字列はそのままの 1 ラン */
  runs?: (string | RunSpec)[];
  /** `a:pPr algn` */
  algn?: string;
  /** 箇条書き：`char` は `a:buChar`、`autoNum` は `a:buAutoNum type`（`startAt` 任意）、`none` は `a:buNone` */
  bullet?: { char: string } | { autoNum: string; startAt?: number } | 'none';
  /** `a:endParaRPr sz`（1/100pt） */
  endSz?: number;
  /** `a:p` の中身をそのまま入れる（`a:pPr` から。上の指定は無視） */
  raw?: string;
}

/**
 * `xdr:style` の指定。`ln`・`fill` は `lnRef idx`・`fillRef idx`。
 * `lnClr`・`fillClr`・`fontClr` は各参照の子の色の要素の XML（省くと `lnRef`・`fillRef` は `accent1`、`fontRef` は `lt1`。空文字なら子を持たない自己閉じ）。
 */
export interface StyleSpec {
  ln: number;
  fill: number;
  lnClr?: string;
  fillClr?: string;
  fontClr?: string;
}

/** 図形（`xdr:sp`）の指定 */
export interface ShapeSpec {
  name: string;
  id?: number;
  txBox?: boolean;
  hidden?: boolean;
  /** `sp textlink`（セル参照の文字） */
  textlink?: string;
  /** `a:prstGeom prst`（既定 `rect`） */
  prst?: string;
  /** `a:avLst` の中身（`a:gd` の並び） */
  avLst?: string;
  /** `a:custGeom` 要素全体（指定すると `a:prstGeom` の代わりに入れる。`custGeomXml` で組める） */
  custGeom?: string;
  /** `a:xfrm` の属性（例 `rot="16200000" flipH="1"`） */
  xfrm?: string;
  /** `a:off`・`a:ext`（x, y, cx, cy。EMU）。指定すると `a:xfrm` を作る（グループの子の座標系で使う） */
  box?: Box;
  /** `spPr` 直下の塗りの XML（例 `<a:noFill/>`） */
  fill?: string;
  /** `a:ln` 要素全体（例 `<a:ln w="9525"><a:noFill/></a:ln>`） */
  ln?: string;
  /** `xdr:style` の `lnRef idx`・`fillRef idx`（省くと `xdr:style` を作らない）。色の子は `StyleSpec` の説明のとおり */
  style?: StyleSpec;
  /** `xdr:txBody` の段落（文字列はそのままの 1 段落・1 ラン）。省くと `txBody` を作らない */
  paras?: (string | ParaSpec)[];
  /** `a:bodyPr` の属性（例 `anchor="ctr" wrap="none"`） */
  bodyPr?: string;
  /** `a:bodyPr` の子（例 `<a:normAutofit fontScale="92500"/>`） */
  bodyPrInner?: string;
}

/** `a:rPr`・`a:endParaRPr` の属性 */
function rPrAttrs(r: RunSpec): string {
  return (
    (r.sz !== undefined ? ` sz="${r.sz}"` : '') +
    (r.b !== undefined ? ` b="${r.b ? 1 : 0}"` : '') +
    (r.i !== undefined ? ` i="${r.i ? 1 : 0}"` : '') +
    (r.u !== undefined ? ` u="${r.u}"` : '') +
    (r.strike !== undefined ? ` strike="${r.strike}"` : '')
  );
}

/** ラン 1 つ（`a:r`・`a:br`・`a:fld`） */
function runXml(r: string | RunSpec): string {
  const o: RunSpec = typeof r === 'string' ? { text: r } : r;
  const rPr =
    o.rPrInner === undefined
      ? `<a:rPr lang="ja-JP"${rPrAttrs(o)}/>`
      : `<a:rPr lang="ja-JP"${rPrAttrs(o)}>${o.rPrInner}</a:rPr>`;
  if (o.br) return `<a:br>${rPr}</a:br>`;
  const t = `<a:t>${o.rawText ?? esc(o.text ?? '')}</a:t>`;
  if (o.fld)
    return `<a:fld id="{00000000-0000-0000-0000-000000000001}" type="TxLink">${rPr}${t}</a:fld>`;
  return `<a:r>${rPr}${t}</a:r>`;
}

/** 段落 1 つ（`a:p`） */
function paraXml(p: string | ParaSpec): string {
  const o: ParaSpec = typeof p === 'string' ? { runs: [p] } : p;
  if (o.raw !== undefined) return `<a:p>${o.raw}</a:p>`;
  let bu = '';
  if (o.bullet === 'none') bu = '<a:buNone/>';
  else if (o.bullet !== undefined && 'char' in o.bullet)
    bu = `<a:buChar char="${esc(o.bullet.char)}"/>`;
  else if (o.bullet !== undefined)
    bu = `<a:buAutoNum type="${o.bullet.autoNum}"${o.bullet.startAt !== undefined ? ` startAt="${o.bullet.startAt}"` : ''}/>`;
  const pPr =
    o.algn !== undefined || bu !== ''
      ? `<a:pPr${o.algn !== undefined ? ` algn="${o.algn}"` : ''}>${bu}</a:pPr>`
      : '';
  const end = o.endSz !== undefined ? `<a:endParaRPr lang="ja-JP" sz="${o.endSz}"/>` : '';
  return `<a:p>${pPr}${(o.runs ?? []).map(runXml).join('')}${end}</a:p>`;
}

/** 文字の本体（`xdr:txBody`） */
function txBodyXml(s: ShapeSpec): string {
  const bodyPr =
    s.bodyPrInner !== undefined
      ? `<a:bodyPr${s.bodyPr ? ` ${s.bodyPr}` : ''}>${s.bodyPrInner}</a:bodyPr>`
      : `<a:bodyPr${s.bodyPr ? ` ${s.bodyPr}` : ''}/>`;
  return `<xdr:txBody>${bodyPr}<a:lstStyle/>${(s.paras ?? []).map(paraXml).join('')}</xdr:txBody>`;
}

/** `spPr`（xfrm・形・塗り・a:ln）と `xdr:style`。`xfrm` を作るのは属性が指定されたときだけ */
function spPrAndStyle(s: {
  prst?: string;
  avLst?: string;
  custGeom?: string;
  xfrm?: string;
  box?: Box;
  fill?: string;
  ln?: string;
  style?: StyleSpec;
}): string {
  const xfrm =
    s.xfrm !== undefined || s.box !== undefined
      ? `<a:xfrm${s.xfrm !== undefined ? ` ${s.xfrm}` : ''}>${offExt(s.box)}</a:xfrm>`
      : '';
  const geom =
    s.custGeom ??
    `<a:prstGeom prst="${s.prst ?? 'rect'}"><a:avLst>${s.avLst ?? ''}</a:avLst></a:prstGeom>`;
  const ref = (tag: string, attrs: string, clr: string | undefined, dflt: string): string => {
    const c = clr ?? dflt;
    return c === '' ? `<a:${tag}${attrs}/>` : `<a:${tag}${attrs}>${c}</a:${tag}>`;
  };
  const accent1 = '<a:schemeClr val="accent1"/>';
  const style = s.style
    ? `<xdr:style>${ref('lnRef', ` idx="${s.style.ln}"`, s.style.lnClr, accent1)}${ref('fillRef', ` idx="${s.style.fill}"`, s.style.fillClr, accent1)}<a:effectRef idx="0">${accent1}</a:effectRef>${ref('fontRef', ' idx="minor"', s.style.fontClr, '<a:schemeClr val="lt1"/>')}</xdr:style>`
    : '';
  return `<xdr:spPr>${xfrm}${geom}${s.fill ?? ''}${s.ln ?? ''}</xdr:spPr>${style}`;
}

/** 文字・回転・枠線・塗りを指定できる図形（`xdr:sp`） */
export function shapeXml(s: ShapeSpec): string {
  const tl = s.textlink !== undefined ? ` textlink="${esc(s.textlink)}"` : ' textlink=""';
  return `<xdr:sp macro=""${tl}><xdr:nvSpPr>${cNvPr(s.id ?? 3, s.name, undefined, s.hidden)}<xdr:cNvSpPr${s.txBox ? ' txBox="1"' : ''}/></xdr:nvSpPr>${spPrAndStyle(s)}${s.paras !== undefined ? txBodyXml(s) : ''}</xdr:sp>`;
}

/** 自由図形（`a:custGeom`）の指定。各リストは中身の XML をそのまま入れる（省くと空の要素） */
export interface CustGeomSpec {
  /** `a:avLst` の中身（`a:gd` の並び） */
  avLst?: string;
  /** `a:gdLst` の中身（`a:gd` の並び） */
  gdLst?: string;
  /** `a:rect` の属性（例 `l="0" t="0" r="r" b="b"`）。省くと `a:rect` を作らない */
  rect?: string;
  /** `a:pathLst` の中身（`a:path` の並び） */
  pathLst?: string;
}

/** `a:custGeom` 要素全体を組む（`ShapeSpec.custGeom` に渡す） */
export function custGeomXml(c: CustGeomSpec): string {
  const rect = c.rect !== undefined ? `<a:rect ${c.rect}/>` : '';
  return `<a:custGeom><a:avLst>${c.avLst ?? ''}</a:avLst><a:gdLst>${c.gdLst ?? ''}</a:gdLst><a:ahLst/><a:cxnLst/>${rect}<a:pathLst>${c.pathLst ?? ''}</a:pathLst></a:custGeom>`;
}

/** 線・矢印の指定（`xdr:cxnSp`。`asSp` なら線の preset の `xdr:sp`） */
export interface LineShapeSpec {
  name: string;
  id?: number;
  hidden?: boolean;
  /** `a:prstGeom prst`（既定 `straightConnector1`） */
  prst?: string;
  avLst?: string;
  xfrm?: string;
  /** `a:off`・`a:ext`（x, y, cx, cy。EMU）。指定すると `a:xfrm` を作る */
  box?: Box;
  /** `a:ln` 要素全体 */
  ln?: string;
  style?: StyleSpec;
  /** true なら `xdr:cxnSp` でなく `xdr:sp` で作る */
  asSp?: boolean;
}

/** xfrm・a:ln・xdr:style を指定できる線・矢印 */
export function lineShapeXml(s: LineShapeSpec): string {
  const props = spPrAndStyle({ ...s, prst: s.prst ?? 'straightConnector1' });
  if (s.asSp)
    return `<xdr:sp macro="" textlink=""><xdr:nvSpPr>${cNvPr(s.id ?? 4, s.name, undefined, s.hidden)}<xdr:cNvSpPr/></xdr:nvSpPr>${props}</xdr:sp>`;
  return `<xdr:cxnSp macro=""><xdr:nvCxnSpPr>${cNvPr(s.id ?? 4, s.name, undefined, s.hidden)}<xdr:cNvCxnSpPr/></xdr:nvCxnSpPr>${props}</xdr:cxnSp>`;
}

// ---- グラフ（ISSUE-101）。Excel が書く chartN.xml に近づけた合成。指定しない項目は出さない ----

/** 項目（`c:cat`）の指定 */
export type ChartCatSpec =
  | { kind: 'str'; values: (string | null)[]; lit?: boolean }
  | { kind: 'num'; values: (number | null)[]; formatCode?: string; lit?: boolean }
  /** 多段の項目。`levels[0]` が最も内側の段 */
  | { kind: 'multi'; levels: string[][] };

/** 系列 1 つの指定 */
export interface ChartSerSpec {
  /** 系列名（`c:tx` の `c:strRef` の保存値）。省くと `c:tx` を出さない */
  name?: string;
  /** 系列名を `c:tx` の `c:v` で出す */
  nameLit?: boolean;
  /**
   * 値（`c:val`）。位置が点の番号。null は `c:pt` を出さない（疎な点）、文字は数値でない値（`c:v` にそのまま）。
   * 省くと空の配列（`c:ptCount` が 0）
   */
  values?: (number | string | null)[];
  /** `c:ptCount` を values の長さでなくこの値にする */
  ptCount?: number;
  /** `c:numRef` に `c:f` だけで `c:numCache` を出さない（保存値の無い系列） */
  noCache?: boolean;
  /** `c:numRef` でなく `c:numLit` で出す */
  lit?: boolean;
  /** `c:numCache` の `c:formatCode` */
  formatCode?: string;
  cat?: ChartCatSpec;
  /** `c:idx`（既定は系列の位置。null で要素を出さない） */
  idx?: number | null;
  order?: number;
  /** `c:spPr` の中身（`<a:solidFill>…`） */
  spPr?: string;
  /** `c:marker` の中身（`<c:symbol val="none"/>` など） */
  marker?: string;
  /** 点ごとの上書き（`c:dPt`）。spPr は `c:spPr` の中身 */
  dPts?: { idx: number; spPr: string }[];
  /** `c:smooth` の val */
  smooth?: boolean;
  /** 系列の中に足す要素（`c:dLbls`・`c:trendline`）。`c:cat` の前に置く */
  extra?: string;
}

/** 軸 1 つの指定（指定した項目だけ出す） */
export interface ChartAxisSpec {
  delete?: boolean;
  axPos?: string;
  min?: number;
  max?: number;
  orientation?: string;
  logBase?: number;
  majorUnit?: number;
  minorUnit?: number;
  crosses?: string;
  crossesAt?: number;
  crossBetween?: string;
  /** 主・補助の目盛線。true は空の要素、文字は `c:spPr` の中身 */
  majorGridlines?: true | string;
  minorGridlines?: true | string;
  majorTickMark?: string;
  minorTickMark?: string;
  tickLblPos?: string;
  tickLblSkip?: number;
  numFmt?: { code: string; sourceLinked?: boolean };
  /** `c:title` の全体（`chartTitleXml`） */
  title?: string;
  /** 軸の `c:spPr` の中身 */
  spPr?: string;
  /** `c:txPr` の全体（`chartTxPrXml`） */
  txPr?: string;
}

/** グラフ全体の指定 */
export interface ChartSpec {
  /** 種類（既定 `bar`） */
  group?: 'bar' | 'line' | 'pie' | 'doughnut';
  /** 種類の要素名をそのまま指定する（例 `areaChart`・`bar3DChart`。`group` より優先） */
  groupName?: string;
  barDir?: string;
  grouping?: string;
  varyColors?: boolean;
  gapWidth?: number;
  overlap?: number;
  firstSliceAng?: number;
  holeSize?: number;
  /** 折れ線のグループの `c:marker val` */
  groupMarker?: boolean;
  /** 種類の要素の中に足す要素（`c:dLbls`） */
  groupExtra?: string;
  series?: ChartSerSpec[];
  /** 種類の要素を足す（組み合わせの試験）。要素の全体 */
  extraGroups?: string[];
  /** 項目軸の指定（null で出さない。既定は棒・折れ線で `{}`、円・ドーナツで出さない） */
  catAx?: ChartAxisSpec | null;
  /** 数値軸の指定（同上） */
  valAx?: ChartAxisSpec | null;
  /** 項目軸を `c:dateAx` にする */
  dateAx?: boolean;
  /** 数値軸をもう 1 本足す（第 2 軸の試験） */
  extraValAx?: boolean;
  /** `c:title` の全体（`chartTitleXml`） */
  title?: string;
  autoTitleDeleted?: boolean;
  view3D?: boolean;
  /** `c:legend` の全体 */
  legend?: string;
  /** プロットの `c:layout` の中身（`<c:manualLayout>…`） */
  plotLayout?: string;
  /** プロットの面の `c:spPr` の中身 */
  plotSpPr?: string;
  /** グラフの面の `c:spPr` の中身 */
  chartSpPr?: string;
  /** グラフの `c:txPr` の全体（`chartTxPrXml`） */
  txPr?: string;
  /** `c:style` を `mc:AlternateContent` の `mc:Fallback` の中に出す（Excel の書き方。Choice の `c14:style` は 100 を足した値） */
  style?: number;
  /** `c:style` を直接出す */
  styleDirect?: number;
  /** `c:roundedCorners` の `val`（文字そのまま。空文字は `val` の無い要素）。省くと要素を出さない */
  roundedCorners?: string;
  /** `c:dispBlanksAs` の val。省くと `gap`、null で要素を出さない */
  dispBlanksAs?: string | null;
}

/** `c:pt` の並び（疎：null は出さない） */
function chartPts(values: readonly (number | string | null)[]): string {
  return values
    .map((v, i) => (v === null ? '' : `<c:pt idx="${i}"><c:v>${esc(v)}</c:v></c:pt>`))
    .join('');
}

/** `c:cat` の XML */
function chartCatXml(cat: ChartCatSpec): string {
  if (cat.kind === 'multi') {
    const n = Math.max(0, ...cat.levels.map((l) => l.length));
    const lvls = cat.levels.map((l) => `<c:lvl>${chartPts(l)}</c:lvl>`).join('');
    return `<c:cat><c:multiLvlStrRef><c:f>Sheet1!$A$2:$B$${n + 1}</c:f><c:multiLvlStrCache><c:ptCount val="${n}"/>${lvls}</c:multiLvlStrCache></c:multiLvlStrRef></c:cat>`;
  }
  const n = cat.values.length;
  if (cat.kind === 'str') {
    const body = `<c:ptCount val="${n}"/>${chartPts(cat.values)}`;
    return cat.lit
      ? `<c:cat><c:strLit>${body}</c:strLit></c:cat>`
      : `<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$${n + 1}</c:f><c:strCache>${body}</c:strCache></c:strRef></c:cat>`;
  }
  const code = `<c:formatCode>${esc(cat.formatCode ?? 'General')}</c:formatCode>`;
  const body = `${code}<c:ptCount val="${n}"/>${chartPts(cat.values)}`;
  return cat.lit
    ? `<c:cat><c:numLit>${body}</c:numLit></c:cat>`
    : `<c:cat><c:numRef><c:f>Sheet1!$A$2:$A$${n + 1}</c:f><c:numCache>${body}</c:numCache></c:numRef></c:cat>`;
}

/** `c:ser` の XML */
function chartSerXml(s: ChartSerSpec, i: number): string {
  const values = s.values ?? [];
  const tx =
    s.name === undefined
      ? ''
      : s.nameLit
        ? `<c:tx><c:v>${esc(s.name)}</c:v></c:tx>`
        : `<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${esc(s.name)}</c:v></c:pt></c:strCache></c:strRef></c:tx>`;
  const dPts = (s.dPts ?? [])
    .map(
      (d) =>
        `<c:dPt><c:idx val="${d.idx}"/><c:invertIfNegative val="0"/><c:bubble3D val="0"/><c:spPr>${d.spPr}</c:spPr></c:dPt>`,
    )
    .join('');
  const cache = `<c:formatCode>${esc(s.formatCode ?? 'General')}</c:formatCode><c:ptCount val="${s.ptCount ?? values.length}"/>${chartPts(values)}`;
  const val = s.lit
    ? `<c:val><c:numLit>${cache}</c:numLit></c:val>`
    : s.noCache
      ? `<c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val>`
      : `<c:val><c:numRef><c:f>Sheet1!$B$2:$B$${values.length + 1}</c:f><c:numCache>${cache}</c:numCache></c:numRef></c:val>`;
  return (
    `<c:ser>${s.idx === null ? '' : `<c:idx val="${s.idx ?? i}"/>`}` +
    (s.order === undefined ? `<c:order val="${i}"/>` : `<c:order val="${s.order}"/>`) +
    tx +
    (s.spPr === undefined ? '' : `<c:spPr>${s.spPr}</c:spPr>`) +
    (s.marker === undefined ? '' : `<c:marker>${s.marker}</c:marker>`) +
    dPts +
    (s.extra ?? '') +
    (s.cat === undefined ? '' : chartCatXml(s.cat)) +
    val +
    (s.smooth === undefined ? '' : `<c:smooth val="${s.smooth ? 1 : 0}"/>`) +
    `</c:ser>`
  );
}

/** 軸 1 つの XML（`kind` は要素名） */
function chartAxisXml(
  kind: 'catAx' | 'dateAx' | 'valAx',
  id: number,
  crossAx: number,
  a: ChartAxisSpec,
): string {
  const n = (tag: string, v: number | string | undefined) =>
    v === undefined ? '' : `<c:${tag} val="${v}"/>`;
  const grid = (tag: string, g: true | string | undefined) =>
    g === undefined
      ? ''
      : g === true
        ? `<c:${tag}/>`
        : `<c:${tag}><c:spPr>${g}</c:spPr></c:${tag}>`;
  const scaling =
    `<c:scaling>${a.logBase === undefined ? '' : `<c:logBase val="${a.logBase}"/>`}` +
    `<c:orientation val="${a.orientation ?? 'minMax'}"/>` +
    n('max', a.max) +
    n('min', a.min) +
    `</c:scaling>`;
  return (
    `<c:${kind}><c:axId val="${id}"/>${scaling}` +
    (a.delete === undefined ? '' : `<c:delete val="${a.delete ? 1 : 0}"/>`) +
    n('axPos', a.axPos) +
    grid('majorGridlines', a.majorGridlines) +
    grid('minorGridlines', a.minorGridlines) +
    (a.title ?? '') +
    (a.numFmt === undefined
      ? ''
      : `<c:numFmt formatCode="${esc(a.numFmt.code)}" sourceLinked="${a.numFmt.sourceLinked ? 1 : 0}"/>`) +
    n('majorTickMark', a.majorTickMark) +
    n('minorTickMark', a.minorTickMark) +
    n('tickLblPos', a.tickLblPos) +
    (a.spPr === undefined ? '' : `<c:spPr>${a.spPr}</c:spPr>`) +
    (a.txPr ?? '') +
    `<c:crossAx val="${crossAx}"/>` +
    n('crosses', a.crosses) +
    n('crossesAt', a.crossesAt) +
    n('crossBetween', a.crossBetween) +
    n('tickLblSkip', a.tickLblSkip) +
    n('majorUnit', a.majorUnit) +
    n('minorUnit', a.minorUnit) +
    `</c:${kind}>`
  );
}

/** グラフ・軸のタイトル（`c:title`）。`lines` は段落（文字列で 1 段落）、extra は title の中に足す要素（`<c:overlay val="1"/>` など） */
export function chartTitleXml(lines: string[], rPr = '', extra = '', bodyPr = ''): string {
  const paras = lines
    .map(
      (l) =>
        `<a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:rPr lang="ja-JP"${rPr}/><a:t>${esc(l)}</a:t></a:r></a:p>`,
    )
    .join('');
  return `<c:title><c:tx><c:rich><a:bodyPr${bodyPr}/><a:lstStyle/>${paras}</c:rich></c:tx>${extra}</c:title>`;
}

/** `c:txPr`（`a:defRPr` の属性と中身を指定する。`bodyPr` は `a:bodyPr` の属性） */
export function chartTxPrXml(defRPrAttrs = '', defRPrInner = '', bodyPr = ''): string {
  const def =
    defRPrInner === ''
      ? `<a:defRPr${defRPrAttrs}/>`
      : `<a:defRPr${defRPrAttrs}>${defRPrInner}</a:defRPr>`;
  return `<c:txPr><a:bodyPr${bodyPr}/><a:lstStyle/><a:p><a:pPr>${def}</a:pPr><a:endParaRPr lang="ja-JP"/></a:p></c:txPr>`;
}

/** グラフの部品（`chartN.xml` の中身）を組む。指定しない項目は出さない */
export function chartXml(spec: ChartSpec = {}): string {
  const group = spec.group ?? 'bar';
  const name = spec.groupName ?? `${group}Chart`;
  const ser = (spec.series ?? []).map(chartSerXml).join('');
  const val = (tag: string, v: number | string | boolean | undefined) =>
    v === undefined ? '' : `<c:${tag} val="${typeof v === 'boolean' ? (v ? 1 : 0) : v}"/>`;
  const hasAxes = group === 'bar' || group === 'line';
  const axIds = hasAxes ? `<c:axId val="111"/><c:axId val="222"/>` : '';
  let head: string;
  let tail: string;
  if (group === 'bar') {
    head =
      val('barDir', spec.barDir) +
      val('grouping', spec.grouping) +
      val('varyColors', spec.varyColors);
    tail = val('gapWidth', spec.gapWidth) + val('overlap', spec.overlap);
  } else if (group === 'line') {
    head = val('grouping', spec.grouping) + val('varyColors', spec.varyColors);
    tail = val('marker', spec.groupMarker);
  } else {
    head = val('varyColors', spec.varyColors);
    tail = val('firstSliceAng', spec.firstSliceAng) + val('holeSize', spec.holeSize);
  }
  const groupXml = `<c:${name}>${head}${ser}${spec.groupExtra ?? ''}${tail}${axIds}</c:${name}>`;

  const catAx = spec.catAx === undefined ? (hasAxes ? {} : null) : spec.catAx;
  const valAx = spec.valAx === undefined ? (hasAxes ? {} : null) : spec.valAx;
  const axes =
    (catAx === null ? '' : chartAxisXml(spec.dateAx ? 'dateAx' : 'catAx', 111, 222, catAx)) +
    (valAx === null ? '' : chartAxisXml('valAx', 222, 111, valAx)) +
    (spec.extraValAx ? chartAxisXml('valAx', 333, 111, {}) : '');

  const style =
    spec.style === undefined
      ? ''
      : `<mc:AlternateContent><mc:Choice Requires="c14" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"><c14:style val="${100 + spec.style}"/></mc:Choice><mc:Fallback><c:style val="${spec.style}"/></mc:Fallback></mc:AlternateContent>`;
  const blanks =
    spec.dispBlanksAs === null ? '' : `<c:dispBlanksAs val="${spec.dispBlanksAs ?? 'gap'}"/>`;
  return (
    `${HEAD}<c:chartSpace xmlns:c="${C_NS}" xmlns:a="${A_NS}" xmlns:r="${REL}" xmlns:mc="${MC_NS}">` +
    `<c:date1904 val="0"/>${spec.roundedCorners === undefined ? '' : `<c:roundedCorners${spec.roundedCorners === '' ? '' : ` val="${spec.roundedCorners}"`}/>`}${style}${val('style', spec.styleDirect)}` +
    `<c:chart>${spec.title ?? ''}${val('autoTitleDeleted', spec.autoTitleDeleted)}` +
    (spec.view3D ? `<c:view3D><c:rotX val="15"/></c:view3D>` : '') +
    `<c:plotArea><c:layout>${spec.plotLayout ?? ''}</c:layout>${groupXml}${(spec.extraGroups ?? []).join('')}${axes}` +
    (spec.plotSpPr === undefined ? '' : `<c:spPr>${spec.plotSpPr}</c:spPr>`) +
    `</c:plotArea>${spec.legend ?? ''}<c:plotVisOnly val="1"/>${blanks}</c:chart>` +
    (spec.chartSpPr === undefined ? '' : `<c:spPr>${spec.chartSpPr}</c:spPr>`) +
    (spec.txPr ?? '') +
    `</c:chartSpace>`
  );
}
