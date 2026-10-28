// 正: InterHub docs/plan/adr/0025-xlsx-own-reader.md
/**
 * workbook.xml の読み取り：シートの一覧（workbook の順）・非表示・1904 年起点・各部品の場所。
 * シートの順は非表示も数える（1 始まり・非表示を含む）。
 */
import type { XlsxLimits } from './limits';
import type { XlsxPackage } from './open';
import { type Attrs, NS, attrByLocalName, looksStrict, parseAttrs } from './xml';
import { XlsxError } from './types';

/** workbook の中の 1 シート */
export interface WorkbookSheet {
  /** 1 始まり・workbook の順 */
  index: number;
  name: string;
  hidden: boolean;
  /** zip の中のパス。rels に見つからなければ null */
  path: string | null;
  /** 普通のワークシートか（グラフシートなどは false） */
  isWorksheet: boolean;
}

/** 読んだ workbook */
export interface Workbook {
  sheets: WorkbookSheet[];
  date1904: boolean;
  /** sharedStrings.xml のパス。無ければ null */
  sharedStringsPath: string | null;
  /** styles.xml のパス。無ければ null */
  stylesPath: string | null;
  /** テーマ（theme1.xml）のパス。rels に無ければ null（既定の 12 色を使う） */
  themePath: string | null;
}

/** 関係（rels）1 件 */
export interface Relationship {
  type: string;
  target: string;
  /** `TargetMode="External"`（zip の外を指す。読み取りでは辿らない） */
  external: boolean;
}

/** `rels` を Id → 関係 の Map にする（キーを Map に入れて `__proto__` を防ぐ） */
export function parseRels(xml: string): Map<string, Relationship> {
  const map = new Map<string, Relationship>();
  for (const m of xml.matchAll(new RegExp(String.raw`<${NS}Relationship\b[^<>]*>`, 'g'))) {
    const a = parseAttrs(m[0]);
    if (a.Id !== undefined && a.Target !== undefined)
      map.set(a.Id, {
        type: a.Type ?? '',
        target: a.Target,
        external: a.TargetMode?.toLowerCase() === 'external',
      });
  }
  return map;
}

/** 関係の Target を zip のパスにする。`/` 始まりは絶対、それ以外は baseDir からの相対（`..` を解く） */
export function resolveTarget(baseDir: string, target: string): string {
  const full = target.startsWith('/') ? target.slice(1) : baseDir + target;
  const out: string[] = [];
  for (const seg of full.split('/')) {
    if (seg === '' || seg === '.') continue;
    if (seg === '..') out.pop();
    else out.push(seg);
  }
  return out.join('/');
}

/** パスのフォルダ部分（末尾の / を含む。無ければ空） */
export function dirOf(path: string): string {
  const i = path.lastIndexOf('/');
  return i < 0 ? '' : path.slice(0, i + 1);
}

/** パスのファイル名の部分 */
export function baseOf(path: string): string {
  const i = path.lastIndexOf('/');
  return i < 0 ? path : path.slice(i + 1);
}

/** 型（URL の最後の語）が一致するか */
export const typeIs = (type: string, last: string) => type === last || type.endsWith('/' + last);

/**
 * workbook.xml を読む。
 * - 本体の場所は `_rels/.rels` の officeDocument から引く（無ければ `xl/workbook.xml`）。
 * - 読めない形式（Strict・xlsb）は unsupported。workbook が無ければ broken。
 */
export async function loadWorkbook(pkg: XlsxPackage, limits: XlsxLimits): Promise<Workbook> {
  const small = limits.maxSmallPartBytes;
  let wbPath = 'xl/workbook.xml';
  const rootRels = await pkg.readText('_rels/.rels', small);
  if (rootRels !== null) {
    for (const rel of parseRels(rootRels).values()) {
      if (typeIs(rel.type, 'officeDocument')) wbPath = resolveTarget('', rel.target);
    }
  }
  const wbXml = await pkg.readText(wbPath, small);
  if (wbXml === null) {
    throw new XlsxError(pkg.has('xl/workbook.bin') ? 'unsupported' : 'broken', 'workbook');
  }
  // ルート要素の名前空間で Strict を見分ける（先頭の数百文字に宣言がある）
  if (looksStrict(wbXml)) throw new XlsxError('unsupported', 'strict');

  const wbDir = dirOf(wbPath);
  const relsXml = await pkg.readText(`${wbDir}_rels/${baseOf(wbPath)}.rels`, small);
  const rels = relsXml === null ? new Map<string, Relationship>() : parseRels(relsXml);

  const date1904 = (() => {
    const m = new RegExp(String.raw`<${NS}workbookPr\b[^<>]*>`).exec(wbXml);
    if (!m) return false;
    const v = parseAttrs(m[0]).date1904;
    return v === '1' || v === 'true';
  })();

  // workbook.xml の <sheet> の並び順がそのまま「N 番目のシート」（非表示を含む）
  const sheets: WorkbookSheet[] = [];
  for (const m of wbXml.matchAll(new RegExp(String.raw`<${NS}sheet\b[^<>]*>`, 'g'))) {
    if (sheets.length >= limits.maxSheets) throw new XlsxError('too-large', 'sheets');
    const a: Attrs = parseAttrs(m[0]);
    const rel = rels.get(attrByLocalName(a, 'id') ?? '');
    sheets.push({
      index: sheets.length + 1,
      name: a.name ?? '',
      hidden: a.state === 'hidden' || a.state === 'veryHidden',
      path: rel ? resolveTarget(wbDir, rel.target) : null,
      isWorksheet: rel ? typeIs(rel.type, 'worksheet') : true,
    });
  }

  let sharedStringsPath: string | null = null;
  let stylesPath: string | null = null;
  let themePath: string | null = null;
  for (const rel of rels.values()) {
    if (typeIs(rel.type, 'sharedStrings')) sharedStringsPath = resolveTarget(wbDir, rel.target);
    else if (typeIs(rel.type, 'styles')) stylesPath = resolveTarget(wbDir, rel.target);
    else if (typeIs(rel.type, 'theme') && !rel.external)
      themePath = resolveTarget(wbDir, rel.target);
  }
  return { sheets, date1904, sharedStringsPath, stylesPath, themePath };
}
