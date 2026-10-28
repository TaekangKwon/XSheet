// build-preset-geom.mjs の型宣言（テストから import するため。本体は依存なしの Node スクリプト）

/** 表に入れない線の preset（10 種） */
export const EXCLUDED_PRESETS: readonly string[];

/** 表の形式の版 */
export const TABLE_VERSION: number;

/** 簡易 XML の木 */
export interface XmlNode {
  name: string;
  attrs: Record<string, string>;
  children: XmlNode[];
}

/** XML を最小の木にする（コメント・処理命令・テキストは捨てる） */
export function parseXml(xml: string): XmlNode;

/** 原本の XML と NOTICE の本文から、形の表の JSON 文字列（末尾改行つき・決定的）を作る */
export function convertPresetGeom(
  xml: string,
  notice: string,
  options?: { allow?: string[] | null },
): string;
