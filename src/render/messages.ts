// 正: docs/02_architecture.md §7
/**
 * render の文言（`RenderOptions.messages`）。InterHub の `t('spreadsheet.…')` の代わり。
 * 既定は日本語。英語も同梱し、使う側は一部のキーだけを渡して差し替えられる。
 * 差し込み（`{detail}`・`{format}`）は render が行う。文言は `textContent`・`title` にだけ入れる。
 */

/** 文言の表（23 キー。元は InterHub `locales/ja.ts` の `spreadsheet.*` と `data.noData`） */
export interface XSheetMessages {
  /** 図形の形の表など、描画に要る部品を読み込めない */
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
export const MESSAGES_JA: Readonly<XSheetMessages> = Object.freeze({
  libLoadFailed: 'Excel 表示ライブラリ（xlsx）を読み込めませんでした。',
  parseFailed: 'Excel ファイルの解析に失敗しました: {detail}',
  noSheets: '（シートがありません）',
  noData: '（データがありません）',
  encrypted: 'このファイルは暗号化されているか、旧形式（.xls）の可能性があるため表示できません。',
  tooLarge: 'Excel ファイルが大きすぎるため表示できません（内部の上限を超えています）。',
  unsupported: 'この形式の Excel ファイル（Strict Open XML・.xlsb など）には対応していません。',
  hiddenSheet: '非表示のシート',
  tabsPrev: '前のシート',
  tabsNext: '次のシート',
  chart: 'グラフ',
  smartart: 'SmartArt',
  shape: '図形',
  textbox: 'テキストボックス',
  group: 'グループ',
  imageFormat: '画像（{format}）',
  imageBroken: '画像（表示できない）',
  noticeTooLarge: '図のデータが大きすぎるため、図は表示していません（表は表示しています）。',
  noticeBroken: '図を読み取れなかったため、図は表示していません（表は表示しています）。',
  noticeTooMany: '図が多いため、先頭の 1,000 個だけ表示しています。',
  noticeTextTruncated: '図形の文字が多いため、一部を省略して表示しています。',
  noticeShapeSimplified:
    '図形の形を描ききれないものがあるため、一部の図形は簡略化して表示しています。',
  noticeOverBudget: '画像の合計が大きいため、一部の画像は枠だけで表示しています。',
});

/** 英語 */
export const MESSAGES_EN: Readonly<XSheetMessages> = Object.freeze({
  libLoadFailed: 'Failed to load the Excel viewer library (xlsx).',
  parseFailed: 'Failed to parse the Excel file: {detail}',
  noSheets: '(No sheets)',
  noData: '(No data)',
  encrypted: 'This file is encrypted, or may be in the old .xls format, so it cannot be displayed.',
  tooLarge: 'The Excel file is too large to display (it exceeds an internal limit).',
  unsupported: 'This Excel format (Strict Open XML, .xlsb, etc.) is not supported.',
  hiddenSheet: 'Hidden sheet',
  tabsPrev: 'Previous sheets',
  tabsNext: 'Next sheets',
  chart: 'Chart',
  smartart: 'SmartArt',
  shape: 'Shape',
  textbox: 'Text box',
  group: 'Group',
  imageFormat: 'Image ({format})',
  imageBroken: 'Image (cannot be displayed)',
  noticeTooLarge: 'The drawing data is too large, so drawings are not shown (the table is shown).',
  noticeBroken: 'The drawings could not be read, so they are not shown (the table is shown).',
  noticeTooMany: 'Too many drawings: only the first 1,000 are shown.',
  noticeTextTruncated: 'The shapes contain a lot of text, so some of it is omitted.',
  noticeShapeSimplified:
    'Some shapes are too complex to draw exactly, so they are shown in a simplified form.',
  noticeOverBudget: 'The images are too large in total, so some images are shown as frames only.',
});

/**
 * 使う文言の表を作る：既定（日本語）に、渡された分だけを上書きする。
 * 文字列でない値・表に無いキーは無視する（使う側の誤りで表示が壊れないように）。
 */
export function resolveMessages(
  messages?: Partial<XSheetMessages>,
  base: Readonly<XSheetMessages> = MESSAGES_JA,
): Readonly<XSheetMessages> {
  if (!messages) return base;
  const out: XSheetMessages = { ...base };
  for (const key of Object.keys(base) as (keyof XSheetMessages)[]) {
    if (!Object.prototype.hasOwnProperty.call(messages, key)) continue;
    const v = messages[key];
    if (typeof v === 'string') out[key] = v;
  }
  return Object.freeze(out);
}

/** `{name}` を差し込む。値の無い `{name}` はそのまま残す（InterHub の `t()` と同じ） */
export function formatMessage(
  template: string,
  params: Readonly<Record<string, string | number>> = {},
): string {
  return template.replace(/\{(\w+)\}/g, (all, name: string) =>
    Object.prototype.hasOwnProperty.call(params, name) ? String(params[name]) : all,
  );
}
