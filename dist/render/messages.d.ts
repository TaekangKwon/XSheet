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
export declare const MESSAGES_JA: Readonly<XSheetMessages>;
/** 英語 */
export declare const MESSAGES_EN: Readonly<XSheetMessages>;
/** 韓国語（M6-2・ADR-0003 Q4） */
export declare const MESSAGES_KO: Readonly<XSheetMessages>;
/**
 * 使う文言の表を作る：既定（日本語）に、渡された分だけを上書きする。
 * 文字列でない値・表に無いキーは無視する（使う側の誤りで表示が壊れないように）。
 */
export declare function resolveMessages(messages?: Partial<XSheetMessages>, base?: Readonly<XSheetMessages>): Readonly<XSheetMessages>;
/** `{name}` を差し込む。値の無い `{name}` はそのまま残す（InterHub の `t()` と同じ） */
export declare function formatMessage(template: string, params?: Readonly<Record<string, string | number>>): string;
