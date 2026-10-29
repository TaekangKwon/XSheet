import type { Styles } from './styles';
/** 組み込みの表示形式の書式コード。無い番号（0 を含む）は General */
export declare function builtinFormatCode(id: number): string;
/** セルの値に表示形式を当てる部品 */
export interface CellFormatter {
    /** 数値のセル。xf はセルの `s` 属性（cellXfs の位置） */
    number(value: number, xf: number): string;
    /** 日付のセル（`t="d"`）。serial は Excel のシリアル値、hasTime は時刻を含むか。書式が日付でなければ既定の日付書式 */
    date(serial: number, xf: number, hasTime: boolean): string;
}
/**
 * 表示形式を当てる部品を作る。
 * 書式コードは cellXfs の位置ごとに 1 回だけ引く（独自の書式 → 組み込みの対応表 → General の順）。
 * ssf が解釈できない書式は General の表示に落とす（壊れずに数字を見せる）。
 * 書式コードは非信頼の styles.xml から来るため、長さが上限（MAX_FORMAT_CODE_LENGTH）を超えるものは
 * ssf に渡さず General として値だけを出す（ReDoS 対策）。
 */
export declare function createFormatter(styles: Styles, date1904: boolean): CellFormatter;
/**
 * ISO 8601 の日付・日時（`2023-03-15`・`2023-03-15T18:00:00`）を Excel のシリアル値にする。
 * 読めなければ null。時刻の無い日付は hasTime=false。タイムゾーンの印は無視する。
 */
export declare function isoToSerial(iso: string, date1904: boolean): {
    serial: number;
    hasTime: boolean;
} | null;
