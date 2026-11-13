// 正: InterHub docs/plan/adr/0025-xlsx-own-reader.md
/**
 * 数値の表示形式を当てて、Excel で見える文字にする。
 * 書式コードの解釈は `ssf` 0.11.2 に任せ、日本語環境に合わせた組み込みの書式の対応表だけ自前で持つ。
 * 和暦（`ggge`）は ssf にも無いので、1 次では対応しない（ADR-0025「悪い点」）。
 * XSheet: 和暦は format-ja.ts で ssf の手前で置き換える（ADR-0007 §6）。
 */
import * as ssfModule from 'ssf';
import { JA_BUILTIN_DATES, withJaTokens, type BuiltinDates } from './format-ja'; // XSheet: 曜日・和暦の口（ADR-0007 §6）
import { MAX_FORMAT_CODE_LENGTH } from './limits';
import type { Styles } from './styles';

/** ssf の `format` と `is_date`（型定義は `ssf/types`。CommonJS の default 経由でも名前付きでも取れるようにする） */
interface SsfApi {
  format(fmt: string | number, val: unknown, opts?: object): string;
  is_date(fmt: string): boolean;
}
// XSheet: 曜日（aaa・aaaa）を ssf の手前で置き換える口（format-ja.ts・ADR-0007 §6）
const ssf: SsfApi = withJaTokens(
  (ssfModule as unknown as { default?: SsfApi }).default ?? (ssfModule as unknown as SsfApi),
);

/**
 * 組み込みの表示形式（numFmtId → 書式コード）。
 * ssf の既定は米国式（14 は `m/d/yy`・22 は `m/d/yy h:mm`）なので、日本語環境の Excel の表示に合わせて上書きする。
 * 27〜36・50〜58 は日本語環境では和暦だが、1 次は西暦で代用する（和暦は範囲外）。
 */
const BUILTIN_FORMATS: ReadonlyMap<number, string> = new Map<number, string>([
  [1, '0'],
  [2, '0.00'],
  [3, '#,##0'],
  [4, '#,##0.00'],
  [9, '0%'],
  [10, '0.00%'],
  [11, '0.00E+00'],
  [12, '# ?/?'],
  [13, '# ??/??'],
  [14, 'yyyy/m/d'],
  [15, 'd-mmm-yy'],
  [16, 'd-mmm'],
  [17, 'mmm-yy'],
  [18, 'h:mm AM/PM'],
  [19, 'h:mm:ss AM/PM'],
  [20, 'h:mm'],
  [21, 'h:mm:ss'],
  [22, 'yyyy/m/d h:mm'],
  [27, 'yyyy/m/d'],
  [28, 'yyyy/m/d'],
  [29, 'yyyy/m/d'],
  [30, 'yyyy/m/d'],
  [31, 'yyyy/m/d'],
  [32, 'h"時"mm"分"'],
  [33, 'h"時"mm"分"ss"秒"'],
  [34, 'AM/PM h"時"mm"分"'],
  [35, 'AM/PM h"時"mm"分"ss"秒"'],
  [36, 'yyyy/m/d'],
  [37, '#,##0_);(#,##0)'],
  [38, '#,##0_);[Red](#,##0)'],
  [39, '#,##0.00_);(#,##0.00)'],
  [40, '#,##0.00_);[Red](#,##0.00)'],
  [45, 'mm:ss'],
  [46, '[h]:mm:ss'],
  [47, 'mm:ss.0'],
  [48, '##0.0E+0'],
  [49, '@'],
  [50, 'yyyy/m/d'],
  [51, 'yyyy/m/d'],
  [52, 'yyyy/m/d'],
  [53, 'yyyy/m/d'],
  [54, 'yyyy/m/d'],
  [55, 'yyyy/m/d'],
  [56, 'yyyy/m/d'],
  [57, 'yyyy/m/d'],
  [58, 'yyyy/m/d'],
]);

/** 組み込みの表示形式の書式コード。無い番号（0 を含む）は General */
export function builtinFormatCode(id: number): string {
  return BUILTIN_FORMATS.get(id) ?? 'General';
}

/** 書式コードが日付・時刻か */
function isDateCode(code: string): boolean {
  try {
    return ssf.is_date(code);
  } catch {
    return false;
  }
}

/** General の表示。整数は桁を落とさず、それ以外は ssf（Excel と同じ有効桁）に任せる */
function formatGeneral(v: number, date1904: boolean): string {
  if (Number.isInteger(v) && Math.abs(v) < 1e11) return String(v);
  try {
    return ssf.format('General', v, { date1904 });
  } catch {
    return String(v);
  }
}

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
export function createFormatter(
  styles: Styles,
  date1904: boolean,
  // XSheet: 組み込みの日付の書式 27〜36・50〜58 を日本語版 Excel の表で引く指定（ADR-0007 §6）
  builtinDates: BuiltinDates = 'default',
): CellFormatter {
  const codes: string[] = styles.xfNumFmtIds.map((id) => {
    const code =
      styles.customFormats.get(id) ??
      (builtinDates === 'japanese-era' ? JA_BUILTIN_DATES.get(id) : undefined) ??
      builtinFormatCode(id);
    return code.length > MAX_FORMAT_CODE_LENGTH ? 'General' : code;
  });
  const codeOf = (xf: number) => codes[xf] ?? 'General';
  return {
    number(value, xf) {
      const code = codeOf(xf);
      // 書式「@」（文字列）の数値は、入力したままの数字で見せる
      if (code === 'General' || code === '@') return formatGeneral(value, date1904);
      try {
        return ssf.format(code, value, { date1904 });
      } catch {
        return formatGeneral(value, date1904);
      }
    },
    date(serial, xf, hasTime) {
      let code = codeOf(xf);
      if (!isDateCode(code)) code = hasTime ? 'yyyy/m/d h:mm:ss' : 'yyyy/m/d';
      try {
        return ssf.format(code, serial, { date1904 });
      } catch {
        return formatGeneral(serial, date1904);
      }
    },
  };
}

/**
 * 書式コードを直接当てて数値を文字にする（グラフの項目・軸の文字用。セルの xf を介さない）。
 * 書式コードは非信頼のファイルから来るため、長さが上限（MAX_FORMAT_CODE_LENGTH）を超えるものは ssf に渡さず General として値だけを出す（ReDoS 対策）。
 * `General`・`@`・解釈できない書式は General の表示（壊れずに数字を見せる）。
 */
export function formatNumberByCode(code: string, value: number, date1904: boolean): string {
  if (code.length > MAX_FORMAT_CODE_LENGTH || code === 'General' || code === '@')
    return formatGeneral(value, date1904);
  try {
    return ssf.format(code, value, { date1904 });
  } catch {
    return formatGeneral(value, date1904);
  }
}

/**
 * ISO 8601 の日付・日時（`2023-03-15`・`2023-03-15T18:00:00`）を Excel のシリアル値にする。
 * 読めなければ null。時刻の無い日付は hasTime=false。タイムゾーンの印は無視する。
 */
export function isoToSerial(
  iso: string,
  date1904: boolean,
): { serial: number; hasTime: boolean } | null {
  const m = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2})(?:\.\d+)?)?)?/.exec(
    iso.trim(),
  );
  if (!m) return null;
  const days = Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3])) / 86_400_000;
  if (!Number.isFinite(days)) return null;
  const hasTime = m[4] !== undefined;
  const frac = hasTime ? (Number(m[4]) * 3600 + Number(m[5]) * 60 + Number(m[6] ?? 0)) / 86_400 : 0;
  // 1970-01-01 は 1900 年起点で 25569、1904 年起点で 24107（Excel の 1900-02-29 は実在しない日だが、3 月以降は 25569 で合う）
  let serial = days + (date1904 ? 24107 : 25569);
  if (!date1904 && serial < 61) serial -= 1;
  return { serial: serial + frac, hasTime };
}
