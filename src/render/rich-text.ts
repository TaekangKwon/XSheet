// 正: docs/adr/0009-rich-text-runs.md §2・§3
/**
 * リッチテキストの区切りを描く（M7-17）。
 * - 区切りの文字を `span.xs-runs` 1 つで包む（`.xs-cell`・`.xs-text` は flex なので、包まずに分けると各区切りが別の
 *   flex 項目になり、端の空白・折り返しが崩れる。`highlight.ts` の `span.xs-hits` と同じ）。
 * - `font` の無い区切りはテキストノード、ある区切りは span に CSSOM で書式を当てる。
 *   値は整数・真偽から組んだ固定の語だけ（ファイルの文字列は入れない）。文字は `createTextNode` だけで入れる。
 */
import { type CellFormat, type RunFont, type TextRun, isPaperFill } from '../core/index';

/** 文字の [from, to)（`CellView.text` 上の位置）を、区切りの書式で分けたノードにする */
export type RunSlicer = (from: number, to: number) => Node[];

/** 区切りの大きさの比の範囲（`format-css.ts` と同じ） */
const MIN_RATIO = 0.5;
const MAX_RATIO = 4;

/** 色の整数 → `#rrggbb` */
function hex(rgb: number): string {
  return `#${(rgb & 0xffffff).toString(16).padStart(6, '0')}`;
}

/**
 * 区切りの span に書式を当てる。区切りはセルの書式を継がない（ADR-0009 §2）。
 * 色が無い区切りは、セルの書式に色があるときだけ自動の色（塗りの上は黒・ほかは画面の文字色）に戻す。
 */
function styleRun(span: HTMLElement, font: RunFont, cell: CellFormat | undefined): void {
  const s = span.style;
  if (font.color !== undefined) {
    s.setProperty('color', hex(font.color.rgb));
  } else if (cell?.color !== undefined) {
    s.setProperty(
      'color',
      cell.fill !== undefined && !isPaperFill(cell.fill) ? '#000000' : 'var(--xs-fg)',
    );
  }
  s.setProperty('font-weight', font.bold ? '700' : '400');
  s.setProperty('font-style', font.italic ? 'italic' : 'normal');
  const lines: string[] = [];
  if (font.underline === 'single' || font.underline === 'double') lines.push('underline');
  if (font.strike) lines.push('line-through');
  if (lines.length > 0) {
    s.setProperty('text-decoration-line', lines.join(' '));
    if (font.underline === 'double') s.setProperty('text-decoration-style', 'double');
  }
  const run = Math.min(Math.max(font.sizeRatio ?? 1, MIN_RATIO), MAX_RATIO);
  const base = cell?.sizeRatio ?? 1;
  const ratio = base > 0 && Number.isFinite(base) ? run / base : 1;
  if (Number.isFinite(ratio) && ratio !== 1)
    s.setProperty('font-size', `${Number(ratio.toFixed(3))}em`);
}

/**
 * 区切りで分けるノードの作り手。
 * @param text セルの文字（`CellView.text`。区切りの連結と同じ）
 * @param runs `CellView.runs`
 * @param cell セルの書式（区切りの大きさの比の基準・自動の色）
 */
export function runSlicer(
  text: string,
  runs: readonly TextRun[],
  cell: CellFormat | undefined,
): RunSlicer {
  return (from, to) => {
    const nodes: Node[] = [];
    let pos = 0;
    for (const run of runs) {
      const end = pos + run.text.length;
      const a = Math.max(from, pos);
      const b = Math.min(to, end);
      if (a < b) {
        const piece = text.slice(a, b);
        if (run.font === undefined) {
          nodes.push(document.createTextNode(piece));
        } else {
          const span = document.createElement('span');
          span.textContent = piece;
          styleRun(span, run.font, cell);
          nodes.push(span);
        }
      }
      if (end >= to) break;
      pos = end;
    }
    return nodes;
  };
}

/**
 * セルの入れ物の文字を区切りで描く（入れ物・はみ出しの span を作った後、検索の一致を付ける前に呼ぶ）。
 * @param box セルの入れ物（文字は直下のテキストか、はみ出しなら `span.xs-text` の中）
 * @param shown 表示した文字（文字列なので `text` と同じ。違えば描かない）
 * @returns 区切りで描いたときの作り手（検索の一致の片の中も区切りで分けるため）。描かなければ undefined
 */
export function applyRuns(
  box: HTMLElement,
  text: string,
  shown: string,
  runs: readonly TextRun[] | undefined,
  cell: CellFormat | undefined,
): RunSlicer | undefined {
  if (runs === undefined || runs.length === 0 || shown !== text) return undefined;
  const slice = runSlicer(text, runs, cell);
  const wrap = document.createElement('span');
  wrap.className = 'xs-runs';
  wrap.append(...slice(0, text.length));
  const target = box.querySelector<HTMLElement>(':scope > .xs-text') ?? box;
  target.replaceChildren(wrap);
  return slice;
}
