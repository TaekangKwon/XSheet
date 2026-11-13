// 正: docs/adr/0007-moobius-viewer-api.md §2
/**
 * 検索の一致のハイライト（M7-5・Moobius の依頼 R2）。
 * - 一致の一覧（`findText` の `matches`）を「行 → 列 → 範囲」の 2 段の表にして表示が持ち、
 *   表を描くたびに描く行・列の一致だけを付け直す（仮想スクロールで描き直しても残る）。
 * - セルの文字を一致の範囲で分け、**分けた全体を `span.xs-hits` 1 つで包む**。`.xs-cell`・`.xs-text` は flex なので、
 *   包まずに分けると各片が別の flex 項目になり、片の端の空白が消える。
 * - 一致の片は `span.xs-hit`（`data-xs-hit` に一覧の番号）。今のヒットは `is-current`。色は CSS 変数
 *   `--xs-hit-bg`・`--xs-hit-current-bg` で、背景だけを付ける（行の高さ・列の幅を変えない）。
 * - 自動の色の文字が塗りをまたぐはみ出し（`background-clip: text` で描く）は、片に `is-ink`（`--xs-fg` の文字の色）を付ける。
 *   不透明なハイライトが、親の背景で描く文字を隠さないようにするため。
 * - `###` で出すセル（数値が列幅に収まらない）は、表示の文字と一致の位置が合わないので、セル全体を光らせる。
 * 文字は `textContent`（テキストノード）だけで入れる（セル値は信頼できない入力）。
 */
import type { TextMatch } from '../core/index';

/** 1 セルの中の一致 */
export interface CellHit {
  /** `CellView.text` 上の位置（UTF-16） */
  start: number;
  /** 終わり（含まない） */
  end: number;
  /** 一致の一覧（`setHighlights` に渡した配列）の番号 */
  index: number;
}

/** 表を描くときに引くハイライト */
export interface GridHits {
  /** 今のヒットの番号（無ければ -1） */
  readonly current: number;
  /** 行 r・列 c（1 始まり）の一致（無ければ undefined） */
  at(r: number, c: number): readonly CellHit[] | undefined;
}

/**
 * 一致の一覧から、表を描くときに引く表を作る。null・空の配列は null（ハイライトなし）。
 * 番号の正しくない一致（整数でない行・列、範囲の逆転）は飛ばす。同じセルの一致は始まりの順に並べる。
 */
export function createGridHits(
  matches: readonly TextMatch[] | null | undefined,
  current = -1,
): GridHits | null {
  if (!matches || matches.length === 0) {
    return null;
  }
  const rows = new Map<number, Map<number, CellHit[]>>();
  matches.forEach((m, index) => {
    if (
      !Number.isInteger(m.r) ||
      !Number.isInteger(m.c) ||
      !Number.isInteger(m.start) ||
      !Number.isInteger(m.end) ||
      m.r < 1 ||
      m.c < 1 ||
      m.start < 0 ||
      m.end <= m.start
    ) {
      return;
    }
    let cols = rows.get(m.r);
    if (!cols) {
      cols = new Map();
      rows.set(m.r, cols);
    }
    let list = cols.get(m.c);
    if (!list) {
      list = [];
      cols.set(m.c, list);
    }
    list.push({ start: m.start, end: m.end, index });
  });
  for (const cols of rows.values()) {
    for (const list of cols.values()) {
      list.sort((a, b) => a.start - b.start || a.index - b.index);
    }
  }
  return { current, at: (r, c) => rows.get(r)?.get(c) };
}

/** 一致の片を付ける（番号と今のヒットの印） */
function markHit(el: HTMLElement, hits: readonly CellHit[], current: number): void {
  el.classList.add('xs-hit');
  el.dataset.xsHit = String(hits[0]!.index);
  if (current >= 0 && hits.some((h) => h.index === current)) {
    el.classList.add('is-current');
  }
}

/**
 * セルの入れ物（`div.xs-cell`）の文字に一致を付ける。`grid-table.ts` が入れ物を作った後に呼ぶ。
 * @param box セルの入れ物（文字は直下のテキストか、はみ出しなら `span.xs-text` の中）
 * @param text セルの文字（`CellView.text`）
 * @param shown 表示した文字（収まらない数値は `###`）
 * @param hits このセルの一致（無ければ何もしない）
 * @param current 今のヒットの番号
 */
export function decorateCell(
  box: HTMLElement,
  text: string,
  shown: string,
  hits: readonly CellHit[] | undefined,
  current: number,
): void {
  if (!hits || hits.length === 0) {
    return;
  }
  if (shown !== text) {
    markHit(box, hits, current); // ###：セル全体
    return;
  }
  const label = box.querySelector<HTMLElement>(':scope > .xs-text');
  const target = label ?? box;
  // 塗りをまたぐ自動の色の文字（inkSpill の background-clip: text）は、片に文字の色を付ける
  const ink = label !== null && label.style.getPropertyValue('background-clip') === 'text';
  const wrap = document.createElement('span');
  wrap.className = 'xs-hits';
  let pos = 0;
  let i = 0;
  while (i < hits.length) {
    const start = Math.max(pos, Math.min(hits[i]!.start, text.length));
    // 重なる一致は 1 つの片にまとめる（番号は先頭の一致）
    let end = Math.min(hits[i]!.end, text.length);
    const group = [hits[i]!];
    i++;
    while (i < hits.length && hits[i]!.start < end) {
      end = Math.max(end, Math.min(hits[i]!.end, text.length));
      group.push(hits[i]!);
      i++;
    }
    if (end <= start) {
      continue;
    }
    if (start > pos) {
      wrap.append(document.createTextNode(text.slice(pos, start)));
    }
    const piece = document.createElement('span');
    piece.textContent = text.slice(start, end);
    markHit(piece, group, current);
    if (ink) {
      piece.classList.add('is-ink');
    }
    wrap.append(piece);
    pos = end;
  }
  if (pos === 0) {
    return; // 文字の範囲に入る一致が無い
  }
  if (pos < text.length) {
    wrap.append(document.createTextNode(text.slice(pos)));
  }
  target.replaceChildren(wrap);
}
