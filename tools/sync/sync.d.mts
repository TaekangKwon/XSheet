// sync.mjs の型宣言（テストから import するため。本体は依存なしの Node スクリプト）
import type { Buffer } from 'node:buffer';

/** XSheet のリポジトリの根 */
export const REPO: string;
/** 対応表（interhub.json）の場所 */
export const CONFIG_PATH: string;
/** 呼んでよい git のサブコマンド（読むだけ） */
export const READ_ONLY_GIT: readonly string[];

/** 読むだけの git の引数を受け取って結果（標準出力）を返す関数 */
export type GitRunner = (args: string[]) => Buffer;

/** 対応表の行き先の種類 */
export type MapKind = 'file' | 'rule' | 'manual' | 'excluded' | 'unmapped';

export interface MapResult {
  kind: MapKind;
  /** XSheet での行き先（file・rule だけ） */
  dest?: string;
  /** 手作業のときの写した先の候補 */
  hint?: string[];
  note?: string;
}

/** compileConfig が作る対応表（中の形は sync.mjs の都合。テストは関数に渡すだけ） */
export interface SyncConfig {
  baseCommit: string;
  defaultDir: string;
  envVar: string;
  root: string;
  watch: string[];
  files: Map<string, string>;
  manualTransforms: { id: string; note: string }[];
  [key: string]: unknown;
}

export interface FileChange {
  status: string;
  path: string;
  oldPath?: string;
  score?: number;
  added: number;
  deleted: number;
  binary: boolean;
  kind: MapKind;
  dest?: string;
  hint?: string[];
  note?: string;
  destExists?: boolean;
}

export interface CommitChange {
  hash: string;
  short: string;
  date: string;
  subject: string;
  files: FileChange[];
}

export interface SyncStatus {
  root: string;
  base: string;
  head: string;
  allCount: number;
  mergeCount: number;
  commits: CommitChange[];
  files: FileChange[];
}

export interface PatchFile {
  status: string;
  path: string;
  oldPath: string | null;
  dest: string;
  kind: MapKind;
}

export interface SkippedFile {
  status: string;
  path: string;
  kind: MapKind;
  hint?: string[];
  note?: string;
}

export interface RewriteResult {
  patch: Buffer;
  manual: Buffer;
  files: PatchFile[];
  skipped: SkippedFile[];
  warnings: string[];
}

export interface PatchResult extends RewriteResult {
  base: string;
  head: string;
  file: string;
  manualFile: string | null;
}

/** 読むだけの git の引数か確かめる。違えば例外 */
export function assertReadOnlyGitArgs(args: unknown): void;
/** InterHub（dir）の git を読むだけで呼ぶ関数を作る */
export function createGitRunner(
  dir: string,
  options?: { exec?: (file: string, args: string[], options: object) => Buffer },
): GitRunner;
/** `<a>..<b>`・`<a>..`・`<a>` を [a, b] にする */
export function parseRange(range: string): [string, string];
/** glob を拾う組つきの正規表現にする */
export function globToRegExp(glob: string): RegExp;
/** interhub.json の中身を確かめて対応表を作る */
export function compileConfig(raw: unknown): SyncConfig;
/** 対応表を読む */
export function loadConfig(file?: string): SyncConfig;
/** git に渡す対象のパス */
export function watchPathspecs(cfg: SyncConfig): string[];
/** InterHub のパス（リポジトリの根から）の行き先を引く */
export function mapPath(cfg: SyncConfig, repoPath: string): MapResult;
/** 相対の import を行き先どうしの相対パスに書き直す。引けなければ null */
export function resolveImport(
  cfg: SyncConfig,
  fromPath: string,
  fromDest: string,
  spec: string,
  warnings?: string[],
): string | null;
/** 差分の 1 行の中身に変換をかける */
export function transformLine(
  cfg: SyncConfig,
  side: { from: string; dest: string },
  body: string,
  warnings?: string[],
): string;
/** InterHub の git diff を XSheet 向けに書き直す */
export function rewritePatch(cfg: SyncConfig, diff: Buffer | string): RewriteResult;
export function parseNameStatusZ(
  buf: Buffer,
): { status: string; path: string; oldPath?: string; score?: number }[];
export function parseNumstatZ(
  buf: Buffer,
): Map<string, { added: number; deleted: number; binary: boolean }>;
/** 取り込み済みのコミットより後のコミットと触れたファイルの一覧を作る */
export function collectStatus(
  git: GitRunner,
  cfg: SyncConfig,
  options?: { range?: string; exists?: (destPath: string) => boolean },
): SyncStatus;
export function formatStatus(st: SyncStatus, options?: { dir?: string }): string;
/** パッチを作って outDir に書く（当てない） */
export function runPatch(
  git: GitRunner,
  cfg: SyncConfig,
  options: { range?: string; outDir?: string; interhubDir?: string },
): PatchResult;
export function formatPatch(r: PatchResult, cfg: SyncConfig): string;
/** CLI の本体。終了コードを返す */
export function main(
  argv: string[],
  options?: { env?: Record<string, string | undefined>; log?: (text: string) => void },
): number;
