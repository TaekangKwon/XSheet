// InterHub の変更を XSheet へ取り込む道具（M6-1・ADR-0003 §4-2・`npm run sync:status`・`npm run sync:patch`）。
// - InterHub は読むだけ。git は読むだけのコマンド（rev-parse・log・diff・show・ls-tree・cat-file）しか呼ばない（`createGitRunner`）。
//   比べるのはコミットどうしだけで、作業ツリー・未コミットの変更は読まない。InterHub の下にファイルを書かない。通信しない。
// - status：取り込み済みのコミット（interhub.json の baseCommit）より後で、対応表のパスに触れたコミットと、
//   ファイルごとの追加・変更・削除・改名・行数・XSheet での行き先を出す。
// - patch：`git diff` のパスを行き先へ置き換え、変換の規則をかけたパッチを一時フォルダ（os.tmpdir() の下の xsheet-sync-*）に書く。
//   当てるのは人（`git apply --3way`）。行き先の無いファイル（手作業・対応なし）の差分は、元のパスのまま別のファイルに書く。
// - 対応表・変換の規則は tools/sync/interhub.json（正）。ここは表を引くだけで、ファイルの対応を持たない。
// - 規則の引き方・変換の中身は純粋な関数にして export する（tests/sync.test.ts。型は sync.d.mts）。CLI は末尾だけ。
import { Buffer } from 'node:buffer';
import { execFileSync } from 'node:child_process';
import { existsSync, mkdirSync, mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import process from 'node:process';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { TextDecoder } from 'node:util';

const HERE = path.dirname(fileURLToPath(import.meta.url));

/** XSheet のリポジトリの根 */
export const REPO = path.resolve(HERE, '../..');

/** 対応表の場所 */
export const CONFIG_PATH = path.join(HERE, 'interhub.json');

// ---------------------------------------------------------------------------
// git（読むだけ）

/** 呼んでよい git のサブコマンド。どれも InterHub のリポジトリにも作業ツリーにも書き込まない */
export const READ_ONLY_GIT = Object.freeze([
  'rev-parse',
  'log',
  'diff',
  'show',
  'ls-tree',
  'cat-file',
]);

// ファイルへ書き出す（--output）・外のプログラムを動かす（外部の diff・textconv・filters）オプションは、
// 読むだけのサブコマンドでも使わない
const FORBIDDEN_OPTION = /^(?:--output(?:=|$)|--ext-diff$|--textconv$|--filters$)/;

/** 読むだけの git の引数か確かめる。違えば例外 */
export function assertReadOnlyGitArgs(args) {
  if (!Array.isArray(args) || args.length === 0 || args.some((a) => typeof a !== 'string')) {
    throw new Error('git の引数は文字列の配列で渡す');
  }
  if (!READ_ONLY_GIT.includes(args[0])) {
    throw new Error(`git ${args[0]} は使わない（読むだけの ${READ_ONLY_GIT.join('・')} に限る）`);
  }
  const bad = args.find((a) => FORBIDDEN_OPTION.test(a));
  if (bad) throw new Error(`git のオプション ${bad} は使わない（書き出し・外のプログラムの実行）`);
}

/**
 * InterHub（`dir`）の git を呼ぶ関数を作る。`args[0]` は {@link READ_ONLY_GIT} のどれか。結果は Buffer。
 * - 引数は配列のまま渡す（シェルを通さない）。
 * - GIT_OPTIONAL_LOCKS=0：読むだけのコマンドでも index の更新などの任意の書き込みをさせない。
 * - 外部の diff（GIT_EXTERNAL_DIFF）は環境から外す。
 */
export function createGitRunner(dir, { exec = execFileSync } = {}) {
  const env = { ...process.env, GIT_OPTIONAL_LOCKS: '0', GIT_TERMINAL_PROMPT: '0', LC_ALL: 'C' };
  delete env.GIT_EXTERNAL_DIFF;
  return (args) => {
    assertReadOnlyGitArgs(args);
    return exec('git', ['-C', dir, '-c', 'core.quotePath=false', ...args], {
      env,
      stdio: ['ignore', 'pipe', 'pipe'],
      maxBuffer: 1024 * 1024 * 1024,
    });
  };
}

// コミットの指定（`943e4f0`・`HEAD~3`・`main` など）。`-` で始まるもの（オプションに化ける）と空白は通さない
const REV = /^[A-Za-z0-9][\w./@{}^~-]*$/;

/** `<a>..<b>`・`<a>..`・`<a>` を [a, b] にする（b を省くと HEAD） */
export function parseRange(range) {
  const parts = range.split('..');
  if (parts.length > 2 || range.includes('...')) {
    throw new Error(`範囲は <コミット>..<コミット> で書く：${range}`);
  }
  const [a, b = 'HEAD'] = parts;
  const head = b === '' ? 'HEAD' : b;
  for (const rev of [a, head]) {
    if (!REV.test(rev)) throw new Error(`コミットの指定が読めない：${rev}`);
  }
  return [a, head];
}

function revParse(git, rev) {
  if (!REV.test(rev)) throw new Error(`コミットの指定が読めない：${rev}`);
  return git(['rev-parse', '--verify', '--end-of-options', `${rev}^{commit}`])
    .toString('utf8')
    .trim();
}

// ---------------------------------------------------------------------------
// 対応表

/** glob（`*` は / を含まない 1 区切り、`**` は / を含む任意の並び）を、拾う組つきの正規表現にする */
export function globToRegExp(glob) {
  let re = '';
  for (let i = 0; i < glob.length; i++) {
    if (glob[i] === '*') {
      if (glob[i + 1] === '*') {
        re += '(.*)';
        i++;
      } else re += '([^/]*)';
    } else re += glob[i].replace(/[.+?^${}()|[\]\\]/g, '\\$&');
  }
  return new RegExp(`^${re}$`);
}

// to の * と ** へ、from で拾った文字を順に入れる
function fillStars(to, captures) {
  let k = 0;
  return to.replace(/\*\*|\*/g, () => captures[k++] ?? '');
}

function regex(find, flags, where) {
  try {
    return new RegExp(find, flags);
  } catch {
    throw new Error(`interhub.json：${where} の正規表現が読めない`);
  }
}

/** interhub.json の中身を確かめ、glob と正規表現を作っておく */
export function compileConfig(raw) {
  const need = (cond, msg) => {
    if (!cond) throw new Error(`interhub.json：${msg}`);
  };
  const isStr = (v) => typeof v === 'string' && v.length > 0;
  const strList = (v) => Array.isArray(v) && v.every(isStr);
  need(raw && typeof raw === 'object', '中身がオブジェクトでない');
  const ih = raw.interhub;
  need(ih && typeof ih === 'object', 'interhub が無い');
  need(
    typeof ih.baseCommit === 'string' && /^[0-9a-f]{7,40}$/.test(ih.baseCommit),
    'interhub.baseCommit はコミットの番号（16 進 7〜40 桁）',
  );
  need(isStr(ih.defaultDir), 'interhub.defaultDir が無い');
  need(isStr(ih.envVar), 'interhub.envVar が無い');
  need(
    typeof ih.root === 'string' && (ih.root === '' || ih.root.endsWith('/')),
    'interhub.root は / で終える',
  );
  need(strList(raw.watch) && raw.watch.length > 0, 'watch はパスの配列');
  need(raw.files && typeof raw.files === 'object', 'files が無い');
  const files = new Map();
  for (const [from, to] of Object.entries(raw.files)) {
    need(isStr(to), `files の ${from} の行き先が無い`);
    files.set(from, to);
  }
  const list = (key) => {
    const v = raw[key] ?? [];
    need(Array.isArray(v), `${key} は配列`);
    return v;
  };
  const manual = list('manual').map((m, i) => {
    need(isStr(m.path) && isStr(m.note), `manual[${i}] は path と note を持つ`);
    need(m.dest === undefined || strList(m.dest), `manual[${i}].dest はパスの配列`);
    return { re: globToRegExp(m.path), path: m.path, dest: m.dest ?? [], note: m.note };
  });
  const exclude = list('exclude').map((x, i) => {
    need(isStr(x.path) && isStr(x.reason), `exclude[${i}] は path と reason を持つ`);
    return { re: globToRegExp(x.path), path: x.path, reason: x.reason };
  });
  const patterns = list('patterns').map((p, i) => {
    need(isStr(p.from) && isStr(p.to), `patterns[${i}] は from と to を持つ`);
    const stars = (s) => (s.match(/\*\*|\*/g) ?? []).length;
    need(stars(p.from) === stars(p.to), `patterns[${i}] の from と to の * の数が違う`);
    return { re: globToRegExp(p.from), from: p.from, to: p.to, note: p.note };
  });
  const scoped = (key, item, i) => {
    need(strList(item.scope), `${key}[${i}].scope はパスの glob の配列`);
    need(isStr(item.find), `${key}[${i}].find が無い`);
    return item.scope.map(globToRegExp);
  };
  const transforms = list('transforms').map((t, i) => {
    need(isStr(t.id) && typeof t.replace === 'string', `transforms[${i}] は id と replace を持つ`);
    return {
      id: t.id,
      scope: scoped('transforms', t, i),
      re: regex(t.find, 'g', `transforms[${i}]`),
      replace: t.replace,
    };
  });
  const overrides = (raw.imports?.overrides ?? []).map((o, i) => {
    need(typeof o.replace === 'string', `imports.overrides[${i}] は replace を持つ`);
    return {
      scope: scoped('imports.overrides', o, i),
      re: regex(o.find, '', `imports.overrides[${i}]`),
      replace: o.replace,
      warn: o.warn,
    };
  });
  const leftovers = list('leftovers').map((l, i) => ({
    scope: scoped('leftovers', l, i),
    re: regex(l.find, '', `leftovers[${i}]`),
    note: l.note ?? l.find,
  }));
  const manualTransforms = list('manualTransforms').map((m, i) => {
    need(isStr(m.id) && isStr(m.note), `manualTransforms[${i}] は id と note を持つ`);
    return { id: m.id, note: m.note };
  });
  return {
    baseCommit: ih.baseCommit,
    defaultDir: ih.defaultDir,
    envVar: ih.envVar,
    root: ih.root,
    watch: [...raw.watch],
    files,
    manual,
    exclude,
    patterns,
    transforms,
    overrides,
    leftovers,
    manualTransforms,
  };
}

/** 対応表を読む */
export function loadConfig(file = CONFIG_PATH) {
  return compileConfig(JSON.parse(readFileSync(file, 'utf8')));
}

/** git に渡す対象のパス（InterHub のリポジトリの根から） */
export function watchPathspecs(cfg) {
  return cfg.watch.map((w) => cfg.root + w);
}

/**
 * InterHub のパス（リポジトリの根から）の行き先を引く。
 * - `file`：写した実績の組（files）。`rule`：まだ写していないファイルの提案（patterns）。どちらも `dest` を持つ。
 * - `manual`：手で直すもの（`hint` は写した先の候補）。`excluded`：写さないもの。`unmapped`：どの規則にも合わない。
 */
export function mapPath(cfg, repoPath) {
  if (!repoPath.startsWith(cfg.root)) return { kind: 'unmapped' };
  const rel = repoPath.slice(cfg.root.length);
  const exact = cfg.files.get(rel);
  if (exact) return { kind: 'file', dest: exact };
  for (const m of cfg.manual) {
    if (m.re.test(rel)) return { kind: 'manual', hint: m.dest, note: m.note };
  }
  for (const x of cfg.exclude) {
    if (x.re.test(rel)) return { kind: 'excluded', note: x.reason };
  }
  for (const p of cfg.patterns) {
    const m = p.re.exec(rel);
    if (m) return { kind: 'rule', dest: fillStars(p.to, m.slice(1)), note: p.note };
  }
  return { kind: 'unmapped' };
}

const inScope = (scope, dest) => scope.some((re) => re.test(dest));

// ---------------------------------------------------------------------------
// 変換（差分の 1 行ずつ）

// 相対の import の指定：from '…'・import '…'・import('…')・vi.mock('…')・vi.importActual('…')・new URL('…'
const IMPORT_SPEC =
  /(\bfrom\s*|\bimport\s*\(\s*|\bimport\s+|\bvi\.mock\(\s*|\bimportActual(?:<[^>]*>)?\(\s*|\bnew URL\(\s*)(['"`])(\.{1,2}\/[^'"`\s]*)/g;
const CODE_FILE = /\.(?:[cm]?[jt]s|tsx|jsx)$/;
const KNOWN_EXT = /\.(?:[cm]?js|json|mts|cts|ts|css|xlsx|png)$/;

/**
 * InterHub のファイル `fromPath`（行き先 `fromDest`）の中の相対の import `spec` を、行き先どうしの相対パスに書き直す。
 * 指す先を対応表で引けなければ null。`${…}` を含む・`/` で終わる指定は、ディレクトリの部分だけを引く（fixtures の new URL など）。
 * 書き直した後の指定に imports.overrides をかける（`warn` を持つものに当たれば warnings に足す）。
 */
export function resolveImport(cfg, fromPath, fromDest, spec, warnings = []) {
  const dir = path.posix.dirname(fromPath);
  const destDir = path.posix.dirname(fromDest);
  const rel = (target) => {
    const r = path.posix.relative(destDir, target);
    return r.startsWith('.') ? r : `./${r}`;
  };
  let out = null;
  const dyn = spec.indexOf('${');
  if (dyn >= 0 || spec.endsWith('/')) {
    const cut = spec.lastIndexOf('/', dyn >= 0 ? dyn : spec.length) + 1;
    const m = mapPath(cfg, path.posix.join(dir, spec.slice(0, cut), '__probe__'));
    if (m.dest) out = `${rel(path.posix.dirname(m.dest))}/${spec.slice(cut)}`;
  } else {
    const base = path.posix.join(dir, spec);
    const candidates = KNOWN_EXT.test(spec)
      ? [base, `${base}.ts`, `${base}/index.ts`]
      : [`${base}.ts`, `${base}/index.ts`, base];
    for (const cand of candidates) {
      const m = mapPath(cfg, cand);
      if (!m.dest) continue;
      out = rel(m.dest);
      // 拡張子を足して見つけたものは、元と同じく拡張子なしで書く
      if (cand !== base) out = out.replace(/\.ts$/, '');
      break;
    }
  }
  if (out === null) return null;
  for (const o of cfg.overrides) {
    if (inScope(o.scope, fromDest) && o.re.test(out)) {
      const next = out.replace(o.re, o.replace);
      if (o.warn) warnings.push(`${fromDest}：import '${spec}' → '${next}'：${o.warn}`);
      out = next;
    }
  }
  return out;
}

/**
 * 差分の 1 行の中身（先頭の `+`・`-`・空白を除いたもの）に変換をかける。
 * 1) コードのファイルなら相対の import を行き先に合わせて書き直す 2) 行き先が scope に合う transforms を順にかける。
 * 行の数は変えない（hunk の行数がずれないように、改行を生む置き換えは例外にする）。
 */
export function transformLine(cfg, { from, dest }, body, warnings = []) {
  let text = body;
  if (CODE_FILE.test(dest)) {
    text = text.replace(IMPORT_SPEC, (all, lead, quote, spec) => {
      const next = resolveImport(cfg, from, dest, spec, warnings);
      if (next === null) {
        warnings.push(`${dest}：import '${spec}' の行き先が対応表に無い（手で直す）`);
        return all;
      }
      return `${lead}${quote}${next}`;
    });
  }
  for (const t of cfg.transforms) {
    if (inScope(t.scope, dest)) text = text.replace(t.re, t.replace);
  }
  if (/[\r\n]/.test(text.replace(/\r$/, '')) && !/[\r\n]/.test(body.replace(/\r$/, ''))) {
    throw new Error(`変換で改行ができた（${dest}）`);
  }
  return text;
}

// ---------------------------------------------------------------------------
// パッチ

const UTF8 = new TextDecoder('utf-8', { fatal: true });

// 区切り（diff --git …）の見出しから、元と先のパス・種類を読む
function parseSectionHeader(lines) {
  const h = { header: lines[0], renameFrom: null, renameTo: null, minus: null, plus: null };
  let added = false;
  let deleted = false;
  for (const line of lines.slice(1)) {
    if (line.startsWith('@@') || line === 'GIT binary patch') break;
    if (line.startsWith('new file mode')) added = true;
    else if (line.startsWith('deleted file mode')) deleted = true;
    else if (line.startsWith('rename from ') || line.startsWith('copy from ')) {
      h.renameFrom = line.slice(line.indexOf(' from ') + 6);
    } else if (line.startsWith('rename to ') || line.startsWith('copy to ')) {
      h.renameTo = line.slice(line.indexOf(' to ') + 4);
    } else if (line.startsWith('--- a/')) h.minus = line.slice(6);
    else if (line.startsWith('+++ b/')) h.plus = line.slice(6);
  }
  // 改名でも ---・+++ も無い（モードだけ・バイナリ）ときは、`diff --git a/P b/P` の左右が同じ長さであることから読む
  let same = null;
  const body = h.header.slice('diff --git '.length);
  if ((body.length - 1) % 2 === 0) {
    const half = (body.length - 1) / 2;
    const a = body.slice(0, half);
    const b = body.slice(half + 1);
    if (a.startsWith('a/') && b.startsWith('b/') && a.slice(2) === b.slice(2)) same = a.slice(2);
  }
  const oldPath = added ? null : (h.renameFrom ?? h.minus ?? same);
  const newPath = deleted ? null : (h.renameTo ?? h.plus ?? same);
  const status = added ? 'A' : deleted ? 'D' : h.renameFrom && h.renameTo ? 'R' : 'M';
  return { oldPath, newPath, status, quoted: body.startsWith('"') };
}

/**
 * InterHub の `git diff`（`--binary`・`a/`・`b/` の前置き）を XSheet 向けに書き直す。
 * - 区切り（ファイル）ごとに、元と先のパスを対応表で引き、見出しのパスを行き先に置き換え、hunk の行に {@link transformLine} をかける。
 * - バイナリの区切り・UTF-8 でない区切りの中身は変えない。
 * - 行き先の無いファイル（手作業・除外・対応なし）は `patch` に入れず、元のまま `manual` に入れる。
 * @returns {{ patch: Buffer, manual: Buffer, files: object[], skipped: object[], warnings: string[] }}
 */
export function rewritePatch(cfg, diff) {
  const all = (Buffer.isBuffer(diff) ? diff : Buffer.from(diff, 'utf8')).toString('latin1');
  const starts = [...all.matchAll(/^diff --git /gm)].map((m) => m.index);
  const out = [];
  const manual = [];
  const files = [];
  const skipped = [];
  const warnings = [];
  starts.forEach((start, i) => {
    const latin = all.slice(start, starts[i + 1] ?? all.length);
    const raw = Buffer.from(latin, 'latin1');
    let text;
    let utf8 = true;
    try {
      text = UTF8.decode(raw);
    } catch {
      text = latin;
      utf8 = false;
    }
    const lines = text.split('\n');
    const head = parseSectionHeader(lines);
    const oldMap = head.oldPath ? mapPath(cfg, head.oldPath) : null;
    const newMap = head.newPath ? mapPath(cfg, head.newPath) : null;
    const which = newMap ?? oldMap;
    const bad = [newMap, oldMap].find((m) => m && !m.dest);
    if (head.quoted || !which || bad) {
      skipped.push({
        status: head.status,
        path: head.newPath ?? head.oldPath ?? lines[0],
        kind: bad?.kind ?? 'unmapped',
        hint: bad?.hint,
        note: bad ? bad.note : 'パスが読めない（特殊な文字など）',
      });
      manual.push(raw);
      return;
    }
    const oldDest = oldMap?.dest ?? null;
    const newDest = newMap?.dest ?? null;
    const sideOf = (c) =>
      c === '-'
        ? { from: head.oldPath ?? head.newPath, dest: oldDest ?? newDest }
        : { from: head.newPath ?? head.oldPath, dest: newDest ?? oldDest };
    const dest = newDest ?? oldDest;
    files.push({
      status: head.status,
      path: head.newPath ?? head.oldPath,
      oldPath: head.oldPath,
      dest,
      kind: which.kind,
    });
    if (!utf8) warnings.push(`${dest}：UTF-8 でないので中身に変換をかけなかった`);
    const res = [];
    let inHunk = false;
    let binary = false;
    const hits = new Map();
    for (const line of lines) {
      if (!inHunk && !binary) {
        if (line.startsWith('diff --git '))
          res.push(`diff --git a/${oldDest ?? newDest} b/${newDest ?? oldDest}`);
        else if (line.startsWith('--- a/')) res.push(`--- a/${oldDest}`);
        else if (line.startsWith('+++ b/')) res.push(`+++ b/${newDest}`);
        else if (line.startsWith('rename from ')) res.push(`rename from ${oldDest}`);
        else if (line.startsWith('rename to ')) res.push(`rename to ${newDest}`);
        else if (line.startsWith('copy from ')) res.push(`copy from ${oldDest}`);
        else if (line.startsWith('copy to ')) res.push(`copy to ${newDest}`);
        else {
          if (line === 'GIT binary patch') binary = true;
          else if (line.startsWith('@@')) inHunk = true;
          res.push(line);
        }
        continue;
      }
      const c = line[0];
      if (binary || !utf8 || (c !== '+' && c !== '-' && c !== ' ')) {
        res.push(line);
        continue;
      }
      const side = sideOf(c);
      const body = transformLine(cfg, side, line.slice(1), c === '+' ? warnings : []);
      if (c === '+') {
        for (const l of cfg.leftovers) {
          if (inScope(l.scope, side.dest) && l.re.test(body))
            hits.set(l.note, (hits.get(l.note) ?? 0) + 1);
        }
      }
      res.push(c + body);
    }
    for (const [note, n] of hits) warnings.push(`${dest}：${note}（${n} 行）`);
    out.push(Buffer.from(res.join('\n'), utf8 ? 'utf8' : 'latin1'));
  });
  return {
    patch: Buffer.concat(out),
    manual: Buffer.concat(manual),
    files,
    skipped,
    warnings: [...new Set(warnings)],
  };
}

// ---------------------------------------------------------------------------
// status

/** `git diff --name-status -z` を読む */
export function parseNameStatusZ(buf) {
  const t = buf.toString('utf8').split('\0');
  const out = [];
  for (let i = 0; i < t.length;) {
    const s = t[i++];
    if (!s) continue;
    const status = s[0];
    if (status === 'R' || status === 'C') {
      out.push({ status, score: Number(s.slice(1)), oldPath: t[i++], path: t[i++] });
    } else out.push({ status, path: t[i++] });
  }
  return out;
}

/** `git diff --numstat -z` を読む（先のパス → 行数。バイナリは binary: true） */
export function parseNumstatZ(buf) {
  const t = buf.toString('utf8').split('\0');
  const out = new Map();
  for (let i = 0; i < t.length;) {
    const s = t[i++];
    if (!s) continue;
    const [a, d, p] = s.split('\t');
    let key = p;
    if (p === '') {
      i++; // 改名の元
      key = t[i++];
    }
    const binary = a === '-';
    out.set(key, { added: binary ? 0 : Number(a), deleted: binary ? 0 : Number(d), binary });
  }
  return out;
}

const DIFF_OPTS = ['--no-ext-diff', '--no-textconv', '--no-color', '--no-relative', '-M'];

function diffFiles(git, cfg, a, b, exists) {
  const spec = ['--', ...watchPathspecs(cfg)];
  const names = parseNameStatusZ(git(['diff', ...DIFF_OPTS, '--name-status', '-z', a, b, ...spec]));
  const nums = parseNumstatZ(git(['diff', ...DIFF_OPTS, '--numstat', '-z', a, b, ...spec]));
  return names.map((e) => {
    const m = mapPath(cfg, e.path);
    const n = nums.get(e.path) ?? { added: 0, deleted: 0, binary: false };
    return {
      ...e,
      ...n,
      kind: m.kind,
      dest: m.dest,
      hint: m.hint,
      note: m.note,
      destExists: m.dest ? exists(m.dest) : undefined,
    };
  });
}

/**
 * 取り込み済みのコミットより後の InterHub のコミットと、触れたファイルの一覧を作る（コミット済みのものだけ）。
 * - commits：対応表のパスに触れたマージ以外のコミット（古い順）と、それぞれの親との差分のファイル。
 * - files：範囲全体（base..head）の差分のファイル。
 * @param exists XSheet にそのパスがあるか（行き先が新しいかの表示に使う）
 */
export function collectStatus(git, cfg, { range, exists = () => false } = {}) {
  const [a, b] = parseRange(range ?? `${cfg.baseCommit}..HEAD`);
  const base = revParse(git, a);
  const head = revParse(git, b);
  const span = `${base}..${head}`;
  const spec = ['--', ...watchPathspecs(cfg)];
  const lines = (buf) => buf.toString('utf8').split('\n').filter(Boolean);
  const allCount = lines(git(['log', '--format=%H', span])).length;
  const mergeCount = lines(git(['log', '--merges', '--format=%H', span])).length;
  const commits = git([
    'log',
    '--no-merges',
    '--reverse',
    '--date=short',
    '--format=%H%x1f%ad%x1f%s%x1e',
    span,
    ...spec,
  ])
    .toString('utf8')
    .split('\x1e')
    .map((r) => r.trim())
    .filter(Boolean)
    .map((r) => {
      const [hash, date, subject] = r.split('\x1f');
      return {
        hash,
        short: hash.slice(0, 7),
        date,
        subject,
        files: diffFiles(git, cfg, `${hash}^`, hash, exists),
      };
    });
  return {
    root: cfg.root,
    base,
    head,
    allCount,
    mergeCount,
    commits,
    files: diffFiles(git, cfg, base, head, exists),
  };
}

const KIND_LABEL = {
  file: '対応表',
  rule: '規則',
  manual: '手作業',
  excluded: '除外',
  unmapped: '対応なし',
};
const STATUS_LABEL = { A: '追加', M: '変更', D: '削除', R: '改名', C: '複写', T: '型' };

// 表示は InterHub の root（src_V1/）からのパスにする
const fromRoot = (root, p) => (p.startsWith(root) ? p.slice(root.length) : p);

function fileLine(root, f) {
  const rel = (p) => fromRoot(root, p);
  const name = f.oldPath ? `${rel(f.oldPath)} → ${rel(f.path)}` : rel(f.path);
  const lines = f.binary ? 'バイナリ' : `+${f.added} -${f.deleted}`;
  let to;
  if (f.dest) {
    const fresh = f.status !== 'D' && !f.destExists ? '・新規' : '';
    to = `${f.dest}（${KIND_LABEL[f.kind]}${fresh}）`;
  } else if (f.kind === 'manual') to = `${f.hint?.length ? f.hint.join('・') : '—'}（手作業）`;
  else to = `—（${KIND_LABEL[f.kind]}）`;
  return `  ${f.status}  ${name}  ${lines}  → ${to}`;
}

/** status の結果を人が読む文にする */
export function formatStatus(st, { dir } = {}) {
  const short = (h) => h.slice(0, 7);
  const out = [];
  out.push(
    `InterHub${dir ? `（${dir}）` : ''}：${short(st.base)}..${short(st.head)}（パスは ${st.root || '根'} から）`,
  );
  out.push(
    `取り込み済みより後のコミット ${st.allCount} 件（マージ ${st.mergeCount} 件を含む）のうち、対応表のパスに触れたもの ${st.commits.length} 件（マージを除く）`,
  );
  for (const c of st.commits) {
    out.push('', `${c.short} ${c.date} ${c.subject}`);
    for (const f of c.files) out.push(fileLine(st.root, f));
  }
  const count = (key, labels) =>
    Object.entries(labels)
      .map(([k, label]) => [label, st.files.filter((f) => f[key] === k).length])
      .filter(([, n]) => n > 0)
      .map(([label, n]) => `${label} ${n}`)
      .join('・') || 'なし';
  out.push(
    '',
    `まとめ（${short(st.base)}..${short(st.head)} の差分・${st.files.length} ファイル）`,
  );
  out.push(`  ${count('status', STATUS_LABEL)}`);
  out.push(`  ${count('kind', KIND_LABEL)}`);
  for (const f of st.files) out.push(fileLine(st.root, f));
  const unmapped = st.files.filter((f) => f.kind === 'unmapped');
  if (unmapped.length) {
    out.push(
      '',
      '対応なし（interhub.json の files か patterns に行き先を足すか、exclude に足す）：',
    );
    for (const f of unmapped) out.push(`  ${fromRoot(st.root, f.path)}`);
  }
  const manual = st.files.filter((f) => f.kind === 'manual');
  if (manual.length) {
    out.push('', '手作業（patch に入らない。差分を読んで写した先を手で直す）：');
    for (const f of manual) out.push(`  ${fromRoot(st.root, f.path)}：${f.note}`);
  }
  return out.join('\n');
}

// ---------------------------------------------------------------------------
// patch

/**
 * InterHub の範囲の差分から XSheet 向けのパッチを作り、`outDir` に書く（当てない）。
 * `interhubDir` の下には書かない（書こうとしたら例外）。
 */
export function runPatch(git, cfg, { range, outDir, interhubDir }) {
  const [a, b] = parseRange(range ?? `${cfg.baseCommit}..HEAD`);
  const base = revParse(git, a);
  const head = revParse(git, b);
  const diff = git([
    'diff',
    ...DIFF_OPTS,
    '--binary',
    '--src-prefix=a/',
    '--dst-prefix=b/',
    base,
    head,
    '--',
    ...watchPathspecs(cfg),
  ]);
  const res = rewritePatch(cfg, diff);
  const dir = path.resolve(outDir ?? mkdtempSync(path.join(tmpdir(), 'xsheet-sync-')));
  if (interhubDir) {
    const r = path.relative(path.resolve(interhubDir), dir);
    if (r === '' || (!r.startsWith('..') && !path.isAbsolute(r))) {
      throw new Error('パッチを InterHub の下には書かない');
    }
  }
  mkdirSync(dir, { recursive: true });
  const name = `interhub-${base.slice(0, 7)}-${head.slice(0, 7)}`;
  const file = path.join(dir, `${name}.patch`);
  writeFileSync(file, res.patch);
  let manualFile = null;
  if (res.manual.length > 0) {
    manualFile = path.join(dir, `${name}.manual.patch`);
    writeFileSync(manualFile, res.manual);
  }
  return { ...res, base, head, file, manualFile };
}

/** patch の結果を人が読む文にする */
export function formatPatch(r, cfg) {
  const out = [];
  out.push(
    `InterHub ${r.base.slice(0, 7)}..${r.head.slice(0, 7)} のパッチを書いた（まだ当てていない）：`,
  );
  out.push(`  ${r.file}（${r.files.length} ファイル）`);
  if (r.manualFile) {
    out.push(
      `  ${r.manualFile}（行き先の無い ${r.skipped.length} ファイル。元のパスのまま。読んで手で写す）`,
    );
    for (const s of r.skipped)
      out.push(
        `    ${s.status}  ${s.path}（${KIND_LABEL[s.kind] ?? s.kind}）${s.note ? `：${s.note}` : ''}`,
      );
  }
  if (r.warnings.length) {
    out.push('', '確かめること：');
    for (const w of r.warnings) out.push(`  ${w}`);
  }
  out.push(
    '',
    '当てる（XSheet のリポジトリの根で。作業用のブランチで）：',
    `  git apply --3way --check "${r.file}"`,
    `  git apply --3way "${r.file}"`,
    '当たらない hunk があれば（Prettier で整えた後のファイルなど）：',
    `  git apply --reject "${r.file}"   （.rej を読んで手で直す）`,
    '',
    '当てた後：',
  );
  for (const m of cfg.manualTransforms) out.push(`  - ${m.note}`);
  out.push(
    '  - npm run typecheck・npm run lint・npm run test を通す',
    `  - tools/sync/interhub.json の interhub.baseCommit を ${r.head.slice(0, 7)} に進め、docs/.history/progress.md に記録する`,
  );
  return out.join('\n');
}

// ---------------------------------------------------------------------------
// CLI

const USAGE = `使い方：
  node tools/sync/sync.mjs status [<範囲>] [--json]
  node tools/sync/sync.mjs patch  [<範囲>] [--out-dir <フォルダ>] [--json]
範囲の既定は <interhub.json の baseCommit>..HEAD。InterHub の場所は環境変数 INTERHUB_DIR（無ければ interhub.json の defaultDir）。`;

function parseArgs(argv) {
  const out = { cmd: argv[0], range: undefined, json: false, outDir: undefined };
  for (let i = 1; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--json') out.json = true;
    else if (a === '--out-dir') out.outDir = argv[++i];
    else if (a.startsWith('-')) throw new Error(`知らないオプション：${a}`);
    else if (out.range === undefined) out.range = a;
    else throw new Error(`引数が多い：${a}`);
  }
  return out;
}

/** CLI の本体（テストからも呼ぶ）。終了コードを返す */
export function main(argv, { env = process.env, log = console.log } = {}) {
  const args = parseArgs(argv);
  if (args.cmd !== 'status' && args.cmd !== 'patch') {
    log(USAGE);
    return args.cmd === undefined || args.cmd === 'help' || args.cmd === '--help' ? 0 : 1;
  }
  const cfg = loadConfig();
  const dir = path.resolve(env[cfg.envVar] || cfg.defaultDir);
  if (!existsSync(dir))
    throw new Error(`InterHub が見つからない（${cfg.envVar} で場所を指す）：${dir}`);
  const git = createGitRunner(dir);
  if (args.cmd === 'status') {
    const st = collectStatus(git, cfg, {
      range: args.range,
      exists: (p) => existsSync(path.join(REPO, p)),
    });
    log(
      args.json ? JSON.stringify({ interhubDir: dir, ...st }, null, 2) : formatStatus(st, { dir }),
    );
    return 0;
  }
  const r = runPatch(git, cfg, { range: args.range, outDir: args.outDir, interhubDir: dir });
  if (args.json) {
    const { patch: _p, manual: _m, ...rest } = r;
    log(JSON.stringify(rest, null, 2));
  } else log(formatPatch(r, cfg));
  return 0;
}

if (process.argv[1] && pathToFileURL(path.resolve(process.argv[1])).href === import.meta.url) {
  try {
    process.exitCode = main(process.argv.slice(2));
  } catch (e) {
    console.error(e instanceof Error ? e.message : e);
    process.exitCode = 1;
  }
}
