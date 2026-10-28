// 正: InterHub docs/plan/adr/0025-xlsx-own-reader.md
/**
 * xlsx（zip）を開く。先頭の印の確認・ファイル数と展開後の大きさの上限・部品の文字での読み出し。
 * 暗号化されたブック（OLE 形式。先頭 D0 CF 11 E0）と、zip でないものは、zip として読む前に弾く。
 */
import JSZip from 'jszip';
import type { XlsxLimits } from './limits';
import { XlsxError } from './types';

/** 開いた zip（部品の読み出し口） */
export interface XlsxPackage {
  /** zip の中のファイルのパス（先頭に / を付けない） */
  has(path: string): boolean;
  /** 部品を文字（UTF-8）で読む。無ければ null。展開後の大きさが maxBytes を超えたら too-large */
  readText(path: string, maxBytes: number): Promise<string | null>;
  /** 部品をバイト列で読む（画像など）。無ければ null。展開しながら数え、maxBytes を超えたらそこで止めて too-large */
  readBytes(path: string, maxBytes: number): Promise<Uint8Array | null>;
  /** 読み込んだ zip を放す */
  dispose(): void;
}

/** Uint8Array へそろえる */
export function toBytes(data: ArrayBuffer | Uint8Array): Uint8Array {
  return data instanceof Uint8Array ? data : new Uint8Array(data);
}

/**
 * 先頭の印を確かめる。
 * - `PK\x03\x04` 以外は zip でない。
 * - OLE（`D0 CF 11 E0`）は、パスワード付きの xlsx（暗号化）か古い形式。どちらにしても読めないので encrypted。
 */
export function checkSignature(bytes: Uint8Array): void {
  if (
    bytes.length >= 4 &&
    bytes[0] === 0xd0 &&
    bytes[1] === 0xcf &&
    bytes[2] === 0x11 &&
    bytes[3] === 0xe0
  ) {
    throw new XlsxError('encrypted');
  }
  if (
    bytes.length < 4 ||
    bytes[0] !== 0x50 ||
    bytes[1] !== 0x4b ||
    bytes[2] !== 0x03 ||
    bytes[3] !== 0x04
  ) {
    throw new XlsxError('not-zip');
  }
}

/** jszip の内部の型（公開の型定義に無い部分） */
interface InternalFile {
  _data?: { uncompressedSize?: number };
  internalStream(type: 'uint8array'): {
    on(event: 'data', cb: (chunk: Uint8Array) => void): unknown;
    on(event: 'end', cb: () => void): unknown;
    on(event: 'error', cb: (e: unknown) => void): unknown;
    resume(): unknown;
    pause(): unknown;
  };
}

/**
 * zip の 1 ファイルを展開しながら、実際に出てきた大きさを数える。maxBytes を超えたらそこで止めて too-large。
 * 見出しの大きさ（偽りうる）が maxBytes を超えるものは、展開前に断る。流れてきた塊は onChunk へ渡し、全体の大きさを返す。
 */
function streamEntry(
  file: InternalFile,
  maxBytes: number,
  onChunk: (chunk: Uint8Array) => void,
): Promise<number> {
  const declared = file._data?.uncompressedSize;
  if (typeof declared === 'number' && declared > maxBytes)
    return Promise.reject(new XlsxError('too-large', 'part'));
  return new Promise<number>((resolve, reject) => {
    let total = 0;
    let done = false;
    const stream = file.internalStream('uint8array');
    stream.on('data', (chunk: Uint8Array) => {
      if (done) return;
      total += chunk.length;
      if (total > maxBytes) {
        done = true;
        stream.pause();
        reject(new XlsxError('too-large', 'part'));
        return;
      }
      onChunk(chunk);
    });
    stream.on('error', () => {
      if (done) return;
      done = true;
      reject(new XlsxError('broken', 'entry'));
    });
    stream.on('end', () => {
      if (done) return;
      done = true;
      resolve(total);
    });
    stream.resume();
  });
}

/** zip を開く。壊れていれば broken、ファイル数が多すぎれば too-large */
export async function openPackage(bytes: Uint8Array, limits: XlsxLimits): Promise<XlsxPackage> {
  let zip: JSZip | null;
  try {
    zip = await JSZip.loadAsync(bytes);
  } catch {
    throw new XlsxError('broken', 'zip');
  }
  if (Object.keys(zip.files).length > limits.maxEntries)
    throw new XlsxError('too-large', 'entries');

  return {
    has: (path) => zip !== null && zip.file(path) !== null,
    async readText(path, maxBytes) {
      if (zip === null) throw new XlsxError('broken', 'closed');
      const file = zip.file(path) as unknown as InternalFile | null;
      if (file === null) return null;
      const decoder = new TextDecoder('utf-8');
      const parts: string[] = [];
      await streamEntry(file, maxBytes, (chunk) =>
        parts.push(decoder.decode(chunk, { stream: true })),
      );
      parts.push(decoder.decode());
      return parts.join('');
    },
    async readBytes(path, maxBytes) {
      if (zip === null) throw new XlsxError('broken', 'closed');
      const file = zip.file(path) as unknown as InternalFile | null;
      if (file === null) return null;
      const chunks: Uint8Array[] = [];
      const total = await streamEntry(file, maxBytes, (chunk) => chunks.push(chunk));
      const out = new Uint8Array(total);
      let at = 0;
      for (const c of chunks) {
        out.set(c, at);
        at += c.length;
      }
      return out;
    },
    dispose() {
      zip = null;
    },
  };
}
