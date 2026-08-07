/**
 * Minimal ZIP reader and writer.
 *
 * Reading is split in two — `listZipEntries` walks the central directory and
 * `readZipMember` decompresses one member — so a caller can inflate members one at
 * a time and yield to the event loop between them. Inflating a whole 38 MB archive
 * in one call would block for seconds with no way to cancel.
 *
 * Method 0 (stored) and method 8 (deflate) are supported in both directions,
 * through ./inflate.ts and ./deflate.ts.
 */
import { NotAZipError } from "../../errors.js";
import { crc32 } from "./crc32.js";
import { deflateRaw } from "./deflate.js";
import { inflateRaw } from "./inflate.js";

export interface ZipEntry {
  name: string;
  data: Uint8Array;
}

/** A member located in the archive but not yet decompressed. */
export interface ZipMember {
  name: string;
  /** ZIP compression method: 0 stored, 8 deflate. */
  method: number;
  /** The member's bytes as stored, still compressed when `method` is 8. */
  compressed: Uint8Array;
}

const SIG_LOCAL = 0x04034b50;
const SIG_CENTRAL = 0x02014b50;
const SIG_EOCD = 0x06054b50;

function readU16(b: Uint8Array, o: number): number {
  return (b[o] as number) | ((b[o + 1] as number) << 8);
}

function readU32(b: Uint8Array, o: number): number {
  return (
    ((b[o] as number) |
      ((b[o + 1] as number) << 8) |
      ((b[o + 2] as number) << 16) |
      ((b[o + 3] as number) << 24)) >>>
    0
  );
}

/** Size of the end-of-central-directory record, absent an archive comment. */
const EOCD_SIZE = 22;

/**
 * Locates every member of a ZIP archive without decompressing any of them.
 *
 * Walks the central directory rather than scanning local headers, so entries with
 * data descriptors (streamed ZIPs) still report correct sizes. The returned
 * `compressed` buffers are views into `bytes`, not copies — listing is cheap even
 * for a large archive, and the expensive part is deferred to `readZipMember`.
 */
export function listZipEntries(bytes: Uint8Array): ZipMember[] {
  let eocd = -1;
  for (let i = bytes.length - EOCD_SIZE; i >= 0; i--) {
    if (readU32(bytes, i) === SIG_EOCD) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) throw new NotAZipError();

  const cdOffset = readU32(bytes, eocd + 16);
  const cdEntries = readU16(bytes, eocd + 10);
  const members: ZipMember[] = [];
  const decoder = new TextDecoder();

  let p = cdOffset;
  for (let i = 0; i < cdEntries; i++) {
    if (readU32(bytes, p) !== SIG_CENTRAL) break;
    const method = readU16(bytes, p + 10);
    const compSize = readU32(bytes, p + 20);
    const nameLen = readU16(bytes, p + 28);
    const extraLen = readU16(bytes, p + 30);
    const commentLen = readU16(bytes, p + 32);
    const localOffset = readU32(bytes, p + 42);
    const name = decoder.decode(bytes.subarray(p + 46, p + 46 + nameLen));

    // The local header's extra field can differ in length from the central one's,
    // so the data offset must be computed from the local header.
    const lfNameLen = readU16(bytes, localOffset + 26);
    const lfExtraLen = readU16(bytes, localOffset + 28);
    const dataStart = localOffset + 30 + lfNameLen + lfExtraLen;

    members.push({
      name,
      method,
      compressed: bytes.subarray(dataStart, dataStart + compSize),
    });
    p += 46 + nameLen + extraLen + commentLen;
  }

  return members;
}

const METHOD_STORE = 0;

/** ZIP compression method 8 — deflate. Anything else is treated as stored. */
const METHOD_DEFLATE = 8;

/** Decompresses one member. The expensive half of reading an archive. */
export function readZipMember(member: ZipMember): Uint8Array {
  return member.method === METHOD_DEFLATE
    ? inflateRaw(member.compressed)
    : member.compressed;
}

/**
 * Below this many bytes, compressing costs more than it saves.
 *
 * A ZIP member carries about 80 bytes of header either way, and DEFLATE's own
 * framing plus a fixed-Huffman literal run can exceed the input on very short
 * parts — the `.rels` files are a few hundred bytes and several of them are
 * incompressible at this size.
 */
const MIN_COMPRESSED_SIZE = 256;

/**
 * Writes entries into a ZIP archive, compressing where it helps.
 *
 * Per member, whichever is smaller: STORE (method 0) or DEFLATE (method 8). The
 * comparison is done rather than assumed, because a stored member that happens
 * to be incompressible — an already-compressed PNG, which is most of the media
 * in a workbook with images — would otherwise grow.
 */
export function makeZip(fileList: ZipEntry[]): Uint8Array {
  const encoder = new TextEncoder();
  const localParts: Uint8Array[] = [];
  const centralParts: Uint8Array[] = [];
  let offset = 0;

  for (const f of fileList) {
    const nameBytes = encoder.encode(f.name);
    const size = f.data.length;
    const crc = crc32(f.data);

    const deflated = size >= MIN_COMPRESSED_SIZE ? deflateRaw(f.data) : undefined;
    const compress = deflated !== undefined && deflated.length < size;
    const data = compress ? deflated : f.data;
    const method = compress ? METHOD_DEFLATE : METHOD_STORE;

    const local = new Uint8Array(30 + nameBytes.length);
    const dv = new DataView(local.buffer);
    dv.setUint32(0, SIG_LOCAL, true);
    dv.setUint16(4, 20, true); // version needed
    dv.setUint16(8, method, true);
    dv.setUint32(14, crc, true);
    dv.setUint32(18, data.length, true);
    dv.setUint32(22, size, true);
    dv.setUint16(26, nameBytes.length, true);
    local.set(nameBytes, 30);
    localParts.push(local, data);

    const central = new Uint8Array(46 + nameBytes.length);
    const cdv = new DataView(central.buffer);
    cdv.setUint32(0, SIG_CENTRAL, true);
    cdv.setUint16(4, 20, true); // version made by
    cdv.setUint16(6, 20, true); // version needed
    cdv.setUint16(10, method, true);
    cdv.setUint32(16, crc, true);
    cdv.setUint32(20, data.length, true);
    cdv.setUint32(24, size, true);
    cdv.setUint16(28, nameBytes.length, true);
    cdv.setUint32(42, offset, true);
    central.set(nameBytes, 46);
    centralParts.push(central);

    offset += local.length + data.length;
  }

  const centralStart = offset;
  const centralSize = centralParts.reduce((n, c) => n + c.length, 0);

  const eocd = new Uint8Array(22);
  const edv = new DataView(eocd.buffer);
  edv.setUint32(0, SIG_EOCD, true);
  edv.setUint16(8, fileList.length, true);
  edv.setUint16(10, fileList.length, true);
  edv.setUint32(12, centralSize, true);
  edv.setUint32(16, centralStart, true);

  const all = [...localParts, ...centralParts, eocd];
  const total = all.reduce((n, a) => n + a.length, 0);
  const out = new Uint8Array(total);
  let pos = 0;
  for (const a of all) {
    out.set(a, pos);
    pos += a.length;
  }
  return out;
}

/** True when the buffer starts with the ZIP local-file magic "PK\x03\x04". */
export function isZip(bytes: Uint8Array): boolean {
  return (
    bytes.length >= 4 &&
    bytes[0] === 0x50 &&
    bytes[1] === 0x4b &&
    bytes[2] === 0x03 &&
    bytes[3] === 0x04
  );
}
