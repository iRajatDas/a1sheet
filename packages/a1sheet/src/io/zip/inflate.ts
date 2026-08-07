/**
 * RFC 1951 DEFLATE decoder, written by hand to keep the dependency count at zero.
 *
 * Handles all three block types: stored (0), fixed Huffman (1), dynamic
 * Huffman (2). There is NO encoder — see ./deflate.ts.
 */

// Length and distance code tables from RFC 1951 §3.2.5.
const LENBASE = [
  3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83,
  99, 115, 131, 163, 195, 227, 258,
];
const LENEXTRA = [
  0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5,
  5, 0,
];
const DISTBASE = [
  1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769,
  1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577,
];
const DISTEXTRA = [
  0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11,
  12, 12, 13, 13,
];
/** Code-length code order, RFC 1951 §3.2.7. */
const CLCIDX = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15];

const MAXBITS = 15;

class BitReader {
  bytes: Uint8Array;
  pos = 0;
  bitpos = 0;

  constructor(bytes: Uint8Array) {
    this.bytes = bytes;
  }

  getBit(): number {
    if (this.pos >= this.bytes.length) return 0;
    const bit = ((this.bytes[this.pos] as number) >> this.bitpos) & 1;
    this.bitpos++;
    if (this.bitpos === 8) {
      this.bitpos = 0;
      this.pos++;
    }
    return bit;
  }

  getBits(n: number): number {
    let v = 0;
    for (let i = 0; i < n; i++) v |= this.getBit() << i;
    return v;
  }

  /** Discards bits up to the next byte boundary, for stored blocks. */
  align(): void {
    if (this.bitpos !== 0) {
      this.bitpos = 0;
      this.pos++;
    }
  }
}

interface HuffmanTable {
  counts: number[];
  symbols: number[];
}

/** Canonical Huffman table from a list of code lengths. */
function buildHuffman(lengths: number[]): HuffmanTable {
  const counts = new Array(MAXBITS + 1).fill(0);
  for (const len of lengths) if (len) counts[len]++;

  const offs = new Array(MAXBITS + 2).fill(0);
  for (let i = 1; i <= MAXBITS; i++) offs[i + 1] = offs[i] + counts[i];

  const symbols = new Array(lengths.length).fill(0);
  const offsCopy = offs.slice();
  for (let i = 0; i < lengths.length; i++) {
    const len = lengths[i] as number;
    if (len) symbols[offsCopy[len]++] = i;
  }

  return { counts, symbols };
}

function decodeSymbol(br: BitReader, table: HuffmanTable): number {
  let code = 0;
  let first = 0;
  let index = 0;
  for (let len = 1; len <= MAXBITS; len++) {
    code |= br.getBit();
    const count = table.counts[len] ?? 0;
    if (code - first < count)
      return table.symbols[index + (code - first)] as number;
    index += count;
    first = (first + count) << 1;
    code <<= 1;
  }
  throw new Error("inflate: bad huffman code");
}

const FIXED_LIT_TABLE = (() => {
  const lens = new Array<number>(288);
  for (let i = 0; i < 144; i++) lens[i] = 8;
  for (let i = 144; i < 256; i++) lens[i] = 9;
  for (let i = 256; i < 280; i++) lens[i] = 7;
  for (let i = 280; i < 288; i++) lens[i] = 8;
  return buildHuffman(lens);
})();

const FIXED_DIST_TABLE = buildHuffman(new Array(30).fill(5));

export function inflateRaw(bytes: Uint8Array): Uint8Array {
  const br = new BitReader(bytes);
  const out: number[] = [];
  let final = 0;

  while (!final) {
    final = br.getBit();
    const type = br.getBits(2);

    if (type === 0) {
      // Stored: byte-aligned LEN, NLEN, then raw bytes.
      br.align();
      const len = (bytes[br.pos] as number) | ((bytes[br.pos + 1] as number) << 8);
      br.pos += 4;
      for (let i = 0; i < len; i++) out.push(bytes[br.pos + i] as number);
      br.pos += len;
      continue;
    }

    let litTable: HuffmanTable;
    let distTable: HuffmanTable;

    if (type === 1) {
      litTable = FIXED_LIT_TABLE;
      distTable = FIXED_DIST_TABLE;
    } else if (type === 2) {
      const hlit = br.getBits(5) + 257;
      const hdist = br.getBits(5) + 1;
      const hclen = br.getBits(4) + 4;

      const clLengths = new Array(19).fill(0);
      for (let i = 0; i < hclen; i++) {
        clLengths[CLCIDX[i] as number] = br.getBits(3);
      }
      const clTable = buildHuffman(clLengths);

      const lengths: number[] = [];
      while (lengths.length < hlit + hdist) {
        const sym = decodeSymbol(br, clTable);
        if (sym < 16) {
          lengths.push(sym);
        } else if (sym === 16) {
          const rep = br.getBits(2) + 3;
          const prev = lengths[lengths.length - 1] ?? 0;
          for (let i = 0; i < rep; i++) lengths.push(prev);
        } else if (sym === 17) {
          const rep = br.getBits(3) + 3;
          for (let i = 0; i < rep; i++) lengths.push(0);
        } else {
          const rep = br.getBits(7) + 11;
          for (let i = 0; i < rep; i++) lengths.push(0);
        }
      }

      litTable = buildHuffman(lengths.slice(0, hlit));
      distTable = buildHuffman(lengths.slice(hlit));
    } else {
      throw new Error("inflate: invalid deflate block type");
    }

    for (;;) {
      const sym = decodeSymbol(br, litTable);
      if (sym < 256) {
        out.push(sym);
      } else if (sym === 256) {
        break;
      } else {
        const li = sym - 257;
        const len = (LENBASE[li] as number) + br.getBits(LENEXTRA[li] as number);
        const dsym = decodeSymbol(br, distTable);
        const dist =
          (DISTBASE[dsym] as number) + br.getBits(DISTEXTRA[dsym] as number);
        const start = out.length - dist;
        // Overlapping copies are legal and intentional in LZ77.
        for (let i = 0; i < len; i++) out.push(out[start + i] as number);
      }
    }
  }

  return new Uint8Array(out);
}
