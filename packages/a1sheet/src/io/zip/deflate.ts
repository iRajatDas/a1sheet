/**
 * DEFLATE encoder — fixed-Huffman, with LZ77 matching.
 *
 * Enough to make an exported .xlsx a normal size. The parts are XML, which is
 * mostly repeated tag names and attribute patterns, so even a modest matcher
 * takes a 4 MB workbook to well under a megabyte.
 *
 * Fixed Huffman rather than dynamic: the fixed tables are defined by the format
 * and need no header, and for XML they cost only a few percent over optimal
 * tables while removing the whole frequency-counting and tree-building pass.
 * `deflateRaw` returns a raw DEFLATE stream — no zlib header, which is what a
 * ZIP member holds.
 *
 * The matcher is a hash chain over three-byte prefixes, the standard approach:
 * every position is filed under the hash of the three bytes starting there, and
 * a candidate match is found by walking the chain of earlier positions with the
 * same hash. `MAX_CHAIN` bounds that walk, which is the compression/speed dial.
 */

const WINDOW_SIZE = 32768;
const MIN_MATCH = 3;
const MAX_MATCH = 258;
/** How many candidates to consider per position. Higher is smaller and slower. */
const MAX_CHAIN = 32;
/** A match at least this long ends the search immediately. */
const GOOD_MATCH = 32;

const HASH_BITS = 15;
const HASH_SIZE = 1 << HASH_BITS;
const HASH_MASK = HASH_SIZE - 1;

/** Writes bits least-significant-first, which is DEFLATE's order. */
class BitWriter {
  private bytes: number[] = [];
  private bitBuffer = 0;
  private bitCount = 0;

  write(value: number, bits: number): void {
    this.bitBuffer |= value << this.bitCount;
    this.bitCount += bits;
    while (this.bitCount >= 8) {
      this.bytes.push(this.bitBuffer & 0xff);
      this.bitBuffer >>>= 8;
      this.bitCount -= 8;
    }
  }

  /** Huffman codes are stored most-significant-first and written reversed. */
  writeCode(code: number, bits: number): void {
    let reversed = 0;
    for (let i = 0; i < bits; i++) {
      reversed = (reversed << 1) | ((code >>> i) & 1);
    }
    this.write(reversed, bits);
  }

  finish(): Uint8Array {
    if (this.bitCount > 0) this.bytes.push(this.bitBuffer & 0xff);
    return new Uint8Array(this.bytes);
  }
}

/**
 * The fixed literal/length code, from RFC 1951 section 3.2.6.
 *
 * 0-143 are eight bits starting at 0b00110000, 144-255 nine bits starting at
 * 0b110010000, 256-279 seven bits from 0, and 280-287 eight bits from
 * 0b11000000. Reproduced as a function rather than a table because the ranges
 * are the specification and a table would hide them.
 */
function literalCode(symbol: number): { code: number; bits: number } {
  if (symbol < 144) return { code: 0b00110000 + symbol, bits: 8 };
  if (symbol < 256) return { code: 0b110010000 + (symbol - 144), bits: 9 };
  if (symbol < 280) return { code: symbol - 256, bits: 7 };
  return { code: 0b11000000 + (symbol - 280), bits: 8 };
}

/** Length symbol 257..285, plus the extra bits that refine it. */
const LENGTH_BASE = [
  3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83,
  99, 115, 131, 163, 195, 227, 258,
];
const LENGTH_EXTRA = [
  0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5,
  5, 0,
];

const DISTANCE_BASE = [
  1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769,
  1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577,
];
const DISTANCE_EXTRA = [
  0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11,
  12, 12, 13, 13,
];

function symbolFor(value: number, bases: readonly number[]): number {
  // Linear from the top: the tables are short and the high symbols are the
  // common ones for XML, where matches tend to be long.
  for (let i = bases.length - 1; i >= 0; i--) {
    if (value >= (bases[i] as number)) return i;
  }
  return 0;
}

const END_OF_BLOCK = 256;
const FIRST_LENGTH_SYMBOL = 257;
/** Fixed distance codes are five bits, straight through. */
const DISTANCE_CODE_BITS = 5;

/**
 * Compresses to a raw DEFLATE stream.
 *
 * One block, `BFINAL=1`, `BTYPE=01` (fixed Huffman). A single block is fine at
 * any size the format allows — block boundaries exist to let the tables change,
 * and the fixed tables never do.
 */
export function deflateRaw(input: Uint8Array): Uint8Array {
  const out = new BitWriter();
  out.write(1, 1); // BFINAL
  out.write(1, 2); // BTYPE = fixed Huffman

  // head[hash] is the most recent position with that hash; prev[pos] chains back.
  const head = new Int32Array(HASH_SIZE).fill(-1);
  const prev = new Int32Array(input.length).fill(-1);

  const hashAt = (pos: number) =>
    (((input[pos] as number) << 10) ^
      ((input[pos + 1] as number) << 5) ^
      (input[pos + 2] as number)) &
    HASH_MASK;

  const emitLiteral = (byte: number) => {
    const { code, bits } = literalCode(byte);
    out.writeCode(code, bits);
  };

  let pos = 0;
  while (pos < input.length) {
    let bestLength = 0;
    let bestDistance = 0;

    if (pos + MIN_MATCH <= input.length) {
      const hash = hashAt(pos);
      let candidate = head[hash] as number;
      let chain = MAX_CHAIN;

      while (candidate >= 0 && chain-- > 0) {
        const distance = pos - candidate;
        if (distance > WINDOW_SIZE) break;

        // Check the byte that would extend the current best first: if it does
        // not match, this candidate cannot beat what we have.
        if (input[candidate + bestLength] === input[pos + bestLength]) {
          let length = 0;
          while (
            length < MAX_MATCH &&
            pos + length < input.length &&
            input[candidate + length] === input[pos + length]
          ) {
            length++;
          }
          if (length > bestLength) {
            bestLength = length;
            bestDistance = distance;
            if (length >= GOOD_MATCH) break;
          }
        }
        candidate = prev[candidate] as number;
      }

      // File this position, and every position the match covers, so later
      // matches can start inside it.
      const covered = Math.max(1, bestLength >= MIN_MATCH ? bestLength : 1);
      for (let i = 0; i < covered && pos + i + MIN_MATCH <= input.length; i++) {
        const h = hashAt(pos + i);
        prev[pos + i] = head[h] as number;
        head[h] = pos + i;
      }
    }

    if (bestLength >= MIN_MATCH) {
      const lengthSymbol = symbolFor(bestLength, LENGTH_BASE);
      const { code, bits } = literalCode(FIRST_LENGTH_SYMBOL + lengthSymbol);
      out.writeCode(code, bits);
      out.write(
        bestLength - (LENGTH_BASE[lengthSymbol] as number),
        LENGTH_EXTRA[lengthSymbol] as number,
      );

      const distanceSymbol = symbolFor(bestDistance, DISTANCE_BASE);
      out.writeCode(distanceSymbol, DISTANCE_CODE_BITS);
      out.write(
        bestDistance - (DISTANCE_BASE[distanceSymbol] as number),
        DISTANCE_EXTRA[distanceSymbol] as number,
      );

      pos += bestLength;
      continue;
    }

    emitLiteral(input[pos] as number);
    pos++;
  }

  const { code, bits } = literalCode(END_OF_BLOCK);
  out.writeCode(code, bits);
  return out.finish();
}

export const HAS_DEFLATE_ENCODER = true;
