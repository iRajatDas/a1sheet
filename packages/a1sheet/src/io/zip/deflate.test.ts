/**
 * The DEFLATE encoder.
 *
 * Asserted against our own decoder rather than against expected bytes: a
 * compressor has no single correct output, and the property that matters is that
 * what comes back is what went in. `inflateRaw` was written first and against
 * real files from Excel, so it is a fair judge.
 */
import { describe, expect, test } from "bun:test";
import { deflateRaw } from "./deflate.js";
import { inflateRaw } from "./inflate.js";
import { listZipEntries, makeZip, readZipMember } from "./zip.js";

const encoder = new TextEncoder();
const decoder = new TextDecoder();

function roundTrip(text: string): string {
  return decoder.decode(inflateRaw(deflateRaw(encoder.encode(text))));
}

describe("round trips", () => {
  test("empty input", () => {
    expect(roundTrip("")).toBe("");
  });

  test("a single byte, which cannot be matched against anything", () => {
    expect(roundTrip("a")).toBe("a");
  });

  test("text with no repetition, which is all literals", () => {
    const text = "The quick brown fox jumps over the lazy dog.";
    expect(roundTrip(text)).toBe(text);
  });

  test("text that is nothing but repetition", () => {
    const text = "ab".repeat(5000);
    expect(roundTrip(text)).toBe(text);
  });

  test("a match at the maximum length", () => {
    // 258 bytes is the longest match DEFLATE can encode; a longer run has to be
    // split, and getting the boundary wrong corrupts everything after it.
    const text = `${"x".repeat(600)}|${"x".repeat(600)}`;
    expect(roundTrip(text)).toBe(text);
  });

  test("a match at the maximum distance", () => {
    // The window is 32768 bytes. A candidate further back than that must be
    // rejected, or the decoder reads from before the start of the stream.
    const text = `MARKER${"-".repeat(40000)}MARKER`;
    expect(roundTrip(text)).toBe(text);
  });

  test("every byte value, so nothing depends on the input being text", () => {
    const bytes = new Uint8Array(256 * 4);
    for (let i = 0; i < bytes.length; i++) bytes[i] = i % 256;
    expect([...inflateRaw(deflateRaw(bytes))]).toEqual([...bytes]);
  });

  test("realistic XML, which is what a workbook actually holds", () => {
    const xml = Array.from(
      { length: 3000 },
      (_, i) => `<c r="A${i}" s="2"><v>${i}</v></c>`,
    ).join("");
    expect(roundTrip(xml)).toBe(xml);
  });
});

describe("it actually compresses", () => {
  test("repetitive XML shrinks several times over", () => {
    const xml = Array.from(
      { length: 3000 },
      (_, i) => `<c r="A${i}" s="2"><v>${i}</v></c>`,
    ).join("");
    const raw = encoder.encode(xml);

    expect(deflateRaw(raw).length).toBeLessThan(raw.length / 3);
  });

  test("incompressible input is never made larger", () => {
    // Media in a workbook is already-compressed PNG, which DEFLATE cannot
    // improve on. Compressing unconditionally would make an image-heavy export
    // LARGER than storing it, so the writer compares and picks.
    //
    // The data is an xorshift sequence: deterministic, so the test cannot flake,
    // and genuinely unstructured — a counter mod 256 looks random and compresses
    // to almost nothing.
    const noise = new Uint8Array(8192);
    let state = 0x9e3779b9;
    for (let i = 0; i < noise.length; i++) {
      state ^= state << 13;
      state ^= state >>> 17;
      state ^= state << 5;
      noise[i] = state & 0xff;
    }

    const zip = makeZip([{ name: "noise.bin", data: noise }]);
    const member = listZipEntries(zip)[0];
    if (!member) throw new Error("no member");

    expect(member.compressed.length).toBeLessThanOrEqual(noise.length);
    expect([...readZipMember(member)]).toEqual([...noise]);
  });
});

describe("through the ZIP writer", () => {
  test("a compressed member reads back byte for byte", () => {
    const text = encoder.encode("<x>hello</x>".repeat(500));
    const zip = makeZip([{ name: "a.xml", data: text }]);
    const member = listZipEntries(zip)[0];

    expect(member?.method).toBe(8);
    expect([...readZipMember(member as never)]).toEqual([...text]);
  });

  test("a tiny member is left stored", () => {
    // Below a couple of hundred bytes the framing costs more than it saves, and
    // several parts in a workbook — the `.rels` files — are exactly that size.
    const zip = makeZip([{ name: "t.xml", data: encoder.encode("<a/>") }]);
    expect(listZipEntries(zip)[0]?.method).toBe(0);
  });

  test("an archive of both kinds reads back correctly", () => {
    const big = encoder.encode("<row/>".repeat(1000));
    const small = encoder.encode("<a/>");
    const zip = makeZip([
      { name: "big.xml", data: big },
      { name: "small.xml", data: small },
    ]);

    const members = listZipEntries(zip);
    expect(members).toHaveLength(2);
    expect([...readZipMember(members[0] as never)]).toEqual([...big]);
    expect([...readZipMember(members[1] as never)]).toEqual([...small]);
  });
});
