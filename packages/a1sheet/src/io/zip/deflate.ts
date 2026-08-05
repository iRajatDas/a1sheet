/**
 * DEFLATE *encoder* — deliberately absent.
 *
 * The POC ships a decoder only, so `makeZip` writes every entry with method 0
 * (STORE, uncompressed). The result is a valid .xlsx that Excel, Sheets, and
 * LibreOffice all open; it is just larger than it needs to be.
 *
 * To close this gap: implement `deflateRaw` here, then switch the `method` field
 * in ./zip.ts from 0 to 8 for the text-heavy entries (every XML part). A
 * fixed-Huffman-only encoder is enough — XML compresses well and there is no
 * need for optimal dynamic tables.
 *
 * This file exists to make the gap discoverable rather than leaving it as a
 * comment buried in the ZIP writer.
 */

export const HAS_DEFLATE_ENCODER = false;
