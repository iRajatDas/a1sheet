/**
 * In-cell images.
 *
 * `=IMAGE("https://…")` does not store a picture in the cell. Excel stores the
 * error value `#VALUE!` there — so a reader that trusts the cell shows `#VALUE!`,
 * which is technically faithful and completely wrong — plus a `vm` attribute
 * pointing into a chain of five parts that ends at a PNG:
 *
 *   <c r="J2" t="e" vm="1"> … <v>#VALUE!</v>
 *     -> metadata.xml <valueMetadata><bk>[vm-1] <rc t="2" v="0"/>
 *        t indexes <metadataTypes>, and must be the XLRICHVALUE one
 *     -> metadata.xml <futureMetadata name="XLRICHVALUE"><bk>[0] <xlrd:rvb i="0"/>
 *     -> richData/rdrichvalue.xml <rv>[0] first <v> = web-image index
 *     -> richData/rdRichValueWebImage.xml <webImageSrd>[i]
 *        <address r:id> is the source URL, <blip r:id> the embedded copy
 *     -> richData/_rels/rdRichValueWebImage.xml.rels resolves both
 *
 * The embedded copy is preferred over the URL: it is already in the file, so it
 * needs no network, works offline, and cannot be swapped out from under the
 * document by whoever controls the address.
 */
import type { CellImage } from "../../model/types.js";
import { findElement, findElements } from "./xml.js";

/** The metadata type whose blocks point at rich values. */
const RICH_VALUE_TYPE = "XLRICHVALUE";

/**
 * Raster types embedded as a data URI.
 *
 * A deliberate allow-list rather than a block-list. An `<img>` will not run
 * script from an SVG, but the whole point of an allow-list is not having to be
 * sure of that for every format a file might carry; these five cover what Excel
 * actually writes for a web image.
 */
const EMBEDDABLE: Record<string, string> = {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
  webp: "image/webp",
  bmp: "image/bmp",
};

/**
 * Total distinct image bytes embedded per workbook.
 *
 * A cap because base64 costs a third again on top and the result is held for the
 * life of the sheet — a workbook of photographs would otherwise pull tens of
 * megabytes of string into memory. 16 MiB covers a set of logos at the sizes
 * Excel's own image search produces (the sample workbook's twenty crests are
 * 1920px PNGs, about 7 MiB together) with room to spare.
 *
 * Past the cap an image falls back to its source URL, which costs a request but
 * not the memory. Counted over distinct parts, so a picture used in a hundred
 * cells is charged once.
 */
const EMBED_BUDGET_BYTES = 16 * 1024 * 1024;

/** Bytes per `btoa` call. Spreading a whole image would overflow the stack. */
const BASE64_CHUNK = 0x8000;

function toBase64(bytes: Uint8Array): string {
  let binary = "";
  for (let i = 0; i < bytes.length; i += BASE64_CHUNK) {
    binary += String.fromCharCode(
      ...bytes.subarray(i, Math.min(i + BASE64_CHUNK, bytes.length)),
    );
  }
  return btoa(binary);
}

interface Rels {
  /** Relationship id -> target, as written (may be `../media/image1.png`). */
  targets: Record<string, string>;
}

function parseRels(xml: string | undefined): Rels {
  const targets: Record<string, string> = {};
  if (!xml) return { targets };
  for (const rel of findElements(xml, "Relationship")) {
    const id = rel.attrs.Id;
    const target = rel.attrs.Target;
    if (id && target) targets[id] = target;
  }
  return { targets };
}

/**
 * `vm` value -> rich-value index.
 *
 * Two hops, both by index, and the first is 1-based while the second is not.
 * A `vm` whose metadata block points at a type other than XLRICHVALUE is not an
 * image — dynamic-array metadata shares the same attribute.
 */
function parseValueMetadata(xml: string | undefined): Map<number, number> {
  const out = new Map<number, number>();
  if (!xml) return out;

  const types = findElements(xml, "metadataType").map((t) => t.attrs.name);
  // `rc/@t` is 1-based over <metadataTypes>.
  const richTypeIndex = types.indexOf(RICH_VALUE_TYPE) + 1;
  if (richTypeIndex === 0) return out;

  const futureBlocks = futureMetadataBlocks(xml);

  const valueMetadata = findElement(xml, "valueMetadata");
  if (!valueMetadata) return out;

  for (const [i, bk] of findElements(valueMetadata.inner, "bk").entries()) {
    const rc = findElement(bk.inner, "rc");
    if (!rc) continue;
    if (Number.parseInt(rc.attrs.t ?? "", 10) !== richTypeIndex) continue;
    const futureIndex = Number.parseInt(rc.attrs.v ?? "", 10);
    const richIndex = futureBlocks[futureIndex];
    // `vm` on a cell is 1-based over these blocks.
    if (richIndex !== undefined) out.set(i + 1, richIndex);
  }
  return out;
}

/** The `<xlrd:rvb i>` of each XLRICHVALUE future-metadata block, in order. */
function futureMetadataBlocks(xml: string): number[] {
  for (const fm of findElements(xml, "futureMetadata")) {
    if (fm.attrs.name !== RICH_VALUE_TYPE) continue;
    return findElements(fm.inner, "bk").map((bk) => {
      const rvb = findElement(bk.inner, "xlrd:rvb");
      const i = Number.parseInt(rvb?.attrs.i ?? "", 10);
      return Number.isFinite(i) ? i : -1;
    });
  }
  return [];
}

/** Rich-value index -> the web-image index its first field holds. */
function parseRichValues(xml: string | undefined): number[] {
  if (!xml) return [];
  return findElements(xml, "rv").map((rv) => {
    const first = findElements(rv.inner, "v")[0];
    const n = Number.parseInt(first?.inner ?? "", 10);
    return Number.isFinite(n) ? n : -1;
  });
}

interface WebImage {
  addressRel?: string;
  blipRel?: string;
}

function parseWebImages(xml: string | undefined): WebImage[] {
  if (!xml) return [];
  return findElements(xml, "webImageSrd").map((srd) => {
    const address = findElement(srd.inner, "address")?.attrs["r:id"];
    const blip = findElement(srd.inner, "blip")?.attrs["r:id"];
    return {
      ...(address ? { addressRel: address } : {}),
      ...(blip ? { blipRel: blip } : {}),
    };
  });
}

export interface ImageSource {
  /** Every inflated part, by name. */
  files: Record<string, Uint8Array>;
  /** Decodes a part to text, or undefined when it is absent. */
  read(name: string | undefined): string | undefined;
  /** Every part name, for case-insensitive lookup. */
  names: readonly string[];
}

/**
 * The workbook's rich-value images, keyed by the `vm` value a cell carries.
 *
 * Built once per workbook rather than per sheet: the chain is workbook-level, and
 * cells on any sheet index into the same table.
 */
export interface ImageTable {
  /** `vm` -> image, for the cells that have one. */
  byValueMetadata: Map<number, CellImage>;
}

export const EMPTY_IMAGE_TABLE: ImageTable = { byValueMetadata: new Map() };

function findPart(names: readonly string[], suffix: string): string | undefined {
  const wanted = suffix.toLowerCase();
  return names.find((n) => n.toLowerCase().endsWith(wanted));
}

/** Resolves a relationship target against `xl/richData/`, where the rels live. */
function mediaPath(target: string): string {
  return `xl/${target.replace(/^\.\.\//, "")}`;
}

export function parseImageTable(source: ImageSource): ImageTable {
  const { names, read, files } = source;

  const metadata = read(findPart(names, "xl/metadata.xml"));
  const vmToRich = parseValueMetadata(metadata);
  if (vmToRich.size === 0) return EMPTY_IMAGE_TABLE;

  const richValues = parseRichValues(
    read(findPart(names, "richdata/rdrichvalue.xml")),
  );
  const webImages = parseWebImages(
    read(findPart(names, "richdata/rdrichvaluewebimage.xml")),
  );
  const rels = parseRels(
    read(findPart(names, "richdata/_rels/rdrichvaluewebimage.xml.rels")),
  );

  const byValueMetadata = new Map<number, CellImage>();
  // One data URI per media part, not per cell. A workbook typically points many
  // cells at the same picture — twenty crests across a hundred and forty cells in
  // the sample file — and encoding each reference separately would multiply both
  // the work and the memory by the number of references.
  const encoded = new Map<string, string | null>();
  let budget = EMBED_BUDGET_BYTES;

  const embed = (target: string): string | null => {
    const cached = encoded.get(target);
    if (cached !== undefined) return cached;

    const mime = EMBEDDABLE[target.split(".").pop()?.toLowerCase() ?? ""];
    const bytes = mime ? files[mediaPath(target)] : undefined;
    if (!mime || !bytes || bytes.length > budget) {
      encoded.set(target, null);
      return null;
    }

    budget -= bytes.length;
    const src = `data:${mime};base64,${toBase64(bytes)}`;
    encoded.set(target, src);
    return src;
  };

  for (const [vm, richIndex] of vmToRich) {
    const imageIndex = richValues[richIndex];
    const image = imageIndex === undefined ? undefined : webImages[imageIndex];
    if (!image) continue;

    const address = image.addressRel ? rels.targets[image.addressRel] : undefined;
    const blip = image.blipRel ? rels.targets[image.blipRel] : undefined;
    const embedded = blip ? embed(blip) : null;

    // Embedded first, then the URL. An image with neither is dropped, which is
    // the only case where something in the file is silently not drawn.
    const src = embedded ?? address;
    if (!src) continue;
    byValueMetadata.set(vm, { src, ...(address ? { alt: address } : {}) });
  }

  return { byValueMetadata };
}
