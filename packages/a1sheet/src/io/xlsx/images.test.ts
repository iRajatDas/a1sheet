/**
 * In-cell images.
 *
 * The cell itself holds `#VALUE!` — Excel keeps the picture five parts away — so
 * a reader that trusts the cell shows an error where a crest should be, which is
 * faithful to the bytes and wrong about the document.
 */
import { describe, expect, test } from "bun:test";
import { imageUrlIn } from "../../formula/imageCall.js";
import { EMPTY_IMAGE_TABLE, type ImageSource, parseImageTable } from "./images.js";

const METADATA = `<metadata>
<metadataTypes count="2">
<metadataType name="XLDAPR"/><metadataType name="XLRICHVALUE"/>
</metadataTypes>
<futureMetadata name="XLRICHVALUE" count="2">
<bk><extLst><ext uri="{x}"><xlrd:rvb i="0"/></ext></extLst></bk>
<bk><extLst><ext uri="{x}"><xlrd:rvb i="1"/></ext></extLst></bk>
</futureMetadata>
<valueMetadata count="2">
<bk><rc t="2" v="0"/></bk>
<bk><rc t="2" v="1"/></bk>
</valueMetadata></metadata>`;

const RICH_VALUES = `<rvData count="2">
<rv s="0"><v>0</v><v>1</v><v>0</v><v>0</v></rv>
<rv s="0"><v>1</v><v>1</v><v>0</v><v>0</v></rv>
</rvData>`;

const WEB_IMAGES = `<webImagesSrd>
<webImageSrd><address r:id="rId1"/><blip r:id="rId2"/></webImageSrd>
<webImageSrd><address r:id="rId3"/><blip r:id="rId4"/></webImageSrd>
</webImagesSrd>`;

const RELS = `<Relationships>
<Relationship Id="rId1" Target="https://example.test/a.png" TargetMode="External"/>
<Relationship Id="rId2" Target="../media/image1.png"/>
<Relationship Id="rId3" Target="https://example.test/b.png" TargetMode="External"/>
<Relationship Id="rId4" Target="../media/image2.png"/>
</Relationships>`;

/** A one-pixel PNG's leading bytes are enough; nothing decodes the image. */
const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

interface Parts {
  [name: string]: string | Uint8Array;
}

function source(parts: Parts): ImageSource {
  const files: Record<string, Uint8Array> = {};
  const text: Record<string, string> = {};
  for (const [name, body] of Object.entries(parts)) {
    if (typeof body === "string") text[name] = body;
    else files[name] = body;
  }
  return {
    files,
    names: Object.keys(parts),
    read: (name) => (name === undefined ? undefined : text[name]),
  };
}

const FULL: Parts = {
  "xl/metadata.xml": METADATA,
  "xl/richData/rdrichvalue.xml": RICH_VALUES,
  "xl/richData/rdRichValueWebImage.xml": WEB_IMAGES,
  "xl/richData/_rels/rdRichValueWebImage.xml.rels": RELS,
  "xl/media/image1.png": PNG,
  "xl/media/image2.png": PNG,
};

describe("the rich-value chain", () => {
  test("a cell's vm resolves all the way to the embedded picture", () => {
    // vm -> valueMetadata -> futureMetadata -> rdrichvalue -> webImage -> rels.
    // Five indexed hops, and the first is 1-based while the rest are not.
    const { byValueMetadata } = parseImageTable(source(FULL));

    expect(byValueMetadata.get(1)?.src).toStartWith("data:image/png;base64,");
    expect(byValueMetadata.get(2)?.src).toStartWith("data:image/png;base64,");
    // vm is 1-based, so there is no entry 0.
    expect(byValueMetadata.get(0)).toBeUndefined();
  });

  test("the source URL becomes the alt text", () => {
    // It is the only description of the image the file contains.
    const { byValueMetadata } = parseImageTable(source(FULL));
    expect(byValueMetadata.get(1)?.alt).toBe("https://example.test/a.png");
  });

  test("the embedded copy is preferred over the URL", () => {
    // No request, works offline, and cannot be changed by whoever owns the address.
    const { byValueMetadata } = parseImageTable(source(FULL));
    expect(byValueMetadata.get(1)?.src).not.toBe("https://example.test/a.png");
  });

  test("an image with no embedded copy falls back to its URL", () => {
    const { byValueMetadata } = parseImageTable(
      source({ ...FULL, "xl/media/image1.png": undefined as never }),
    );
    expect(byValueMetadata.get(1)?.src).toBe("https://example.test/a.png");
  });

  test("one data URI is shared by every cell pointing at the same part", () => {
    // The sample workbook references twenty pictures from a hundred and forty
    // cells; encoding per reference would multiply the memory sevenfold.
    const shared = `<webImagesSrd>
<webImageSrd><address r:id="rId1"/><blip r:id="rId2"/></webImageSrd>
<webImageSrd><address r:id="rId3"/><blip r:id="rId2"/></webImageSrd>
</webImagesSrd>`;
    const { byValueMetadata } = parseImageTable(
      source({ ...FULL, "xl/richData/rdRichValueWebImage.xml": shared }),
    );

    expect(byValueMetadata.get(1)?.src).toBe(byValueMetadata.get(2)?.src as string);
  });

  test("a non-raster part is not embedded", () => {
    // An allow-list, so a format nobody has vetted never becomes a data URI.
    const svgRels = RELS.replace("../media/image1.png", "../media/image1.svg");
    const { byValueMetadata } = parseImageTable(
      source({
        ...FULL,
        "xl/richData/_rels/rdRichValueWebImage.xml.rels": svgRels,
        "xl/media/image1.svg": PNG,
      }),
    );

    expect(byValueMetadata.get(1)?.src).toBe("https://example.test/a.png");
  });
});

describe("a workbook with no images", () => {
  test("costs nothing beyond looking for the metadata part", () => {
    expect(parseImageTable(source({})).byValueMetadata.size).toBe(0);
    expect(EMPTY_IMAGE_TABLE.byValueMetadata.size).toBe(0);
  });

  test("dynamic-array metadata is not mistaken for an image", () => {
    // `vm` is shared with XLDAPR, which marks a spill anchor and has no picture.
    const dynamicOnly = METADATA.replace(/t="2"/g, 't="1"');
    const { byValueMetadata } = parseImageTable(
      source({ ...FULL, "xl/metadata.xml": dynamicOnly }),
    );
    expect(byValueMetadata.size).toBe(0);
  });
});

describe("recognizing an IMAGE call", () => {
  test("finds the URL, with or without Excel's modern-function prefix", () => {
    expect(imageUrlIn('=IMAGE("https://example.test/a.png")')).toBe(
      "https://example.test/a.png",
    );
    expect(imageUrlIn('=_xlfn.IMAGE("https://example.test/a.png")')).toBe(
      "https://example.test/a.png",
    );
  });

  test("ignores anything that is not an IMAGE call", () => {
    expect(imageUrlIn("=SUM(A1:A2)")).toBeNull();
    expect(imageUrlIn('IMAGE("x")')).toBeNull();
    expect(imageUrlIn("")).toBeNull();
  });

  test("tolerates whitespace and extra arguments", () => {
    expect(imageUrlIn('= IMAGE( "x.png" , "alt", 1)')).toBe("x.png");
  });
});
