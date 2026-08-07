/**
 * Cancellation and progress, tested through the public read API.
 *
 * The fixtures are built big enough that the reader must cross its frame budget —
 * a read that never yields would pass a cancellation test vacuously, because the
 * abort would land only after the work had already finished.
 */
import { describe, expect, test } from "bun:test";
import { readWorkbookFile } from "./index.js";
import { createPacer, type ReadProgress } from "./progress.js";
import { writeXlsx } from "./xlsx/write.js";

/** Rows chosen so parsing takes several frames on a fast machine. */
const BIG_ROWS = 20_000;

function bigCsvBlob(): Blob {
  const lines: string[] = [];
  for (let r = 0; r < BIG_ROWS; r++) {
    lines.push(`Item ${r},${r},${r * 2},some longer text value for row ${r}`);
  }
  return new Blob([lines.join("\n")], { type: "text/csv" });
}

function bigXlsxBlob(): Blob {
  const cells: Record<string, string> = {};
  for (let r = 0; r < BIG_ROWS; r++) {
    cells[`${r}_0`] = `Item ${r}`;
    cells[`${r}_1`] = String(r);
    cells[`${r}_2`] = `some longer text value for row ${r}`;
  }
  const bytes = writeXlsx([
    { name: "Big", cells, styles: {}, merges: [], namedRanges: {} },
  ]);
  return new Blob([bytes as BlobPart]);
}

describe("progress reporting", () => {
  test("a large CSV reports progress that ends at 1", async () => {
    const seen: ReadProgress[] = [];
    const result = await readWorkbookFile(bigCsvBlob(), {
      onProgress: (p) => seen.push(p),
    });

    expect(result.sheets[0]?.rows).toBe(BIG_ROWS);
    expect(seen.length).toBeGreaterThan(1);
    expect(seen.at(-1)?.ratio).toBeCloseTo(1, 5);
  });

  test("progress never goes backwards and stays within 0..1", async () => {
    const ratios: number[] = [];
    await readWorkbookFile(bigXlsxBlob(), {
      onProgress: (p) => ratios.push(p.ratio),
    });

    expect(ratios.length).toBeGreaterThan(1);
    for (const [i, ratio] of ratios.entries()) {
      expect(ratio).toBeGreaterThanOrEqual(0);
      expect(ratio).toBeLessThanOrEqual(1);
      if (i > 0) expect(ratio).toBeGreaterThanOrEqual(ratios[i - 1] as number);
    }
  });

  test("an xlsx passes through both phases in order", async () => {
    const phases: string[] = [];
    await readWorkbookFile(bigXlsxBlob(), {
      onProgress: (p) => {
        if (phases.at(-1) !== p.phase) phases.push(p.phase);
      },
    });

    expect(phases).toEqual(["decompressing", "parsing"]);
  });

  test("omitting onProgress is not an error", async () => {
    const result = await readWorkbookFile(bigCsvBlob());
    expect(result.format).toBe("csv");
  });
});

describe("cancellation", () => {
  test("aborting mid-parse rejects with code ABORTED", async () => {
    const controller = new AbortController();
    const pending = readWorkbookFile(bigCsvBlob(), {
      signal: controller.signal,
      // Abort on the first yield, which only happens once a frame is spent.
      onProgress: () => controller.abort(),
    });

    await expect(pending).rejects.toMatchObject({ code: "ABORTED" });
  });

  test("an already-aborted signal stops the read at its first checkpoint", async () => {
    const controller = new AbortController();
    controller.abort();

    await expect(
      readWorkbookFile(bigXlsxBlob(), { signal: controller.signal }),
    ).rejects.toMatchObject({ code: "ABORTED" });
  });

  test("aborting one read does not affect a concurrent one", async () => {
    const doomed = new AbortController();
    const cancelled = readWorkbookFile(bigCsvBlob(), {
      signal: doomed.signal,
      onProgress: () => doomed.abort(),
    });
    const survivor = readWorkbookFile(bigCsvBlob(), {
      signal: new AbortController().signal,
    });

    await expect(cancelled).rejects.toMatchObject({ code: "ABORTED" });
    expect((await survivor).sheets[0]?.rows).toBe(BIG_ROWS);
  });

  test("a cancelled read produces nothing — no partial workbook", async () => {
    const controller = new AbortController();
    let result: unknown = "untouched";
    try {
      result = await readWorkbookFile(bigCsvBlob(), {
        signal: controller.signal,
        onProgress: () => controller.abort(),
      });
    } catch {
      // expected
    }
    expect(result).toBe("untouched");
  });
});

describe("yielding", () => {
  test("a checkpoint past the frame budget hands the thread back", async () => {
    // Driven by an injected clock rather than by how long a real read happens to
    // take. The old version scheduled a timer and read a large CSV, which only
    // proved anything while the machine was slow enough for the read to overrun a
    // frame — so it passed on CI and failed at random on a fast laptop.
    //
    // Asserted by ORDERING rather than by a timer: yielding means the checkpoint
    // resolves on a later macrotask, so a microtask queued after it must run
    // first. A timer cannot see this — the yielder uses a MessageChannel, which
    // the event loop runs ahead of a setTimeout(0).
    let clock = 0;
    const pacer = createPacer({ now: () => clock });
    const order: string[] = [];

    const inBudget = pacer
      .checkpoint("parsing", 0.1, "x")
      .then(() => order.push("checkpoint"));
    await Promise.resolve().then(() => order.push("microtask"));
    await inBudget;
    expect(order).toEqual(["checkpoint", "microtask"]);

    order.length = 0;
    clock += 1000;
    const overBudget = pacer
      .checkpoint("parsing", 0.2, "x")
      .then(() => order.push("checkpoint"));
    await Promise.resolve().then(() => order.push("microtask"));
    await overBudget;
    expect(order).toEqual(["microtask", "checkpoint"]);

    // The yielder holds a MessageChannel open; without this the process never
    // exits.
    pacer.finish("done");
  });

  test("a real read still yields when it is big enough to need to", async () => {
    // The end-to-end version of the same property, kept because the unit test
    // above cannot show that `readWorkbookFile` actually threads a pacer through.
    // Asserted on the progress reports rather than on a timer: several reports
    // can only happen across several checkpoints.
    const seen: number[] = [];
    await readWorkbookFile(bigCsvBlob(), {
      onProgress: (p) => seen.push(p.ratio),
    });
    expect(seen.length).toBeGreaterThan(1);
  });
});
