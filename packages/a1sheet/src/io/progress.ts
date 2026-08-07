/**
 * Cooperative pacing for long parses.
 *
 * A 38 MB xlsx is tens of millions of characters of XML. Parsing it in one
 * synchronous burst freezes the tab for seconds: no repaint, no scroll, no way to
 * change your mind. §6 of the engineering rules requires long work to be async,
 * cancellable, and to report progress — this module is the one mechanism that
 * provides all three, so the readers stay readable.
 *
 * A reader calls `pacer.checkpoint(...)` every few thousand units of work. The
 * checkpoint:
 *
 *   1. throws `AbortedError` if the caller's signal has fired,
 *   2. yields to the event loop if this task has held the thread for a frame,
 *   3. reports progress, rate-limited so a React consumer is not re-rendered
 *      thousands of times.
 *
 * On a small file no checkpoint ever exceeds the frame budget, so nothing yields
 * and the read costs the same as the synchronous version did. The cost is paid
 * only by files big enough to need it.
 */
import { AbortedError } from "../errors.js";

/** The stages a read passes through, in order. Progress weights follow this order. */
export const READ_PHASES = ["decompressing", "parsing"] as const;

export type ReadPhase = (typeof READ_PHASES)[number];

/**
 * Share of total progress each phase accounts for. Rough but stable: inflating is
 * fast per byte, XML scanning is not. Must sum to 1.
 */
const PHASE_WEIGHTS: Readonly<Record<ReadPhase, number>> = {
  decompressing: 0.3,
  parsing: 0.7,
} as const;

export interface ReadProgress {
  phase: ReadPhase;
  /** Fraction of the whole read that is done, 0..1. Never decreases. */
  ratio: number;
  /** What is being worked on right now — a ZIP member or a sheet name. */
  detail: string;
}

/** Accepted by every long-running read. Both fields are optional; both are honored. */
export interface AsyncReadOptions {
  /** Aborting rejects the read with `AbortedError` at the next checkpoint. */
  signal?: AbortSignal;
  /** Called at most once per frame, plus once per 1% of progress. */
  onProgress?: (progress: ReadProgress) => void;
}

export interface Pacer {
  /**
   * Abort check, event-loop yield, and progress report.
   *
   * Call this per chunk of work — a few thousand cells, one ZIP member — never per
   * cell: every call allocates a promise, and awaiting one per cell would cost more
   * than the parsing does.
   *
   * @param localRatio Progress within `phase`, 0..1.
   */
  checkpoint(phase: ReadPhase, localRatio: number, detail: string): Promise<void>;

  /**
   * Reports 1 unconditionally. Call once when the read completes.
   *
   * Checkpoints are rate-limited and land on chunk boundaries, so the last one
   * lands short of the end — a bar left sitting at 99% for no reason. This closes
   * it, and guarantees the contract that a successful read's final report is 1.
   */
  finish(detail: string): void;
}

/**
 * Longest the parse loop may hold the thread before yielding — one 60 Hz frame.
 *
 * Not lower, though a scheduler like React's would use ~5 ms. Two measurements on
 * a 37 MB, 600k-cell workbook set this:
 *
 *   - Pacing is not free. Unpaced the read takes ~1.1 s; paced it takes ~1.7 s.
 *     Almost none of that is the yield itself (4 ms across 135 yields) — it is the
 *     runtime finally getting to run garbage collection, which a read that never
 *     releases the thread defers until it is over. Halving the budget roughly
 *     doubles the yields and the overhead with them.
 *   - Lowering it buys little. A few steps cannot be chunked at all — inflating one
 *     large ZIP member, decoding it to a string — and those alone block ~50 ms. A
 *     budget below that floor makes the parse loop more granular than the rest of
 *     the read already is.
 */
const FRAME_BUDGET_MS = 16;

/** Smallest progress change worth reporting. Caps report count at ~100 per read. */
const PROGRESS_STEP = 0.01;

/**
 * Counts non-overlapping occurrences of `needle` in `haystack`.
 *
 * Here because a progress bar needs a denominator and the readers can only get one
 * by pre-scanning. `indexOf` is native and vectorized, so this sweep costs a small
 * fraction of the parse it measures — a per-character loop would not.
 */
export function countOccurrences(haystack: string, needle: string): number {
  let count = 0;
  let at = haystack.indexOf(needle);
  while (at !== -1) {
    count++;
    at = haystack.indexOf(needle, at + needle.length);
  }
  return count;
}

/** Cumulative weight of every phase before this one. */
function phaseOffset(phase: ReadPhase): number {
  let offset = 0;
  for (const p of READ_PHASES) {
    if (p === phase) break;
    offset += PHASE_WEIGHTS[p];
  }
  return offset;
}

function clamp01(n: number): number {
  if (!Number.isFinite(n)) return 0;
  return Math.min(1, Math.max(0, n));
}

interface Yielder {
  /** Hands the thread back so the host can paint and run pending input. */
  yieldToEventLoop(): Promise<void>;
  /** Releases anything the strategy holds. Safe to call more than once. */
  dispose(): void;
}

interface YieldGlobals {
  scheduler?: { yield?: () => Promise<void> };
  setImmediate?: (callback: () => void) => unknown;
  MessageChannel?: new () => {
    port1: { onmessage: (() => void) | null; close(): void };
    port2: { postMessage(value: unknown): void; close(): void };
  };
}

/**
 * Picks the cheapest way to reach the next macrotask, once per read.
 *
 * The ladder exists because `setTimeout(0)` is not free and a read yields on the
 * order of a hundred times. Hosts clamp it to about 1 ms, and browsers clamp it to
 * 4 ms once timers nest more than five deep — which is exactly the shape of a
 * paced loop. Measured here: `setTimeout` 1.16 ms per yield, `setImmediate`
 * 0.003 ms, `MessageChannel` 0.011 ms. Over 135 yields that is the difference
 * between a rounding error and half a second.
 *
 * A microtask (`queueMicrotask`, `await null`) is not an option: it runs before
 * the host gets control back, so it would cancel and paint nothing.
 */
function createYielder(): Yielder {
  const host = globalThis as YieldGlobals;

  // Chromium. Resumes at the front of the queue, so unrelated timers cannot
  // starve the read.
  const schedulerYield = host.scheduler?.yield;
  if (typeof schedulerYield === "function") {
    const scheduler = host.scheduler as { yield: () => Promise<void> };
    return { yieldToEventLoop: () => scheduler.yield(), dispose() {} };
  }

  // Node, Bun, Workers under those runtimes. A true macrotask with no clamp.
  const setImmediateFn = host.setImmediate;
  if (typeof setImmediateFn === "function") {
    return {
      yieldToEventLoop: () => new Promise((resolve) => setImmediateFn(resolve)),
      dispose() {},
    };
  }

  // Safari and Firefox. A message post is a macrotask and is not clamped; this is
  // the same trick React's scheduler uses, and for the same reason.
  const MessageChannelCtor = host.MessageChannel;
  if (typeof MessageChannelCtor === "function") {
    // Built on first use so a read that never yields never allocates a port pair.
    let channel: InstanceType<typeof MessageChannelCtor> | null = null;
    let pending: (() => void) | null = null;
    return {
      yieldToEventLoop() {
        if (!channel) {
          channel = new MessageChannelCtor();
          channel.port1.onmessage = () => {
            const resume = pending;
            pending = null;
            resume?.();
          };
        }
        const port2 = channel.port2;
        return new Promise((resolve) => {
          pending = resolve;
          port2.postMessage(null);
        });
      },
      dispose() {
        channel?.port1.close();
        channel?.port2.close();
        channel = null;
      },
    };
  }

  return {
    yieldToEventLoop: () => new Promise((resolve) => setTimeout(resolve, 0)),
    dispose() {},
  };
}

/**
 * Builds the pacer a read threads through its loops. Returns a no-progress,
 * no-abort pacer that still yields when given no options — yielding is about
 * responsiveness, which nobody has to opt into.
 */
/**
 * `AsyncReadOptions` plus the clock, which is internal.
 *
 * Not on the public options: a consumer has no reason to move the reader's clock,
 * and a test that races the real one is a test that fails on a fast machine.
 */
export interface PacerOptions extends AsyncReadOptions {
  now?: () => number;
}

export function createPacer(options: PacerOptions = {}): Pacer {
  const { signal, onProgress } = options;
  const yielder = createYielder();
  // Injectable so a test can decide when the budget is spent instead of racing
  // a real clock. `readWorkbookFile` never passes one.
  const now = options.now ?? Date.now;
  let deadline = now() + FRAME_BUDGET_MS;
  let ratio = 0;
  let reportedRatio = -1;
  let reportedPhase: ReadPhase | null = null;

  return {
    async checkpoint(phase, localRatio, detail) {
      if (signal?.aborted) {
        yielder.dispose();
        throw new AbortedError("Reading the spreadsheet");
      }

      // Monotonic: a phase that overruns its estimate must not walk the bar back.
      ratio = Math.max(
        ratio,
        clamp01(phaseOffset(phase) + PHASE_WEIGHTS[phase] * clamp01(localRatio)),
      );

      const overBudget = now() >= deadline;
      const worthReporting =
        phase !== reportedPhase || ratio - reportedRatio >= PROGRESS_STEP;

      if (onProgress && (overBudget || worthReporting)) {
        reportedRatio = ratio;
        reportedPhase = phase;
        onProgress({ phase, ratio, detail });
      }

      if (!overBudget) return;
      await yielder.yieldToEventLoop();
      deadline = now() + FRAME_BUDGET_MS;
    },

    finish(detail) {
      yielder.dispose();
      ratio = 1;
      reportedRatio = 1;
      const phase = READ_PHASES[READ_PHASES.length - 1] as ReadPhase;
      reportedPhase = phase;
      onProgress?.({ phase, ratio, detail });
    },
  };
}
