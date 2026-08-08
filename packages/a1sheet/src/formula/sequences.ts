/**
 * Named cyclic sequences for the fill handle — weekdays and months.
 *
 * Detected before numeric/cyclic rules in `extrapolateSeries`. Matching is
 * case-preserving: the output keeps the casing of the last source value's form.
 */
export interface Sequence {
  readonly id: string;
  readonly values: readonly string[];
}

export const SEQUENCES = [
  {
    id: "weekday",
    values: [
      "Monday",
      "Tuesday",
      "Wednesday",
      "Thursday",
      "Friday",
      "Saturday",
      "Sunday",
    ],
  },
  {
    id: "weekday-abbr",
    values: ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"],
  },
  {
    id: "month",
    values: [
      "January",
      "February",
      "March",
      "April",
      "May",
      "June",
      "July",
      "August",
      "September",
      "October",
      "November",
      "December",
    ],
  },
  {
    id: "month-abbr",
    values: [
      "Jan",
      "Feb",
      "Mar",
      "Apr",
      "May",
      "Jun",
      "Jul",
      "Aug",
      "Sep",
      "Oct",
      "Nov",
      "Dec",
    ],
  },
] as const satisfies readonly Sequence[];

function indexIn(seq: Sequence, value: string): number {
  const lower = value.toLowerCase();
  return seq.values.findIndex((v) => v.toLowerCase() === lower);
}

/**
 * If every source value belongs to the same sequence and advances by a constant
 * step (wrapping), return that sequence and step. Otherwise undefined.
 */
export function matchSequence(
  values: readonly string[],
): { sequence: Sequence; step: number; lastIndex: number } | undefined {
  if (values.length === 0) return undefined;
  if (values.some((v) => v === "")) return undefined;

  for (const sequence of SEQUENCES) {
    const indexes = values.map((v) => indexIn(sequence, v));
    if (indexes.some((i) => i < 0)) continue;

    const first = indexes[0] as number;
    if (values.length === 1) {
      return { sequence, step: 1, lastIndex: first };
    }

    const second = indexes[1] as number;
    const len = sequence.values.length;
    const step = (second - first + len) % len;
    if (step === 0) continue;

    let ok = true;
    for (let i = 1; i < indexes.length; i++) {
      const prev = indexes[i - 1] as number;
      const cur = indexes[i] as number;
      if ((prev + step) % len !== cur) {
        ok = false;
        break;
      }
    }
    if (!ok) continue;

    return {
      sequence,
      step,
      lastIndex: indexes[indexes.length - 1] as number,
    };
  }

  return undefined;
}

/** Continue a matched sequence for `count` steps, preserving the last value's case. */
export function extendSequence(
  values: readonly string[],
  count: number,
): string[] | undefined {
  const match = matchSequence(values);
  if (!match) return undefined;

  const { sequence, step, lastIndex } = match;
  const lastRaw = values[values.length - 1] ?? "";
  const template = sequence.values[lastIndex] ?? lastRaw;
  const out: string[] = [];
  const len = sequence.values.length;

  for (let i = 0; i < count; i++) {
    const idx = (lastIndex + step * (i + 1)) % len;
    const next = sequence.values[idx] ?? "";
    out.push(applyCase(template, lastRaw, next));
  }
  return out;
}

function applyCase(template: string, sample: string, next: string): string {
  if (sample === sample.toUpperCase() && sample !== sample.toLowerCase()) {
    return next.toUpperCase();
  }
  if (sample === sample.toLowerCase()) return next.toLowerCase();
  // Title / mixed — keep the sequence's canonical form (template casing).
  if (sample[0] === sample[0]?.toLowerCase()) {
    return next.charAt(0).toLowerCase() + next.slice(1);
  }
  return template[0] === next[0] ? next : next;
}
