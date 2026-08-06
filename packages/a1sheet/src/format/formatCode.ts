/**
 * Rendering a value through an Excel number-format code.
 *
 * A format code is up to four sections separated by `;` — positive, negative,
 * zero, text — and each section is a pattern of literal characters mixed with
 * placeholders: `0` and `#` for digits, `.` for the decimal point, `,` for
 * thousands, `y`/`m`/`d`/`h`/`s` for date parts, `@` for the text itself.
 *
 * This covers what spreadsheets actually contain: signs, fixed decimals,
 * thousands separators, percentages, currency prefixes, dates, times, and colour
 * hints. It does NOT implement fractions (`# ?/?`), scientific notation beyond
 * passing it through, repeat (`*`), or width padding (`_`) — those degrade to the
 * plain number rather than to an error, so an unhandled code never breaks a cell.
 */

/** A parsed code, ready to apply. Cheap to build but built once per code. */
export interface NumberFormat {
  sections: readonly string[];
  /** Colour named in the applicable section, e.g. `[Red]`. */
  colorFor(value: number): string | undefined;
}

const SECTION_SPLIT = /;/;

/** `[Red]`, `[Blue]`, and friends. Excel allows eight named colours. */
const COLOR_NAMES: Record<string, string> = {
  black: "#000000",
  blue: "#0000ff",
  cyan: "#00ffff",
  green: "#008000",
  magenta: "#ff00ff",
  red: "#ff0000",
  white: "#ffffff",
  yellow: "#ffff00",
};

const BRACKETED = /\[([^\]]*)\]/g;

/**
 * Strips `[...]` directives from a section, returning the pattern and any colour.
 *
 * `[h]` and `[mm]` are elapsed-time markers rather than directives and must
 * survive, so a bracket whose contents are only time letters is left in place.
 */
function stripDirectives(section: string): {
  pattern: string;
  color?: string;
} {
  let color: string | undefined;
  const pattern = section.replace(BRACKETED, (whole, inner: string) => {
    if (/^[hms]+$/i.test(inner)) return whole;
    const named = COLOR_NAMES[inner.toLowerCase()];
    if (named) {
      color = named;
      return "";
    }
    // A condition like [>100] or a locale like [$-409] is not applied; dropping
    // it renders the rest of the section, which beats rendering nothing.
    return "";
  });
  return color === undefined ? { pattern } : { pattern, color };
}

export function parseFormatCode(code: string): NumberFormat {
  const sections = code.split(SECTION_SPLIT);
  return {
    sections,
    colorFor(value: number) {
      return stripDirectives(sectionFor(sections, value)).color;
    },
  };
}

/** Positive, negative, zero — falling back to the first section when absent. */
function sectionFor(sections: readonly string[], value: number): string {
  const positive = sections[0] ?? "";
  if (value < 0) return sections[1] ?? positive;
  if (value === 0) return sections[2] ?? positive;
  return positive;
}

const DATE_LETTERS = /[ymdhs]/i;
/** `mmm`/`mmmm` and `AM/PM` are the only letter runs that are not digits. */
const TEXT_SECTION_INDEX = 3;

/**
 * True when a code renders a date rather than a number.
 *
 * Testing for date letters alone is not enough: `"0.00"` has none, but a currency
 * code like `"$#,##0"` has an `m`-free `d`-free body while `[Red]` contains a `d`.
 * Directives are stripped first, and quoted literals with them.
 */
export function isDateFormat(code: string): boolean {
  const body = stripDirectives(sectionFor(code.split(SECTION_SPLIT), 1)).pattern;
  return DATE_LETTERS.test(stripLiterals(body));
}

const QUOTED = /"[^"]*"|\\./g;

function stripLiterals(pattern: string): string {
  return pattern.replace(QUOTED, "");
}

const PERCENT_SCALE = 100;
const MS_PER_DAY = 86400000;
const MINUTES_PER_HOUR = 60;
const SECONDS_PER_MINUTE = 60;
const HOURS_PER_DAY = 24;

const MONTH_NAMES = [
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
] as const;
const DAY_NAMES = [
  "Sunday",
  "Monday",
  "Tuesday",
  "Wednesday",
  "Thursday",
  "Friday",
  "Saturday",
] as const;
const SHORT_NAME_LENGTH = 3;

interface DateParts {
  year: number;
  month: number;
  day: number;
  hours: number;
  minutes: number;
  seconds: number;
  weekday: number;
  /** Whole days, for the elapsed-time forms. */
  totalDays: number;
}

/** A day serial in this engine's Unix-based epoch, split into parts in UTC. */
function dateParts(serial: number): DateParts {
  const d = new Date(Math.round(serial * MS_PER_DAY));
  return {
    year: d.getUTCFullYear(),
    month: d.getUTCMonth() + 1,
    day: d.getUTCDate(),
    hours: d.getUTCHours(),
    minutes: d.getUTCMinutes(),
    seconds: d.getUTCSeconds(),
    weekday: d.getUTCDay(),
    totalDays: serial,
  };
}

function pad(n: number, width: number): string {
  return String(Math.abs(n)).padStart(width, "0");
}

/**
 * Renders a date/time pattern.
 *
 * The `m` ambiguity is the whole difficulty: it means month, except immediately
 * after an hour or before a second, where it means minute. Resolved by looking at
 * the nearest date letter on either side, which is the rule Excel documents.
 */
function renderDate(pattern: string, parts: DateParts, use12Hour: boolean): string {
  let out = "";
  let i = 0;
  while (i < pattern.length) {
    const ch = pattern[i] as string;

    if (ch === '"') {
      const end = pattern.indexOf('"', i + 1);
      out += pattern.slice(i + 1, end === -1 ? pattern.length : end);
      i = end === -1 ? pattern.length : end + 1;
      continue;
    }
    if (ch === "\\") {
      out += pattern[i + 1] ?? "";
      i += 2;
      continue;
    }
    if (pattern.startsWith("AM/PM", i) || pattern.startsWith("am/pm", i)) {
      out += parts.hours < HOURS_PER_DAY / 2 ? "AM" : "PM";
      i += "AM/PM".length;
      continue;
    }
    if (ch === "[") {
      const end = pattern.indexOf("]", i);
      const inner = pattern.slice(i + 1, end === -1 ? pattern.length : end);
      out += elapsed(inner, parts);
      i = end === -1 ? pattern.length : end + 1;
      continue;
    }

    const run = runLength(pattern, i);
    const token = pattern.slice(i, i + run).toLowerCase();

    if (token[0] === "y") {
      out += run <= 2 ? pad(parts.year % PERCENT_SCALE, 2) : String(parts.year);
    } else if (token[0] === "d") {
      if (run === 1 || run === 2)
        out += run === 1 ? String(parts.day) : pad(parts.day, 2);
      else {
        const name = DAY_NAMES[parts.weekday] as string;
        out += run === SHORT_NAME_LENGTH ? name.slice(0, SHORT_NAME_LENGTH) : name;
      }
    } else if (token[0] === "h") {
      const h = use12Hour
        ? parts.hours % (HOURS_PER_DAY / 2) || HOURS_PER_DAY / 2
        : parts.hours;
      out += run === 1 ? String(h) : pad(h, 2);
    } else if (token[0] === "s") {
      out += run === 1 ? String(parts.seconds) : pad(parts.seconds, 2);
    } else if (token[0] === "m") {
      out += isMinute(pattern, i, run)
        ? run === 1
          ? String(parts.minutes)
          : pad(parts.minutes, 2)
        : monthText(run, parts.month);
    } else {
      out += pattern.slice(i, i + run);
    }
    i += run;
  }
  return out;
}

function monthText(run: number, month: number): string {
  if (run === 1) return String(month);
  if (run === 2) return pad(month, 2);
  const name = MONTH_NAMES[month - 1] as string;
  if (run === SHORT_NAME_LENGTH) return name.slice(0, SHORT_NAME_LENGTH);
  // "mmmmm" is the single-letter form, e.g. J for January.
  return run > 4 ? (name[0] as string) : name;
}

function elapsed(marker: string, parts: DateParts): string {
  const total = parts.totalDays;
  if (/^h+$/i.test(marker)) return String(Math.floor(total * HOURS_PER_DAY));
  if (/^m+$/i.test(marker)) {
    return String(Math.floor(total * HOURS_PER_DAY * MINUTES_PER_HOUR));
  }
  if (/^s+$/i.test(marker)) {
    return String(
      Math.floor(total * HOURS_PER_DAY * MINUTES_PER_HOUR * SECONDS_PER_MINUTE),
    );
  }
  return "";
}

function runLength(pattern: string, at: number): number {
  const ch = pattern[at];
  let n = 1;
  while (pattern[at + n] === ch) n++;
  return n;
}

/**
 * Whether an `m` run means minutes: true when the nearest significant neighbour
 * is an hour before it or a second after it.
 */
function isMinute(pattern: string, at: number, run: number): boolean {
  const before = pattern.slice(0, at).match(/[ydhs](?=[^ydhs]*$)/i)?.[0];
  if (before?.toLowerCase() === "h") return true;
  const after = pattern.slice(at + run).match(/[ydhs]/i)?.[0];
  return after?.toLowerCase() === "s";
}

/** Does the pattern ask for a 12-hour clock? */
function has12Hour(pattern: string): boolean {
  return /am\/pm|a\/p/i.test(pattern);
}

/**
 * Renders a numeric pattern: `#,##0.00`, `+0`, `$#,##0`, `0%`.
 *
 * The pattern's digit placeholders decide the decimal places and whether
 * thousands are grouped; everything that is not a placeholder is a literal and
 * comes out where it sits, which is what makes `"+0;-0;0"` produce `+45`.
 */
function renderNumber(pattern: string, value: number, locale?: string): string {
  const scaled = pattern.includes("%") ? value * PERCENT_SCALE : value;
  const body = stripLiterals(pattern);
  const decimals = body.split(".")[1]?.match(/[0#]/g)?.length ?? 0;
  const grouped = /[#0],[#0]/.test(body.split(".")[0] ?? "");
  const minIntegerDigits = (body.split(".")[0]?.match(/0/g) ?? []).length;

  const magnitude = Math.abs(scaled);
  const digits = magnitude.toLocaleString(locale, {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals,
    minimumIntegerDigits: Math.max(1, minIntegerDigits),
    useGrouping: grouped,
  });

  // Walk the pattern and emit literals in place, substituting the digits for the
  // first placeholder run. A sign in the pattern is a literal — that is how the
  // negative section of "+0;-0;0" produces its own minus.
  let out = "";
  let placed = false;
  let i = 0;
  while (i < pattern.length) {
    const ch = pattern[i] as string;
    if (ch === '"') {
      const end = pattern.indexOf('"', i + 1);
      out += pattern.slice(i + 1, end === -1 ? pattern.length : end);
      i = end === -1 ? pattern.length : end + 1;
      continue;
    }
    if (ch === "\\") {
      out += pattern[i + 1] ?? "";
      i += 2;
      continue;
    }
    if (ch === "_" || ch === "*") {
      // Width padding and fill: skip the directive and the character it applies to.
      i += 2;
      continue;
    }
    if (/[#0?.,]/.test(ch)) {
      if (!placed) {
        out += digits;
        placed = true;
      }
      i++;
      continue;
    }
    if (ch === "%") {
      out += "%";
      i++;
      continue;
    }
    out += ch;
    i++;
  }

  return placed ? out : out + digits;
}

/**
 * The text section: `@` stands for the value and everything else is a literal,
 * with the same quoting rules as a numeric pattern. `"<"@">"` wraps in angle
 * brackets, and the quotes are not part of the output.
 */
function renderText(pattern: string, value: string): string {
  let out = "";
  let i = 0;
  while (i < pattern.length) {
    const ch = pattern[i] as string;
    if (ch === '"') {
      const end = pattern.indexOf('"', i + 1);
      out += pattern.slice(i + 1, end === -1 ? pattern.length : end);
      i = end === -1 ? pattern.length : end + 1;
      continue;
    }
    if (ch === "\\") {
      out += pattern[i + 1] ?? "";
      i += 2;
      continue;
    }
    out += ch === "@" ? value : ch;
    i++;
  }
  return out;
}

export interface ApplyOptions {
  locale?: string;
}

/**
 * Renders a value through a format code.
 *
 * Returns null when the code cannot be applied to this value — a text value
 * against a numeric code, or a code with no section for it — so the caller can
 * fall back rather than show something wrong.
 */
export function applyFormatCode(
  code: string,
  value: string | number,
  opts: ApplyOptions = {},
): string | null {
  const sections = code.split(SECTION_SPLIT);

  if (typeof value === "string") {
    const textSection = sections[TEXT_SECTION_INDEX];
    if (!textSection) return null;
    return renderText(stripDirectives(textSection).pattern, value);
  }

  const { pattern } = stripDirectives(sectionFor(sections, value));
  if (pattern === "") return "";
  if (pattern.trim() === "@") return String(value);

  if (DATE_LETTERS.test(stripLiterals(pattern))) {
    return renderDate(pattern, dateParts(value), has12Hour(pattern));
  }

  // A negative value rendered by its own section must not also carry a minus,
  // since the section supplies whatever sign it wants.
  const usesOwnSection = value < 0 && sections[1] !== undefined;
  const rendered = renderNumber(pattern, value, opts.locale);
  return usesOwnSection || value >= 0 ? rendered : `-${rendered}`;
}
