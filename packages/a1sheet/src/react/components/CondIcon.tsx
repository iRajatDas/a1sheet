"use client";

/**
 * The icon a conditional-format icon-set rule draws.
 *
 * Excel names about twenty sets and draws them from its own resources. Rather
 * than reproduce each one, sets are grouped by what they mean — a traffic light,
 * an arrow, a flag, a rating — and drawn from the same three-colour palette,
 * because what a reader takes from an icon set is which band a value fell into,
 * not the exact glyph.
 *
 * Inline SVG on the same 24-unit grid as the rest of the icons, so it inherits
 * the cell's size and sits on the text's line.
 */
import type { ReactNode } from "react";

const ICON_SIZE = 12;
const VIEW_BOX = "0 0 24 24";

/** Low, middle, high. Every set is coloured from these, in band order. */
const BAND_COLORS = ["#dc2626", "#f59e0b", "#16a34a"] as const;

type Shape = "circle" | "arrow" | "flag" | "triangle" | "star";

/**
 * A set name to the shape that carries its meaning. Matched by substring, since
 * the names are systematic — `3Arrows`, `4ArrowsGray`, `5Arrows` all arrows.
 */
const SHAPES: readonly [string, Shape][] = [
  ["Arrow", "arrow"],
  ["Flag", "flag"],
  ["Triangle", "triangle"],
  ["Star", "star"],
  ["Rating", "star"],
  ["Quarters", "circle"],
  ["Boxes", "circle"],
];

function shapeFor(set: string | undefined): Shape {
  if (!set) return "circle";
  return SHAPES.find(([name]) => set.includes(name))?.[1] ?? "circle";
}

/**
 * The colour for a band, stretched across however many bands the set has.
 *
 * A five-band set still reads low to high through the same three colours; using
 * five distinct ones would invent a scale the file does not describe.
 */
function colorFor(index: number, bands: number): string {
  if (bands <= 1) return BAND_COLORS[1] as string;
  const at = Math.round((index / (bands - 1)) * (BAND_COLORS.length - 1));
  return BAND_COLORS[Math.max(0, Math.min(BAND_COLORS.length - 1, at))] as string;
}

/** Bands in a set name: the leading digit of `3TrafficLights1`. */
function bandsIn(set: string | undefined): number {
  const n = Number.parseInt(set ?? "", 10);
  return Number.isFinite(n) && n > 0 ? n : BAND_COLORS.length;
}

export interface CondIconProps {
  set: string | undefined;
  /** 0-based band, counting up from the lowest. */
  index: number;
}

export function CondIcon({ set, index }: CondIconProps): ReactNode {
  const bands = bandsIn(set);
  const color = colorFor(index, bands);
  const shape = shapeFor(set);

  // An arrow points up in the top band, down in the bottom, and sideways between.
  const arrowAngle = index === 0 ? 90 : index >= bands - 1 ? -90 : 0;

  return (
    <svg
      xmlns="http://www.w3.org/2000/svg"
      width={ICON_SIZE}
      height={ICON_SIZE}
      viewBox={VIEW_BOX}
      aria-hidden="true"
      focusable="false"
      style={{ display: "block" }}
    >
      {shape === "circle" && <circle cx="12" cy="12" r="8" fill={color} />}
      {shape === "arrow" && (
        <g
          transform={`rotate(${arrowAngle} 12 12)`}
          stroke={color}
          strokeWidth="2.5"
          strokeLinecap="round"
          strokeLinejoin="round"
          fill="none"
        >
          <path d="M5 12h14" />
          <path d="M13 6l6 6-6 6" />
        </g>
      )}
      {shape === "flag" && (
        <path d="M6 21V4h12l-3 4 3 4H6" fill={color} stroke="none" />
      )}
      {shape === "triangle" && (
        <path
          d={index === 0 ? "M12 20L3 6h18z" : "M12 4l9 14H3z"}
          fill={color}
          stroke="none"
        />
      )}
      {shape === "star" && (
        <path
          d="M12 3l2.8 5.7 6.2.9-4.5 4.4 1.1 6.2L12 17.3 6.4 20.2l1.1-6.2L3 9.6l6.2-.9z"
          fill={index === 0 ? "none" : color}
          stroke={color}
          strokeWidth="1.5"
        />
      )}
    </svg>
  );
}
