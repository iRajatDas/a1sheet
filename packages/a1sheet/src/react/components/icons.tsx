"use client";

/**
 * The icon set the built-in primitives draw with.
 *
 * Hand-drawn glyphs (`↶`, `⯇`, `▾`) and emoji (`🔒`) were doing this job before.
 * Both are the wrong tool: a glyph renders in whatever font the host page
 * happens to resolve, so it changes shape, weight, and baseline between
 * machines, and several of the ones used here have no coverage in the common
 * UI fonts at all — they fell back to a tofu box. Emoji are worse still, being
 * colour bitmaps on most platforms: they ignore `color`, so the lock did not
 * turn white when its button went active.
 *
 * The paths below are Tabler Icons 3.44.0 (MIT), copied in rather than
 * installed — `dependencies` stays empty, which is a hard rule (CLAUDE.md §1).
 * See THIRD-PARTY-NOTICES.md for the licence. Each icon names the Tabler icon
 * it came from, so it can be re-copied when the set is updated.
 *
 * They are drawn at `ICON_SIZE` and stroked in `currentColor`, so they inherit
 * the button's colour — including the active and disabled states — and stay
 * crisp at any zoom.
 */
import type { ReactNode } from "react";

/**
 * Drawn slightly below the 16px cap-height of the surrounding text so an icon
 * button and a text button come out the same height.
 */
const ICON_SIZE = 15;

/** For icons that sit inside the grid, where a toolbar-sized one would crowd. */
const HEADER_ICON_SIZE = 12;

/** The grid every path below is drawn on, matching Tabler's own. */
const VIEW_BOX = "0 0 24 24";

/** Tabler's stroke weight. Changing it would break the set's consistency. */
const STROKE_WIDTH = 2;

interface IconProps {
  /** Paths, on the 24-unit grid. */
  children: ReactNode;
  /** Defaults to `ICON_SIZE`. */
  size?: number;
}

/**
 * Always `aria-hidden`. An icon is never the accessible name — the button
 * around it carries a `title` and an `aria-label`, which is what a tooltip and
 * a screen reader respectively need.
 */
function Icon({ children, size = ICON_SIZE }: IconProps): ReactNode {
  return (
    <svg
      xmlns="http://www.w3.org/2000/svg"
      width={size}
      height={size}
      viewBox={VIEW_BOX}
      fill="none"
      stroke="currentColor"
      strokeWidth={STROKE_WIDTH}
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
      focusable="false"
      style={{ display: "block", flex: "none" }}
    >
      {children}
    </svg>
  );
}

/** Tabler `arrow-back-up`. */
export function UndoIcon(): ReactNode {
  return (
    <Icon>
      <path d="M9 14l-4 -4l4 -4" />
      <path d="M5 10h11a4 4 0 1 1 0 8h-1" />
    </Icon>
  );
}

/** Tabler `arrow-forward-up`. */
export function RedoIcon(): ReactNode {
  return (
    <Icon>
      <path d="M15 14l4 -4l-4 -4" />
      <path d="M19 10h-11a4 4 0 1 0 0 8h1" />
    </Icon>
  );
}

/** Tabler `bold`. */
export function BoldIcon(): ReactNode {
  return (
    <Icon>
      <path d="M7 5h6a3.5 3.5 0 0 1 0 7h-6l0 -7" />
      <path d="M13 12h1a3.5 3.5 0 0 1 0 7h-7v-7" />
    </Icon>
  );
}

/** Tabler `italic`. */
export function ItalicIcon(): ReactNode {
  return (
    <Icon>
      <path d="M11 5l6 0" />
      <path d="M7 19l6 0" />
      <path d="M14 5l-4 14" />
    </Icon>
  );
}

/** Tabler `underline`. */
export function UnderlineIcon(): ReactNode {
  return (
    <Icon>
      <path d="M7 5v5a5 5 0 0 0 10 0v-5" />
      <path d="M5 19h14" />
    </Icon>
  );
}

/** Tabler `align-left`. */
export function AlignLeftIcon(): ReactNode {
  return (
    <Icon>
      <path d="M4 6l16 0" />
      <path d="M4 12l10 0" />
      <path d="M4 18l14 0" />
    </Icon>
  );
}

/** Tabler `align-center`. */
export function AlignCenterIcon(): ReactNode {
  return (
    <Icon>
      <path d="M4 6l16 0" />
      <path d="M8 12l8 0" />
      <path d="M6 18l12 0" />
    </Icon>
  );
}

/** Tabler `align-right`. */
export function AlignRightIcon(): ReactNode {
  return (
    <Icon>
      <path d="M4 6l16 0" />
      <path d="M10 12l10 0" />
      <path d="M6 18l14 0" />
    </Icon>
  );
}

/** Tabler `lock`. */
export function LockIcon(): ReactNode {
  return (
    <Icon>
      <path d="M5 13a2 2 0 0 1 2 -2h10a2 2 0 0 1 2 2v6a2 2 0 0 1 -2 2h-10a2 2 0 0 1 -2 -2v-6" />
      <path d="M11 16a1 1 0 1 0 2 0a1 1 0 0 0 -2 0" />
      <path d="M8 11v-4a4 4 0 1 1 8 0v4" />
    </Icon>
  );
}

/** Tabler `lock-open`. */
export function UnlockIcon(): ReactNode {
  return (
    <Icon>
      <path d="M5 13a2 2 0 0 1 2 -2h10a2 2 0 0 1 2 2v6a2 2 0 0 1 -2 2h-10a2 2 0 0 1 -2 -2l0 -6" />
      <path d="M11 16a1 1 0 1 0 2 0a1 1 0 1 0 -2 0" />
      <path d="M8 11v-5a4 4 0 0 1 8 0" />
    </Icon>
  );
}

/** Tabler `row-insert-bottom`. */
export function InsertRowIcon(): ReactNode {
  return (
    <Icon>
      <path d="M20 6v4a1 1 0 0 1 -1 1h-14a1 1 0 0 1 -1 -1v-4a1 1 0 0 1 1 -1h14a1 1 0 0 1 1 1" />
      <path d="M12 15l0 4" />
      <path d="M14 17l-4 0" />
    </Icon>
  );
}

/** Tabler `row-remove`. */
export function DeleteRowIcon(): ReactNode {
  return (
    <Icon>
      <path d="M20 6v4a1 1 0 0 1 -1 1h-14a1 1 0 0 1 -1 -1v-4a1 1 0 0 1 1 -1h14a1 1 0 0 1 1 1" />
      <path d="M10 16l4 4" />
      <path d="M10 20l4 -4" />
    </Icon>
  );
}

/** Tabler `column-insert-right`. */
export function InsertColIcon(): ReactNode {
  return (
    <Icon>
      <path d="M6 4h4a1 1 0 0 1 1 1v14a1 1 0 0 1 -1 1h-4a1 1 0 0 1 -1 -1v-14a1 1 0 0 1 1 -1" />
      <path d="M15 12l4 0" />
      <path d="M17 10l0 4" />
    </Icon>
  );
}

/** Tabler `column-remove`. */
export function DeleteColIcon(): ReactNode {
  return (
    <Icon>
      <path d="M6 4h4a1 1 0 0 1 1 1v14a1 1 0 0 1 -1 1h-4a1 1 0 0 1 -1 -1v-14a1 1 0 0 1 1 -1" />
      <path d="M16 10l4 4" />
      <path d="M16 14l4 -4" />
    </Icon>
  );
}

/** Tabler `arrows-join-2`. */
export function MergeIcon(): ReactNode {
  return (
    <Icon>
      <path d="M3 7h1.948c1.913 0 3.705 .933 4.802 2.5a5.861 5.861 0 0 0 4.802 2.5h6.448" />
      <path d="M3 17h1.95a5.854 5.854 0 0 0 4.798 -2.5a5.854 5.854 0 0 1 4.798 -2.5h5.454" />
      <path d="M18 15l3 -3l-3 -3" />
    </Icon>
  );
}

/** Tabler `arrows-split-2`. */
export function UnmergeIcon(): ReactNode {
  return (
    <Icon>
      <path d="M21 17h-5.397a5 5 0 0 1 -4.096 -2.133l-.514 -.734a5 5 0 0 0 -4.096 -2.133h-3.897" />
      <path d="M21 7h-5.395a5 5 0 0 0 -4.098 2.135l-.51 .73a5 5 0 0 1 -4.097 2.135h-3.9" />
      <path d="M18 10l3 -3l-3 -3" />
      <path d="M18 20l3 -3l-3 -3" />
    </Icon>
  );
}

/** Tabler `freeze-row-column`: the pinned band is the hatched corner. */
export function FreezeIcon(): ReactNode {
  return (
    <Icon>
      <path d="M3 5a2 2 0 0 1 2 -2h14a2 2 0 0 1 2 2v14a2 2 0 0 1 -2 2h-14a2 2 0 0 1 -2 -2v-14" />
      <path d="M15 3l-12 12" />
      <path d="M9.5 3l-6 6" />
      <path d="M20 3.5l-5.5 5.5" />
      <path d="M9 15l-5 5" />
      <path d="M21 9h-12v12" />
    </Icon>
  );
}

/** Tabler `table-dashed`: the same grid with nothing pinned. */
export function UnfreezeIcon(): ReactNode {
  return (
    <Icon>
      <path d="M3 5a2 2 0 0 1 2 -2h14a2 2 0 0 1 2 2v14a2 2 0 0 1 -2 2h-14a2 2 0 0 1 -2 -2l0 -14" />
      <path d="M3 10h18" />
      <path d="M10 3v18" />
    </Icon>
  );
}

/** Tabler `upload`. */
export function ImportIcon(): ReactNode {
  return (
    <Icon>
      <path d="M4 17v2a2 2 0 0 0 2 2h12a2 2 0 0 0 2 -2v-2" />
      <path d="M7 9l5 -5l5 5" />
      <path d="M12 4l0 12" />
    </Icon>
  );
}

/** Tabler `download`. */
export function ExportIcon(): ReactNode {
  return (
    <Icon>
      <path d="M4 17v2a2 2 0 0 0 2 2h12a2 2 0 0 0 2 -2v-2" />
      <path d="M7 11l5 5l5 -5" />
      <path d="M12 4l0 12" />
    </Icon>
  );
}

/** Tabler `chevron-down`, at header size — it sits in a column header. */
export function ChevronDownIcon(): ReactNode {
  return (
    <Icon size={HEADER_ICON_SIZE}>
      <path d="M6 9l6 6l6 -6" />
    </Icon>
  );
}

/** Tabler `plus`. */
export function PlusIcon(): ReactNode {
  return (
    <Icon>
      <path d="M12 5l0 14" />
      <path d="M5 12l14 0" />
    </Icon>
  );
}

/** Tabler `trash`. */
export function TrashIcon(): ReactNode {
  return (
    <Icon>
      <path d="M4 7l16 0" />
      <path d="M10 11l0 6" />
      <path d="M14 11l0 6" />
      <path d="M5 7l1 12a2 2 0 0 0 2 2h8a2 2 0 0 0 2 -2l1 -12" />
      <path d="M9 7v-3a1 1 0 0 1 1 -1h4a1 1 0 0 1 1 1v3" />
    </Icon>
  );
}
