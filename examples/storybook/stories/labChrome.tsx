/**
 * Shared chrome for Feature Lab demos (Storybook + Vite).
 * Deliberately product-shaped: grouped tools, primary actions, live status —
 * not a dump of equally-weighted buttons.
 */
import { toA1 } from "a1sheet";
import { useSheet } from "a1sheet/react";
import {
  type CSSProperties,
  type ReactNode,
  useEffect,
  useId,
  useState,
} from "react";

const font =
  '13px/1.45 "IBM Plex Sans", "Segoe UI", system-ui, sans-serif';
const mono =
  '12px/1.4 ui-monospace, "SF Mono", Menlo, Consolas, monospace';

export const labTokens = {
  ink: "#0f172a",
  muted: "#64748b",
  line: "#e2e8f0",
  surface: "#f8fafc",
  surfaceRaised: "#ffffff",
  accent: "#0d9488",
  accentSoft: "rgba(13, 148, 136, 0.12)",
  accentInk: "#0f766e",
  danger: "#b91c1c",
  darkBg: "#0b1220",
  darkSurface: "#111827",
  darkRaised: "#1e293b",
  darkLine: "#334155",
  darkMuted: "#94a3b8",
  darkInk: "#e2e8f0",
  darkAccent: "#2dd4bf",
  darkAccentSoft: "rgba(45, 212, 191, 0.14)",
} as const;

export type LabTone = "light" | "dark";

function t(tone: LabTone) {
  return tone === "dark"
    ? {
        ink: labTokens.darkInk,
        muted: labTokens.darkMuted,
        line: labTokens.darkLine,
        surface: labTokens.darkSurface,
        raised: labTokens.darkRaised,
        accent: labTokens.darkAccent,
        accentSoft: labTokens.darkAccentSoft,
        accentInk: labTokens.darkAccent,
      }
    : {
        ink: labTokens.ink,
        muted: labTokens.muted,
        line: labTokens.line,
        surface: labTokens.surface,
        raised: labTokens.surfaceRaised,
        accent: labTokens.accent,
        accentSoft: labTokens.accentSoft,
        accentInk: labTokens.accentInk,
      };
}

/** Page intro above the sheet — one job, one sentence. */
export function LabIntro({
  title,
  body,
  tone = "light",
}: {
  title: string;
  body: string;
  tone?: LabTone;
}): ReactNode {
  const c = t(tone);
  return (
    <header
      style={{
        font,
        padding: "14px 16px 12px",
        borderBottom: `1px solid ${c.line}`,
        background: c.surface,
        color: c.ink,
      }}
    >
      <div
        style={{
          font: '600 15px/1.3 "IBM Plex Sans", system-ui, sans-serif',
          letterSpacing: "-0.01em",
          marginBottom: 4,
        }}
      >
        {title}
      </div>
      <p style={{ margin: 0, color: c.muted, maxWidth: 52 * 8 }}>{body}</p>
    </header>
  );
}

export function LabToolbar({
  children,
  tone = "light",
}: {
  children: ReactNode;
  tone?: LabTone;
}): ReactNode {
  const c = t(tone);
  return (
    <div
      style={{
        font,
        display: "grid",
        gap: 12,
        padding: "12px 16px",
        borderBottom: `1px solid ${c.line}`,
        background: c.raised,
        color: c.ink,
      }}
    >
      {children}
    </div>
  );
}

export function LabSection({
  label,
  hint,
  children,
  tone = "light",
}: {
  label: string;
  hint?: string;
  children: ReactNode;
  tone?: LabTone;
}): ReactNode {
  const c = t(tone);
  return (
    <section style={{ display: "grid", gap: 8 }}>
      <div
        style={{
          display: "flex",
          alignItems: "baseline",
          justifyContent: "space-between",
          gap: 12,
        }}
      >
        <h3
          style={{
            margin: 0,
            font: "600 10px/1 system-ui, sans-serif",
            letterSpacing: "0.08em",
            textTransform: "uppercase",
            color: c.muted,
          }}
        >
          {label}
        </h3>
        {hint ? (
          <span style={{ fontSize: 12, color: c.muted }}>{hint}</span>
        ) : null}
      </div>
      {children}
    </section>
  );
}

export function LabRow({ children }: { children: ReactNode }): ReactNode {
  return (
    <div
      style={{
        display: "flex",
        flexWrap: "wrap",
        gap: 8,
        alignItems: "center",
      }}
    >
      {children}
    </div>
  );
}

type LabButtonKind = "primary" | "secondary" | "ghost";

export function LabButton({
  children,
  onClick,
  kind = "secondary",
  tone = "light",
  disabled,
  title,
}: {
  children: ReactNode;
  onClick: () => void;
  kind?: LabButtonKind;
  tone?: LabTone;
  disabled?: boolean;
  title?: string;
}): ReactNode {
  const c = t(tone);
  const base: CSSProperties = {
    font,
    fontWeight: 550,
    padding: "7px 12px",
    borderRadius: 8,
    border: "1px solid transparent",
    cursor: disabled ? "not-allowed" : "pointer",
    opacity: disabled ? 0.45 : 1,
    transition: "background 120ms ease, border-color 120ms ease, color 120ms ease",
  };
  const styles: Record<LabButtonKind, CSSProperties> = {
    primary: {
      ...base,
      background: c.accent,
      color: tone === "dark" ? "#042f2e" : "#fff",
      borderColor: c.accent,
    },
    secondary: {
      ...base,
      background: c.surface,
      color: c.ink,
      borderColor: c.line,
    },
    ghost: {
      ...base,
      background: "transparent",
      color: c.muted,
      borderColor: "transparent",
    },
  };
  return (
    <button
      type="button"
      title={title}
      disabled={disabled}
      onClick={onClick}
      style={styles[kind]}
    >
      {children}
    </button>
  );
}

export function LabField({
  label,
  value,
  onChange,
  placeholder,
  tone = "light",
  width = 160,
}: {
  label: string;
  value: string;
  onChange: (value: string) => void;
  placeholder?: string;
  tone?: LabTone;
  width?: number;
}): ReactNode {
  const c = t(tone);
  const id = useId();
  return (
    <label
      htmlFor={id}
      style={{ display: "grid", gap: 4, minWidth: width }}
    >
      <span
        style={{
          font: "600 10px/1 system-ui, sans-serif",
          letterSpacing: "0.06em",
          textTransform: "uppercase",
          color: c.muted,
        }}
      >
        {label}
      </span>
      <input
        id={id}
        value={value}
        placeholder={placeholder}
        onChange={(e) => onChange(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === "Enter") (e.target as HTMLInputElement).blur();
        }}
        style={{
          font,
          padding: "7px 10px",
          borderRadius: 8,
          border: `1px solid ${c.line}`,
          background: c.raised,
          color: c.ink,
          outline: "none",
          width: "100%",
          boxSizing: "border-box",
        }}
      />
    </label>
  );
}

/** Clickable colour chip for filter demos. */
export function LabSwatch({
  color,
  label,
  selected,
  onClick,
  tone = "light",
}: {
  color: string;
  label: string;
  selected?: boolean;
  onClick: () => void;
  tone?: LabTone;
}): ReactNode {
  const c = t(tone);
  return (
    <button
      type="button"
      onClick={onClick}
      aria-pressed={selected === true}
      title={label}
      style={{
        display: "inline-flex",
        alignItems: "center",
        gap: 8,
        padding: "6px 10px 6px 6px",
        borderRadius: 999,
        border: `1px solid ${selected ? c.accent : c.line}`,
        background: selected ? c.accentSoft : c.raised,
        color: c.ink,
        cursor: "pointer",
        font,
      }}
    >
      <span
        aria-hidden
        style={{
          width: 18,
          height: 18,
          borderRadius: 6,
          background: color,
          boxShadow: `inset 0 0 0 1px ${c.line}`,
        }}
      />
      {label}
    </button>
  );
}

export function LabStatusBar({ tone = "light" }: { tone?: LabTone }): ReactNode {
  const api = useSheet();
  const c = t(tone);
  const [flash, setFlash] = useState(false);

  useEffect(() => {
    if (!api.status) return;
    setFlash(true);
    const id = window.setTimeout(() => setFlash(false), 900);
    return () => window.clearTimeout(id);
  }, [api.status]);

  return (
    <div
      role="status"
      aria-live="polite"
      style={{
        display: "flex",
        alignItems: "center",
        gap: 10,
        minHeight: 28,
        padding: "6px 10px",
        borderRadius: 8,
        background: flash ? c.accentSoft : c.surface,
        border: `1px solid ${flash ? c.accent : c.line}`,
        color: api.status ? c.ink : c.muted,
        font,
        transition: "background 200ms ease, border-color 200ms ease",
      }}
    >
      <span
        style={{
          font: "600 10px/1 system-ui, sans-serif",
          letterSpacing: "0.08em",
          textTransform: "uppercase",
          color: c.muted,
          flex: "none",
        }}
      >
        Status
      </span>
      <span style={{ fontFamily: mono.split(",")[0], overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
        {api.status || "Ready — try an action above"}
      </span>
    </div>
  );
}

export function LabSteps({
  steps,
  tone = "light",
}: {
  steps: readonly string[];
  tone?: LabTone;
}): ReactNode {
  const c = t(tone);
  return (
    <ol
      style={{
        margin: 0,
        padding: 0,
        listStyle: "none",
        display: "grid",
        gap: 6,
      }}
    >
      {steps.map((step, i) => (
        <li
          key={step}
          style={{
            display: "grid",
            gridTemplateColumns: "22px 1fr",
            gap: 8,
            alignItems: "start",
            color: c.ink,
          }}
        >
          <span
            style={{
              width: 22,
              height: 22,
              borderRadius: 6,
              display: "grid",
              placeItems: "center",
              background: c.accentSoft,
              color: c.accentInk,
              font: "600 11px/1 system-ui, sans-serif",
            }}
          >
            {i + 1}
          </span>
          <span style={{ paddingTop: 2 }}>{step}</span>
        </li>
      ))}
    </ol>
  );
}

export function LabMeta({
  items,
  tone = "light",
}: {
  items: readonly { label: string; value: string }[];
  tone?: LabTone;
}): ReactNode {
  const c = t(tone);
  return (
    <dl
      style={{
        margin: 0,
        display: "flex",
        flexWrap: "wrap",
        gap: "6px 16px",
        color: c.muted,
        fontSize: 12,
      }}
    >
      {items.map((item) => (
        <div key={item.label} style={{ display: "flex", gap: 6 }}>
          <dt style={{ margin: 0 }}>{item.label}</dt>
          <dd style={{ margin: 0, color: c.ink, fontFamily: mono }}>{item.value}</dd>
        </div>
      ))}
    </dl>
  );
}

/** Segmented control for lab sheet kinds / modes. */
export function LabSegmented<T extends string>({
  value,
  onChange,
  options,
  tone = "light",
}: {
  value: T;
  onChange: (value: T) => void;
  options: readonly { id: T; label: string }[];
  tone?: LabTone;
}): ReactNode {
  const c = t(tone);
  return (
    <div
      role="tablist"
      style={{
        display: "inline-flex",
        padding: 3,
        gap: 2,
        borderRadius: 10,
        background: c.surface,
        border: `1px solid ${c.line}`,
      }}
    >
      {options.map((opt) => {
        const active = opt.id === value;
        return (
          <button
            key={opt.id}
            type="button"
            role="tab"
            aria-selected={active}
            onClick={() => onChange(opt.id)}
            style={{
              font,
              fontWeight: active ? 600 : 500,
              padding: "6px 12px",
              borderRadius: 8,
              border: "none",
              cursor: "pointer",
              background: active ? c.raised : "transparent",
              color: active ? c.ink : c.muted,
              boxShadow: active ? `0 0 0 1px ${c.line}` : "none",
            }}
          >
            {opt.label}
          </button>
        );
      })}
    </div>
  );
}

export function scrollHitIntoView(api: ReturnType<typeof useSheet>, row: number): void {
  const top = api.rowWindow.rowTop(row);
  const scroller = document.querySelector(".a1s-scroller");
  if (top !== null && scroller) scroller.scrollTop = Math.max(0, top - 40);
}

export function formatActive(api: ReturnType<typeof useSheet>): string {
  return toA1(api.active.row, api.active.col);
}
