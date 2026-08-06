/**
 * Chrome for the stories — the little panels that show state next to a sheet.
 * Not part of the library, and deliberately plain so it never competes with
 * what a story is demonstrating.
 */
import type { ReactNode } from "react";

export function Panel({
  title,
  children,
}: {
  title: string;
  children: ReactNode;
}): ReactNode {
  return (
    <section
      style={{
        font: "13px/1.5 system-ui, sans-serif",
        background: "#f8fafc",
        border: "1px solid #e2e8f0",
        borderRadius: 8,
        padding: "10px 12px",
        marginBottom: 12,
        display: "grid",
        gap: 4,
      }}
    >
      <h3
        style={{
          margin: 0,
          font: "600 11px/1.4 system-ui, sans-serif",
          letterSpacing: "0.06em",
          textTransform: "uppercase",
          color: "#64748b",
        }}
      >
        {title}
      </h3>
      {children}
    </section>
  );
}

export function Row({
  label,
  children,
}: {
  label: string;
  children: ReactNode;
}): ReactNode {
  return (
    <div style={{ display: "flex", gap: 8 }}>
      <span style={{ color: "#64748b", minWidth: 168 }}>{label}</span>
      <span style={{ color: "#0f172a" }}>{children}</span>
    </div>
  );
}

/** A short "the problem / the solution" header, for the recipe stories. */
export function Problem({
  problem,
  solution,
}: {
  problem: string;
  solution: string;
}): ReactNode {
  return (
    <section
      style={{
        font: "13px/1.6 system-ui, sans-serif",
        borderLeft: "3px solid #0d9488",
        padding: "2px 0 2px 12px",
        margin: "0 0 12px",
        display: "grid",
        gap: 2,
      }}
    >
      <div>
        <strong>Problem — </strong>
        {problem}
      </div>
      <div style={{ color: "#334155" }}>
        <strong>Solution — </strong>
        {solution}
      </div>
    </section>
  );
}
