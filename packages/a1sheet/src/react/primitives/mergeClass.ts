/**
 * Concatenate class names, skipping falsy entries. Zero runtime dependencies.
 */
export function mergeClass(
  ...parts: (string | undefined | false | null)[]
): string | undefined {
  const merged = parts.filter(Boolean).join(" ");
  return merged || undefined;
}
