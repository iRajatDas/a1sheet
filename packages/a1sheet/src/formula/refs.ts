/**
 * Relative-reference rewriting. Ported from ref/formulaEngine.js:364-381.
 *
 * Operates at the TOKEN level, not the AST level: re-tokenize, shift every ref
 * token lacking the relevant `$` marker, reassemble. A range is just two
 * independent ref tokens either side of `:`, so it needs no special handling.
 *
 * Used by both the fill handle and internal paste.
 */
import { colToLetters, parseCellRef } from "../model/address.js";
import { tokenize } from "./tokenize.js";

/**
 * Shifts every non-`$` reference by (dRow, dCol) — what Excel does when a formula
 * is dragged or copy-pasted elsewhere. A ref that would move off the sheet
 * becomes "#REF!".
 *
 * Takes the formula body WITHOUT the leading "=".
 */
export function shiftFormulaRefs(
  formulaText: string,
  dRow: number,
  dCol: number,
): string {
  return tokenize(formulaText)
    .map((t) => {
      if (t.type === "ref") {
        const { row, col } = parseCellRef(t.value);
        const newCol = t.colAbs ? col : col + dCol;
        const newRow = t.rowAbs ? row : row + dRow;
        if (newCol < 0 || newRow < 0) return "#REF!";
        const colPart = (t.colAbs ? "$" : "") + colToLetters(newCol);
        const rowPart = (t.rowAbs ? "$" : "") + (newRow + 1);
        return colPart + rowPart;
      }
      if (t.type === "num") return String(t.value);
      if (t.type === "str") return `"${t.value}"`;
      if (t.type === "name") return t.value;
      if (t.type === "cmp") return t.value;
      return t.type;
    })
    .join("");
}
