"use client";

/**
 * `<Spreadsheet />` — a PRESET, not the extension point.
 *
 * It composes the primitives in the default arrangement and takes no layout props:
 * no `showToolbar`, no `showStatusBar`. If you want a different arrangement, compose
 * the primitives yourself — that is the supported path, and this component is a
 * 30-line example of it:
 *
 *   <Sheet.Root defaultWorkbook={wb}>
 *     <Sheet.Toolbar />
 *     <Sheet.Grid />
 *     <MyFooter />
 *   </Sheet.Root>
 *
 * File I/O is wired here because a preset should work out of the box. A composed
 * tree that omits `onExportXlsx` never pulls the XLSX writer into its bundle.
 */
import { forwardRef, type ReactNode, useCallback, useEffect, useRef } from "react";
import { isA1SheetError } from "../errors.js";
import { downloadCsv, downloadXlsx } from "../io/download.js";
import { readWorkbookFile } from "../io/index.js";
import { makeSheet } from "../model/sheet.js";
import type { Sheet as SheetModel } from "../model/types.js";
import { ColumnMenu } from "./components/ColumnMenu.js";
import { ContextMenu } from "./components/ContextMenu.js";
import { FormulaBar } from "./components/FormulaBar.js";
import { Grid } from "./components/Grid.js";
import { SheetTabs } from "./components/SheetTabs.js";
import { StatusBar } from "./components/StatusBar.js";
import { Toolbar } from "./components/Toolbar.js";
import { useSheetContext } from "./context.js";
import { Root, type SheetRootHandle, type SheetRootProps } from "./Root.js";

export type SpreadsheetProps = SheetRootProps;

const PERCENT = 100;

function percent(ratio: number): string {
  return `${Math.round(ratio * PERCENT)}%`;
}

/** Inside Root, so it can reach the api for the file handlers. */
function PresetBody(): ReactNode {
  const { api } = useSheetContext("Spreadsheet");
  const { sheet } = api;

  // One in-flight import at a time. Picking a second file supersedes the first
  // rather than racing it, and unmounting cancels whatever is still parsing.
  const importRef = useRef<AbortController | null>(null);
  useEffect(() => () => importRef.current?.abort(), []);

  const handleImport = useCallback(
    async (file: File) => {
      importRef.current?.abort();
      const controller = new AbortController();
      importRef.current = controller;

      try {
        const { sheets } = await readWorkbookFile(file, {
          signal: controller.signal,
          onProgress: ({ ratio, detail }) => {
            api.setStatus(`Importing ${file.name} — ${percent(ratio)} (${detail})`);
          },
        });
        const rebuilt: SheetModel[] = sheets.map((s) => ({
          ...makeSheet(s.name),
          cells: s.cells,
          styles: s.styles,
          merges: s.merges,
          colWidths: s.colWidths,
          rowHeights: s.rowHeights,
          numRows: Math.max(200, s.rows),
          numCols: Math.max(26, s.cols),
        }));
        api.replaceWorkbook({
          sheets: rebuilt,
          activeSheetIndex: 0,
          namedRanges: {},
        });
        api.setStatus(`Imported ${file.name}`);
      } catch (err) {
        // A superseded or unmounted import is not a failure — say nothing, and
        // leave the status the newer import is already writing.
        if (isA1SheetError(err) && err.code === "ABORTED") return;
        api.setStatus(
          `Import failed: ${err instanceof Error ? err.message : String(err)}`,
        );
      } finally {
        if (importRef.current === controller) importRef.current = null;
      }
    },
    [api],
  );

  return (
    <>
      <Toolbar
        onImport={handleImport}
        onExportCsv={() =>
          downloadCsv(sheet.cells, api.evaluator, `${sheet.name}.csv`)
        }
        onExportXlsx={() =>
          downloadXlsx(
            api.workbook.sheets.map((s) => ({
              name: s.name,
              cells: s.cells,
              styles: s.styles,
              merges: s.merges,
              colWidths: s.colWidths,
              rowHeights: s.rowHeights,
              namedRanges: api.workbook.namedRanges,
            })),
          )
        }
      />
      <FormulaBar />
      <Grid />
      <SheetTabs />
      <StatusBar />
      <ContextMenu />
      <ColumnMenu />
    </>
  );
}

export const Spreadsheet = forwardRef<SheetRootHandle, SpreadsheetProps>(
  function Spreadsheet(props, ref) {
    return (
      <Root {...props} ref={ref}>
        <PresetBody />
      </Root>
    );
  },
);
