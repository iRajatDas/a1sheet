"use client";

/**
 * Import and export buttons — a primitive, because file I/O is the heaviest
 * thing in the library and not every sheet wants it.
 *
 * This used to be three optional callbacks on `Toolbar`, which had two problems.
 * A `<Sheet.Toolbar />` composed by hand got none of them and silently rendered
 * no Import button at all, and passing behaviour down as props is the
 * configuration pattern the toolbar is supposed to avoid. As a primitive it is
 * a child you render:
 *
 *   <Sheet.Toolbar>
 *     <Sheet.FileMenu />
 *   </Sheet.Toolbar>
 *
 * Not rendering it keeps the XLSX writer and the ZIP reader out of the bundle,
 * which is the property the callbacks were protecting in the first place.
 */
import { type ReactNode, useCallback, useEffect, useRef } from "react";
import { isA1SheetError } from "../../errors.js";
import { downloadCsv, downloadXlsx } from "../../io/download.js";
import { readWorkbookFile } from "../../io/index.js";
import {
  DEFAULT_NUM_COLS,
  DEFAULT_NUM_ROWS,
  makeSheet,
} from "../../model/sheet.js";
import type { Sheet as SheetModel } from "../../model/types.js";
import { useSheetContext } from "../context.js";
import { ExportIcon, ImportIcon } from "./icons.js";

const PERCENT = 100;

function percent(ratio: number): string {
  return `${Math.round(ratio * PERCENT)}%`;
}

export function FileMenu(): ReactNode {
  const { api, prefix } = useSheetContext("Sheet.FileMenu");
  const { sheet } = api;
  const fileRef = useRef<HTMLInputElement>(null);

  // One in-flight import at a time. Picking a second file supersedes the first
  // rather than racing it, and unmounting cancels whatever is still parsing.
  const importRef = useRef<AbortController | null>(null);
  useEffect(() => () => importRef.current?.abort(), []);

  const setStatus = api.setStatus;
  const replaceWorkbook = api.replaceWorkbook;

  const handleImport = useCallback(
    async (file: File) => {
      importRef.current?.abort();
      const controller = new AbortController();
      importRef.current = controller;

      try {
        const { sheets, namedRanges, namedFormulas } = await readWorkbookFile(
          file,
          {
            signal: controller.signal,
            onProgress: ({ ratio, detail }) => {
              setStatus(`Importing ${file.name} — ${percent(ratio)} (${detail})`);
            },
          },
        );
        const rebuilt: SheetModel[] = sheets.map((s) => ({
          ...makeSheet(s.name),
          cells: s.cells,
          styles: s.styles,
          cachedValues: s.cachedValues,
          condFormats: s.condFormats,
          images: s.images,
          tables: s.tables,
          validations: s.validations,
          hiddenCols: new Set(s.hiddenCols),
          hiddenRows: new Set(s.hiddenRows),
          spillRanges: s.spillRanges,
          merges: s.merges,
          colWidths: s.colWidths,
          rowHeights: s.rowHeights,
          numRows: Math.max(DEFAULT_NUM_ROWS, s.rows),
          numCols: Math.max(DEFAULT_NUM_COLS, s.cols),
        }));
        replaceWorkbook({
          sheets: rebuilt,
          activeSheetIndex: 0,
          namedRanges,
          namedFormulas,
        });
        setStatus(`Imported ${file.name}`);
      } catch (err) {
        // A superseded or unmounted import is not a failure — say nothing, and
        // leave the status the newer import is already writing.
        if (isA1SheetError(err) && err.code === "ABORTED") return;
        setStatus(
          `Import failed: ${err instanceof Error ? err.message : String(err)}`,
        );
      } finally {
        if (importRef.current === controller) importRef.current = null;
      }
    },
    [setStatus, replaceWorkbook],
  );

  return (
    <>
      <button
        type="button"
        className={`${prefix}btn`}
        onClick={() => fileRef.current?.click()}
      >
        <ImportIcon />
        Import
      </button>
      <input
        ref={fileRef}
        type="file"
        aria-label="Import a spreadsheet"
        accept=".csv,.xlsx,.xlsm"
        style={{ display: "none" }}
        onChange={(e) => {
          const file = e.target.files?.[0];
          if (file) void handleImport(file);
          // Cleared so that re-picking the same file fires change again.
          e.target.value = "";
        }}
      />
      <button
        type="button"
        className={`${prefix}btn`}
        onClick={() => downloadCsv(sheet.cells, api.evaluator, `${sheet.name}.csv`)}
      >
        <ExportIcon />
        Export CSV
      </button>
      <button
        type="button"
        className={`${prefix}btn`}
        onClick={() =>
          downloadXlsx(
            api.workbook.sheets.map((s) => ({
              name: s.name,
              cells: s.cells,
              styles: s.styles,
              cachedValues: s.cachedValues,
              merges: s.merges,
              colWidths: s.colWidths,
              rowHeights: s.rowHeights,
              tables: s.tables,
              condFormats: s.condFormats,
              images: s.images,
              namedRanges: api.workbook.namedRanges,
            })),
          )
        }
      >
        <ExportIcon />
        Export XLSX
      </button>
    </>
  );
}
