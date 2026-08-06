"use client";

/**
 * `<Spreadsheet />` — a PRESET, not the extension point.
 *
 * It composes the primitives in the default arrangement and takes no layout props:
 * no `showToolbar`, no `showStatusBar`. If you want a different arrangement, compose
 * the primitives yourself — that is the supported path, and this component is the
 * shortest possible example of it. Everything below is a child you could move,
 * drop, or replace:
 *
 *   <Sheet.Root defaultWorkbook={wb}>
 *     <Sheet.Toolbar />
 *     <Sheet.Grid />
 *     <MyFooter />
 *   </Sheet.Root>
 */
import { forwardRef, type ReactNode } from "react";
import { AddRows } from "./components/AddRows.js";
import { ColumnMenu } from "./components/ColumnMenu.js";
import { ContextMenu } from "./components/ContextMenu.js";
import { FileMenu } from "./components/FileMenu.js";
import { FormulaBar } from "./components/FormulaBar.js";
import { Grid } from "./components/Grid.js";
import { SheetTabs } from "./components/SheetTabs.js";
import { StatusBar } from "./components/StatusBar.js";
import { Toolbar } from "./components/Toolbar.js";
import { Root, type SheetRootHandle, type SheetRootProps } from "./Root.js";

export type SpreadsheetProps = SheetRootProps;

export const Spreadsheet = forwardRef<SheetRootHandle, SpreadsheetProps>(
  function Spreadsheet(props, ref) {
    return (
      <Root {...props} ref={ref}>
        <Toolbar>
          <FileMenu />
        </Toolbar>
        <FormulaBar />
        <Grid>
          <AddRows />
        </Grid>
        <SheetTabs />
        <StatusBar />
        <ContextMenu />
        <ColumnMenu />
      </Root>
    );
  },
);
