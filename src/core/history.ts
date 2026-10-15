import { Workbook, Cell, MergeRange } from './model';

export interface CellPatch {
  key: string;
  before?: Cell;
  after?: Cell;
}

export interface FullSheetSnapshot {
  cells: Record<string, Cell>;
  merges: MergeRange[];
  colWidths?: Record<number, number>;
  rowHeights?: Record<number, number>;
  frozenRows?: number;
  frozenCols?: number;
  rowCount?: number;
  colCount?: number;
}

export interface SheetPatch {
  cellPatches?: CellPatch[];
  mergesBefore?: MergeRange[];
  mergesAfter?: MergeRange[];
  fullSheetBefore?: FullSheetSnapshot;
  fullSheetAfter?: FullSheetSnapshot;
}

export interface WorkbookSnapshot {
  sheets: Sheet[];
  activeSheet: number;
}

export interface HistoryAction {
  description: string;
  sheetIndex?: number;
  patch?: SheetPatch;
  workbookBefore?: WorkbookSnapshot;
  workbookAfter?: WorkbookSnapshot;
  activeSheetBefore?: number;
  activeSheetAfter?: number;
}

export class HistoryManager {
  private undoStack: HistoryAction[] = [];
  private redoStack: HistoryAction[] = [];
  private readonly maxSteps: number;

  constructor(maxSteps = 100) {
    this.maxSteps = maxSteps;
  }

  public push(action: HistoryAction) {
    this.undoStack.push(action);
    if (this.undoStack.length > this.maxSteps) {
      this.undoStack.shift();
    }
    // Any new action clears the redo stack
    this.redoStack = [];
  }

  public canUndo(): boolean {
    return this.undoStack.length > 0;
  }

  public canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  public undo(wb: Workbook): HistoryAction | null {
    const action = this.undoStack.pop();
    if (!action) return null;

    if (action.workbookBefore) {
      wb.sheets = JSON.parse(JSON.stringify(action.workbookBefore.sheets));
      wb.activeSheet = action.workbookBefore.activeSheet;
      this.redoStack.push(action);
      return action;
    }

    if (action.sheetIndex !== undefined && action.patch) {
      const sheet = wb.sheets[action.sheetIndex];
      if (sheet) {
        if (action.patch.fullSheetBefore) {
          sheet.cells = JSON.parse(JSON.stringify(action.patch.fullSheetBefore.cells));
          sheet.merges = JSON.parse(JSON.stringify(action.patch.fullSheetBefore.merges));
          if (action.patch.fullSheetBefore.colWidths) sheet.colWidths = { ...action.patch.fullSheetBefore.colWidths };
          if (action.patch.fullSheetBefore.rowHeights) sheet.rowHeights = { ...action.patch.fullSheetBefore.rowHeights };
          if (action.patch.fullSheetBefore.frozenRows !== undefined) sheet.frozenRows = action.patch.fullSheetBefore.frozenRows;
          if (action.patch.fullSheetBefore.frozenCols !== undefined) sheet.frozenCols = action.patch.fullSheetBefore.frozenCols;
        } else {
          // Revert cell patches
          if (action.patch.cellPatches) {
            for (const p of action.patch.cellPatches) {
              if (p.before) {
                sheet.cells[p.key] = { ...p.before };
              } else {
                delete sheet.cells[p.key];
              }
            }
          }

          // Revert merges
          if (action.patch.mergesBefore !== undefined) {
            sheet.merges = [...action.patch.mergesBefore];
          }
        }
      }
    }

    this.redoStack.push(action);
    return action;
  }

  public redo(wb: Workbook): HistoryAction | null {
    const action = this.redoStack.pop();
    if (!action) return null;

    if (action.workbookAfter) {
      wb.sheets = JSON.parse(JSON.stringify(action.workbookAfter.sheets));
      wb.activeSheet = action.workbookAfter.activeSheet;
      this.undoStack.push(action);
      return action;
    }

    if (action.sheetIndex !== undefined && action.patch) {
      const sheet = wb.sheets[action.sheetIndex];
      if (sheet) {
        if (action.patch.fullSheetAfter) {
          sheet.cells = JSON.parse(JSON.stringify(action.patch.fullSheetAfter.cells));
          sheet.merges = JSON.parse(JSON.stringify(action.patch.fullSheetAfter.merges));
          if (action.patch.fullSheetAfter.colWidths) sheet.colWidths = { ...action.patch.fullSheetAfter.colWidths };
          if (action.patch.fullSheetAfter.rowHeights) sheet.rowHeights = { ...action.patch.fullSheetAfter.rowHeights };
          if (action.patch.fullSheetAfter.frozenRows !== undefined) sheet.frozenRows = action.patch.fullSheetAfter.frozenRows;
          if (action.patch.fullSheetAfter.frozenCols !== undefined) sheet.frozenCols = action.patch.fullSheetAfter.frozenCols;
        } else {
          // Re-apply cell patches
          if (action.patch.cellPatches) {
            for (const p of action.patch.cellPatches) {
              if (p.after) {
                sheet.cells[p.key] = { ...p.after };
              } else {
                delete sheet.cells[p.key];
              }
            }
          }

          // Re-apply merges
          if (action.patch.mergesAfter !== undefined) {
            sheet.merges = [...action.patch.mergesAfter];
          }
        }
      }
    }

    this.undoStack.push(action);
    return action;
  }

  public clear() {
    this.undoStack = [];
    this.redoStack = [];
  }
}
