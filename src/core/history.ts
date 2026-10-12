import { Workbook, Cell, MergeRange } from './model';

export interface CellPatch {
  key: string;
  before?: Cell;
  after?: Cell;
}

export interface SheetPatch {
  cellPatches?: CellPatch[];
  mergesBefore?: MergeRange[];
  mergesAfter?: MergeRange[];
}

export interface HistoryAction {
  description: string;
  sheetIndex: number;
  patch: SheetPatch;
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

    const sheet = wb.sheets[action.sheetIndex];
    if (sheet) {
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

    this.redoStack.push(action);
    return action;
  }

  public redo(wb: Workbook): HistoryAction | null {
    const action = this.redoStack.pop();
    if (!action) return null;

    const sheet = wb.sheets[action.sheetIndex];
    if (sheet) {
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

    this.undoStack.push(action);
    return action;
  }

  public clear() {
    this.undoStack = [];
    this.redoStack = [];
  }
}
