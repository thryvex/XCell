import { Workbook, Sheet, Cell, MergeRange } from './model';

export interface CellPatch {
  key: string;
  before?: Cell;
  after?: Cell;
}

/** cell patch ที่ระบุชีต ใช้กับสูตรในชีตอื่นที่ถูกปรับตามการแทรก/ลบแถว-คอลัมน์ */
export interface SheetCellPatch extends CellPatch {
  sheetIndex: number;
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
  /** ชีตที่ได้รับผลกระทบ (Undo/Redo จะสลับไปชีตนี้) */
  sheetIndex?: number;
  patch?: SheetPatch;
  /** สูตรในชีตอื่นที่ถูกเขียนใหม่พร้อมกัน */
  extraCellPatches?: SheetCellPatch[];
  /** การเปลี่ยนโครงสร้างสมุดงาน (เพิ่ม/ลบ/เปลี่ยนชื่อ/ทำสำเนา/ย้ายชีต) เก็บทั้งชุดชีต */
  workbookBefore?: WorkbookSnapshot;
  workbookAfter?: WorkbookSnapshot;
}

function cloneSheets(sheets: Sheet[]): Sheet[] {
  return JSON.parse(JSON.stringify(sheets));
}

export function snapshotWorkbook(wb: Workbook): WorkbookSnapshot {
  return { sheets: cloneSheets(wb.sheets), activeSheet: wb.activeSheet };
}

function applySnapshot(sheet: Sheet, snap: FullSheetSnapshot) {
  sheet.cells = JSON.parse(JSON.stringify(snap.cells));
  sheet.merges = JSON.parse(JSON.stringify(snap.merges));
  if (snap.colWidths) sheet.colWidths = { ...snap.colWidths };
  if (snap.rowHeights) sheet.rowHeights = { ...snap.rowHeights };
  if (snap.frozenRows !== undefined) sheet.frozenRows = snap.frozenRows;
  if (snap.frozenCols !== undefined) sheet.frozenCols = snap.frozenCols;
}

function applyCellPatches(sheet: Sheet, patches: CellPatch[], direction: 'before' | 'after') {
  for (const p of patches) {
    const value = p[direction];
    if (value) {
      sheet.cells[p.key] = { ...value };
    } else {
      delete sheet.cells[p.key];
    }
  }
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
    this.apply(wb, action, 'before');
    this.redoStack.push(action);
    return action;
  }

  public redo(wb: Workbook): HistoryAction | null {
    const action = this.redoStack.pop();
    if (!action) return null;
    this.apply(wb, action, 'after');
    this.undoStack.push(action);
    return action;
  }

  private apply(wb: Workbook, action: HistoryAction, direction: 'before' | 'after') {
    const wbSnap = direction === 'before' ? action.workbookBefore : action.workbookAfter;
    if (wbSnap) {
      wb.sheets = cloneSheets(wbSnap.sheets);
      wb.activeSheet = wbSnap.activeSheet;
      return;
    }

    if (action.sheetIndex !== undefined && action.patch) {
      const sheet = wb.sheets[action.sheetIndex];
      if (sheet) {
        const full = direction === 'before' ? action.patch.fullSheetBefore : action.patch.fullSheetAfter;
        if (full) {
          applySnapshot(sheet, full);
        } else {
          if (action.patch.cellPatches) {
            applyCellPatches(sheet, action.patch.cellPatches, direction);
          }
          const merges = direction === 'before' ? action.patch.mergesBefore : action.patch.mergesAfter;
          if (merges !== undefined) {
            sheet.merges = [...merges];
          }
        }
      }
    }

    if (action.extraCellPatches) {
      for (const p of action.extraCellPatches) {
        const sheet = wb.sheets[p.sheetIndex];
        if (sheet) applyCellPatches(sheet, [p], direction);
      }
    }
  }

  public clear() {
    this.undoStack = [];
    this.redoStack = [];
  }
}
