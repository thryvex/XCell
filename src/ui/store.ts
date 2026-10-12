import { create } from 'zustand';
import {
  Workbook,
  Sheet,
  NumFmt,
  MergeRange,
  createEmptyWorkbook,
  CellValue,
} from '../core/model';
import { Evaluator } from '../core/evaluator';
import { HistoryManager, HistoryAction } from '../core/history';
import { convertShorthand } from '../core/shorthand';
import { cellToA1, parseA1, isCellInRange } from '../core/address';
import { exportToXlsxBuffer, importFromXlsxBuffer } from '../core/xlsx';
import { formatDisplayValue } from '../core/numfmt';

export interface SelectionState {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

export interface SheetState {
  workbook: Workbook;
  activeSheetIndex: number;
  activeCell: { r: number; c: number };
  selection: SelectionState;
  isEditing: boolean;
  editingValue: string;
  filePath: string | null;
  fileName: string;
  isDirty: boolean;
  evaluator: Evaluator;
  history: HistoryManager;
  rowCount: number;
  colCount: number;

  // Actions
  setActiveCell: (r: number, c: number, extendSelection?: boolean) => void;
  setSelection: (r1: number, c1: number, r2: number, c2: number) => void;
  startEditing: (initialValue?: string) => void;
  setEditingValue: (val: string) => void;
  commitEditing: (dr?: number, dc?: number) => void;
  cancelEditing: () => void;
  clearSelectionContent: () => void;
  setNumFmt: (fmt: NumFmt) => void;
  mergeCells: () => boolean;
  unmergeCells: () => void;
  applySigma: (fnName: string) => void;
  undo: () => void;
  redo: () => void;
  newWorkbook: () => void;
  loadWorkbook: (data: Uint8Array, filePath?: string) => Promise<void>;
  saveWorkbook: () => Promise<Uint8Array>;
  markSaved: (filePath: string) => void;
  getDisplayValue: (r: number, c: number) => string;
  getRawValue: (r: number, c: number) => string;
  getComputedValue: (r: number, c: number) => CellValue;
  updateTitle: () => void;
}

const initialWb = createEmptyWorkbook();
const initialEvaluator = new Evaluator(initialWb.sheets[0]);
const initialHistory = new HistoryManager(100);

export const useSheetStore = create<SheetState>((set, get) => ({
  workbook: initialWb,
  activeSheetIndex: 0,
  activeCell: { r: 0, c: 0 },
  selection: { r1: 0, c1: 0, r2: 0, c2: 0 },
  isEditing: false,
  editingValue: '',
  filePath: null,
  fileName: 'ตารางบันทึก',
  isDirty: false,
  evaluator: initialEvaluator,
  history: initialHistory,
  rowCount: 100,
  colCount: 26,

  updateTitle: () => {
    const { fileName, isDirty } = get();
    const title = `${fileName}${isDirty ? ' ●' : ''} - SheetEasy`;
    if (typeof window !== 'undefined' && (window as any).electronAPI) {
      (window as any).electronAPI.setTitle(title);
    }
  },

  setActiveCell: (r, c, extend = false) => {
    const { rowCount, colCount, activeCell, selection, isEditing, commitEditing } = get();
    if (isEditing) {
      commitEditing(0, 0);
    }

    const clampedR = Math.max(0, Math.min(rowCount - 1, r));
    const clampedC = Math.max(0, Math.min(colCount - 1, c));

    if (extend) {
      const r1 = Math.min(selection.r1, clampedR);
      const c1 = Math.min(selection.c1, clampedC);
      const r2 = Math.max(selection.r2, clampedR);
      const c2 = Math.max(selection.c2, clampedC);
      set({
        activeCell: { r: clampedR, c: clampedC },
        selection: { r1, c1, r2, c2 },
      });
    } else {
      set({
        activeCell: { r: clampedR, c: clampedC },
        selection: { r1: clampedR, c1: clampedC, r2: clampedR, c2: clampedC },
      });
    }
  },

  setSelection: (r1, c1, r2, c2) => {
    set({
      selection: {
        r1: Math.min(r1, r2),
        c1: Math.min(c1, c2),
        r2: Math.max(r1, r2),
        c2: Math.max(c1, c2),
      },
    });
  },

  startEditing: (initialValue) => {
    const { activeCell, workbook, activeSheetIndex } = get();
    const sheet = workbook.sheets[activeSheetIndex];
    const key = cellToA1(activeCell.r, activeCell.c);
    const existing = sheet.cells[key]?.raw || '';
    const val = initialValue !== undefined ? initialValue : existing;

    set({
      isEditing: true,
      editingValue: val,
    });
  },

  setEditingValue: (val) => {
    set({ editingValue: val });
  },

  commitEditing: (dr = 0, dc = 0) => {
    const {
      isEditing,
      editingValue,
      activeCell,
      workbook,
      activeSheetIndex,
      evaluator,
      history,
      setActiveCell,
      updateTitle,
    } = get();

    if (!isEditing) return;

    const sheet = workbook.sheets[activeSheetIndex];
    const key = cellToA1(activeCell.r, activeCell.c);
    const prevCell = sheet.cells[key] ? { ...sheet.cells[key] } : undefined;

    // Apply shorthand conversion if user typed a shorthand
    let finalValue = editingValue.trim();
    const shorthandConverted = convertShorthand(finalValue);
    if (shorthandConverted) {
      finalValue = shorthandConverted;
    }

    const currentCell = sheet.cells[key];
    const isSame = (currentCell?.raw || '') === finalValue;

    if (!isSame) {
      if (finalValue === '') {
        delete sheet.cells[key];
      } else {
        sheet.cells[key] = {
          raw: finalValue,
          fmt: currentCell?.fmt,
          style: currentCell?.style,
        };
      }

      evaluator.invalidate(key);

      const action: HistoryAction = {
        description: `Edit ${key}`,
        sheetIndex: activeSheetIndex,
        patch: {
          cellPatches: [
            {
              key,
              before: prevCell,
              after: sheet.cells[key] ? { ...sheet.cells[key] } : undefined,
            },
          ],
        },
      };
      history.push(action);

      set({
        workbook: { ...workbook },
        isEditing: false,
        editingValue: '',
        isDirty: true,
      });
      updateTitle();
    } else {
      set({
        isEditing: false,
        editingValue: '',
      });
    }

    if (dr !== 0 || dc !== 0) {
      setActiveCell(activeCell.r + dr, activeCell.c + dc, false);
    }
  },

  cancelEditing: () => {
    set({
      isEditing: false,
      editingValue: '',
    });
  },

  clearSelectionContent: () => {
    const { selection, workbook, activeSheetIndex, evaluator, history, updateTitle } = get();
    const sheet = workbook.sheets[activeSheetIndex];
    const patches: any[] = [];

    for (let r = selection.r1; r <= selection.r2; r++) {
      for (let c = selection.c1; c <= selection.c2; c++) {
        const key = cellToA1(r, c);
        if (sheet.cells[key]) {
          patches.push({
            key,
            before: { ...sheet.cells[key] },
            after: undefined,
          });
          delete sheet.cells[key];
          evaluator.invalidate(key);
        }
      }
    }

    if (patches.length > 0) {
      history.push({
        description: 'Clear cells',
        sheetIndex: activeSheetIndex,
        patch: { cellPatches: patches },
      });
      set({ workbook: { ...workbook }, isDirty: true });
      updateTitle();
    }
  },

  setNumFmt: (fmt: NumFmt) => {
    const { selection, workbook, activeSheetIndex, history, updateTitle } = get();
    const sheet = workbook.sheets[activeSheetIndex];
    const patches: any[] = [];

    for (let r = selection.r1; r <= selection.r2; r++) {
      for (let c = selection.c1; c <= selection.c2; c++) {
        const key = cellToA1(r, c);
        const cell = sheet.cells[key] || { raw: '' };
        patches.push({
          key,
          before: sheet.cells[key] ? { ...sheet.cells[key] } : undefined,
          after: { ...cell, fmt },
        });
        sheet.cells[key] = { ...cell, fmt };
      }
    }

    history.push({
      description: `Format as ${fmt}`,
      sheetIndex: activeSheetIndex,
      patch: { cellPatches: patches },
    });
    set({ workbook: { ...workbook }, isDirty: true });
    updateTitle();
  },

  mergeCells: () => {
    const { selection, workbook, activeSheetIndex, history, updateTitle } = get();
    if (selection.r1 === selection.r2 && selection.c1 === selection.c2) {
      return false;
    }

    const sheet = workbook.sheets[activeSheetIndex];
    const mergesBefore = [...sheet.merges];

    // Remove any overlapping merges
    sheet.merges = sheet.merges.filter(
      (m) =>
        m.r2 < selection.r1 ||
        m.r1 > selection.r2 ||
        m.c2 < selection.c1 ||
        m.c1 > selection.c2
    );

    // Keep top-left value, clear others
    const topLeftKey = cellToA1(selection.r1, selection.c1);
    const topLeftCell = sheet.cells[topLeftKey];

    const patches: any[] = [];
    for (let r = selection.r1; r <= selection.r2; r++) {
      for (let c = selection.c1; c <= selection.c2; c++) {
        if (r === selection.r1 && c === selection.c1) continue;
        const key = cellToA1(r, c);
        if (sheet.cells[key]) {
          patches.push({
            key,
            before: { ...sheet.cells[key] },
            after: undefined,
          });
          delete sheet.cells[key];
        }
      }
    }

    const newMerge: MergeRange = {
      r1: selection.r1,
      c1: selection.c1,
      r2: selection.r2,
      c2: selection.c2,
    };
    sheet.merges.push(newMerge);

    history.push({
      description: 'Merge cells',
      sheetIndex: activeSheetIndex,
      patch: {
        cellPatches: patches,
        mergesBefore,
        mergesAfter: [...sheet.merges],
      },
    });

    set({ workbook: { ...workbook }, isDirty: true });
    updateTitle();
    return true;
  },

  unmergeCells: () => {
    const { activeCell, workbook, activeSheetIndex, history, updateTitle } = get();
    const sheet = workbook.sheets[activeSheetIndex];
    const mergesBefore = [...sheet.merges];

    sheet.merges = sheet.merges.filter((m) => !isCellInRange(activeCell.r, activeCell.c, m));

    if (sheet.merges.length !== mergesBefore.length) {
      history.push({
        description: 'Unmerge cells',
        sheetIndex: activeSheetIndex,
        patch: {
          mergesBefore,
          mergesAfter: [...sheet.merges],
        },
      });
      set({ workbook: { ...workbook }, isDirty: true });
      updateTitle();
    }
  },

  applySigma: (fnName: string) => {
    const {
      selection,
      workbook,
      activeSheetIndex,
      startEditing,
      setEditingValue,
      evaluator,
      history,
      updateTitle,
    } = get();
    const isMultiCell = selection.r1 !== selection.r2 || selection.c1 !== selection.c2;

    if (!isMultiCell) {
      // Case 2: single cell -> insert =FUNCTION( into formula bar
      startEditing(`=${fnName}(`);
      return;
    }

    // Case 1: multi-cell -> insert formula below the range
    const sheet = workbook.sheets[activeSheetIndex];
    const targetRow = selection.r2 + 1;
    const patches: any[] = [];

    // For each column in selection, insert formula below
    for (let c = selection.c1; c <= selection.c2; c++) {
      const formulaKey = cellToA1(targetRow, c);
      const rangeStr = `${cellToA1(selection.r1, c)}:${cellToA1(selection.r2, c)}`;
      const formula = `=${fnName}(${rangeStr})`;

      patches.push({
        key: formulaKey,
        before: sheet.cells[formulaKey] ? { ...sheet.cells[formulaKey] } : undefined,
        after: { raw: formula },
      });

      sheet.cells[formulaKey] = { raw: formula };
      evaluator.invalidate(formulaKey);
    }

    history.push({
      description: `Insert ${fnName}`,
      sheetIndex: activeSheetIndex,
      patch: { cellPatches: patches },
    });

    set({
      workbook: { ...workbook },
      activeCell: { r: targetRow, c: selection.c1 },
      selection: { r1: targetRow, c1: selection.c1, r2: targetRow, c2: selection.c2 },
      isDirty: true,
    });
    updateTitle();
  },

  undo: () => {
    const { history, workbook, evaluator, updateTitle } = get();
    if (!history.canUndo()) return;
    history.undo(workbook);
    evaluator.setSheet(workbook.sheets[get().activeSheetIndex]);
    set({ workbook: { ...workbook }, isDirty: true });
    updateTitle();
  },

  redo: () => {
    const { history, workbook, evaluator, updateTitle } = get();
    if (!history.canRedo()) return;
    history.redo(workbook);
    evaluator.setSheet(workbook.sheets[get().activeSheetIndex]);
    set({ workbook: { ...workbook }, isDirty: true });
    updateTitle();
  },

  newWorkbook: () => {
    const newWb = createEmptyWorkbook();
    const newEval = new Evaluator(newWb.sheets[0]);
    const newHist = new HistoryManager(100);

    set({
      workbook: newWb,
      activeSheetIndex: 0,
      activeCell: { r: 0, c: 0 },
      selection: { r1: 0, c1: 0, r2: 0, c2: 0 },
      isEditing: false,
      editingValue: '',
      filePath: null,
      fileName: 'ตารางบันทึก',
      isDirty: false,
      evaluator: newEval,
      history: newHist,
    });
    get().updateTitle();
  },

  loadWorkbook: async (data: Uint8Array, filePath?: string) => {
    const wb = await importFromXlsxBuffer(data);
    const evalInstance = new Evaluator(wb.sheets[0]);
    const histInstance = new HistoryManager(100);
    const fileName = filePath
      ? filePath.split(/[\\/]/).pop()?.replace(/\.[^.]+$/, '') || 'ตารางบันทึก'
      : 'ตารางบันทึก';

    set({
      workbook: wb,
      activeSheetIndex: 0,
      activeCell: { r: 0, c: 0 },
      selection: { r1: 0, c1: 0, r2: 0, c2: 0 },
      isEditing: false,
      editingValue: '',
      filePath: filePath || null,
      fileName,
      isDirty: false,
      evaluator: evalInstance,
      history: histInstance,
    });
    get().updateTitle();
  },

  saveWorkbook: async () => {
    const { workbook } = get();
    return await exportToXlsxBuffer(workbook);
  },

  markSaved: (filePath: string) => {
    const fileName = filePath.split(/[\\/]/).pop()?.replace(/\.[^.]+$/, '') || 'ตารางบันทึก';
    set({
      filePath,
      fileName,
      isDirty: false,
    });
    get().updateTitle();
  },

  getRawValue: (r: number, c: number) => {
    const { workbook, activeSheetIndex } = get();
    const key = cellToA1(r, c);
    return workbook.sheets[activeSheetIndex]?.cells[key]?.raw || '';
  },

  getComputedValue: (r: number, c: number) => {
    const { evaluator } = get();
    const key = cellToA1(r, c);
    return evaluator.getCellValue(key);
  },

  getDisplayValue: (r: number, c: number) => {
    const { workbook, activeSheetIndex, evaluator } = get();
    const key = cellToA1(r, c);
    const cell = workbook.sheets[activeSheetIndex]?.cells[key];
    const val = evaluator.getCellValue(key);
    return formatDisplayValue(val, cell?.fmt);
  },
}));
