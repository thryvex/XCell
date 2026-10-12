import { create } from 'zustand';
import {
  Workbook,
  Sheet,
  Cell,
  NumFmt,
  MergeRange,
  createEmptyWorkbook,
  CellValue,
} from '../core/model';
import { Evaluator } from '../core/evaluator';
import { HistoryManager, HistoryAction } from '../core/history';
import { convertShorthand } from '../core/shorthand';
import { cellToA1, parseA1, isCellInRange, cycleReference } from '../core/address';
import { exportToXlsxBuffer, importFromXlsxBuffer } from '../core/xlsx';
import { formatDisplayValue } from '../core/numfmt';
import { shiftFormula } from '../core/refshift';
import { insertRow as opInsertRow, deleteRow as opDeleteRow, insertCol as opInsertCol, deleteCol as opDeleteCol } from '../core/sheetOps';
import { rangeToTsv, rangeToHtml, parseTsv } from '../core/clipboard';
import { computeFill, fillDown as opFillDown, fillRight as opFillRight } from '../core/fill';
import { detectSigmaRange } from '../core/sigma';

export interface SelectionState {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

export interface InternalClipboard {
  cells: Record<string, Cell>;
  r1: number;
  c1: number;
  r2: number;
  c2: number;
  isCut?: boolean;
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

  // Phase 2 state
  clipboard: InternalClipboard | null;
  cutRange: SelectionState | null;

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

  // Phase 2 actions
  copySelection: () => void;
  cutSelection: () => void;
  paste: (valuesOnly?: boolean) => Promise<void>;
  insertRowAt: (atRow: number) => void;
  deleteRowAt: (atRow: number) => void;
  insertColAt: (atCol: number) => void;
  deleteColAt: (atCol: number) => void;
  setColWidth: (colIdx: number, width: number) => void;
  setRowHeight: (rowIdx: number, height: number) => void;
  freezeTo: (rows: number, cols: number) => void;
  applyFillRange: (target: SelectionState) => void;
  fillDownAction: () => void;
  fillRightAction: () => void;
  cycleCurrentReference: () => void;
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
  clipboard: null,
  cutRange: null,

  updateTitle: () => {
    const { fileName, isDirty } = get();
    const title = `${fileName}${isDirty ? ' ●' : ''} - SheetEasy`;
    if (typeof window !== 'undefined' && (window as any).electronAPI) {
      (window as any).electronAPI.setTitle(title);
    }
  },

  setActiveCell: (r, c, extend = false) => {
    const { rowCount, colCount, selection, isEditing, commitEditing } = get();
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

    sheet.merges = sheet.merges.filter(
      (m) =>
        m.r2 < selection.r1 ||
        m.r1 > selection.r2 ||
        m.c2 < selection.c1 ||
        m.c1 > selection.c2
    );

    const topLeftKey = cellToA1(selection.r1, selection.c1);
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
      activeCell,
      workbook,
      activeSheetIndex,
      startEditing,
      evaluator,
      history,
      updateTitle,
    } = get();
    const isMultiCell = selection.r1 !== selection.r2 || selection.c1 !== selection.c2;
    const sheet = workbook.sheets[activeSheetIndex];

    if (!isMultiCell) {
      // Auto-detect range above or to the left (Phase 2 feature)
      const detected = detectSigmaRange(sheet, activeCell.r, activeCell.c);
      if (detected) {
        const formula = `=${fnName}(${detected})`;
        const key = cellToA1(activeCell.r, activeCell.c);
        const prev = sheet.cells[key] ? { ...sheet.cells[key] } : undefined;

        sheet.cells[key] = { raw: formula };
        evaluator.invalidate(key);

        history.push({
          description: `Auto ${fnName}`,
          sheetIndex: activeSheetIndex,
          patch: { cellPatches: [{ key, before: prev, after: { raw: formula } }] },
        });

        set({ workbook: { ...workbook }, isDirty: true });
        updateTitle();
        return;
      }

      // Fallback: insert =FUNCTION( into editor
      startEditing(`=${fnName}(`);
      return;
    }

    // Multi-cell selection: insert formula below the range
    const targetRow = selection.r2 + 1;
    const patches: any[] = [];

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
      clipboard: null,
      cutRange: null,
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
      clipboard: null,
      cutRange: null,
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

  // --- Phase 2 Actions ---

  copySelection: () => {
    const { selection, workbook, activeSheetIndex, getDisplayValue } = get();
    const sheet = workbook.sheets[activeSheetIndex];
    const copiedCells: Record<string, Cell> = {};

    for (let r = selection.r1; r <= selection.r2; r++) {
      for (let c = selection.c1; c <= selection.c2; c++) {
        const key = cellToA1(r, c);
        if (sheet.cells[key]) {
          copiedCells[key] = { ...sheet.cells[key] };
        }
      }
    }

    const tsv = rangeToTsv(sheet, selection, getDisplayValue);
    const html = rangeToHtml(sheet, selection, getDisplayValue);

    if (navigator?.clipboard?.write) {
      const textBlob = new Blob([tsv], { type: 'text/plain' });
      const htmlBlob = new Blob([html], { type: 'text/html' });
      navigator.clipboard.write([
        new ClipboardItem({
          'text/plain': textBlob,
          'text/html': htmlBlob,
        }),
      ]).catch(() => {});
    }

    set({
      clipboard: {
        cells: copiedCells,
        r1: selection.r1,
        c1: selection.c1,
        r2: selection.r2,
        c2: selection.c2,
        isCut: false,
      },
      cutRange: null,
    });
  },

  cutSelection: () => {
    get().copySelection();
    set({
      clipboard: {
        ...get().clipboard!,
        isCut: true,
      },
      cutRange: { ...get().selection },
    });
  },

  paste: async (valuesOnly = false) => {
    const {
      clipboard,
      activeCell,
      selection,
      workbook,
      activeSheetIndex,
      evaluator,
      history,
      updateTitle,
    } = get();

    const sheet = workbook.sheets[activeSheetIndex];
    const patches: any[] = [];

    if (clipboard && Object.keys(clipboard.cells).length > 0) {
      const dr = activeCell.r - clipboard.r1;
      const dc = activeCell.c - clipboard.c1;

      // Check if copying single cell into multi-cell target selection
      const isSingleCopied = clipboard.r1 === clipboard.r2 && clipboard.c1 === clipboard.c2;
      const isTargetMulti = selection.r1 !== selection.r2 || selection.c1 !== selection.c2;

      if (isSingleCopied && isTargetMulti) {
        const srcCell = clipboard.cells[cellToA1(clipboard.r1, clipboard.c1)];
        for (let r = selection.r1; r <= selection.r2; r++) {
          for (let c = selection.c1; c <= selection.c2; c++) {
            const destKey = cellToA1(r, c);
            const prev = sheet.cells[destKey] ? { ...sheet.cells[destKey] } : undefined;

            let rawVal = srcCell ? srcCell.raw : '';
            if (rawVal.startsWith('=')) {
              if (valuesOnly) {
                rawVal = String(evaluator.getCellValue(cellToA1(clipboard.r1, clipboard.c1)));
              } else {
                rawVal = shiftFormula(rawVal, r - clipboard.r1, c - clipboard.c1);
              }
            }

            sheet.cells[destKey] = {
              raw: rawVal,
              fmt: srcCell?.fmt,
              style: srcCell?.style,
            };
            evaluator.invalidate(destKey);
            patches.push({ key: destKey, before: prev, after: { ...sheet.cells[destKey] } });
          }
        }
      } else {
        // Normal paste with dr, dc
        for (const [srcKey, cell] of Object.entries(clipboard.cells)) {
          const p = parseA1(srcKey);
          if (!p) continue;

          const destR = p.r + dr;
          const destC = p.c + dc;
          const destKey = cellToA1(destR, destC);
          const prev = sheet.cells[destKey] ? { ...sheet.cells[destKey] } : undefined;

          let rawVal = cell.raw;
          if (rawVal.startsWith('=')) {
            if (valuesOnly) {
              rawVal = String(evaluator.getCellValue(srcKey));
            } else {
              rawVal = shiftFormula(rawVal, dr, dc);
            }
          }

          sheet.cells[destKey] = {
            raw: rawVal,
            fmt: cell.fmt,
            style: cell.style,
          };
          evaluator.invalidate(destKey);
          patches.push({ key: destKey, before: prev, after: { ...sheet.cells[destKey] } });
        }
      }

      // If cut operation, clear original cut cells
      if (clipboard.isCut) {
        for (const srcKey of Object.keys(clipboard.cells)) {
          if (!sheet.cells[srcKey] || patches.some((p) => p.key === srcKey)) continue;
          patches.push({ key: srcKey, before: { ...sheet.cells[srcKey] }, after: undefined });
          delete sheet.cells[srcKey];
          evaluator.invalidate(srcKey);
        }
      }

      history.push({
        description: valuesOnly ? 'Paste values' : 'Paste',
        sheetIndex: activeSheetIndex,
        patch: { cellPatches: patches },
      });

      set({
        workbook: { ...workbook },
        cutRange: null,
        clipboard: clipboard.isCut ? null : clipboard,
        isDirty: true,
      });
      updateTitle();
      return;
    }

    // External clipboard fallback
    if (navigator?.clipboard?.readText) {
      try {
        const text = await navigator.clipboard.readText();
        const grid = parseTsv(text);
        if (grid.length === 0) return;

        for (let i = 0; i < grid.length; i++) {
          for (let j = 0; j < grid[i].length; j++) {
            const destR = activeCell.r + i;
            const destC = activeCell.c + j;
            const destKey = cellToA1(destR, destC);
            const prev = sheet.cells[destKey] ? { ...sheet.cells[destKey] } : undefined;
            const rawVal = grid[i][j];

            sheet.cells[destKey] = { raw: rawVal };
            evaluator.invalidate(destKey);
            patches.push({ key: destKey, before: prev, after: { raw: rawVal } });
          }
        }

        history.push({
          description: 'Paste external',
          sheetIndex: activeSheetIndex,
          patch: { cellPatches: patches },
        });

        set({ workbook: { ...workbook }, isDirty: true });
        updateTitle();
      } catch {}
    }
  },

  insertRowAt: (atRow: number) => {
    const { workbook, activeSheetIndex, evaluator, history, rowCount, updateTitle } = get();
    const sheet = workbook.sheets[activeSheetIndex];
    opInsertRow(sheet, atRow, 1);
    evaluator.setSheet(sheet);

    set({
      workbook: { ...workbook },
      rowCount: rowCount + 1,
      isDirty: true,
    });
    updateTitle();
  },

  deleteRowAt: (atRow: number) => {
    const { workbook, activeSheetIndex, evaluator, history, rowCount, updateTitle } = get();
    const sheet = workbook.sheets[activeSheetIndex];
    opDeleteRow(sheet, atRow, 1);
    evaluator.setSheet(sheet);

    set({
      workbook: { ...workbook },
      rowCount: Math.max(1, rowCount - 1),
      isDirty: true,
    });
    updateTitle();
  },

  insertColAt: (atCol: number) => {
    const { workbook, activeSheetIndex, evaluator, history, colCount, updateTitle } = get();
    const sheet = workbook.sheets[activeSheetIndex];
    opInsertCol(sheet, atCol, 1);
    evaluator.setSheet(sheet);

    set({
      workbook: { ...workbook },
      colCount: colCount + 1,
      isDirty: true,
    });
    updateTitle();
  },

  deleteColAt: (atCol: number) => {
    const { workbook, activeSheetIndex, evaluator, history, colCount, updateTitle } = get();
    const sheet = workbook.sheets[activeSheetIndex];
    opDeleteCol(sheet, atCol, 1);
    evaluator.setSheet(sheet);

    set({
      workbook: { ...workbook },
      colCount: Math.max(1, colCount - 1),
      isDirty: true,
    });
    updateTitle();
  },

  setColWidth: (colIdx: number, width: number) => {
    const { workbook, activeSheetIndex } = get();
    const sheet = workbook.sheets[activeSheetIndex];
    sheet.colWidths[colIdx] = Math.max(30, width);
    set({ workbook: { ...workbook }, isDirty: true });
  },

  setRowHeight: (rowIdx: number, height: number) => {
    const { workbook, activeSheetIndex } = get();
    const sheet = workbook.sheets[activeSheetIndex];
    sheet.rowHeights[rowIdx] = Math.max(18, height);
    set({ workbook: { ...workbook }, isDirty: true });
  },

  freezeTo: (rows: number, cols: number) => {
    const { workbook, activeSheetIndex } = get();
    const sheet = workbook.sheets[activeSheetIndex];
    sheet.frozenRows = rows;
    sheet.frozenCols = cols;
    set({ workbook: { ...workbook }, isDirty: true });
  },

  applyFillRange: (target: SelectionState) => {
    const { selection, workbook, activeSheetIndex, evaluator, history, updateTitle } = get();
    const sheet = workbook.sheets[activeSheetIndex];
    const fillResult = computeFill(sheet, selection, target);

    const patches: any[] = [];
    for (const [key, cell] of Object.entries(fillResult.cellsToSet)) {
      const prev = sheet.cells[key] ? { ...sheet.cells[key] } : undefined;
      sheet.cells[key] = { ...cell };
      evaluator.invalidate(key);
      patches.push({ key, before: prev, after: { ...cell } });
    }

    if (patches.length > 0) {
      history.push({
        description: 'Auto fill',
        sheetIndex: activeSheetIndex,
        patch: { cellPatches: patches },
      });
      set({
        workbook: { ...workbook },
        selection: { ...target },
        isDirty: true,
      });
      updateTitle();
    }
  },

  fillDownAction: () => {
    const { selection, workbook, activeSheetIndex, evaluator, history, updateTitle } = get();
    const sheet = workbook.sheets[activeSheetIndex];
    const res = opFillDown(sheet, selection);

    const patches: any[] = [];
    for (const [key, cell] of Object.entries(res)) {
      const prev = sheet.cells[key] ? { ...sheet.cells[key] } : undefined;
      sheet.cells[key] = { ...cell };
      evaluator.invalidate(key);
      patches.push({ key, before: prev, after: { ...cell } });
    }

    if (patches.length > 0) {
      history.push({
        description: 'Fill down',
        sheetIndex: activeSheetIndex,
        patch: { cellPatches: patches },
      });
      set({ workbook: { ...workbook }, isDirty: true });
      updateTitle();
    }
  },

  fillRightAction: () => {
    const { selection, workbook, activeSheetIndex, evaluator, history, updateTitle } = get();
    const sheet = workbook.sheets[activeSheetIndex];
    const res = opFillRight(sheet, selection);

    const patches: any[] = [];
    for (const [key, cell] of Object.entries(res)) {
      const prev = sheet.cells[key] ? { ...sheet.cells[key] } : undefined;
      sheet.cells[key] = { ...cell };
      evaluator.invalidate(key);
      patches.push({ key, before: prev, after: { ...cell } });
    }

    if (patches.length > 0) {
      history.push({
        description: 'Fill right',
        sheetIndex: activeSheetIndex,
        patch: { cellPatches: patches },
      });
      set({ workbook: { ...workbook }, isDirty: true });
      updateTitle();
    }
  },

  cycleCurrentReference: () => {
    const { isEditing, editingValue, setEditingValue } = get();
    if (!isEditing || !editingValue.startsWith('=')) return;

    // Find the last reference in formula
    const match = /([$]?[A-Za-z]+[$]?\d+)(?!.*[$]?[A-Za-z]+[$]?\d+)/.exec(editingValue);
    if (match) {
      const ref = match[1];
      const cycled = cycleReference(ref);
      const start = match.index;
      const end = start + ref.length;
      const nextVal = editingValue.slice(0, start) + cycled + editingValue.slice(end);
      setEditingValue(nextVal);
    }
  },
}));
