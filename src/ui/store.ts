import { create } from 'zustand';
import {
  Workbook,
  Sheet,
  Cell,
  CellStyle,
  NumFmt,
  MergeRange,
  createEmptyWorkbook,
  CellValue,
  ensureSheetIds,
} from '../core/model';
import { Evaluator } from '../core/evaluator';
import { HistoryManager, HistoryAction, snapshotWorkbook, SheetCellPatch } from '../core/history';
import {
  addSheet as opAddSheet,
  duplicateSheet as opDuplicateSheet,
  renameSheet as opRenameSheet,
  deleteSheet as opDeleteSheet,
  moveSheet as opMoveSheet,
  adjacentVisibleSheet,
  nearestVisibleSheet,
  SheetOpResult,
  countSheetReferences,
} from '../core/workbookOps';

import { th } from '../i18n/th';
import { convertShorthand } from '../core/shorthand';
import { cellToA1, parseA1, isCellInRange, cycleReference } from '../core/address';
import { exportToXlsxBuffer, importFromXlsxBuffer } from '../core/xlsx';
import { formatDisplayValue } from '../core/numfmt';
import { shiftFormula } from '../core/refshift';
import {
  insertRow as opInsertRow,
  deleteRow as opDeleteRow,
  insertCol as opInsertCol,
  deleteCol as opDeleteCol,
  countFormulasBrokenByDelete,
  adjustOtherSheetsForRowCol,
} from '../core/sheetOps';
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
  sourceSheetId?: string; // ชีตต้นทาง (สำหรับตัดข้ามชีต และวางเฉพาะค่า)
}

/** สถานะการมองเห็นที่แยกอิสระต่อชีต (ความกว้าง/สูง/การตรึง เก็บอยู่ใน Sheet แล้ว) */
export interface SheetView {
  activeCell: { r: number; c: number };
  selection: SelectionState;
  rowCount: number;
  colCount: number;
  scrollTop: number;
  scrollLeft: number;
}

export interface PendingDeleteSheet {
  sheetIndex: number;
  sheetName: string;
  referencingFormulaCount: number;
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
  clearSelectionFormat: () => void;
  insertCurrentDate: () => void;
  selectAll: () => void;
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
  deleteRowAt: (atRow: number, force?: boolean) => void;
  insertColAt: (atCol: number) => void;
  deleteColAt: (atCol: number, force?: boolean) => void;
  addRowsAtBottom: (count?: number) => void;
  setColWidth: (colIdx: number, width: number) => void;
  setRowHeight: (rowIdx: number, height: number) => void;
  freezeTo: (rows: number, cols: number) => void;
  applyFillRange: (target: SelectionState) => void;
  fillDownAction: () => void;
  fillRightAction: () => void;
  cycleCurrentReference: () => void;

  // Phase 3A styling actions
  getActiveCellStyle: () => CellStyle;
  setCellStyle: (stylePatch: Partial<CellStyle>) => void;
  toggleBold: () => void;
  toggleItalic: () => void;
  toggleUnderline: () => void;
  toggleWrap: () => void;
  setAlignment: (align?: 'left' | 'center' | 'right') => void;
  setTextColor: (color?: string) => void;
  setFillColor: (fill?: string) => void;
  setBorder: (border?: 'none' | 'all' | 'outer') => void;

  // Phase 3B state & actions
  pendingDelete: { type: 'row' | 'col'; index: number; count: number } | null;
  confirmPendingDelete: () => void;
  cancelPendingDelete: () => void;
  isFormulaHelperOpen: boolean;
  openFormulaHelper: () => void;
  closeFormulaHelper: () => void;
  insertFormulaFunction: (fnName: string) => void;

  // Phase 3C state & actions
  recentFiles: string[];
  fetchRecentFiles: () => Promise<void>;
  addRecentFile: (filePath: string) => Promise<void>;
  removeRecentFile: (filePath: string) => Promise<void>;
  recoverySessionId: string;
  saveRecoverySnapshot: () => Promise<void>;
  clearRecoverySnapshot: () => Promise<void>;
  restoreRecoveryWorkbook: (recoveredWb: Workbook, origPath: string | null) => void;
  unsupportedFeatures: string[];
  hasWarnedUnsupportedOverwrite: boolean;
  setHasWarnedUnsupportedOverwrite: (warned: boolean) => void;
  dismissUnsupportedBanner: () => void;
  loadTemplate: (wb: Workbook, templateName: string) => void;

  // Phase 4B: หลายชีต
  sheetViews: Record<string, SheetView>;
  pendingDeleteSheet: PendingDeleteSheet | null;
  switchSheet: (index: number) => void;
  nextSheet: () => void;
  prevSheet: () => void;
  addSheet: () => void;
  duplicateSheet: (index: number) => void;
  renameSheet: (index: number, newName: string) => { success: boolean; error?: string };
  deleteSheet: (index: number) => { success: boolean; error?: string };
  confirmDeleteSheet: () => void;
  cancelDeleteSheet: () => void;
  moveSheet: (from: number, to: number) => void;
  rememberScroll: (scrollTop: number, scrollLeft: number) => void;
  getActiveSheetView: () => SheetView;
}

const DEFAULT_ROW_COUNT = 100;
const DEFAULT_COL_COUNT = 26;

function defaultSheetView(activeCell = { r: 0, c: 0 }): SheetView {
  return {
    activeCell: { ...activeCell },
    selection: { r1: activeCell.r, c1: activeCell.c, r2: activeCell.r, c2: activeCell.c },
    rowCount: DEFAULT_ROW_COUNT,
    colCount: DEFAULT_COL_COUNT,
    scrollTop: 0,
    scrollLeft: 0,
  };
}

/**
 * สถานะเริ่มต้นเมื่อโหลดสมุดงานใหม่ (ใหม่ / เปิดไฟล์ / กู้คืน / เทมเพลต)
 */
function freshWorkbookState(wb: Workbook, activeCell = { r: 0, c: 0 }) {
  ensureSheetIds(wb);
  const activeSheetIndex = nearestVisibleSheet(wb, wb.activeSheet || 0);
  wb.activeSheet = activeSheetIndex;
  const view = defaultSheetView(activeCell);
  return {
    workbook: wb,
    activeSheetIndex,
    activeCell: view.activeCell,
    selection: view.selection,
    rowCount: view.rowCount,
    colCount: view.colCount,
    sheetViews: { [wb.sheets[activeSheetIndex].id!]: view } as Record<string, SheetView>,
    isEditing: false,
    editingValue: '',
    evaluator: new Evaluator(wb),
    history: new HistoryManager(100),
    clipboard: null,
    cutRange: null,
    pendingDelete: null,
    pendingDeleteSheet: null,
  };
}

type StoreGet = () => SheetState;
type StoreSet = (partial: Partial<SheetState>) => void;

/** บันทึกสถานะการมองเห็นของชีตที่เปิดอยู่ลงใน sheetViews (คืนค่า object ใหม่) */
function captureView(get: StoreGet): Record<string, SheetView> {
  const s = get();
  const id = s.workbook.sheets[s.activeSheetIndex]?.id;
  if (!id) return s.sheetViews;
  const prev = s.sheetViews[id];
  return {
    ...s.sheetViews,
    [id]: {
      activeCell: { ...s.activeCell },
      selection: { ...s.selection },
      rowCount: s.rowCount,
      colCount: s.colCount,
      scrollTop: prev?.scrollTop ?? 0,
      scrollLeft: prev?.scrollLeft ?? 0,
    },
  };
}

/** สถานะ store ที่ใช้แสดงชีต index (ถ้ายังไม่มีสถานะเดิม จะสร้างค่าเริ่มต้นใน views) */
function viewStateFor(wb: Workbook, index: number, views: Record<string, SheetView>) {
  const id = wb.sheets[index]?.id;
  let view = id ? views[id] : undefined;
  if (!view) {
    view = defaultSheetView();
    if (id) views[id] = view;
  }
  return {
    activeSheetIndex: index,
    activeCell: { ...view.activeCell },
    selection: { ...view.selection },
    rowCount: view.rowCount,
    colCount: view.colCount,
  };
}

/** หลัง Undo/Redo: คำนวณใหม่ทั้งสมุดงาน และสลับไปชีตที่ได้รับผลกระทบ */
function finishHistoryStep(
  get: StoreGet,
  set: StoreSet,
  views: Record<string, SheetView>,
  targetIndex: number,
  full?: { rowCount?: number; colCount?: number }
) {
  const { workbook, evaluator, activeSheetIndex, updateTitle } = get();
  const nextWb = { ...workbook };
  const index = nearestVisibleSheet(nextWb, targetIndex);
  nextWb.activeSheet = index;
  evaluator.setWorkbook(nextWb);

  const viewState =
    index === activeSheetIndex && nextWb.sheets[index]?.id === workbook.sheets[activeSheetIndex]?.id
      ? {}
      : viewStateFor(nextWb, index, views);

  set({
    workbook: nextWb,
    ...viewState,
    activeSheetIndex: index,
    ...(full?.rowCount !== undefined ? { rowCount: full.rowCount } : {}),
    ...(full?.colCount !== undefined ? { colCount: full.colCount } : {}),
    sheetViews: views,
    isEditing: false,
    editingValue: '',
    isDirty: true,
  });
  updateTitle();
}

/**
 * ทำรายการระดับชีต (เพิ่ม/ลบ/เปลี่ยนชื่อ/ทำสำเนา/ย้าย) พร้อมบันทึกประวัติ Undo/Redo
 * - affectedBefore: ชีตที่ Undo จะสลับกลับไป
 * - show: 'result' = แสดงชีตที่ได้จากรายการ, 'keep' = คงชีตที่เปิดอยู่เดิม
 */
function runSheetOp(
  get: StoreGet,
  set: StoreSet,
  description: string,
  mutate: (wb: Workbook) => SheetOpResult,
  opts: { affectedBefore: number; show: 'result' | 'keep' }
): SheetOpResult {
  const { isEditing, commitEditing } = get();
  if (isEditing) commitEditing(0, 0);

  const { workbook, activeSheetIndex, history, evaluator, updateTitle } = get();
  const activeId = workbook.sheets[activeSheetIndex]?.id;
  const views = captureView(get);
  const before = snapshotWorkbook(workbook);
  before.activeSheet = opts.affectedBefore;

  const res = mutate(workbook);
  if (!res.ok) return res;

  const affectedAfter = res.index ?? activeSheetIndex;
  let showIndex = affectedAfter;
  if (opts.show === 'keep') {
    const kept = workbook.sheets.findIndex((s) => s.id === activeId);
    showIndex = kept >= 0 ? kept : affectedAfter;
  }

  const after = snapshotWorkbook(workbook);
  after.activeSheet = affectedAfter;
  history.push({ description, workbookBefore: before, workbookAfter: after });

  const nextWb = { ...workbook, activeSheet: showIndex };
  evaluator.setWorkbook(nextWb);
  set({
    workbook: nextWb,
    ...viewStateFor(nextWb, showIndex, views),
    sheetViews: views,
    isDirty: true,
  });
  updateTitle();
  return res;
}

/**
 * แทรก/ลบแถว-คอลัมน์ในชีตที่เปิดอยู่ และปรับสูตรในชีตอื่นที่อ้างถึงชีตนี้ด้วย
 */
function applyRowColOp(
  get: StoreGet,
  set: StoreSet,
  op: 'insert' | 'delete',
  axis: 'row' | 'col',
  index: number
) {
  const { workbook, activeSheetIndex, evaluator, history, rowCount, colCount, updateTitle } = get();
  const sheet = workbook.sheets[activeSheetIndex];
  const fullBefore = {
    cells: JSON.parse(JSON.stringify(sheet.cells)),
    merges: JSON.parse(JSON.stringify(sheet.merges)),
    rowCount,
    colCount,
  };

  if (op === 'insert') {
    if (axis === 'row') opInsertRow(sheet, index, 1);
    else opInsertCol(sheet, index, 1);
  } else {
    if (axis === 'row') opDeleteRow(sheet, index, 1);
    else opDeleteCol(sheet, index, 1);
  }
  const extraCellPatches = adjustOtherSheetsForRowCol(workbook, activeSheetIndex, op, axis, index, 1);
  evaluator.setWorkbook(workbook);

  const delta = op === 'insert' ? 1 : -1;
  const newRowCount = axis === 'row' ? Math.max(1, rowCount + delta) : rowCount;
  const newColCount = axis === 'col' ? Math.max(1, colCount + delta) : colCount;
  const fullAfter = {
    cells: JSON.parse(JSON.stringify(sheet.cells)),
    merges: JSON.parse(JSON.stringify(sheet.merges)),
    rowCount: newRowCount,
    colCount: newColCount,
  };

  const descriptions = {
    'insert-row': 'Insert row',
    'delete-row': 'Delete row',
    'insert-col': 'Insert column',
    'delete-col': 'Delete column',
  } as const;

  history.push({
    description: descriptions[`${op}-${axis}`],
    sheetIndex: activeSheetIndex,
    patch: { fullSheetBefore: fullBefore, fullSheetAfter: fullAfter },
    extraCellPatches: extraCellPatches.length > 0 ? extraCellPatches : undefined,
  });

  set({
    workbook: { ...workbook },
    rowCount: newRowCount,
    colCount: newColCount,
    isDirty: true,
  });
  updateTitle();
}

const initialWb = ensureSheetIds(createEmptyWorkbook());
const initialEvaluator = new Evaluator(initialWb);
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
  pendingDelete: null,
  isFormulaHelperOpen: false,
  recentFiles: [],
  recoverySessionId: `recovery-${Date.now()}`,
  unsupportedFeatures: [],
  hasWarnedUnsupportedOverwrite: false,
  sheetViews: { [initialWb.sheets[0].id!]: defaultSheetView() },
  pendingDeleteSheet: null,

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
    if (finalValue.startsWith('=')) {
      const openCount = (finalValue.match(/\(/g) || []).length;
      const closeCount = (finalValue.match(/\)/g) || []).length;
      if (openCount > closeCount) {
        finalValue += ')'.repeat(openCount - closeCount);
      }
    }
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

      evaluator.invalidate(key, activeSheetIndex);

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
        const cell = sheet.cells[key];
        if (cell && (cell.raw !== '' || cell.raw !== undefined)) {
          const before = { ...cell };
          let after: Cell | undefined;
          if (cell.style || cell.fmt) {
            after = { raw: '', fmt: cell.fmt, style: cell.style ? { ...cell.style } : undefined };
            sheet.cells[key] = after;
          } else {
            after = undefined;
            delete sheet.cells[key];
          }
          patches.push({ key, before, after });
          evaluator.invalidate(key, activeSheetIndex);
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

  clearSelectionFormat: () => {
    const { selection, workbook, activeSheetIndex, history, updateTitle } = get();
    const sheet = workbook.sheets[activeSheetIndex];
    const patches: any[] = [];

    for (let r = selection.r1; r <= selection.r2; r++) {
      for (let c = selection.c1; c <= selection.c2; c++) {
        const key = cellToA1(r, c);
        const cell = sheet.cells[key];
        if (cell && (cell.style || cell.fmt)) {
          const before = { ...cell };
          let after: Cell | undefined;
          if (cell.raw) {
            after = { raw: cell.raw };
            sheet.cells[key] = after;
          } else {
            after = undefined;
            delete sheet.cells[key];
          }
          patches.push({ key, before, after });
        }
      }
    }

    if (patches.length > 0) {
      history.push({
        description: 'Clear formatting',
        sheetIndex: activeSheetIndex,
        patch: { cellPatches: patches },
      });
      set({ workbook: { ...workbook }, isDirty: true });
      updateTitle();
    }
  },

  insertCurrentDate: () => {
    const { selection, workbook, activeSheetIndex, evaluator, history, updateTitle, isEditing, editingValue, setEditingValue } = get();
    const now = new Date();
    const day = String(now.getDate()).padStart(2, '0');
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const year = now.getFullYear();
    const dateStr = `${day}/${month}/${year}`;

    if (isEditing) {
      setEditingValue(editingValue + dateStr);
      return;
    }

    const sheet = workbook.sheets[activeSheetIndex];
    const patches: any[] = [];

    for (let r = selection.r1; r <= selection.r2; r++) {
      for (let c = selection.c1; c <= selection.c2; c++) {
        const key = cellToA1(r, c);
        const cell = sheet.cells[key];
        const before = cell ? { ...cell } : undefined;
        const after: Cell = {
          raw: dateStr,
          fmt: 'date',
          style: cell?.style ? { ...cell.style } : undefined,
        };
        sheet.cells[key] = after;
        patches.push({ key, before, after });
        evaluator.invalidate(key, activeSheetIndex);
      }
    }

    if (patches.length > 0) {
      history.push({
        description: 'Insert current date',
        sheetIndex: activeSheetIndex,
        patch: { cellPatches: patches },
      });
      set({ workbook: { ...workbook }, isDirty: true });
      updateTitle();
    }
  },

  selectAll: () => {
    const { rowCount, colCount } = get();
    get().setSelection(0, 0, rowCount - 1, colCount - 1);
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

  getActiveCellStyle: () => {
    const { workbook, activeSheetIndex, activeCell } = get();
    const sheet = workbook.sheets[activeSheetIndex];
    const key = cellToA1(activeCell.r, activeCell.c);
    return sheet.cells[key]?.style || {};
  },

  setCellStyle: (stylePatch: Partial<CellStyle>) => {
    const { selection, workbook, activeSheetIndex, history, updateTitle } = get();
    const sheet = workbook.sheets[activeSheetIndex];
    const patches: any[] = [];

    for (let r = selection.r1; r <= selection.r2; r++) {
      for (let c = selection.c1; c <= selection.c2; c++) {
        const key = cellToA1(r, c);
        const cell = sheet.cells[key] || { raw: '' };
        const prevStyle = cell.style ? { ...cell.style } : {};
        const mergedStyle: CellStyle = { ...prevStyle, ...stylePatch };

        const cleanStyle: CellStyle = {};
        for (const [k, v] of Object.entries(mergedStyle)) {
          if (v !== undefined && v !== null && v !== '' && v !== 'none') {
            (cleanStyle as any)[k] = v;
          }
        }

        const newStyle = Object.keys(cleanStyle).length > 0 ? cleanStyle : undefined;
        const newCell: Cell = { ...cell, style: newStyle };

        patches.push({
          key,
          before: sheet.cells[key] ? { ...sheet.cells[key] } : undefined,
          after: newCell,
        });

        sheet.cells[key] = newCell;
      }
    }

    if (patches.length > 0) {
      history.push({
        description: 'Format cells',
        sheetIndex: activeSheetIndex,
        patch: { cellPatches: patches },
      });
      set({ workbook: { ...workbook }, isDirty: true });
      updateTitle();
    }
  },

  toggleBold: () => {
    const current = get().getActiveCellStyle();
    get().setCellStyle({ bold: current.bold ? undefined : true });
  },

  toggleItalic: () => {
    const current = get().getActiveCellStyle();
    get().setCellStyle({ italic: current.italic ? undefined : true });
  },

  toggleUnderline: () => {
    const current = get().getActiveCellStyle();
    get().setCellStyle({ underline: current.underline ? undefined : true });
  },

  toggleWrap: () => {
    const current = get().getActiveCellStyle();
    get().setCellStyle({ wrap: current.wrap ? undefined : true });
  },

  setAlignment: (align?: 'left' | 'center' | 'right') => {
    const current = get().getActiveCellStyle();
    const newAlign = current.align === align ? undefined : align;
    get().setCellStyle({ align: newAlign });
  },

  setTextColor: (color?: string) => {
    get().setCellStyle({ color: color || undefined });
  },

  setFillColor: (fill?: string) => {
    get().setCellStyle({ fill: fill || undefined });
  },

  setBorder: (border?: 'none' | 'all' | 'outer') => {
    get().setCellStyle({ border: border === 'none' ? undefined : border });
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
        evaluator.invalidate(key, activeSheetIndex);

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
      evaluator.invalidate(formulaKey, activeSheetIndex);
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
    const { history, workbook } = get();
    if (!history.canUndo()) return;
    const views = captureView(get);
    const action = history.undo(workbook);
    if (!action) return;
    const target = action.workbookBefore
      ? action.workbookBefore.activeSheet
      : action.sheetIndex ?? get().activeSheetIndex;
    finishHistoryStep(get, set, views, target, action.patch?.fullSheetBefore);
  },

  redo: () => {
    const { history, workbook } = get();
    if (!history.canRedo()) return;
    const views = captureView(get);
    const action = history.redo(workbook);
    if (!action) return;
    const target = action.workbookAfter
      ? action.workbookAfter.activeSheet
      : action.sheetIndex ?? get().activeSheetIndex;
    finishHistoryStep(get, set, views, target, action.patch?.fullSheetAfter);
  },

  newWorkbook: () => {
    get().clearRecoverySnapshot();
    set({
      ...freshWorkbookState(createEmptyWorkbook()),
      filePath: null,
      fileName: 'ตารางบันทึก',
      isDirty: false,
      recoverySessionId: `recovery-${Date.now()}`,
      unsupportedFeatures: [],
      hasWarnedUnsupportedOverwrite: false,
    });
    get().updateTitle();
  },

  loadWorkbook: async (data: Uint8Array, filePath?: string) => {
    // Clear old recovery snapshot before loading new file
    await get().clearRecoverySnapshot();

    const wb = await importFromXlsxBuffer(data);
    const fileName = filePath
      ? filePath.split(/[\\/]/).pop()?.replace(/\.[^.]+$/, '') || 'ตารางบันทึก'
      : 'ตารางบันทึก';

    if (filePath) {
      get().addRecentFile(filePath);
    }

    set({
      ...freshWorkbookState(wb),
      filePath: filePath || null,
      fileName,
      isDirty: false,
      recoverySessionId: `recovery-${Date.now()}`,
      unsupportedFeatures: wb.unsupportedFeatures || [],
      hasWarnedUnsupportedOverwrite: false,
    });
    get().updateTitle();
  },

  saveWorkbook: async () => {
    const { workbook, activeSheetIndex } = get();
    return await exportToXlsxBuffer({ ...workbook, activeSheet: activeSheetIndex });
  },

  markSaved: (filePath: string) => {
    const fileName = filePath.split(/[\\/]/).pop()?.replace(/\.[^.]+$/, '') || 'ตารางบันทึก';
    get().addRecentFile(filePath);
    get().clearRecoverySnapshot();
    set({
      filePath,
      fileName,
      isDirty: false,
      recoverySessionId: `recovery-${Date.now()}`,
    });
    get().updateTitle();
  },

  fetchRecentFiles: async () => {
    const electronAPI = typeof window !== 'undefined' ? (window as any).electronAPI : undefined;
    if (electronAPI?.getRecentFiles) {
      try {
        const list = await electronAPI.getRecentFiles();
        if (Array.isArray(list)) set({ recentFiles: list });
      } catch {}
    }
  },

  addRecentFile: async (filePath: string) => {
    const electronAPI = typeof window !== 'undefined' ? (window as any).electronAPI : undefined;
    if (electronAPI?.addRecentFile) {
      try {
        const list = await electronAPI.addRecentFile(filePath);
        if (Array.isArray(list)) set({ recentFiles: list });
      } catch {}
    }
  },

  removeRecentFile: async (filePath: string) => {
    const electronAPI = typeof window !== 'undefined' ? (window as any).electronAPI : undefined;
    if (electronAPI?.removeRecentFile) {
      try {
        const list = await electronAPI.removeRecentFile(filePath);
        if (Array.isArray(list)) set({ recentFiles: list });
      } catch {}
    }
  },

  saveRecoverySnapshot: async () => {
    const { workbook, activeSheetIndex, filePath, fileName, isDirty, recoverySessionId } = get();
    if (!isDirty) return;
    const electronAPI = typeof window !== 'undefined' ? (window as any).electronAPI : undefined;
    if (electronAPI?.saveRecovery) {
      try {
        await electronAPI.saveRecovery({
          id: recoverySessionId,
          originalPath: filePath,
          fileName,
          savedAt: Date.now(),
          // เก็บทุกชีต (รวมชีตซ่อน) และชีตที่เปิดอยู่
          workbook: { ...workbook, activeSheet: activeSheetIndex },
        });
      } catch {}
    }
  },

  clearRecoverySnapshot: async () => {
    const { recoverySessionId } = get();
    const electronAPI = typeof window !== 'undefined' ? (window as any).electronAPI : undefined;
    if (electronAPI?.deleteRecovery) {
      try {
        await electronAPI.deleteRecovery(recoverySessionId);
      } catch {}
    }
  },

  restoreRecoveryWorkbook: (recoveredWb: Workbook, origPath: string | null) => {
    const fileName = origPath
      ? origPath.split(/[\\/]/).pop()?.replace(/\.[^.]+$/, '') || 'ตารางกู้คืน'
      : 'ตารางกู้คืน';

    set({
      ...freshWorkbookState(recoveredWb),
      filePath: origPath,
      fileName,
      isDirty: true,
      unsupportedFeatures: recoveredWb.unsupportedFeatures || [],
      hasWarnedUnsupportedOverwrite: false,
    });
    get().updateTitle();
  },

  setHasWarnedUnsupportedOverwrite: (warned: boolean) => {
    set({ hasWarnedUnsupportedOverwrite: warned });
  },

  dismissUnsupportedBanner: () => {
    set({ unsupportedFeatures: [] });
  },

  loadTemplate: (wb: Workbook, templateName: string) => {
    get().clearRecoverySnapshot();
    set({
      ...freshWorkbookState(wb, { r: 1, c: 0 }),
      filePath: null,
      fileName: templateName,
      isDirty: false,
      recoverySessionId: `recovery-${Date.now()}`,
      unsupportedFeatures: [],
      hasWarnedUnsupportedOverwrite: false,
    });
    get().updateTitle();
  },

  // --- Phase 4B: หลายชีต ---

  getActiveSheetView: () => {
    const { workbook, activeSheetIndex, sheetViews } = get();
    const id = workbook.sheets[activeSheetIndex]?.id;
    return (id && sheetViews[id]) || defaultSheetView();
  },

  rememberScroll: (scrollTop: number, scrollLeft: number) => {
    // แก้ค่าตรงๆ ไม่เรียก set เพื่อไม่ให้ตาราง re-render ทุกครั้งที่เลื่อน
    const { workbook, activeSheetIndex, sheetViews } = get();
    const id = workbook.sheets[activeSheetIndex]?.id;
    if (!id) return;
    if (!sheetViews[id]) sheetViews[id] = defaultSheetView();
    sheetViews[id].scrollTop = scrollTop;
    sheetViews[id].scrollLeft = scrollLeft;
  },

  switchSheet: (index: number) => {
    const { workbook, activeSheetIndex, isEditing, commitEditing } = get();
    const target = workbook.sheets[index];
    if (!target || target.hidden || index === activeSheetIndex) return;

    if (isEditing) {
      commitEditing(0, 0);
    }

    const views = captureView(get);
    workbook.activeSheet = index;
    set({ ...viewStateFor(workbook, index, views), sheetViews: views });
  },

  nextSheet: () => {
    const { workbook, activeSheetIndex, switchSheet } = get();
    const next = adjacentVisibleSheet(workbook, activeSheetIndex, 1);
    if (next !== null) switchSheet(next);
  },

  prevSheet: () => {
    const { workbook, activeSheetIndex, switchSheet } = get();
    const prev = adjacentVisibleSheet(workbook, activeSheetIndex, -1);
    if (prev !== null) switchSheet(prev);
  },

  addSheet: () => {
    runSheetOp(get, set, 'Add sheet', (wb) => opAddSheet(wb), {
      affectedBefore: get().activeSheetIndex,
      show: 'result',
    });
  },

  duplicateSheet: (index: number) => {
    runSheetOp(get, set, 'Duplicate sheet', (wb) => opDuplicateSheet(wb, index), {
      affectedBefore: index,
      show: 'result',
    });
  },

  renameSheet: (index: number, newName: string) => {
    const { workbook } = get();
    const sheet = workbook.sheets[index];
    if (!sheet) return { success: false };
    if (sheet.name === newName.trim()) return { success: true };

    const res = runSheetOp(get, set, 'Rename sheet', (wb) => opRenameSheet(wb, index, newName), {
      affectedBefore: index,
      show: 'keep',
    });
    return { success: res.ok, error: res.error };
  },

  deleteSheet: (index: number) => {
    const { workbook } = get();
    const sheet = workbook.sheets[index];
    if (!sheet) return { success: false };

    const visibleCount = workbook.sheets.filter((s) => !s.hidden).length;
    if (visibleCount <= 1) {
      return { success: false, error: th.sheetTabs.cannotDeleteLastSheet };
    }

    set({
      pendingDeleteSheet: {
        sheetIndex: index,
        sheetName: sheet.name,
        referencingFormulaCount: countSheetReferences(workbook, index),
      },
    });
    return { success: true };
  },

  confirmDeleteSheet: () => {
    const { pendingDeleteSheet } = get();
    if (!pendingDeleteSheet) return;
    set({ pendingDeleteSheet: null });
    const index = pendingDeleteSheet.sheetIndex;
    runSheetOp(get, set, 'Delete sheet', (wb) => opDeleteSheet(wb, index), {
      affectedBefore: index,
      show: get().activeSheetIndex === index ? 'result' : 'keep',
    });
  },

  cancelDeleteSheet: () => {
    set({ pendingDeleteSheet: null });
  },

  moveSheet: (from: number, to: number) => {
    runSheetOp(get, set, 'Move sheet', (wb) => opMoveSheet(wb, from, to), {
      affectedBefore: from,
      show: 'keep',
    });
  },

  getRawValue: (r: number, c: number) => {
    const { workbook, activeSheetIndex } = get();
    const key = cellToA1(r, c);
    return workbook.sheets[activeSheetIndex]?.cells[key]?.raw || '';
  },

  getComputedValue: (r: number, c: number) => {
    const { evaluator, activeSheetIndex } = get();
    const key = cellToA1(r, c);
    return evaluator.getCellValue(key, activeSheetIndex);
  },

  getDisplayValue: (r: number, c: number) => {
    const { workbook, activeSheetIndex, evaluator } = get();
    const key = cellToA1(r, c);
    const cell = workbook.sheets[activeSheetIndex]?.cells[key];
    const val = evaluator.getCellValue(key, activeSheetIndex);
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
        sourceSheetId: workbook.sheets[activeSheetIndex]?.id,
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
      // ชีตต้นทางของคลิปบอร์ด (อาจเป็นชีตอื่น) ใช้คำนวณค่าตอนวางเฉพาะค่า และล้างเซลตอนตัด
      const srcIndex = clipboard.sourceSheetId
        ? workbook.sheets.findIndex((s) => s.id === clipboard.sourceSheetId) // -1 = ชีตต้นทางถูกลบไปแล้ว
        : activeSheetIndex;
      const valueSheetIndex = srcIndex >= 0 ? srcIndex : activeSheetIndex;

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
                rawVal = String(evaluator.getCellValue(cellToA1(clipboard.r1, clipboard.c1), valueSheetIndex));
              } else {
                rawVal = shiftFormula(rawVal, r - clipboard.r1, c - clipboard.c1);
              }
            }

            sheet.cells[destKey] = {
              raw: rawVal,
              fmt: valuesOnly ? prev?.fmt : srcCell?.fmt,
              style: valuesOnly
                ? (prev?.style ? { ...prev.style } : undefined)
                : (srcCell?.style ? { ...srcCell.style } : undefined),
            };
            evaluator.invalidate(destKey, activeSheetIndex);
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
              rawVal = String(evaluator.getCellValue(srcKey, valueSheetIndex));
            } else {
              rawVal = shiftFormula(rawVal, dr, dc);
            }
          }

          sheet.cells[destKey] = {
            raw: rawVal,
            fmt: valuesOnly ? prev?.fmt : cell.fmt,
            style: valuesOnly
              ? (prev?.style ? { ...prev.style } : undefined)
              : (cell.style ? { ...cell.style } : undefined),
          };
          evaluator.invalidate(destKey, activeSheetIndex);
          patches.push({ key: destKey, before: prev, after: { ...sheet.cells[destKey] } });
        }
      }

      // If cut operation, clear original cut cells (ในชีตต้นทาง ซึ่งอาจเป็นชีตอื่น)
      const sourcePatches: SheetCellPatch[] = [];
      if (clipboard.isCut && srcIndex >= 0) {
        const srcSheet = workbook.sheets[srcIndex];
        const sameSheet = srcIndex === activeSheetIndex;
        for (const srcKey of Object.keys(clipboard.cells)) {
          if (!srcSheet.cells[srcKey]) continue;
          if (sameSheet && patches.some((p) => p.key === srcKey)) continue;
          const patch = { key: srcKey, before: { ...srcSheet.cells[srcKey] }, after: undefined };
          if (sameSheet) {
            patches.push(patch);
          } else {
            sourcePatches.push({ ...patch, sheetIndex: srcIndex });
          }
          delete srcSheet.cells[srcKey];
          evaluator.invalidate(srcKey, srcIndex);
        }
      }

      history.push({
        description: valuesOnly ? 'Paste values' : 'Paste',
        sheetIndex: activeSheetIndex,
        patch: { cellPatches: patches },
        extraCellPatches: sourcePatches.length > 0 ? sourcePatches : undefined,
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

            sheet.cells[destKey] = {
              raw: rawVal,
              fmt: prev?.fmt,
              style: prev?.style ? { ...prev.style } : undefined,
            };
            evaluator.invalidate(destKey, activeSheetIndex);
            patches.push({ key: destKey, before: prev, after: { ...sheet.cells[destKey] } });
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
    applyRowColOp(get, set, 'insert', 'row', atRow);
  },

  deleteRowAt: (atRow: number, force = false) => {
    const { workbook, activeSheetIndex } = get();
    if (!force) {
      const refCount = countFormulasBrokenByDelete(workbook, activeSheetIndex, 'row', atRow, 1);
      if (refCount > 0) {
        set({ pendingDelete: { type: 'row', index: atRow, count: refCount } });
        return;
      }
    }
    applyRowColOp(get, set, 'delete', 'row', atRow);
  },

  insertColAt: (atCol: number) => {
    applyRowColOp(get, set, 'insert', 'col', atCol);
  },

  deleteColAt: (atCol: number, force = false) => {
    const { workbook, activeSheetIndex } = get();
    if (!force) {
      const refCount = countFormulasBrokenByDelete(workbook, activeSheetIndex, 'col', atCol, 1);
      if (refCount > 0) {
        set({ pendingDelete: { type: 'col', index: atCol, count: refCount } });
        return;
      }
    }
    applyRowColOp(get, set, 'delete', 'col', atCol);
  },

  confirmPendingDelete: () => {
    const { pendingDelete, deleteRowAt, deleteColAt } = get();
    if (!pendingDelete) return;
    if (pendingDelete.type === 'row') {
      deleteRowAt(pendingDelete.index, true);
    } else {
      deleteColAt(pendingDelete.index, true);
    }
    set({ pendingDelete: null });
  },

  cancelPendingDelete: () => {
    set({ pendingDelete: null });
  },

  openFormulaHelper: () => set({ isFormulaHelperOpen: true }),
  closeFormulaHelper: () => set({ isFormulaHelperOpen: false }),

  insertFormulaFunction: (fnName: string) => {
    const { activeCell, selection, startEditing } = get();
    const isRangeSelected =
      selection.r1 !== selection.r2 || selection.c1 !== selection.c2;

    set({ isFormulaHelperOpen: false });

    if (isRangeSelected) {
      const startA1 = cellToA1(selection.r1, selection.c1);
      const endA1 = cellToA1(selection.r2, selection.c2);
      startEditing(`=${fnName}(${startA1}:${endA1})`);
    } else {
      startEditing(`=${fnName}(`);
    }
  },

  addRowsAtBottom: (count = 50) => {
    const { rowCount, updateTitle } = get();
    set({
      rowCount: rowCount + count,
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
      evaluator.invalidate(key, activeSheetIndex);
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
      evaluator.invalidate(key, activeSheetIndex);
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
      evaluator.invalidate(key, activeSheetIndex);
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
