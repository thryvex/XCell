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
  ColumnFilter,
  PrintSettings,
  DEFAULT_PRINT_SETTINGS,
} from '../core/model';
import { Evaluator } from '../core/evaluator';
import {
  HistoryManager,
  HistoryAction,
  snapshotWorkbook,
  SheetCellPatch,
  captureFilterState,
} from '../core/history';
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
import { formatDisplayValue, convertTypedDate, editTextForCell, todaySerial, formatDate, setDisplayYearMode, YearMode } from '../core/numfmt';
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
import {
  findMatches,
  nextMatchIndex,
  SearchMatch,
  SearchOptions,
  DEFAULT_SEARCH_OPTIONS,
} from '../core/search';
import {
  detectDataRegion,
  detectHeaderRow,
  rangeHasMerge,
  rangeHasData,
  formulasReferenceOtherRows,
  sortRange,
  CellRange,
  SortKey,
  SortDirection,
} from '../core/sort';
import {
  computeHiddenRows,
  createFilter,
} from '../core/filter';
import {
  CsvEncoding,
  detectEncoding,
  decodeBytes,
  detectDelimiter,
  looksBinary,
  parseCsv,
  csvExceedsLimits,
  buildSheetFromCsv,
  sheetNameFromFileName,
  sheetToCsvRows,
  writeCsv,
  encodeUtf8WithBom,
} from '../core/csv';

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
  // key ในคลิปบอร์ด → เซลต้นทางจริง (ต่างกันเมื่อข้ามแถวที่ตัวกรองซ่อน)
  sourceKeys?: Record<string, string>;
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
  // ชีตและเซลที่กำลังแก้ไข (ระหว่าง point mode ข้ามชีต ชีตที่แสดงอาจไม่ใช่ชีตนี้)
  editingSheetIndex: number | null;
  editingCell: { r: number; c: number } | null;
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
  getEditText: (r: number, c: number) => string; // ข้อความในแถบสูตร/ช่องแก้ไข (วันที่ตามโหมดปี)
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
  hiddenSheetNotice: number; // จำนวนชีตซ่อนในไฟล์ที่เปิด (แสดงในแถบแจ้งเตือน)
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

  // Phase 4C: ค้นหา (ไม่บันทึกลงไฟล์ ไม่อยู่ใน Undo)
  search: SearchState;
  scrollToActiveToken: number; // เพิ่มค่าเมื่อต้องเลื่อนตารางให้เห็นเซลที่เลือก (เช่น ไปยังผลค้นหา)
  openSearch: () => void;
  closeSearch: () => void;
  setSearchQuery: (query: string) => void;
  setSearchOption: (key: keyof SearchOptions, value: boolean) => void;
  searchNext: () => void;
  searchPrev: () => void;

  // Phase 4C: เรียงลำดับ
  pendingSort: PendingSort | null;
  notice: string | null; // ข้อความแจ้ง (ปุ่มตกลงอย่างเดียว)
  dismissNotice: () => void;
  requestSort: (dir: SortDirection) => void;
  openAdvancedSort: () => void;
  updatePendingSort: (patch: Partial<Pick<PendingSort, 'hasHeader' | 'keys'>>) => void;
  confirmSort: () => void;
  cancelSort: () => void;

  // Phase 4C: ตัวกรอง (แถวที่ซ่อนเกิดจากตัวกรองเท่านั้น)
  filteredRowsNotice: boolean; // ไฟล์ที่เปิดมีแถวที่ถูกซ่อนด้วยตัวกรอง
  toggleFilter: () => void;
  setColumnFilter: (col: number, cf: ColumnFilter | null) => void;
  reapplyFilter: () => void;
  sortFilterColumn: (col: number, dir: SortDirection) => void;
  showAllFilteredRows: () => void;
  dismissFilteredRowsNotice: () => void;

  // Phase 4C: ซูม (ต่อชีต, เก็บใน Sheet.zoom; ไม่อยู่ใน Undo)
  setZoom: (percent: number) => void;
  zoomIn: () => void;
  zoomOut: () => void;
  resetZoom: () => void;

  // Phase 4D: CSV
  pendingCsvImport: PendingCsvImport | null;
  beginCsvImport: (bytes: Uint8Array, filePath: string | null, mode: 'open' | 'newSheet') => void;
  updateCsvImport: (patch: Partial<Pick<PendingCsvImport, 'encoding' | 'delimiter' | 'autoConvert'>>) => void;
  confirmCsvImport: () => void;
  cancelCsvImport: () => void;
  getCsvExport: () => { bytes: Uint8Array; defaultName: string; hasFilter: boolean; hasDates: boolean };

  // Phase 5A: รูปแบบปีที่แสดง (ค่าของแอป เก็บใน userData)
  yearMode: YearMode;
  setYearMode: (mode: YearMode, persist?: boolean) => void;

  // Phase 5B: ตัวเลือกพิมพ์ (จำต่อไฟล์ เก็บใน workbook จึงอยู่ในไฟล์กู้คืนด้วย)
  setPrintSettings: (patch: Partial<PrintSettings>) => void;
}

export interface PendingCsvImport {
  mode: 'open' | 'newSheet'; // open = งานใหม่จาก CSV, newSheet = เพิ่มเป็นชีตใหม่
  filePath: string | null;
  fileName: string;
  bytes: Uint8Array;
  encoding: CsvEncoding;
  delimiter: string;
  autoConvert: boolean;
  step: 'options' | 'limitWarning';
}

export interface PendingSort {
  mode: 'quick' | 'advanced';
  range: CellRange;
  hasHeader: boolean;
  keys: SortKey[];
  step: 'options' | 'formulaWarning';
  // เรียงจากเมนูตัวกรอง: เรียงเฉพาะข้อมูลใต้หัวตาราง (ไม่มีหน้าต่างตัวเลือก)
  fromFilter?: boolean;
}

export interface SearchState {
  isOpen: boolean;
  query: string;
  options: SearchOptions;
  matches: SearchMatch[];
  current: number; // index ใน matches หรือ -1
}

const closedSearch: SearchState = {
  isOpen: false,
  query: '',
  options: { ...DEFAULT_SEARCH_OPTIONS },
  matches: [],
  current: -1,
};

export const ZOOM_MIN = 50;
export const ZOOM_MAX = 200;
export const ZOOM_STEP = 10;

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
    ...endEditingState,
    evaluator: new Evaluator(wb),
    history: new HistoryManager(100),
    clipboard: null,
    cutRange: null,
    pendingDelete: null,
    pendingDeleteSheet: null,
    search: closedSearch,
    filteredRowsNotice: false,
    pendingSort: null,
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

const endEditingState = {
  isEditing: false,
  editingValue: '',
  editingSheetIndex: null,
  editingCell: null,
};

/**
 * ถ้ากำลังดูชีตอื่นระหว่าง point mode ให้กลับไปชีตและเซลที่กำลังแก้สูตรอยู่
 */
function returnToEditingSheet(get: StoreGet, set: StoreSet) {
  const { editingSheetIndex, editingCell, activeSheetIndex, workbook } = get();
  if (editingSheetIndex === null || editingSheetIndex === activeSheetIndex) return;
  if (!workbook.sheets[editingSheetIndex]) return;

  const views = captureView(get);
  workbook.activeSheet = editingSheetIndex;
  const viewState = viewStateFor(workbook, editingSheetIndex, views);
  if (editingCell) {
    viewState.activeCell = { ...editingCell };
  }
  set({ ...viewState, sheetViews: views });
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
    ...endEditingState,
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
    filterState: captureFilterState(sheet),
  };
  const criteriaBefore = Object.keys(sheet.filter?.columns ?? {}).length;
  if (op === 'insert') {
    if (axis === 'row') opInsertRow(sheet, index, 1);
    else opInsertCol(sheet, index, 1);
  } else {
    if (axis === 'row') opDeleteRow(sheet, index, 1);
    else opDeleteCol(sheet, index, 1);
  }
  const extraCellPatches = adjustOtherSheetsForRowCol(workbook, activeSheetIndex, op, axis, index, 1);
  evaluator.setWorkbook(workbook);
  // ลบคอลัมน์ที่มีเงื่อนไขกรอง: คำนวณแถวที่ซ่อนใหม่จากเงื่อนไขที่เหลือ
  if (sheet.filter && Object.keys(sheet.filter.columns).length !== criteriaBefore) {
    setHiddenRows(sheet, computeHiddenRows(sheet, evaluator, activeSheetIndex));
  }

  const delta = op === 'insert' ? 1 : -1;
  const newRowCount = axis === 'row' ? Math.max(1, rowCount + delta) : rowCount;
  const newColCount = axis === 'col' ? Math.max(1, colCount + delta) : colCount;
  const fullAfter = {
    cells: JSON.parse(JSON.stringify(sheet.cells)),
    merges: JSON.parse(JSON.stringify(sheet.merges)),
    rowCount: newRowCount,
    colCount: newColCount,
    filterState: captureFilterState(sheet),
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

/**
 * ค้นหาใหม่แล้วไปยังผลถัดไป/ก่อนหน้าจากเซลปัจจุบัน (สลับชีตได้เมื่อค้นทุกชีต)
 * inclusive = true ใช้ตอนพิมพ์คำค้น/เปลี่ยนตัวเลือก (เซลปัจจุบันนับเป็นผลแรกได้)
 */
function runSearch(get: StoreGet, set: StoreSet, direction: 1 | -1, inclusive: boolean) {
  const { search, workbook, evaluator, activeSheetIndex, activeCell } = get();
  const matches = findMatches(workbook, evaluator, search.query, search.options, activeSheetIndex);
  if (matches.length === 0) {
    set({ search: { ...search, matches, current: -1 } });
    return;
  }

  const from = { sheetIndex: activeSheetIndex, r: activeCell.r, c: activeCell.c };
  const idx = nextMatchIndex(matches, from, direction, inclusive);
  const m = matches[idx];
  if (m.sheetIndex !== activeSheetIndex) {
    get().switchSheet(m.sheetIndex);
  }
  const { rowCount, colCount } = get();
  if (m.r >= rowCount || m.c >= colCount) {
    set({ rowCount: Math.max(rowCount, m.r + 1), colCount: Math.max(colCount, m.c + 1) });
  }
  get().setActiveCell(m.r, m.c);
  set({
    search: { ...get().search, matches, current: idx },
    scrollToActiveToken: get().scrollToActiveToken + 1,
  });
}

function rowIsHidden(sheet: Sheet, key: string): boolean {
  const p = parseA1(key);
  return !!p && !!sheet.hiddenRows?.includes(p.r);
}

function setHiddenRows(sheet: Sheet, rows: number[]) {
  if (rows.length > 0) sheet.hiddenRows = [...rows].sort((a, b) => a - b);
  else delete sheet.hiddenRows;
}

/**
 * เปลี่ยนตัวกรองของชีตที่เปิดอยู่ + บันทึก Undo (เก็บตัวกรองและแถวที่ซ่อนทั้งก่อนและหลัง)
 * mutate คืน false เมื่อไม่มีอะไรเปลี่ยน
 */
function commitFilterChange(
  get: StoreGet,
  set: StoreSet,
  description: string,
  mutate: (sheet: Sheet) => boolean
) {
  const { isEditing, commitEditing } = get();
  if (isEditing) commitEditing(0, 0);

  const { workbook, activeSheetIndex, history, updateTitle } = get();
  const sheet = workbook.sheets[activeSheetIndex];
  const before = captureFilterState(sheet);
  if (!mutate(sheet)) return;
  const after = captureFilterState(sheet);
  if (JSON.stringify(before) === JSON.stringify(after)) return;

  history.push({
    description,
    sheetIndex: activeSheetIndex,
    patch: { filterBefore: before, filterAfter: after },
  });
  set({ workbook: { ...workbook }, isDirty: true });
  updateTitle();
}

/** แถวที่มองเห็นได้ (ไม่ถูกตัวกรองซ่อน) ในช่วง r1..r2 */
function visibleRowsInRange(sheet: Sheet, r1: number, r2: number): number[] {
  const hidden = new Set(sheet.hiddenRows ?? []);
  const rows: number[] = [];
  for (let r = r1; r <= r2; r++) if (!hidden.has(r)) rows.push(r);
  return rows;
}

/**
 * เตรียมการเรียงลำดับจากเซล/ช่วงที่เลือก: หาช่วง, ตรวจเซลผสาน, เดาหัวตาราง
 * คืน null ถ้าเรียงไม่ได้ (และแสดงข้อความถ้าจำเป็น)
 */
function prepareSort(get: StoreGet, set: StoreSet, mode: 'quick' | 'advanced'): PendingSort | null {
  const { isEditing, commitEditing } = get();
  if (isEditing) commitEditing(0, 0);

  const { workbook, activeSheetIndex, activeCell, selection, evaluator } = get();
  const sheet = workbook.sheets[activeSheetIndex];
  const isMulti = selection.r1 !== selection.r2 || selection.c1 !== selection.c2;
  const range: CellRange = isMulti
    ? { ...selection }
    : detectDataRegion(sheet, activeCell.r, activeCell.c);

  if (!rangeHasData(sheet, range) || range.r1 === range.r2) {
    set({ notice: th.sort.noData });
    return null;
  }
  if (rangeHasMerge(sheet, range)) {
    set({ notice: th.sort.mergedError });
    return null;
  }

  const col = Math.min(range.c2, Math.max(range.c1, activeCell.c));
  return {
    mode,
    range,
    hasHeader: detectHeaderRow(sheet, evaluator, activeSheetIndex, range),
    keys: [{ col, dir: 'asc' }],
    step: 'options',
  };
}

/** เรียงลำดับจริง + บันทึก Undo หนึ่งขั้น */
function performSort(get: StoreGet, set: StoreSet, pending: PendingSort) {
  const { workbook, activeSheetIndex, evaluator, history, updateTitle } = get();
  const sheet = workbook.sheets[activeSheetIndex];
  if (rangeHasMerge(sheet, pending.range)) {
    set({ notice: th.sort.mergedError });
    return;
  }

  const filterBefore = captureFilterState(sheet);
  const order: number[] = [];
  const patches = sortRange(
    sheet,
    evaluator,
    activeSheetIndex,
    pending.range,
    pending.keys,
    pending.hasHeader,
    order
  );
  if (patches.length === 0) return;

  // แถวย้ายที่แล้ว: ปรับแถวที่ซ่อนให้ตรงกับข้อมูล (เรียงรวมแถวที่ซ่อนอยู่ด้วย)
  evaluator.setWorkbook(workbook);
  let filterPatch = {};
  if (sheet.filter && Object.keys(sheet.filter.columns).length > 0) {
    // มีเงื่อนไข: คำนวณจากเงื่อนไขใหม่
    setHiddenRows(sheet, computeHiddenRows(sheet, evaluator, activeSheetIndex));
  } else if (sheet.hiddenRows && sheet.hiddenRows.length > 0) {
    // ไม่รู้เงื่อนไข (เช่น แถวซ่อนจากไฟล์): ให้สถานะซ่อนย้ายตามแถวของมัน
    const dataR1 = pending.hasHeader ? pending.range.r1 + 1 : pending.range.r1;
    const hidden = new Set(sheet.hiddenRows);
    const outside = sheet.hiddenRows.filter((r) => r < dataR1 || r > pending.range.r2);
    const moved = order.flatMap((oldRow, i) => (hidden.has(oldRow) ? [dataR1 + i] : []));
    setHiddenRows(sheet, [...outside, ...moved]);
  }
  if (JSON.stringify(filterBefore) !== JSON.stringify(captureFilterState(sheet))) {
    filterPatch = { filterBefore, filterAfter: captureFilterState(sheet) };
  }

  history.push({
    description: 'Sort range',
    sheetIndex: activeSheetIndex,
    patch: { cellPatches: patches, ...filterPatch },
  });
  const { range } = pending;
  set({
    workbook: { ...workbook },
    selection: { ...range },
    isDirty: true,
  });
  updateTitle();
}

/** path ของไฟล์ CSV: ห้ามใช้เป็นที่บันทึก (บันทึกเป็น .xlsx เสมอ) */
export function isCsvPath(path: string | null | undefined): boolean {
  return !!path && /\.(csv|txt)$/i.test(path);
}

/**
 * นำเข้า CSV จริง:
 * - open: งานใหม่ 1 ชีต ยังไม่มีชื่อไฟล์ (Ctrl+S = บันทึกเป็น .xlsx) ชื่อชีตตามชื่อไฟล์
 * - newSheet: เพิ่มเป็นชีตใหม่ท้ายสมุดงานที่เปิดอยู่ (Undo ได้)
 */
function performCsvImport(get: StoreGet, set: StoreSet, p: PendingCsvImport, rows: string[][]) {
  if (p.mode === 'newSheet') {
    runSheetOp(
      get,
      set,
      'Import CSV',
      (wb) => {
        const name = sheetNameFromFileName(p.fileName, wb.sheets.map((s) => s.name));
        wb.sheets.push(buildSheetFromCsv(rows, name, p.autoConvert));
        return { ok: true, index: wb.sheets.length - 1 };
      },
      { affectedBefore: get().activeSheetIndex, show: 'result' }
    );
    return;
  }

  get().clearRecoverySnapshot();
  const sheet = buildSheetFromCsv(rows, sheetNameFromFileName(p.fileName), p.autoConvert);
  if (p.filePath) get().addRecentFile(p.filePath);
  set({
    ...freshWorkbookState({ sheets: [sheet], activeSheet: 0 }),
    filePath: null, // ห้ามเขียนทับไฟล์ .csv ต้นฉบับ
    fileName: p.fileName.replace(/\.[^.]+$/, '') || th.csv.defaultName,
    isDirty: false,
    recoverySessionId: `recovery-${Date.now()}`,
    unsupportedFeatures: [],
    hiddenSheetNotice: 0,
    hasWarnedUnsupportedOverwrite: false,
  });
  get().updateTitle();
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
  editingSheetIndex: null,
  editingCell: null,
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
  hiddenSheetNotice: 0,
  hasWarnedUnsupportedOverwrite: false,
  sheetViews: { [initialWb.sheets[0].id!]: defaultSheetView() },
  pendingDeleteSheet: null,
  search: closedSearch,
  scrollToActiveToken: 0,
  pendingSort: null,
  notice: null,
  filteredRowsNotice: false,
  pendingCsvImport: null,
  yearMode: 'ce',

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
    const existing = editTextForCell(sheet.cells[key]?.raw || '', sheet.cells[key]?.fmt);
    const val = initialValue !== undefined ? initialValue : existing;

    set({
      isEditing: true,
      editingValue: val,
      editingSheetIndex: activeSheetIndex,
      editingCell: { ...activeCell },
    });
  },

  setEditingValue: (val) => {
    set({ editingValue: val });
  },

  commitEditing: (dr = 0, dc = 0) => {
    const { isEditing, editingValue, workbook, evaluator, history, setActiveCell, updateTitle } = get();

    if (!isEditing) return;

    // เซลที่กำลังแก้อาจอยู่คนละชีตกับที่แสดงอยู่ (point mode ข้ามชีต)
    const editSheetIndex = get().editingSheetIndex ?? get().activeSheetIndex;
    const editCell = get().editingCell ?? get().activeCell;
    returnToEditingSheet(get, set);

    const sheet = workbook.sheets[editSheetIndex];
    const key = cellToA1(editCell.r, editCell.c);
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
    // วันที่ที่พิมพ์ (d/m, d/m/yyyy, dd-mm-yyyy, yyyy-mm-dd; ปี >= 2400 = พ.ศ.) → วันที่จริง (SPEC 5.9 / 5.15)
    const typedDate = convertTypedDate(finalValue);
    if (typedDate) {
      finalValue = typedDate.raw;
    }

    const currentCell = sheet.cells[key];
    const isSame = (currentCell?.raw || '') === finalValue && (!typedDate || currentCell?.fmt === 'date');

    if (!isSame) {
      if (finalValue === '') {
        delete sheet.cells[key];
      } else {
        sheet.cells[key] = {
          raw: finalValue,
          fmt: typedDate ? 'date' : currentCell?.fmt,
          style: currentCell?.style,
        };
      }

      evaluator.invalidate(key, editSheetIndex);

      const action: HistoryAction = {
        description: `Edit ${key}`,
        sheetIndex: editSheetIndex,
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
        workbook: { ...get().workbook },
        ...endEditingState,
        isDirty: true,
      });
      updateTitle();
    } else {
      set({ ...endEditingState });
    }

    if (dr !== 0 || dc !== 0) {
      setActiveCell(editCell.r + dr, editCell.c + dc, false);
    }
  },

  cancelEditing: () => {
    returnToEditingSheet(get, set);
    set({ ...endEditingState });
  },

  clearSelectionContent: () => {
    const { selection, workbook, activeSheetIndex, evaluator, history, updateTitle } = get();
    const sheet = workbook.sheets[activeSheetIndex];
    const patches: any[] = [];

    const hiddenRows = new Set(sheet.hiddenRows ?? []);
    for (let r = selection.r1; r <= selection.r2; r++) {
      if (hiddenRows.has(r)) continue; // ไม่แก้ข้อมูลในแถวที่ตัวกรองซ่อน
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
    // ระหว่างแก้ไข: แทรกข้อความวันที่ตามโหมดปี (ยืนยันแล้วจะแปลงเป็นวันที่จริง)
    const dateStr = formatDate(todaySerial());

    if (isEditing) {
      setEditingValue(editingValue + dateStr);
      return;
    }

    const sheet = workbook.sheets[activeSheetIndex];
    const patches: any[] = [];

    const hiddenRows = new Set(sheet.hiddenRows ?? []);
    for (let r = selection.r1; r <= selection.r2; r++) {
      if (hiddenRows.has(r)) continue; // ไม่แก้ข้อมูลในแถวที่ตัวกรองซ่อน
      for (let c = selection.c1; c <= selection.c2; c++) {
        const key = cellToA1(r, c);
        const cell = sheet.cells[key];
        const before = cell ? { ...cell } : undefined;
        const after: Cell = {
          raw: String(todaySerial()), // วันที่จริง (serial ค.ศ.)
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
      hiddenSheetNotice: 0,
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
      filteredRowsNotice: wb.sheets.some((s) => (s.hiddenRows?.length ?? 0) > 0),
      hiddenSheetNotice: wb.sheets.filter((s) => s.hidden).length,
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
      // งานที่มาจาก CSV: ไม่ผูกกับไฟล์ .csv เดิม (บันทึกเป็น .xlsx เสมอ)
      filePath: isCsvPath(origPath) ? null : origPath,
      fileName,
      isDirty: true,
      unsupportedFeatures: recoveredWb.unsupportedFeatures || [],
      hiddenSheetNotice: recoveredWb.sheets.filter((s) => s.hidden).length,
      hasWarnedUnsupportedOverwrite: false,
    });
    get().updateTitle();
  },

  setHasWarnedUnsupportedOverwrite: (warned: boolean) => {
    set({ hasWarnedUnsupportedOverwrite: warned });
  },

  dismissUnsupportedBanner: () => {
    set({ unsupportedFeatures: [], hiddenSheetNotice: 0 });
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
      hiddenSheetNotice: 0,
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

    // ระหว่างพิมพ์สูตร: สลับไปดูชีตอื่นได้โดยไม่ยืนยันการแก้ไข (point mode ข้ามชีต)
    // ข้อความธรรมดา: ยืนยันการแก้ไขก่อนสลับ
    const isFormulaEditing = isEditing && get().editingValue.trimStart().startsWith('=');
    if (isEditing && !isFormulaEditing) {
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

  // --- Phase 4C: ค้นหา ---

  openSearch: () => {
    const { search, isEditing, commitEditing } = get();
    if (isEditing) commitEditing(0, 0);
    set({ search: { ...search, isOpen: true } });
  },

  closeSearch: () => {
    // เซลที่ค้นพบล่าสุดยังถูกเลือกค้างไว้ (activeCell ไม่เปลี่ยน)
    set({ search: { ...closedSearch, options: get().search.options } });
  },

  setSearchQuery: (query: string) => {
    const { search } = get();
    set({ search: { ...search, query } });
    runSearch(get, set, 1, true);
  },

  setSearchOption: (key: keyof SearchOptions, value: boolean) => {
    const { search } = get();
    set({ search: { ...search, options: { ...search.options, [key]: value } } });
    runSearch(get, set, 1, true);
  },

  searchNext: () => {
    runSearch(get, set, 1, false);
  },

  searchPrev: () => {
    runSearch(get, set, -1, false);
  },

  // --- Phase 4C: เรียงลำดับ ---

  dismissNotice: () => set({ notice: null }),

  requestSort: (dir: SortDirection) => {
    const pending = prepareSort(get, set, 'quick');
    if (!pending) return;
    pending.keys = [{ col: pending.keys[0].col, dir }];
    set({ pendingSort: pending });
  },

  openAdvancedSort: () => {
    const pending = prepareSort(get, set, 'advanced');
    if (pending) set({ pendingSort: pending });
  },

  updatePendingSort: (patch) => {
    const { pendingSort } = get();
    if (!pendingSort) return;
    set({ pendingSort: { ...pendingSort, ...patch } });
  },

  confirmSort: () => {
    const { pendingSort, workbook, activeSheetIndex } = get();
    if (!pendingSort) return;
    const sheet = workbook.sheets[activeSheetIndex];

    if (
      pendingSort.step === 'options' &&
      formulasReferenceOtherRows(sheet, pendingSort.range, pendingSort.hasHeader)
    ) {
      set({ pendingSort: { ...pendingSort, step: 'formulaWarning' } });
      return;
    }

    set({ pendingSort: null });
    performSort(get, set, pendingSort);
  },

  cancelSort: () => set({ pendingSort: null }),

  // --- Phase 4C: ตัวกรอง ---

  toggleFilter: () => {
    const { workbook, activeSheetIndex, selection, activeCell } = get();
    const sheet = workbook.sheets[activeSheetIndex];

    if (sheet.filter) {
      // ปิดตัวกรอง: แสดงทุกแถวกลับมา
      commitFilterChange(get, set, 'Remove filter', (s) => {
        delete s.filter;
        delete s.hiddenRows;
        return true;
      });
      return;
    }

    const isMulti = selection.r1 !== selection.r2 || selection.c1 !== selection.c2;
    const range: CellRange = isMulti
      ? { ...selection }
      : detectDataRegion(sheet, activeCell.r, activeCell.c);
    if (!rangeHasData(sheet, range)) {
      set({ notice: th.filter.noData });
      return;
    }

    commitFilterChange(get, set, 'Create filter', (s) => {
      s.filter = createFilter(range);
      return true;
    });
  },

  setColumnFilter: (col: number, cf: ColumnFilter | null) => {
    const { evaluator, activeSheetIndex } = get();
    commitFilterChange(get, set, 'Change filter', (s) => {
      if (!s.filter) return false;
      const columns = { ...s.filter.columns };
      if (cf) columns[col] = cf;
      else delete columns[col];
      s.filter = { ...s.filter, columns };
      setHiddenRows(s, computeHiddenRows(s, evaluator, activeSheetIndex));
      return true;
    });
  },

  reapplyFilter: () => {
    const { evaluator, activeSheetIndex } = get();
    commitFilterChange(get, set, 'Reapply filter', (s) => {
      if (!s.filter) return false;
      // ไม่มีเงื่อนไขที่รู้จัก (เช่น ตัวกรองจากไฟล์) → ไม่ไปแตะแถวที่ซ่อนอยู่
      if (Object.keys(s.filter.columns).length === 0) return false;
      setHiddenRows(s, computeHiddenRows(s, evaluator, activeSheetIndex));
      return true;
    });
  },

  sortFilterColumn: (col: number, dir: SortDirection) => {
    const { workbook, activeSheetIndex } = get();
    const filter = workbook.sheets[activeSheetIndex].filter;
    if (!filter) return;
    const range = { r1: filter.r1, c1: filter.c1, r2: filter.r2, c2: filter.c2 };
    if (rangeHasMerge(workbook.sheets[activeSheetIndex], range)) {
      set({ notice: th.sort.mergedError });
      return;
    }
    set({
      pendingSort: {
        mode: 'quick',
        range,
        hasHeader: true,
        keys: [{ col, dir }],
        step: 'options',
        fromFilter: true,
      },
    });
    // เรียงจากเมนูตัวกรองไม่มีหน้าต่างตัวเลือก (หัวตารางคือแถวแรกของตัวกรองเสมอ)
    get().confirmSort();
  },

  showAllFilteredRows: () => {
    // จากแถบแจ้งไฟล์ที่มีแถวซ่อน: แสดงทุกแถวในทุกชีต (Undo ได้)
    const { workbook, history, evaluator, activeSheetIndex, updateTitle } = get();
    if (!workbook.sheets.some((s) => s.hiddenRows && s.hiddenRows.length > 0)) {
      set({ filteredRowsNotice: false });
      return;
    }
    const before = snapshotWorkbook(workbook);
    before.activeSheet = activeSheetIndex;
    for (const s of workbook.sheets) delete s.hiddenRows;
    const after = snapshotWorkbook(workbook);
    after.activeSheet = activeSheetIndex;
    history.push({ description: 'Show all rows', workbookBefore: before, workbookAfter: after });
    const nextWb = { ...workbook };
    evaluator.setWorkbook(nextWb);
    set({ workbook: nextWb, filteredRowsNotice: false, isDirty: true });
    updateTitle();
  },

  dismissFilteredRowsNotice: () => set({ filteredRowsNotice: false }),

  // --- Phase 4C: ซูม ---

  setZoom: (percent: number) => {
    const { workbook, activeSheetIndex } = get();
    const sheet = workbook.sheets[activeSheetIndex];
    const clamped = Math.max(ZOOM_MIN, Math.min(ZOOM_MAX, Math.round(percent / ZOOM_STEP) * ZOOM_STEP));
    if ((sheet.zoom ?? 100) === clamped) return;
    if (clamped === 100) delete sheet.zoom;
    else sheet.zoom = clamped;
    // ระดับซูมเป็นสถานะการมองเห็น: ไม่ทำให้ไฟล์เป็น "ยังไม่ได้บันทึก" แต่จะถูกเขียนลงไฟล์ตอนบันทึกครั้งถัดไป
    set({ workbook: { ...workbook } });
  },

  zoomIn: () => {
    const { workbook, activeSheetIndex, setZoom } = get();
    setZoom((workbook.sheets[activeSheetIndex].zoom ?? 100) + ZOOM_STEP);
  },

  zoomOut: () => {
    const { workbook, activeSheetIndex, setZoom } = get();
    setZoom((workbook.sheets[activeSheetIndex].zoom ?? 100) - ZOOM_STEP);
  },

  resetZoom: () => get().setZoom(100),

  // --- Phase 4D: CSV ---

  beginCsvImport: (bytes: Uint8Array, filePath: string | null, mode: 'open' | 'newSheet') => {
    if (bytes.length === 0) {
      set({ notice: th.csv.emptyFile });
      return;
    }
    const encoding = detectEncoding(bytes);
    if (looksBinary(bytes, encoding)) {
      set({ notice: th.csv.notText });
      return;
    }
    const text = decodeBytes(bytes, encoding);
    if (text.trim() === '') {
      set({ notice: th.csv.emptyFile });
      return;
    }
    const fileName = (filePath ?? th.csv.defaultName).split(/[\\/]/).pop()!;
    set({
      pendingCsvImport: {
        mode,
        filePath,
        fileName,
        bytes,
        encoding,
        delimiter: detectDelimiter(text),
        autoConvert: true,
        step: 'options',
      },
    });
  },

  updateCsvImport: (patch) => {
    const { pendingCsvImport } = get();
    if (!pendingCsvImport) return;
    set({ pendingCsvImport: { ...pendingCsvImport, ...patch, step: 'options' } });
  },

  confirmCsvImport: () => {
    const { pendingCsvImport: p } = get();
    if (!p) return;
    const rows = parseCsv(decodeBytes(p.bytes, p.encoding), p.delimiter);
    if (rows.length === 0 || rows.every((r) => r.every((v) => v === ''))) {
      set({ pendingCsvImport: null, notice: th.csv.emptyFile });
      return;
    }
    // เกินขีดจำกัด: ต้องให้ผู้ใช้ยืนยันก่อน ห้ามตัดข้อมูลเงียบๆ
    if (p.step === 'options' && csvExceedsLimits(rows)) {
      set({ pendingCsvImport: { ...p, step: 'limitWarning' } });
      return;
    }
    set({ pendingCsvImport: null });
    performCsvImport(get, set, p, rows);
  },

  cancelCsvImport: () => set({ pendingCsvImport: null }),

  // --- Phase 5A: ปี พ.ศ. ---

  setPrintSettings: (patch) => {
    const { workbook } = get();
    // ค่าการพิมพ์เป็นค่าของไฟล์ แต่ไม่ทำให้ขึ้น "ยังไม่ได้บันทึก" (เหมือนระดับซูม)
    const next = { ...workbook, printSettings: { ...DEFAULT_PRINT_SETTINGS, ...workbook.printSettings, ...patch } };
    set({ workbook: next });
  },

  setYearMode: (mode: YearMode, persist = true) => {
    setDisplayYearMode(mode);
    // สร้าง object ใหม่ให้ตารางแสดงผลใหม่ทันที (ข้อมูลไม่เปลี่ยน)
    set({ yearMode: mode, workbook: { ...get().workbook } });
    if (!persist) return;
    const electronAPI = typeof window !== 'undefined' ? (window as any).electronAPI : undefined;
    if (electronAPI?.setPreferences) {
      electronAPI.setPreferences({ yearMode: mode }).catch(() => {});
    } else {
      try {
        localStorage.setItem('sheeteasy_year_mode', mode);
      } catch {}
    }
  },

  getCsvExport: () => {
    const { workbook, activeSheetIndex, evaluator, fileName } = get();
    const sheet = workbook.sheets[activeSheetIndex];
    const text = writeCsv(sheetToCsvRows(sheet, evaluator, activeSheetIndex));
    // ชื่อไฟล์เริ่มต้น = ชื่อไฟล์ปัจจุบัน + ชื่อชีต (ตัดอักขระที่ใช้ในชื่อไฟล์ไม่ได้)
    const defaultName = `${fileName} - ${sheet.name}.csv`.replace(/[<>:"/\\|?*]/g, '');
    const hasDates = Object.values(sheet.cells).some((c) => c.fmt === 'date' && c.raw !== '');
    return { bytes: encodeUtf8WithBom(text), defaultName, hasFilter: !!sheet.filter, hasDates };
  },

  getEditText: (r: number, c: number) => {
    const { workbook, activeSheetIndex } = get();
    const cell = workbook.sheets[activeSheetIndex]?.cells[cellToA1(r, c)];
    return cell ? editTextForCell(cell.raw, cell.fmt) : '';
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
    const sourceKeys: Record<string, string> = {};

    // คัดลอกเฉพาะแถวที่มองเห็น (แถวที่ตัวกรองซ่อนไม่ถูกคัดลอก) แล้วต่อแถวให้ชิดกัน
    const rows = visibleRowsInRange(sheet, selection.r1, selection.r2);
    rows.forEach((srcR, i) => {
      const dstR = selection.r1 + i;
      for (let c = selection.c1; c <= selection.c2; c++) {
        const key = cellToA1(srcR, c);
        const cell = sheet.cells[key];
        if (!cell) continue;
        const copy = { ...cell };
        if (copy.raw && copy.raw.startsWith('=') && dstR !== srcR) {
          copy.raw = shiftFormula(copy.raw, dstR - srcR, 0);
        }
        const compactKey = cellToA1(dstR, c);
        copiedCells[compactKey] = copy;
        sourceKeys[compactKey] = key;
      }
    });

    const skipRows = new Set(sheet.hiddenRows ?? []);
    const tsv = rangeToTsv(sheet, selection, getDisplayValue, skipRows);
    const html = rangeToHtml(sheet, selection, getDisplayValue, skipRows);

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
        r2: selection.r1 + Math.max(0, rows.length - 1),
        c2: selection.c2,
        isCut: false,
        sourceKeys,
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
    // แถวที่ตัวกรองซ่อนอยู่จะไม่ถูกเขียนทับ: วางลงเฉพาะแถวที่มองเห็น ตามลำดับ
    const hidden = new Set(sheet.hiddenRows ?? []);
    const visibleRowFrom = (start: number, offset: number): number => {
      let r = start;
      let seen = -1;
      while (true) {
        if (!hidden.has(r)) seen++;
        if (seen === offset) return r;
        r++;
      }
    };

    const writeCell = (destR: number, destC: number, cell: Cell | undefined, rawVal: string) => {
      const destKey = cellToA1(destR, destC);
      const prev = sheet.cells[destKey] ? { ...sheet.cells[destKey] } : undefined;
      sheet.cells[destKey] = {
        raw: rawVal,
        fmt: valuesOnly ? prev?.fmt : cell?.fmt,
        style: valuesOnly
          ? prev?.style
            ? { ...prev.style }
            : undefined
          : cell?.style
            ? { ...cell.style }
            : undefined,
      };
      evaluator.invalidate(destKey, activeSheetIndex);
      patches.push({ key: destKey, before: prev, after: { ...sheet.cells[destKey] } });
    };

    if (clipboard && Object.keys(clipboard.cells).length > 0) {
      // ชีตต้นทางของคลิปบอร์ด (อาจเป็นชีตอื่น) ใช้คำนวณค่าตอนวางเฉพาะค่า และล้างเซลตอนตัด
      const srcIndex = clipboard.sourceSheetId
        ? workbook.sheets.findIndex((s) => s.id === clipboard.sourceSheetId) // -1 = ชีตต้นทางถูกลบไปแล้ว
        : activeSheetIndex;
      const valueSheetIndex = srcIndex >= 0 ? srcIndex : activeSheetIndex;
      const originalKey = (key: string) => clipboard.sourceKeys?.[key] ?? key;
      const dc = activeCell.c - clipboard.c1;

      // Check if copying single cell into multi-cell target selection
      const isSingleCopied = clipboard.r1 === clipboard.r2 && clipboard.c1 === clipboard.c2;
      const isTargetMulti = selection.r1 !== selection.r2 || selection.c1 !== selection.c2;

      if (isSingleCopied && isTargetMulti) {
        const srcKey = cellToA1(clipboard.r1, clipboard.c1);
        const srcCell = clipboard.cells[srcKey];
        for (let r = selection.r1; r <= selection.r2; r++) {
          if (hidden.has(r)) continue;
          for (let c = selection.c1; c <= selection.c2; c++) {
            let rawVal = srcCell ? srcCell.raw : '';
            if (rawVal.startsWith('=')) {
              rawVal = valuesOnly
                ? String(evaluator.getCellValue(originalKey(srcKey), valueSheetIndex))
                : shiftFormula(rawVal, r - clipboard.r1, c - clipboard.c1);
            }
            writeCell(r, c, srcCell, rawVal);
          }
        }
      } else {
        // Normal paste: แถวที่ i ของคลิปบอร์ดไปที่แถวที่มองเห็นลำดับที่ i นับจากเซลปัจจุบัน
        for (const [srcKey, cell] of Object.entries(clipboard.cells)) {
          const p = parseA1(srcKey);
          if (!p) continue;

          const destR = visibleRowFrom(activeCell.r, p.r - clipboard.r1);
          const destC = p.c + dc;

          let rawVal = cell.raw;
          if (rawVal.startsWith('=')) {
            rawVal = valuesOnly
              ? String(evaluator.getCellValue(originalKey(srcKey), valueSheetIndex))
              : shiftFormula(rawVal, destR - p.r, dc);
          }
          writeCell(destR, destC, cell, rawVal);
        }
      }

      // If cut operation, clear original cut cells (ในชีตต้นทาง ซึ่งอาจเป็นชีตอื่น)
      const sourcePatches: SheetCellPatch[] = [];
      if (clipboard.isCut && srcIndex >= 0) {
        const srcSheet = workbook.sheets[srcIndex];
        const sameSheet = srcIndex === activeSheetIndex;
        for (const key of Object.keys(clipboard.cells)) {
          const srcKey = originalKey(key);
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
          const destR = visibleRowFrom(activeCell.r, i);
          for (let j = 0; j < grid[i].length; j++) {
            const destKey = cellToA1(destR, activeCell.c + j);
            const prev = sheet.cells[destKey] ? { ...sheet.cells[destKey] } : undefined;
            // วันที่จากภายนอก (รวมปี พ.ศ. >= 2400) แปลงเป็นวันที่จริงเหมือนการพิมพ์
            const pastedDate = convertTypedDate(grid[i][j]);
            sheet.cells[destKey] = {
              raw: pastedDate ? pastedDate.raw : grid[i][j],
              fmt: pastedDate ? 'date' : prev?.fmt,
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
      if (rowIsHidden(sheet, key)) continue; // ไม่เติมลงแถวที่ตัวกรองซ่อน
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
      if (rowIsHidden(sheet, key)) continue; // ไม่เติมลงแถวที่ตัวกรองซ่อน
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
      if (rowIsHidden(sheet, key)) continue; // ไม่เติมลงแถวที่ตัวกรองซ่อน
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
