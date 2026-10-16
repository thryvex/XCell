import React, { useRef, useEffect, useLayoutEffect, useState, useMemo, useCallback } from 'react';
import { useSheetStore, SelectionState } from './store';
import { Cell } from './Cell';
import { ContextMenu, ContextMenuTarget } from './ContextMenu';
import { FilterMenu } from './FilterMenu';
import { colToStr, cellToA1, parseRange } from '../core/address';
import { tokenize } from '../core/tokenizer';
import { sameSheetName, formatSheetReference } from '../core/sheetName';
import { MergeRange, DEFAULT_COL_WIDTH_PX, DEFAULT_ROW_HEIGHT_PX } from '../core/model';
import { th } from '../i18n/th';
import { formatDate, todaySerial } from '../core/numfmt';
import {
  NavGrid,
  Direction,
  stepCell,
  jumpToDataEdge,
  moveVisibleRows,
  homePos,
  endPos,
  sheetStartPos,
  sheetEndPos,
} from '../core/navigation';

const DEFAULT_COL_WIDTH = DEFAULT_COL_WIDTH_PX;
const DEFAULT_ROW_HEIGHT = DEFAULT_ROW_HEIGHT_PX;
const HEADER_ROW_HEIGHT = 28;
const HEADER_COL_WIDTH = 45;

// ชื่อ token สี reference ตอนแก้สูตร (ค่าสีจริงอยู่ใน theme.css: --ref-1 … --ref-6)
const HIGHLIGHT_COLORS = ['ref-1', 'ref-2', 'ref-3', 'ref-4', 'ref-5', 'ref-6'];

export const Grid: React.FC = () => {
  const containerRef = useRef<HTMLDivElement>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  const {
    workbook,
    activeSheetIndex,
    activeCell,
    selection,
    isEditing,
    editingValue,
    editingSheetIndex,
    rowCount,
    colCount,
    cutRange,
    clipboard,
    nextSheet,
    prevSheet,
    rememberScroll,
    getActiveSheetView,
    setActiveCell,
    setSelection,
    startEditing,
    setEditingValue,
    commitEditing,
    cancelEditing,
    clearSelectionContent,
    clearSelectionFormat,
    insertCurrentDate,
    selectAll,
    copySelection,
    cutSelection,
    paste,
    insertRowAt,
    deleteRowAt,
    insertColAt,
    deleteColAt,
    setColWidth,
    setRowHeight,
    freezeTo,
    applyFillRange,
    addRowsAtBottom,
    fillDownAction,
    fillRightAction,
    cycleCurrentReference,
    toggleBold,
    toggleItalic,
    toggleUnderline,
    undo,
    redo,
    mergeCells,
    unmergeCells,
    getDisplayValue,
    getRawValue,
    getComputedValue,
    search,
    openSearch,
    searchNext,
    searchPrev,
    scrollToActiveToken,
    scrollTarget,
    extendSelectionTo,
    getSelectionFocus,
    scrollToCell,
    requestSort,
    openAdvancedSort,
    toggleFilter,
    zoomIn,
    zoomOut,
    resetZoom,
  } = useSheetStore();
  // พื้นตาราง "ขาวเสมอ" (5C): เปลี่ยนเฉพาะ token ของพื้นที่ตาราง (theme.css [data-grid-white])
  const gridWhite = useSheetStore((s) => s.gridWhite);

  const sheet = workbook.sheets[activeSheetIndex];
  const sheetId = sheet.id;

  // แถวที่ซ่อนด้วยตัวกรอง (4C)
  const hiddenSet = useMemo(() => new Set(sheet.hiddenRows ?? []), [sheet.hiddenRows]);
  // แถวถัดไป/ก่อนหน้าที่มองเห็น (ข้ามแถวที่ซ่อน) สำหรับลูกศร/Enter
  const stepRow = useCallback(
    (from: number, dir: 1 | -1) => {
      let r = from + dir;
      while (r >= 0 && r < rowCount && hiddenSet.has(r)) r += dir;
      return r < 0 || r >= rowCount ? from : r;
    },
    [hiddenSet, rowCount]
  );
  const [filterMenu, setFilterMenu] = useState<{ col: number; x: number; y: number } | null>(null);

  // ซูมต่อชีต (4C)
  const zoomFactor = (sheet.zoom ?? 100) / 100;

  // Ctrl+ล้อเมาส์ = ซูม (ต้องเป็น listener แบบ non-passive จึงกัน preventDefault ได้)
  useEffect(() => {
    const el = containerRef.current;
    if (!el) return;
    const onWheel = (e: WheelEvent) => {
      if (!e.ctrlKey && !e.metaKey) return;
      e.preventDefault();
      if (e.deltaY < 0) zoomIn();
      else if (e.deltaY > 0) zoomOut();
    };
    el.addEventListener('wheel', onWheel, { passive: false });
    return () => el.removeEventListener('wheel', onWheel);
  }, [zoomIn, zoomOut]);
  // ชีตของเซลที่กำลังแก้สูตร (point mode ข้ามชีต: ชีตที่แสดงอาจไม่ใช่ชีตนี้)
  const editingSheet = editingSheetIndex !== null ? workbook.sheets[editingSheetIndex] : undefined;
  const editingSheetName = editingSheet?.name ?? sheet.name;
  const isOnEditingSheet = editingSheetIndex === null || editingSheetIndex === activeSheetIndex;

  // แสดงกรอบ "ตัด" เฉพาะบนชีตต้นทาง
  const visibleCutRange =
    cutRange && (!clipboard?.sourceSheetId || clipboard.sourceSheetId === sheetId) ? cutRange : null;

  // ผลค้นหาบนชีตนี้: ทุกผล (สีอ่อน) และผลปัจจุบัน (สีเด่น)
  const searchCells = useMemo(() => {
    const map = new Map<string, 'match' | 'current'>();
    if (!search.isOpen) return map;
    search.matches.forEach((m, i) => {
      if (m.sheetIndex !== activeSheetIndex) return;
      map.set(cellToA1(m.r, m.c), i === search.current ? 'current' : 'match');
    });
    return map;
  }, [search.isOpen, search.matches, search.current, activeSheetIndex]);

  // เลื่อนตารางให้เห็นเซลที่เลือก (เมื่อไปยังผลค้นหา)
  useEffect(() => {
    const el = containerRef.current;
    if (!el || scrollToActiveToken === 0) return;
    const target = scrollTarget ?? activeCell;
    const hidden = new Set(sheet.hiddenRows ?? []);
    const rowH = (r: number) => (hidden.has(r) ? 0 : sheet.rowHeights[r] || DEFAULT_ROW_HEIGHT);
    const colW = (c: number) => sheet.colWidths[c] || DEFAULT_COL_WIDTH;
    let top = HEADER_ROW_HEIGHT;
    for (let r = 0; r < target.r; r++) top += rowH(r);
    let left = HEADER_COL_WIDTH;
    for (let c = 0; c < target.c; c++) left += colW(c);
    // พื้นที่ที่แถว/คอลัมน์ตรึงบังไว้ด้านบน/ซ้าย
    let frozenH = 0;
    for (let r = 0; r < sheet.frozenRows; r++) frozenH += rowH(r);
    let frozenW = 0;
    for (let c = 0; c < sheet.frozenCols; c++) frozenW += colW(c);
    const zoom = (sheet.zoom ?? 100) / 100;
    const height = rowH(target.r) * zoom;
    const width = colW(target.c) * zoom;
    top *= zoom;
    left *= zoom;
    const viewTop = el.scrollTop + (HEADER_ROW_HEIGHT + frozenH) * zoom;
    const viewLeft = el.scrollLeft + (HEADER_COL_WIDTH + frozenW) * zoom;
    if (target.r >= sheet.frozenRows && (top < viewTop || top + height > el.scrollTop + el.clientHeight)) {
      el.scrollTop = Math.max(0, top - el.clientHeight / 2);
    }
    if (target.c >= sheet.frozenCols && (left < viewLeft || left + width > el.scrollLeft + el.clientWidth)) {
      el.scrollLeft = Math.max(0, left - el.clientWidth / 2);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [scrollToActiveToken]);

  // คืนตำแหน่งเลื่อนของชีตเมื่อสลับชีต (สถานะแยกต่อชีต)
  useLayoutEffect(() => {
    const el = containerRef.current;
    if (!el) return;
    const view = getActiveSheetView();
    el.scrollTop = view.scrollTop;
    el.scrollLeft = view.scrollLeft;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [sheetId]);

  // Mouse states
  const [isMouseDown, setIsMouseDown] = useState(false);
  const [isFillDragging, setIsFillDragging] = useState(false);
  const [fillTarget, setFillTarget] = useState<SelectionState | null>(null);

  // Resize states
  const [resizingCol, setResizingCol] = useState<{ colIdx: number; startX: number; startW: number } | null>(null);
  const [resizingRow, setResizingRow] = useState<{ rowIdx: number; startY: number; startH: number } | null>(null);

  // Context Menu state
  const [contextMenuPos, setContextMenuPos] = useState<{ x: number; y: number; targetType: ContextMenuTarget } | null>(null);

  // Point mode drag state
  const [pointModeAnchor, setPointModeAnchor] = useState<{ r: number; c: number } | null>(null);
  const [pointModeBaseFormula, setPointModeBaseFormula] = useState<string>('');
  const [pointModePrefix, setPointModePrefix] = useState<string>('');

  // Precompute merge maps
  const { anchorMap, coveredMap } = useMemo(() => {
    const anchors: Record<string, MergeRange> = {};
    const covered: Set<string> = new Set();

    for (const m of sheet.merges) {
      anchors[`${m.r1},${m.c1}`] = m;
      for (let r = m.r1; r <= m.r2; r++) {
        for (let c = m.c1; c <= m.c2; c++) {
          if (r !== m.r1 || c !== m.c1) {
            covered.add(`${r},${c}`);
          }
        }
      }
    }

    return { anchorMap: anchors, coveredMap: covered };
  }, [sheet.merges]);

  // Formula Reference Highlights (colored borders for referenced cells)
  const refHighlights = useMemo(() => {
    const map: Record<string, string> = {};
    if (!isEditing || !editingValue.startsWith('=')) return map;

    try {
      const tokens = tokenize(editingValue);
      let colorIdx = 0;

      for (const t of tokens) {
        if (t.type !== 'CELL_REF' && t.type !== 'RANGE') continue;
        const color = HIGHLIGHT_COLORS[colorIdx % HIGHLIGHT_COLORS.length];
        colorIdx++;

        // ไฮไลต์เฉพาะ reference ที่ชี้ชีตที่แสดงอยู่ (ไม่มีชื่อชีต = ชีตของเซลที่กำลังแก้)
        const refSheet = t.sheet ?? editingSheetName;
        if (!sameSheetName(refSheet, sheet.name)) continue;

        if (t.type === 'CELL_REF') {
          map[t.value.toUpperCase()] = color;
        } else {
          const rng = parseRange(t.value);
          if (rng) {
            for (let r = rng.r1; r <= rng.r2; r++) {
              for (let c = rng.c1; c <= rng.c2; c++) {
                map[cellToA1(r, c)] = color;
              }
            }
          }
        }
      }
    } catch {}

    return map;
  }, [isEditing, editingValue, editingSheetName, sheet.name]);

  // Focus input when editing starts
  useEffect(() => {
    if (isEditing && isOnEditingSheet && inputRef.current) {
      inputRef.current.focus();
    }
  }, [isEditing, isOnEditingSheet]);

  // Check if cursor in editing formula is ready for Point Mode reference insertion
  const isPointModeReady = useCallback(() => {
    if (!isEditing || !editingValue.startsWith('=')) return false;
    const trimmed = editingValue.trimEnd();
    const lastChar = trimmed[trimmed.length - 1];
    return '=+-*/^%&,(<>'.includes(lastChar);
  }, [isEditing, editingValue]);

  // ตำแหน่งปลายทางของปุ่มเลื่อน (null = ไม่ใช่ปุ่มเลื่อน)
  const navigationTarget = (
    e: KeyboardEvent,
    isCtrl: boolean
  ): { r: number; c: number; scrollBy?: { top: number; left: number } } | null => {
    const nav: NavGrid = {
      rowCount,
      colCount,
      hiddenRows: hiddenSet,
      merges: sheet.merges,
      hasValue: (r, c) => Boolean(sheet.cells[cellToA1(r, c)]?.raw),
    };
    const from = e.shiftKey ? getSelectionFocus() : activeCell;
    const dirs: Record<string, Direction> = { ArrowUp: 'up', ArrowDown: 'down', ArrowLeft: 'left', ArrowRight: 'right' };
    const dir = dirs[e.key];
    if (dir && !e.altKey) {
      if (isCtrl) return jumpToDataEdge(nav, from, dir);
      return stepCell(nav, from, dir) ?? from;
    }
    if (e.key === 'Home' && !e.altKey) return isCtrl ? sheetStartPos(nav) : homePos(nav, from);
    if (e.key === 'End' && !e.altKey) return isCtrl ? sheetEndPos(nav) : endPos(nav, from);
    if ((e.key === 'PageDown' || e.key === 'PageUp') && !isCtrl) {
      const step: 1 | -1 = e.key === 'PageDown' ? 1 : -1;
      const el = containerRef.current;
      const zoom = zoomFactor;
      if (e.altKey) {
        // Alt+PageDown / Alt+PageUp: เลื่อนขวา/ซ้ายทีละความกว้างที่เห็น (ไม่นับคอลัมน์ตรึง)
        let frozenW = 0;
        for (let c = 0; c < sheet.frozenCols; c++) frozenW += sheet.colWidths[c] || DEFAULT_COL_WIDTH;
        const avail = (el?.clientWidth ?? 800) - (HEADER_COL_WIDTH + frozenW) * zoom;
        let c = from.c;
        let used = 0;
        let moved = 0;
        while (c + step >= 0 && c + step < colCount) {
          const w = (sheet.colWidths[c + step] || DEFAULT_COL_WIDTH) * zoom;
          if (moved > 0 && used + w > avail) break;
          used += w;
          c += step;
          moved++;
        }
        return { r: from.r, c, scrollBy: { top: 0, left: step * used } };
      }
      // PageDown / PageUp: เลื่อนทีละจำนวนแถวที่เห็นบนจอ (ไม่นับแถวตรึง) และเลื่อนหน้าจอเท่ากัน
      let frozenH = 0;
      for (let r = 0; r < sheet.frozenRows; r++) {
        if (!hiddenSet.has(r)) frozenH += sheet.rowHeights[r] || DEFAULT_ROW_HEIGHT;
      }
      const avail = (el?.clientHeight ?? 600) - (HEADER_ROW_HEIGHT + frozenH) * zoom;
      let count = 0;
      let used = 0;
      let r = from.r;
      for (;;) {
        const n = moveVisibleRows(nav, r, 1, step);
        if (n === r) break;
        const h = (sheet.rowHeights[n] || DEFAULT_ROW_HEIGHT) * zoom;
        if (count > 0 && used + h > avail) break;
        used += h;
        r = n;
        count++;
      }
      return { r, c: from.c, scrollBy: { top: step * used, left: 0 } };
    }
    return null;
  };

  // Global keyboard shortcuts in grid
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      // ซูม: Ctrl+= / Ctrl+- / Ctrl+0 (ใช้ได้ตลอด)
      if ((e.ctrlKey || e.metaKey) && !e.altKey) {
        if (e.key === '=' || e.key === '+') {
          e.preventDefault();
          zoomIn();
          return;
        }
        if (e.key === '-' || e.key === '_') {
          e.preventDefault();
          zoomOut();
          return;
        }
        if (e.key === '0') {
          e.preventDefault();
          resetZoom();
          return;
        }
      }

      // Ctrl+F เปิดแถบค้นหา, F3 / Shift+F3 ผลถัดไป/ก่อนหน้า
      if ((e.ctrlKey || e.metaKey) && !e.shiftKey && e.key.toLowerCase() === 'f') {
        e.preventDefault();
        openSearch();
        return;
      }
      if (e.key === 'F3' && !isEditing) {
        e.preventDefault();
        if (!search.isOpen) openSearch();
        else if (e.shiftKey) searchPrev();
        else searchNext();
        return;
      }

      // F4 key works both inside editor and grid
      if (e.key === 'F4') {
        e.preventDefault();
        cycleCurrentReference();
        return;
      }

      if (isEditing) return;

      // ไม่แย่งคีย์จากช่องพิมพ์อื่น (เช่น ช่องเปลี่ยนชื่อแท็บชีต, แถบสูตร)
      const target = e.target as HTMLElement | null;
      if (
        target &&
        target !== inputRef.current &&
        (target.tagName === 'INPUT' || target.tagName === 'TEXTAREA' || target.isContentEditable)
      ) {
        return;
      }

      const isCtrl = e.ctrlKey || e.metaKey;

      // Ctrl+PageUp / Ctrl+PageDown สลับชีต
      if (isCtrl && e.key === 'PageUp') {
        e.preventDefault();
        prevSheet();
        return;
      }
      if (isCtrl && e.key === 'PageDown') {
        e.preventDefault();
        nextSheet();
        return;
      }

      if (isCtrl && e.key.toLowerCase() === 'z') {
        e.preventDefault();
        if (e.shiftKey) redo();
        else undo();
        return;
      }

      if (isCtrl && e.key.toLowerCase() === 'y') {
        e.preventDefault();
        redo();
        return;
      }

      if (isCtrl && e.key.toLowerCase() === 'c') {
        e.preventDefault();
        copySelection();
        return;
      }

      if (isCtrl && e.key.toLowerCase() === 'x') {
        e.preventDefault();
        cutSelection();
        return;
      }

      if (isCtrl && e.key.toLowerCase() === 'v') {
        e.preventDefault();
        paste(e.shiftKey);
        return;
      }

      if (isCtrl && e.key.toLowerCase() === 'd') {
        e.preventDefault();
        fillDownAction();
        return;
      }

      if (isCtrl && e.key.toLowerCase() === 'r') {
        e.preventDefault();
        fillRightAction();
        return;
      }

      if (isCtrl && e.key.toLowerCase() === 'b') {
        e.preventDefault();
        toggleBold();
        return;
      }

      if (isCtrl && e.key.toLowerCase() === 'i') {
        e.preventDefault();
        toggleItalic();
        return;
      }

      if (isCtrl && e.key.toLowerCase() === 'u') {
        e.preventDefault();
        toggleUnderline();
        return;
      }

      if (isCtrl && (e.key === ';' || e.code === 'Semicolon')) {
        e.preventDefault();
        insertCurrentDate();
        return;
      }

      // Ctrl+Shift+L เปิด/ปิดตัวกรอง
      if (isCtrl && e.shiftKey && e.key.toLowerCase() === 'l') {
        e.preventDefault();
        toggleFilter();
        return;
      }

      if (isCtrl && e.key.toLowerCase() === 'a') {
        e.preventDefault();
        selectAll();
        return;
      }

      // ปุ่มเลื่อน (SPEC 5.1): ลูกศร, Ctrl+ลูกศร, Home/End, Ctrl+Home/End, PageUp/PageDown, Alt+PageUp/PageDown
      // Shift ร่วม = ขยายการเลือกจากขอบที่กำลังขยาย (focus) โดยเซลที่ใช้งานอยู่ไม่ย้าย
      const navTarget = navigationTarget(e, isCtrl);
      if (navTarget) {
        e.preventDefault();
        if (e.shiftKey) extendSelectionTo(navTarget.r, navTarget.c);
        else setActiveCell(navTarget.r, navTarget.c, false);
        const el = containerRef.current;
        if (navTarget.scrollBy && el) {
          el.scrollTop = Math.max(0, el.scrollTop + navTarget.scrollBy.top);
          el.scrollLeft = Math.max(0, el.scrollLeft + navTarget.scrollBy.left);
        }
        scrollToCell(navTarget.r, navTarget.c);
        return;
      }

      if (e.key === 'Tab') {
        e.preventDefault();
        setActiveCell(activeCell.r, activeCell.c + (e.shiftKey ? -1 : 1), false);
      } else if (e.key === 'Enter') {
        e.preventDefault();
        setActiveCell(stepRow(activeCell.r, e.shiftKey ? -1 : 1), activeCell.c, false);
      } else if (e.key === 'F2') {
        e.preventDefault();
        startEditing();
      } else if (e.key === 'Delete' || e.key === 'Backspace') {
        e.preventDefault();
        clearSelectionContent();
      } else if (e.key.length === 1 && !isCtrl && !e.altKey) {
        e.preventDefault();
        startEditing(e.key);
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [
    isEditing,
    activeCell,
    setActiveCell,
    startEditing,
    clearSelectionContent,
    insertCurrentDate,
    selectAll,
    copySelection,
    cutSelection,
    paste,
    fillDownAction,
    fillRightAction,
    cycleCurrentReference,
    undo,
    redo,
    toggleBold,
    toggleItalic,
    toggleUnderline,
    nextSheet,
    search.isOpen,
    openSearch,
    searchNext,
    searchPrev,
    prevSheet,
    stepRow,
    navigationTarget,
    extendSelectionTo,
    scrollToCell,
    toggleFilter,
    zoomIn,
    zoomOut,
    resetZoom,
  ]);

  // Point mode cell click / regular selection
  const handleCellMouseDown = (r: number, c: number, e: React.MouseEvent) => {
    if (e.button !== 0) return;

    if (isEditing && isPointModeReady()) {
      // Point mode: insert reference into formula and allow range drag
      // ถ้ากำลังดูชีตอื่น ใส่ชื่อชีตนำหน้า เช่น 'ชีต2'!A1
      e.preventDefault();
      const prefix = isOnEditingSheet ? '' : `${formatSheetReference(sheet.name)}!`;
      setPointModeAnchor({ r, c });
      setPointModeBaseFormula(editingValue);
      setPointModePrefix(prefix);
      setEditingValue(editingValue + prefix + cellToA1(r, c));
      return;
    }

    if (isEditing && !isOnEditingSheet) {
      // ไม่อยู่ในตำแหน่งที่รับ reference ได้ = ยืนยันการแก้ไข แล้วกลับชีตเดิม
      e.preventDefault();
      commitEditing(0, 0);
      return;
    }

    setIsMouseDown(true);
    if (e.shiftKey) {
      setActiveCell(r, c, true);
    } else {
      setActiveCell(r, c, false);
    }
  };

  const handleCellMouseEnter = (r: number, c: number) => {
    if (pointModeAnchor) {
      // Expand point mode range reference
      const r1 = Math.min(pointModeAnchor.r, r);
      const c1 = Math.min(pointModeAnchor.c, c);
      const r2 = Math.max(pointModeAnchor.r, r);
      const c2 = Math.max(pointModeAnchor.c, c);
      const refStr =
        r1 === r2 && c1 === c2
          ? cellToA1(r, c)
          : `${cellToA1(r1, c1)}:${cellToA1(r2, c2)}`;
      setEditingValue(pointModeBaseFormula + pointModePrefix + refStr);
      return;
    }

    if (isFillDragging) {
      // Expand fill preview
      const r1 = Math.min(selection.r1, r);
      const c1 = Math.min(selection.c1, c);
      const r2 = Math.max(selection.r2, r);
      const c2 = Math.max(selection.c2, c);
      setFillTarget({ r1, c1, r2, c2 });
      return;
    }

    if (!isMouseDown) return;
    setSelection(activeCell.r, activeCell.c, r, c);
  };

  // Fill handle mouse interactions
  const handleFillHandleMouseDown = (e: React.MouseEvent) => {
    e.stopPropagation();
    e.preventDefault();
    setIsFillDragging(true);
    setFillTarget({ ...selection });
  };

  const handleFillHandleDoubleClick = (e: React.MouseEvent) => {
    e.stopPropagation();
    // Auto-fill down until last non-empty row of adjacent column
    const adjCol = selection.c1 > 0 ? selection.c1 - 1 : selection.c2 + 1;
    let lastRow = selection.r2;

    for (let r = selection.r2 + 1; r < rowCount; r++) {
      const adjKey = cellToA1(r, adjCol);
      if (sheet.cells[adjKey] && sheet.cells[adjKey].raw) {
        lastRow = r;
      } else {
        break;
      }
    }

    if (lastRow > selection.r2) {
      applyFillRange({
        r1: selection.r1,
        c1: selection.c1,
        r2: lastRow,
        c2: selection.c2,
      });
    }
  };

  const handleMouseUp = () => {
    if (pointModeAnchor) {
      setPointModeAnchor(null);
      setPointModeBaseFormula('');
      setPointModePrefix('');
    }
    if (isFillDragging && fillTarget) {
      applyFillRange(fillTarget);
      setIsFillDragging(false);
      setFillTarget(null);
    }
    setIsMouseDown(false);
    setResizingCol(null);
    setResizingRow(null);
  };

  // Resize mouse move listener
  useEffect(() => {
    const handleMouseMove = (e: MouseEvent) => {
      if (resizingCol) {
        const delta = e.clientX - resizingCol.startX;
        setColWidth(resizingCol.colIdx, resizingCol.startW + delta / zoomFactor);
      } else if (resizingRow) {
        const delta = e.clientY - resizingRow.startY;
        setRowHeight(resizingRow.rowIdx, resizingRow.startH + delta / zoomFactor);
      }
    };

    window.addEventListener('mousemove', handleMouseMove);
    window.addEventListener('mouseup', handleMouseUp);
    return () => {
      window.removeEventListener('mousemove', handleMouseMove);
      window.removeEventListener('mouseup', handleMouseUp);
    };
  }, [resizingCol, resizingRow, isFillDragging, fillTarget, pointModeAnchor, setColWidth, setRowHeight, zoomFactor]);

  // Context Menu triggers
  const handleCellContextMenu = (r: number, c: number, e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (isEditing) {
      commitEditing();
    }
    const isInSelection =
      r >= selection.r1 &&
      r <= selection.r2 &&
      c >= selection.c1 &&
      c <= selection.c2;

    if (!isInSelection) {
      setActiveCell(r, c, false);
      setSelection(r, c, r, c);
    }
    setContextMenuPos({ x: e.clientX, y: e.clientY, targetType: 'cell' });
  };

  const handleRowHeaderContextMenu = (rowIdx: number, e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (isEditing) {
      commitEditing();
    }
    const isRowSelected =
      rowIdx >= selection.r1 &&
      rowIdx <= selection.r2 &&
      selection.c1 === 0 &&
      selection.c2 === colCount - 1;

    if (!isRowSelected) {
      setActiveCell(rowIdx, 0, false);
      setSelection(rowIdx, 0, rowIdx, colCount - 1);
    }
    setContextMenuPos({ x: e.clientX, y: e.clientY, targetType: 'row' });
  };

  const handleColHeaderContextMenu = (colIdx: number, e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (isEditing) {
      commitEditing();
    }
    const isColSelected =
      colIdx >= selection.c1 &&
      colIdx <= selection.c2 &&
      selection.r1 === 0 &&
      selection.r2 === rowCount - 1;

    if (!isColSelected) {
      setActiveCell(0, colIdx, false);
      setSelection(0, colIdx, rowCount - 1, colIdx);
    }
    setContextMenuPos({ x: e.clientX, y: e.clientY, targetType: 'col' });
  };

  // Header clicks
  const handleColHeaderClick = (colIdx: number) => {
    setActiveCell(0, colIdx, false);
    setSelection(0, colIdx, rowCount - 1, colIdx);
  };

  const handleRowHeaderClick = (rowIdx: number) => {
    setActiveCell(rowIdx, 0, false);
    setSelection(rowIdx, 0, rowIdx, colCount - 1);
  };

  const handleCornerClick = () => {
    setActiveCell(0, 0, false);
    setSelection(0, 0, rowCount - 1, colCount - 1);
  };

  // Auto-fit column width on double-click
  const handleColResizeDoubleClick = (colIdx: number) => {
    let maxLen = 4;
    for (let r = 0; r < rowCount; r++) {
      const val = getDisplayValue(r, colIdx);
      if (val) maxLen = Math.max(maxLen, val.length);
    }
    setColWidth(colIdx, Math.min(400, Math.max(60, maxLen * 11 + 24)));
  };

  // Calculate cell position for inline editor overlay
  const editorStyle = useMemo(() => {
    if (!isEditing) return { display: 'none' };

    let left = HEADER_COL_WIDTH;
    for (let c = 0; c < activeCell.c; c++) {
      left += sheet.colWidths[c] || DEFAULT_COL_WIDTH;
    }

    let top = HEADER_ROW_HEIGHT;
    for (let r = 0; r < activeCell.r; r++) {
      if (hiddenSet.has(r)) continue;
      top += sheet.rowHeights[r] || DEFAULT_ROW_HEIGHT;
    }

    const width = sheet.colWidths[activeCell.c] || DEFAULT_COL_WIDTH;
    const height = sheet.rowHeights[activeCell.r] || DEFAULT_ROW_HEIGHT;

    return {
      position: 'absolute' as const,
      left: `${left}px`,
      top: `${top}px`,
      width: `${Math.max(width, 140)}px`,
      height: `${height}px`,
      zIndex: 10,
    };
  }, [isEditing, activeCell, sheet.colWidths, sheet.rowHeights, hiddenSet]);

  // Position of fill handle at bottom-right corner of selection
  const fillHandleStyle = useMemo(() => {
    let left = HEADER_COL_WIDTH;
    for (let c = 0; c <= selection.c2; c++) {
      left += sheet.colWidths[c] || DEFAULT_COL_WIDTH;
    }

    let top = HEADER_ROW_HEIGHT;
    for (let r = 0; r <= selection.r2; r++) {
      if (hiddenSet.has(r)) continue;
      top += sheet.rowHeights[r] || DEFAULT_ROW_HEIGHT;
    }

    return {
      position: 'absolute' as const,
      left: `${left - 4}px`,
      top: `${top - 4}px`,
      width: '7px',
      height: '7px',
      backgroundColor: 'var(--fill-handle)',
      border: '1px solid var(--fill-handle-border)',
      cursor: 'crosshair',
      zIndex: 8,
    };
  }, [selection, sheet.colWidths, sheet.rowHeights, hiddenSet]);

  return (
    <div
      ref={containerRef}
      data-grid-white={gridWhite ? 'true' : undefined}
      onScroll={(e) => rememberScroll(e.currentTarget.scrollTop, e.currentTarget.scrollLeft)}
      style={{
        flex: 1,
        overflow: 'auto',
        position: 'relative',
        backgroundColor: 'var(--grid-bg)',
      }}
    >
      {/* ซูมเฉพาะพื้นที่ตาราง (CSS zoom): การคลิก/ลากใช้พิกัดของ element จึงตรงทุกระดับซูม */}
      <div style={{ position: 'relative', display: 'inline-block', zoom: zoomFactor }}>
        {/* Fill handle square */}
        {!isEditing && (
          <div
            style={fillHandleStyle}
            onMouseDown={handleFillHandleMouseDown}
            onDoubleClick={handleFillHandleDoubleClick}
            title="ลากเพื่อเติมข้อมูล / ดับเบิลคลิกเพื่อเติมลงล่าง"
          />
        )}

        {/* In-place editor */}
        {isEditing && isOnEditingSheet && (
          <input
            ref={inputRef}
            style={{
              ...editorStyle,
              border: '2px solid var(--grid-active-border)',
              padding: '0 6px',
              fontSize: '14px',
              fontFamily: 'inherit',
              boxSizing: 'border-box',
              outline: 'none',
              backgroundColor: 'var(--grid-bg)',
              color: 'var(--grid-fg)',
            }}
            value={editingValue}
            onChange={(e) => setEditingValue(e.target.value)}
            onKeyDown={(e) => {
              const isCtrl = e.ctrlKey || e.metaKey;
              if (isCtrl && (e.key === ';' || e.code === 'Semicolon')) {
                e.preventDefault();
                const dateStr = formatDate(todaySerial()); // ตามโหมดปี (พ.ศ./ค.ศ.)

                const input = inputRef.current;
                if (input) {
                  const start = input.selectionStart || 0;
                  const end = input.selectionEnd || 0;
                  const val = input.value;
                  const nextVal = val.slice(0, start) + dateStr + val.slice(end);
                  setEditingValue(nextVal);
                  setTimeout(() => {
                    input.selectionStart = input.selectionEnd = start + dateStr.length;
                  }, 0);
                } else {
                  setEditingValue(editingValue + dateStr);
                }
                return;
              }

              if (e.key === 'Enter') {
                e.preventDefault();
                commitEditing(e.shiftKey ? -1 : 1, 0);
              } else if (e.key === 'Tab') {
                e.preventDefault();
                commitEditing(0, e.shiftKey ? -1 : 1);
              } else if (e.key === 'Escape') {
                e.preventDefault();
                cancelEditing();
              }
            }}
          />
        )}

        <table
          style={{
            borderCollapse: 'collapse',
            tableLayout: 'fixed',
            userSelect: 'none',
          }}
        >
          <thead>
            <tr>
              {/* Corner header */}
              <th
                onClick={handleCornerClick}
                style={{
                  position: 'sticky',
                  top: 0,
                  left: 0,
                  zIndex: 5,
                  width: `${HEADER_COL_WIDTH}px`,
                  minWidth: `${HEADER_COL_WIDTH}px`,
                  height: `${HEADER_ROW_HEIGHT}px`,
                  backgroundColor: 'var(--grid-header-bg)',
                  border: '1px solid var(--grid-line)',
                  cursor: 'pointer',
                }}
              />

              {/* Column headers (A, B, C...) */}
              {Array.from({ length: colCount }, (_, c) => {
                const width = sheet.colWidths[c] || DEFAULT_COL_WIDTH;
                const isColSelected =
                  c >= selection.c1 &&
                  c <= selection.c2 &&
                  selection.r1 === 0 &&
                  selection.r2 === rowCount - 1;
                const isFrozen = c < sheet.frozenCols;

                return (
                  <th
                    key={c}
                    onClick={() => handleColHeaderClick(c)}
                    onContextMenu={(e) => handleColHeaderContextMenu(c, e)}
                    style={{
                      position: 'sticky',
                      top: 0,
                      zIndex: isFrozen ? 4 : 3,
                      width: `${width}px`,
                      minWidth: `${width}px`,
                      maxWidth: `${width}px`,
                      height: `${HEADER_ROW_HEIGHT}px`,
                      backgroundColor: isColSelected ? 'var(--grid-selection)' : 'var(--grid-header-bg)',
                      color: 'var(--grid-header-fg)',
                      border: '1px solid var(--grid-line)',
                      borderRight: isFrozen && c === sheet.frozenCols - 1 ? '3px solid var(--grid-active-border)' : '1px solid var(--grid-line)',
                      fontWeight: 600,
                      textAlign: 'center',
                      cursor: 'pointer',
                    }}
                  >
                    {colToStr(c)}

                    {/* Column resize drag handle */}
                    <div
                      onMouseDown={(e) => {
                        e.stopPropagation();
                        setResizingCol({ colIdx: c, startX: e.clientX, startW: width });
                      }}
                      onDoubleClick={(e) => {
                        e.stopPropagation();
                        handleColResizeDoubleClick(c);
                      }}
                      style={{
                        position: 'absolute',
                        top: 0,
                        right: 0,
                        width: '5px',
                        height: '100%',
                        cursor: 'col-resize',
                        zIndex: 10,
                      }}
                    />
                  </th>
                );
              })}
            </tr>
          </thead>

          <tbody>
            {Array.from({ length: rowCount }, (_, r) => {
              const height = sheet.rowHeights[r] || DEFAULT_ROW_HEIGHT;
              if (hiddenSet.has(r)) return null; // แถวที่ตัวกรองซ่อน (เลขแถวที่เหลือเป็นเลขจริง)
              const isRowSelected =
                r >= selection.r1 &&
                r <= selection.r2 &&
                selection.c1 === 0 &&
                selection.c2 === colCount - 1;
              const isFrozenRow = r < sheet.frozenRows;

              return (
                <tr key={r}>
                  {/* Row header (1, 2, 3...) */}
                  <th
                    onClick={() => handleRowHeaderClick(r)}
                    onContextMenu={(e) => handleRowHeaderContextMenu(r, e)}
                    style={{
                      position: 'sticky',
                      left: 0,
                      zIndex: isFrozenRow ? 4 : 2,
                      width: `${HEADER_COL_WIDTH}px`,
                      minWidth: `${HEADER_COL_WIDTH}px`,
                      height: `${height}px`,
                      backgroundColor: isRowSelected ? 'var(--grid-selection)' : 'var(--grid-header-bg)',
                      color: 'var(--grid-header-fg)',
                      border: '1px solid var(--grid-line)',
                      borderBottom: isFrozenRow && r === sheet.frozenRows - 1 ? '3px solid var(--grid-active-border)' : '1px solid var(--grid-line)',
                      fontWeight: 500,
                      textAlign: 'center',
                      cursor: 'pointer',
                    }}
                  >
                    {r + 1}

                    {/* Row resize drag handle */}
                    <div
                      onMouseDown={(e) => {
                        e.stopPropagation();
                        setResizingRow({ rowIdx: r, startY: e.clientY, startH: height });
                      }}
                      style={{
                        position: 'absolute',
                        bottom: 0,
                        left: 0,
                        width: '100%',
                        height: '5px',
                        cursor: 'row-resize',
                        zIndex: 10,
                      }}
                    />
                  </th>

                  {/* Grid cells */}
                  {Array.from({ length: colCount }, (_, c) => {
                    const coordKey = `${r},${c}`;
                    if (coveredMap.has(coordKey)) {
                      return null;
                    }

                    const a1 = cellToA1(r, c);
                    const cellData = sheet.cells[a1];
                    const merge = anchorMap[coordKey];
                    const width = sheet.colWidths[c] || DEFAULT_COL_WIDTH;

                    const isSelected =
                      r >= selection.r1 &&
                      r <= selection.r2 &&
                      c >= selection.c1 &&
                      c <= selection.c2;
                    const isActive = r === activeCell.r && c === activeCell.c;
                    const isCut =
                      visibleCutRange !== null &&
                      r >= visibleCutRange.r1 &&
                      r <= visibleCutRange.r2 &&
                      c >= visibleCutRange.c1 &&
                      c <= visibleCutRange.c2;

                    // Reference color highlight
                    const highlightColor = refHighlights[a1];

                    return (
                      <Cell
                        key={c}
                        r={r}
                        c={c}
                        a1={a1}
                        displayValue={getDisplayValue(r, c)}
                        rawValue={getRawValue(r, c)}
                        computedValue={getComputedValue(r, c)}
                        isSelected={isSelected}
                        isActive={isActive}
                        isCut={isCut}
                        isUnsupportedFormula={Boolean(cellData?.unsupportedFormula)}
                        highlightColor={highlightColor}
                        searchHighlight={searchCells.get(a1)}
                        filterButton={
                          sheet.filter && r === sheet.filter.r1 && c >= sheet.filter.c1 && c <= sheet.filter.c2
                            ? {
                                active: c in sheet.filter.columns,
                                title: th.filter.headerButtonTooltip,
                                onOpen: (x, y) => setFilterMenu({ col: c, x, y }),
                              }
                            : undefined
                        }
                        style={cellData?.style}
                        rowSpan={merge ? merge.r2 - merge.r1 + 1 : undefined}
                        colSpan={merge ? merge.c2 - merge.c1 + 1 : undefined}
                        width={width}
                        height={height}
                        onMouseDown={(e) => handleCellMouseDown(r, c, e)}
                        onMouseEnter={() => handleCellMouseEnter(r, c)}
                        onDoubleClick={() => startEditing()}
                        onContextMenu={(e) => handleCellContextMenu(r, c, e)}
                      />
                    );
                  })}
                </tr>
              );
            })}
          </tbody>
        </table>

        {/* Append rows bar at bottom */}
        <div
          style={{
            padding: '16px 20px',
            display: 'flex',
            alignItems: 'center',
            gap: '10px',
            backgroundColor: 'var(--grid-bg)',
          }}
        >
          <span style={{ fontSize: '13px', color: 'var(--grid-header-fg)' }}>
            {th.grid.appendRowsLabel}
          </span>
          <button
            onClick={() => addRowsAtBottom(50)}
            style={{
              padding: '4px 12px',
              fontSize: '13px',
              fontWeight: 500,
              backgroundColor: 'var(--grid-header-bg)',
              color: 'var(--grid-fg)',
              border: '1px solid var(--control-border)',
              borderRadius: '4px',
              cursor: 'pointer',
            }}
          >
            {th.grid.add50Rows}
          </button>
          <button
            onClick={() => addRowsAtBottom(100)}
            style={{
              padding: '4px 12px',
              fontSize: '13px',
              fontWeight: 500,
              backgroundColor: 'var(--grid-header-bg)',
              color: 'var(--grid-fg)',
              border: '1px solid var(--control-border)',
              borderRadius: '4px',
              cursor: 'pointer',
            }}
          >
            {th.grid.add100Rows}
          </button>
          <span style={{ fontSize: '12px', color: 'var(--grid-header-fg)', marginLeft: '8px' }}>
            {th.grid.currentRows.replace('{count}', String(rowCount))}
          </span>
        </div>
      </div>

      {/* เมนูตัวกรองของคอลัมน์ (4C) */}
      {filterMenu && (
        <FilterMenu
          col={filterMenu.col}
          x={filterMenu.x}
          y={filterMenu.y}
          onClose={() => setFilterMenu(null)}
        />
      )}

      {/* Context Menu */}
      {contextMenuPos && (
        <ContextMenu
          x={contextMenuPos.x}
          y={contextMenuPos.y}
          targetType={contextMenuPos.targetType}
          onClose={() => setContextMenuPos(null)}
          onCopy={copySelection}
          onCut={cutSelection}
          onPaste={() => paste(false)}
          onPasteValuesOnly={() => paste(true)}
          onInsertRowAbove={() => insertRowAt(activeCell.r)}
          onInsertRowBelow={() => insertRowAt(activeCell.r + 1)}
          onInsertColLeft={() => insertColAt(activeCell.c)}
          onInsertColRight={() => insertColAt(activeCell.c + 1)}
          onDeleteRow={() => deleteRowAt(activeCell.r)}
          onDeleteCol={() => deleteColAt(activeCell.c)}
          onClearContent={clearSelectionContent}
          onClearFormat={clearSelectionFormat}
          onMerge={mergeCells}
          onUnmerge={unmergeCells}
          onFreezeRow={() => freezeTo(activeCell.r + 1, sheet.frozenCols)}
          onFreezeCol={() => freezeTo(sheet.frozenRows, activeCell.c + 1)}
          onUnfreeze={() => freezeTo(0, 0)}
          onSortAsc={() => requestSort('asc')}
          onSortDesc={() => requestSort('desc')}
          onSortAdvanced={openAdvancedSort}
        />
      )}
    </div>
  );
};
