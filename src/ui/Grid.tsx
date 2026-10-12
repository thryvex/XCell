import React, { useRef, useEffect, useState, useMemo, useCallback } from 'react';
import { useSheetStore, SelectionState } from './store';
import { Cell } from './Cell';
import { ContextMenu } from './ContextMenu';
import { colToStr, cellToA1, parseRange } from '../core/address';
import { tokenize } from '../core/tokenizer';
import { MergeRange } from '../core/model';

const DEFAULT_COL_WIDTH = 110;
const DEFAULT_ROW_HEIGHT = 28;
const HEADER_ROW_HEIGHT = 28;
const HEADER_COL_WIDTH = 45;

const HIGHLIGHT_COLORS = [
  '#2563eb', // Blue
  '#dc2626', // Red
  '#16a34a', // Green
  '#9333ea', // Purple
  '#ea580c', // Orange
  '#0891b2', // Cyan
];

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
    rowCount,
    colCount,
    cutRange,
    setActiveCell,
    setSelection,
    startEditing,
    setEditingValue,
    commitEditing,
    cancelEditing,
    clearSelectionContent,
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
    undo,
    redo,
    mergeCells,
    unmergeCells,
    getDisplayValue,
    getRawValue,
    getComputedValue,
  } = useSheetStore();

  const sheet = workbook.sheets[activeSheetIndex];

  // Mouse states
  const [isMouseDown, setIsMouseDown] = useState(false);
  const [isFillDragging, setIsFillDragging] = useState(false);
  const [fillTarget, setFillTarget] = useState<SelectionState | null>(null);

  // Resize states
  const [resizingCol, setResizingCol] = useState<{ colIdx: number; startX: number; startW: number } | null>(null);
  const [resizingRow, setResizingRow] = useState<{ rowIdx: number; startY: number; startH: number } | null>(null);

  // Context Menu state
  const [contextMenuPos, setContextMenuPos] = useState<{ x: number; y: number } | null>(null);

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
        if (t.type === 'CELL_REF') {
          const color = HIGHLIGHT_COLORS[colorIdx % HIGHLIGHT_COLORS.length];
          map[t.value.toUpperCase()] = color;
          colorIdx++;
        } else if (t.type === 'RANGE') {
          const rng = parseRange(t.value);
          if (rng) {
            const color = HIGHLIGHT_COLORS[colorIdx % HIGHLIGHT_COLORS.length];
            for (let r = rng.r1; r <= rng.r2; r++) {
              for (let c = rng.c1; c <= rng.c2; c++) {
                map[cellToA1(r, c)] = color;
              }
            }
            colorIdx++;
          }
        }
      }
    } catch {}

    return map;
  }, [isEditing, editingValue]);

  // Focus input when editing starts
  useEffect(() => {
    if (isEditing && inputRef.current) {
      inputRef.current.focus();
    }
  }, [isEditing]);

  // Check if cursor in editing formula is ready for Point Mode reference insertion
  const isPointModeReady = useCallback(() => {
    if (!isEditing || !editingValue.startsWith('=')) return false;
    const trimmed = editingValue.trimEnd();
    const lastChar = trimmed[trimmed.length - 1];
    return '=+-*/^%&,(<>'.includes(lastChar);
  }, [isEditing, editingValue]);

  // Global keyboard shortcuts in grid
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      // F4 key works both inside editor and grid
      if (e.key === 'F4') {
        e.preventDefault();
        cycleCurrentReference();
        return;
      }

      if (isEditing) return;

      const isCtrl = e.ctrlKey || e.metaKey;

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

      if (e.key === 'ArrowUp') {
        e.preventDefault();
        setActiveCell(activeCell.r - 1, activeCell.c, e.shiftKey);
      } else if (e.key === 'ArrowDown') {
        e.preventDefault();
        setActiveCell(activeCell.r + 1, activeCell.c, e.shiftKey);
      } else if (e.key === 'ArrowLeft') {
        e.preventDefault();
        setActiveCell(activeCell.r, activeCell.c - 1, e.shiftKey);
      } else if (e.key === 'ArrowRight') {
        e.preventDefault();
        setActiveCell(activeCell.r, activeCell.c + 1, e.shiftKey);
      } else if (e.key === 'Tab') {
        e.preventDefault();
        setActiveCell(activeCell.r, activeCell.c + (e.shiftKey ? -1 : 1), false);
      } else if (e.key === 'Enter') {
        e.preventDefault();
        setActiveCell(activeCell.r + (e.shiftKey ? -1 : 1), activeCell.c, false);
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
    copySelection,
    cutSelection,
    paste,
    fillDownAction,
    fillRightAction,
    cycleCurrentReference,
    undo,
    redo,
  ]);

  // Point mode cell click / regular selection
  const handleCellMouseDown = (r: number, c: number, e: React.MouseEvent) => {
    if (e.button !== 0) return;

    if (isEditing && isPointModeReady()) {
      // Point mode: insert reference into formula
      e.preventDefault();
      const clickedA1 = cellToA1(r, c);
      setEditingValue(editingValue + clickedA1);
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
        setColWidth(resizingCol.colIdx, resizingCol.startW + delta);
      } else if (resizingRow) {
        const delta = e.clientY - resizingRow.startY;
        setRowHeight(resizingRow.rowIdx, resizingRow.startH + delta);
      }
    };

    window.addEventListener('mousemove', handleMouseMove);
    window.addEventListener('mouseup', handleMouseUp);
    return () => {
      window.removeEventListener('mousemove', handleMouseMove);
      window.removeEventListener('mouseup', handleMouseUp);
    };
  }, [resizingCol, resizingRow, isFillDragging, fillTarget, setColWidth, setRowHeight]);

  // Context Menu trigger
  const handleCellContextMenu = (r: number, c: number, e: React.MouseEvent) => {
    e.preventDefault();
    if (r < selection.r1 || r > selection.r2 || c < selection.c1 || c > selection.c2) {
      setActiveCell(r, c, false);
    }
    setContextMenuPos({ x: e.clientX, y: e.clientY });
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
  }, [isEditing, activeCell, sheet.colWidths, sheet.rowHeights]);

  // Position of fill handle at bottom-right corner of selection
  const fillHandleStyle = useMemo(() => {
    let left = HEADER_COL_WIDTH;
    for (let c = 0; c <= selection.c2; c++) {
      left += sheet.colWidths[c] || DEFAULT_COL_WIDTH;
    }

    let top = HEADER_ROW_HEIGHT;
    for (let r = 0; r <= selection.r2; r++) {
      top += sheet.rowHeights[r] || DEFAULT_ROW_HEIGHT;
    }

    return {
      position: 'absolute' as const,
      left: `${left - 4}px`,
      top: `${top - 4}px`,
      width: '7px',
      height: '7px',
      backgroundColor: 'var(--color-primary)',
      border: '1px solid #ffffff',
      cursor: 'crosshair',
      zIndex: 8,
    };
  }, [selection, sheet.colWidths, sheet.rowHeights]);

  return (
    <div
      ref={containerRef}
      style={{
        flex: 1,
        overflow: 'auto',
        position: 'relative',
        backgroundColor: 'var(--bg-app)',
      }}
    >
      <div style={{ position: 'relative', display: 'inline-block' }}>
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
        {isEditing && (
          <input
            ref={inputRef}
            style={{
              ...editorStyle,
              border: '2px solid var(--color-primary)',
              padding: '0 6px',
              fontSize: '14px',
              fontFamily: 'inherit',
              boxSizing: 'border-box',
              outline: 'none',
              backgroundColor: 'var(--bg-app)',
              color: 'var(--fg-main)',
            }}
            value={editingValue}
            onChange={(e) => setEditingValue(e.target.value)}
            onKeyDown={(e) => {
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
                  backgroundColor: 'var(--bg-header)',
                  border: '1px solid var(--border-color)',
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
                    onContextMenu={(e) => {
                      e.preventDefault();
                      handleColHeaderClick(c);
                      setContextMenuPos({ x: e.clientX, y: e.clientY });
                    }}
                    style={{
                      position: 'sticky',
                      top: 0,
                      zIndex: isFrozen ? 4 : 3,
                      width: `${width}px`,
                      minWidth: `${width}px`,
                      maxWidth: `${width}px`,
                      height: `${HEADER_ROW_HEIGHT}px`,
                      backgroundColor: isColSelected ? 'var(--bg-selection)' : 'var(--bg-header)',
                      color: 'var(--fg-muted)',
                      border: '1px solid var(--border-color)',
                      borderRight: isFrozen && c === sheet.frozenCols - 1 ? '3px solid var(--color-primary)' : '1px solid var(--border-color)',
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
                    onContextMenu={(e) => {
                      e.preventDefault();
                      handleRowHeaderClick(r);
                      setContextMenuPos({ x: e.clientX, y: e.clientY });
                    }}
                    style={{
                      position: 'sticky',
                      left: 0,
                      zIndex: isFrozenRow ? 4 : 2,
                      width: `${HEADER_COL_WIDTH}px`,
                      minWidth: `${HEADER_COL_WIDTH}px`,
                      height: `${height}px`,
                      backgroundColor: isRowSelected ? 'var(--bg-selection)' : 'var(--bg-header)',
                      color: 'var(--fg-muted)',
                      border: '1px solid var(--border-color)',
                      borderBottom: isFrozenRow && r === sheet.frozenRows - 1 ? '3px solid var(--color-primary)' : '1px solid var(--border-color)',
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
                      cutRange !== null &&
                      r >= cutRange.r1 &&
                      r <= cutRange.r2 &&
                      c >= cutRange.c1 &&
                      c <= cutRange.c2;

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
                        highlightColor={highlightColor}
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
            backgroundColor: 'var(--bg-app)',
          }}
        >
          <span style={{ fontSize: '13px', color: 'var(--fg-muted)' }}>
            เพิ่มแถวที่ด้านล่าง:
          </span>
          <button
            onClick={() => addRowsAtBottom(50)}
            style={{
              padding: '4px 12px',
              fontSize: '13px',
              fontWeight: 500,
              backgroundColor: 'var(--bg-toolbar)',
              color: 'var(--fg-primary)',
              border: '1px solid var(--border-color)',
              borderRadius: '4px',
              cursor: 'pointer',
            }}
          >
            ＋ เพิ่ม 50 แถว
          </button>
          <button
            onClick={() => addRowsAtBottom(100)}
            style={{
              padding: '4px 12px',
              fontSize: '13px',
              fontWeight: 500,
              backgroundColor: 'var(--bg-toolbar)',
              color: 'var(--fg-primary)',
              border: '1px solid var(--border-color)',
              borderRadius: '4px',
              cursor: 'pointer',
            }}
          >
            ＋ เพิ่ม 100 แถว
          </button>
          <span style={{ fontSize: '12px', color: 'var(--fg-muted)', marginLeft: '8px' }}>
            (ปัจจุบันมี {rowCount} แถว)
          </span>
        </div>
      </div>

      {/* Context Menu */}
      {contextMenuPos && (
        <ContextMenu
          x={contextMenuPos.x}
          y={contextMenuPos.y}
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
          onMerge={mergeCells}
          onUnmerge={unmergeCells}
          onFreezeRow={() => freezeTo(activeCell.r + 1, sheet.frozenCols)}
          onFreezeCol={() => freezeTo(sheet.frozenRows, activeCell.c + 1)}
          onUnfreeze={() => freezeTo(0, 0)}
        />
      )}
    </div>
  );
};
