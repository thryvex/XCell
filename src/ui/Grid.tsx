import React, { useRef, useEffect, useState, useMemo, useCallback } from 'react';
import { useSheetStore } from './store';
import { Cell } from './Cell';
import { colToStr, cellToA1 } from '../core/address';
import { MergeRange } from '../core/model';

const DEFAULT_COL_WIDTH = 110;
const DEFAULT_ROW_HEIGHT = 28;
const HEADER_ROW_HEIGHT = 28;
const HEADER_COL_WIDTH = 45;

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
    setActiveCell,
    setSelection,
    startEditing,
    setEditingValue,
    commitEditing,
    cancelEditing,
    clearSelectionContent,
    undo,
    redo,
    getDisplayValue,
    getRawValue,
    getComputedValue,
  } = useSheetStore();

  const sheet = workbook.sheets[activeSheetIndex];
  const [isMouseDown, setIsMouseDown] = useState(false);

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

  // Focus input when editing starts
  useEffect(() => {
    if (isEditing && inputRef.current) {
      inputRef.current.focus();
    }
  }, [isEditing]);

  // Handle global keyboard shortcuts in grid
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
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
        // Typing printable character replaces cell content
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
    undo,
    redo,
  ]);

  const handleCellMouseDown = (r: number, c: number, e: React.MouseEvent) => {
    if (e.button !== 0) return;
    setIsMouseDown(true);
    if (e.shiftKey) {
      setActiveCell(r, c, true);
    } else {
      setActiveCell(r, c, false);
    }
  };

  const handleCellMouseEnter = (r: number, c: number) => {
    if (!isMouseDown) return;
    setSelection(activeCell.r, activeCell.c, r, c);
  };

  const handleMouseUp = () => {
    setIsMouseDown(false);
  };

  useEffect(() => {
    window.addEventListener('mouseup', handleMouseUp);
    return () => window.removeEventListener('mouseup', handleMouseUp);
  }, []);

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
                  zIndex: 4,
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

                return (
                  <th
                    key={c}
                    onClick={() => handleColHeaderClick(c)}
                    style={{
                      position: 'sticky',
                      top: 0,
                      zIndex: 3,
                      width: `${width}px`,
                      minWidth: `${width}px`,
                      maxWidth: `${width}px`,
                      height: `${HEADER_ROW_HEIGHT}px`,
                      backgroundColor: isColSelected ? 'var(--bg-selection)' : 'var(--bg-header)',
                      color: 'var(--fg-muted)',
                      border: '1px solid var(--border-color)',
                      fontWeight: 600,
                      textAlign: 'center',
                      cursor: 'pointer',
                    }}
                  >
                    {colToStr(c)}
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

              return (
                <tr key={r}>
                  {/* Row header (1, 2, 3...) */}
                  <th
                    onClick={() => handleRowHeaderClick(r)}
                    style={{
                      position: 'sticky',
                      left: 0,
                      zIndex: 2,
                      width: `${HEADER_COL_WIDTH}px`,
                      minWidth: `${HEADER_COL_WIDTH}px`,
                      height: `${height}px`,
                      backgroundColor: isRowSelected ? 'var(--bg-selection)' : 'var(--bg-header)',
                      color: 'var(--fg-muted)',
                      border: '1px solid var(--border-color)',
                      fontWeight: 500,
                      textAlign: 'center',
                      cursor: 'pointer',
                    }}
                  >
                    {r + 1}
                  </th>

                  {/* Grid cells */}
                  {Array.from({ length: colCount }, (_, c) => {
                    const coordKey = `${r},${c}`;
                    if (coveredMap.has(coordKey)) {
                      return null; // Covered by merged cell
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
                        style={cellData?.style}
                        rowSpan={merge ? merge.r2 - merge.r1 + 1 : undefined}
                        colSpan={merge ? merge.c2 - merge.c1 + 1 : undefined}
                        width={width}
                        height={height}
                        onMouseDown={(e) => handleCellMouseDown(r, c, e)}
                        onMouseEnter={() => handleCellMouseEnter(r, c)}
                        onDoubleClick={() => startEditing()}
                      />
                    );
                  })}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
};
