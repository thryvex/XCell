import React, { useMemo } from 'react';
import { th } from '../i18n/th';
import { IconMinus, IconPlus } from './icons';
import { useSheetStore, ZOOM_MIN, ZOOM_MAX } from './store';
import { cellToA1 } from '../core/address';
import { formatNumberWithCommas, formatDate } from '../core/numfmt';
import { filterDataRowCount } from '../core/filter';

export const StatusBar: React.FC = () => {
  const { selection, getComputedValue, isDirty, workbook, activeSheetIndex, zoomIn, zoomOut, resetZoom, yearMode } =
    useSheetStore();
  const sheet = workbook.sheets[activeSheetIndex];
  const filter = sheet?.filter;
  const hiddenRows = sheet?.hiddenRows;
  const zoom = sheet?.zoom ?? 100;
  const zoomButton: React.CSSProperties = {
    border: '1px solid transparent',
    background: 'transparent',
    padding: '0 6px',
    fontSize: '13px',
    cursor: 'pointer',
    color: 'var(--fg-main)',
  };

  const isSingleCell = selection.r1 === selection.r2 && selection.c1 === selection.c2;
  const rangeText = isSingleCell
    ? cellToA1(selection.r1, selection.c1)
    : `${cellToA1(selection.r1, selection.c1)}:${cellToA1(selection.r2, selection.c2)}`;

  const stats = useMemo(() => {
    const numbers: number[] = [];
    let count = 0;
    let allDates = true; // ตัวเลขทุกเซลเป็นวันที่ → แสดงสถิติเป็นวันที่ (ตามโหมดปี) แทนเลข serial

    // รวม/เฉลี่ย/นับ ของช่วงที่เลือก นับเฉพาะแถวที่มองเห็น (แถวที่ตัวกรองซ่อนไม่นับ)
    const hidden = new Set(hiddenRows ?? []);
    for (let r = selection.r1; r <= selection.r2; r++) {
      if (hidden.has(r)) continue;
      for (let c = selection.c1; c <= selection.c2; c++) {
        const val = getComputedValue(r, c);
        if (val !== '' && val !== null && val !== undefined) {
          count++;
        }
        if (typeof val === 'number' && !isNaN(val) && isFinite(val)) {
          if (sheet?.cells[cellToA1(r, c)]?.fmt !== 'date') allDates = false;
          numbers.push(val);
        }
      }
    }

    if (numbers.length === 0) {
      return null;
    }

    const sum = numbers.reduce((a, b) => a + b, 0);
    const avg = sum / numbers.length;
    const min = Math.min(...numbers);
    const max = Math.max(...numbers);
    const fmt = (n: number) => (allDates ? formatDate(Math.round(n)) : formatNumberWithCommas(n, 2));

    return {
      sum: fmt(sum),
      avg: fmt(avg),
      count: numbers.length,
      min: fmt(min),
      max: fmt(max),
    };
  }, [selection, getComputedValue, hiddenRows, sheet, yearMode]);

  return (
    <div
      style={{
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'space-between',
        padding: '4px 12px',
        borderTop: '1px solid var(--border-color)',
        backgroundColor: 'var(--status-bg)',
        color: 'var(--status-fg)',
        fontSize: '13px',
      }}
    >
      <div style={{ display: 'flex', alignItems: 'center', gap: '16px' }}>
        <span>{rangeText ? th.statusBar.selection.replace('{range}', rangeText) : ''}</span>

        {filter && (
          <span style={{ color: 'var(--color-primary)', fontWeight: 500 }}>
            {th.filter.statusFiltered
              .replace('{shown}', String(filterDataRowCount(filter) - (hiddenRows?.length ?? 0)))
              .replace('{total}', String(filterDataRowCount(filter)))}
          </span>
        )}

        {stats ? (
          <span style={{ color: 'var(--fg-main)', fontWeight: 500 }}>
            {th.statusBar.stats
              .replace('{sum}', stats.sum)
              .replace('{avg}', stats.avg)
              .replace('{count}', String(stats.count))
              .replace('{min}', stats.min)
              .replace('{max}', stats.max)}
          </span>
        ) : (
          <span>{th.statusBar.ready}</span>
        )}
      </div>

      <div style={{ display: 'flex', alignItems: 'center', gap: '14px' }}>
        {/* ซูม 50%–200% (4C) */}
        <div style={{ display: 'flex', alignItems: 'center', gap: '2px' }}>
          <button style={zoomButton} onClick={zoomOut} title={th.zoom.zoomOut} disabled={zoom <= ZOOM_MIN}>
            <IconMinus size={16} />
          </button>
          <button style={{ ...zoomButton, minWidth: '46px' }} onClick={resetZoom} title={th.zoom.reset}>
            {th.zoom.percent.replace('{percent}', String(zoom))}
          </button>
          <button style={zoomButton} onClick={zoomIn} title={th.zoom.zoomIn} disabled={zoom >= ZOOM_MAX}>
            <IconPlus size={16} />
          </button>
        </div>
        {isDirty ? (
          <span style={{ color: 'var(--color-danger)' }}>{th.statusBar.unsaved}</span>
        ) : (
          <span style={{ color: 'var(--color-success)' }}>{th.statusBar.saved}</span>
        )}
      </div>
    </div>
  );
};
