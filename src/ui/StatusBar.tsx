import React, { useMemo } from 'react';
import { th } from '../i18n/th';
import { useSheetStore } from './store';
import { cellToA1 } from '../core/address';
import { formatNumberWithCommas } from '../core/numfmt';

export const StatusBar: React.FC = () => {
  const { selection, getComputedValue, isDirty } = useSheetStore();

  const isSingleCell = selection.r1 === selection.r2 && selection.c1 === selection.c2;
  const rangeText = isSingleCell
    ? cellToA1(selection.r1, selection.c1)
    : `${cellToA1(selection.r1, selection.c1)}:${cellToA1(selection.r2, selection.c2)}`;

  const stats = useMemo(() => {
    const numbers: number[] = [];
    let count = 0;

    for (let r = selection.r1; r <= selection.r2; r++) {
      for (let c = selection.c1; c <= selection.c2; c++) {
        const val = getComputedValue(r, c);
        if (val !== '' && val !== null && val !== undefined) {
          count++;
        }
        if (typeof val === 'number' && !isNaN(val) && isFinite(val)) {
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

    return {
      sum: formatNumberWithCommas(sum, 2),
      avg: formatNumberWithCommas(avg, 2),
      count: numbers.length,
      min: formatNumberWithCommas(min, 2),
      max: formatNumberWithCommas(max, 2),
    };
  }, [selection, getComputedValue]);

  return (
    <div
      style={{
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'space-between',
        padding: '4px 12px',
        borderTop: '1px solid var(--border-color)',
        backgroundColor: 'var(--bg-surface)',
        color: 'var(--fg-muted)',
        fontSize: '13px',
      }}
    >
      <div style={{ display: 'flex', alignItems: 'center', gap: '16px' }}>
        <span>{rangeText ? th.statusBar.selection.replace('{range}', rangeText) : ''}</span>

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

      <div>
        {isDirty ? (
          <span style={{ color: 'var(--color-danger)' }}>{th.statusBar.unsaved}</span>
        ) : (
          <span style={{ color: 'var(--color-success)' }}>{th.statusBar.saved}</span>
        )}
      </div>
    </div>
  );
};
