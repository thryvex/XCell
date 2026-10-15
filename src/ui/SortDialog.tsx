import React from 'react';
import { useSheetStore } from './store';
import { th } from '../i18n/th';
import { colToStr } from '../core/address';
import { rangeLabel, SortKey } from '../core/sort';

const MAX_SORT_KEYS = 3;

const overlay: React.CSSProperties = {
  position: 'fixed',
  inset: 0,
  backgroundColor: 'var(--overlay)',
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  zIndex: 3000,
};

const panel: React.CSSProperties = {
  backgroundColor: 'var(--bg-app)',
  borderRadius: '8px',
  padding: '20px 24px',
  width: '440px',
  maxWidth: '92%',
  boxShadow: 'var(--shadow-dialog)',
  display: 'flex',
  flexDirection: 'column',
  gap: '14px',
  fontSize: '14px',
};

/**
 * หน้าต่างเรียงลำดับ (แบบเร็วและแบบหลายคอลัมน์), คำเตือนสูตรอ้างแถวอื่น และข้อความแจ้ง
 */
export const SortDialog: React.FC = () => {
  const {
    pendingSort,
    notice,
    dismissNotice,
    updatePendingSort,
    confirmSort,
    cancelSort,
    getDisplayValue,
  } = useSheetStore();

  if (notice) {
    return (
      <div style={overlay}>
        <div style={panel} role="alertdialog">
          <p style={{ margin: 0 }}>{notice}</p>
          <div style={{ display: 'flex', justifyContent: 'flex-end' }}>
            <button autoFocus onClick={dismissNotice}>
              {th.dialog.ok}
            </button>
          </div>
        </div>
      </div>
    );
  }

  if (!pendingSort) return null;
  const { range, hasHeader, keys, mode, step } = pendingSort;

  if (step === 'formulaWarning') {
    return (
      <div style={overlay}>
        <div style={panel} role="alertdialog">
          <h3 style={{ margin: 0, fontSize: '16px' }}>⚠️ {th.sort.dialogTitle}</h3>
          <p style={{ margin: 0 }}>{th.sort.formulaWarning}</p>
          <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '8px' }}>
            <button onClick={cancelSort}>{th.dialog.cancel}</button>
            <button autoFocus onClick={confirmSort} style={{ fontWeight: 600 }}>
              {th.sort.continue}
            </button>
          </div>
        </div>
      </div>
    );
  }

  const columnLabel = (col: number) => {
    const letter = colToStr(col);
    if (hasHeader) {
      const header = getDisplayValue(range.r1, col);
      if (header) return th.sort.columnWithHeader.replace('{col}', letter).replace('{header}', header);
    }
    return th.sort.column.replace('{col}', letter);
  };

  const columns: number[] = [];
  for (let c = range.c1; c <= range.c2; c++) columns.push(c);

  const setKey = (i: number, patch: Partial<SortKey>) => {
    updatePendingSort({ keys: keys.map((k, j) => (j === i ? { ...k, ...patch } : k)) });
  };

  const unusedColumn = columns.find((c) => !keys.some((k) => k.col === c));
  const dirLabel = (dir: 'asc' | 'desc') => (dir === 'asc' ? th.sort.dirAsc : th.sort.dirDesc);

  return (
    <div style={overlay}>
      <div
        style={panel}
        role="dialog"
        onKeyDown={(e) => {
          if (e.key === 'Escape') cancelSort();
        }}
      >
        <h3 style={{ margin: 0, fontSize: '16px' }}>
          {mode === 'advanced' ? th.sort.advancedTitle : th.sort.dialogTitle}
        </h3>
        <div style={{ color: 'var(--fg-muted)' }}>
          {th.sort.rangeLabel.replace('{range}', rangeLabel(range))}
        </div>

        <label style={{ display: 'flex', alignItems: 'center', gap: '8px', cursor: 'pointer' }}>
          <input
            type="checkbox"
            checked={hasHeader}
            onChange={(e) => updatePendingSort({ hasHeader: e.target.checked })}
          />
          {th.sort.hasHeader}
        </label>

        {mode === 'quick' ? (
          <div>
            {th.sort.sortBy} {columnLabel(keys[0].col)} ({dirLabel(keys[0].dir)})
          </div>
        ) : (
          <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
            {keys.map((k, i) => (
              <div key={i} style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                <span style={{ width: '84px' }}>{i === 0 ? th.sort.sortBy : th.sort.thenBy}</span>
                <select
                  value={k.col}
                  onChange={(e) => setKey(i, { col: Number(e.target.value) })}
                  style={{ flex: 1, minWidth: 0 }}
                >
                  {columns.map((c) => (
                    <option key={c} value={c}>
                      {columnLabel(c)}
                    </option>
                  ))}
                </select>
                <select
                  value={k.dir}
                  onChange={(e) => setKey(i, { dir: e.target.value as 'asc' | 'desc' })}
                >
                  <option value="asc">{th.sort.dirAsc}</option>
                  <option value="desc">{th.sort.dirDesc}</option>
                </select>
                {i > 0 && (
                  <button onClick={() => updatePendingSort({ keys: keys.filter((_, j) => j !== i) })}>
                    {th.sort.removeKey}
                  </button>
                )}
              </div>
            ))}
            {keys.length < MAX_SORT_KEYS && unusedColumn !== undefined && (
              <button
                style={{ alignSelf: 'flex-start' }}
                onClick={() => updatePendingSort({ keys: [...keys, { col: unusedColumn, dir: 'asc' }] })}
              >
                {th.sort.addKey}
              </button>
            )}
          </div>
        )}

        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '8px' }}>
          <button onClick={cancelSort}>{th.dialog.cancel}</button>
          <button autoFocus onClick={confirmSort} style={{ fontWeight: 600 }}>
            {th.sort.confirm}
          </button>
        </div>
      </div>
    </div>
  );
};
