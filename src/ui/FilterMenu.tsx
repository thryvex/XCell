import React, { useMemo, useState } from 'react';
import { useSheetStore } from './store';
import { th } from '../i18n/th';
import { uniqueColumnValues } from '../core/filter';
import { ColumnFilter, FilterCondition } from '../core/model';

type ConditionType = FilterCondition['type'] | 'none';

const CONDITION_LABELS: Record<ConditionType, string> = {
  none: th.filter.conditionNone,
  empty: th.filter.conditionEmpty,
  notEmpty: th.filter.conditionNotEmpty,
  textContains: th.filter.conditionTextContains,
  textNotContains: th.filter.conditionTextNotContains,
  startsWith: th.filter.conditionStartsWith,
  equals: th.filter.conditionEquals,
  greaterThan: th.filter.conditionGreaterThan,
  lessThan: th.filter.conditionLessThan,
  between: th.filter.conditionBetween,
};

const NUMERIC_CONDITIONS: ConditionType[] = ['greaterThan', 'lessThan', 'between'];

interface FilterMenuProps {
  col: number;
  x: number;
  y: number;
  onClose: () => void;
}

/**
 * เมนูของปุ่มลูกศรที่หัวคอลัมน์ตัวกรอง: เรียง, กรองตามค่า, กรองตามเงื่อนไข, ใช้ตัวกรองอีกครั้ง
 */
export const FilterMenu: React.FC<FilterMenuProps> = ({ col, x, y, onClose }) => {
  const {
    workbook,
    activeSheetIndex,
    evaluator,
    setColumnFilter,
    reapplyFilter,
    sortFilterColumn,
  } = useSheetStore();
  const sheet = workbook.sheets[activeSheetIndex];
  const existing: ColumnFilter | undefined = sheet.filter?.columns[col];

  const values = useMemo(
    () => uniqueColumnValues(sheet, evaluator, activeSheetIndex, col),
    [sheet, evaluator, activeSheetIndex, col]
  );

  const [mode, setMode] = useState<'values' | 'condition'>(existing?.kind === 'condition' ? 'condition' : 'values');
  const [hiddenValues, setHiddenValues] = useState<Set<string>>(
    new Set(existing?.kind === 'values' ? existing.hiddenValues : [])
  );
  const [search, setSearch] = useState('');
  const initialCond = existing?.kind === 'condition' ? existing.condition : undefined;
  const [condType, setCondType] = useState<ConditionType>(initialCond?.type ?? 'none');
  const [v1, setV1] = useState(
    initialCond && 'value' in initialCond ? initialCond.value : initialCond?.type === 'between' ? initialCond.min : ''
  );
  const [v2, setV2] = useState(initialCond?.type === 'between' ? initialCond.max : '');

  const shownValues = values.filter((v) =>
    (v.display || th.filter.blankValue).toLocaleLowerCase().includes(search.toLocaleLowerCase())
  );

  const toggleValue = (display: string) => {
    const next = new Set(hiddenValues);
    if (next.has(display)) next.delete(display);
    else next.add(display);
    setHiddenValues(next);
  };

  const buildCondition = (): FilterCondition | null => {
    switch (condType) {
      case 'none':
        return null;
      case 'empty':
      case 'notEmpty':
        return { type: condType };
      case 'between':
        return { type: 'between', min: v1, max: v2 };
      default:
        return { type: condType, value: v1 };
    }
  };

  const apply = () => {
    if (mode === 'values') {
      setColumnFilter(col, hiddenValues.size > 0 ? { kind: 'values', hiddenValues: [...hiddenValues] } : null);
    } else {
      const condition = buildCondition();
      setColumnFilter(col, condition ? { kind: 'condition', condition } : null);
    }
    onClose();
  };

  const left = Math.min(x, window.innerWidth - 300);
  const top = Math.min(y, window.innerHeight - 460);
  const itemButton: React.CSSProperties = {
    border: 'none',
    background: 'transparent',
    textAlign: 'left',
    padding: '5px 4px',
    cursor: 'pointer',
    justifyContent: 'flex-start',
  };
  const tab = (active: boolean): React.CSSProperties => ({
    flex: 1,
    padding: '4px',
    border: '1px solid var(--border-color)',
    background: active ? 'var(--color-primary)' : 'transparent',
    color: active ? 'var(--fg-on-primary)' : 'var(--fg-main)',
    cursor: 'pointer',
  });

  return (
    <>
      {/* คลิกนอกเมนูเพื่อปิด */}
      <div style={{ position: 'fixed', inset: 0, zIndex: 1999 }} onMouseDown={onClose} />
      <div
        role="dialog"
        aria-label={th.filter.headerButtonTooltip}
        onKeyDown={(e) => {
          if (e.key === 'Escape') onClose();
        }}
        style={{
          position: 'fixed',
          left: Math.max(4, left),
          top: Math.max(4, top),
          width: '280px',
          zIndex: 2000,
          backgroundColor: 'var(--bg-app)',
          border: '1px solid var(--border-color)',
          borderRadius: '6px',
          boxShadow: 'var(--shadow-menu)',
          padding: '8px',
          fontSize: '13px',
          display: 'flex',
          flexDirection: 'column',
          gap: '6px',
        }}
      >
        <button style={itemButton} onClick={() => { onClose(); sortFilterColumn(col, 'asc'); }}>
          {th.filter.sortAsc}
        </button>
        <button style={itemButton} onClick={() => { onClose(); sortFilterColumn(col, 'desc'); }}>
          {th.filter.sortDesc}
        </button>
        <div style={{ height: '1px', background: 'var(--border-color)' }} />

        <div style={{ display: 'flex' }}>
          <button style={tab(mode === 'values')} onClick={() => setMode('values')}>
            {th.filter.byValues}
          </button>
          <button style={tab(mode === 'condition')} onClick={() => setMode('condition')}>
            {th.filter.byCondition}
          </button>
        </div>

        {mode === 'values' ? (
          <>
            <input
              autoFocus
              value={search}
              placeholder={th.filter.searchValues}
              onChange={(e) => setSearch(e.target.value)}
              style={{ padding: '4px 6px', border: '1px solid var(--control-border)', borderRadius: '4px' }}
            />
            <div style={{ display: 'flex', gap: '8px' }}>
              <button
                style={{ ...itemButton, color: 'var(--color-primary)' }}
                onClick={() => {
                  const next = new Set(hiddenValues);
                  shownValues.forEach((v) => next.delete(v.display));
                  setHiddenValues(next);
                }}
              >
                {th.filter.selectAll}
              </button>
              <button
                style={{ ...itemButton, color: 'var(--color-primary)' }}
                onClick={() => {
                  const next = new Set(hiddenValues);
                  shownValues.forEach((v) => next.add(v.display));
                  setHiddenValues(next);
                }}
              >
                {th.filter.clearAll}
              </button>
            </div>
            <div style={{ maxHeight: '200px', overflowY: 'auto', border: '1px solid var(--border-color)', borderRadius: '4px', padding: '4px' }}>
              {shownValues.map((v) => (
                <label key={v.display} style={{ display: 'flex', alignItems: 'center', gap: '6px', padding: '2px 0', cursor: 'pointer' }}>
                  <input
                    type="checkbox"
                    checked={!hiddenValues.has(v.display)}
                    onChange={() => toggleValue(v.display)}
                  />
                  <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                    {th.filter.valueCount
                      .replace('{value}', v.display || th.filter.blankValue)
                      .replace('{count}', String(v.count))}
                  </span>
                </label>
              ))}
            </div>
          </>
        ) : (
          <>
            <select value={condType} onChange={(e) => setCondType(e.target.value as ConditionType)}>
              {(Object.keys(CONDITION_LABELS) as ConditionType[]).map((t) => (
                <option key={t} value={t}>
                  {CONDITION_LABELS[t]}
                </option>
              ))}
            </select>
            {condType !== 'none' && condType !== 'empty' && condType !== 'notEmpty' && (
              <input
                autoFocus
                value={v1}
                placeholder={condType === 'between' ? th.filter.conditionMin : th.filter.conditionValue}
                onChange={(e) => setV1(e.target.value)}
                style={{ padding: '4px 6px', border: '1px solid var(--control-border)', borderRadius: '4px' }}
              />
            )}
            {condType === 'between' && (
              <input
                value={v2}
                placeholder={th.filter.conditionMax}
                onChange={(e) => setV2(e.target.value)}
                style={{ padding: '4px 6px', border: '1px solid var(--control-border)', borderRadius: '4px' }}
              />
            )}
            {NUMERIC_CONDITIONS.includes(condType) && (
              <span style={{ color: 'var(--fg-muted)', fontSize: '12px' }}>{th.filter.numberHint}</span>
            )}
          </>
        )}

        <div style={{ height: '1px', background: 'var(--border-color)' }} />
        <div style={{ display: 'flex', flexWrap: 'wrap', gap: '6px', justifyContent: 'space-between' }}>
          <button
            style={{ ...itemButton, color: 'var(--color-primary)' }}
            onClick={() => {
              onClose();
              reapplyFilter();
            }}
          >
            ↻ {th.filter.reapply}
          </button>
          {existing && (
            <button
              style={{ ...itemButton, color: 'var(--color-danger)' }}
              onClick={() => {
                onClose();
                setColumnFilter(col, null);
              }}
            >
              {th.filter.clearColumn}
            </button>
          )}
        </div>
        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '6px' }}>
          <button onClick={onClose}>{th.filter.cancel}</button>
          <button onClick={apply} style={{ fontWeight: 600 }}>
            {th.filter.apply}
          </button>
        </div>
      </div>
    </>
  );
};
