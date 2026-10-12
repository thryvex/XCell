import React from 'react';
import { th } from '../i18n/th';
import { useSheetStore } from './store';
import { cellToA1 } from '../core/address';

export const FormulaBar: React.FC = () => {
  const {
    activeCell,
    isEditing,
    editingValue,
    startEditing,
    setEditingValue,
    commitEditing,
    cancelEditing,
    getRawValue,
    openFormulaHelper,
  } = useSheetStore();

  const activeA1 = cellToA1(activeCell.r, activeCell.c);
  const rawValue = getRawValue(activeCell.r, activeCell.c);
  const displayVal = isEditing ? editingValue : rawValue;

  const handleInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (!isEditing) {
      startEditing(e.target.value);
    } else {
      setEditingValue(e.target.value);
    }
  };

  const handleKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    const isCtrl = e.ctrlKey || e.metaKey;
    if (isCtrl && (e.key === ';' || e.code === 'Semicolon')) {
      e.preventDefault();
      const now = new Date();
      const day = String(now.getDate()).padStart(2, '0');
      const month = String(now.getMonth() + 1).padStart(2, '0');
      const year = now.getFullYear();
      const dateStr = `${day}/${month}/${year}`;

      const input = e.currentTarget;
      const start = input.selectionStart || 0;
      const end = input.selectionEnd || 0;
      const val = input.value;
      const nextVal = val.slice(0, start) + dateStr + val.slice(end);
      if (!isEditing) {
        startEditing(nextVal);
      } else {
        setEditingValue(nextVal);
      }
      setTimeout(() => {
        input.selectionStart = input.selectionEnd = start + dateStr.length;
      }, 0);
      return;
    }

    if (e.key === 'Enter') {
      e.preventDefault();
      commitEditing(1, 0);
    } else if (e.key === 'Escape') {
      e.preventDefault();
      cancelEditing();
    }
  };

  return (
    <div
      style={{
        display: 'flex',
        alignItems: 'center',
        gap: '8px',
        padding: '6px 12px',
        borderBottom: '1px solid var(--border-color)',
        backgroundColor: 'var(--bg-app)',
      }}
    >
      <input
        value={activeA1}
        readOnly
        style={{
          width: '65px',
          textAlign: 'center',
          fontWeight: 600,
          backgroundColor: 'var(--bg-header)',
          border: '1px solid var(--border-color)',
          borderRadius: '4px',
          padding: '4px 6px',
        }}
      />
      <button
        type="button"
        onClick={openFormulaHelper}
        title={th.formulaHelper.title}
        style={{
          border: 'none',
          background: 'transparent',
          color: 'var(--color-primary)',
          fontWeight: 700,
          fontStyle: 'italic',
          fontSize: '16px',
          cursor: 'pointer',
          padding: '2px 6px',
          borderRadius: '4px',
        }}
      >
        {th.formulaBar.fx}
      </button>
      <input
        value={displayVal}
        onChange={handleInputChange}
        onKeyDown={handleKeyDown}
        placeholder={th.formulaBar.placeholder}
        style={{
          flex: 1,
          border: '1px solid var(--border-color)',
          borderRadius: '4px',
          padding: '4px 8px',
          fontSize: '14px',
        }}
      />
    </div>
  );
};
