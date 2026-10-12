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
      <span style={{ color: 'var(--fg-muted)', fontWeight: 600, fontStyle: 'italic', fontSize: '15px' }}>
        {th.formulaBar.fx}
      </span>
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
