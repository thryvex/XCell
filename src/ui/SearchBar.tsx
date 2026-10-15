import React, { useEffect, useRef, useState } from 'react';
import { useSheetStore } from './store';
import { th } from '../i18n/th';
import { SearchOptions } from '../core/search';

/**
 * แถบค้นหามุมขวาบนของตาราง (Ctrl+F)
 * Enter / F3 = ถัดไป, Shift+Enter / Shift+F3 = ก่อนหน้า, Esc = ปิด
 */
export const SearchBar: React.FC = () => {
  const { search, closeSearch, setSearchQuery, setSearchOption, searchNext, searchPrev } =
    useSheetStore();
  const inputRef = useRef<HTMLInputElement>(null);
  const [showOptions, setShowOptions] = useState(false);

  useEffect(() => {
    if (search.isOpen && inputRef.current) {
      inputRef.current.focus();
      inputRef.current.select();
    }
  }, [search.isOpen]);

  if (!search.isOpen) return null;

  const total = search.matches.length;
  let status = '';
  if (search.query) {
    if (total === 0) status = th.search.notFound;
    else if (search.current >= 0)
      status = th.search.resultCount
        .replace('{current}', String(search.current + 1))
        .replace('{total}', String(total));
    else status = th.search.resultCountNoCurrent.replace('{total}', String(total));
  }

  const handleKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Enter' || e.key === 'F3') {
      e.preventDefault();
      if (e.shiftKey) searchPrev();
      else searchNext();
    } else if (e.key === 'Escape') {
      e.preventDefault();
      closeSearch();
    }
  };

  const optionKeys: (keyof SearchOptions)[] = ['matchCase', 'wholeCell', 'includeFormulas', 'allSheets'];

  const iconButton: React.CSSProperties = {
    border: 'none',
    background: 'transparent',
    padding: '2px 6px',
    cursor: 'pointer',
    fontSize: '13px',
    color: 'var(--fg-main)',
  };

  return (
    <div
      role="search"
      style={{
        position: 'absolute',
        top: '6px',
        right: '22px',
        zIndex: 50,
        backgroundColor: 'var(--bg-app)',
        border: '1px solid var(--border-color)',
        borderRadius: '6px',
        boxShadow: '0 4px 14px rgba(0,0,0,0.15)',
        padding: '6px 8px',
        fontSize: '13px',
        minWidth: '320px',
      }}
    >
      <div style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
        <input
          ref={inputRef}
          value={search.query}
          placeholder={th.search.placeholder}
          onChange={(e) => setSearchQuery(e.target.value)}
          onKeyDown={handleKeyDown}
          style={{
            flex: 1,
            border: '1px solid var(--border-color)',
            borderRadius: '4px',
            padding: '4px 6px',
            fontSize: '13px',
            minWidth: 0,
          }}
        />
        <span
          aria-live="polite"
          style={{
            minWidth: '64px',
            textAlign: 'center',
            color: total === 0 && search.query ? 'var(--color-danger)' : 'var(--fg-muted)',
            whiteSpace: 'nowrap',
          }}
        >
          {status}
        </span>
        <button style={iconButton} title={th.search.prev} onClick={searchPrev} disabled={total === 0}>
          ▲
        </button>
        <button style={iconButton} title={th.search.next} onClick={searchNext} disabled={total === 0}>
          ▼
        </button>
        <button
          style={{ ...iconButton, fontWeight: showOptions ? 700 : 400 }}
          title={th.search.options}
          onClick={() => setShowOptions((v) => !v)}
        >
          ⚙
        </button>
        <button style={iconButton} title={th.search.close} onClick={closeSearch}>
          ✕
        </button>
      </div>
      {showOptions && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: '4px', marginTop: '6px' }}>
          {optionKeys.map((key) => (
            <label key={key} style={{ display: 'flex', alignItems: 'center', gap: '6px', cursor: 'pointer' }}>
              <input
                type="checkbox"
                checked={search.options[key]}
                onChange={(e) => setSearchOption(key, e.target.checked)}
              />
              {th.search[key]}
            </label>
          ))}
        </div>
      )}
    </div>
  );
};
