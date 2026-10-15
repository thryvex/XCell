import React, { useEffect, useLayoutEffect, useRef, useState } from 'react';
import { th } from '../i18n/th';

export type ContextMenuTarget = 'cell' | 'row' | 'col';

export interface ContextMenuProps {
  x: number;
  y: number;
  targetType?: ContextMenuTarget;
  onClose: () => void;
  onCopy: () => void;
  onCut: () => void;
  onPaste: () => void;
  onPasteValuesOnly: () => void;
  onInsertRowAbove: () => void;
  onInsertRowBelow: () => void;
  onInsertColLeft: () => void;
  onInsertColRight: () => void;
  onDeleteRow: () => void;
  onDeleteCol: () => void;
  onClearContent: () => void;
  onClearFormat: () => void;
  onMerge: () => void;
  onUnmerge: () => void;
  onFreezeRow: () => void;
  onFreezeCol: () => void;
  onUnfreeze: () => void;
  onSortAsc?: () => void;
  onSortDesc?: () => void;
  onSortAdvanced?: () => void;
}

export const ContextMenu: React.FC<ContextMenuProps> = ({
  x,
  y,
  targetType = 'cell',
  onClose,
  onCopy,
  onCut,
  onPaste,
  onPasteValuesOnly,
  onInsertRowAbove,
  onInsertRowBelow,
  onInsertColLeft,
  onInsertColRight,
  onDeleteRow,
  onDeleteCol,
  onClearContent,
  onClearFormat,
  onMerge,
  onUnmerge,
  onFreezeRow,
  onFreezeCol,
  onUnfreeze,
  onSortAsc,
  onSortDesc,
  onSortAdvanced,
}) => {
  const menuRef = useRef<HTMLDivElement>(null);

  // Position calculation and viewport clamping
  const [menuRect, setMenuRect] = useState<{ width: number; height: number }>({
    width: 240,
    height: targetType === 'cell' ? 480 : 360,
  });

  useLayoutEffect(() => {
    if (menuRef.current) {
      setMenuRect({
        width: menuRef.current.offsetWidth || 240,
        height: menuRef.current.offsetHeight || (targetType === 'cell' ? 480 : 360),
      });
    }
  }, [targetType]);

  const margin = 8;
  const posX = Math.max(margin, Math.min(x, window.innerWidth - menuRect.width - margin));
  const posY = Math.max(margin, Math.min(y, window.innerHeight - menuRect.height - margin));

  useEffect(() => {
    const handleOutsidePointer = (e: MouseEvent | PointerEvent) => {
      if (menuRef.current && !menuRef.current.contains(e.target as Node)) {
        onClose();
      }
    };

    const handleOutsideContextMenu = (e: MouseEvent) => {
      if (menuRef.current && !menuRef.current.contains(e.target as Node)) {
        onClose();
      }
    };

    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.preventDefault();
        e.stopPropagation();
        onClose();
      }
    };

    const handleScroll = (e: Event) => {
      if (menuRef.current && menuRef.current.contains(e.target as Node)) {
        return;
      }
      onClose();
    };

    // Attach listeners after a short delay (40ms) so the event triggering contextmenu doesn't dismiss it
    const timer = setTimeout(() => {
      document.addEventListener('pointerdown', handleOutsidePointer);
      document.addEventListener('contextmenu', handleOutsideContextMenu);
      document.addEventListener('keydown', handleKeyDown);
      window.addEventListener('scroll', handleScroll, true);
    }, 40);

    return () => {
      clearTimeout(timer);
      document.removeEventListener('pointerdown', handleOutsidePointer);
      document.removeEventListener('contextmenu', handleOutsideContextMenu);
      document.removeEventListener('keydown', handleKeyDown);
      window.removeEventListener('scroll', handleScroll, true);
    };
  }, [onClose]);

  const menuItemStyle: React.CSSProperties = {
    padding: '6px 14px',
    cursor: 'pointer',
    fontSize: '13px',
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    gap: '20px',
    userSelect: 'none',
  };

  const shortcutStyle: React.CSSProperties = {
    color: 'var(--fg-muted)',
    fontSize: '12px',
  };

  const dividerStyle: React.CSSProperties = {
    height: '1px',
    backgroundColor: 'var(--border-color)',
    margin: '4px 0',
  };

  const renderItem = (label: string, shortcut: string | undefined, onClick: () => void, isDanger = false) => (
    <div
      style={{
        ...menuItemStyle,
        color: isDanger ? 'var(--color-danger)' : undefined,
      }}
      onClick={(e) => {
        e.stopPropagation();
        onClick();
        onClose();
      }}
      onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = 'var(--bg-selection)')}
      onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'transparent')}
    >
      <span>{label}</span>
      {shortcut && <span style={shortcutStyle}>{shortcut}</span>}
    </div>
  );

  return (
    <div
      ref={menuRef}
      onPointerDown={(e) => e.stopPropagation()}
      onMouseDown={(e) => e.stopPropagation()}
      onContextMenu={(e) => {
        e.preventDefault();
        e.stopPropagation();
      }}
      onClick={(e) => e.stopPropagation()}
      style={{
        position: 'fixed',
        left: `${posX}px`,
        top: `${posY}px`,
        zIndex: 2000,
        backgroundColor: 'var(--bg-app)',
        border: '1px solid var(--border-color)',
        borderRadius: '6px',
        boxShadow: 'var(--shadow-menu)',
        width: '240px',
        padding: '4px 0',
        display: 'flex',
        flexDirection: 'column',
      }}
    >
      {/* 1. Row Header Menu */}
      {targetType === 'row' && (
        <>
          {renderItem(th.contextMenu.insertRowAbove, undefined, onInsertRowAbove)}
          {renderItem(th.contextMenu.insertRowBelow, undefined, onInsertRowBelow)}
          {renderItem(th.contextMenu.deleteRow, undefined, onDeleteRow, true)}
          <div style={dividerStyle} />
          {renderItem(th.contextMenu.clearContent, 'Delete', onClearContent)}
          {renderItem(th.contextMenu.clearFormat, undefined, onClearFormat)}
          <div style={dividerStyle} />
          {renderItem(th.contextMenu.copy, 'Ctrl+C', onCopy)}
          {renderItem(th.contextMenu.cut, 'Ctrl+X', onCut)}
          {renderItem(th.contextMenu.paste, 'Ctrl+V', onPaste)}
          {renderItem(th.contextMenu.pasteValuesOnly, 'Ctrl+Shift+V', onPasteValuesOnly)}
          <div style={dividerStyle} />
          {renderItem(th.contextMenu.freezeRow, undefined, onFreezeRow)}
          {renderItem(th.contextMenu.unfreeze, undefined, onUnfreeze)}
        </>
      )}

      {/* 2. Column Header Menu */}
      {targetType === 'col' && (
        <>
          {renderItem(th.contextMenu.insertColLeft, undefined, onInsertColLeft)}
          {renderItem(th.contextMenu.insertColRight, undefined, onInsertColRight)}
          {renderItem(th.contextMenu.deleteCol, undefined, onDeleteCol, true)}
          <div style={dividerStyle} />
          {renderItem(th.contextMenu.clearContent, 'Delete', onClearContent)}
          {renderItem(th.contextMenu.clearFormat, undefined, onClearFormat)}
          <div style={dividerStyle} />
          {renderItem(th.contextMenu.copy, 'Ctrl+C', onCopy)}
          {renderItem(th.contextMenu.cut, 'Ctrl+X', onCut)}
          {renderItem(th.contextMenu.paste, 'Ctrl+V', onPaste)}
          {renderItem(th.contextMenu.pasteValuesOnly, 'Ctrl+Shift+V', onPasteValuesOnly)}
          <div style={dividerStyle} />
          {renderItem(th.contextMenu.freezeCol, undefined, onFreezeCol)}
          {renderItem(th.contextMenu.unfreeze, undefined, onUnfreeze)}
        </>
      )}

      {/* 3. Cell Menu (Full) */}
      {targetType === 'cell' && (
        <>
          {renderItem(th.contextMenu.copy, 'Ctrl+C', onCopy)}
          {renderItem(th.contextMenu.cut, 'Ctrl+X', onCut)}
          {renderItem(th.contextMenu.paste, 'Ctrl+V', onPaste)}
          {renderItem(th.contextMenu.pasteValuesOnly, 'Ctrl+Shift+V', onPasteValuesOnly)}
          <div style={dividerStyle} />
          {renderItem(th.contextMenu.insertRowAbove, undefined, onInsertRowAbove)}
          {renderItem(th.contextMenu.insertRowBelow, undefined, onInsertRowBelow)}
          {renderItem(th.contextMenu.insertColLeft, undefined, onInsertColLeft)}
          {renderItem(th.contextMenu.insertColRight, undefined, onInsertColRight)}
          {renderItem(th.contextMenu.deleteRow, undefined, onDeleteRow, true)}
          {renderItem(th.contextMenu.deleteCol, undefined, onDeleteCol, true)}
          <div style={dividerStyle} />
          {renderItem(th.contextMenu.clearContent, 'Delete', onClearContent)}
          {renderItem(th.contextMenu.clearFormat, undefined, onClearFormat)}
          {renderItem(th.contextMenu.merge, undefined, onMerge)}
          {renderItem(th.contextMenu.unmerge, undefined, onUnmerge)}
          {onSortAsc && onSortDesc && onSortAdvanced && (
            <>
              <div style={dividerStyle} />
              {renderItem(th.sort.asc, undefined, onSortAsc)}
              {renderItem(th.sort.desc, undefined, onSortDesc)}
              {renderItem(th.sort.advanced, undefined, onSortAdvanced)}
            </>
          )}
          <div style={dividerStyle} />
          {renderItem(th.contextMenu.freezeRow, undefined, onFreezeRow)}
          {renderItem(th.contextMenu.freezeCol, undefined, onFreezeCol)}
          {renderItem(th.contextMenu.unfreeze, undefined, onUnfreeze)}
        </>
      )}
    </div>
  );
};
