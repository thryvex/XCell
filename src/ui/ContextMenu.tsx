import React, { useEffect } from 'react';

export interface ContextMenuProps {
  x: number;
  y: number;
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
  onMerge: () => void;
  onUnmerge: () => void;
  onFreezeRow: () => void;
  onFreezeCol: () => void;
  onUnfreeze: () => void;
}

export const ContextMenu: React.FC<ContextMenuProps> = ({
  x,
  y,
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
  onMerge,
  onUnmerge,
  onFreezeRow,
  onFreezeCol,
  onUnfreeze,
}) => {
  useEffect(() => {
    const handleOutsideClick = (e: MouseEvent) => {
      onClose();
    };
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose();
    };

    window.addEventListener('click', handleOutsideClick);
    window.addEventListener('contextmenu', handleOutsideClick);
    window.addEventListener('keydown', handleKeyDown);

    return () => {
      window.removeEventListener('click', handleOutsideClick);
      window.removeEventListener('contextmenu', handleOutsideClick);
      window.removeEventListener('keydown', handleKeyDown);
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

  // Prevent menu overflow off right or bottom of window
  const menuWidth = 240;
  const menuHeight = 440;
  const posX = Math.min(x, window.innerWidth - menuWidth - 10);
  const posY = Math.min(y, window.innerHeight - menuHeight - 10);

  return (
    <div
      onClick={(e) => e.stopPropagation()}
      style={{
        position: 'fixed',
        left: `${posX}px`,
        top: `${posY}px`,
        zIndex: 2000,
        backgroundColor: 'var(--bg-app)',
        border: '1px solid var(--border-color)',
        borderRadius: '6px',
        boxShadow: '0 6px 18px rgba(0, 0, 0, 0.15)',
        width: `${menuWidth}px`,
        padding: '4px 0',
        display: 'flex',
        flexDirection: 'column',
      }}
    >
      <div
        style={menuItemStyle}
        onClick={() => {
          onCopy();
          onClose();
        }}
        onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = 'var(--bg-selection)')}
        onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'transparent')}
      >
        <span>คัดลอก</span>
        <span style={shortcutStyle}>Ctrl+C</span>
      </div>

      <div
        style={menuItemStyle}
        onClick={() => {
          onCut();
          onClose();
        }}
        onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = 'var(--bg-selection)')}
        onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'transparent')}
      >
        <span>ตัด</span>
        <span style={shortcutStyle}>Ctrl+X</span>
      </div>

      <div
        style={menuItemStyle}
        onClick={() => {
          onPaste();
          onClose();
        }}
        onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = 'var(--bg-selection)')}
        onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'transparent')}
      >
        <span>วาง</span>
        <span style={shortcutStyle}>Ctrl+V</span>
      </div>

      <div
        style={menuItemStyle}
        onClick={() => {
          onPasteValuesOnly();
          onClose();
        }}
        onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = 'var(--bg-selection)')}
        onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'transparent')}
      >
        <span>วางเฉพาะค่า</span>
        <span style={shortcutStyle}>Ctrl+Shift+V</span>
      </div>

      <div style={dividerStyle} />

      <div
        style={menuItemStyle}
        onClick={() => {
          onInsertRowAbove();
          onClose();
        }}
        onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = 'var(--bg-selection)')}
        onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'transparent')}
      >
        <span>แทรก 1 แถวด้านบน</span>
      </div>

      <div
        style={menuItemStyle}
        onClick={() => {
          onInsertRowBelow();
          onClose();
        }}
        onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = 'var(--bg-selection)')}
        onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'transparent')}
      >
        <span>แทรก 1 แถวด้านล่าง</span>
      </div>

      <div
        style={menuItemStyle}
        onClick={() => {
          onInsertColLeft();
          onClose();
        }}
        onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = 'var(--bg-selection)')}
        onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'transparent')}
      >
        <span>แทรก 1 คอลัมน์ทางซ้าย</span>
      </div>

      <div
        style={menuItemStyle}
        onClick={() => {
          onInsertColRight();
          onClose();
        }}
        onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = 'var(--bg-selection)')}
        onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'transparent')}
      >
        <span>แทรก 1 คอลัมน์ทางขวา</span>
      </div>

      <div
        style={menuItemStyle}
        onClick={() => {
          onDeleteRow();
          onClose();
        }}
        onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = 'var(--bg-selection)')}
        onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'transparent')}
      >
        <span>ลบแถว</span>
      </div>

      <div
        style={menuItemStyle}
        onClick={() => {
          onDeleteCol();
          onClose();
        }}
        onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = 'var(--bg-selection)')}
        onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'transparent')}
      >
        <span>ลบคอลัมน์</span>
      </div>

      <div style={dividerStyle} />

      <div
        style={menuItemStyle}
        onClick={() => {
          onClearContent();
          onClose();
        }}
        onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = 'var(--bg-selection)')}
        onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'transparent')}
      >
        <span>ล้างเนื้อหา</span>
        <span style={shortcutStyle}>Delete</span>
      </div>

      <div
        style={menuItemStyle}
        onClick={() => {
          onMerge();
          onClose();
        }}
        onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = 'var(--bg-selection)')}
        onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'transparent')}
      >
        <span>ผสานเซล</span>
      </div>

      <div
        style={menuItemStyle}
        onClick={() => {
          onUnmerge();
          onClose();
        }}
        onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = 'var(--bg-selection)')}
        onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'transparent')}
      >
        <span>แยกเซล</span>
      </div>

      <div style={dividerStyle} />

      <div
        style={menuItemStyle}
        onClick={() => {
          onFreezeRow();
          onClose();
        }}
        onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = 'var(--bg-selection)')}
        onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'transparent')}
      >
        <span>ตรึงถึงแถวนี้</span>
      </div>

      <div
        style={menuItemStyle}
        onClick={() => {
          onFreezeCol();
          onClose();
        }}
        onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = 'var(--bg-selection)')}
        onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'transparent')}
      >
        <span>ตรึงถึงคอลัมน์นี้</span>
      </div>

      <div
        style={menuItemStyle}
        onClick={() => {
          onUnfreeze();
          onClose();
        }}
        onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = 'var(--bg-selection)')}
        onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'transparent')}
      >
        <span>ยกเลิกการตรึง</span>
      </div>
    </div>
  );
};
