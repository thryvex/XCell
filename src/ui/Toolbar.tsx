import React, { useState } from 'react';
import { th } from '../i18n/th';
import { useSheetStore } from './store';
import { NumFmt } from '../core/model';

interface ToolbarProps {
  onNew: () => void;
  onOpen: () => void;
  onSave: () => void;
  onSaveAs: () => void;
}

const PRESET_COLORS = [
  '#000000', '#434343', '#666666', '#999999', '#cccccc', '#ffffff',
  '#ef4444', '#f97316', '#f59e0b', '#84cc16', '#10b981', '#06b6d4',
  '#3b82f6', '#6366f1', '#8b5cf6', '#d946ef', '#fecaca', '#fed7aa',
  '#bbf7d0', '#bfdbfe',
];

export const Toolbar: React.FC<ToolbarProps> = ({ onNew, onOpen, onSave, onSaveAs }) => {
  const {
    undo,
    redo,
    history,
    mergeCells,
    unmergeCells,
    setNumFmt,
    applySigma,
    activeCell,
    insertRowAt,
    deleteRowAt,
    insertColAt,
    deleteColAt,
    addRowsAtBottom,
    getActiveCellStyle,
    toggleBold,
    toggleItalic,
    toggleUnderline,
    toggleWrap,
    setAlignment,
    setTextColor,
    setFillColor,
    setBorder,
    openFormulaHelper,
  } = useSheetStore();

  const [sigmaMenuOpen, setSigmaMenuOpen] = useState(false);
  const [rowColMenuOpen, setRowColMenuOpen] = useState(false);
  const [textColorOpen, setTextColorOpen] = useState(false);
  const [fillColorOpen, setFillColorOpen] = useState(false);
  const [borderMenuOpen, setBorderMenuOpen] = useState(false);

  const activeStyle = getActiveCellStyle();

  const handleFmtChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    const val = e.target.value as NumFmt;
    if (val) {
      setNumFmt(val);
      e.target.value = '';
    }
  };

  const handleSigmaSelect = (fn: string) => {
    applySigma(fn);
    setSigmaMenuOpen(false);
  };

  const closeAllMenus = () => {
    setSigmaMenuOpen(false);
    setRowColMenuOpen(false);
    setTextColorOpen(false);
    setFillColorOpen(false);
    setBorderMenuOpen(false);
  };

  const activeBtnStyle = (isActive?: boolean): React.CSSProperties => ({
    backgroundColor: isActive ? 'var(--color-primary)' : 'transparent',
    color: isActive ? '#ffffff' : 'var(--fg-main)',
    borderColor: isActive ? 'var(--color-primary)' : 'var(--border-color)',
    fontWeight: isActive ? 600 : 400,
  });

  return (
    <div
      style={{
        display: 'flex',
        alignItems: 'center',
        gap: '6px',
        padding: '6px 12px',
        borderBottom: '1px solid var(--border-color)',
        backgroundColor: 'var(--bg-surface)',
        flexWrap: 'wrap',
      }}
    >
      {/* File group */}
      <button onClick={onNew} title={th.toolbar.newFileTooltip}>
        📄 {th.toolbar.newFile}
      </button>
      <button onClick={onOpen} title={th.toolbar.openFileTooltip}>
        📂 {th.toolbar.openFile}
      </button>
      <button className="primary" onClick={onSave} title={th.toolbar.saveFileTooltip}>
        💾 {th.toolbar.saveFile}
      </button>
      <button onClick={onSaveAs} title={th.toolbar.saveAsFileTooltip}>
        {th.toolbar.saveAsFile}
      </button>

      <span style={{ width: '1px', height: '20px', background: 'var(--border-color)', margin: '0 4px' }} />

      {/* History group */}
      <button onClick={undo} disabled={!history.canUndo()} title={th.toolbar.undoTooltip}>
        ↶ {th.toolbar.undo}
      </button>
      <button onClick={redo} disabled={!history.canRedo()} title={th.toolbar.redoTooltip}>
        ↷ {th.toolbar.redo}
      </button>

      <span style={{ width: '1px', height: '20px', background: 'var(--border-color)', margin: '0 4px' }} />

      {/* Text formatting: Bold, Italic, Underline */}
      <button
        style={{ ...activeBtnStyle(activeStyle.bold), fontWeight: 'bold', width: '32px' }}
        onClick={toggleBold}
        title={th.toolbar.bold}
      >
        B
      </button>
      <button
        style={{ ...activeBtnStyle(activeStyle.italic), fontStyle: 'italic', width: '32px' }}
        onClick={toggleItalic}
        title={th.toolbar.italic}
      >
        I
      </button>
      <button
        style={{ ...activeBtnStyle(activeStyle.underline), textDecoration: 'underline', width: '32px' }}
        onClick={toggleUnderline}
        title={th.toolbar.underline}
      >
        U
      </button>

      {/* Text color dropdown */}
      <div style={{ position: 'relative', display: 'inline-block' }}>
        <button
          onClick={() => {
            const next = !textColorOpen;
            closeAllMenus();
            setTextColorOpen(next);
          }}
          title={th.toolbar.textColor}
          style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', minWidth: '34px', padding: '2px 6px' }}
        >
          <span style={{ fontWeight: 600, fontSize: '13px' }}>A</span>
          <span
            style={{
              width: '18px',
              height: '3px',
              backgroundColor: activeStyle.color || 'var(--fg-main)',
              marginTop: '1px',
              borderRadius: '1px',
            }}
          />
        </button>

        {textColorOpen && (
          <div
            style={{
              position: 'absolute',
              top: '100%',
              left: 0,
              zIndex: 100,
              backgroundColor: 'var(--bg-app)',
              border: '1px solid var(--border-color)',
              borderRadius: '6px',
              boxShadow: '0 4px 12px rgba(0,0,0,0.15)',
              padding: '8px',
              minWidth: '180px',
              marginTop: '4px',
            }}
          >
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(5, 1fr)', gap: '6px', marginBottom: '8px' }}>
              {PRESET_COLORS.map((hex) => (
                <div
                  key={hex}
                  onClick={() => {
                    setTextColor(hex);
                    setTextColorOpen(false);
                  }}
                  title={hex}
                  style={{
                    width: '24px',
                    height: '24px',
                    backgroundColor: hex,
                    borderRadius: '4px',
                    border: activeStyle.color === hex ? '2px solid var(--color-primary)' : '1px solid var(--border-color)',
                    cursor: 'pointer',
                  }}
                />
              ))}
            </div>
            <div style={{ display: 'flex', alignItems: 'center', gap: '6px', marginBottom: '6px' }}>
              <label style={{ fontSize: '12px', flex: 1 }}>{th.toolbar.customColor}:</label>
              <input
                type="color"
                value={activeStyle.color || '#000000'}
                onChange={(e) => {
                  setTextColor(e.target.value);
                  setTextColorOpen(false);
                }}
                style={{ width: '28px', height: '24px', padding: 0, border: 'none', cursor: 'pointer' }}
              />
            </div>
            <button
              onClick={() => {
                setTextColor(undefined);
                setTextColorOpen(false);
              }}
              style={{ width: '100%', fontSize: '12px', padding: '4px' }}
            >
              {th.toolbar.colorDefault}
            </button>
          </div>
        )}
      </div>

      {/* Fill color dropdown */}
      <div style={{ position: 'relative', display: 'inline-block' }}>
        <button
          onClick={() => {
            const next = !fillColorOpen;
            closeAllMenus();
            setFillColorOpen(next);
          }}
          title={th.toolbar.fillColor}
          style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', minWidth: '34px', padding: '2px 6px' }}
        >
          <span style={{ fontSize: '13px' }}>🎨</span>
          <span
            style={{
              width: '18px',
              height: '3px',
              backgroundColor: activeStyle.fill || 'transparent',
              border: activeStyle.fill ? 'none' : '1px dashed var(--border-color)',
              marginTop: '1px',
              borderRadius: '1px',
            }}
          />
        </button>

        {fillColorOpen && (
          <div
            style={{
              position: 'absolute',
              top: '100%',
              left: 0,
              zIndex: 100,
              backgroundColor: 'var(--bg-app)',
              border: '1px solid var(--border-color)',
              borderRadius: '6px',
              boxShadow: '0 4px 12px rgba(0,0,0,0.15)',
              padding: '8px',
              minWidth: '180px',
              marginTop: '4px',
            }}
          >
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(5, 1fr)', gap: '6px', marginBottom: '8px' }}>
              {PRESET_COLORS.map((hex) => (
                <div
                  key={hex}
                  onClick={() => {
                    setFillColor(hex);
                    setFillColorOpen(false);
                  }}
                  title={hex}
                  style={{
                    width: '24px',
                    height: '24px',
                    backgroundColor: hex,
                    borderRadius: '4px',
                    border: activeStyle.fill === hex ? '2px solid var(--color-primary)' : '1px solid var(--border-color)',
                    cursor: 'pointer',
                  }}
                />
              ))}
            </div>
            <div style={{ display: 'flex', alignItems: 'center', gap: '6px', marginBottom: '6px' }}>
              <label style={{ fontSize: '12px', flex: 1 }}>{th.toolbar.customColor}:</label>
              <input
                type="color"
                value={activeStyle.fill || '#ffffff'}
                onChange={(e) => {
                  setFillColor(e.target.value);
                  setFillColorOpen(false);
                }}
                style={{ width: '28px', height: '24px', padding: 0, border: 'none', cursor: 'pointer' }}
              />
            </div>
            <button
              onClick={() => {
                setFillColor(undefined);
                setFillColorOpen(false);
              }}
              style={{ width: '100%', fontSize: '12px', padding: '4px' }}
            >
              {th.toolbar.colorDefault}
            </button>
          </div>
        )}
      </div>

      <span style={{ width: '1px', height: '20px', background: 'var(--border-color)', margin: '0 4px' }} />

      {/* Alignment group: Left, Center, Right */}
      <button
        style={{ ...activeBtnStyle(activeStyle.align === 'left'), width: '32px' }}
        onClick={() => setAlignment('left')}
        title={th.toolbar.alignLeft}
      >
        ≡⇤
      </button>
      <button
        style={{ ...activeBtnStyle(activeStyle.align === 'center'), width: '32px' }}
        onClick={() => setAlignment('center')}
        title={th.toolbar.alignCenter}
      >
        ≡
      </button>
      <button
        style={{ ...activeBtnStyle(activeStyle.align === 'right'), width: '32px' }}
        onClick={() => setAlignment('right')}
        title={th.toolbar.alignRight}
      >
        ⇥≡
      </button>

      {/* Text wrap toggle */}
      <button
        style={activeBtnStyle(activeStyle.wrap)}
        onClick={toggleWrap}
        title={th.toolbar.wrapText}
      >
        ↵ {th.toolbar.wrapText}
      </button>

      {/* Border menu */}
      <div style={{ position: 'relative', display: 'inline-block' }}>
        <button
          onClick={() => {
            const next = !borderMenuOpen;
            closeAllMenus();
            setBorderMenuOpen(next);
          }}
          title={th.toolbar.borders}
          style={activeBtnStyle(activeStyle.border && activeStyle.border !== 'none')}
        >
          田 {th.toolbar.borders} ▾
        </button>

        {borderMenuOpen && (
          <div
            style={{
              position: 'absolute',
              top: '100%',
              left: 0,
              zIndex: 100,
              backgroundColor: 'var(--bg-app)',
              border: '1px solid var(--border-color)',
              borderRadius: '6px',
              boxShadow: '0 4px 12px rgba(0,0,0,0.15)',
              display: 'flex',
              flexDirection: 'column',
              minWidth: '150px',
              marginTop: '4px',
              padding: '4px 0',
            }}
          >
            <button
              style={{
                border: 'none',
                background: 'transparent',
                justifyContent: 'flex-start',
                padding: '6px 12px',
                fontWeight: !activeStyle.border || activeStyle.border === 'none' ? 600 : 400,
              }}
              onClick={() => {
                setBorder('none');
                setBorderMenuOpen(false);
              }}
            >
              ◻ {th.toolbar.borderNone}
            </button>
            <button
              style={{
                border: 'none',
                background: 'transparent',
                justifyContent: 'flex-start',
                padding: '6px 12px',
                fontWeight: activeStyle.border === 'all' ? 600 : 400,
              }}
              onClick={() => {
                setBorder('all');
                setBorderMenuOpen(false);
              }}
            >
              田 {th.toolbar.borderAll}
            </button>
            <button
              style={{
                border: 'none',
                background: 'transparent',
                justifyContent: 'flex-start',
                padding: '6px 12px',
                fontWeight: activeStyle.border === 'outer' ? 600 : 400,
              }}
              onClick={() => {
                setBorder('outer');
                setBorderMenuOpen(false);
              }}
            >
              回 {th.toolbar.borderOuter}
            </button>
          </div>
        )}
      </div>

      <span style={{ width: '1px', height: '20px', background: 'var(--border-color)', margin: '0 4px' }} />

      {/* Merge group */}
      <button onClick={() => mergeCells()} title={th.toolbar.mergeCellsTooltip}>
        ⊞ {th.toolbar.mergeCells}
      </button>
      <button onClick={() => unmergeCells()} title={th.toolbar.unmergeCellsTooltip}>
        ⊟ {th.toolbar.unmergeCells}
      </button>

      <span style={{ width: '1px', height: '20px', background: 'var(--border-color)', margin: '0 4px' }} />

      {/* Number Formats group */}
      <select onChange={handleFmtChange} defaultValue="" aria-label={th.toolbar.numFmt}>
        <option value="" disabled>
          {th.toolbar.numFmt}
        </option>
        <option value="general">{th.toolbar.fmtGeneral}</option>
        <option value="number2">{th.toolbar.fmtNumber2}</option>
        <option value="integer">{th.toolbar.fmtInteger}</option>
        <option value="baht">{th.toolbar.fmtBaht}</option>
        <option value="percent">{th.toolbar.fmtPercent}</option>
        <option value="date">{th.toolbar.fmtDate}</option>
      </select>

      <button onClick={() => setNumFmt('baht')} title="จัดรูปแบบ ฿ บาท">
        ฿
      </button>
      <button onClick={() => setNumFmt('percent')} title="จัดรูปแบบ %">
        %
      </button>

      <span style={{ width: '1px', height: '20px', background: 'var(--border-color)', margin: '0 4px' }} />

      {/* Row and Column Management */}
      <div style={{ position: 'relative', display: 'inline-block' }}>
        <button
          onClick={() => {
            const next = !rowColMenuOpen;
            closeAllMenus();
            setRowColMenuOpen(next);
          }}
          title="จัดการแถวและคอลัมน์"
        >
          ☵ แถว/คอลัมน์ ▾
        </button>
        {rowColMenuOpen && (
          <div
            style={{
              position: 'absolute',
              top: '100%',
              left: 0,
              zIndex: 100,
              backgroundColor: 'var(--bg-app)',
              border: '1px solid var(--border-color)',
              borderRadius: '6px',
              boxShadow: '0 4px 12px rgba(0,0,0,0.15)',
              display: 'flex',
              flexDirection: 'column',
              minWidth: '180px',
              marginTop: '4px',
              padding: '4px 0',
            }}
          >
            <button
              style={{ border: 'none', background: 'transparent', justifyContent: 'flex-start', padding: '6px 12px' }}
              onClick={() => {
                insertRowAt(activeCell.r);
                setRowColMenuOpen(false);
              }}
            >
              ＋ แทรกแถวด้านบน
            </button>
            <button
              style={{ border: 'none', background: 'transparent', justifyContent: 'flex-start', padding: '6px 12px' }}
              onClick={() => {
                insertRowAt(activeCell.r + 1);
                setRowColMenuOpen(false);
              }}
            >
              ＋ แทรกแถวด้านล่าง
            </button>
            <button
              style={{ border: 'none', background: 'transparent', justifyContent: 'flex-start', padding: '6px 12px', color: 'var(--color-danger)' }}
              onClick={() => {
                deleteRowAt(activeCell.r);
                setRowColMenuOpen(false);
              }}
            >
              🗑️ ลบแถวปัจจุบัน
            </button>
            <span style={{ height: '1px', background: 'var(--border-color)', margin: '4px 0' }} />
            <button
              style={{ border: 'none', background: 'transparent', justifyContent: 'flex-start', padding: '6px 12px' }}
              onClick={() => {
                insertColAt(activeCell.c);
                setRowColMenuOpen(false);
              }}
            >
              ＋ แทรกคอลัมน์ทางซ้าย
            </button>
            <button
              style={{ border: 'none', background: 'transparent', justifyContent: 'flex-start', padding: '6px 12px' }}
              onClick={() => {
                insertColAt(activeCell.c + 1);
                setRowColMenuOpen(false);
              }}
            >
              ＋ แทรกคอลัมน์ทางขวา
            </button>
            <button
              style={{ border: 'none', background: 'transparent', justifyContent: 'flex-start', padding: '6px 12px', color: 'var(--color-danger)' }}
              onClick={() => {
                deleteColAt(activeCell.c);
                setRowColMenuOpen(false);
              }}
            >
              🗑️ ลบคอลัมน์ปัจจุบัน
            </button>
            <span style={{ height: '1px', background: 'var(--border-color)', margin: '4px 0' }} />
            <button
              style={{ border: 'none', background: 'transparent', justifyContent: 'flex-start', padding: '6px 12px', fontWeight: 600 }}
              onClick={() => {
                addRowsAtBottom(50);
                setRowColMenuOpen(false);
              }}
            >
              ＋ เพิ่ม 50 แถวล่างสุด
            </button>
          </div>
        )}
      </div>

      <span style={{ width: '1px', height: '20px', background: 'var(--border-color)', margin: '0 4px' }} />

      {/* Sigma menu */}
      <div style={{ position: 'relative', display: 'inline-block' }}>
        <button
          onClick={() => {
            const next = !sigmaMenuOpen;
            closeAllMenus();
            setSigmaMenuOpen(next);
          }}
          title={th.toolbar.sigmaTooltip}
          style={{ fontWeight: 600 }}
        >
          Σ {th.toolbar.sigma} ▾
        </button>
        {sigmaMenuOpen && (
          <div
            style={{
              position: 'absolute',
              top: '100%',
              left: 0,
              zIndex: 100,
              backgroundColor: 'var(--bg-app)',
              border: '1px solid var(--border-color)',
              borderRadius: '6px',
              boxShadow: '0 4px 12px rgba(0,0,0,0.15)',
              display: 'flex',
              flexDirection: 'column',
              minWidth: '160px',
              marginTop: '4px',
              padding: '4px 0',
            }}
          >
            <button
              style={{ border: 'none', background: 'transparent', justifyContent: 'flex-start', padding: '6px 12px' }}
              onClick={() => handleSigmaSelect('SUM')}
            >
              {th.toolbar.sum}
            </button>
            <button
              style={{ border: 'none', background: 'transparent', justifyContent: 'flex-start', padding: '6px 12px' }}
              onClick={() => handleSigmaSelect('AVERAGE')}
            >
              {th.toolbar.average}
            </button>
            <button
              style={{ border: 'none', background: 'transparent', justifyContent: 'flex-start', padding: '6px 12px' }}
              onClick={() => handleSigmaSelect('COUNT')}
            >
              {th.toolbar.count}
            </button>
            <button
              style={{ border: 'none', background: 'transparent', justifyContent: 'flex-start', padding: '6px 12px' }}
              onClick={() => handleSigmaSelect('MIN')}
            >
              {th.toolbar.min}
            </button>
            <button
              style={{ border: 'none', background: 'transparent', justifyContent: 'flex-start', padding: '6px 12px' }}
              onClick={() => handleSigmaSelect('MAX')}
            >
              {th.toolbar.max}
            </button>
            <button
              style={{ border: 'none', background: 'transparent', justifyContent: 'flex-start', padding: '6px 12px' }}
              onClick={() => handleSigmaSelect('MEDIAN')}
            >
              {th.toolbar.median}
            </button>
            <span style={{ height: '1px', background: 'var(--border-color)', margin: '4px 0' }} />
            <button
              style={{
                border: 'none',
                background: 'transparent',
                justifyContent: 'flex-start',
                padding: '6px 12px',
                fontWeight: 600,
                color: 'var(--color-primary)',
              }}
              onClick={() => {
                openFormulaHelper();
                setSigmaMenuOpen(false);
              }}
            >
              ƒx {th.toolbar.otherFunctions}
            </button>
          </div>
        )}
      </div>
    </div>
  );
};
