import React, { useState } from 'react';
import { th } from '../i18n/th';
import { IconAlignCenter, IconAlignLeft, IconAlignRight, IconCheck, IconChevronDown, IconFillColor } from './icons';

/** ช่องเครื่องหมายถูกหน้ารายการในเมนูตั้งค่า (กว้างเท่ากันทุกแถว) */
const checkSlotStyle: React.CSSProperties = { display: 'inline-flex', width: '16px', marginRight: '6px' };
import { useSheetStore } from './store';
import { NumFmt } from '../core/model';

interface ToolbarProps {
  onNew: () => void;
  onOpen: () => void;
  onOpenRecent: (path: string) => void;
  onSave: () => void;
  onSaveAs: () => void;
  onExportCsv: () => void;
  onImportCsv: () => void;
  onPrint: () => void;
  onExportPdf: () => void;
  onOpenHelp: () => void;
}

// theme-exempt-start: จานสีให้ผู้ใช้เลือกสีเซล (ค่าสีที่เขียนลงไฟล์ ไม่ใช่สี UI)
const PRESET_COLORS = [
  '#000000', '#434343', '#666666', '#999999', '#cccccc', '#ffffff',
  '#ef4444', '#f97316', '#f59e0b', '#84cc16', '#10b981', '#06b6d4',
  '#3b82f6', '#6366f1', '#8b5cf6', '#d946ef', '#fecaca', '#fed7aa',
  '#bbf7d0', '#bfdbfe',
];
// theme-exempt-end

export const Toolbar: React.FC<ToolbarProps> = ({
  onNew,
  onOpen,
  onOpenRecent,
  onSave,
  onSaveAs,
  onOpenHelp,
  onExportCsv,
  onImportCsv,
  onPrint,
  onExportPdf,
}) => {
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
    recentFiles,
    requestSort,
    openAdvancedSort,
    toggleFilter,
    workbook,
    activeSheetIndex,
    yearMode,
    setYearMode,
    themePref,
    setThemePref,
    gridWhite,
    setGridWhite,
  } = useSheetStore();

  const [recentMenuOpen, setRecentMenuOpen] = useState(false);
  const [sigmaMenuOpen, setSigmaMenuOpen] = useState(false);
  const [rowColMenuOpen, setRowColMenuOpen] = useState(false);
  const [textColorOpen, setTextColorOpen] = useState(false);
  const [fillColorOpen, setFillColorOpen] = useState(false);
  const [borderMenuOpen, setBorderMenuOpen] = useState(false);
  const [sortMenuOpen, setSortMenuOpen] = useState(false);
  const [fileMenuOpen, setFileMenuOpen] = useState(false);
  const [settingsMenuOpen, setSettingsMenuOpen] = useState(false);

  const dropdownStyle: React.CSSProperties = {
    position: 'absolute',
    top: '100%',
    left: 0,
    zIndex: 100,
    backgroundColor: 'var(--menu-bg)',
    border: '1px solid var(--border-color)',
    borderRadius: '6px',
    boxShadow: 'var(--shadow-menu)',
    display: 'flex',
    flexDirection: 'column',
    minWidth: '220px',
    marginTop: '4px',
    padding: '4px 0',
  };
  const dropdownItemStyle: React.CSSProperties = {
    border: 'none',
    background: 'transparent',
    justifyContent: 'flex-start',
    padding: '6px 12px',
  };

  const activeStyle = getActiveCellStyle();
  const hasFilter = !!workbook.sheets[activeSheetIndex]?.filter;

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
    setRecentMenuOpen(false);
    setSigmaMenuOpen(false);
    setRowColMenuOpen(false);
    setTextColorOpen(false);
    setFillColorOpen(false);
    setBorderMenuOpen(false);
    setSortMenuOpen(false);
    setFileMenuOpen(false);
    setSettingsMenuOpen(false);
  };

  const activeBtnStyle = (isActive?: boolean): React.CSSProperties => ({
    backgroundColor: isActive ? 'var(--color-primary)' : 'transparent',
    color: isActive ? 'var(--fg-on-primary)' : 'var(--fg-main)',
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
        {th.toolbar.newFile}
      </button>
      <button onClick={onOpen} title={th.toolbar.openFileTooltip}>
        {th.toolbar.openFile}
      </button>

      {/* Recent Files Dropdown */}
      <div style={{ position: 'relative', display: 'inline-block' }}>
        <button
          onClick={() => {
            const next = !recentMenuOpen;
            closeAllMenus();
            setRecentMenuOpen(next);
          }}
          title={th.toolbar.recentFiles}
        >
          {th.toolbar.recentFiles} <IconChevronDown />
        </button>
        {recentMenuOpen && (
          <div
            style={{
              position: 'absolute',
              top: '100%',
              left: 0,
              zIndex: 100,
              backgroundColor: 'var(--bg-surface)',
              border: '1px solid var(--border-color)',
              borderRadius: '6px',
              boxShadow: 'var(--shadow-menu)',
              minWidth: '240px',
              maxWidth: '380px',
              marginTop: '4px',
              padding: '6px 0',
            }}
          >
            {recentFiles.length === 0 ? (
              <div style={{ padding: '8px 14px', fontSize: '13px', color: 'var(--fg-muted)' }}>
                {th.toolbar.noRecentFiles}
              </div>
            ) : (
              recentFiles.map((p) => {
                const baseName = p.split(/[\\/]/).pop() || p;
                return (
                  <div
                    key={p}
                    onClick={() => {
                      setRecentMenuOpen(false);
                      onOpenRecent(p);
                    }}
                    title={p}
                    style={{
                      padding: '8px 14px',
                      cursor: 'pointer',
                      borderBottom: '1px solid var(--border-color)',
                    }}
                    onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = 'var(--bg-hover)')}
                    onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'transparent')}
                  >
                    <div style={{ fontSize: '13px', fontWeight: 500, color: 'var(--fg-main)' }}>
                      {baseName}
                    </div>
                    <div
                      style={{
                        fontSize: '11px',
                        color: 'var(--fg-muted)',
                        overflow: 'hidden',
                        textOverflow: 'ellipsis',
                        whiteSpace: 'nowrap',
                      }}
                    >
                      {p}
                    </div>
                  </div>
                );
              })
            )}
          </div>
        )}
      </div>

      <button className="primary" onClick={onSave} title={th.toolbar.saveFileTooltip}>
        {th.toolbar.saveFile}
      </button>
      <button onClick={onSaveAs} title={th.toolbar.saveAsFileTooltip}>
        {th.toolbar.saveAsFile}
      </button>

      {/* ไฟล์ > ส่งออก / นำเข้า CSV (4D) */}
      <div style={{ position: 'relative', display: 'inline-block' }}>
        <button
          onClick={() => {
            const next = !fileMenuOpen;
            closeAllMenus();
            setFileMenuOpen(next);
          }}
          title={th.csv.menuTooltip}
        >
          {th.csv.menu} <IconChevronDown />
        </button>
        {fileMenuOpen && (
          <div style={dropdownStyle}>
            <button style={dropdownItemStyle} onClick={() => { setFileMenuOpen(false); onPrint(); }}>
              {th.print.menuPrint}
            </button>
            <button style={dropdownItemStyle} onClick={() => { setFileMenuOpen(false); onExportPdf(); }}>
              {th.print.menuPdf}
            </button>
            <span style={{ height: '1px', background: 'var(--border-color)', margin: '4px 0' }} />
            <button style={dropdownItemStyle} onClick={() => { setFileMenuOpen(false); onExportCsv(); }}>
              {th.csv.exportCsv}
            </button>
            <button style={dropdownItemStyle} onClick={() => { setFileMenuOpen(false); onImportCsv(); }}>
              {th.csv.importCsv}
            </button>
          </div>
        )}
      </div>

      <span style={{ width: '1px', height: '20px', background: 'var(--border-color)', margin: '0 4px' }} />

      {/* History group */}
      <button onClick={undo} disabled={!history.canUndo()} title={th.toolbar.undoTooltip}>
        {th.toolbar.undo}
      </button>
      <button onClick={redo} disabled={!history.canRedo()} title={th.toolbar.redoTooltip}>
        {th.toolbar.redo}
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
              backgroundColor: 'var(--menu-bg)',
              border: '1px solid var(--border-color)',
              borderRadius: '6px',
              boxShadow: 'var(--shadow-menu)',
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
                value={activeStyle.color || '#000000'} // theme-exempt: ค่าเริ่มต้นของตัวเลือกสีเซล
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
          <IconFillColor size={18} />
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
              backgroundColor: 'var(--menu-bg)',
              border: '1px solid var(--border-color)',
              borderRadius: '6px',
              boxShadow: 'var(--shadow-menu)',
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
                value={activeStyle.fill || '#ffffff'} // theme-exempt: ค่าเริ่มต้นของตัวเลือกสีเซล
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
        <IconAlignLeft />
      </button>
      <button
        style={{ ...activeBtnStyle(activeStyle.align === 'center'), width: '32px' }}
        onClick={() => setAlignment('center')}
        title={th.toolbar.alignCenter}
      >
        <IconAlignCenter />
      </button>
      <button
        style={{ ...activeBtnStyle(activeStyle.align === 'right'), width: '32px' }}
        onClick={() => setAlignment('right')}
        title={th.toolbar.alignRight}
      >
        <IconAlignRight />
      </button>

      {/* Text wrap toggle */}
      <button
        style={activeBtnStyle(activeStyle.wrap)}
        onClick={toggleWrap}
        title={th.toolbar.wrapText}
      >
        {th.toolbar.wrapText}
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
          {th.toolbar.borders} <IconChevronDown />
        </button>

        {borderMenuOpen && (
          <div
            style={{
              position: 'absolute',
              top: '100%',
              left: 0,
              zIndex: 100,
              backgroundColor: 'var(--menu-bg)',
              border: '1px solid var(--border-color)',
              borderRadius: '6px',
              boxShadow: 'var(--shadow-menu)',
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
              {th.toolbar.borderNone}
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
              {th.toolbar.borderAll}
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
              {th.toolbar.borderOuter}
            </button>
          </div>
        )}
      </div>

      <span style={{ width: '1px', height: '20px', background: 'var(--border-color)', margin: '0 4px' }} />

      {/* Merge group */}
      <button onClick={() => mergeCells()} title={th.toolbar.mergeCellsTooltip}>
        {th.toolbar.mergeCells}
      </button>
      <button onClick={() => unmergeCells()} title={th.toolbar.unmergeCellsTooltip}>
        {th.toolbar.unmergeCells}
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

      <button onClick={() => setNumFmt('baht')} title={th.toolbar.bahtTooltip}>
        ฿
      </button>
      <button onClick={() => setNumFmt('percent')} title={th.toolbar.percentTooltip}>
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
          title={th.toolbar.rowColTooltip}
        >
          {th.toolbar.rowCol} <IconChevronDown />
        </button>
        {rowColMenuOpen && (
          <div
            style={{
              position: 'absolute',
              top: '100%',
              left: 0,
              zIndex: 100,
              backgroundColor: 'var(--menu-bg)',
              border: '1px solid var(--border-color)',
              borderRadius: '6px',
              boxShadow: 'var(--shadow-menu)',
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
              {th.toolbar.insertRowAbove}
            </button>
            <button
              style={{ border: 'none', background: 'transparent', justifyContent: 'flex-start', padding: '6px 12px' }}
              onClick={() => {
                insertRowAt(activeCell.r + 1);
                setRowColMenuOpen(false);
              }}
            >
              {th.toolbar.insertRowBelow}
            </button>
            <button
              style={{ border: 'none', background: 'transparent', justifyContent: 'flex-start', padding: '6px 12px', color: 'var(--color-danger)' }}
              onClick={() => {
                deleteRowAt(activeCell.r);
                setRowColMenuOpen(false);
              }}
            >
              {th.toolbar.deleteRow}
            </button>
            <span style={{ height: '1px', background: 'var(--border-color)', margin: '4px 0' }} />
            <button
              style={{ border: 'none', background: 'transparent', justifyContent: 'flex-start', padding: '6px 12px' }}
              onClick={() => {
                insertColAt(activeCell.c);
                setRowColMenuOpen(false);
              }}
            >
              {th.toolbar.insertColLeft}
            </button>
            <button
              style={{ border: 'none', background: 'transparent', justifyContent: 'flex-start', padding: '6px 12px' }}
              onClick={() => {
                insertColAt(activeCell.c + 1);
                setRowColMenuOpen(false);
              }}
            >
              {th.toolbar.insertColRight}
            </button>
            <button
              style={{ border: 'none', background: 'transparent', justifyContent: 'flex-start', padding: '6px 12px', color: 'var(--color-danger)' }}
              onClick={() => {
                deleteColAt(activeCell.c);
                setRowColMenuOpen(false);
              }}
            >
              {th.toolbar.deleteCol}
            </button>
            <span style={{ height: '1px', background: 'var(--border-color)', margin: '4px 0' }} />
            <button
              style={{ border: 'none', background: 'transparent', justifyContent: 'flex-start', padding: '6px 12px', fontWeight: 600 }}
              onClick={() => {
                addRowsAtBottom(50);
                setRowColMenuOpen(false);
              }}
            >
              {th.toolbar.addRowsBottom}
            </button>
          </div>
        )}
      </div>

      <span style={{ width: '1px', height: '20px', background: 'var(--border-color)', margin: '0 4px' }} />

      {/* ตัวกรอง (4C): กดซ้ำเพื่อปิด */}
      <button
        onClick={() => {
          closeAllMenus();
          toggleFilter();
        }}
        title={th.filter.buttonTooltip}
        aria-pressed={hasFilter}
        style={activeBtnStyle(hasFilter)}
      >
        {th.filter.button}
      </button>

      {/* Sort menu (4C) */}
      <div style={{ position: 'relative', display: 'inline-block' }}>
        <button
          onClick={() => {
            const next = !sortMenuOpen;
            closeAllMenus();
            setSortMenuOpen(next);
          }}
          title={th.sort.menuTooltip}
        >
          {th.sort.menu} <IconChevronDown />
        </button>
        {sortMenuOpen && (
          <div style={dropdownStyle}>
            <button style={dropdownItemStyle} onClick={() => { setSortMenuOpen(false); requestSort('asc'); }}>
              {th.sort.asc}
            </button>
            <button style={dropdownItemStyle} onClick={() => { setSortMenuOpen(false); requestSort('desc'); }}>
              {th.sort.desc}
            </button>
            <span style={{ height: '1px', background: 'var(--border-color)', margin: '4px 0' }} />
            <button style={dropdownItemStyle} onClick={() => { setSortMenuOpen(false); openAdvancedSort(); }}>
              {th.sort.advanced}
            </button>
          </div>
        )}
      </div>

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
          {th.toolbar.sigma} <IconChevronDown />
        </button>
        {sigmaMenuOpen && (
          <div
            style={{
              position: 'absolute',
              top: '100%',
              left: 0,
              zIndex: 100,
              backgroundColor: 'var(--menu-bg)',
              border: '1px solid var(--border-color)',
              borderRadius: '6px',
              boxShadow: 'var(--shadow-menu)',
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
              {th.toolbar.otherFunctions}
            </button>
          </div>
        )}
      </div>

      {/* Help Button + ตั้งค่า (5A) */}
      <div style={{ marginLeft: 'auto', display: 'flex', gap: '6px', alignItems: 'center' }}>
        <div style={{ position: 'relative', display: 'inline-block' }}>
          <button
            onClick={() => {
              const next = !settingsMenuOpen;
              closeAllMenus();
              setSettingsMenuOpen(next);
            }}
            title={th.settings.menuTooltip}
          >
            {th.settings.menu} <IconChevronDown />
          </button>
          {settingsMenuOpen && (
            <div style={{ ...dropdownStyle, left: 'auto', right: 0 }} role="menu">
              <div style={{ padding: '4px 12px', fontSize: '12px', color: 'var(--fg-muted)' }}>
                {th.settings.yearFormat}
              </div>
              {(['ce', 'be'] as const).map((mode) => (
                <button
                  key={mode}
                  role="menuitemradio"
                  aria-checked={yearMode === mode}
                  style={{ ...dropdownItemStyle, fontWeight: yearMode === mode ? 600 : 400 }}
                  onClick={() => {
                    setSettingsMenuOpen(false);
                    setYearMode(mode);
                  }}
                >
                  <span style={checkSlotStyle}>{yearMode === mode && <IconCheck />}</span>
                  {mode === 'ce' ? th.settings.yearCE : th.settings.yearBE}
                </button>
              ))}
              <div style={{ height: '1px', backgroundColor: 'var(--border-color)', margin: '4px 0' }} />
              <div style={{ padding: '4px 12px', fontSize: '12px', color: 'var(--fg-muted)' }}>
                {th.settings.theme}
              </div>
              {(['system', 'light', 'dark'] as const).map((pref) => (
                <button
                  key={pref}
                  role="menuitemradio"
                  aria-checked={themePref === pref}
                  style={{ ...dropdownItemStyle, fontWeight: themePref === pref ? 600 : 400 }}
                  onClick={() => {
                    setSettingsMenuOpen(false);
                    setThemePref(pref);
                  }}
                >
                  <span style={checkSlotStyle}>{themePref === pref && <IconCheck />}</span>
                  {pref === 'system' ? th.settings.themeSystem : pref === 'light' ? th.settings.themeLight : th.settings.themeDark}
                </button>
              ))}
              <div style={{ height: '1px', backgroundColor: 'var(--border-color)', margin: '4px 0' }} />
              <div style={{ padding: '4px 12px', fontSize: '12px', color: 'var(--fg-muted)' }}>
                {th.settings.gridBackground}
              </div>
              {([false, true] as const).map((white) => (
                <button
                  key={String(white)}
                  role="menuitemradio"
                  aria-checked={gridWhite === white}
                  style={{ ...dropdownItemStyle, fontWeight: gridWhite === white ? 600 : 400 }}
                  onClick={() => {
                    setSettingsMenuOpen(false);
                    setGridWhite(white);
                  }}
                >
                  <span style={checkSlotStyle}>{gridWhite === white && <IconCheck />}</span>
                  {white ? th.settings.gridAlwaysWhite : th.settings.gridFollowTheme}
                </button>
              ))}
            </div>
          )}
        </div>
        <button onClick={onOpenHelp} title={th.toolbar.helpTooltip}>
          {th.toolbar.help}
        </button>
      </div>
    </div>
  );
};
