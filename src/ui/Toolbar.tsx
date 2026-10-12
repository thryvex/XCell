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

export const Toolbar: React.FC<ToolbarProps> = ({ onNew, onOpen, onSave, onSaveAs }) => {
  const {
    undo,
    redo,
    history,
    mergeCells,
    unmergeCells,
    setNumFmt,
    applySigma,
  } = useSheetStore();

  const [sigmaMenuOpen, setSigmaMenuOpen] = useState(false);

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

      {/* Sigma menu */}
      <div style={{ position: 'relative', display: 'inline-block' }}>
        <button
          onClick={() => setSigmaMenuOpen(!sigmaMenuOpen)}
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
          </div>
        )}
      </div>
    </div>
  );
};
