import React, { useMemo, useState } from 'react';
import { useSheetStore } from './store';
import { th } from '../i18n/th';
import { CsvEncoding, decodeBytes, parseCsv, csvSize, convertCsvValue } from '../core/csv';
import { formatDisplayValue } from '../core/numfmt';

const PREVIEW_ROWS = 10;
const PRESET_DELIMITERS = [',', ';', '\t'];

const overlay: React.CSSProperties = {
  position: 'fixed',
  inset: 0,
  backgroundColor: 'var(--overlay)',
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  zIndex: 3000,
};

/** แสดงค่าตัวอย่างตามที่จะได้หลังนำเข้า (เช่น 1,234.50 / 5.00% / 01/10/2026) */
function previewText(text: string, autoConvert: boolean): string {
  const cell = convertCsvValue(text, autoConvert);
  if (!cell) return '';
  if (cell.raw.startsWith("'")) return cell.raw.slice(1);
  const n = Number(cell.raw);
  return isNaN(n) ? cell.raw : formatDisplayValue(n, cell.fmt);
}

/**
 * หน้าต่างตัวเลือกก่อนนำเข้า CSV: encoding, ตัวคั่น, แปลงอัตโนมัติ, ตัวอย่าง 10 แถว (อัปเดตทันที)
 */
export const CsvImportDialog: React.FC = () => {
  const { pendingCsvImport: p, updateCsvImport, confirmCsvImport, cancelCsvImport } = useSheetStore();
  const [otherDelimiter, setOtherDelimiter] = useState('|');

  const parsed = useMemo(() => {
    if (!p || !p.delimiter) return { rows: [] as string[][], size: { rows: 0, cols: 0 } };
    const rows = parseCsv(decodeBytes(p.bytes, p.encoding), p.delimiter);
    return { rows, size: csvSize(rows) };
  }, [p]);

  if (!p) return null;

  if (p.step === 'limitWarning') {
    return (
      <div style={overlay}>
        <div role="alertdialog" style={panelStyle}>
          <h3 style={{ margin: 0, fontSize: '16px' }}>⚠️ {th.csv.importTitle}</h3>
          <p style={{ margin: 0 }}>
            {th.csv.limitWarning
              .replace('{rows}', parsed.size.rows.toLocaleString('th-TH'))
              .replace('{cols}', parsed.size.cols.toLocaleString('th-TH'))}
          </p>
          <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '8px' }}>
            <button onClick={cancelCsvImport}>{th.dialog.cancel}</button>
            <button autoFocus onClick={confirmCsvImport} style={{ fontWeight: 600 }}>
              {th.csv.importPartial}
            </button>
          </div>
        </div>
      </div>
    );
  }

  const isPreset = PRESET_DELIMITERS.includes(p.delimiter);
  const preview = parsed.rows.slice(0, PREVIEW_ROWS);
  const previewCols = Math.min(parsed.size.cols, 12);

  return (
    <div style={overlay}>
      <div
        role="dialog"
        style={{ ...panelStyle, width: '720px' }}
        onKeyDown={(e) => {
          if (e.key === 'Escape') cancelCsvImport();
        }}
      >
        <h3 style={{ margin: 0, fontSize: '16px' }}>
          {p.mode === 'newSheet' ? th.csv.importTitleNewSheet : th.csv.importTitle}
        </h3>
        <div style={{ color: 'var(--fg-muted)' }}>{th.csv.fileLabel.replace('{name}', p.fileName)}</div>

        <div style={{ display: 'flex', flexWrap: 'wrap', gap: '16px', alignItems: 'center' }}>
          <label style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
            {th.csv.encoding}
            <select
              value={p.encoding}
              onChange={(e) => updateCsvImport({ encoding: e.target.value as CsvEncoding })}
            >
              <option value="utf-8">{th.csv.encUtf8}</option>
              <option value="windows-874">{th.csv.encTis}</option>
              <option value="utf-16le">{th.csv.encUtf16le}</option>
              <option value="utf-16be">{th.csv.encUtf16be}</option>
            </select>
          </label>

          <label style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
            {th.csv.delimiter}
            <select
              value={isPreset ? p.delimiter : 'other'}
              onChange={(e) =>
                updateCsvImport({ delimiter: e.target.value === 'other' ? otherDelimiter : e.target.value })
              }
            >
              <option value=",">{th.csv.delimComma}</option>
              <option value=";">{th.csv.delimSemicolon}</option>
              <option value={'\t'}>{th.csv.delimTab}</option>
              <option value="other">{th.csv.delimOther}</option>
            </select>
            {!isPreset && (
              <input
                value={otherDelimiter}
                placeholder={th.csv.otherPlaceholder}
                maxLength={3}
                onChange={(e) => {
                  setOtherDelimiter(e.target.value);
                  if (e.target.value) updateCsvImport({ delimiter: e.target.value });
                }}
                style={{ width: '48px' }}
              />
            )}
          </label>

          <label style={{ display: 'flex', alignItems: 'center', gap: '6px', cursor: 'pointer' }}>
            <input
              type="checkbox"
              checked={p.autoConvert}
              onChange={(e) => updateCsvImport({ autoConvert: e.target.checked })}
            />
            {th.csv.autoConvert}
          </label>
        </div>

        <div>
          <div style={{ fontWeight: 600, marginBottom: '4px' }}>
            {th.csv.preview}{' '}
            <span style={{ fontWeight: 400, color: 'var(--fg-muted)' }}>
              ({th.csv.size
                .replace('{rows}', parsed.size.rows.toLocaleString('th-TH'))
                .replace('{cols}', parsed.size.cols.toLocaleString('th-TH'))})
            </span>
          </div>
          <div style={{ overflow: 'auto', maxHeight: '260px', border: '1px solid var(--border-color)' }}>
            <table style={{ borderCollapse: 'collapse', fontSize: '12px' }}>
              <tbody>
                {preview.map((row, r) => (
                  <tr key={r}>
                    {Array.from({ length: previewCols }, (_, c) => (
                      <td
                        key={c}
                        style={{
                          border: '1px solid var(--border-color)',
                          padding: '2px 6px',
                          maxWidth: '160px',
                          overflow: 'hidden',
                          textOverflow: 'ellipsis',
                          whiteSpace: 'nowrap',
                        }}
                      >
                        {previewText(row[c] ?? '', p.autoConvert)}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '8px' }}>
          <button onClick={cancelCsvImport}>{th.dialog.cancel}</button>
          <button autoFocus onClick={confirmCsvImport} style={{ fontWeight: 600 }} disabled={!p.delimiter}>
            {th.csv.import}
          </button>
        </div>
      </div>
    </div>
  );
};

const panelStyle: React.CSSProperties = {
  backgroundColor: 'var(--bg-app)',
  borderRadius: '8px',
  padding: '20px 24px',
  width: '460px',
  maxWidth: '94%',
  boxShadow: 'var(--shadow-dialog)',
  display: 'flex',
  flexDirection: 'column',
  gap: '12px',
  fontSize: '14px',
};
