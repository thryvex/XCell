import React, { useEffect, useMemo, useState } from 'react';
import { useSheetStore } from './store';
import { th } from '../i18n/th';
import {
  DEFAULT_PRINT_SETTINGS,
  PrintSettings,
  PaperSize,
  PrintMargins,
} from '../core/model';
import { buildPrintJob, pageGeometry, MANY_PAGES_WARNING, MIN_SCALE, MAX_SCALE } from '../core/paginate';
import { renderPrintPage, renderPrintDocument, PrintMeta } from '../core/printRender';
import { formatDate, todaySerial } from '../core/numfmt';
import { getPrintFontCss } from './printFonts';

const MM_TO_PX = 96 / 25.4;
const PREVIEW_MAX_W = 520;
const PREVIEW_MAX_H = 640;

interface PrintDialogProps {
  isOpen: boolean;
  focus: 'print' | 'pdf';
  onClose: () => void;
  onToast: (msg: string) => void;
}

/**
 * หน้าต่าง "พิมพ์": ตัวเลือกด้านซ้าย, ตัวอย่างก่อนพิมพ์ด้านขวา (วาดเฉพาะหน้าที่แสดง)
 * ตัวอย่างและเอกสารที่ส่งพิมพ์สร้างจากฟังก์ชันชุดเดียวกัน (renderPrintPage / renderPrintDocument)
 */
export const PrintDialog: React.FC<PrintDialogProps> = ({ isOpen, focus, onClose, onToast }) => {
  const { workbook, activeSheetIndex, selection, evaluator, fileName, yearMode, setPrintSettings } =
    useSheetStore();
  const settings: PrintSettings = { ...DEFAULT_PRINT_SETTINGS, ...workbook.printSettings };
  const [pageIndex, setPageIndex] = useState(0);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);
  // ฟอนต์ Sarabun ฝังในเอกสาร (โหลดครั้งเดียว)
  const [fontCss, setFontCss] = useState('');
  useEffect(() => {
    if (isOpen) getPrintFontCss().then(setFontCss);
  }, [isOpen]);

  const hasSelection = selection.r1 !== selection.r2 || selection.c1 !== selection.c2;
  const effective: PrintSettings =
    settings.what === 'selection' && !hasSelection ? { ...settings, what: 'sheet' } : settings;

  const job = useMemo(
    () => (isOpen ? buildPrintJob(workbook, activeSheetIndex, selection, effective) : { pages: [], mergeSplit: false }),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [isOpen, workbook, activeSheetIndex, selection, JSON.stringify(effective), yearMode]
  );
  const total = job.pages.length;

  useEffect(() => {
    if (pageIndex >= total) setPageIndex(Math.max(0, total - 1));
  }, [total, pageIndex]);
  useEffect(() => {
    if (isOpen) {
      setPageIndex(0);
      setError(null);
    }
  }, [isOpen]);

  const meta: PrintMeta = { fileName, printedAt: formatDate(todaySerial()) };
  const labels = { pageNumber: th.print.pageNumberText };

  // ตัวอย่าง: วาดเฉพาะหน้าที่แสดงอยู่ (lazy)
  const previewDoc = useMemo(() => {
    if (!isOpen || total === 0) return '';
    const page = job.pages[Math.min(pageIndex, total - 1)];
    const html = renderPrintPage(workbook, evaluator, page, pageIndex + 1, total, effective, meta, labels);
    return renderPrintDocument([html], effective, fontCss);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isOpen, job, pageIndex, total, fontCss]);

  if (!isOpen) return null;

  const geo = pageGeometry(effective);
  const pageWpx = geo.pageW * MM_TO_PX;
  const pageHpx = geo.pageH * MM_TO_PX;
  const previewScale = Math.min(PREVIEW_MAX_W / pageWpx, PREVIEW_MAX_H / pageHpx);

  const buildFullDocument = () => {
    const pagesHtml = job.pages.map((page, i) =>
      renderPrintPage(workbook, evaluator, page, i + 1, total, effective, meta, labels)
    );
    return renderPrintDocument(pagesHtml, effective, fontCss);
  };

  const run = async (action: 'print' | 'pdf') => {
    const electronAPI = (window as any).electronAPI;
    if (!electronAPI?.printDocument) {
      setError(th.print.desktopOnly);
      return;
    }
    setBusy(true);
    setError(null);
    try {
      const req = { html: buildFullDocument(), pageWidthMm: geo.pageW, pageHeightMm: geo.pageH };
      if (action === 'print') {
        const res = await electronAPI.printDocument(req);
        if (res?.success) onClose();
        else if (!res?.canceled) setError(th.print.printFailed + (res?.error ?? ''));
      } else {
        const res = await electronAPI.printToPdf({ ...req, defaultPath: `${fileName}.pdf` });
        if (res?.success) {
          onToast(th.print.pdfSaved);
          onClose();
        } else if (!res?.canceled) {
          setError(th.print.pdfFailed + (res?.error ?? ''));
        }
      }
    } catch (err: any) {
      setError((action === 'print' ? th.print.printFailed : th.print.pdfFailed) + (err?.message || String(err)));
    } finally {
      setBusy(false);
    }
  };

  const row: React.CSSProperties = { display: 'flex', flexDirection: 'column', gap: '4px' };
  const label: React.CSSProperties = { fontWeight: 600, fontSize: '13px' };
  const check = (key: keyof PrintSettings, text: string) => (
    <label style={{ display: 'flex', alignItems: 'center', gap: '6px', cursor: 'pointer' }}>
      <input
        type="checkbox"
        checked={Boolean(settings[key])}
        onChange={(e) => setPrintSettings({ [key]: e.target.checked } as Partial<PrintSettings>)}
      />
      {text}
    </label>
  );
  const disabled = busy || total === 0;

  return (
    <div
      style={{
        position: 'fixed',
        inset: 0,
        backgroundColor: 'rgba(0,0,0,0.4)',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        zIndex: 3000,
      }}
    >
      <div
        role="dialog"
        aria-label={th.print.title}
        onKeyDown={(e) => {
          if (e.key === 'Escape') onClose();
        }}
        style={{
          backgroundColor: 'var(--bg-app)',
          borderRadius: '8px',
          boxShadow: '0 8px 28px rgba(0,0,0,0.25)',
          display: 'flex',
          gap: '20px',
          padding: '20px',
          maxWidth: '96vw',
          maxHeight: '94vh',
          fontSize: '13px',
        }}
      >
        {/* ตัวเลือก */}
        <div style={{ width: '250px', display: 'flex', flexDirection: 'column', gap: '12px', overflowY: 'auto' }}>
          <h3 style={{ margin: 0, fontSize: '16px' }}>🖨 {th.print.title}</h3>

          <div style={row}>
            <span style={label}>{th.print.what}</span>
            <select
              value={effective.what}
              onChange={(e) => setPrintSettings({ what: e.target.value as PrintSettings['what'] })}
            >
              <option value="sheet">{th.print.whatSheet}</option>
              <option value="all">{th.print.whatAll}</option>
              <option value="selection" disabled={!hasSelection}>
                {th.print.whatSelection}
              </option>
            </select>
          </div>

          <div style={{ display: 'flex', gap: '8px' }}>
            <div style={{ ...row, flex: 1 }}>
              <span style={label}>{th.print.paper}</span>
              <select value={settings.paper} onChange={(e) => setPrintSettings({ paper: e.target.value as PaperSize })}>
                <option value="A4">A4</option>
                <option value="Letter">Letter</option>
                <option value="A5">A5</option>
              </select>
            </div>
            <div style={{ ...row, flex: 1 }}>
              <span style={label}>{th.print.orientation}</span>
              <select
                value={settings.orientation}
                onChange={(e) => setPrintSettings({ orientation: e.target.value as PrintSettings['orientation'] })}
              >
                <option value="portrait">{th.print.portrait}</option>
                <option value="landscape">{th.print.landscape}</option>
              </select>
            </div>
          </div>

          <div style={row}>
            <span style={label}>{th.print.scale}</span>
            <select
              value={settings.scaleMode}
              onChange={(e) => setPrintSettings({ scaleMode: e.target.value as PrintSettings['scaleMode'] })}
            >
              <option value="100">{th.print.scale100}</option>
              <option value="fitWidth">{th.print.scaleFit}</option>
              <option value="custom">{th.print.scaleCustom}</option>
            </select>
            {settings.scaleMode === 'custom' && (
              <label style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                <input
                  type="number"
                  min={MIN_SCALE * 100}
                  max={MAX_SCALE * 100}
                  value={settings.scalePercent}
                  onChange={(e) => {
                    const n = Number(e.target.value);
                    if (!isNaN(n)) setPrintSettings({ scalePercent: Math.max(30, Math.min(200, Math.round(n))) });
                  }}
                  style={{ width: '70px' }}
                />
                % <span style={{ color: 'var(--fg-muted)' }}>({th.print.scalePercentHint})</span>
              </label>
            )}
          </div>

          <div style={row}>
            <span style={label}>{th.print.margins}</span>
            <select value={settings.margins} onChange={(e) => setPrintSettings({ margins: e.target.value as PrintMargins })}>
              <option value="normal">{th.print.marginsNormal}</option>
              <option value="narrow">{th.print.marginsNarrow}</option>
              <option value="wide">{th.print.marginsWide}</option>
            </select>
          </div>

          <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
            {check('gridlines', th.print.gridlines)}
            {check('fills', th.print.fills)}
            {check('repeatFrozen', th.print.repeatFrozen)}
            {check('pageNumbers', th.print.pageNumbers)}
            {check('header', th.print.header)}
          </div>

          {total > MANY_PAGES_WARNING && (
            <div role="alert" style={{ color: '#b45309' }}>
              ⚠️ {th.print.manyPages.replace('{count}', total.toLocaleString('th-TH'))}
            </div>
          )}
          {job.mergeSplit && <div style={{ color: '#b45309' }}>⚠️ {th.print.mergeSplit}</div>}
          {error && (
            <div role="alert" style={{ color: 'var(--color-danger)' }}>
              {error}
            </div>
          )}

          <div style={{ marginTop: 'auto', display: 'flex', flexDirection: 'column', gap: '6px' }}>
            {total === 0 && <div style={{ color: 'var(--color-danger)', fontWeight: 600 }}>{th.print.noData}</div>}
            {busy && <div style={{ color: 'var(--fg-muted)' }}>{th.print.working}</div>}
            <button
              autoFocus={focus === 'print'}
              disabled={disabled}
              onClick={() => run('print')}
              style={{ fontWeight: 600 }}
              title={total === 0 ? th.print.noData : undefined}
            >
              🖨 {th.print.printButton}
            </button>
            <button autoFocus={focus === 'pdf'} disabled={disabled} onClick={() => run('pdf')}>
              📄 {th.print.pdfButton}
            </button>
            <button onClick={onClose} disabled={busy}>
              {th.dialog.cancel}
            </button>
          </div>
        </div>

        {/* ตัวอย่างก่อนพิมพ์ */}
        <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '8px' }}>
          <div
            style={{
              width: pageWpx * previewScale,
              height: pageHpx * previewScale,
              boxShadow: '0 2px 10px rgba(0,0,0,0.25)',
              backgroundColor: '#ffffff',
              overflow: 'hidden',
            }}
          >
            {total > 0 && (
              // sandbox ว่าง = ไม่รันสคริปต์ ไม่มีสิทธิ์ใดๆ (เอกสารเองก็มี CSP)
              <iframe
                title={th.print.title}
                sandbox=""
                srcDoc={previewDoc}
                style={{
                  width: pageWpx,
                  height: pageHpx,
                  border: 'none',
                  transform: `scale(${previewScale})`,
                  transformOrigin: 'top left',
                  backgroundColor: '#ffffff',
                }}
              />
            )}
          </div>
          <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
            <button disabled={pageIndex <= 0} onClick={() => setPageIndex((i) => Math.max(0, i - 1))}>
              {th.print.prevPage}
            </button>
            <span aria-live="polite">
              {th.print.previewPage
                .replace('{page}', String(total === 0 ? 0 : pageIndex + 1))
                .replace('{total}', String(total))}
            </span>
            <button disabled={pageIndex >= total - 1} onClick={() => setPageIndex((i) => Math.min(total - 1, i + 1))}>
              {th.print.nextPage}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};
