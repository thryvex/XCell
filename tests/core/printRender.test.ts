import { describe, it, expect, afterEach } from 'vitest';
import { escapeHtml, renderPrintPage, renderPrintDocument } from '../../src/core/printRender';
import { paginateSheet, printExtent } from '../../src/core/paginate';
import { Evaluator } from '../../src/core/evaluator';
import { createEmptySheet, Workbook, DEFAULT_PRINT_SETTINGS, PrintSettings } from '../../src/core/model';
import { dateToSerial, setDisplayYearMode } from '../../src/core/numfmt';

const labels = { pageNumber: 'หน้า {page} / {total}' };
const meta = { fileName: 'บัญชี <ทดสอบ>', printedAt: '16/10/2569' };

function renderFirstPage(wb: Workbook, settings: PrintSettings = DEFAULT_PRINT_SETTINGS): string {
  const sheet = wb.sheets[0];
  const { pages } = paginateSheet(sheet, printExtent(sheet), settings, 0);
  return renderPrintPage(wb, new Evaluator(wb), pages[0], 1, pages.length, settings, meta, labels);
}

afterEach(() => setDisplayYearMode('ce'));

describe('printRender: ความปลอดภัย', () => {
  it('escapes all HTML special characters', () => {
    expect(escapeHtml(`<a href="x">'&'</a>`)).toBe('&lt;a href=&quot;x&quot;&gt;&#39;&amp;&#39;&lt;/a&gt;');
  });

  it('cell text with script/img injection comes out as plain text', () => {
    const sheet = createEmptySheet('ชีต1');
    sheet.cells['A1'] = { raw: '<script>alert(1)</script>' };
    sheet.cells['A2'] = { raw: '"><img src=x onerror=alert(1)>' };
    const html = renderFirstPage({ sheets: [sheet], activeSheet: 0 }, { ...DEFAULT_PRINT_SETTINGS, header: true });
    expect(html).not.toContain('<script');
    expect(html).not.toContain('<img');
    expect(html).toContain('&lt;script&gt;alert(1)&lt;/script&gt;');
    expect(html).toContain('&quot;&gt;&lt;img src=x onerror=alert(1)&gt;');
    expect(html).toContain('บัญชี &lt;ทดสอบ&gt;'); // ชื่อไฟล์ที่หัวกระดาษก็ถูก escape
  });

  it('colors and font names that could inject CSS are ignored', () => {
    const sheet = createEmptySheet('ชีต1');
    sheet.cells['A1'] = {
      raw: 'x',
      style: { color: 'red;background:url(http://evil)', fill: '#12345g', fontFamily: 'A"};x{' },
    };
    const html = renderFirstPage({ sheets: [sheet], activeSheet: 0 });
    expect(html).not.toContain('evil');
    expect(html).not.toContain('#12345g');
    expect(html).not.toContain('x{');
  });

  it('document has a CSP that blocks scripts and remote resources', () => {
    const doc = renderPrintDocument(['<div class="page"></div>'], DEFAULT_PRINT_SETTINGS, '');
    expect(doc).toContain(`default-src 'none'; style-src 'unsafe-inline'; font-src data:; img-src data:`);
    expect(doc).not.toMatch(/<script/i);
  });
});

describe('printRender: ค่าที่แสดงและสไตล์', () => {
  it('shows values exactly as in the grid (baht, %, BE dates, Thai errors)', () => {
    setDisplayYearMode('be');
    const sheet = createEmptySheet('ชีต1');
    sheet.cells['A1'] = { raw: '1234.5', fmt: 'baht' };
    sheet.cells['B1'] = { raw: '0.075', fmt: 'percent' };
    sheet.cells['C1'] = { raw: String(dateToSerial(new Date(2026, 9, 15))), fmt: 'date' };
    sheet.cells['D1'] = { raw: '=1/0' };
    sheet.cells['E1'] = { raw: '=A1*2', fmt: 'baht' };
    const html = renderFirstPage({ sheets: [sheet], activeSheet: 0 });
    expect(html).toContain('>฿1,234.50</td>');
    expect(html).toContain('>7.50%</td>');
    expect(html).toContain('>15/10/2569</td>');
    expect(html).toMatch(/color:#dc2626[^>]*>#หารศูนย์<\/td>/);
    expect(html).toContain('>฿2,469.00</td>');
  });

  it('applies bold/italic/underline/colors/align/wrap/border; numbers align right', () => {
    const sheet = createEmptySheet('ชีต1');
    sheet.cells['A1'] = {
      raw: 'หัว',
      style: { bold: true, italic: true, underline: true, color: '#ff0000', fill: '#00ff00', align: 'center', wrap: true, border: 'all' },
    };
    sheet.cells['B1'] = { raw: '5' };
    const html = renderFirstPage({ sheets: [sheet], activeSheet: 0 });
    const a1 = html.match(/<td style="([^"]*)">หัว<\/td>/)![1];
    for (const part of [
      'font-weight:bold',
      'font-style:italic',
      'text-decoration:underline',
      'color:#ff0000',
      'background-color:#00ff00',
      'text-align:center',
      'white-space:normal',
      'border:0.3mm solid #000000',
    ]) {
      expect(a1).toContain(part);
    }
    expect(html).toMatch(/text-align:right">5<\/td>/);
  });

  it('fill colours can be turned off; gridlines can be turned on', () => {
    const sheet = createEmptySheet('ชีต1');
    sheet.cells['A1'] = { raw: 'x', style: { fill: '#00ff00' } };
    const wb = { sheets: [sheet], activeSheet: 0 };
    expect(renderFirstPage(wb, { ...DEFAULT_PRINT_SETTINGS, fills: false })).not.toContain('#00ff00');
    expect(renderPrintDocument([], { ...DEFAULT_PRINT_SETTINGS, gridlines: true }, '')).toContain('border: 0.2mm solid #c0c0c0');
    expect(renderPrintDocument([], DEFAULT_PRINT_SETTINGS, '')).not.toContain('#c0c0c0');
  });

  it('merged cells get rowspan/colspan and covered cells are omitted', () => {
    const sheet = createEmptySheet('ชีต1');
    sheet.cells['A1'] = { raw: 'ผสาน' };
    sheet.cells['C3'] = { raw: 'ท้าย' };
    sheet.merges = [{ r1: 0, c1: 0, r2: 1, c2: 1 }];
    const html = renderFirstPage({ sheets: [sheet], activeSheet: 0 });
    expect(html).toMatch(/<td rowspan="2" colspan="2" style="[^"]*">ผสาน<\/td>/);
    const rows = html.match(/<tr[^>]*>.*?<\/tr>/g)!;
    expect(rows).toHaveLength(3);
    expect((rows[0].match(/<td/g) ?? []).length).toBe(2); // A1(ผสาน) + C1
    expect((rows[1].match(/<td/g) ?? []).length).toBe(1); // C2 เท่านั้น
    expect((rows[2].match(/<td/g) ?? []).length).toBe(3);
  });

  it('page number footer and header text', () => {
    const sheet = createEmptySheet('ชีต1');
    sheet.cells['A1'] = { raw: 'x' };
    const html = renderFirstPage({ sheets: [sheet], activeSheet: 0 }, { ...DEFAULT_PRINT_SETTINGS, header: true });
    expect(html).toContain('<div class="page-footer">หน้า 1 / 1</div>');
    expect(html).toContain('<span>16/10/2569</span>');
    const noFooter = renderFirstPage({ sheets: [sheet], activeSheet: 0 }, { ...DEFAULT_PRINT_SETTINGS, pageNumbers: false });
    expect(noFooter).not.toContain('page-footer');
  });

  it('@page size follows paper and orientation with zero margins', () => {
    expect(renderPrintDocument([], DEFAULT_PRINT_SETTINGS, '')).toContain('@page { size: 210mm 297mm; margin: 0; }');
    expect(renderPrintDocument([], { ...DEFAULT_PRINT_SETTINGS, orientation: 'landscape', paper: 'Letter' }, '')).toContain(
      '@page { size: 279.4mm 215.9mm; margin: 0; }'
    );
  });

  it('rows hidden by a filter are not rendered', () => {
    const sheet = createEmptySheet('ชีต1');
    sheet.cells['A1'] = { raw: 'แสดง' };
    sheet.cells['A2'] = { raw: 'ซ่อน' };
    sheet.cells['A3'] = { raw: 'แสดง2' };
    sheet.hiddenRows = [1];
    const html = renderFirstPage({ sheets: [sheet], activeSheet: 0 });
    expect(html).not.toContain('ซ่อน');
    expect(html).toContain('แสดง2');
  });
});
