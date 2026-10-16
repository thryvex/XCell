import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import ExcelJS from 'exceljs';
import { exportToXlsxBuffer, importFromXlsxBuffer } from '../../src/core/xlsx';
import { buildChartXml, chartRef, parseChartRef, parseChartXml, readChartsFromXlsx } from '../../src/core/chartXlsx';
import { getChartData } from '../../src/core/chart';
import { parseXml } from '../../src/core/xmlTree';
import { Workbook, Sheet, SheetChart, createEmptySheet } from '../../src/core/model';

function dataSheet(name: string): Sheet {
  const s = createEmptySheet(name);
  s.cells['A1'] = { raw: 'หมวด' };
  s.cells['B1'] = { raw: 'รายจ่าย' };
  s.cells['C1'] = { raw: 'รายรับ' };
  s.cells['A2'] = { raw: 'อาหาร' };
  s.cells['B2'] = { raw: '120' };
  s.cells['C2'] = { raw: '500' };
  s.cells['A3'] = { raw: 'เดินทาง' };
  s.cells['B3'] = { raw: '=B2/2' };
  s.cells['C3'] = { raw: '' };
  s.cells['A4'] = { raw: 'ที่พัก' };
  s.cells['B4'] = { raw: '300' };
  s.cells['C4'] = { raw: '200' };
  return s;
}

function chart(type: SheetChart['type'], title = 'สรุป & "ทดสอบ"'): SheetChart {
  return { id: 'c1', type, title, range: { r1: 0, c1: 0, r2: 3, c2: 2 }, anchor: { row: 1, col: 4, dx: 5, dy: 6 }, width: 480, height: 300 };
}

async function unzip(bytes: Uint8Array) {
  return JSZip.loadAsync(bytes);
}

describe('chart refs', () => {
  it('always quotes the sheet name and escapes quotes', () => {
    expect(chartRef('Data', 1, 1, 1, 3)).toBe("'Data'!$B$2:$B$4");
    expect(chartRef("ชีต'1", 0, 0)).toBe("'ชีต''1'!$A$1");
  });

  it('parses quoted, unquoted and no-sheet refs', () => {
    expect(parseChartRef("'ชีต''1'!$B$2:$B$4")).toEqual({ sheet: "ชีต'1", r1: 1, c1: 1, r2: 3, c2: 1 });
    expect(parseChartRef('Sheet1!A1')).toEqual({ sheet: 'Sheet1', r1: 0, c1: 0, r2: 0, c2: 0 });
    expect(parseChartRef('C5:A1')).toEqual({ sheet: null, r1: 0, c1: 0, r2: 4, c2: 2 });
    expect(parseChartRef('ไม่ใช่ช่วง')).toBeNull();
  });
});

describe('buildChartXml / parseChartXml round-trip', () => {
  const sheet = dataSheet("ข้อมูล ปี '69");
  const get = (k: string) => {
    const raw = sheet.cells[k]?.raw ?? '';
    return raw !== '' && !isNaN(Number(raw)) ? Number(raw) : raw;
  };

  for (const type of ['bar', 'line', 'pie'] as const) {
    it(`${type}: type, title and range come back; no formatting flag`, () => {
      const c = chart(type);
      const xml = buildChartXml(sheet.name, c, getChartData(c.range, get, (k) => String(get(k))));
      expect(() => parseXml(xml)).not.toThrow();
      expect(xml).not.toMatch(/spPr|txPr/);
      const parsed = parseChartXml(xml, sheet.name);
      // วงกลมใช้ชุดข้อมูลแรกชุดเดียว → ช่วงที่อ่านกลับแคบลงเหลือ A:B
      expect(parsed).toEqual({
        type,
        title: c.title,
        range: type === 'pie' ? { r1: 0, c1: 0, r2: 3, c2: 1 } : c.range,
        formatting: false,
      });
    });
  }

  it('without a header row or category column', () => {
    const c: SheetChart = { ...chart('line', ''), range: { r1: 1, c1: 1, r2: 3, c2: 1 } };
    const xml = buildChartXml(sheet.name, c, getChartData(c.range, get, (k) => String(get(k))));
    expect(xml).toContain('autoTitleDeleted val="1"');
    expect(parseChartXml(xml, sheet.name)).toMatchObject({ type: 'line', title: '', range: c.range });
  });

  it('rejects chart types and shapes the app does not support', () => {
    const c = chart('bar');
    const xml = buildChartXml(sheet.name, c, getChartData(c.range, get, (k) => String(get(k))));
    expect(parseChartXml(xml.replace('barDir val="col"', 'barDir val="bar"'), sheet.name)).toBeNull();
    expect(parseChartXml(xml.replace('grouping val="clustered"', 'grouping val="stacked"'), sheet.name)).toBeNull();
    expect(parseChartXml(xml.replace(/barChart/g, 'areaChart'), sheet.name)).toBeNull();
    // ช่วงข้ามชีต
    expect(parseChartXml(xml, 'ชีตอื่น')).toBeNull();
  });

  it('flags formatting (colors, data labels) on otherwise supported charts', () => {
    const c = chart('bar');
    const xml = buildChartXml(sheet.name, c, getChartData(c.range, get, (k) => String(get(k))));
    const colored = xml.replace('<c:invertIfNegative', '<c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></c:spPr><c:invertIfNegative');
    expect(parseChartXml(colored, sheet.name)?.formatting).toBe(true);
  });
});

describe('.xlsx with charts', () => {
  function workbook(): Workbook {
    const s = dataSheet('ข้อมูล');
    s.colWidths = { 0: 150 };
    s.charts = [chart('bar'), { ...chart('pie', 'วงกลม'), id: 'c2', anchor: { row: 20, col: 0, dx: 0, dy: 0 } }];
    const other = createEmptySheet('ชีต2');
    other.cells['A1'] = { raw: 'x' };
    other.cells['B1'] = { raw: '1' };
    other.charts = [{ ...chart('line', ''), id: 'c3', range: { r1: 0, c1: 0, r2: 0, c2: 1 } }];
    return { sheets: [s, other], activeSheet: 0 };
  }

  it('writes chart parts with cached values; ExcelJS can still read the file', async () => {
    const bytes = await exportToXlsxBuffer(workbook());
    const zip = await unzip(bytes);
    const files = Object.keys(zip.files);
    expect(files.filter((f) => /^xl\/charts\/chart\d+\.xml$/.test(f))).toHaveLength(3);
    expect(files.filter((f) => /^xl\/drawings\/drawing\d+\.xml$/.test(f))).toHaveLength(2);
    const ct = await zip.file('[Content_Types].xml')!.async('string');
    expect(ct).toContain('/xl/charts/chart1.xml');
    expect(ct).toContain('drawingml.chart+xml');
    const chart1 = await zip.file('xl/charts/chart1.xml')!.async('string');
    expect(chart1).toContain("'ข้อมูล'!$B$2:$B$4");
    // ค่าจากสูตร =B2/2 อยู่ใน cache, เซลว่างไม่มีจุดข้อมูล
    expect(chart1).toContain('<c:v>60</c:v>');
    expect(chart1).toContain('<c:v>เดินทาง</c:v>');
    const sheet1 = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
    expect(sheet1).toMatch(/<drawing r:id="rId\d+"\/>/);

    const excel = new ExcelJS.Workbook();
    await excel.xlsx.load(bytes as any);
    expect(excel.getWorksheet('ข้อมูล')!.getCell('B3').result).toBe(60);
  });

  it('round-trips type, title, range and position; no warning for app-written charts', async () => {
    const wb = await importFromXlsxBuffer(await exportToXlsxBuffer(workbook()));
    expect(wb.unsupportedFeatures).toBeUndefined();
    const [a, b] = wb.sheets[0].charts!;
    expect(a).toMatchObject({ type: 'bar', title: 'สรุป & "ทดสอบ"', range: { r1: 0, c1: 0, r2: 3, c2: 2 }, anchor: { row: 1, col: 4, dx: 5, dy: 6 }, width: 480, height: 300 });
    expect(b).toMatchObject({ type: 'pie', title: 'วงกลม', range: { r1: 0, c1: 0, r2: 3, c2: 1 }, anchor: { row: 20, col: 0, dx: 0, dy: 0 } });
    expect(wb.sheets[1].charts).toHaveLength(1);
    expect(wb.sheets[1].charts![0]).toMatchObject({ type: 'line', title: '', range: { r1: 0, c1: 0, r2: 0, c2: 1 } });
  });

  it('charts whose range was deleted are not written', async () => {
    const wb = workbook();
    wb.sheets[1].charts![0].range = null;
    const zip = await unzip(await exportToXlsxBuffer(wb));
    expect(Object.keys(zip.files).filter((f) => /^xl\/charts\/chart\d+\.xml$/.test(f))).toHaveLength(2);
  });

  it('workbooks without charts are written exactly as ExcelJS writes them', async () => {
    const zip = await unzip(await exportToXlsxBuffer({ sheets: [dataSheet('A')], activeSheet: 0 }));
    expect(Object.keys(zip.files).some((f) => f.startsWith('xl/charts/'))).toBe(false);
  });

  it('unsupported charts and chart formatting are reported (SPEC 6.4)', async () => {
    const bytes = await exportToXlsxBuffer(workbook());
    const zip = await unzip(bytes);
    const c1 = await zip.file('xl/charts/chart1.xml')!.async('string');
    zip.file('xl/charts/chart1.xml', c1.replace(/barChart/g, 'areaChart'));
    const c2 = await zip.file('xl/charts/chart2.xml')!.async('string');
    zip.file('xl/charts/chart2.xml', c2.replace('<c:varyColors val="1"/>', '<c:varyColors val="1"/><c:dLbls><c:showVal val="1"/></c:dLbls>'));
    const wb = await importFromXlsxBuffer(await zip.generateAsync({ type: 'uint8array' }));
    expect(wb.unsupportedFeatures).toEqual(expect.arrayContaining(['แผนภูมิ/กราฟ', 'การจัดรูปแบบกราฟ']));
    // กราฟแท่งที่กลายเป็นกราฟพื้นที่ไม่ถูกเก็บ เหลือวงกลม
    expect(wb.sheets[0].charts!.map((c) => c.type)).toEqual(['pie']);
  });

  it('chart parts not attached to a worksheet (e.g. chartsheet) are reported', async () => {
    const zip = await unzip(await exportToXlsxBuffer({ sheets: [dataSheet('A')], activeSheet: 0 }));
    zip.file('xl/charts/chart9.xml', '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>');
    const bytes = await zip.generateAsync({ type: 'uint8array' });
    const sheets = [dataSheet('A')];
    const r = await readChartsFromXlsx(bytes, sheets);
    expect(r.unsupported).toBe(true);
    expect(r.chartsBySheet).toEqual([[]]);
  });
});

describe('written package parts are well-formed XML', () => {
  it('every .xml and .rels part parses', async () => {
    const s = dataSheet("ชีต & 'อัญ' ประกาศ");
    s.charts = [chart('bar')];
    const zip = await unzip(await exportToXlsxBuffer({ sheets: [s], activeSheet: 0 }));
    const parts = Object.keys(zip.files).filter((n) => n.endsWith('.xml') || n.endsWith('.rels'));
    for (const f of parts) {
      const xml = await zip.file(f)!.async('string');
      expect(() => parseXml(xml), f).not.toThrow();
    }
    const chartXml = await zip.file('xl/charts/chart1.xml')!.async('string');
    expect(chartXml).toContain("'ชีต &amp; ''อัญ'' ประกาศ'!$B$2:$B$4");
  });
});
