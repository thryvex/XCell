import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import ExcelJS from 'exceljs';
import { exportToXlsxBuffer, importFromXlsxBuffer } from '../../src/core/xlsx';
import { buildChartXml, parseChartXml, readChartsFromXlsx, xlsxChartRef, hasUnkeptFormatting } from '../../src/core/chartXlsx';
import { newChart, workbookChartContext, EMU_PER_PX } from '../../src/core/chart';
import { Evaluator } from '../../src/core/evaluator';
import { parseXml, MAX_XML_LENGTH } from '../../src/core/xmlTree';
import { Workbook, Sheet, SheetChart, createEmptySheet } from '../../src/core/model';
import { dateToSerial } from '../../src/core/numfmt';
import { th } from '../../src/i18n/th';
import { useSheetStore } from '../../src/ui/store';

const C_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const XDR_NS = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';

function dataSheet(name: string): Sheet {
  const s = createEmptySheet(name);
  const put = (k: string, raw: string, fmt?: Sheet['cells'][string]['fmt']) => (s.cells[k] = fmt ? { raw, fmt } : { raw });
  put('A1', 'หมวด');
  put('B1', 'รายจ่าย');
  put('C1', 'รายรับ');
  put('A2', 'อาหาร');
  put('B2', '120', 'baht');
  put('C2', '500');
  put('A3', 'เดินทาง');
  put('B3', '=B2/2');
  put('C3', '');
  put('A4', 'ที่พัก');
  put('B4', 'ไม่ทราบ');
  put('C4', '200');
  return s;
}

/** กราฟที่ใช้ทุกฟิลด์ */
function fullChart(sheetName: string, type: SheetChart['type'] = 'bar'): SheetChart {
  const q = `'${sheetName.replace(/'/g, "''")}'`;
  return {
    ...newChart(
      type,
      [
        { name: { ref: `${q}!B1` }, categories: `${q}!A2:A4`, values: `${q}!B2:B4`, color: '#4E79A7' },
        { name: { text: 'รับ <เงิน> & "สด"' }, categories: `${q}!A2:A4`, values: `${q}!C2:C4`, color: '#E15759' },
      ],
      { row: 1, col: 4, dx: 5, dy: 6 }
    ),
    title: 'สรุป & <ทดสอบ> "ไทย"',
    xTitle: 'หมวด',
    yTitle: 'บาท',
    dataLabels: true,
    legend: 'right',
  };
}

const ctxOf = (wb: Workbook, forFile = true) => workbookChartContext(wb, new Evaluator(wb), forFile);
const stripIds = (charts: SheetChart[] | undefined) => (charts ?? []).map(({ id: _id, ...rest }) => rest);
const unzip = (bytes: Uint8Array) => JSZip.loadAsync(bytes);
const chartFiles = (zip: JSZip) => Object.keys(zip.files).filter((f) => /^xl\/charts\/chart\d+\.xml$/.test(f)).sort();

// ---------------------------------------------------------------------------
// ตัวเขียน
// ---------------------------------------------------------------------------

describe('ตัวเขียน chart XML', () => {
  const s = dataSheet("ข้อมูล & 'ร้าน' ป้า");
  const wb: Workbook = { sheets: [s], activeSheet: 0 };

  it("reference ใส่ '...' รอบชื่อชีตเสมอ และใช้ $", () => {
    expect(xlsxChartRef({ sheet: 'Data', r1: 1, c1: 1, r2: 3, c2: 1 })).toBe("'Data'!$B$2:$B$4");
    expect(xlsxChartRef({ sheet: "ชีต'1", r1: 0, c1: 0, r2: 0, c2: 0 })).toBe("'ชีต''1'!$A$1");
  });

  for (const type of ['bar', 'line', 'pie'] as const) {
    it(`${type}: เป็น XML ที่ถูกต้อง มีองค์ประกอบที่จำเป็น และ escape ข้อความ`, () => {
      const xml = buildChartXml(fullChart(s.name, type), ctxOf(wb))!;
      const root = parseXml(xml);
      expect(root.local).toBe('chartSpace');
      expect(xml).toContain('<c:plotVisOnly val="1"/>');
      expect(xml).toContain('<c:dispBlanksAs val="gap"/>');
      expect(xml).toContain(`<c:${type === 'bar' ? 'barChart' : type === 'line' ? 'lineChart' : 'pieChart'}>`);
      expect(xml).toContain('สรุป &amp; &lt;ทดสอบ&gt; &quot;ไทย&quot;');
      expect(xml).toContain("<c:f>'ข้อมูล &amp; ''ร้าน'' ป้า'!$B$2:$B$4</c:f>");
      expect(xml).toContain('<c:legendPos val="r"/>');
      expect(xml).toContain('<c:showVal val="1"/>');
      if (type === 'bar') {
        expect(xml).toContain('<c:barDir val="col"/><c:grouping val="clustered"/>');
        expect(xml).toContain('<a:srgbClr val="4E79A7"/>');
        expect(xml).toContain('<a:srgbClr val="E15759"/>');
        expect(xml).toContain('<a:t>บาท</a:t>');
        expect(xml).toContain('รับ &lt;เงิน&gt; &amp; &quot;สด&quot;');
      }
      if (type === 'pie') {
        // ชุดแรกชุดเดียว สีชิ้นจากจานสี
        expect(xml.match(/<c:ser>/g)).toHaveLength(1);
        expect(xml.match(/<c:dPt>/g)).toHaveLength(3);
        expect(xml).not.toContain('<c:catAx>');
      }
    });
  }

  it('ค่า cache ตรงกับค่าที่คำนวณ: สูตร, ข้อความ = 0, ช่องว่างไม่มีจุด, ป้ายวันที่เป็น ค.ศ.', () => {
    const d = dataSheet('D');
    d.cells.A2 = { raw: String(dateToSerial(new Date(2026, 0, 15))), fmt: 'date' };
    const x = buildChartXml(fullChart('D'), ctxOf({ sheets: [d], activeSheet: 0 }))!;
    const ser = x.split('<c:ser>');
    expect(ser[1]).toContain('<c:pt idx="0"><c:v>120</c:v></c:pt><c:pt idx="1"><c:v>60</c:v></c:pt><c:pt idx="2"><c:v>0</c:v></c:pt>');
    expect(ser[2]).toContain('<c:ptCount val="3"/><c:pt idx="0"><c:v>500</c:v></c:pt><c:pt idx="2"><c:v>200</c:v></c:pt>');
    expect(ser[1]).toContain('<c:v>รายจ่าย</c:v>');
    expect(ser[1]).toMatch(/<c:v>[^<]*2026[^<]*<\/c:v>/);
    expect(ser[1]).toContain('<c:v>เดินทาง</c:v>');
  });

  it('ชุดที่ reference หายไม่ถูกเขียน; หายทุกชุด = ไม่เขียนกราฟ', () => {
    const c = fullChart(s.name);
    c.series[0].values = '#REF!';
    expect(buildChartXml(c, ctxOf(wb))!.match(/<c:ser>/g)).toHaveLength(1);
    c.series[1].values = '#REF!';
    expect(buildChartXml(c, ctxOf(wb))).toBeNull();
  });

  it('ไม่มี legend เมื่อเลือก "ไม่แสดง"; ไม่มีชื่อกราฟ = autoTitleDeleted', () => {
    const c = { ...fullChart(s.name), legend: 'none' as const, title: '', dataLabels: false };
    const xml = buildChartXml(c, ctxOf(wb))!;
    expect(xml).not.toContain('<c:legend>');
    expect(xml).toContain('<c:autoTitleDeleted val="1"/>');
    expect(xml).not.toContain('<c:dLbls>');
  });
});

describe('.xlsx ที่มีกราฟ', () => {
  function workbook(): Workbook {
    const data = dataSheet('ข้อมูล ปี 69');
    data.colWidths = { 0: 150, 4: 60, 5: 200 };
    data.rowHeights = { 2: 50 };
    data.charts = [fullChart(data.name), { ...fullChart(data.name, 'pie'), anchor: { row: 20, col: 0, dx: 0, dy: 0 }, legend: 'bottom' }];
    const summary = createEmptySheet('สรุป');
    // อ้างข้อมูลข้ามชีต
    summary.charts = [
      {
        ...newChart('line', [{ name: { text: '' }, values: "'ข้อมูล ปี 69'!C2:C4", color: '#59A14F' }], { row: 0, col: 0, dx: 0, dy: 0 }),
        width: 333,
        height: 222,
      },
    ];
    return { sheets: [data, summary], activeSheet: 0 };
  }

  it('เขียนส่วนกราฟครบ ชื่อไฟล์ไม่ชน และ ExcelJS ยังอ่านไฟล์ได้', async () => {
    const bytes = await exportToXlsxBuffer(workbook());
    const zip = await unzip(bytes);
    expect(chartFiles(zip)).toEqual(['xl/charts/chart1.xml', 'xl/charts/chart2.xml', 'xl/charts/chart3.xml']);
    expect(Object.keys(zip.files).filter((f) => /^xl\/drawings\/drawing\d+\.xml$/.test(f))).toHaveLength(2);
    const ct = await zip.file('[Content_Types].xml')!.async('string');
    for (const f of chartFiles(zip)) expect(ct).toContain(`PartName="/${f}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"`);
    expect(ct).toContain('PartName="/xl/drawings/drawing1.xml"');
    const sheet1 = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
    expect(sheet1).toMatch(/<drawing r:id="rId\d+"\/>/);
    const rels = await zip.file('xl/worksheets/_rels/sheet1.xml.rels')!.async('string');
    expect(rels).toContain('Target="../drawings/drawing1.xml"');
    const drawing = await zip.file('xl/drawings/drawing1.xml')!.async('string');
    expect(drawing.match(/<xdr:twoCellAnchor editAs="oneCell">/g)).toHaveLength(2);
    // ทุก part เป็น XML ที่ถูกต้อง
    for (const f of Object.keys(zip.files).filter((n) => n.endsWith('.xml') || n.endsWith('.rels'))) {
      const xml = await zip.file(f)!.async('string');
      expect(() => parseXml(xml), f).not.toThrow();
    }
    const excel = new ExcelJS.Workbook();
    await excel.xlsx.load(bytes as never);
    expect(excel.getWorksheet('ข้อมูล ปี 69')!.getCell('B3').result).toBe(60);
  });

  it('anchor ในไฟล์คิดจากความกว้างคอลัมน์/ความสูงแถวจริง', async () => {
    const zip = await unzip(await exportToXlsxBuffer(workbook()));
    const drawing = parseXml(await zip.file('xl/drawings/drawing1.xml')!.async('string'));
    const first = drawing.children[0];
    const pos = (tag: string) => {
      const n = first.children.find((c) => c.uri === XDR_NS && c.local === tag)!;
      const v = (l: string) => Number(n.children.find((c) => c.local === l)!.text);
      return { col: v('col'), colOff: v('colOff'), row: v('row'), rowOff: v('rowOff') };
    };
    expect(pos('from')).toEqual({ col: 4, colOff: 5 * EMU_PER_PX, row: 1, rowOff: 6 * EMU_PER_PX });
    // คิดจากขนาดที่ถูกเขียนลงไฟล์: คอลัมน์ E 60 px → 8 หน่วย = 64 px, แถว 3 50 px → 38 pt = 51 px
    // กว้าง 480: E (64-5=59) + F (200) + G (110) = 369, เหลือ 111 → H (110) → เหลือ 1 ใน I
    // สูง 300: แถว 2 (28-6=22) + แถว 3 (51) + 28×8 = 297, เหลือ 3 ในแถว 12
    expect(pos('to')).toEqual({ col: 8, colOff: 1 * EMU_PER_PX, row: 11, rowOff: 3 * EMU_PER_PX });
  });

  it('round-trip: ได้กราฟเท่าเดิมทุกฟิลด์ (รวมอ้างข้ามชีต) และไม่มีแถบเตือน', async () => {
    const original = workbook();
    const wb = await importFromXlsxBuffer(await exportToXlsxBuffer(original));
    expect(wb.unsupportedFeatures).toBeUndefined();
    expect(stripIds(wb.sheets[0].charts)).toEqual(stripIds(original.sheets[0].charts).map((c, i) =>
      // วงกลมเขียนเฉพาะชุดแรก
      i === 1 ? { ...c, series: c.series.slice(0, 1), xTitle: '', yTitle: '' } : c
    ));
    expect(stripIds(wb.sheets[1].charts)).toEqual(stripIds(original.sheets[1].charts));
  });

  it('เปิดแล้วบันทึกซ้ำ 3 รอบ กราฟไม่เปลี่ยน', async () => {
    let wb = await importFromXlsxBuffer(await exportToXlsxBuffer(workbook()));
    const first = wb.sheets.map((s) => stripIds(s.charts));
    for (let i = 0; i < 3; i++) wb = await importFromXlsxBuffer(await exportToXlsxBuffer(wb));
    expect(wb.sheets.map((s) => stripIds(s.charts))).toEqual(first);
  });

  it('ไฟล์ที่ไม่มีกราฟเขียนตามเดิม ไม่มีส่วนกราฟ', async () => {
    const zip = await unzip(await exportToXlsxBuffer({ sheets: [dataSheet('A')], activeSheet: 0 }));
    expect(Object.keys(zip.files).some((f) => f.startsWith('xl/charts/') || f.startsWith('xl/drawings/'))).toBe(false);
  });
});

// ---------------------------------------------------------------------------
// ตัวอ่าน — XML ที่สร้างเลียนแบบไฟล์จาก Google Sheets / Excel (ยังไม่ได้ทดสอบกับไฟล์จริง)
// ---------------------------------------------------------------------------

const SPACE_OPEN = `<c:chartSpace xmlns:c="${C_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}">`;

/** แบบ Google Sheets: ไม่มี grouping, ชื่อชีตไม่มี '...', สี srgbClr, ฟอนต์ใน rich text */
const GOOGLE_BAR = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
${SPACE_OPEN}<c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr lvl="0"><a:defRPr b="0"><a:solidFill><a:srgbClr val="757575"/></a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr></a:pPr><a:r><a:rPr b="0"><a:solidFill><a:srgbClr val="757575"/></a:solidFill></a:rPr><a:t>ยอดขาย</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>
<c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>Data!$B$1</c:f></c:strRef></c:tx><c:spPr><a:solidFill><a:srgbClr val="4285F4"/></a:solidFill><a:ln cmpd="sng"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></c:spPr><c:cat><c:strRef><c:f>Data!$A$2:$A$4</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>Data!$B$2:$B$4</c:f></c:numRef></c:val></c:ser>
<c:axId val="1"/><c:axId val="2"/></c:barChart>
<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>หมวด</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr lvl="0"><a:defRPr b="0"/></a:pPr><a:endParaRPr lang="en"/></a:p></c:txPr><c:crossAx val="2"/></c:catAx>
<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines><c:spPr><a:ln><a:solidFill><a:srgbClr val="B7B7B7"/></a:solidFill></a:ln></c:spPr></c:majorGridlines><c:numFmt formatCode="General" sourceLinked="1"/><c:crossAx val="1"/></c:valAx></c:plotArea>
<c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/></c:chart></c:chartSpace>`;

/** แบบ Excel: mc:AlternateContent, schemeClr, extLst ใน ser, dLbls ปิด, grouping clustered */
const EXCEL_LINE = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="${C_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:c16r2="http://schemas.microsoft.com/office/drawing/2015/06/chart"><c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/>
<mc:AlternateContent><mc:Choice Requires="c14" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"><c14:style val="102"/></mc:Choice><mc:Fallback><c:style val="2"/></mc:Fallback></mc:AlternateContent>
<c:chart><c:autoTitleDeleted val="0"/><c:plotArea><c:layout/><c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>
<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>'ข้อมูล ปี 69'!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>ยอด</c:v></c:pt></c:strCache></c:strRef></c:tx><c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:round/></a:ln><a:effectLst/></c:spPr><c:marker><c:symbol val="none"/></c:marker>
<c:cat><c:strRef><c:f>'ข้อมูล ปี 69'!$A$2:$A$4</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>'ข้อมูล ปี 69'!$B$2:$B$4</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/></c:numCache></c:numRef></c:val><c:smooth val="0"/>
<c:extLst><c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}" xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart"><c16:uniqueId val="{00000000-0001}"/></c:ext></c:extLst></c:ser>
<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls><c:smooth val="0"/><c:axId val="11"/><c:axId val="12"/></c:lineChart>
<c:catAx><c:axId val="11"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="12"/></c:catAx><c:valAx><c:axId val="12"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="11"/></c:valAx>
<c:spPr><a:noFill/></c:spPr></c:plotArea><c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>
<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr></c:chartSpace>`;

const EXCEL_PIE = `${SPACE_OPEN}<c:chart><c:title><c:tx><c:strRef><c:f>Data!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>ยอดขาย</c:v></c:pt></c:strCache></c:strRef></c:tx><c:overlay val="0"/></c:title><c:plotArea><c:layout/><c:pieChart><c:varyColors val="1"/>
<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>ชิ้น</c:v></c:tx><c:dPt><c:idx val="0"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></c:spPr></c:dPt>
<c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>
<c:cat><c:strRef><c:f>Data!$A$2:$A$4</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>Data!$B$2:$B$4</c:f></c:numRef></c:val></c:ser><c:firstSliceAng val="0"/></c:pieChart></c:plotArea><c:plotVisOnly val="1"/></c:chart></c:chartSpace>`;

describe('ตัวอ่าน chart XML (เลียนแบบ Google Sheets / Excel)', () => {
  it('Google Sheets: แท่ง ไม่มี grouping, ชื่อชีตไม่มี quote, สีและชื่อแกน', () => {
    expect(parseChartXml(GOOGLE_BAR, 'Data')).toEqual({
      type: 'bar',
      title: 'ยอดขาย',
      series: [{ name: { ref: 'Data!B1' }, categories: 'Data!A2:A4', values: 'Data!B2:B4', color: '#4285F4' }],
      legend: 'right',
      xTitle: 'หมวด',
      yTitle: '',
      dataLabels: false,
    });
  });

  it('Excel: เส้น ชื่อชีตไทยมีเว้นวรรค, สีจากธีม Office → จานสีของแอป, ไม่มี legend เป็นล่าง', () => {
    const c = parseChartXml(EXCEL_LINE, 'อื่น')!;
    expect(c.type).toBe('line');
    expect(c.series[0]).toEqual({
      name: { ref: "'ข้อมูล ปี 69'!B1" },
      categories: "'ข้อมูล ปี 69'!A2:A4",
      values: "'ข้อมูล ปี 69'!B2:B4",
      color: '#4E79A7',
    });
    expect(c.legend).toBe('bottom');
    expect(c.dataLabels).toBe(false);
  });

  it('Excel: วงกลม ชื่อกราฟจากเซล (ใช้ค่า cache), ชื่อชุดเป็นข้อความ, ตัวเลขบนกราฟ, ไม่มี legend', () => {
    const c = parseChartXml(EXCEL_PIE, 'Data')!;
    expect(c).toMatchObject({ type: 'pie', title: 'ยอดขาย', legend: 'none', dataLabels: true });
    expect(c.series[0].name).toEqual({ text: 'ชิ้น' });
  });

  it('reference ที่ไม่มีชื่อชีต = ชีตที่กราฟอยู่', () => {
    const xml = GOOGLE_BAR.replace(/Data!/g, '');
    expect(parseChartXml(xml, 'ชีต 1')!.series[0].values).toBe("'ชีต 1'!B2:B4");
  });

  const outOfScope: Array<[string, string]> = [
    ['แท่งแนวนอน', GOOGLE_BAR.replace('<c:barDir val="col"/>', '<c:barDir val="bar"/>')],
    ['stacked', GOOGLE_BAR.replace('<c:barDir val="col"/>', '<c:barDir val="col"/><c:grouping val="stacked"/>')],
    ['percentStacked', GOOGLE_BAR.replace('<c:barDir val="col"/>', '<c:barDir val="col"/><c:grouping val="percentStacked"/>')],
    ['เส้น stacked', EXCEL_LINE.replace('<c:grouping val="standard"/>', '<c:grouping val="stacked"/>')],
    ['combo (แท่ง + เส้น)', GOOGLE_BAR.replace('</c:barChart>', '</c:barChart><c:lineChart><c:grouping val="standard"/><c:ser><c:idx val="1"/><c:order val="1"/><c:val><c:numRef><c:f>Data!$C$2:$C$4</c:f></c:numRef></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:lineChart>')],
    ['scatter', GOOGLE_BAR.replace(/barChart/g, 'scatterChart')],
    ['area', GOOGLE_BAR.replace(/barChart/g, 'areaChart')],
    ['3D', GOOGLE_BAR.replace(/barChart/g, 'bar3DChart')],
    ['doughnut', EXCEL_PIE.replace(/pieChart/g, 'doughnutChart')],
    ['trendline', EXCEL_LINE.replace('<c:smooth val="0"/>\n<c:extLst>', '<c:trendline><c:trendlineType val="linear"/></c:trendline><c:smooth val="0"/>\n<c:extLst>')],
    ['แกนรอง', GOOGLE_BAR.replace('</c:plotArea>', '<c:valAx><c:axId val="3"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="r"/><c:crossAx val="1"/></c:valAx></c:plotArea>')],
    ['ค่าเป็นตัวเลขคงที่ (numLit)', GOOGLE_BAR.replace('<c:numRef><c:f>Data!$B$2:$B$4</c:f></c:numRef>', '<c:numLit><c:ptCount val="1"/></c:numLit>')],
    ['ช่วงค่าสองมิติ', GOOGLE_BAR.replace('Data!$B$2:$B$4', 'Data!$B$2:$C$4')],
    ['ช่วงหลายส่วน', GOOGLE_BAR.replace('Data!$B$2:$B$4', '(Data!$B$2:$B$3,Data!$B$5)')],
    ['อ้างไฟล์อื่น', GOOGLE_BAR.replace('Data!$B$2:$B$4', '[1]Data!$B$2:$B$4')],
  ];
  for (const [name, xml] of outOfScope) {
    it(`เกินขอบเขต: ${name} → ไม่นำเข้า`, () => {
      expect(parseChartXml(xml, 'Data')).toBeNull();
    });
  }

  it('ปฏิเสธ XML ที่มี DOCTYPE / entity และ XML ที่ใหญ่เกินไป', () => {
    const doctype = `<?xml version="1.0"?><!DOCTYPE x [<!ENTITY big "aaaa">]>${GOOGLE_BAR.split('?>')[1].replace('ยอดขาย', '&big;')}`;
    expect(() => parseXml(doctype)).toThrow();
    expect(() => parseChartXml(doctype, 'Data')).toThrow();
    expect(() => parseXml(GOOGLE_BAR.replace('ยอดขาย', '&undefinedEntity;'))).toThrow();
    expect(() => parseXml(`<x>${'a'.repeat(MAX_XML_LENGTH)}</x>`)).toThrow();
    // entity มาตรฐานยังใช้ได้
    expect(parseChartXml(GOOGLE_BAR.replace('ยอดขาย', 'A &amp; B'), 'Data')!.title).toBe('A & B');
  });
});

// ---------------------------------------------------------------------------
// อ่านทั้งไฟล์: แทรกกราฟแบบไฟล์ภายนอกลงไฟล์ที่ไม่มีกราฟ
// ---------------------------------------------------------------------------

async function withExternalCharts(base: Workbook, charts: Array<{ xml: string; anchor: string }>, sheetPart = 'xl/worksheets/sheet1.xml') {
  const zip = await unzip(await exportToXlsxBuffer(base));
  const anchors: string[] = [];
  const rels: string[] = [];
  charts.forEach((c, i) => {
    zip.file(`xl/charts/chart${i + 1}.xml`, c.xml);
    rels.push(`<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart${i + 1}.xml"/>`);
    anchors.push(c.anchor.replace('RID', `rId${i + 1}`));
  });
  zip.file('xl/drawings/drawing1.xml', `<xdr:wsDr xmlns:xdr="${XDR_NS}" xmlns:a="${A_NS}">${anchors.join('')}</xdr:wsDr>`);
  zip.file('xl/drawings/_rels/drawing1.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join('')}</Relationships>`);
  const relsPath = sheetPart.replace('worksheets/', 'worksheets/_rels/') + '.rels';
  const old = await zip.file(relsPath)?.async('string');
  const drawingRel = '<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>';
  zip.file(relsPath, old ? old.replace('</Relationships>', `${drawingRel}</Relationships>`) : `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${drawingRel}</Relationships>`);
  return zip.generateAsync({ type: 'uint8array' });
}

const graphic = `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${C_NS}"><c:chart xmlns:c="${C_NS}" xmlns:r="${R_NS}" r:id="RID"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/>`;
const ONE_CELL = `<xdr:oneCellAnchor><xdr:from><xdr:col>3</xdr:col><xdr:colOff>95250</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="${600 * EMU_PER_PX}" cy="${371 * EMU_PER_PX}"/>${graphic}</xdr:oneCellAnchor>`;
const TWO_CELL = `<xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>20</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>${graphic}</xdr:twoCellAnchor>`;

describe('เปิดไฟล์ที่มีกราฟจากภายนอก (เลียนแบบ)', () => {
  const base = (): Workbook => ({ sheets: [dataSheet('Data')], activeSheet: 0 });

  it('นำเข้ากราฟที่รองรับพร้อมตำแหน่ง/ขนาด ไม่มีแถบเตือน', async () => {
    const bytes = await withExternalCharts(base(), [
      { xml: GOOGLE_BAR, anchor: ONE_CELL },
      { xml: EXCEL_PIE, anchor: TWO_CELL },
    ]);
    const wb = await importFromXlsxBuffer(bytes);
    expect(wb.unsupportedFeatures).toBeUndefined();
    const [bar, pie] = wb.sheets[0].charts!;
    expect(bar).toMatchObject({ type: 'bar', anchor: { row: 2, col: 3, dx: 10, dy: 0 }, width: 600, height: 371 });
    expect(pie).toMatchObject({ type: 'pie', anchor: { row: 10, col: 0, dx: 0, dy: 0 }, width: 4 * 110, height: 10 * 28 });
  });

  it('กราฟที่เกินขอบเขต: ไม่นำเข้า + แถบเตือน "กราฟชนิดที่แอปไม่รองรับ จะหายเมื่อบันทึก"', async () => {
    const bytes = await withExternalCharts(base(), [
      { xml: GOOGLE_BAR, anchor: ONE_CELL },
      { xml: GOOGLE_BAR.replace(/barChart/g, 'scatterChart'), anchor: TWO_CELL },
    ]);
    const wb = await importFromXlsxBuffer(bytes);
    expect(wb.unsupportedFeatures).toContain(th.chart.unsupportedInFile);
    expect(th.chart.unsupportedInFile).toBe('กราฟชนิดที่แอปไม่รองรับ จะหายเมื่อบันทึก');
    expect(wb.sheets[0].charts!.map((c) => c.type)).toEqual(['bar']);
  });

  it('กราฟที่มี DOCTYPE ถูกรายงานเป็นไม่รองรับ (ไม่ crash)', async () => {
    const evil = `<?xml version="1.0"?><!DOCTYPE c:chartSpace [<!ENTITY x "y">]>${GOOGLE_BAR.split('?>')[1]}`;
    const wb = await importFromXlsxBuffer(await withExternalCharts(base(), [{ xml: evil, anchor: ONE_CELL }]));
    expect(wb.unsupportedFeatures).toContain(th.chart.unsupportedInFile);
    expect(wb.sheets[0].charts).toBeUndefined();
  });

  it('ส่วนกราฟที่ไม่ได้ผูกกับชีต (เช่น ชีตกราฟแยก) ถูกรายงาน', async () => {
    const zip = await unzip(await exportToXlsxBuffer(base()));
    zip.file('xl/charts/chart9.xml', GOOGLE_BAR);
    const r = await readChartsFromXlsx(await zip.generateAsync({ type: 'uint8array' }), [dataSheet('Data')]);
    expect(r.unsupported).toBe(true);
    expect(r.chartsBySheet).toEqual([[]]);
  });

  it('เปิดไฟล์ภายนอกแล้วบันทึกกลับ กราฟที่รองรับไม่หาย/ไม่เปลี่ยน', async () => {
    const opened = await importFromXlsxBuffer(await withExternalCharts(base(), [{ xml: GOOGLE_BAR, anchor: ONE_CELL }]));
    const again = await importFromXlsxBuffer(await exportToXlsxBuffer(opened));
    expect(stripIds(again.sheets[0].charts)).toEqual(stripIds(opened.sheets[0].charts));
  });
});

describe('แถบข้อมูล "การจัดรูปแบบกราฟ"', () => {
  it('กราฟที่แอปเขียนเองทุกชนิดและทุกตัวเลือกไม่ถูกนับ', () => {
    const s = dataSheet('D');
    const wb: Workbook = { sheets: [s], activeSheet: 0 };
    for (const type of ['bar', 'line', 'pie'] as const) {
      for (const dataLabels of [true, false]) {
        const xml = buildChartXml({ ...fullChart('D', type), dataLabels }, ctxOf(wb))!;
        expect(hasUnkeptFormatting(parseXml(xml), type), `${type} ${dataLabels}`).toBe(false);
      }
    }
  });

  it('ฟอนต์/ขอบ (Google Sheets), สีจากธีม Office (Excel), ตัวเลขบนกราฟแบบเปอร์เซ็นต์ ถูกนับ', () => {
    expect(hasUnkeptFormatting(parseXml(GOOGLE_BAR), 'bar')).toBe(true);
    expect(hasUnkeptFormatting(parseXml(EXCEL_LINE), 'line')).toBe(true);
    const plainPie = EXCEL_PIE.replace(/<c:dPt>.*?<\/c:dPt>/s, '');
    expect(hasUnkeptFormatting(parseXml(plainPie), 'pie')).toBe(false);
    expect(hasUnkeptFormatting(parseXml(plainPie.replace('<c:showPercent val="0"/>', '<c:showPercent val="1"/>')), 'pie')).toBe(true);
    expect(hasUnkeptFormatting(parseXml(EXCEL_PIE), 'pie')).toBe(true); // สีชิ้นจากธีม Office
  });

  it('เปิดไฟล์: ตั้งธง chartFormatting เฉพาะเมื่อมีกราฟที่จัดรูปแบบ ไม่ใช่แถบเตือน 6.4', async () => {
    const base: Workbook = { sheets: [dataSheet('Data')], activeSheet: 0 };
    const ext = await importFromXlsxBuffer(await withExternalCharts(base, [{ xml: GOOGLE_BAR, anchor: ONE_CELL }]));
    expect(ext.chartFormatting).toBe(true);
    expect(ext.unsupportedFeatures).toBeUndefined();
    const own = createEmptySheet('Data');
    Object.assign(own.cells, dataSheet('Data').cells);
    own.charts = [fullChart('Data')];
    const back = await importFromXlsxBuffer(await exportToXlsxBuffer({ sheets: [own], activeSheet: 0 }));
    expect(back.chartFormatting).toBeUndefined();
  });
});

describe('store: แถบข้อมูลการจัดรูปแบบกราฟ', () => {
  it('ขึ้นเมื่อเปิดไฟล์ที่มีกราฟจัดรูปแบบ ปิดได้ และงานใหม่ไม่มีแถบ', async () => {
    const S = useSheetStore.getState;
    const bytes = await withExternalCharts({ sheets: [dataSheet('Data')], activeSheet: 0 }, [{ xml: GOOGLE_BAR, anchor: ONE_CELL }]);
    await S().loadWorkbook(bytes, 'C:\ไฟล์\กราฟ.xlsx');
    expect(S().chartFormattingNotice).toBe(true);
    expect(S().unsupportedFeatures).toEqual([]);
    expect(S().workbook.sheets[0].charts).toHaveLength(1);
    S().dismissChartFormattingNotice();
    expect(S().chartFormattingNotice).toBe(false);
    await S().loadWorkbook(bytes, 'x.xlsx');
    S().newWorkbook();
    expect(S().chartFormattingNotice).toBe(false);
  });
});
