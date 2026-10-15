import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { useSheetStore } from '../../src/ui/store';
import { exportToXlsxBuffer } from '../../src/core/xlsx';
import { Workbook, createEmptySheet } from '../../src/core/model';
import { serializeRecoveryData, deserializeRecoveryData } from '../../src/core/recovery';
import { buildTemplateWorkbook } from '../../src/templates';
import { th } from '../../src/i18n/th';

const S = () => useSheetStore.getState();

function fileWorkbook(): Workbook {
  const a = createEmptySheet('หน้าแรก');
  const secret = createEmptySheet('ซ่อนไว้');
  secret.hidden = 'hidden';
  secret.cells['A1'] = { raw: '50' };
  const b = createEmptySheet('รายการ');
  b.cells['A1'] = { raw: '=ซ่อนไว้!A1*2' };
  return { sheets: [a, secret, b], activeSheet: 2 };
}

describe('Phase 4B ขั้น 3: เปิดไฟล์หลายชีต / ชีตซ่อน', () => {
  beforeEach(async () => {
    S().newWorkbook();
    await S().loadWorkbook(await exportToXlsxBuffer(fileWorkbook()), 'C:\\งาน\\บัญชี.xlsx');
  });

  it('opens on the sheet that was active in the file', () => {
    expect(S().activeSheetIndex).toBe(2);
    expect(S().workbook.sheets[2].name).toBe('รายการ');
    expect(S().getComputedValue(0, 0)).toBe(100);
  });

  it('reports hidden sheets in the banner but not as unsupported', () => {
    expect(S().hiddenSheetNotice).toBe(1);
    expect(S().unsupportedFeatures).toEqual([]);
    expect(th.banner.hiddenSheets.replace('{count}', '1')).toBe('มีชีตซ่อน 1 ชีต ถูกเก็บไว้ตามเดิม');
  });

  it('skips hidden sheets when switching', () => {
    S().prevSheet();
    expect(S().activeSheetIndex).toBe(0);
    S().nextSheet();
    expect(S().activeSheetIndex).toBe(2);
    S().switchSheet(1);
    expect(S().activeSheetIndex).toBe(2);
  });

  it('cannot delete the last visible sheet while a hidden one remains', () => {
    S().deleteSheet(0);
    S().confirmDeleteSheet();
    expect(S().workbook.sheets.map((s) => s.name)).toEqual(['ซ่อนไว้', 'รายการ']);
    expect(S().activeSheetIndex).toBe(1);
    expect(S().deleteSheet(1)).toEqual({
      success: false,
      error: th.sheetTabs.cannotDeleteLastSheet,
    });
  });

  it('saving keeps the hidden sheet and the current active sheet', async () => {
    S().switchSheet(0);
    const { importFromXlsxBuffer } = await import('../../src/core/xlsx');
    const reloaded = await importFromXlsxBuffer(await S().saveWorkbook());
    expect(reloaded.sheets.map((s) => s.name)).toEqual(['หน้าแรก', 'ซ่อนไว้', 'รายการ']);
    expect(reloaded.sheets[1].hidden).toBe('hidden');
    expect(reloaded.activeSheet).toBe(0);
  });
});

describe('Phase 4B ขั้น 3: ไฟล์กู้คืนหลายชีต', () => {
  const saved: any[] = [];

  beforeEach(() => {
    saved.length = 0;
    (globalThis as any).window = {
      electronAPI: {
        saveRecovery: async (data: any) => {
          saved.push(JSON.parse(serializeRecoveryData(data)));
        },
        deleteRecovery: async () => {},
        setTitle: () => {},
      },
    };
    S().newWorkbook();
  });

  afterEach(() => {
    delete (globalThis as any).window;
  });

  it('snapshot stores every sheet and the active sheet, and restore reopens it', async () => {
    S().startEditing('หนึ่ง');
    S().commitEditing();
    S().addSheet();
    S().startEditing('=ชีต1!A1');
    S().commitEditing();
    S().addSheet();
    S().switchSheet(1);

    await S().saveRecoverySnapshot();
    expect(saved).toHaveLength(1);
    const data = deserializeRecoveryData(JSON.stringify(saved[0]))!;
    expect(data.workbook.sheets.map((s) => s.name)).toEqual(['ชีต1', 'ชีต2', 'ชีต3']);
    expect(data.workbook.activeSheet).toBe(1);

    S().newWorkbook();
    S().restoreRecoveryWorkbook(data.workbook, null);
    expect(S().workbook.sheets).toHaveLength(3);
    expect(S().activeSheetIndex).toBe(1);
    expect(S().getRawValue(0, 0)).toBe('=ชีต1!A1');
    expect(S().getComputedValue(0, 0)).toBe('หนึ่ง');
  });
});

describe('Phase 4B ขั้น 3: point mode ข้ามชีต (ระดับ store)', () => {
  beforeEach(() => {
    S().newWorkbook();
    S().addSheet(); // ชีต2
    S().setActiveCell(0, 0);
    S().startEditing('7');
    S().commitEditing();
    S().switchSheet(0);
    S().setActiveCell(2, 1); // B3 ของชีต1
  });

  it('switching tabs while typing a formula keeps editing; Enter commits into the original cell and returns', () => {
    S().startEditing('=');
    S().switchSheet(1);
    expect(S().isEditing).toBe(true);
    expect(S().activeSheetIndex).toBe(1);
    expect(S().editingSheetIndex).toBe(0);

    // Grid ใส่ reference พร้อมชื่อชีตเมื่อคลิกเซลบนชีตอื่น
    S().setEditingValue(S().editingValue + "'ชีต2'!A1");
    S().setEditingValue(S().editingValue + '*3');
    S().commitEditing(1, 0); // Enter

    expect(S().activeSheetIndex).toBe(0);
    expect(S().isEditing).toBe(false);
    expect(S().workbook.sheets[0].cells['B3'].raw).toBe("='ชีต2'!A1*3");
    expect(S().workbook.sheets[1].cells['B3']).toBeUndefined();
    expect(S().getComputedValue(2, 1)).toBe(21);
    expect(S().activeCell).toEqual({ r: 3, c: 1 });
  });

  it('Esc cancels and returns to the original sheet and cell', () => {
    S().startEditing('=SUM(');
    S().switchSheet(1);
    S().setEditingValue("=SUM('ชีต2'!A1:A3");
    S().cancelEditing();
    expect(S().activeSheetIndex).toBe(0);
    expect(S().activeCell).toEqual({ r: 2, c: 1 });
    expect(S().isEditing).toBe(false);
    expect(S().workbook.sheets[0].cells['B3']).toBeUndefined();
  });

  it('undo of a cross-sheet point-mode edit targets the original sheet', () => {
    S().startEditing('=');
    S().switchSheet(1);
    S().setEditingValue("='ชีต2'!A1");
    S().commitEditing(0, 0);
    S().switchSheet(1);
    S().undo();
    expect(S().activeSheetIndex).toBe(0);
    expect(S().workbook.sheets[0].cells['B3']).toBeUndefined();
  });

  it('plain text editing is committed before switching sheets', () => {
    S().startEditing('ข้อความ');
    S().switchSheet(1);
    expect(S().isEditing).toBe(false);
    expect(S().activeSheetIndex).toBe(1);
    expect(S().workbook.sheets[0].cells['B3'].raw).toBe('ข้อความ');
  });
});

describe('Phase 4B ขั้น 3: เทมเพลตนิยามได้หลายชีต', () => {
  it('builds a multi-sheet template workbook and loads it', () => {
    const data = createEmptySheet('บันทึก');
    data.cells['A1'] = { raw: '10' };
    const summary = createEmptySheet('สรุป');
    summary.cells['A1'] = { raw: '=บันทึก!A1*2' };
    const wb = buildTemplateWorkbook([data, summary], 1);
    expect(wb.activeSheet).toBe(1);

    S().loadTemplate(wb, 'ทดสอบ');
    expect(S().workbook.sheets.map((s) => s.name)).toEqual(['บันทึก', 'สรุป']);
    expect(S().activeSheetIndex).toBe(1);
    expect(S().getComputedValue(0, 0)).toBe(20);
  });

  it('rejects duplicate sheet names in a template definition', () => {
    expect(() => buildTemplateWorkbook([createEmptySheet('ก'), createEmptySheet('ก')])).toThrow();
  });
});
