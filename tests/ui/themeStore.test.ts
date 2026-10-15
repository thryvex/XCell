import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import ExcelJS from 'exceljs';
import { useSheetStore, selectResolvedTheme } from '../../src/ui/store';
import { exportToXlsxBuffer, importFromXlsxBuffer } from '../../src/core/xlsx';

describe('ธีมสีใน store (5C)', () => {
  let setPreferences: ReturnType<typeof vi.fn>;
  let saveRecovery: ReturnType<typeof vi.fn>;

  beforeEach(() => {
    setPreferences = vi.fn(() => Promise.resolve({ success: true }));
    saveRecovery = vi.fn(() => Promise.resolve({ success: true }));
    (globalThis as any).window = { electronAPI: { setPreferences, saveRecovery, setTitle: () => {} } };
    useSheetStore.getState().newWorkbook();
    useSheetStore.setState({ themePref: 'system', gridWhite: false, systemPrefersDark: false });
  });

  afterEach(() => {
    delete (globalThis as any).window;
  });

  it('resolves system/light/dark and follows OS changes', () => {
    const s = useSheetStore.getState();
    expect(selectResolvedTheme(useSheetStore.getState())).toBe('light');
    s.setSystemPrefersDark(true);
    expect(selectResolvedTheme(useSheetStore.getState())).toBe('dark');
    s.setThemePref('light');
    expect(selectResolvedTheme(useSheetStore.getState())).toBe('light');
    s.setThemePref('dark');
    s.setSystemPrefersDark(false);
    expect(selectResolvedTheme(useSheetStore.getState())).toBe('dark');
  });

  it('saves theme and grid background as app preferences only', () => {
    const s = useSheetStore.getState();
    s.setThemePref('dark');
    s.setGridWhite(true);
    expect(setPreferences).toHaveBeenCalledWith({ theme: 'dark' });
    expect(setPreferences).toHaveBeenCalledWith({ gridWhite: true });
    // ค่าที่โหลดตอนเริ่มไม่ต้องบันทึกซ้ำ
    setPreferences.mockClear();
    s.setThemePref('light', false);
    expect(setPreferences).not.toHaveBeenCalled();
    // ธีมไม่ทำให้งานกลายเป็นยังไม่บันทึก
    expect(useSheetStore.getState().isDirty).toBe(false);
  });

  it('recovery snapshot does not contain the theme', async () => {
    const s = useSheetStore.getState();
    s.setThemePref('dark');
    s.setGridWhite(true);
    s.startEditing('ข้อมูล');
    s.commitEditing();
    await useSheetStore.getState().saveRecoverySnapshot();
    expect(saveRecovery).toHaveBeenCalledTimes(1);
    const json = JSON.stringify(saveRecovery.mock.calls[0][0]);
    expect(json).not.toMatch(/theme|gridWhite|dark/i);
  });

  it('saving .xlsx in dark mode writes no color for unset cells and keeps user colors', async () => {
    const s = useSheetStore.getState();
    s.setThemePref('dark');
    s.setGridWhite(false);
    const wb = useSheetStore.getState().workbook;
    wb.sheets[0].cells['A1'] = { raw: 'ไม่ตั้งสี' };
    wb.sheets[0].cells['A2'] = { raw: 'ตั้งสีเอง', style: { color: '#ff0000', fill: '#ffff00' } };
    wb.sheets[0].cells['A3'] = { raw: 'ตัวหนา', style: { bold: true } };

    const bytes = await exportToXlsxBuffer(wb);
    const excel = new ExcelJS.Workbook();
    await excel.xlsx.load(bytes as any);
    const ws = excel.worksheets[0];

    for (const ref of ['A1', 'A3']) {
      const cell = ws.getCell(ref);
      expect(cell.font?.color, `${ref} font color`).toBeUndefined();
      expect((cell.fill as any)?.fgColor, `${ref} fill`).toBeUndefined();
    }
    expect(ws.getCell('A2').font?.color?.argb).toBe('FFFF0000');
    expect((ws.getCell('A2').fill as any).fgColor.argb).toBe('FFFFFF00');

    // เปิดกลับมา: สีที่ตั้งเองเหมือนเดิม เซลที่ไม่ตั้งสียังไม่มีสี
    const back = await importFromXlsxBuffer(bytes);
    expect(back.sheets[0].cells['A2'].style).toMatchObject({ color: '#ff0000', fill: '#ffff00' });
    expect(back.sheets[0].cells['A1'].style?.color).toBeUndefined();
    expect(back.sheets[0].cells['A1'].style?.fill).toBeUndefined();
    expect(JSON.stringify(back)).not.toMatch(/theme/i);
  });
});
