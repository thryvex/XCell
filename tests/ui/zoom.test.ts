import { describe, it, expect, beforeEach } from 'vitest';
import { useSheetStore } from '../../src/ui/store';
import { exportToXlsxBuffer, importFromXlsxBuffer, workbookToExcelJS } from '../../src/core/xlsx';
import { serializeRecoveryData, deserializeRecoveryData } from '../../src/core/recovery';

const S = () => useSheetStore.getState();
const zoomOf = (i = S().activeSheetIndex) => S().workbook.sheets[i].zoom ?? 100;

describe('ซูม (Phase 4C)', () => {
  beforeEach(() => {
    S().newWorkbook();
  });

  it('steps by 10% and clamps to 50%–200%', () => {
    S().zoomIn();
    expect(zoomOf()).toBe(110);
    for (let i = 0; i < 20; i++) S().zoomIn();
    expect(zoomOf()).toBe(200);
    S().resetZoom();
    expect(zoomOf()).toBe(100);
    for (let i = 0; i < 20; i++) S().zoomOut();
    expect(zoomOf()).toBe(50);
    S().setZoom(137);
    expect(zoomOf()).toBe(140);
  });

  it('is kept per sheet', () => {
    S().setZoom(150);
    S().addSheet();
    expect(zoomOf()).toBe(100);
    S().setZoom(70);
    S().switchSheet(0);
    expect(zoomOf()).toBe(150);
    S().switchSheet(1);
    expect(zoomOf()).toBe(70);
  });

  it('is not an undo step and does not mark the file unsaved', () => {
    S().setZoom(120);
    expect(S().history.canUndo()).toBe(false);
    expect(S().isDirty).toBe(false);
  });

  it('round-trips through .xlsx as zoomScale per sheet', async () => {
    S().setZoom(150);
    S().addSheet();
    S().setZoom(80);
    const excel = await workbookToExcelJS(S().workbook);
    expect((excel.worksheets[0].views[0] as any).zoomScale).toBe(150);
    expect((excel.worksheets[1].views[0] as any).zoomScale).toBe(80);

    const reloaded = await importFromXlsxBuffer(await exportToXlsxBuffer(S().workbook));
    expect(reloaded.sheets.map((s) => s.zoom ?? 100)).toEqual([150, 80]);
  });

  it('recovery keeps zoom', () => {
    S().setZoom(130);
    const data = deserializeRecoveryData(
      serializeRecoveryData({ id: 'r', originalPath: null, fileName: 'x', savedAt: 0, workbook: S().workbook })
    )!;
    S().restoreRecoveryWorkbook(data.workbook, null);
    expect(zoomOf()).toBe(130);
  });
});
