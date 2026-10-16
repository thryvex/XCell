import { describe, it, expect, beforeEach } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';
import { useSheetStore } from '../../src/ui/store';
import { exportToXlsxBuffer } from '../../src/core/xlsx';
import { createEmptyWorkbook } from '../../src/core/model';
import { serializeRecoveryData, deserializeRecoveryData } from '../../src/core/recovery';

const S = () => useSheetStore.getState();
const chartFile = () => new Uint8Array(fs.readFileSync(path.join(__dirname, '../fixtures', 'chart-bar.xlsx')));
const ORIGINAL = 'C:\\งาน\\กราฟ.xlsx';

describe('คำเตือนไฟล์ที่มีกราฟ (store)', () => {
  beforeEach(async () => {
    S().newWorkbook();
    await S().loadWorkbook(chartFile(), ORIGINAL);
  });

  it('เปิดไฟล์ที่มีกราฟแล้วแถบเตือนกราฟแสดง', () => {
    expect(S().workbook.sourceHasCharts).toBe(true);
    expect(S().chartBannerDismissed).toBe(false);
    expect(S().filePath).toBe(ORIGINAL);
  });

  it('ปิดแถบเตือน (ทั้งแถบกราฟและแถบทั่วไป) แล้วยังต้องถามก่อนบันทึกทับ', () => {
    S().dismissChartBanner();
    S().dismissUnsupportedBanner();
    expect(S().chartBannerDismissed).toBe(true);
    expect(S().workbook.sourceHasCharts).toBe(true);
  });

  it('บันทึกสำเร็จแล้ว (ทับหรือเป็นไฟล์ใหม่) ไม่ต้องถามอีก', () => {
    S().markSaved('C:\\งาน\\กราฟ-ใหม่.xlsx');
    expect(S().workbook.sourceHasCharts).toBeUndefined();
  });

  it('เปิดไฟล์อื่นที่ไม่มีกราฟแล้วล้างสถานะ และแถบเตือนกลับมาเมื่อเปิดไฟล์กราฟอีกครั้ง', async () => {
    S().dismissChartBanner();
    const plain = createEmptyWorkbook();
    plain.sheets[0].cells['A1'] = { raw: '1' };
    await S().loadWorkbook(await exportToXlsxBuffer(plain), 'C:\\งาน\\ธรรมดา.xlsx');
    expect(S().workbook.sourceHasCharts).toBeUndefined();

    await S().loadWorkbook(chartFile(), ORIGINAL);
    expect(S().workbook.sourceHasCharts).toBe(true);
    expect(S().chartBannerDismissed).toBe(false);
  });

  it('สร้างงานใหม่ไม่มีสถานะกราฟ', () => {
    S().newWorkbook();
    expect(S().workbook.sourceHasCharts).toBeUndefined();
  });

  it('ไฟล์กู้คืนเก็บสถานะกราฟไว้ กู้คืนแล้วยังต้องถามก่อนบันทึกทับไฟล์ต้นฉบับ', () => {
    const json = serializeRecoveryData({
      id: 'r1',
      originalPath: ORIGINAL,
      fileName: 'กราฟ',
      savedAt: 0,
      workbook: S().workbook,
    });
    const item = deserializeRecoveryData(json)!;
    S().newWorkbook();
    S().restoreRecoveryWorkbook(item.workbook, item.originalPath);
    expect(S().workbook.sourceHasCharts).toBe(true);
    expect(S().filePath).toBe(ORIGINAL);
  });
});
