import { describe, it, expect } from 'vitest';
import { formatFsError } from '../../src/core/fsError';

describe('src/core/fsError (SPEC 6.2)', () => {
  it('formats EBUSY and EPERM errors to Thai message explaining file is locked', () => {
    const errBusy = { code: 'EBUSY', message: 'resource busy or locked' };
    const errPerm = { code: 'EPERM', message: 'operation not permitted' };

    expect(formatFsError(errBusy)).toBe(
      'ไฟล์กำลังถูกเปิดใช้งานโดยโปรแกรมอื่น กรุณาปิดไฟล์ในโปรแกรมอื่นก่อนบันทึก'
    );
    expect(formatFsError(errPerm)).toBe(
      'ไฟล์กำลังถูกเปิดใช้งานโดยโปรแกรมอื่น กรุณาปิดไฟล์ในโปรแกรมอื่นก่อนบันทึก'
    );
  });

  it('formats EACCES error to Thai message explaining permission denied', () => {
    const errAccess = { code: 'EACCES', message: 'permission denied' };
    expect(formatFsError(errAccess)).toBe(
      'ไม่มีสิทธิ์เขียนไฟล์ในโฟลเดอร์นี้ กรุณาเลือกโฟลเดอร์อื่น'
    );
  });

  it('formats ENOSPC error to Thai message explaining disk is full', () => {
    const errNoSpace = { code: 'ENOSPC', message: 'no space left on device' };
    expect(formatFsError(errNoSpace)).toBe(
      'พื้นที่ดิสก์ไม่เพียงพอ ไม่สามารถบันทึกไฟล์ได้'
    );
  });

  it('formats EROFS error to Thai message explaining read-only drive', () => {
    const errReadOnly = { code: 'EROFS', message: 'read-only file system' };
    expect(formatFsError(errReadOnly)).toBe(
      'ไดรฟ์นี้ถูกป้องกันการเขียน (Read-only)'
    );
  });

  it('formats ENOENT error to Thai message explaining directory not found', () => {
    const errNotFound = { code: 'ENOENT', message: 'no such file or directory' };
    expect(formatFsError(errNotFound)).toBe(
      'ไม่พบโฟลเดอร์หรือไฟล์ปลายทาง'
    );
  });

  it('falls back to generic Thai error message for unmapped errors', () => {
    const errUnknown = { message: 'Network path failed' };
    expect(formatFsError(errUnknown)).toContain('เกิดข้อผิดพลาดในการบันทึกไฟล์');
    expect(formatFsError(errUnknown)).toContain('Network path failed');
  });
});
