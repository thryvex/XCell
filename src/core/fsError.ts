/**
 * Formats file system error codes into friendly Thai messages (SPEC 6.2)
 */
export function formatFsError(err: any): string {
  const code = err?.code;
  if (code === 'EBUSY' || code === 'EPERM') {
    return 'ไฟล์กำลังถูกเปิดใช้งานโดยโปรแกรมอื่น กรุณาปิดไฟล์ในโปรแกรมอื่นก่อนบันทึก';
  }
  if (code === 'EACCES') {
    return 'ไม่มีสิทธิ์เขียนไฟล์ในโฟลเดอร์นี้ กรุณาเลือกโฟลเดอร์อื่น';
  }
  if (code === 'ENOSPC') {
    return 'พื้นที่ดิสก์ไม่เพียงพอ ไม่สามารถบันทึกไฟล์ได้';
  }
  if (code === 'EROFS') {
    return 'ไดรฟ์นี้ถูกป้องกันการเขียน (Read-only)';
  }
  if (code === 'ENOENT') {
    return 'ไม่พบโฟลเดอร์หรือไฟล์ปลายทาง';
  }
  return err?.message ? `เกิดข้อผิดพลาดในการบันทึกไฟล์ (${err.message})` : 'เกิดข้อผิดพลาดในการบันทึกไฟล์';
}
