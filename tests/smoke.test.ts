import { describe, it, expect } from 'vitest';
import { APP_INFO } from '../src/appInfo';
import { th } from '../src/i18n/th';

describe('Phase 0 Smoke Test', () => {
  it('loads app info correctly', () => {
    expect(APP_INFO.name).toBe('SheetEasy');
  });

  it('contains essential Thai translations', () => {
    expect(th.app.title).toBe('SheetEasy');
    expect(th.statusBar.ready).toBe('พร้อมใช้งาน');
    expect(th.errors.div0).toBe('#หารศูนย์');
  });
});
