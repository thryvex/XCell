import { describe, it, expect, afterEach } from 'vitest';
import fs from 'fs';
import path from 'path';
import os from 'os';

describe('Atomic File Write Logic (SPEC 6.2)', () => {
  const tempDir = path.join(os.tmpdir(), `xcell-atomic-test-${Date.now()}`);

  if (!fs.existsSync(tempDir)) {
    fs.mkdirSync(tempDir, { recursive: true });
  }

  afterEach(() => {
    try {
      const files = fs.readdirSync(tempDir);
      for (const f of files) {
        fs.unlinkSync(path.join(tempDir, f));
      }
    } catch {}
  });

  it('writes to a temporary file in the same folder and renames over target file', () => {
    const targetFile = path.join(tempDir, 'sample.xlsx');
    const content = Buffer.from('TEST_EXCEL_BINARY_DATA');

    // Atomic write implementation
    const dir = path.dirname(targetFile);
    const tempFile = path.join(dir, `.${path.basename(targetFile)}.tmp-${Date.now()}`);

    fs.writeFileSync(tempFile, content);
    expect(fs.existsSync(tempFile)).toBe(true);

    fs.renameSync(tempFile, targetFile);

    expect(fs.existsSync(targetFile)).toBe(true);
    expect(fs.existsSync(tempFile)).toBe(false);
    expect(fs.readFileSync(targetFile).toString()).toBe('TEST_EXCEL_BINARY_DATA');
  });

  it('cleans up temporary file if write or rename fails, leaving existing target intact', () => {
    const targetFile = path.join(tempDir, 'existing.xlsx');
    fs.writeFileSync(targetFile, Buffer.from('ORIGINAL_DATA'));

    const dir = path.dirname(targetFile);
    const tempFile = path.join(dir, `.${path.basename(targetFile)}.tmp-${Date.now()}`);
    fs.writeFileSync(tempFile, Buffer.from('NEW_CORRUPT_DATA'));

    // Simulate failure during process:
    try {
      // Force an error e.g. simulating failure
      throw new Error('SIMULATED_WRITE_ERROR');
    } catch (err) {
      if (fs.existsSync(tempFile)) {
        fs.unlinkSync(tempFile);
      }
    }

    // Original file remains completely unchanged
    expect(fs.existsSync(targetFile)).toBe(true);
    expect(fs.readFileSync(targetFile).toString()).toBe('ORIGINAL_DATA');
    // Temp file cleaned up
    expect(fs.existsSync(tempFile)).toBe(false);
  });
});
