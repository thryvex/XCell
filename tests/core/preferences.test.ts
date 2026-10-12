import { describe, it, expect } from 'vitest';
import {
  DEFAULT_PREFERENCES,
  parsePreferences,
  serializePreferences,
  AppPreferences,
} from '../../src/core/preferences';

describe('Preferences (Phase 3D)', () => {
  it('has showWelcomeGuide default to true', () => {
    expect(DEFAULT_PREFERENCES.showWelcomeGuide).toBe(true);
  });

  it('parses valid preferences JSON', () => {
    const raw = JSON.stringify({ showWelcomeGuide: false });
    const parsed = parsePreferences(raw);
    expect(parsed.showWelcomeGuide).toBe(false);
  });

  it('handles invalid or empty JSON safely by returning default preferences', () => {
    expect(parsePreferences(null)).toEqual(DEFAULT_PREFERENCES);
    expect(parsePreferences('')).toEqual(DEFAULT_PREFERENCES);
    expect(parsePreferences('   ')).toEqual(DEFAULT_PREFERENCES);
    expect(parsePreferences('invalid json')).toEqual(DEFAULT_PREFERENCES);
    expect(parsePreferences('{ invalid: }')).toEqual(DEFAULT_PREFERENCES);
  });

  it('handles partial JSON safely by merging with defaults', () => {
    const parsed = parsePreferences(JSON.stringify({ someOtherField: 123 }));
    expect(parsed.showWelcomeGuide).toBe(true);
  });

  it('serializes preferences to formatted JSON string', () => {
    const prefs: AppPreferences = { showWelcomeGuide: false };
    const serialized = serializePreferences(prefs);
    expect(serialized).toContain('"showWelcomeGuide": false');
    expect(JSON.parse(serialized)).toEqual({ showWelcomeGuide: false });
  });
});
