export interface AppPreferences {
  showWelcomeGuide: boolean;
}

export const DEFAULT_PREFERENCES: AppPreferences = {
  showWelcomeGuide: true,
};

export function parsePreferences(jsonStr: string | null | undefined): AppPreferences {
  if (!jsonStr) return { ...DEFAULT_PREFERENCES };
  try {
    const data = JSON.parse(jsonStr);
    return {
      showWelcomeGuide: typeof data.showWelcomeGuide === 'boolean' ? data.showWelcomeGuide : true,
    };
  } catch {
    return { ...DEFAULT_PREFERENCES };
  }
}

export function serializePreferences(prefs: AppPreferences): string {
  return JSON.stringify(prefs, null, 2);
}
