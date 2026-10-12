/**
 * Manages the recent files list (SPEC 6.2 - max 10 items)
 */
export function updateRecentFilesList(currentList: string[], newPath: string, maxItems = 10): string[] {
  if (!newPath || typeof newPath !== 'string') return currentList;
  const filtered = currentList.filter(
    (p) => p && p.trim().toLowerCase() !== newPath.trim().toLowerCase()
  );
  return [newPath.trim(), ...filtered].slice(0, maxItems);
}

export function removeRecentFileItem(currentList: string[], pathToRemove: string): string[] {
  if (!pathToRemove || typeof pathToRemove !== 'string') return currentList;
  return currentList.filter(
    (p) => p && p.trim().toLowerCase() !== pathToRemove.trim().toLowerCase()
  );
}
