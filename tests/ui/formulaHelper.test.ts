import { describe, it, expect, beforeEach } from 'vitest';
import { CORE_FUNCTIONS } from '../../src/ui/FormulaHelper';
import { useSheetStore } from '../../src/ui/store';

describe('Phase 3B: Formula Helper & Delete Reference Warning', () => {
  beforeEach(() => {
    useSheetStore.getState().newWorkbook();
  });

  describe('CORE_FUNCTIONS catalog', () => {
    it('contains all 21 Core functions from SPEC 4.2', () => {
      const names = CORE_FUNCTIONS.map((f) => f.name);
      const expectedCore = [
        'SUM',
        'AVERAGE',
        'MEDIAN',
        'MIN',
        'MAX',
        'COUNT',
        'COUNTA',
        'PRODUCT',
        'ROUND',
        'ROUNDUP',
        'ROUNDDOWN',
        'ABS',
        'INT',
        'SQRT',
        'POWER',
        'MOD',
        'IF',
        'IFERROR',
        'AND',
        'OR',
        'NOT',
        // Phase 4A functions
        'SUMIF',
        'COUNTIF',
        'AVERAGEIF',
        'SUMIFS',
        'COUNTIFS',
        'TODAY',
        'NOW',
        'YEAR',
        'MONTH',
        'DAY',
        'DATE',
        'TEXT',
        'CONCAT',
        'LEN',
      ];

      for (const expected of expectedCore) {
        expect(names).toContain(expected);
      }
      expect(CORE_FUNCTIONS.length).toBe(35);
    });

    it('each function provides Thai name, syntax, description, and Thai example', () => {
      for (const fn of CORE_FUNCTIONS) {
        expect(fn.name).toBeTruthy();
        expect(fn.thaiName).toBeTruthy();
        expect(fn.syntax).toBeTruthy();
        expect(fn.description).toBeTruthy();
        expect(fn.example).toBeTruthy();
        expect(fn.example.startsWith('=')).toBe(true);
      }
    });
  });

  describe('Formula Helper actions in store', () => {
    it('opens and closes formula helper dialog', () => {
      const store = useSheetStore.getState();
      expect(store.isFormulaHelperOpen).toBe(false);

      store.openFormulaHelper();
      expect(useSheetStore.getState().isFormulaHelperOpen).toBe(true);

      store.closeFormulaHelper();
      expect(useSheetStore.getState().isFormulaHelperOpen).toBe(false);
    });

    it('inserts formula function into active cell ready for point mode', () => {
      const store = useSheetStore.getState();
      store.setActiveCell(1, 1); // B2

      store.insertFormulaFunction('SUM');

      const state = useSheetStore.getState();
      expect(state.isEditing).toBe(true);
      expect(state.editingValue).toBe('=SUM(');
      expect(state.isFormulaHelperOpen).toBe(false);
    });

    it('inserts formula function wrapping existing range when range is selected', () => {
      const store = useSheetStore.getState();
      store.setActiveCell(0, 0);
      store.setSelection(0, 0, 4, 0); // A1:A5

      store.insertFormulaFunction('AVERAGE');

      const state = useSheetStore.getState();
      expect(state.isEditing).toBe(true);
      expect(state.editingValue).toBe('=AVERAGE(A1:A5)');
    });
  });

  describe('Row/Column delete reference warnings', () => {
    it('deletes row immediately without prompt when no formula references it', () => {
      const store = useSheetStore.getState();
      store.setActiveCell(0, 0);
      store.startEditing('100');
      store.commitEditing();

      // Delete row 1 (0-based index 0)
      store.deleteRowAt(0);

      const state = useSheetStore.getState();
      expect(state.pendingDelete).toBeNull();
      // Cell was deleted
      expect(state.workbook.sheets[0].cells['A1']).toBeUndefined();
    });

    it('prompts confirmation when row is referenced by formulas in other cells', () => {
      const store = useSheetStore.getState();
      const sheet = store.workbook.sheets[0];
      sheet.cells['A1'] = { raw: '10' };
      sheet.cells['A2'] = { raw: '=A1*2' };
      sheet.cells['A3'] = { raw: '=A1+10' };

      // Try to delete row 1 (0-based 0)
      store.deleteRowAt(0);

      const state = useSheetStore.getState();
      expect(state.pendingDelete).not.toBeNull();
      expect(state.pendingDelete?.type).toBe('row');
      expect(state.pendingDelete?.index).toBe(0);
      expect(state.pendingDelete?.count).toBe(2); // A2 and A3 reference A1

      // Cancel keeps the row intact
      store.cancelPendingDelete();
      expect(useSheetStore.getState().pendingDelete).toBeNull();
      expect(useSheetStore.getState().workbook.sheets[0].cells['A1']?.raw).toBe('10');

      // Re-trigger and confirm deletion
      store.deleteRowAt(0);
      store.confirmPendingDelete();

      const afterState = useSheetStore.getState();
      expect(afterState.pendingDelete).toBeNull();
      // A1 deleted, formulas shifted with #REF!
      expect(afterState.workbook.sheets[0].cells['A1']?.raw).toContain('#REF!');
    });

    it('prompts confirmation when column is referenced by formulas', () => {
      const store = useSheetStore.getState();
      const sheet = store.workbook.sheets[0];
      sheet.cells['A1'] = { raw: '50' };
      sheet.cells['B1'] = { raw: '=A1+10' };

      // Delete col A (0-based 0)
      store.deleteColAt(0);

      const state = useSheetStore.getState();
      expect(state.pendingDelete).not.toBeNull();
      expect(state.pendingDelete?.type).toBe('col');
      expect(state.pendingDelete?.index).toBe(0);
      expect(state.pendingDelete?.count).toBe(1);

      store.confirmPendingDelete();
      expect(useSheetStore.getState().pendingDelete).toBeNull();
      expect(useSheetStore.getState().workbook.sheets[0].cells['A1']?.raw).toBe('=#REF!+10');
    });
  });
});
