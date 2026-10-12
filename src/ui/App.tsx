import React, { useState, useEffect, useCallback } from 'react';
import { Toolbar } from './Toolbar';
import { FormulaBar } from './FormulaBar';
import { Grid } from './Grid';
import { StatusBar } from './StatusBar';
import { ConfirmDialog } from './Dialogs';
import { FormulaHelper } from './FormulaHelper';
import { useSheetStore } from './store';
import { th } from '../i18n/th';

export const App: React.FC = () => {
  const {
    isDirty,
    filePath,
    newWorkbook,
    loadWorkbook,
    saveWorkbook,
    markSaved,
    updateTitle,
    pendingDelete,
    confirmPendingDelete,
    cancelPendingDelete,
  } = useSheetStore();

  const [confirmNewOpen, setConfirmNewOpen] = useState(false);
  const [toastMessage, setToastMessage] = useState<string | null>(null);

  const showToast = useCallback((msg: string) => {
    setToastMessage(msg);
    setTimeout(() => {
      setToastMessage(null);
    }, 3500);
  }, []);

  const handleNew = () => {
    if (isDirty) {
      setConfirmNewOpen(true);
    } else {
      newWorkbook();
    }
  };

  const handleOpen = async () => {
    try {
      const electronAPI = (window as any).electronAPI;
      if (electronAPI) {
        const res = await electronAPI.openFile();
        if (res && !res.canceled && res.data) {
          const uint8 = new Uint8Array(res.data);
          await loadWorkbook(uint8, res.filePath);
          showToast(th.messages.fileOpened);
        }
      } else {
        // Fallback for browser testing
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = '.xlsx,.csv';
        input.onchange = async (e: any) => {
          const file = e.target.files?.[0];
          if (file) {
            const buf = await file.arrayBuffer();
            await loadWorkbook(new Uint8Array(buf), file.name);
            showToast(th.messages.fileOpened);
          }
        };
        input.click();
      }
    } catch (err: any) {
      showToast(th.messages.fileOpenFailed + (err.message || String(err)));
    }
  };

  const handleSaveAs = async () => {
    try {
      const electronAPI = (window as any).electronAPI;
      const buffer = await saveWorkbook();

      if (electronAPI) {
        const saveDialogRes = await electronAPI.saveFile(filePath || undefined);
        if (saveDialogRes && !saveDialogRes.canceled && saveDialogRes.filePath) {
          const writeRes = await electronAPI.writeFile(
            saveDialogRes.filePath,
            Array.from(buffer)
          );
          if (writeRes.success) {
            markSaved(saveDialogRes.filePath);
            showToast(th.messages.fileSaved);
          } else {
            showToast(th.messages.fileSaveFailed + writeRes.error);
          }
        }
      } else {
        // Fallback for browser download
        const blob = new Blob([buffer], {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'ตารางรายรับรายจ่าย.xlsx';
        a.click();
        URL.revokeObjectURL(url);
        showToast(th.messages.fileSaved);
      }
    } catch (err: any) {
      showToast(th.messages.fileSaveFailed + (err.message || String(err)));
    }
  };

  const handleSave = async () => {
    if (!filePath) {
      return handleSaveAs();
    }

    try {
      const electronAPI = (window as any).electronAPI;
      const buffer = await saveWorkbook();

      if (electronAPI) {
        const writeRes = await electronAPI.writeFile(filePath, Array.from(buffer));
        if (writeRes.success) {
          markSaved(filePath);
          showToast(th.messages.fileSaved);
        } else {
          showToast(th.messages.fileSaveFailed + writeRes.error);
        }
      } else {
        handleSaveAs();
      }
    } catch (err: any) {
      showToast(th.messages.fileSaveFailed + (err.message || String(err)));
    }
  };

  // Keyboard shortcuts for File operations
  useEffect(() => {
    updateTitle();

    const handleKeyDown = (e: KeyboardEvent) => {
      const isCtrl = e.ctrlKey || e.metaKey;

      if (isCtrl && e.key.toLowerCase() === 's') {
        e.preventDefault();
        if (e.shiftKey) {
          handleSaveAs();
        } else {
          handleSave();
        }
      } else if (isCtrl && e.key.toLowerCase() === 'o') {
        e.preventDefault();
        handleOpen();
      } else if (isCtrl && e.key.toLowerCase() === 'n') {
        e.preventDefault();
        handleNew();
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [filePath, isDirty, handleSave, handleSaveAs, handleOpen, handleNew, updateTitle]);

  return (
    <div
      style={{
        display: 'flex',
        flexDirection: 'column',
        height: '100vh',
        width: '100vw',
        overflow: 'hidden',
        position: 'relative',
      }}
    >
      <Toolbar
        onNew={handleNew}
        onOpen={handleOpen}
        onSave={handleSave}
        onSaveAs={handleSaveAs}
      />
      <FormulaBar />
      <Grid />
      <StatusBar />

      {/* Toast Notification */}
      {toastMessage && (
        <div
          style={{
            position: 'fixed',
            bottom: '36px',
            right: '16px',
            backgroundColor: 'var(--color-primary)',
            color: '#ffffff',
            padding: '8px 16px',
            borderRadius: '6px',
            boxShadow: '0 4px 12px rgba(0,0,0,0.2)',
            zIndex: 1000,
            fontSize: '13px',
            fontWeight: 500,
          }}
        >
          {toastMessage}
        </div>
      )}

      {/* Confirmation Modal */}
      <ConfirmDialog
        isOpen={confirmNewOpen}
        message={th.dialog.confirmNew}
        onConfirm={() => {
          setConfirmNewOpen(false);
          newWorkbook();
        }}
        onCancel={() => setConfirmNewOpen(false)}
      />

      {/* Delete Row/Col Referenced Warning Modal */}
      <ConfirmDialog
        isOpen={pendingDelete !== null}
        message={
          pendingDelete
            ? (pendingDelete.type === 'row'
                ? th.dialog.confirmDeleteReferencedRow
                : th.dialog.confirmDeleteReferencedCol
              ).replace('{count}', String(pendingDelete.count))
            : ''
        }
        confirmText={th.dialog.delete}
        cancelText={th.dialog.cancel}
        danger={true}
        onConfirm={confirmPendingDelete}
        onCancel={cancelPendingDelete}
      />

      {/* Formula Helper Dialog */}
      <FormulaHelper />
    </div>
  );
};

export default App;
