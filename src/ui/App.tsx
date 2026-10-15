import React, { useState, useEffect, useCallback, useRef } from 'react';
import { Toolbar } from './Toolbar';
import { FormulaBar } from './FormulaBar';
import { Grid } from './Grid';
import { StatusBar } from './StatusBar';
import { SheetTabs } from './SheetTabs';
import { SearchBar } from './SearchBar';
import { SortDialog } from './SortDialog';
import {
  ConfirmDialog,
  UnsavedChangesDialog,
  RecoveryDialog,
  UnsupportedOverwriteDialog,
} from './Dialogs';
import { TemplateModal } from './TemplateModal';
import { WelcomeGuideModal } from './WelcomeGuideModal';
import { HelpModal } from './HelpModal';
import { FormulaHelper } from './FormulaHelper';
import { useSheetStore } from './store';
import { th } from '../i18n/th';
import { createLedgerWorkbook, createInventoryWorkbook } from '../templates';

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
    // Phase 3C & 3D
    recentFiles,
    fetchRecentFiles,
    removeRecentFile,
    saveRecoverySnapshot,
    clearRecoverySnapshot,
    restoreRecoveryWorkbook,
    unsupportedFeatures,
    hiddenSheetNotice,
    filteredRowsNotice,
    showAllFilteredRows,
    dismissFilteredRowsNotice,
    hasWarnedUnsupportedOverwrite,
    setHasWarnedUnsupportedOverwrite,
    dismissUnsupportedBanner,
    loadTemplate,
  } = useSheetStore();

  const [toastMessage, setToastMessage] = useState<string | null>(null);
  const [isDraggingOver, setIsDraggingOver] = useState(false);

  // Dialog states
  const [unsavedDialogOpen, setUnsavedDialogOpen] = useState(false);
  const [pendingAction, setPendingAction] = useState<(() => Promise<void> | void) | 'close' | null>(null);

  const [recoveryDialogOpen, setRecoveryDialogOpen] = useState(false);
  const [recoveryList, setRecoveryList] = useState<any[]>([]);

  const [unsupportedOverwriteOpen, setUnsupportedOverwriteOpen] = useState(false);

  // Phase 3D Modal states
  const [templateModalOpen, setTemplateModalOpen] = useState(false);
  const [welcomeGuideModalOpen, setWelcomeGuideModalOpen] = useState(false);
  const [helpModalOpen, setHelpModalOpen] = useState(false);
  const openedGuideFromHelp = useRef(false);

  const showToast = useCallback((msg: string) => {
    setToastMessage(msg);
    setTimeout(() => {
      setToastMessage(null);
    }, 3500);
  }, []);

  // Check whether to show Welcome Guide or Template Modal on startup
  const checkWelcomeOrTemplate = useCallback(async () => {
    const electronAPI = typeof window !== 'undefined' ? (window as any).electronAPI : undefined;
    let showWelcome = true;
    if (electronAPI?.getPreferences) {
      try {
        const prefs = await electronAPI.getPreferences();
        if (prefs && typeof prefs.showWelcomeGuide === 'boolean') {
          showWelcome = prefs.showWelcomeGuide;
        }
      } catch {}
    } else {
      try {
        const stored = localStorage.getItem('sheeteasy_show_welcome');
        if (stored !== null) {
          showWelcome = stored === 'true';
        }
      } catch {}
    }

    if (showWelcome) {
      openedGuideFromHelp.current = false;
      setWelcomeGuideModalOpen(true);
    } else {
      setTemplateModalOpen(true);
    }
  }, []);

  const savePreferences = async (newPrefs: { showWelcomeGuide: boolean }) => {
    const electronAPI = typeof window !== 'undefined' ? (window as any).electronAPI : undefined;
    if (electronAPI?.setPreferences) {
      try {
        await electronAPI.setPreferences(newPrefs);
      } catch {}
    } else {
      try {
        localStorage.setItem('sheeteasy_show_welcome', String(newPrefs.showWelcomeGuide));
      } catch {}
    }
  };

  const handleWelcomeFinish = async (_dontShowAgain: boolean) => {
    setWelcomeGuideModalOpen(false);
    await savePreferences({ showWelcomeGuide: false });
    if (!openedGuideFromHelp.current) {
      setTemplateModalOpen(true);
    }
    openedGuideFromHelp.current = false;
  };

  const handleWelcomeSkip = async (dontShowAgain: boolean) => {
    setWelcomeGuideModalOpen(false);
    if (dontShowAgain) {
      await savePreferences({ showWelcomeGuide: false });
    }
    if (!openedGuideFromHelp.current) {
      setTemplateModalOpen(true);
    }
    openedGuideFromHelp.current = false;
  };

  // Fetch recent files and check for recovery files on mount
  useEffect(() => {
    fetchRecentFiles();

    const checkStartup = async () => {
      const electronAPI = typeof window !== 'undefined' ? (window as any).electronAPI : undefined;
      let hasRecovery = false;
      if (electronAPI?.getRecoveryList) {
        try {
          const list = await electronAPI.getRecoveryList();
          if (Array.isArray(list) && list.length > 0) {
            setRecoveryList(list);
            setRecoveryDialogOpen(true);
            hasRecovery = true;
          }
        } catch {}
      }

      if (!hasRecovery) {
        await checkWelcomeOrTemplate();
      }
    };

    checkStartup();
  }, [fetchRecentFiles, checkWelcomeOrTemplate]);

  // Periodic auto-recovery snapshot every 30 seconds (SPEC 6.3)
  useEffect(() => {
    if (!isDirty) return;

    const timer = setInterval(() => {
      saveRecoverySnapshot();
    }, 30000);

    return () => clearInterval(timer);
  }, [isDirty, saveRecoverySnapshot]);

  // Execute save to a given file path
  const executeSaveToPath = async (targetPath: string): Promise<boolean> => {
    try {
      const electronAPI = (window as any).electronAPI;
      const buffer = await saveWorkbook();

      if (electronAPI) {
        const writeRes = await electronAPI.writeFile(targetPath, Array.from(buffer));
        if (writeRes.success) {
          markSaved(targetPath);
          showToast(th.messages.fileSaved);
          return true;
        } else {
          showToast(th.messages.fileSaveFailed + writeRes.error);
          return false;
        }
      } else {
        markSaved(targetPath);
        showToast(th.messages.fileSaved);
        return true;
      }
    } catch (err: any) {
      showToast(th.messages.fileSaveFailed + (err.message || String(err)));
      return false;
    }
  };

  // Save As
  const handleSaveAs = async (): Promise<boolean> => {
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
            return true;
          } else {
            showToast(th.messages.fileSaveFailed + writeRes.error);
            return false;
          }
        }
        return false;
      } else {
        // Fallback for browser download
        const blob = new Blob([new Uint8Array(buffer)], {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'ตารางรายรับรายจ่าย.xlsx';
        a.click();
        URL.revokeObjectURL(url);
        markSaved('ตารางรายรับรายจ่าย.xlsx');
        showToast(th.messages.fileSaved);
        return true;
      }
    } catch (err: any) {
      showToast(th.messages.fileSaveFailed + (err.message || String(err)));
      return false;
    }
  };

  // Save (handles unsupported overwrite warning on first Ctrl+S)
  const handleSave = async (): Promise<boolean> => {
    if (!filePath) {
      return handleSaveAs();
    }

    if (unsupportedFeatures.length > 0 && !hasWarnedUnsupportedOverwrite) {
      setUnsupportedOverwriteOpen(true);
      return false;
    }

    return executeSaveToPath(filePath);
  };

  // Unsaved Action Guard: if dirty, ask user before continuing (บันทึก / ไม่บันทึก / ยกเลิก)
  const requestActionWithGuard = (action: (() => Promise<void> | void) | 'close') => {
    if (isDirty) {
      setPendingAction(() => (typeof action === 'function' ? action : action));
      setUnsavedDialogOpen(true);
    } else {
      if (action === 'close') {
        const electronAPI = (window as any).electronAPI;
        if (electronAPI?.forceClose) {
          electronAPI.forceClose();
        }
      } else {
        action();
      }
    }
  };

  // Handler for Unsaved Changes Dialog "Save" button
  const handleUnsavedDialogSave = async () => {
    setUnsavedDialogOpen(false);
    const saveSuccess = await handleSave();
    if (saveSuccess && pendingAction) {
      if (pendingAction === 'close') {
        const electronAPI = (window as any).electronAPI;
        if (electronAPI?.forceClose) electronAPI.forceClose();
      } else if (typeof pendingAction === 'function') {
        await pendingAction();
      }
    }
    setPendingAction(null);
  };

  // Handler for Unsaved Changes Dialog "Don't Save" button
  const handleUnsavedDialogDontSave = async () => {
    setUnsavedDialogOpen(false);
    await clearRecoverySnapshot();
    if (pendingAction) {
      if (pendingAction === 'close') {
        const electronAPI = (window as any).electronAPI;
        if (electronAPI?.forceClose) electronAPI.forceClose();
      } else if (typeof pendingAction === 'function') {
        await pendingAction();
      }
    }
    setPendingAction(null);
  };

  // Handler for Unsaved Changes Dialog "Cancel" button
  const handleUnsavedDialogCancel = () => {
    setUnsavedDialogOpen(false);
    setPendingAction(null);
  };

  // New Workbook (Phase 3D: opens Template Selection modal)
  const handleNew = () => {
    requestActionWithGuard(() => {
      setTemplateModalOpen(true);
    });
  };

  const handleSelectBlank = () => {
    newWorkbook();
    setTemplateModalOpen(false);
  };

  const handleSelectLedger = () => {
    const wb = createLedgerWorkbook();
    loadTemplate(wb, 'ตารางรายรับรายจ่าย');
    setTemplateModalOpen(false);
  };

  const handleSelectInventory = () => {
    const wb = createInventoryWorkbook();
    loadTemplate(wb, 'ตารางรายการสินค้า');
    setTemplateModalOpen(false);
  };

  const handleSelectRecent = (targetPath: string) => {
    setTemplateModalOpen(false);
    handleOpenRecent(targetPath);
  };

  const handleCloseTemplateModal = () => {
    newWorkbook();
    setTemplateModalOpen(false);
  };

  // Open File
  const handleOpen = () => {
    requestActionWithGuard(async () => {
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
    });
  };

  // Open Recent File
  const handleOpenRecent = (targetPath: string) => {
    requestActionWithGuard(async () => {
      try {
        const electronAPI = (window as any).electronAPI;
        if (electronAPI) {
          const exists = await electronAPI.checkFileExists(targetPath);
          if (!exists) {
            showToast(th.messages.recentFileNotFound.replace('{path}', targetPath));
            await removeRecentFile(targetPath);
            return;
          }
          const res = await electronAPI.readFile(targetPath);
          if (res.success && res.data) {
            await loadWorkbook(new Uint8Array(res.data), targetPath);
            showToast(th.messages.fileOpened);
          } else {
            showToast(th.messages.fileOpenFailed + (res.error || ''));
          }
        }
      } catch (err: any) {
        showToast(th.messages.fileOpenFailed + (err.message || String(err)));
      }
    });
  };

  // Open file from external source (Open with / drag-drop / second instance)
  const handleOpenExternal = useCallback(
    (targetPath: string) => {
      requestActionWithGuard(async () => {
        try {
          const electronAPI = (window as any).electronAPI;
          if (electronAPI?.readFile) {
            const res = await electronAPI.readFile(targetPath);
            if (res.success && res.data) {
              await loadWorkbook(new Uint8Array(res.data), targetPath);
              showToast(th.messages.fileOpened);
            } else {
              showToast(th.messages.fileOpenFailed + (res.error || ''));
            }
          }
        } catch (err: any) {
          showToast(th.messages.fileOpenFailed + (err.message || String(err)));
        }
      });
    },
    [isDirty, loadWorkbook, showToast]
  );

  // Drag and Drop Handlers
  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDraggingOver(true);
  };

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDraggingOver(false);
  };

  const handleDrop = async (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDraggingOver(false);

    const file = e.dataTransfer.files?.[0];
    if (!file) return;

    const lowerName = file.name.toLowerCase();
    if (!lowerName.endsWith('.xlsx') && !lowerName.endsWith('.csv')) {
      showToast(th.errors.dragDropUnsupported);
      return;
    }

    const filePath = (file as any).path || file.name;

    requestActionWithGuard(async () => {
      try {
        const buf = await file.arrayBuffer();
        await loadWorkbook(new Uint8Array(buf), filePath);
        showToast(th.messages.fileOpened);
      } catch (err: any) {
        showToast(th.messages.fileOpenFailed + (err.message || String(err)));
      }
    });
  };

  // Listen for Electron events (second instance Open with & window close interception)
  useEffect(() => {
    const electronAPI = (window as any).electronAPI;
    if (!electronAPI) return;

    const unsubscribeOpenExternal = electronAPI.onOpenExternalFile((filePath: string) => {
      handleOpenExternal(filePath);
    });

    const unsubscribeRequestClose = electronAPI.onRequestClose(() => {
      requestActionWithGuard('close');
    });

    return () => {
      if (unsubscribeOpenExternal) unsubscribeOpenExternal();
      if (unsubscribeRequestClose) unsubscribeRequestClose();
    };
  }, [handleOpenExternal, isDirty]);

  // Keyboard shortcuts
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
      onDragOver={handleDragOver}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}
      style={{
        display: 'flex',
        flexDirection: 'column',
        height: '100vh',
        width: '100vw',
        overflow: 'hidden',
        position: 'relative',
      }}
    >
      {/* Drag & Drop Visual Overlay */}
      {isDraggingOver && (
        <div
          style={{
            position: 'absolute',
            top: 0,
            left: 0,
            right: 0,
            bottom: 0,
            backgroundColor: 'rgba(59, 130, 246, 0.15)',
            border: '3px dashed var(--color-primary)',
            zIndex: 9999,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            fontSize: '18px',
            fontWeight: 600,
            color: 'var(--color-primary)',
            pointerEvents: 'none',
          }}
        >
          📂 วางไฟล์ .xlsx หรือ .csv เพื่อเปิดใน SheetEasy
        </div>
      )}

      <Toolbar
        onNew={handleNew}
        onOpen={handleOpen}
        onOpenRecent={handleOpenRecent}
        onSave={handleSave}
        onSaveAs={handleSaveAs}
        onOpenHelp={() => setHelpModalOpen(true)}
      />

      <FormulaBar />

      {/* Unsupported Features Warning Banner (SPEC 6.4) */}
      {(unsupportedFeatures.length > 0 || hiddenSheetNotice > 0) && (
        <div
          style={{
            backgroundColor: '#fef3c7',
            color: '#92400e',
            borderBottom: '1px solid #fcd34d',
            padding: '7px 16px',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            fontSize: '13px',
            zIndex: 5,
          }}
        >
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            <span>⚠️</span>
            <span>
              {unsupportedFeatures.length > 0 &&
                th.banner.unsupportedFeatures.replace('{features}', unsupportedFeatures.join(', '))}
              {unsupportedFeatures.length > 0 && hiddenSheetNotice > 0 && ' · '}
              {hiddenSheetNotice > 0 &&
                th.banner.hiddenSheets.replace('{count}', String(hiddenSheetNotice))}
            </span>
          </div>
          <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
            {unsupportedFeatures.length > 0 && (
            <button
              style={{
                padding: '4px 10px',
                fontSize: '12px',
                fontWeight: 600,
                backgroundColor: '#ffffff',
                border: '1px solid #d97706',
                color: '#92400e',
                borderRadius: '4px',
                cursor: 'pointer',
              }}
              onClick={handleSaveAs}
            >
              💾 {th.banner.saveAsNewBtn}
            </button>
            )}
            <button
              style={{
                padding: '2px 8px',
                fontSize: '13px',
                background: 'transparent',
                border: 'none',
                color: '#92400e',
                cursor: 'pointer',
              }}
              onClick={dismissUnsupportedBanner}
              title={th.banner.dismiss}
            >
              ✕
            </button>
          </div>
        </div>
      )}

      {/* ไฟล์ที่เปิดมีแถวที่ซ่อนด้วยตัวกรอง (เงื่อนไขกรองไม่ถูกเก็บใน .xlsx) */}
      {filteredRowsNotice && (
        <div
          role="status"
          style={{
            backgroundColor: '#e0f2fe',
            color: '#075985',
            borderBottom: '1px solid #7dd3fc',
            padding: '7px 16px',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            fontSize: '13px',
          }}
        >
          <span>ℹ️ {th.filter.fileHasFilteredRows}</span>
          <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
            <button onClick={showAllFilteredRows} style={{ fontSize: '12px', fontWeight: 600 }}>
              {th.filter.showAllRows}
            </button>
            <button
              onClick={dismissFilteredRowsNotice}
              title={th.banner.dismiss}
              style={{ background: 'transparent', border: 'none', cursor: 'pointer' }}
            >
              ✕
            </button>
          </div>
        </div>
      )}

      {/* พื้นที่ตาราง + แถบค้นหามุมขวาบน */}
      <div style={{ position: 'relative', flex: 1, display: 'flex', flexDirection: 'column', minHeight: 0 }}>
        <Grid />
        <SearchBar />
      </div>
      <SheetTabs />
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
            zIndex: 2000,
            fontSize: '13px',
            fontWeight: 500,
          }}
        >
          {toastMessage}
        </div>
      )}

      {/* Unsaved Changes Dialog (SPEC 6.2 - บันทึก / ไม่บันทึก / ยกเลิก) */}
      <UnsavedChangesDialog
        isOpen={unsavedDialogOpen}
        onSave={handleUnsavedDialogSave}
        onDontSave={handleUnsavedDialogDontSave}
        onCancel={handleUnsavedDialogCancel}
      />

      {/* Auto-Recovery Prompt Dialog (SPEC 6.3) */}
      <RecoveryDialog
        isOpen={recoveryDialogOpen}
        recoveryItems={recoveryList}
        onRecover={(item) => {
          restoreRecoveryWorkbook(item.workbook, item.originalPath);
          setRecoveryDialogOpen(false);
          showToast(th.messages.recoveryRestored);
        }}
        onDiscard={async () => {
          const electronAPI = typeof window !== 'undefined' ? (window as any).electronAPI : undefined;
          if (electronAPI?.clearAllRecovery) {
            await electronAPI.clearAllRecovery();
          }
          setRecoveryDialogOpen(false);
          showToast(th.messages.recoveryDiscarded);
          await checkWelcomeOrTemplate();
        }}
      />

      {/* Unsupported Overwrite Warning Dialog (SPEC 6.4) */}
      <UnsupportedOverwriteDialog
        isOpen={unsupportedOverwriteOpen}
        features={unsupportedFeatures}
        onOverwrite={async () => {
          setHasWarnedUnsupportedOverwrite(true);
          setUnsupportedOverwriteOpen(false);
          if (filePath) {
            await executeSaveToPath(filePath);
          }
        }}
        onSaveAs={async () => {
          setUnsupportedOverwriteOpen(false);
          await handleSaveAs();
        }}
        onCancel={() => {
          setUnsupportedOverwriteOpen(false);
        }}
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

      {/* เรียงลำดับ (4C) */}
      <SortDialog />

      {/* Formula Helper Dialog */}
      <FormulaHelper />

      {/* Template Selection Modal (Phase 3D) */}
      <TemplateModal
        isOpen={templateModalOpen}
        onSelectBlank={handleSelectBlank}
        onSelectLedger={handleSelectLedger}
        onSelectInventory={handleSelectInventory}
        onSelectRecent={handleSelectRecent}
        onClose={handleCloseTemplateModal}
        recentFiles={recentFiles}
      />

      {/* Welcome Guide Modal (Phase 3D) */}
      <WelcomeGuideModal
        isOpen={welcomeGuideModalOpen}
        onFinish={handleWelcomeFinish}
        onSkip={handleWelcomeSkip}
      />

      {/* Help Modal (Phase 3D) */}
      <HelpModal
        isOpen={helpModalOpen}
        onClose={() => setHelpModalOpen(false)}
        onOpenWelcomeGuide={() => {
          openedGuideFromHelp.current = true;
          setWelcomeGuideModalOpen(true);
        }}
      />
    </div>
  );
};

export default App;
