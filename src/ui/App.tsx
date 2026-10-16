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
import { useSheetStore, isCsvPath, selectResolvedTheme } from './store';
import { needsLowContrastWarning, WINDOW_BACKGROUND, ThemePref } from '../core/theme';
import { CsvImportDialog } from './CsvImportDialog';
import { PrintDialog } from './PrintDialog';
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
    chartBannerDismissed,
    dismissChartBanner,
    loadTemplate,
    fileName,
    beginCsvImport,
    getCsvExport,
  } = useSheetStore();

  const [toastMessage, setToastMessage] = useState<string | null>(null);
  const [isDraggingOver, setIsDraggingOver] = useState(false);

  // Dialog states
  const [unsavedDialogOpen, setUnsavedDialogOpen] = useState(false);
  const [pendingAction, setPendingAction] = useState<(() => Promise<void> | void) | 'close' | null>(null);

  const [recoveryDialogOpen, setRecoveryDialogOpen] = useState(false);
  const [recoveryList, setRecoveryList] = useState<any[]>([]);

  const [unsupportedOverwriteOpen, setUnsupportedOverwriteOpen] = useState(false);
  const [chartOverwriteOpen, setChartOverwriteOpen] = useState(false);
  // ไฟล์ต้นฉบับมีกราฟ (ล้างเมื่อบันทึกสำเร็จ) — ไม่ขึ้นกับการปิดแถบเตือน
  const sourceHasCharts = useSheetStore((s) => Boolean(s.workbook.sourceHasCharts));

  // Phase 3D Modal states
  const [templateModalOpen, setTemplateModalOpen] = useState(false);
  const [welcomeGuideModalOpen, setWelcomeGuideModalOpen] = useState(false);
  const [helpModalOpen, setHelpModalOpen] = useState(false);
  const openedGuideFromHelp = useRef(false);
  const openExternalRef = useRef<((path: string) => void) | null>(null);

  const showToast = useCallback((msg: string) => {
    setToastMessage(msg);
    setTimeout(() => {
      setToastMessage(null);
    }, 3500);
  }, []);

  // --- ปี พ.ศ. (5A) ---
  // โหลดรูปแบบปีจากค่าของแอป (userData) ตอนเริ่ม
  const hideBeNoticeRef = useRef(false);
  const beNoticeShownRef = useRef(false);
  const [beNoticeOpen, setBeNoticeOpen] = useState(false);
  const [beNoticeDontShow, setBeNoticeDontShow] = useState(false);
  useEffect(() => {
    (async () => {
      const electronAPI = typeof window !== 'undefined' ? (window as any).electronAPI : undefined;
      let mode: string | undefined;
      try {
        if (electronAPI?.getPreferences) {
          const prefs = await electronAPI.getPreferences();
          mode = prefs?.yearMode;
          hideBeNoticeRef.current = prefs?.hideBuddhistSaveNotice === true;
        } else {
          mode = localStorage.getItem('sheeteasy_year_mode') ?? undefined;
          hideBeNoticeRef.current = localStorage.getItem('sheeteasy_hide_be_notice') === 'true';
        }
      } catch {}
      if (mode === 'be') useSheetStore.getState().setYearMode('be', false);
    })();
  }, []);

  // --- ธีมสี (5C) ---
  const resolvedTheme = useSheetStore(selectResolvedTheme);
  const gridWhite = useSheetStore((s) => s.gridWhite);
  useEffect(() => {
    // ค่าที่บันทึกไว้ (userData) — main ตั้ง nativeTheme ตามค่านี้แล้วตั้งแต่สร้างหน้าต่าง
    (async () => {
      const electronAPI = typeof window !== 'undefined' ? (window as any).electronAPI : undefined;
      let theme: string | null | undefined;
      let white = false;
      try {
        if (electronAPI?.getPreferences) {
          const prefs = await electronAPI.getPreferences();
          theme = prefs?.theme;
          white = prefs?.gridWhite === true;
        } else {
          theme = localStorage.getItem('sheeteasy_theme');
          white = localStorage.getItem('sheeteasy_grid_white') === '1';
        }
      } catch {}
      const store = useSheetStore.getState();
      if (theme === 'light' || theme === 'dark') store.setThemePref(theme as ThemePref, false);
      if (white) store.setGridWhite(true, false);
    })();
  }, []);
  useEffect(() => {
    // โหมดของระบบ: ตามการเปลี่ยนทันที
    const electronAPI = typeof window !== 'undefined' ? (window as any).electronAPI : undefined;
    if (electronAPI?.onSystemThemeChanged) {
      // Electron: main ส่งค่า nativeTheme.shouldUseDarkColors มาเมื่อเปลี่ยน
      useSheetStore.getState().setSystemPrefersDark(electronAPI.initialTheme === 'dark');
      return electronAPI.onSystemThemeChanged((dark: boolean) => useSheetStore.getState().setSystemPrefersDark(dark));
    }
    if (typeof matchMedia !== 'function') return;
    const mq = matchMedia('(prefers-color-scheme: dark)');
    useSheetStore.getState().setSystemPrefersDark(mq.matches);
    const onChange = (e: MediaQueryListEvent) => useSheetStore.getState().setSystemPrefersDark(e.matches);
    mq.addEventListener('change', onChange);
    return () => mq.removeEventListener('change', onChange);
  }, []);
  useEffect(() => {
    document.documentElement.dataset.theme = resolvedTheme;
  }, [resolvedTheme]);

  // เปิดไฟล์ในโหมดมืดแล้วมีสีตัวอักษรที่อาจมองยาก: เตือนครั้งเดียวต่อไฟล์ต่อรอบการใช้งาน
  const [lowContrastNotice, setLowContrastNotice] = useState(false);
  const lowContrastShownRef = useRef<Set<string>>(new Set());
  const checkLowContrast = (path: string) => {
    if (lowContrastShownRef.current.has(path)) return;
    const s = useSheetStore.getState();
    const gridBg =
      getComputedStyle(document.documentElement).getPropertyValue('--grid-bg').trim() || WINDOW_BACKGROUND.dark;
    if (needsLowContrastWarning(s.workbook, selectResolvedTheme(s), s.gridWhite, gridBg)) {
      lowContrastShownRef.current.add(path);
      setLowContrastNotice(true);
    }
  };

  // บันทึกสำเร็จ: ถ้าแสดงปีเป็น พ.ศ. แจ้งครั้งเดียวว่าไฟล์เก็บปีเป็น ค.ศ.
  const markSavedAndNotify = (path: string) => {
    markSaved(path);
    if (
      useSheetStore.getState().yearMode === 'be' &&
      !hideBeNoticeRef.current &&
      !beNoticeShownRef.current
    ) {
      beNoticeShownRef.current = true;
      setBeNoticeDontShow(false);
      setBeNoticeOpen(true);
    }
  };
  const closeBeNotice = () => {
    setBeNoticeOpen(false);
    if (!beNoticeDontShow) return;
    hideBeNoticeRef.current = true;
    const electronAPI = (window as any).electronAPI;
    if (electronAPI?.setPreferences) {
      electronAPI.setPreferences({ hideBuddhistSaveNotice: true }).catch(() => {});
    } else {
      try {
        localStorage.setItem('sheeteasy_hide_be_notice', 'true');
      } catch {}
    }
  };

  // เปิดไฟล์จากไบต์: .csv ไปหน้าต่างนำเข้า CSV (งานใหม่), นอกนั้นเปิดเป็น .xlsx
  const openFileBytes = useCallback(
    async (data: Uint8Array, path: string) => {
      if (isCsvPath(path)) {
        beginCsvImport(data, path, 'open');
        return;
      }
      setLowContrastNotice(false);
      await loadWorkbook(data, path);
      showToast(th.messages.fileOpened);
      if (useSheetStore.getState().filePath === path) checkLowContrast(path);
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [beginCsvImport, loadWorkbook, showToast]
  );

  // ส่งออก CSV (ชีตที่เปิดอยู่, ค่าที่แสดง)
  const [csvFilterConfirmOpen, setCsvFilterConfirmOpen] = useState(false);
  const [printDialog, setPrintDialog] = useState<'print' | 'pdf' | null>(null);
  const runCsvExport = async () => {
    try {
      const { bytes, defaultName } = getCsvExport();
      const electronAPI = (window as any).electronAPI;
      if (electronAPI?.exportCsvDialog) {
        const res = await electronAPI.exportCsvDialog(defaultName);
        if (!res || res.canceled || !res.filePath) return;
        const writeRes = await electronAPI.writeFile(res.filePath, Array.from(bytes)); // atomic (3C)
        showToast(writeRes.success ? th.csv.exported : th.csv.exportFailed + writeRes.error);
      } else {
        const blob = new Blob([new Uint8Array(bytes)], { type: 'text/csv;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = defaultName;
        a.click();
        URL.revokeObjectURL(url);
        showToast(th.csv.exported);
      }
    } catch (err: any) {
      showToast(th.csv.exportFailed + (err.message || String(err)));
    }
  };
  const [csvConfirmMessage, setCsvConfirmMessage] = useState('');
  const handleExportCsv = () => {
    const { hasFilter, hasDates } = getCsvExport();
    if (!hasFilter && !hasDates) {
      runCsvExport();
      return;
    }
    // บอกก่อนส่งออก: แถวที่ซ่อนด้วยตัวกรองถูกส่งออกด้วย และปีของวันที่เป็น พ.ศ. หรือ ค.ศ. ตามที่แสดง
    const lines: string[] = [];
    if (hasFilter) lines.push(th.csv.exportFilterConfirm);
    if (hasDates) {
      const mode = useSheetStore.getState().yearMode === 'be' ? th.csv.yearBE : th.csv.yearCE;
      lines.push(th.csv.exportYear.replace('{mode}', mode));
    }
    setCsvConfirmMessage(lines.join(' '));
    setCsvFilterConfirmOpen(true);
  };

  // นำเข้า CSV เป็นชีตใหม่ในไฟล์ที่เปิดอยู่
  const handleImportCsvAsSheet = async () => {
    const electronAPI = (window as any).electronAPI;
    try {
      if (electronAPI?.openCsvFile) {
        const res = await electronAPI.openCsvFile();
        if (!res || res.canceled) return;
        if (res.error || !res.data) {
          showToast(th.csv.readFailed + (res.error || ''));
          return;
        }
        beginCsvImport(new Uint8Array(res.data), res.filePath, 'newSheet');
      } else {
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = '.csv,.txt';
        input.onchange = async (e: any) => {
          const file = e.target.files?.[0];
          if (file) beginCsvImport(new Uint8Array(await file.arrayBuffer()), file.name, 'newSheet');
        };
        input.click();
      }
    } catch (err: any) {
      showToast(th.csv.readFailed + (err.message || String(err)));
    }
  };

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

      // เปิดแอปด้วยไฟล์ (Open with): ขอ path จาก main เมื่อหน้าพร้อมแล้ว แล้วเปิดแทนหน้าต่างเลือกเทมเพลต
      let initialFile: string | null = null;
      if (electronAPI?.getInitialFile) {
        try {
          initialFile = await electronAPI.getInitialFile();
        } catch {}
      }
      if (initialFile) {
        openExternalRef.current?.(initialFile);
        return;
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
          markSavedAndNotify(targetPath);
          showToast(th.messages.fileSaved);
          return true;
        } else {
          showToast(th.messages.fileSaveFailed + writeRes.error);
          return false;
        }
      } else {
        markSavedAndNotify(targetPath);
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
        const defaultSavePath = filePath && !isCsvPath(filePath) ? filePath : `${fileName}.xlsx`;
        const saveDialogRes = await electronAPI.saveFile(defaultSavePath);
        if (saveDialogRes && !saveDialogRes.canceled && saveDialogRes.filePath) {
          const writeRes = await electronAPI.writeFile(
            saveDialogRes.filePath,
            Array.from(buffer)
          );
          if (writeRes.success) {
            markSavedAndNotify(saveDialogRes.filePath);
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
        markSavedAndNotify('ตารางรายรับรายจ่าย.xlsx');
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
    // งานที่มาจาก CSV ไม่มีไฟล์ปลายทาง: บันทึกเป็น .xlsx เสมอ (ห้ามเขียนทับ .csv)
    if (!filePath || isCsvPath(filePath)) {
      return handleSaveAs();
    }

    // บันทึกทับไฟล์ต้นฉบับที่มีกราฟ: ถามทุกครั้งจนกว่าจะบันทึกสำเร็จ (หลังจากนั้นไฟล์ไม่มีกราฟแล้ว)
    if (sourceHasCharts) {
      setChartOverwriteOpen(true);
      return false;
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
            await openFileBytes(uint8, res.filePath);
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
              await openFileBytes(new Uint8Array(buf), file.name);
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
            await openFileBytes(new Uint8Array(res.data), targetPath);
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
              await openFileBytes(new Uint8Array(res.data), targetPath);
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
  // ให้ effect ตอนเริ่มแอปเรียก handleOpenExternal ตัวล่าสุดได้
  openExternalRef.current = handleOpenExternal;

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
        await openFileBytes(new Uint8Array(buf), filePath);
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
      } else if (isCtrl && e.key.toLowerCase() === 'p') {
        e.preventDefault();
        setPrintDialog('print');
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
            backgroundColor: 'var(--drop-overlay)',
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
        onExportCsv={handleExportCsv}
        onImportCsv={handleImportCsvAsSheet}
        onPrint={() => setPrintDialog('print')}
        onExportPdf={() => setPrintDialog('pdf')}
      />

      <FormulaBar />

      {/* แถบเตือนไฟล์ที่มีกราฟ (กราฟเลื่อนไป v0.2) */}
      {sourceHasCharts && !chartBannerDismissed && (
        <div
          role="status"
          style={{
            backgroundColor: 'var(--banner-warn-bg)',
            color: 'var(--banner-warn-fg)',
            borderBottom: '1px solid var(--banner-warn-border)',
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
            <span>{th.banner.charts}</span>
          </div>
          <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
            <button
              style={{
                padding: '4px 10px',
                fontSize: '12px',
                fontWeight: 600,
                backgroundColor: 'var(--bg-app)',
                border: '1px solid var(--banner-warn-border)',
                color: 'var(--banner-warn-fg)',
                borderRadius: '4px',
                cursor: 'pointer',
              }}
              onClick={handleSaveAs}
            >
              💾 {th.banner.saveAsNewBtn}
            </button>
            <button
              style={{
                padding: '2px 8px',
                fontSize: '13px',
                background: 'transparent',
                border: 'none',
                color: 'var(--banner-warn-fg)',
                cursor: 'pointer',
              }}
              onClick={dismissChartBanner}
              title={th.banner.dismiss}
            >
              ✕
            </button>
          </div>
        </div>
      )}

      {/* Unsupported Features Warning Banner (SPEC 6.4) */}
      {(unsupportedFeatures.length > 0 || hiddenSheetNotice > 0) && (
        <div
          style={{
            backgroundColor: 'var(--banner-warn-bg)',
            color: 'var(--banner-warn-fg)',
            borderBottom: '1px solid var(--banner-warn-border)',
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
                backgroundColor: 'var(--bg-app)',
                border: '1px solid var(--banner-warn-border)',
                color: 'var(--banner-warn-fg)',
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
                color: 'var(--banner-warn-fg)',
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
            backgroundColor: 'var(--banner-info-bg)',
            color: 'var(--banner-info-fg)',
            borderBottom: '1px solid var(--banner-info-border)',
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
              style={{ background: 'transparent', border: 'none', cursor: 'pointer', color: 'var(--banner-info-fg)' }}
            >
              ✕
            </button>
          </div>
        </div>
      )}

      {/* สีตัวอักษรในไฟล์อาจมองยากบนพื้นมืด (5C) */}
      {lowContrastNotice && resolvedTheme === 'dark' && !gridWhite && (
        <div
          role="status"
          style={{
            backgroundColor: 'var(--banner-warn-bg)',
            color: 'var(--banner-warn-fg)',
            borderBottom: '1px solid var(--banner-warn-border)',
            padding: '7px 16px',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            fontSize: '13px',
          }}
        >
          <span>⚠️ {th.theme.lowContrastBanner}</span>
          <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
            <button
              onClick={() => {
                useSheetStore.getState().setGridWhite(true);
                setLowContrastNotice(false);
              }}
              style={{ fontSize: '12px', fontWeight: 600 }}
            >
              {th.theme.useWhiteGrid}
            </button>
            <button
              onClick={() => setLowContrastNotice(false)}
              title={th.banner.dismiss}
              aria-label={th.banner.dismiss}
              style={{ background: 'transparent', border: 'none', cursor: 'pointer', color: 'var(--banner-warn-fg)' }}
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
            backgroundColor: 'var(--toast-bg)',
            color: 'var(--toast-fg)',
            padding: '8px 16px',
            borderRadius: '6px',
            boxShadow: 'var(--shadow-menu)',
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

      {/* กล่องยืนยันบันทึกทับไฟล์ที่มีกราฟ */}
      <UnsupportedOverwriteDialog
        isOpen={chartOverwriteOpen}
        features={[]}
        charts
        onOverwrite={async () => {
          // ยอมรับว่าส่วนที่ไม่รองรับจะหายแล้ว ไม่ต้องถามซ้ำด้วยกล่องทั่วไป
          setHasWarnedUnsupportedOverwrite(true);
          setChartOverwriteOpen(false);
          if (filePath) {
            await executeSaveToPath(filePath);
          }
        }}
        onSaveAs={async () => {
          setChartOverwriteOpen(false);
          await handleSaveAs();
        }}
        onCancel={() => {
          setChartOverwriteOpen(false);
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

      {/* แจ้งเรื่องปี พ.ศ. หลังบันทึก (5A) */}
      {beNoticeOpen && (
        <div
          style={{
            position: 'fixed',
            inset: 0,
            backgroundColor: 'var(--overlay)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            zIndex: 3000,
          }}
        >
          <div
            role="alertdialog"
            style={{
              backgroundColor: 'var(--bg-app)',
              borderRadius: '8px',
              padding: '20px 24px',
              width: '420px',
              maxWidth: '92%',
              display: 'flex',
              flexDirection: 'column',
              gap: '12px',
              fontSize: '14px',
              boxShadow: 'var(--shadow-dialog)',
            }}
          >
            <h3 style={{ margin: 0, fontSize: '16px' }}>ℹ️ {th.yearNotice.title}</h3>
            <p style={{ margin: 0 }}>{th.yearNotice.message}</p>
            <label style={{ display: 'flex', alignItems: 'center', gap: '8px', cursor: 'pointer' }}>
              <input
                type="checkbox"
                checked={beNoticeDontShow}
                onChange={(e) => setBeNoticeDontShow(e.target.checked)}
              />
              {th.yearNotice.dontShowAgain}
            </label>
            <div style={{ display: 'flex', justifyContent: 'flex-end' }}>
              <button autoFocus onClick={closeBeNotice} style={{ fontWeight: 600 }}>
                {th.dialog.ok}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* พิมพ์ / PDF (5B) */}
      <PrintDialog
        isOpen={printDialog !== null}
        focus={printDialog ?? 'print'}
        onClose={() => setPrintDialog(null)}
        onToast={showToast}
      />

      {/* นำเข้า CSV (4D) */}
      <CsvImportDialog />

      {/* ส่งออก CSV ขณะมีตัวกรอง: บอกก่อนว่าจะส่งออกทุกแถว */}
      <ConfirmDialog
        isOpen={csvFilterConfirmOpen}
        message={csvConfirmMessage}
        confirmText={th.csv.exportContinue}
        cancelText={th.dialog.cancel}
        onConfirm={() => {
          setCsvFilterConfirmOpen(false);
          runCsvExport();
        }}
        onCancel={() => setCsvFilterConfirmOpen(false)}
      />

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
