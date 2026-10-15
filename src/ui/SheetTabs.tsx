import React, { useState, useRef, useEffect } from 'react';
import { useSheetStore } from './store';
import { th } from '../i18n/th';
import { ConfirmDialog } from './Dialogs';

export const SheetTabs: React.FC = () => {
  const {
    workbook,
    activeSheetIndex,
    switchSheet,
    addSheet,
    duplicateSheet,
    moveSheet,
    renameSheet,
    deleteSheet,
    pendingDeleteSheet,
    confirmDeleteSheet,
    cancelDeleteSheet,
  } = useSheetStore();

  const [editingIndex, setEditingIndex] = useState<number | null>(null);
  const [editingName, setEditingName] = useState<string>('');
  const [renameError, setRenameError] = useState<string | null>(null);

  const [contextMenu, setContextMenu] = useState<{
    x: number;
    y: number;
    sheetIndex: number;
  } | null>(null);

  const [draggedIndex, setDraggedIndex] = useState<number | null>(null);
  const [dragOverIndex, setDragOverIndex] = useState<number | null>(null);

  const editInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    if (editingIndex !== null && editInputRef.current) {
      editInputRef.current.focus();
      editInputRef.current.select();
    }
  }, [editingIndex]);

  // Close context menu on window click
  useEffect(() => {
    const handleWindowClick = () => {
      setContextMenu(null);
    };
    window.addEventListener('click', handleWindowClick);
    return () => window.removeEventListener('click', handleWindowClick);
  }, []);

  const handleStartRename = (index: number) => {
    setEditingIndex(index);
    setEditingName(workbook.sheets[index].name);
    setRenameError(null);
    setContextMenu(null);
  };

  const handleCommitRename = (index: number) => {
    if (editingIndex === null) return;
    const res = renameSheet(index, editingName);
    if (!res.success) {
      setRenameError(res.error || th.sheetTabs.nameEmpty);
    } else {
      setEditingIndex(null);
      setRenameError(null);
    }
  };

  const handleCancelRename = () => {
    setEditingIndex(null);
    setRenameError(null);
  };

  // Drag & drop handlers
  const handleDragStart = (e: React.DragEvent, index: number) => {
    setDraggedIndex(index);
    e.dataTransfer.effectAllowed = 'move';
  };

  const handleDragOver = (e: React.DragEvent, index: number) => {
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    if (dragOverIndex !== index) {
      setDragOverIndex(index);
    }
  };

  const handleDrop = (e: React.DragEvent, dropIndex: number) => {
    e.preventDefault();
    if (draggedIndex !== null && draggedIndex !== dropIndex) {
      moveSheet(draggedIndex, dropIndex);
    }
    setDraggedIndex(null);
    setDragOverIndex(null);
  };

  const handleDragEnd = () => {
    setDraggedIndex(null);
    setDragOverIndex(null);
  };

  return (
    <div
      style={{
        display: 'flex',
        alignItems: 'center',
        height: '34px',
        backgroundColor: 'var(--bg-toolbar, #f8fafc)',
        borderTop: '1px solid var(--border-color, #e2e8f0)',
        borderBottom: '1px solid var(--border-color, #e2e8f0)',
        padding: '0 8px',
        userSelect: 'none',
        gap: '4px',
        zIndex: 5,
        overflowX: 'auto',
      }}
    >
      {/* Add Sheet (+) Button */}
      <button
        onClick={addSheet}
        title={th.sheetTabs.addSheet}
        style={{
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          width: '26px',
          height: '26px',
          borderRadius: '4px',
          border: '1px solid transparent',
          backgroundColor: 'transparent',
          color: 'var(--fg-main, #334155)',
          cursor: 'pointer',
          fontSize: '18px',
          fontWeight: 600,
          padding: 0,
          flexShrink: 0,
        }}
        onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = 'rgba(0,0,0,0.06)')}
        onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'transparent')}
      >
        ＋
      </button>

      <div style={{ width: '1px', height: '18px', backgroundColor: 'var(--border-color, #cbd5e1)', margin: '0 4px', flexShrink: 0 }} />

      {/* Sheet Tabs List */}
      <div style={{ display: 'flex', alignItems: 'center', gap: '4px', flex: 1, overflowX: 'auto' }}>
        {workbook.sheets.map((sheet, idx) => {
          const isActive = idx === activeSheetIndex;
          const isBeingEdited = idx === editingIndex;
          const isDragOver = idx === dragOverIndex && draggedIndex !== idx;

          return (
            <div
              key={sheet.name + '-' + idx}
              draggable={!isBeingEdited}
              onDragStart={(e) => handleDragStart(e, idx)}
              onDragOver={(e) => handleDragOver(e, idx)}
              onDrop={(e) => handleDrop(e, idx)}
              onDragEnd={handleDragEnd}
              onClick={() => {
                if (!isBeingEdited && idx !== activeSheetIndex) {
                  switchSheet(idx);
                }
              }}
              onDoubleClick={() => handleStartRename(idx)}
              onContextMenu={(e) => {
                e.preventDefault();
                setContextMenu({
                  x: e.clientX,
                  y: e.clientY,
                  sheetIndex: idx,
                });
              }}
              style={{
                display: 'flex',
                alignItems: 'center',
                padding: '4px 12px',
                height: '26px',
                borderRadius: '4px 4px 0 0',
                fontSize: '13px',
                fontWeight: isActive ? 600 : 400,
                backgroundColor: isActive ? 'var(--bg-app, #ffffff)' : 'transparent',
                color: isActive ? 'var(--color-primary, #0f766e)' : 'var(--fg-muted, #64748b)',
                borderBottom: isActive ? '2px solid var(--color-primary, #0f766e)' : '2px solid transparent',
                borderLeft: isDragOver ? '2px solid var(--color-primary, #0f766e)' : '1px solid transparent',
                borderRight: '1px solid transparent',
                borderTop: isActive ? '1px solid var(--border-color, #e2e8f0)' : '1px solid transparent',
                cursor: 'pointer',
                whiteSpace: 'nowrap',
                transition: 'all 0.15s ease',
                position: 'relative',
                flexShrink: 0,
              }}
              title={`${sheet.name} (ดับเบิลคลิกเพื่อเปลี่ยนชื่อ, คลิกขวาสำหรับตัวเลือก)`}
            >
              {isBeingEdited ? (
                <div style={{ display: 'flex', alignItems: 'center', position: 'relative' }}>
                  <input
                    ref={editInputRef}
                    type="text"
                    value={editingName}
                    maxLength={31}
                    onChange={(e) => {
                      setEditingName(e.target.value);
                      setRenameError(null);
                    }}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter') {
                        e.preventDefault();
                        handleCommitRename(idx);
                      } else if (e.key === 'Escape') {
                        e.preventDefault();
                        handleCancelRename();
                      }
                    }}
                    onBlur={() => handleCommitRename(idx)}
                    style={{
                      height: '20px',
                      padding: '0 4px',
                      fontSize: '13px',
                      border: '1px solid var(--color-primary, #0f766e)',
                      borderRadius: '2px',
                      outline: 'none',
                      backgroundColor: 'var(--bg-app, #ffffff)',
                      color: 'var(--fg-main, #0f172a)',
                      width: '120px',
                    }}
                  />
                  {renameError && (
                    <div
                      style={{
                        position: 'absolute',
                        bottom: '28px',
                        left: 0,
                        backgroundColor: '#ef4444',
                        color: '#ffffff',
                        padding: '4px 8px',
                        borderRadius: '4px',
                        fontSize: '11px',
                        whiteSpace: 'nowrap',
                        zIndex: 100,
                        boxShadow: '0 2px 8px rgba(0,0,0,0.15)',
                      }}
                    >
                      {renameError}
                    </div>
                  )}
                </div>
              ) : (
                <span>{sheet.name}</span>
              )}
            </div>
          );
        })}
      </div>

      {/* Context Menu for Sheet Tabs */}
      {contextMenu && (
        <div
          onClick={(e) => e.stopPropagation()}
          style={{
            position: 'fixed',
            top: contextMenu.y - 120,
            left: contextMenu.x,
            backgroundColor: 'var(--bg-app, #ffffff)',
            border: '1px solid var(--border-color, #e2e8f0)',
            borderRadius: '6px',
            boxShadow: '0 4px 16px rgba(0,0,0,0.12)',
            zIndex: 1000,
            padding: '4px 0',
            minWidth: '150px',
            fontSize: '13px',
          }}
        >
          <div
            onClick={() => handleStartRename(contextMenu.sheetIndex)}
            style={{ padding: '6px 12px', cursor: 'pointer' }}
            onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = 'var(--bg-toolbar, #f1f5f9)')}
            onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'transparent')}
          >
            ✏️ {th.sheetTabs.rename}
          </div>
          <div
            onClick={() => {
              duplicateSheet(contextMenu.sheetIndex);
              setContextMenu(null);
            }}
            style={{ padding: '6px 12px', cursor: 'pointer' }}
            onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = 'var(--bg-toolbar, #f1f5f9)')}
            onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'transparent')}
          >
            📋 {th.sheetTabs.duplicate}
          </div>
          <div style={{ height: '1px', backgroundColor: 'var(--border-color, #e2e8f0)', margin: '4px 0' }} />
          <div
            onClick={() => {
              if (contextMenu.sheetIndex > 0) {
                moveSheet(contextMenu.sheetIndex, contextMenu.sheetIndex - 1);
              }
              setContextMenu(null);
            }}
            style={{
              padding: '6px 12px',
              cursor: contextMenu.sheetIndex > 0 ? 'pointer' : 'not-allowed',
              opacity: contextMenu.sheetIndex > 0 ? 1 : 0.5,
            }}
            onMouseEnter={(e) => {
              if (contextMenu.sheetIndex > 0) e.currentTarget.style.backgroundColor = 'var(--bg-toolbar, #f1f5f9)';
            }}
            onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'transparent')}
          >
            ◀ {th.sheetTabs.moveLeft}
          </div>
          <div
            onClick={() => {
              if (contextMenu.sheetIndex < workbook.sheets.length - 1) {
                moveSheet(contextMenu.sheetIndex, contextMenu.sheetIndex + 1);
              }
              setContextMenu(null);
            }}
            style={{
              padding: '6px 12px',
              cursor: contextMenu.sheetIndex < workbook.sheets.length - 1 ? 'pointer' : 'not-allowed',
              opacity: contextMenu.sheetIndex < workbook.sheets.length - 1 ? 1 : 0.5,
            }}
            onMouseEnter={(e) => {
              if (contextMenu.sheetIndex < workbook.sheets.length - 1)
                e.currentTarget.style.backgroundColor = 'var(--bg-toolbar, #f1f5f9)';
            }}
            onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'transparent')}
          >
            ▶ {th.sheetTabs.moveRight}
          </div>
          <div style={{ height: '1px', backgroundColor: 'var(--border-color, #e2e8f0)', margin: '4px 0' }} />
          <div
            onClick={() => {
              if (workbook.sheets.length > 1) {
                deleteSheet(contextMenu.sheetIndex);
              }
              setContextMenu(null);
            }}
            style={{
              padding: '6px 12px',
              color: workbook.sheets.length > 1 ? '#ef4444' : '#94a3b8',
              cursor: workbook.sheets.length > 1 ? 'pointer' : 'not-allowed',
            }}
            onMouseEnter={(e) => {
              if (workbook.sheets.length > 1) e.currentTarget.style.backgroundColor = '#fef2f2';
            }}
            onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'transparent')}
          >
            🗑️ {th.sheetTabs.delete}
          </div>
        </div>
      )}

      {/* Confirm Sheet Deletion Modal */}
      {pendingDeleteSheet && (
        <ConfirmDialog
          isOpen={true}
          danger={true}
          title={th.sheetTabs.confirmDeleteTitle}
          message={
            pendingDeleteSheet.referencingFormulaCount > 0
              ? `${th.sheetTabs.confirmDeleteWithRefs.replace(
                  '{count}',
                  String(pendingDeleteSheet.referencingFormulaCount)
                )} (${th.sheetTabs.confirmDeleteMessage.replace('{name}', pendingDeleteSheet.sheetName)})`
              : th.sheetTabs.confirmDeleteMessage.replace('{name}', pendingDeleteSheet.sheetName)
          }
          confirmText={th.dialog.delete}
          cancelText={th.dialog.cancel}
          onConfirm={confirmDeleteSheet}
          onCancel={cancelDeleteSheet}
        />
      )}
    </div>
  );
};
