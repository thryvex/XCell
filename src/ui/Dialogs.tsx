import React, { useState } from 'react';
import { th } from '../i18n/th';

interface ConfirmDialogProps {
  isOpen: boolean;
  title?: string;
  message: string;
  confirmText?: string;
  cancelText?: string;
  danger?: boolean;
  onConfirm: () => void;
  onCancel: () => void;
}

export const ConfirmDialog: React.FC<ConfirmDialogProps> = ({
  isOpen,
  title = th.dialog.warning,
  message,
  confirmText = th.dialog.ok,
  cancelText = th.dialog.cancel,
  danger = false,
  onConfirm,
  onCancel,
}) => {
  if (!isOpen) return null;

  return (
    <div
      style={{
        position: 'fixed',
        top: 0,
        left: 0,
        right: 0,
        bottom: 0,
        backgroundColor: 'var(--overlay)',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        zIndex: 1000,
      }}
    >
      <div
        style={{
          backgroundColor: 'var(--bg-app)',
          border: '1px solid var(--border-color)',
          borderRadius: '8px',
          padding: '20px 24px',
          width: '420px',
          maxWidth: '90%',
          boxShadow: 'var(--shadow-dialog)',
          display: 'flex',
          flexDirection: 'column',
          gap: '16px',
        }}
      >
        <h3 style={{ margin: 0, fontSize: '16px', fontWeight: 600 }}>{title}</h3>
        <p style={{ margin: 0, color: 'var(--fg-main)', lineHeight: 1.5 }}>{message}</p>
        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '8px', marginTop: '8px' }}>
          <button onClick={onCancel}>{cancelText}</button>
          <button
            className={danger ? '' : 'primary'}
            style={
              danger
                ? {
                    backgroundColor: 'var(--color-danger)',
                    color: 'var(--fg-on-primary)',
                    border: 'none',
                    borderRadius: '4px',
                    padding: '6px 14px',
                    fontWeight: 600,
                    cursor: 'pointer',
                  }
                : undefined
            }
            onClick={onConfirm}
          >
            {confirmText}
          </button>
        </div>
      </div>
    </div>
  );
};

interface UnsavedChangesDialogProps {
  isOpen: boolean;
  onSave: () => void;
  onDontSave: () => void;
  onCancel: () => void;
}

export const UnsavedChangesDialog: React.FC<UnsavedChangesDialogProps> = ({
  isOpen,
  onSave,
  onDontSave,
  onCancel,
}) => {
  if (!isOpen) return null;

  return (
    <div
      style={{
        position: 'fixed',
        top: 0,
        left: 0,
        right: 0,
        bottom: 0,
        backgroundColor: 'var(--overlay)',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        zIndex: 1000,
      }}
    >
      <div
        style={{
          backgroundColor: 'var(--bg-surface)',
          border: '1px solid var(--border-color)',
          borderRadius: '8px',
          padding: '20px 24px',
          width: '440px',
          maxWidth: '90%',
          boxShadow: 'var(--shadow-dialog)',
          display: 'flex',
          flexDirection: 'column',
          gap: '16px',
        }}
      >
        <h3 style={{ margin: 0, fontSize: '16px', fontWeight: 600, color: 'var(--fg-main)' }}>
          ⚠️ {th.dialog.unsavedChangesTitle}
        </h3>
        <p style={{ margin: 0, color: 'var(--fg-main)', lineHeight: 1.5 }}>
          {th.dialog.confirmUnsaved}
        </p>
        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '8px', marginTop: '8px' }}>
          <button onClick={onCancel}>{th.dialog.cancel}</button>
          <button
            style={{
              backgroundColor: 'var(--color-danger)',
              color: 'var(--fg-on-primary)',
              border: 'none',
              borderRadius: '4px',
              padding: '6px 14px',
              fontWeight: 600,
              cursor: 'pointer',
            }}
            onClick={onDontSave}
          >
            {th.dialog.dontSave}
          </button>
          <button className="primary" onClick={onSave}>
            💾 {th.dialog.save}
          </button>
        </div>
      </div>
    </div>
  );
};

export interface RecoveryDialogProps {
  isOpen: boolean;
  recoveryItems: any[];
  onRecover: (item: any) => void;
  onDiscard: () => void;
}

export const RecoveryDialog: React.FC<RecoveryDialogProps> = ({
  isOpen,
  recoveryItems,
  onRecover,
  onDiscard,
}) => {
  const [selectedIndex, setSelectedIndex] = useState(0);

  if (!isOpen || recoveryItems.length === 0) return null;

  return (
    <div
      style={{
        position: 'fixed',
        top: 0,
        left: 0,
        right: 0,
        bottom: 0,
        backgroundColor: 'var(--overlay)',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        zIndex: 1100,
      }}
    >
      <div
        style={{
          backgroundColor: 'var(--bg-surface)',
          border: '1px solid var(--border-color)',
          borderRadius: '8px',
          padding: '24px',
          width: '520px',
          maxWidth: '92%',
          boxShadow: 'var(--shadow-dialog)',
          display: 'flex',
          flexDirection: 'column',
          gap: '16px',
        }}
      >
        <h3 style={{ margin: 0, fontSize: '18px', fontWeight: 600, color: 'var(--color-primary)' }}>
          🔄 {th.dialog.foundRecoveryTitle}
        </h3>
        <p style={{ margin: 0, color: 'var(--fg-main)', lineHeight: 1.5, fontSize: '14px' }}>
          {th.dialog.foundRecoveryPrompt}
        </p>

        {recoveryItems.length > 1 && (
          <div style={{ fontSize: '13px', fontWeight: 500, color: 'var(--fg-muted)' }}>
            {th.dialog.recoveryListTitle}
          </div>
        )}

        <div
          style={{
            maxHeight: '180px',
            overflowY: 'auto',
            border: '1px solid var(--border-color)',
            borderRadius: '6px',
            backgroundColor: 'var(--bg-app)',
          }}
        >
          {recoveryItems.map((item, idx) => {
            const isSelected = idx === selectedIndex;
            const dateStr = item.savedAt
              ? new Date(item.savedAt).toLocaleString('th-TH')
              : '-';

            return (
              <div
                key={item.id || idx}
                onClick={() => setSelectedIndex(idx)}
                style={{
                  padding: '10px 14px',
                  borderBottom: idx < recoveryItems.length - 1 ? '1px solid var(--border-color)' : 'none',
                  backgroundColor: isSelected ? 'var(--bg-selection)' : 'transparent',
                  cursor: 'pointer',
                  display: 'flex',
                  justifyContent: 'space-between',
                  alignItems: 'center',
                }}
              >
                <div>
                  <div style={{ fontWeight: 600, fontSize: '13px', color: 'var(--fg-main)' }}>
                    📄 {item.fileName || 'ตารางไม่ระบุชื่อ'}
                  </div>
                  <div style={{ fontSize: '12px', color: 'var(--fg-muted)' }}>
                    {item.originalPath || '(ยังไม่เคยบันทึกเป็นไฟล์)'}
                  </div>
                </div>
                <div style={{ fontSize: '11px', color: 'var(--fg-muted)', textAlign: 'right' }}>
                  {dateStr}
                </div>
              </div>
            );
          })}
        </div>

        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '8px', marginTop: '8px' }}>
          <button
            style={{
              backgroundColor: 'var(--color-danger)',
              color: 'var(--fg-on-primary)',
              border: 'none',
              borderRadius: '4px',
              padding: '6px 14px',
              fontWeight: 600,
              cursor: 'pointer',
            }}
            onClick={onDiscard}
          >
            🗑️ {th.dialog.discard}
          </button>
          <button
            className="primary"
            onClick={() => onRecover(recoveryItems[selectedIndex])}
          >
            ✓ {th.dialog.recover}
          </button>
        </div>
      </div>
    </div>
  );
};

interface UnsupportedOverwriteDialogProps {
  isOpen: boolean;
  features: string[];
  onOverwrite: () => void;
  onSaveAs: () => void;
  onCancel: () => void;
}

export const UnsupportedOverwriteDialog: React.FC<UnsupportedOverwriteDialogProps> = ({
  isOpen,
  features,
  onOverwrite,
  onSaveAs,
  onCancel,
}) => {
  if (!isOpen) return null;

  return (
    <div
      style={{
        position: 'fixed',
        top: 0,
        left: 0,
        right: 0,
        bottom: 0,
        backgroundColor: 'var(--overlay)',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        zIndex: 1000,
      }}
    >
      <div
        style={{
          backgroundColor: 'var(--bg-surface)',
          border: '1px solid var(--border-color)',
          borderRadius: '8px',
          padding: '20px 24px',
          width: '480px',
          maxWidth: '90%',
          boxShadow: 'var(--shadow-dialog)',
          display: 'flex',
          flexDirection: 'column',
          gap: '16px',
        }}
      >
        <h3 style={{ margin: 0, fontSize: '16px', fontWeight: 600, color: 'var(--color-warning)' }}>
          ⚠️ {th.dialog.overwriteWarningTitle}
        </h3>
        <p style={{ margin: 0, color: 'var(--fg-main)', lineHeight: 1.5, fontSize: '14px' }}>
          {th.dialog.overwriteWarningPrompt.replace('{features}', features.join(', '))}
        </p>
        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '8px', marginTop: '8px' }}>
          <button onClick={onCancel}>{th.dialog.cancel}</button>
          <button
            style={{
              backgroundColor: 'var(--color-danger)',
              color: 'var(--fg-on-primary)',
              border: 'none',
              borderRadius: '4px',
              padding: '6px 14px',
              fontWeight: 600,
              cursor: 'pointer',
            }}
            onClick={onOverwrite}
          >
            {th.dialog.overwrite}
          </button>
          <button className="primary" onClick={onSaveAs}>
            💾 {th.dialog.saveAsNew}
          </button>
        </div>
      </div>
    </div>
  );
};
