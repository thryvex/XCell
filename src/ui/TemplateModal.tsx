import React, { useEffect, useRef } from 'react';
import { th } from '../i18n/th';

interface TemplateModalProps {
  isOpen: boolean;
  onSelectBlank: () => void;
  onSelectLedger: () => void;
  onSelectInventory: () => void;
  onSelectRecent: (filePath: string) => void;
  onClose: () => void; // Closing modal = opens blank table
  recentFiles: string[];
}

export const TemplateModal: React.FC<TemplateModalProps> = ({
  isOpen,
  onSelectBlank,
  onSelectLedger,
  onSelectInventory,
  onSelectRecent,
  onClose,
  recentFiles,
}) => {
  const blankCardRef = useRef<HTMLButtonElement>(null);

  useEffect(() => {
    if (isOpen) {
      // Focus "ตารางว่าง" by default so pressing Enter immediately selects blank table
      setTimeout(() => {
        blankCardRef.current?.focus();
      }, 50);
    }
  }, [isOpen]);

  useEffect(() => {
    if (!isOpen) return;

    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.preventDefault();
        onClose(); // Close = Blank table
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [isOpen, onClose]);

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
        zIndex: 1050,
      }}
      onClick={(e) => {
        if (e.target === e.currentTarget) {
          onClose(); // Click outside = blank table
        }
      }}
    >
      <div
        style={{
          backgroundColor: 'var(--bg-surface)',
          border: '1px solid var(--border-color)',
          borderRadius: '10px',
          padding: '24px 28px',
          width: '640px',
          maxWidth: '92%',
          maxHeight: '90vh',
          overflowY: 'auto',
          boxShadow: 'var(--shadow-dialog)',
          display: 'flex',
          flexDirection: 'column',
          gap: '20px',
        }}
      >
        {/* Header */}
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
          <div>
            <h2 style={{ margin: 0, fontSize: '20px', fontWeight: 600, color: 'var(--fg-main)' }}>
              {th.templates.title}
            </h2>
            <p style={{ margin: '4px 0 0', fontSize: '13px', color: 'var(--fg-muted)' }}>
              {th.templates.subtitle}
            </p>
          </div>
          <button
            onClick={onClose}
            style={{
              background: 'transparent',
              border: 'none',
              fontSize: '18px',
              cursor: 'pointer',
              color: 'var(--fg-muted)',
              padding: '2px 8px',
            }}
            title="ปิด (เริ่มตารางว่าง)"
          >
            ✕
          </button>
        </div>

        {/* Template choices grid */}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '14px' }}>
          {/* 1. Blank Workbook (Focused default) */}
          <button
            ref={blankCardRef}
            onClick={onSelectBlank}
            style={{
              display: 'flex',
              flexDirection: 'column',
              alignItems: 'center',
              textAlign: 'center',
              padding: '20px 14px',
              borderRadius: '8px',
              border: '2px solid var(--color-primary)',
              backgroundColor: 'var(--bg-app)',
              cursor: 'pointer',
              gap: '10px',
              outline: 'none',
              boxShadow: 'var(--shadow-menu)',
              transition: 'transform 0.1s, border-color 0.1s',
            }}
            onMouseEnter={(e) => (e.currentTarget.style.transform = 'translateY(-2px)')}
            onMouseLeave={(e) => (e.currentTarget.style.transform = 'translateY(0)')}
          >
            <div style={{ fontSize: '32px' }}>📄</div>
            <div style={{ fontWeight: 600, fontSize: '15px', color: 'var(--color-primary)' }}>
              {th.templates.blank}
            </div>
            <div style={{ fontSize: '12px', color: 'var(--fg-muted)', lineHeight: 1.4 }}>
              {th.templates.blankDesc}
            </div>
          </button>

          {/* 2. Ledger Template */}
          <button
            onClick={onSelectLedger}
            style={{
              display: 'flex',
              flexDirection: 'column',
              alignItems: 'center',
              textAlign: 'center',
              padding: '20px 14px',
              borderRadius: '8px',
              border: '1px solid var(--border-color)',
              backgroundColor: 'var(--bg-app)',
              cursor: 'pointer',
              gap: '10px',
              transition: 'transform 0.1s, border-color 0.1s',
            }}
            onMouseEnter={(e) => {
              e.currentTarget.style.transform = 'translateY(-2px)';
              e.currentTarget.style.borderColor = 'var(--color-primary)';
            }}
            onMouseLeave={(e) => {
              e.currentTarget.style.transform = 'translateY(0)';
              e.currentTarget.style.borderColor = 'var(--border-color)';
            }}
          >
            <div style={{ fontSize: '32px' }}>💰</div>
            <div style={{ fontWeight: 600, fontSize: '15px', color: 'var(--fg-main)' }}>
              {th.templates.ledger}
            </div>
            <div style={{ fontSize: '12px', color: 'var(--fg-muted)', lineHeight: 1.4 }}>
              {th.templates.ledgerDesc}
            </div>
          </button>

          {/* 3. Product Inventory Template */}
          <button
            onClick={onSelectInventory}
            style={{
              display: 'flex',
              flexDirection: 'column',
              alignItems: 'center',
              textAlign: 'center',
              padding: '20px 14px',
              borderRadius: '8px',
              border: '1px solid var(--border-color)',
              backgroundColor: 'var(--bg-app)',
              cursor: 'pointer',
              gap: '10px',
              transition: 'transform 0.1s, border-color 0.1s',
            }}
            onMouseEnter={(e) => {
              e.currentTarget.style.transform = 'translateY(-2px)';
              e.currentTarget.style.borderColor = 'var(--color-primary)';
            }}
            onMouseLeave={(e) => {
              e.currentTarget.style.transform = 'translateY(0)';
              e.currentTarget.style.borderColor = 'var(--border-color)';
            }}
          >
            <div style={{ fontSize: '32px' }}>📦</div>
            <div style={{ fontWeight: 600, fontSize: '15px', color: 'var(--fg-main)' }}>
              {th.templates.inventory}
            </div>
            <div style={{ fontSize: '12px', color: 'var(--fg-muted)', lineHeight: 1.4 }}>
              {th.templates.inventoryDesc}
            </div>
          </button>
        </div>

        {/* Recent files section */}
        <div style={{ marginTop: '8px' }}>
          <div style={{ fontSize: '13px', fontWeight: 600, color: 'var(--fg-main)', marginBottom: '8px' }}>
            🕒 {th.templates.recentFilesTitle}
          </div>
          {recentFiles.length === 0 ? (
            <div style={{ fontSize: '13px', color: 'var(--fg-muted)', fontStyle: 'italic', padding: '6px 0' }}>
              {th.templates.noRecent}
            </div>
          ) : (
            <div
              style={{
                maxHeight: '140px',
                overflowY: 'auto',
                border: '1px solid var(--border-color)',
                borderRadius: '6px',
                backgroundColor: 'var(--bg-app)',
              }}
            >
              {recentFiles.slice(0, 5).map((filePath) => {
                const baseName = filePath.split(/[\\/]/).pop() || filePath;
                return (
                  <div
                    key={filePath}
                    onClick={() => onSelectRecent(filePath)}
                    style={{
                      padding: '8px 12px',
                      borderBottom: '1px solid var(--border-color)',
                      cursor: 'pointer',
                      display: 'flex',
                      justifyContent: 'space-between',
                      alignItems: 'center',
                    }}
                    onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = 'var(--bg-hover)')}
                    onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'transparent')}
                  >
                    <div style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                      <span style={{ fontWeight: 500, fontSize: '13px', color: 'var(--fg-main)' }}>
                        📄 {baseName}
                      </span>
                      <span style={{ fontSize: '11px', color: 'var(--fg-muted)', marginLeft: '8px' }}>
                        {filePath}
                      </span>
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>
      </div>
    </div>
  );
};
