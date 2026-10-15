import React, { useState } from 'react';
import { th } from '../i18n/th';
import { APP_INFO } from '../appInfo';

interface HelpModalProps {
  isOpen: boolean;
  onClose: () => void;
  onOpenWelcomeGuide: () => void;
}

export const HelpModal: React.FC<HelpModalProps> = ({
  isOpen,
  onClose,
  onOpenWelcomeGuide,
}) => {
  const [activeTab, setActiveTab] = useState<'gs' | 'shortcuts' | 'features' | 'about'>('gs');

  if (!isOpen) return null;

  const SHORTCUTS = th.help.shortcuts;

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
      onClick={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
    >
      <div
        style={{
          backgroundColor: 'var(--bg-surface)',
          border: '1px solid var(--border-color)',
          borderRadius: '10px',
          width: '680px',
          maxWidth: '94%',
          height: '520px',
          maxHeight: '90vh',
          boxShadow: 'var(--shadow-dialog)',
          display: 'flex',
          flexDirection: 'column',
          overflow: 'hidden',
        }}
      >
        {/* Header */}
        <div
          style={{
            padding: '16px 20px',
            borderBottom: '1px solid var(--border-color)',
            display: 'flex',
            justifyContent: 'space-between',
            alignItems: 'center',
          }}
        >
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            <span style={{ fontSize: '20px' }}>❓</span>
            <h2 style={{ margin: 0, fontSize: '18px', fontWeight: 600, color: 'var(--fg-main)' }}>
              {th.help.title}
            </h2>
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
          >
            ✕
          </button>
        </div>

        {/* Navigation Tabs */}
        <div
          style={{
            display: 'flex',
            borderBottom: '1px solid var(--border-color)',
            backgroundColor: 'var(--bg-app)',
            padding: '0 16px',
            gap: '8px',
          }}
        >
          <button
            onClick={() => setActiveTab('gs')}
            style={{
              border: 'none',
              borderBottom: activeTab === 'gs' ? '2px solid var(--color-primary)' : '2px solid transparent',
              backgroundColor: 'transparent',
              padding: '10px 14px',
              fontWeight: activeTab === 'gs' ? 600 : 400,
              color: activeTab === 'gs' ? 'var(--color-primary)' : 'var(--fg-main)',
              cursor: 'pointer',
              borderRadius: 0,
            }}
          >
            🌐 {th.help.tabGoogleSheets}
          </button>
          <button
            onClick={() => setActiveTab('shortcuts')}
            style={{
              border: 'none',
              borderBottom: activeTab === 'shortcuts' ? '2px solid var(--color-primary)' : '2px solid transparent',
              backgroundColor: 'transparent',
              padding: '10px 14px',
              fontWeight: activeTab === 'shortcuts' ? 600 : 400,
              color: activeTab === 'shortcuts' ? 'var(--color-primary)' : 'var(--fg-main)',
              cursor: 'pointer',
              borderRadius: 0,
            }}
          >
            ⌨️ {th.help.tabShortcuts}
          </button>
          <button
            onClick={() => setActiveTab('features')}
            style={{
              border: 'none',
              borderBottom: activeTab === 'features' ? '2px solid var(--color-primary)' : '2px solid transparent',
              backgroundColor: 'transparent',
              padding: '10px 14px',
              fontWeight: activeTab === 'features' ? 600 : 400,
              color: activeTab === 'features' ? 'var(--color-primary)' : 'var(--fg-main)',
              cursor: 'pointer',
              borderRadius: 0,
            }}
          >
            ✨ {th.help.tabFeatures}
          </button>
          <button
            onClick={() => setActiveTab('about')}
            style={{
              border: 'none',
              borderBottom: activeTab === 'about' ? '2px solid var(--color-primary)' : '2px solid transparent',
              backgroundColor: 'transparent',
              padding: '10px 14px',
              fontWeight: activeTab === 'about' ? 600 : 400,
              color: activeTab === 'about' ? 'var(--color-primary)' : 'var(--fg-main)',
              cursor: 'pointer',
              borderRadius: 0,
            }}
          >
            ℹ️ {th.help.tabAbout}
          </button>
        </div>

        {/* Tab Content */}
        <div style={{ flex: 1, padding: '20px', overflowY: 'auto' }}>
          {/* Tab 1: Google Sheets Guide */}
          {activeTab === 'gs' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
              <h3 style={{ margin: 0, fontSize: '16px', color: 'var(--fg-main)' }}>
                {th.help.gsTitle}
              </h3>
              <div
                style={{
                  backgroundColor: 'var(--bg-app)',
                  border: '1px solid var(--border-color)',
                  borderRadius: '8px',
                  padding: '16px',
                  display: 'flex',
                  flexDirection: 'column',
                  gap: '12px',
                  lineHeight: 1.6,
                  fontSize: '14px',
                }}
              >
                <div>{th.help.gsStep1}</div>
                <div>{th.help.gsStep2}</div>
                <div>{th.help.gsStep3}</div>
                <div style={{ marginTop: '8px', color: 'var(--fg-muted)' }}>{th.help.gsYearNote}</div>
              </div>
              <div
                style={{
                  padding: '12px 14px',
                  backgroundColor: 'var(--bg-selection)',
                  border: '1px solid var(--color-primary)',
                  borderRadius: '6px',
                  fontSize: '13px',
                  color: 'var(--fg-main)',
                }}
              >
                💡 <b>ข้อดี:</b> สูตรคำนวณทั้งหมด, เซลที่ผสาน, สีและฟอนต์ รวมถึงรูปแบบ ฿ บาท จะแสดงผลและทำงานได้ทันทีบน Google Sheets โดยไม่ต้องปรับแต่งเพิ่ม
              </div>
            </div>
          )}

          {/* Tab: ฟีเจอร์เพิ่มเติม (Phase 4) */}
          {activeTab === 'features' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px', fontSize: '13px' }}>
              {th.help.features.map((section) => (
                <div key={section.title}>
                  <div style={{ fontWeight: 600, marginBottom: '4px' }}>{section.title}</div>
                  <ul style={{ margin: 0, paddingLeft: '20px', lineHeight: 1.6 }}>
                    {section.items.map((item) => (
                      <li key={item}>{item}</li>
                    ))}
                  </ul>
                </div>
              ))}
            </div>
          )}

          {/* Tab 2: Keyboard Shortcuts */}
          {activeTab === 'shortcuts' && (
            <div>
              <div style={{ marginBottom: '10px', fontSize: '12px', color: 'var(--fg-muted)' }}>
                * สำหรับผู้ใช้งาน macOS ให้ใช้ปุ่ม <code>Cmd (⌘)</code> แทน <code>Ctrl</code>
              </div>
              <table
                style={{
                  width: '100%',
                  borderCollapse: 'collapse',
                  fontSize: '13px',
                }}
              >
                <thead>
                  <tr style={{ backgroundColor: 'var(--bg-app)', borderBottom: '2px solid var(--border-color)' }}>
                    <th style={{ textAlign: 'left', padding: '8px 10px', width: '220px' }}>คีย์ลัด</th>
                    <th style={{ textAlign: 'left', padding: '8px 10px' }}>การทำงาน</th>
                  </tr>
                </thead>
                <tbody>
                  {SHORTCUTS.map((s, idx) => (
                    <tr
                      key={s.key}
                      style={{
                        borderBottom: '1px solid var(--border-color)',
                        backgroundColor: idx % 2 === 0 ? 'transparent' : 'var(--bg-app)',
                      }}
                    >
                      <td style={{ padding: '8px 10px', fontFamily: 'monospace', fontWeight: 600, color: 'var(--color-primary)' }}>
                        {s.key}
                      </td>
                      <td style={{ padding: '8px 10px', color: 'var(--fg-main)' }}>
                        {s.desc}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}

          {/* Tab 3: About */}
          {activeTab === 'about' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '16px', alignItems: 'center', textAlign: 'center', padding: '20px 0' }}>
              <div style={{ fontSize: '48px' }}>📊</div>
              <h3 style={{ margin: 0, fontSize: '20px', fontWeight: 600, color: 'var(--fg-main)' }}>
                {APP_INFO.name}
              </h3>
              <div style={{ fontSize: '13px', color: 'var(--fg-muted)' }}>
                {th.help.aboutVersion.replace('{version}', APP_INFO.version)}
              </div>
              <p style={{ margin: 0, maxWidth: '440px', fontSize: '14px', color: 'var(--fg-main)', lineHeight: 1.6 }}>
                {th.help.aboutDesc}
              </p>
              <div style={{ fontSize: '12px', color: 'var(--fg-muted)' }}>
                {th.help.aboutLicense}
              </div>

              <div style={{ marginTop: '12px' }}>
                <button
                  className="primary"
                  onClick={() => {
                    onClose();
                    onOpenWelcomeGuide();
                  }}
                >
                  🚀 {th.help.reopenWelcome}
                </button>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
};
