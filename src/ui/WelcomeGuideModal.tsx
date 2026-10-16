import React, { useState } from 'react';
import { th } from '../i18n/th';

interface WelcomeGuideModalProps {
  isOpen: boolean;
  onFinish: (dontShowAgain: boolean) => void;
  onSkip: (dontShowAgain: boolean) => void;
}

export const WelcomeGuideModal: React.FC<WelcomeGuideModalProps> = ({
  isOpen,
  onFinish,
  onSkip,
}) => {
  const [step, setStep] = useState(0);
  const [dontShowAgain, setDontShowAgain] = useState(false);

  if (!isOpen) return null;

  const g = th.welcomeGuide;
  const STEPS = [
    { title: g.step1Title, desc: g.step1Desc, tip: g.step1Tip, badge: g.step1Badge },
    { title: g.step2Title, desc: g.step2Desc, tip: g.step2Tip, badge: g.step2Badge },
    { title: g.step3Title, desc: g.step3Desc, tip: g.step3Tip, badge: g.step3Badge },
    { title: g.step4Title, desc: g.step4Desc, tip: g.step4Tip, badge: g.step4Badge },
  ];

  const currentStepData = STEPS[step];

  const handleNext = () => {
    if (step < STEPS.length - 1) {
      setStep(step + 1);
    } else {
      onFinish(dontShowAgain);
    }
  };

  const handlePrev = () => {
    if (step > 0) {
      setStep(step - 1);
    }
  };

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
          borderRadius: '12px',
          padding: '28px 32px',
          width: '560px',
          maxWidth: '92%',
          boxShadow: 'var(--shadow-dialog)',
          display: 'flex',
          flexDirection: 'column',
          gap: '20px',
        }}
      >
        {/* Header */}
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <div>
            <h2 style={{ margin: 0, fontSize: '18px', fontWeight: 600, color: 'var(--color-primary)' }}>
              {th.welcomeGuide.title}
            </h2>
            <div style={{ fontSize: '12px', color: 'var(--fg-muted)', marginTop: '2px' }}>
              {th.welcomeGuide.subtitle} ({step + 1}/{STEPS.length})
            </div>
          </div>
          <button
            onClick={() => onSkip(dontShowAgain)}
            style={{
              background: 'transparent',
              border: 'none',
              fontSize: '13px',
              color: 'var(--fg-muted)',
              cursor: 'pointer',
              padding: '4px 8px',
            }}
          >
            {th.welcomeGuide.skip}
          </button>
        </div>

        {/* Step Content Card */}
        <div
          style={{
            backgroundColor: 'var(--bg-app)',
            border: '1px solid var(--border-color)',
            borderRadius: '8px',
            padding: '24px 20px',
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            textAlign: 'center',
            gap: '12px',
            minHeight: '210px',
            justifyContent: 'center',
          }}
        >
          <div
            style={{
              fontSize: '11px',
              fontWeight: 600,
              color: 'var(--color-primary)',
              textTransform: 'uppercase',
              letterSpacing: '0.5px',
            }}
          >
            {currentStepData.badge}
          </div>
          <h3 style={{ margin: 0, fontSize: '17px', fontWeight: 600, color: 'var(--fg-main)' }}>
            {currentStepData.title}
          </h3>
          <p style={{ margin: 0, fontSize: '13px', color: 'var(--fg-muted)', lineHeight: 1.6, maxWidth: '440px' }}>
            {currentStepData.desc}
          </p>
          <div
            style={{
              marginTop: '8px',
              backgroundColor: 'var(--bg-selection)',
              border: '1px dashed var(--color-primary)',
              borderRadius: '6px',
              padding: '8px 14px',
              fontSize: '12px',
              color: 'var(--fg-main)',
              fontWeight: 500,
            }}
          >
            {currentStepData.tip}
          </div>
        </div>

        {/* Step Indicators */}
        <div style={{ display: 'flex', justifyContent: 'center', gap: '8px' }}>
          {STEPS.map((_, i) => (
            <div
              key={i}
              onClick={() => setStep(i)}
              style={{
                width: i === step ? '20px' : '8px',
                height: '8px',
                borderRadius: '4px',
                backgroundColor: i === step ? 'var(--color-primary)' : 'var(--border-color)',
                cursor: 'pointer',
                transition: 'all 0.2s',
              }}
            />
          ))}
        </div>

        {/* Footer actions */}
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: '4px' }}>
          <label
            style={{
              display: 'flex',
              alignItems: 'center',
              gap: '6px',
              fontSize: '12px',
              color: 'var(--fg-muted)',
              cursor: 'pointer',
              userSelect: 'none',
            }}
          >
            <input
              type="checkbox"
              checked={dontShowAgain}
              onChange={(e) => setDontShowAgain(e.target.checked)}
              style={{ cursor: 'pointer' }}
            />
            {th.welcomeGuide.dontShowAgain}
          </label>

          <div style={{ display: 'flex', gap: '8px' }}>
            <button onClick={handlePrev} disabled={step === 0}>
              {th.welcomeGuide.prev}
            </button>
            <button className="primary" onClick={handleNext}>
              {step === STEPS.length - 1 ? th.welcomeGuide.finish : th.welcomeGuide.next}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};
