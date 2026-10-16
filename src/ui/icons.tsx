import React from 'react';

/**
 * ไอคอนเส้นชุดเดียวของแอป (แนวเดียวกับ Google Sheets): viewBox 24, เส้น 2, ปลายมน, สีตาม currentColor
 * ใช้เฉพาะปุ่มที่ไม่มีข้อความ (toolbar, ปิด, ลูกศร) — ปุ่มที่มีข้อความใช้ข้อความล้วน
 */
type IconProps = { size?: number; title?: string };

const Svg: React.FC<IconProps & { children: React.ReactNode }> = ({ size = 20, title, children }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth={2}
    strokeLinecap="round"
    strokeLinejoin="round"
    aria-hidden={title ? undefined : true}
    role={title ? 'img' : undefined}
    focusable="false"
    style={{ display: 'block', flexShrink: 0 }}
  >
    {title && <title>{title}</title>}
    {children}
  </svg>
);

export const IconAlignLeft: React.FC<IconProps> = (p) => (
  <Svg {...p}>
    <path d="M4 6h16M4 10h10M4 14h16M4 18h10" />
  </Svg>
);

export const IconAlignCenter: React.FC<IconProps> = (p) => (
  <Svg {...p}>
    <path d="M4 6h16M7 10h10M4 14h16M7 18h10" />
  </Svg>
);

export const IconAlignRight: React.FC<IconProps> = (p) => (
  <Svg {...p}>
    <path d="M4 6h16M10 10h10M4 14h16M10 18h10" />
  </Svg>
);

/** สีพื้นเซล (ถังสี) */
export const IconFillColor: React.FC<IconProps> = (p) => (
  <Svg {...p}>
    <path d="M5 11l6-6 7 7-6 6a1.4 1.4 0 0 1-2 0l-5-5a1.4 1.4 0 0 1 0-2z" />
    <path d="M9 3l2 2M5 12h13" />
    <path d="M20 15s1.5 2 1.5 3a1.5 1.5 0 0 1-3 0c0-1 1.5-3 1.5-3z" />
  </Svg>
);

/** ลูกศรเมนูดรอปดาวน์ */
export const IconChevronDown: React.FC<IconProps> = ({ size = 16, ...p }) => (
  <Svg size={size} {...p}>
    <path d="M7 10l5 5 5-5" />
  </Svg>
);

export const IconChevronUp: React.FC<IconProps> = ({ size = 16, ...p }) => (
  <Svg size={size} {...p}>
    <path d="M7 14l5-5 5 5" />
  </Svg>
);

export const IconClose: React.FC<IconProps> = ({ size = 18, ...p }) => (
  <Svg size={size} {...p}>
    <path d="M6 6l12 12M18 6L6 18" />
  </Svg>
);

export const IconPlus: React.FC<IconProps> = ({ size = 18, ...p }) => (
  <Svg size={size} {...p}>
    <path d="M12 5v14M5 12h14" />
  </Svg>
);

export const IconMinus: React.FC<IconProps> = ({ size = 18, ...p }) => (
  <Svg size={size} {...p}>
    <path d="M5 12h14" />
  </Svg>
);

export const IconCheck: React.FC<IconProps> = ({ size = 16, ...p }) => (
  <Svg size={size} {...p}>
    <path d="M5 12.5l4.5 4.5L19 7.5" />
  </Svg>
);

/** ตัวเลือก (ใช้ในแถบค้นหา) */
export const IconOptions: React.FC<IconProps> = ({ size = 18, ...p }) => (
  <Svg size={size} {...p}>
    <path d="M4 7h10M18 7h2M4 17h4M12 17h8" />
    <circle cx="16" cy="7" r="2" />
    <circle cx="10" cy="17" r="2" />
  </Svg>
);

/** ตัวกรองที่ใช้งานอยู่ (กรวย) */
export const IconFilter: React.FC<IconProps> = ({ size = 14, ...p }) => (
  <Svg size={size} {...p}>
    <path d="M4 5h16l-6 7.5V19l-4-2v-4.5z" />
  </Svg>
);
