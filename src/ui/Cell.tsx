import React from 'react';
import { CellStyle, CellError } from '../core/model';
import { ERROR_TOOLTIP_MAP } from '../core/numfmt';
import { defaultTextColorOnFill } from '../core/theme';
import { th } from '../i18n/th';
import { IconChevronDown, IconFilter } from './icons';

interface CellProps {
  r: number;
  c: number;
  a1: string;
  displayValue: string;
  rawValue: string;
  computedValue: any;
  isSelected: boolean;
  isActive: boolean;
  isCut?: boolean;
  isUnsupportedFormula?: boolean;
  highlightColor?: string;
  searchHighlight?: 'match' | 'current'; // ผลค้นหา (4C)
  // ปุ่มลูกศรตัวกรองที่เซลหัวตาราง (4C): active = คอลัมน์นี้กำลังกรองอยู่
  filterButton?: { active: boolean; title: string; onOpen: (x: number, y: number) => void };
  style?: CellStyle;
  rowSpan?: number;
  colSpan?: number;
  width: number;
  height: number;
  onMouseDown: (e: React.MouseEvent) => void;
  onMouseEnter: (e: React.MouseEvent) => void;
  onDoubleClick: () => void;
  onContextMenu: (e: React.MouseEvent) => void;
}

export const Cell: React.FC<CellProps> = ({
  displayValue,
  computedValue,
  isSelected,
  isActive,
  isCut,
  isUnsupportedFormula,
  highlightColor,
  searchHighlight,
  filterButton,
  style,
  rowSpan,
  colSpan,
  width,
  height,
  onMouseDown,
  onMouseEnter,
  onDoubleClick,
  onContextMenu,
}) => {
  const isNumber = typeof computedValue === 'number';
  const isError =
    typeof computedValue === 'string' &&
    ['#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#CYCLE!', '#ERR'].includes(computedValue);

  let tooltip = isError ? ERROR_TOOLTIP_MAP[computedValue as CellError] : undefined;
  if (isUnsupportedFormula) {
    tooltip = th.errors.unsupportedFormulaTooltip;
  }

  let textAlign = style?.align;
  if (!textAlign) {
    textAlign = isNumber ? 'right' : 'left';
  }

  // Determine background color
  let bgColor = style?.fill || 'transparent';
  if (highlightColor) {
    bgColor = `var(--${highlightColor}-bg)`; // พื้นอ่อนของสี reference
  } else if (searchHighlight === 'current') {
    bgColor = 'var(--search-current-bg)'; // ผลค้นหาปัจจุบัน (สีเด่น)
  } else if (searchHighlight === 'match') {
    bgColor = 'var(--search-match-bg)'; // ผลค้นหาอื่น (สีอ่อน)
  } else if (isActive) {
    bgColor = 'var(--grid-active-bg)';
  } else if (isSelected) {
    bgColor = 'var(--grid-selection)';
  }

  // สีตัวอักษร: สีที่ผู้ใช้ตั้งเสมอ, ไม่ตั้งแต่มีสีพื้น (และกำลังแสดงสีพื้นนั้น) ใช้สีธีมที่อ่านบนพื้นนั้นได้ชัด
  const textColor =
    style?.color ||
    (style?.fill && bgColor === style.fill ? defaultTextColorOnFill(style.fill) : undefined) ||
    'var(--grid-fg)';

  // Determine outline border
  let outline = 'none';
  if (highlightColor) {
    outline = `2px solid var(--${highlightColor})`;
  } else if (isActive) {
    outline = '2px solid var(--grid-active-border)';
  } else if (isCut) {
    outline = '2px dashed var(--grid-active-border)';
  }

  return (
    <td
      rowSpan={rowSpan}
      colSpan={colSpan}
      title={tooltip}
      onMouseDown={onMouseDown}
      onMouseEnter={onMouseEnter}
      onDoubleClick={onDoubleClick}
      onContextMenu={onContextMenu}
      style={{
        width: `${width}px`,
        minWidth: `${width}px`,
        maxWidth: `${width}px`,
        height: `${height}px`,
        padding: '0 6px',
        border: style?.border && style.border !== 'none' ? '1.5px solid var(--fg-main)' : '1px solid var(--grid-line)',
        backgroundColor: bgColor,
        color: isError ? 'var(--grid-error)' : textColor,
        fontWeight: style?.bold ? 'bold' : 'normal',
        fontStyle: style?.italic ? 'italic' : 'normal',
        textDecoration: style?.underline ? 'underline' : 'none',
        fontSize: style?.fontSize ? `${style.fontSize}pt` : undefined,
        fontFamily: style?.fontFamily || undefined,
        textAlign,
        whiteSpace: style?.wrap ? 'normal' : 'nowrap',
        wordBreak: style?.wrap ? 'break-word' : 'normal',
        lineHeight: style?.wrap ? 1.3 : undefined,
        overflow: 'hidden',
        textOverflow: 'ellipsis',
        userSelect: 'none',
        cursor: 'cell',
        outline,
        outlineOffset: '-2px',
        position: 'relative',
        zIndex: isActive || highlightColor ? 2 : 1,
      }}
    >
      {isUnsupportedFormula && (
        // สามเหลี่ยมเล็กที่มุมขวาบนของเซล (สีเตือนเดิม --grid-unsupported)
        <span
          aria-hidden="true"
          style={{
            position: 'absolute',
            top: 0,
            right: 0,
            width: 0,
            height: 0,
            borderTop: '7px solid var(--grid-unsupported)',
            borderLeft: '7px solid transparent',
            pointerEvents: 'none',
          }}
        />
      )}
      {displayValue}
      {filterButton && (
        <button
          type="button"
          title={filterButton.title}
          aria-pressed={filterButton.active}
          onMouseDown={(e) => e.stopPropagation()}
          onDoubleClick={(e) => e.stopPropagation()}
          onClick={(e) => {
            e.stopPropagation();
            const rect = e.currentTarget.getBoundingClientRect();
            filterButton.onOpen(rect.left, rect.bottom + 2);
          }}
          style={{
            position: 'absolute',
            right: '2px',
            top: '50%',
            transform: 'translateY(-50%)',
            width: '18px',
            height: '18px',
            padding: 0,
            lineHeight: '16px',
            fontSize: '10px',
            border: '1px solid var(--control-border)',
            borderRadius: '3px',
            cursor: 'pointer',
            // คอลัมน์ที่กรองอยู่: ปุ่มสีเด่น + สัญลักษณ์กรวย
            backgroundColor: filterButton.active ? 'var(--color-primary)' : 'var(--grid-bg)',
            color: filterButton.active ? 'var(--fg-on-primary)' : 'var(--grid-fg)',
            zIndex: 3,
          }}
        >
          {filterButton.active ? <IconFilter size={12} /> : <IconChevronDown size={12} />}
        </button>
      )}
    </td>
  );
};
