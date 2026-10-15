import React from 'react';
import { CellStyle, CellError } from '../core/model';
import { ERROR_TOOLTIP_MAP } from '../core/numfmt';

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
    tooltip = 'ฟังก์ชันนี้แอปยังไม่รองรับ แสดงค่าเดิมจากไฟล์';
  }

  let textAlign = style?.align;
  if (!textAlign) {
    textAlign = isNumber ? 'right' : 'left';
  }

  // Determine background color
  let bgColor = style?.fill || 'transparent';
  if (highlightColor) {
    bgColor = `${highlightColor}18`; // light tint
  } else if (searchHighlight === 'current') {
    bgColor = '#fdba74'; // ผลค้นหาปัจจุบัน (สีเด่น)
  } else if (searchHighlight === 'match') {
    bgColor = '#fef08a'; // ผลค้นหาอื่น (สีอ่อน)
  } else if (isActive) {
    bgColor = 'var(--bg-active)';
  } else if (isSelected) {
    bgColor = 'var(--bg-selection)';
  }

  // Determine outline border
  let outline = 'none';
  if (highlightColor) {
    outline = `2px solid ${highlightColor}`;
  } else if (isActive) {
    outline = '2px solid var(--color-primary)';
  } else if (isCut) {
    outline = '2px dashed var(--color-primary)';
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
        border: style?.border && style.border !== 'none' ? '1.5px solid var(--fg-main)' : '1px solid var(--border-grid)',
        backgroundColor: bgColor,
        color: isError ? 'var(--color-danger)' : style?.color || 'var(--fg-main)',
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
        <span
          style={{
            position: 'absolute',
            top: '2px',
            right: '2px',
            fontSize: '10px',
            lineHeight: 1,
            color: '#d97706',
            pointerEvents: 'none',
          }}
          title={tooltip}
        >
          ⚠️
        </span>
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
            border: '1px solid var(--border-color)',
            borderRadius: '3px',
            cursor: 'pointer',
            // คอลัมน์ที่กรองอยู่: ปุ่มสีเด่น + สัญลักษณ์กรวย
            backgroundColor: filterButton.active ? 'var(--color-primary)' : 'var(--bg-app)',
            color: filterButton.active ? '#ffffff' : 'var(--fg-main)',
            zIndex: 3,
          }}
        >
          {filterButton.active ? '▼' : '▾'}
        </button>
      )}
    </td>
  );
};
