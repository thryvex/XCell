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
  highlightColor?: string;
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
  highlightColor,
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

  const tooltip = isError ? ERROR_TOOLTIP_MAP[computedValue as CellError] : undefined;

  let textAlign = style?.align;
  if (!textAlign) {
    textAlign = isNumber ? 'right' : 'left';
  }

  // Determine background color
  let bgColor = style?.fill || 'transparent';
  if (highlightColor) {
    bgColor = `${highlightColor}18`; // light tint
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
        border: '1px solid var(--border-grid)',
        backgroundColor: bgColor,
        color: isError ? 'var(--color-danger)' : style?.color || 'var(--fg-main)',
        fontWeight: style?.bold ? 'bold' : 'normal',
        fontStyle: style?.italic ? 'italic' : 'normal',
        textDecoration: style?.underline ? 'underline' : 'none',
        textAlign,
        whiteSpace: style?.wrap ? 'normal' : 'nowrap',
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
      {displayValue}
    </td>
  );
};
