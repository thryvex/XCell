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
  style?: CellStyle;
  rowSpan?: number;
  colSpan?: number;
  width: number;
  height: number;
  onMouseDown: (e: React.MouseEvent) => void;
  onMouseEnter: (e: React.MouseEvent) => void;
  onDoubleClick: () => void;
}

export const Cell: React.FC<CellProps> = ({
  displayValue,
  computedValue,
  isSelected,
  isActive,
  style,
  rowSpan,
  colSpan,
  width,
  height,
  onMouseDown,
  onMouseEnter,
  onDoubleClick,
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

  return (
    <td
      rowSpan={rowSpan}
      colSpan={colSpan}
      title={tooltip}
      onMouseDown={onMouseDown}
      onMouseEnter={onMouseEnter}
      onDoubleClick={onDoubleClick}
      style={{
        width: `${width}px`,
        minWidth: `${width}px`,
        maxWidth: `${width}px`,
        height: `${height}px`,
        padding: '0 6px',
        border: '1px solid var(--border-grid)',
        backgroundColor: isActive
          ? 'var(--bg-active)'
          : isSelected
          ? 'var(--bg-selection)'
          : style?.fill || 'transparent',
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
        outline: isActive ? '2px solid var(--color-primary)' : 'none',
        outlineOffset: '-2px',
        position: 'relative',
        zIndex: isActive ? 2 : 1,
      }}
    >
      {displayValue}
    </td>
  );
};
