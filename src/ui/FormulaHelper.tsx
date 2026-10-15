import React, { useState, useMemo } from 'react';
import { th } from '../i18n/th';
import { useSheetStore } from './store';

export interface FunctionItem {
  name: string;
  thaiName: string;
  category: 'math' | 'stat' | 'logic' | 'date' | 'text';
  description: string;
  syntax: string;
  example: string;
}

export const CORE_FUNCTIONS: FunctionItem[] = [
  // Math functions
  {
    name: 'SUM',
    thaiName: 'ผลรวม',
    category: 'math',
    description: 'หาผลรวมของตัวเลขทั้งหมดในช่วงเซล',
    syntax: 'SUM(ช่วงข้อมูล)',
    example: '=SUM(A1:A10)',
  },
  {
    name: 'PRODUCT',
    thaiName: 'ผลคูณ',
    category: 'math',
    description: 'หาผลคูณของตัวเลขทั้งหมดในช่วงเซล',
    syntax: 'PRODUCT(ช่วงข้อมูล)',
    example: '=PRODUCT(A1:A5)',
  },
  {
    name: 'ROUND',
    thaiName: 'ปัดเศษทศนิยม',
    category: 'math',
    description: 'ปัดเศษตัวเลขตามจำนวนหลักทศนิยมที่กำหนด',
    syntax: 'ROUND(ตัวเลข, จำนวนหลัก)',
    example: '=ROUND(A1, 2)',
  },
  {
    name: 'ROUNDUP',
    thaiName: 'ปัดเศษขึ้น',
    category: 'math',
    description: 'ปัดเศษตัวเลขขึ้นเสมอตามจำนวนหลักที่กำหนด',
    syntax: 'ROUNDUP(ตัวเลข, จำนวนหลัก)',
    example: '=ROUNDUP(A1, 0)',
  },
  {
    name: 'ROUNDDOWN',
    thaiName: 'ปัดเศษลง',
    category: 'math',
    description: 'ปัดเศษตัวเลขลงเสมอตามจำนวนหลักที่กำหนด',
    syntax: 'ROUNDDOWN(ตัวเลข, จำนวนหลัก)',
    example: '=ROUNDDOWN(A1, 0)',
  },
  {
    name: 'ABS',
    thaiName: 'ค่าสัมบูรณ์',
    category: 'math',
    description: 'แปลงค่าตัวเลขให้เป็นค่าบวกเสมอ',
    syntax: 'ABS(ตัวเลข)',
    example: '=ABS(A1)',
  },
  {
    name: 'INT',
    thaiName: 'ปัดลงเป็นจำนวนเต็ม',
    category: 'math',
    description: 'ปัดเศษตัวเลขลงเป็นจำนวนเต็มที่ใกล้ที่สุด',
    syntax: 'INT(ตัวเลข)',
    example: '=INT(A1)',
  },
  {
    name: 'SQRT',
    thaiName: 'สแควร์รูท (รากที่สอง)',
    category: 'math',
    description: 'หาค่ารากที่สองของตัวเลขบวก',
    syntax: 'SQRT(ตัวเลข)',
    example: '=SQRT(A1)',
  },
  {
    name: 'POWER',
    thaiName: 'เลขยกกำลัง',
    category: 'math',
    description: 'คำนวณผลลัพธ์ของตัวเลขยกกำลัง',
    syntax: 'POWER(ฐาน, เลขชี้กำลัง)',
    example: '=POWER(A1, 2)',
  },
  {
    name: 'MOD',
    thaiName: 'เศษจากการหาร',
    category: 'math',
    description: 'หาเศษที่เหลือจากการหารตัวเลข',
    syntax: 'MOD(ตัวตั้ง, ตัวหาร)',
    example: '=MOD(A1, 3)',
  },

  // Statistical functions
  {
    name: 'AVERAGE',
    thaiName: 'ค่าเฉลี่ย',
    category: 'stat',
    description: 'หาค่าเฉลี่ยเลขคณิตของตัวเลขทั้งหมดในช่วงเซล',
    syntax: 'AVERAGE(ช่วงข้อมูล)',
    example: '=AVERAGE(B2:B20)',
  },
  {
    name: 'MEDIAN',
    thaiName: 'มัธยฐาน',
    category: 'stat',
    description: 'หาค่ากึ่งกลางของชุดข้อมูลตัวเลข',
    syntax: 'MEDIAN(ช่วงข้อมูล)',
    example: '=MEDIAN(C1:C15)',
  },
  {
    name: 'MIN',
    thaiName: 'ค่าน้อยสุด',
    category: 'stat',
    description: 'หาค่าตัวเลขที่น้อยที่สุดในช่วงเซล',
    syntax: 'MIN(ช่วงข้อมูล)',
    example: '=MIN(D1:D10)',
  },
  {
    name: 'MAX',
    thaiName: 'ค่ามากสุด',
    category: 'stat',
    description: 'หาค่าตัวเลขที่มากที่สุดในช่วงเซล',
    syntax: 'MAX(ช่วงข้อมูล)',
    example: '=MAX(D1:D10)',
  },
  {
    name: 'COUNT',
    thaiName: 'นับจำนวนตัวเลข',
    category: 'stat',
    description: 'นับจำนวนเซลที่มีค่าเป็นตัวเลข',
    syntax: 'COUNT(ช่วงข้อมูล)',
    example: '=COUNT(A1:A50)',
  },
  {
    name: 'COUNTA',
    thaiName: 'นับจำนวนเซลไม่ว่าง',
    category: 'stat',
    description: 'นับจำนวนเซลทั้งหมดที่มีข้อมูล (ไม่ว่าง)',
    syntax: 'COUNTA(ช่วงข้อมูล)',
    example: '=COUNTA(A1:A50)',
  },

  // Logical functions
  {
    name: 'IF',
    thaiName: 'ถ้า (เงื่อนไข)',
    category: 'logic',
    description: 'ตรวจสอบเงื่อนไขแล้วส่งคืนค่าตามผลจริงหรือเท็จ',
    syntax: 'IF(เงื่อนไข, ค่าเมื่อจริง, ค่าเมื่อเท็จ)',
    example: '=IF(A1>=50, "ผ่าน", "ตก")',
  },
  {
    name: 'IFERROR',
    thaiName: 'ถ้าเกิดข้อผิดพลาด',
    category: 'logic',
    description: 'ส่งคืนค่าสำรองหากสูตรเกิดข้อผิดพลาด',
    syntax: 'IFERROR(สูตร, ค่าสำรองเมื่อเกิดข้อผิดพลาด)',
    example: '=IFERROR(A1/B1, 0)',
  },
  {
    name: 'AND',
    thaiName: 'และ (จริงทุกข้อ)',
    category: 'logic',
    description: 'ตรวจสอบว่าเงื่อนไขทุกข้อเป็นจริงทั้งหมดหรือไม่',
    syntax: 'AND(เงื่อนไข1, เงื่อนไข2, ...)',
    example: '=AND(A1>0, B1>0)',
  },
  {
    name: 'OR',
    thaiName: 'หรือ (จริงข้อใดข้อหนึ่ง)',
    category: 'logic',
    description: 'ตรวจสอบว่ามีเงื่อนไขใดข้อหนึ่งเป็นจริงหรือไม่',
    syntax: 'OR(เงื่อนไข1, เงื่อนไข2, ...)',
    example: '=OR(A1>50, B1>50)',
  },
  {
    name: 'NOT',
    thaiName: 'ไม่ (สลับค่าความจริง)',
    category: 'logic',
    description: 'สลับค่าความจริงจากจริงเป็นเท็จ และเท็จเป็นจริง',
    syntax: 'NOT(เงื่อนไข)',
    example: '=NOT(A1=0)',
  },

  // Phase 4A: Conditional Aggregations
  {
    name: 'SUMIF',
    thaiName: 'รวมตามเงื่อนไข',
    category: 'math',
    description: 'หาผลรวมของตัวเลขในช่วงที่ตรงตามเงื่อนไขที่กำหนด',
    syntax: 'SUMIF(ช่วงเงื่อนไข, เงื่อนไข, [ช่วงรวม])',
    example: '=SUMIF(A1:A10, ">100", B1:B10)',
  },
  {
    name: 'COUNTIF',
    thaiName: 'นับตามเงื่อนไข',
    category: 'stat',
    description: 'นับจำนวนเซลในช่วงที่ตรงตามเงื่อนไขที่กำหนด',
    syntax: 'COUNTIF(ช่วงเงื่อนไข, เงื่อนไข)',
    example: '=COUNTIF(A1:A10, "อาหาร*")',
  },
  {
    name: 'AVERAGEIF',
    thaiName: 'เฉลี่ยตามเงื่อนไข',
    category: 'stat',
    description: 'หาค่าเฉลี่ยของตัวเลขในช่วงที่ตรงตามเงื่อนไข',
    syntax: 'AVERAGEIF(ช่วงเงื่อนไข, เงื่อนไข, [ช่วงเฉลี่ย])',
    example: '=AVERAGEIF(A1:A10, "<>0", B1:B10)',
  },
  {
    name: 'SUMIFS',
    thaiName: 'รวมหลายเงื่อนไข',
    category: 'math',
    description: 'หาผลรวมของตัวเลขที่ตรงตามหลายเงื่อนไขพร้อมกัน',
    syntax: 'SUMIFS(ช่วงรวม, ช่วงเงื่อนไข1, เงื่อนไข1, [ช่วงเงื่อนไข2, เงื่อนไข2, ...])',
    example: '=SUMIFS(C1:C10, A1:A10, "รายรับ", B1:B10, ">500")',
  },
  {
    name: 'COUNTIFS',
    thaiName: 'นับหลายเงื่อนไข',
    category: 'stat',
    description: 'นับจำนวนแถวที่ตรงตามหลายเงื่อนไขพร้อมกัน',
    syntax: 'COUNTIFS(ช่วงเงื่อนไข1, เงื่อนไข1, [ช่วงเงื่อนไข2, เงื่อนไข2, ...])',
    example: '=COUNTIFS(A1:A10, "อาหาร*", B1:B10, ">100")',
  },

  // Phase 4A: Date & Time functions
  {
    name: 'TODAY',
    thaiName: 'วันที่ปัจจุบัน',
    category: 'date',
    description: 'ส่งกลับวันที่ปัจจุบันของระบบ (อัปเดตเมื่อคำนวณใหม่)',
    syntax: 'TODAY()',
    example: '=TODAY()',
  },
  {
    name: 'NOW',
    thaiName: 'วันและเวลาปัจจุบัน',
    category: 'date',
    description: 'ส่งกลับวันและเวลาปัจจุบันของระบบ',
    syntax: 'NOW()',
    example: '=NOW()',
  },
  {
    name: 'YEAR',
    thaiName: 'ปี',
    category: 'date',
    description: 'ดึงเลขปี (ค.ศ.) ออกจากวันที่',
    syntax: 'YEAR(วันที่)',
    example: '=YEAR(A1)',
  },
  {
    name: 'MONTH',
    thaiName: 'เดือน',
    category: 'date',
    description: 'ดึงเลขเดือน (1-12) ออกจากวันที่',
    syntax: 'MONTH(วันที่)',
    example: '=MONTH(A1)',
  },
  {
    name: 'DAY',
    thaiName: 'วัน',
    category: 'date',
    description: 'ดึงเลขวันในเดือน (1-31) ออกจากวันที่',
    syntax: 'DAY(วันที่)',
    example: '=DAY(A1)',
  },
  {
    name: 'DATE',
    thaiName: 'สร้างวันที่',
    category: 'date',
    description: 'สร้างวันที่จากปี เดือน และวันที่ระบุ (รองรับเดือน/วันล้น)',
    syntax: 'DATE(ปี, เดือน, วัน)',
    example: '=DATE(2026, 10, 1)',
  },

  // Phase 4A: Text functions
  {
    name: 'TEXT',
    thaiName: 'แปลงรูปแบบข้อความ',
    category: 'text',
    description: 'แปลงตัวเลขหรือวันที่ให้อยู่ในรูปแบบข้อความที่กำหนด',
    syntax: 'TEXT(ค่า, รูปแบบ)',
    example: '=TEXT(A1, "#,##0.00")',
  },
  {
    name: 'CONCAT',
    thaiName: 'ต่อข้อความ',
    category: 'text',
    description: 'รวมข้อความหลายค่าหรือช่วงเซลเข้าด้วยกัน',
    syntax: 'CONCAT(ข้อความ1, [ข้อความ2, ...])',
    example: '=CONCAT("สวัสดี", " ", A1)',
  },
  {
    name: 'LEN',
    thaiName: 'ความยาวข้อความ',
    category: 'text',
    description: 'นับจำนวนตัวอักษรในข้อความ',
    syntax: 'LEN(ข้อความ)',
    example: '=LEN(A1)',
  },
];

const OPERATORS = ['+', '-', '*', '/', '(', ')', '^', '%'];

export const FormulaHelper: React.FC = () => {
  const { isFormulaHelperOpen, closeFormulaHelper, insertFormulaFunction } = useSheetStore();
  const [searchTerm, setSearchTerm] = useState('');
  const [selectedCategory, setSelectedCategory] = useState<'all' | 'math' | 'stat' | 'logic' | 'date' | 'text'>('all');
  const [selectedFnName, setSelectedFnName] = useState<string>('SUM');

  const filteredFunctions = useMemo(() => {
    const term = searchTerm.trim().toLowerCase();
    return CORE_FUNCTIONS.filter((fn) => {
      const matchCategory = selectedCategory === 'all' || fn.category === selectedCategory;
      if (!matchCategory) return false;
      if (!term) return true;
      return (
        fn.name.toLowerCase().includes(term) ||
        fn.thaiName.toLowerCase().includes(term) ||
        fn.description.toLowerCase().includes(term) ||
        (term === 'mean' && fn.name === 'AVERAGE') ||
        (term === 'avg' && fn.name === 'AVERAGE')
      );
    });
  }, [searchTerm, selectedCategory]);

  const currentFn = useMemo(() => {
    return CORE_FUNCTIONS.find((fn) => fn.name === selectedFnName) || filteredFunctions[0] || CORE_FUNCTIONS[0];
  }, [selectedFnName, filteredFunctions]);

  if (!isFormulaHelperOpen) return null;

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
      onClick={closeFormulaHelper}
    >
      <div
        style={{
          backgroundColor: 'var(--bg-app)',
          border: '1px solid var(--border-color)',
          borderRadius: '10px',
          padding: '20px 24px',
          width: '640px',
          maxWidth: '92%',
          maxHeight: '85vh',
          boxShadow: 'var(--shadow-dialog)',
          display: 'flex',
          flexDirection: 'column',
          gap: '16px',
        }}
        onClick={(e) => e.stopPropagation()}
      >
        {/* Header */}
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            <span style={{ fontSize: '18px', fontWeight: 700, color: 'var(--color-primary)' }}>ƒx</span>
            <h3 style={{ margin: 0, fontSize: '17px', fontWeight: 600 }}>{th.formulaHelper.title}</h3>
          </div>
          <button
            onClick={closeFormulaHelper}
            style={{
              border: 'none',
              background: 'transparent',
              fontSize: '18px',
              cursor: 'pointer',
              color: 'var(--fg-muted)',
              padding: '2px 6px',
            }}
          >
            ✕
          </button>
        </div>

        {/* Search input & Category tabs */}
        <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
          <input
            type="text"
            autoFocus
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            placeholder={th.formulaHelper.searchPlaceholder}
            style={{
              padding: '8px 12px',
              borderRadius: '6px',
              border: '1px solid var(--control-border)',
              fontSize: '14px',
              width: '100%',
              boxSizing: 'border-box',
            }}
          />

          <div style={{ display: 'flex', gap: '6px' }}>
            {[
              { id: 'all', label: th.formulaHelper.categoryAll },
              { id: 'math', label: th.formulaHelper.categoryMath },
              { id: 'stat', label: th.formulaHelper.categoryStat },
              { id: 'logic', label: th.formulaHelper.categoryLogic },
              { id: 'date', label: th.formulaHelper.categoryDate },
              { id: 'text', label: th.formulaHelper.categoryText },
            ].map((cat) => (
              <button
                key={cat.id}
                onClick={() => setSelectedCategory(cat.id as any)}
                style={{
                  padding: '4px 10px',
                  fontSize: '12px',
                  borderRadius: '16px',
                  border: '1px solid var(--border-color)',
                  backgroundColor: selectedCategory === cat.id ? 'var(--color-primary)' : 'transparent',
                  color: selectedCategory === cat.id ? 'var(--fg-on-primary)' : 'var(--fg-main)',
                  fontWeight: selectedCategory === cat.id ? 600 : 400,
                  cursor: 'pointer',
                }}
              >
                {cat.label}
              </button>
            ))}
          </div>
        </div>

        {/* Content Area: List + Details */}
        <div style={{ display: 'flex', gap: '16px', minHeight: '260px', maxHeight: '340px' }}>
          {/* Function List */}
          <div
            style={{
              flex: '0 0 210px',
              border: '1px solid var(--border-color)',
              borderRadius: '6px',
              overflowY: 'auto',
              padding: '4px',
              display: 'flex',
              flexDirection: 'column',
              gap: '2px',
            }}
          >
            {filteredFunctions.length === 0 ? (
              <div style={{ padding: '16px', textAlign: 'center', color: 'var(--fg-muted)', fontSize: '13px' }}>
                {th.formulaHelper.noResults}
              </div>
            ) : (
              filteredFunctions.map((fn) => {
                const isSelected = currentFn?.name === fn.name;
                return (
                  <div
                    key={fn.name}
                    onClick={() => setSelectedFnName(fn.name)}
                    style={{
                      padding: '6px 10px',
                      borderRadius: '4px',
                      cursor: 'pointer',
                      backgroundColor: isSelected ? 'var(--bg-selection)' : 'transparent',
                      color: isSelected ? 'var(--color-primary)' : 'var(--fg-main)',
                      fontWeight: isSelected ? 600 : 400,
                      display: 'flex',
                      justifyContent: 'space-between',
                      alignItems: 'center',
                    }}
                  >
                    <span>{fn.name}</span>
                    <span style={{ fontSize: '11px', color: 'var(--fg-muted)' }}>{fn.thaiName}</span>
                  </div>
                );
              })
            )}
          </div>

          {/* Details Panel */}
          {currentFn && (
            <div
              style={{
                flex: 1,
                border: '1px solid var(--border-color)',
                borderRadius: '6px',
                padding: '14px 16px',
                display: 'flex',
                flexDirection: 'column',
                justifyContent: 'space-between',
                backgroundColor: 'var(--bg-surface)',
              }}
            >
              <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                <div style={{ display: 'flex', alignItems: 'baseline', gap: '8px' }}>
                  <span style={{ fontSize: '18px', fontWeight: 700, color: 'var(--color-primary)' }}>
                    {currentFn.name}
                  </span>
                  <span style={{ fontSize: '14px', color: 'var(--fg-muted)', fontWeight: 500 }}>
                    ({currentFn.thaiName})
                  </span>
                </div>

                <p style={{ margin: 0, fontSize: '13px', lineHeight: 1.5, color: 'var(--fg-main)' }}>
                  {currentFn.description}
                </p>

                <div style={{ marginTop: '4px' }}>
                  <div style={{ fontSize: '12px', fontWeight: 600, color: 'var(--fg-muted)', marginBottom: '3px' }}>
                    {th.formulaHelper.syntaxTitle}
                  </div>
                  <code
                    style={{
                      display: 'block',
                      backgroundColor: 'var(--bg-app)',
                      border: '1px solid var(--border-color)',
                      borderRadius: '4px',
                      padding: '6px 10px',
                      fontFamily: 'monospace',
                      fontSize: '12.5px',
                    }}
                  >
                    ={currentFn.syntax}
                  </code>
                </div>

                <div>
                  <div style={{ fontSize: '12px', fontWeight: 600, color: 'var(--fg-muted)', marginBottom: '3px' }}>
                    {th.formulaHelper.examplesTitle}
                  </div>
                  <code
                    style={{
                      display: 'block',
                      backgroundColor: 'var(--bg-app)',
                      border: '1px solid var(--border-color)',
                      borderRadius: '4px',
                      padding: '6px 10px',
                      fontFamily: 'monospace',
                      fontSize: '12.5px',
                      color: 'var(--color-primary)',
                    }}
                  >
                    {currentFn.example}
                  </code>
                </div>
              </div>

              <div style={{ marginTop: '16px', display: 'flex', justifyContent: 'flex-end' }}>
                <button
                  className="primary"
                  onClick={() => insertFormulaFunction(currentFn.name)}
                  style={{ padding: '8px 18px', fontSize: '13.5px', fontWeight: 600 }}
                >
                  ✓ {th.formulaHelper.useFormula}
                </button>
              </div>
            </div>
          )}
        </div>

        {/* Quick Operators Panel */}
        <div
          style={{
            display: 'flex',
            alignItems: 'center',
            gap: '8px',
            paddingTop: '8px',
            borderTop: '1px solid var(--border-color)',
          }}
        >
          <span style={{ fontSize: '12px', color: 'var(--fg-muted)', fontWeight: 500 }}>
            {th.formulaHelper.operators}
          </span>
          <div style={{ display: 'flex', gap: '6px' }}>
            {OPERATORS.map((op) => (
              <button
                key={op}
                onClick={() => {
                  const state = useSheetStore.getState();
                  if (!state.isEditing) {
                    state.startEditing(`=${op}`);
                  } else {
                    state.setEditingValue(state.editingValue + op);
                  }
                  closeFormulaHelper();
                }}
                style={{
                  width: '28px',
                  height: '28px',
                  padding: 0,
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  fontSize: '14px',
                  fontWeight: 600,
                  cursor: 'pointer',
                  borderRadius: '4px',
                }}
              >
                {op}
              </button>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
};
