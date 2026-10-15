# Spike: กราฟที่ติดไปกับไฟล์ .xlsx

งานทดลองเพื่อประเมินความเป็นไปได้ ไม่ใช่โค้ดของแอป (ไม่ได้แก้ `src/`, `electron/`)

## วิธีรัน

```
node scripts/chart-spike/run.mjs [ไฟล์กราฟจาก Excel/Google Sheets ...]
npx vite-node scripts/chart-spike/appRoundtrip.ts [ไฟล์ .xlsx ...]
```

- `run.mjs`: ขั้นที่ 1 (ตรวจ ExcelJS), ขั้นที่ 2 (สร้างไฟล์ + แทรกกราฟ + ตรวจโครงสร้าง + อ่าน/เขียนกลับด้วย ExcelJS)
- `appRoundtrip.ts`: ส่งไฟล์ผ่านเส้นทางเปิด/บันทึกของแอปจริง (`src/core/xlsx.ts`) และดูว่าแถบเตือน "ส่วนที่ไม่รองรับ" ขึ้นไหม
- `chartXml.mjs`, `injectChart.mjs`: ตัวสร้าง XML กราฟ/drawing และตัวแทรกลง zip

## ไฟล์ผลลัพธ์ (`out/`)

| ไฟล์ | คืออะไร |
|---|---|
| `chart-bar.xlsx` | ข้อมูล 4 หมวด + กราฟแท่ง (ใส่กราฟเอง) |
| `chart-line.xlsx` | ข้อมูลเดียวกัน + กราฟเส้น |
| `chart-pie.xlsx` | ข้อมูลเดียวกัน + กราฟวงกลม |
| `*.exceljs-roundtrip.xlsx` | ไฟล์ข้างบนหลัง ExcelJS `load` → `writeFile` |
| `*.app-saved.xlsx` | ไฟล์ข้างบนหลังเปิด/บันทึกด้วยโค้ดของแอป |

ข้อมูล: ชีต `ข้อมูล` A1:B5 (หมวด / จำนวนเงิน: อาหาร 4500, เดินทาง 1200, ที่พัก 8000, อื่นๆ 650)
กราฟอ้าง `'ข้อมูล'!$A$2:$A$5` และ `'ข้อมูล'!$B$2:$B$5` พร้อมค่า cache วางที่ D2:K18 โดยประมาณ

## ผลที่ได้

1. **ExcelJS 4.4.0 เขียนกราฟไม่ได้:**
   - ไม่มี `addChart`
   - `index.d.ts` ไม่มีคำว่า chart เลย
   - `lib/xlsx/xlsx.js` ไม่จัดการ `xl/charts/`
2. **ExcelJS ไม่คงกราฟเมื่ออ่านแล้วเขียนใหม่:** กราฟหายครบทั้ง 3 ไฟล์ ทั้ง chart, drawing และแท็ก `<drawing>` ในชีต
   - โค้ด: `lib/xlsx/xform/drawing/two-cell-anchor-xform.js` รู้จักแค่ `xdr:pic` ไม่รู้จัก `xdr:graphicFrame`
   - `lib/xlsx/xform/sheet/worksheet-xform.js` (ช่วงบรรทัด 486–505) เก็บไว้เฉพาะ anchor ที่เป็นรูปภาพ
3. **บั๊กเดิมของแอป:** แถบเตือน "ไฟล์นี้มีบางส่วนที่แอปไม่รองรับ" ไม่ขึ้นกับไฟล์ที่มีกราฟ (ผลคือ `[]`)
   - สาเหตุ: `detectUnsupportedFeatures` ดู `ws.model.drawing` แต่ ExcelJS ทิ้งกราฟไปก่อนแล้ว
   - ผลคือผู้ใช้กด Ctrl+S แล้วกราฟหายโดยไม่มีคำเตือน ซึ่งขัดกับ SPEC 6.4
4. **แทรกกราฟเองหลัง ExcelJS เขียนไฟล์ทำได้:**
   - ไฟล์มี chart, drawing, rels, content type และ `<drawing r:id>` ในตำแหน่งที่ schema กำหนด
   - XML ทุกไฟล์ใน zip parse ผ่าน (well-formed)
   - **ยังไม่ได้เปิดใน Excel/Google Sheets จริง** (เครื่องนี้ไม่มี Excel หรือ LibreOffice)
5. **ไลบรารี zip:**
   - ใช้ `jszip` 3.10.2 ที่ติดมากับ exceljs (`exceljs → jszip ^3.10.1`) ไม่ได้ดาวน์โหลดอะไรเพิ่ม
   - ถ้าจะใช้ในแอปจริง ควรประกาศ `jszip` ใน `package.json` เป็น dependency ตรง (เวอร์ชันเดิม ไม่มีแพ็กเกจใหม่) เพราะการ import dependency ทางอ้อมอาจพังเมื่อ exceljs เปลี่ยน
   - `saxes` ที่ใช้ตรวจ XML ในสคริปต์ก็ติดมากับ exceljs เช่นกัน
