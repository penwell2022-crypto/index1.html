/* =========================================================	# 🧭 แผนภาพธาตุ 4 – อาการแยกตามธาตุ
   แผนภาพธาตุ 4 – อาการแยกตามธาตุ	
   © 2026 [ชื่อผู้จัดทำ] สงวนลิขสิทธิ์ (All rights reserved)	© 2026 [ชื่อผู้จัดทำ] — สงวนลิขสิทธิ์ (All rights reserved)
   ห้ามคัดลอก ดัดแปลง หรือนำโค้ดไปใช้ต่อโดยไม่ได้รับอนุญาต	ห้ามคัดลอก ดัดแปลง หรือนำโค้ด/เนื้อหาไปใช้ต่อโดยไม่ได้รับอนุญาต · แชร์ได้เฉพาะลิงก์หน้าเว็บเพื่ออ่านทบทวน (ดูไฟล์ LICENSE)
   ข้อมูล: B=รหัส C=ชื่อ D=คำอธิบาย E=ระดับ F=ภาวะ G..Z=อาการ 1-20 (เริ่มแถว 3)	
   ========================================================= */	## เมนูบนหน้าเว็บ
	| เมนู | คำอธิบาย |
const DATA_URL = "data/data.xlsx";	|---|---|
const TABS = ["cards", "ladder", "map", "explore", "quiz"];	| ① การ์ดธาตุ | การ์ดสี 5 ธาตุ จำนวนภาวะ/อาการ คำช่วยจำ |
	| ② บันไดความรุนแรง | หย่อน → กำเริบ → พิการ → ออกจากกาย → แตก → ขาด |
// ธาตุหลัก (คำช่วยจำ = ตัวช่วยเรียนที่เพิ่มขึ้น ไม่ได้มาจากตำรา)	| ③ แผนที่กลุ่มธาตุ | ตาราง ● ธาตุที่เกี่ยวข้องในแต่ละกลุ่ม + กราฟแท่ง |
const ELEMENTS = [	| ④ สำรวจเอง | เลือกกลุ่ม / ค้นหาอาการ |
  { key: "ไฟ",   code: "A01", title: "🔥 ไฟ (เตโช)",   prefix: "เตโชธาตุ", c: "var(--fire)",  l: "var(--fire-l)",  mn: "ร้อน • แสบ • ไข้ • ดี/กำเดา (ระบบเผาผลาญ)" },	| ⑤ แบบทดสอบ | ก) เลือกตอบ 8 ข้อ · ข) ลากวางจับคู่ |
  { key: "ดิน",  code: "A02", title: "⛰️ ดิน (ปถวี)",  prefix: "ปถวีธาตุ", c: "var(--earth)", l: "var(--earth-l)", mn: "ของแข็งในกาย: ผม ฟัน กระดูก อวัยวะ ท้อง-ไส้" },	
  { key: "น้ำ",  code: "A03", title: "💧 น้ำ (อาโป)",  prefix: "อาโปธาตุ", c: "var(--water)", l: "var(--water-l)", mn: "ของเหลว: เสมหะ เลือด หนอง เหงื่อ ปัสสาวะ" },	## โครงสร้างไฟล์
  { key: "ลม",   code: "A04", title: "🌬️ ลม (วาโย)",  prefix: "วาโยธาตุ", c: "var(--wind)",  l: "var(--wind-l)",  mn: "การเคลื่อนไหว: หาวเรอ ผายลม เมื่อย วิงเวียน" },	```
  { key: "อากาศ", code: "A05", title: "🌌 อากาศ",       prefix: "อากาศธาตุ", c: "var(--space)", l: "var(--space-l)", mn: "ช่องว่างในกาย: หู ตา (ข้อมูลในตำราน้อย)" },	index.html        หน้าเว็บ
];	style.css         สี/รูปแบบ
const STAGES = ["หย่อน", "กำเริบ", "พิการ", "ออกจากกาย", "แตก", "ขาด"];	app.js            โค้ดอ่าน Excel และสร้างแผนภาพ
const STAGE_COLORS = ["#FFF2CC", "#FFE699", "#F8CBAD", "#F4B183", "#E06666", "#990000"];	data/data.xlsx    ไฟล์ข้อมูล (สำเนาที่มีเฉพาะชีตข้อมูล)
	LICENSE           ข้อกำหนดลิขสิทธิ์
let ROWS = [];	.nojekyll
	```
const $ = (id) => document.getElementById(id);	
const esc = (s) => String(s).replace(/[&<>"]/g, (m) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[m]));	## อัปเดตข้อมูล
const shuffle = (a) => { a = [...a]; for (let i = a.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); [a[i], a[j]] = [a[j], a[i]]; } return a; };	แก้ข้อมูลใน Excel → บันทึกสำเนาที่มีเฉพาะชีตข้อมูล ชื่อ `data.xlsx` → อัปโหลดทับใน `data/`
const realSymptoms = (r) => r.symptoms.filter((s) => s !== "/");	
const elementOf = (desc) => ELEMENTS.filter((e) => desc.includes(e.key));	## ทดสอบบนเครื่อง
const asEl = (t) => (t && t.nodeType === 1 ? t : t && t.parentElement);	```bash
	python -m http.server 8000
/* ---------- ลดการคัดลอก (ยับยั้งผู้ใช้ทั่วไปได้ระดับหนึ่ง) ---------- */	# แล้วเปิด http://localhost:8000
function protectPage() {	```
<img width="1681" height="929" alt="image" src="https://github.com/user-attachments/assets/7a4fa241-db12-4922-96af-eb0d1ee37261" />

git add .
git commit -m "AASYMA visual learning site"
git branch -M main
git remote add origin https://github.com/<ชื่อผู้ใช้>/<ชื่อ-repo>.git
git push -u origin main
```

## อัปเดตข้อมูล
แก้ข้อมูลในชีต AASYMA ใน Excel → อัปโหลด `data/aasyma.xlsx` ทับไฟล์เดิม → หน้าเว็บอัปเดตเอง
(หรือลากไฟล์ Excel ไปวางบนหน้าเว็บเพื่อดูข้อมูลชั่วคราวได้ทันที โดยไม่ต้องอัปโหลด)

## ทดสอบบนเครื่อง
```bash
python -m http.server 8000
# แล้วเปิด http://localhost:8000
```
(การเปิด index.html ตรง ๆ แบบ double-click จะโหลด data/aasyma.xlsx ไม่ได้ ให้ใช้การลากไฟล์มาวางแทน)

## หมายเหตุ
- รูปแบบข้อมูล: ข้อมูลเริ่มแถว 3, คอลัมน์ B=รหัส, C=ชื่อ, D=คำอธิบาย, E=ระดับ, F=ภาวะ, G–Z=อาการ 1–20
- "คำช่วยจำ" ในการ์ดเป็นตัวช่วยเรียนที่เพิ่มขึ้น ไม่ได้มาจากตำรา
- ใช้ไลบรารี [SheetJS](https://sheetjs.com) อ่านไฟล์ Excel ในเบราว์เซอร์
![Uploading image.png…]()
