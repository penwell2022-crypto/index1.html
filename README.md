# 🧭 แผนภาพธาตุ 4 – AASYMA (Visual · Kinesthetic · Active Learning)

เว็บไซต์แสดงข้อมูล "อาการแยกตามธาตุ" จากชีต **AASYMA** ในไฟล์ Excel
เป็นเว็บแบบ static (HTML/CSS/JS) จึงเปิดบน **GitHub Pages** ได้โดยไม่ต้องมีเซิร์ฟเวอร์

## ส่วนต่าง ๆ บนหน้าเว็บ
| ส่วน | รูปแบบการเรียน | คำอธิบาย |
|---|---|---|
| ① การ์ดธาตุ | Visual | การ์ดสี 5 ธาตุ จำนวนภาวะ/อาการ คำช่วยจำ (คลิกการ์ดเพื่อไปสำรวจ) |
| ② บันไดความรุนแรง | Visual | หย่อน → กำเริบ → พิการ → ออกจากกาย → แตก → ขาด (คลิกช่องเพื่อดูอาการทั้งหมด) |
| ③ แผนที่กลุ่มธาตุ | Visual | ตาราง ● แสดงธาตุที่เกี่ยวข้องในกลุ่ม A01–A21 + กราฟแท่ง |
| ④ สำรวจเอง | Active | เลือกกลุ่ม / ค้นหาอาการ แล้วไฮไลต์อาการที่ตรง |
| ⑤ แบบทดสอบ | Active | สุ่ม 8 ข้อ กดเลือกธาตุ ตรวจทันที นับคะแนน |
| ⑥ ลากวางจับคู่ | Kinesthetic | ลากอาการไปใส่กล่องธาตุ (มือถือ: แตะอาการ แล้วแตะกล่อง) |

## โครงสร้างไฟล์
```
index.html      หน้าเว็บ
style.css       สี/รูปแบบ (สีเดียวกับชีต AASYMA_Visual)
app.js          โค้ดอ่าน Excel และสร้างแผนภาพทั้งหมด
data/aasyma.xlsx  ← ใส่ไฟล์ Excel ของคุณที่นี่ (ต้องมีชีตชื่อ AASYMA)
.nojekyll
```

## วิธีขึ้น GitHub และเปิดเป็นเว็บไซต์
1. สร้าง repository ใหม่บน GitHub (เช่น `thai-element-aasyma`) ตั้งเป็น **Public**
2. อัปโหลดไฟล์ทั้งหมดในโฟลเดอร์นี้ (Add file → Upload files)
3. นำไฟล์ Excel (`ZoodsP_ธาตุ4.xlsx`) มาเปลี่ยนชื่อเป็น **`aasyma.xlsx`** แล้วอัปโหลดไว้ในโฟลเดอร์ **`data/`**
4. ไปที่ **Settings → Pages** → Source: *Deploy from a branch* → Branch: `main` / `(root)` → Save
5. รอ 1–2 นาที เว็บจะอยู่ที่ `https://<ชื่อผู้ใช้>.github.io/<ชื่อ-repo>/`

หรือใช้คำสั่ง git:
```bash
git init
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
