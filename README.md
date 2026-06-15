# Dashboard แผนปฏิบัติการบริหารความเสี่ยง ปี 2569 (RM & IC)

เว็บเผยแพร่ 👉 **https://atomnet18-lang.github.io/Risk-Dashboard/**
ติดตามความก้าวหน้า แผน vs ผลจริง 3 แผนหลัก · 48 กิจกรรม · ผู้รับผิดชอบกรอกผ่าน Microsoft Lists

---

## 📁 โครงสร้างโฟลเดอร์

```
Risk-Dashboard/
├── index.html                    ← เว็บ Dashboard (สร้างจากระบบ อย่าแก้มือ)
├── data/
│   └── rm_actual.csv             ← ★ ไฟล์ผลจริง — อัปเดตทุกเดือน (จุดเดียวที่ต้องแก้)
│
├── ฟอร์ม_SharePoint/             ← ไฟล์สำหรับ SharePoint
│   ├── SharePoint_List_RM_2569.xlsx   (อัปโหลดสร้าง List — มี 48 กิจกรรม)
│   └── MS_List_template.xlsx          (อ้างอิงคอลัมน์ + รหัส 48 รายการ)
│
├── ระบบสร้าง_Dashboard/          ← โค้ด (แตะเฉพาะตอน "แผน" เปลี่ยน)
│   ├── parse_plan.py  build_dashboard.py  make_sample.py
│   ├── make_sharepoint.py  common_items.py  template.html
│   ├── plan_data.json  items_master.csv  actuals_template.csv
│   └── แผนปฏิบัติการบริหารความเสี่ยง_2569.xlsx   (ไฟล์แผนต้นฉบับ)
│
├── README.md   .nojekyll   .gitignore
```

> 🟢 ปกติยุ่งแค่ `data/rm_actual.csv` · 🟡 `ฟอร์ม_SharePoint/` ใช้ตอนตั้ง List · 🔴 `ระบบสร้าง_Dashboard/` ใช้ตอนแผนเปลี่ยน

---

## 🔄 อัปเดตข้อมูลรายเดือน (ทำบ่อย — ง่ายสุด)

1. ผู้รับผิดชอบกรอกข้อมูลใน **SharePoint List**
2. List → **Export to CSV** → เปลี่ยนชื่อเป็น `rm_actual.csv`
3. วางทับไฟล์เดิมที่ **`data/rm_actual.csv`**
4. push ขึ้น GitHub:
```bash
git add data/rm_actual.csv
git commit -m "update actuals - <เดือน>"
git push
```
→ เว็บอัปเดตเองใน ~2 นาที (ไม่ต้องแตะโค้ด)

> รายเดือนถัดไปให้ **เพิ่มแถวใหม่** (เดือนใหม่ + % สะสม) ใน List เพื่อให้กราฟ S-Curve มีจุดข้อมูลครบทุกเดือน

---

## ⚙️ ตั้งค่า SharePoint List (ครั้งเดียว)

เปิด `ฟอร์ม_SharePoint/SharePoint_List_RM_2569.xlsx` (ชีตเดียว ตาราง `RMActual` มีครบ 48 กิจกรรม) →
SharePoint > **+ New > List > From Excel** → เลือกตาราง `RMActual` → ตั้งชนิดคอลัมน์ → Create

| # | คอลัมน์ | ชนิด | สถานะ |
|---|---|---|---|
| 1 | รหัสกิจกรรม | Text (`M_2.1.1`,`P_3.1`) — **ห้ามชื่อ "ID"** | เติมให้แล้ว |
| 2 | แผนหลัก | Text | เติมให้แล้ว |
| 3 | แผนงานย่อย | Text | เติมให้แล้ว |
| 4 | ชื่อกิจกรรม/มาตรการ | Text | เติมให้แล้ว |
| 5 | ผู้รับผิดชอบ | Text/Person | เติมให้แล้ว |
| 6 | เดือน | Choice (ม.ค.–ธ.ค.) | **กรอก** |
| 7 | percent_จริง | Number 0–100 | **กรอก** |
| 8 | สถานะงาน | Choice: เสร็จสิ้น/อยู่ระหว่างดำเนินการ/ยังไม่ดำเนินการ | **กรอก** |
| 9 | รายละเอียดการดำเนินการ | Multiple lines | **กรอก** |
| 10 | ปัญหาอุปสรรค | Multiple lines | **กรอก** |
| 11 | ผู้รายงาน | Person/Text | **กรอก** |
| 12 | วันที่รายงาน | Date | **กรอก** |

> Dashboard อ่านเฉพาะ: รหัสกิจกรรม, เดือน, percent_จริง, สถานะงาน, รายละเอียด, ปัญหา

---

## 🚀 เผยแพร่ครั้งแรก (ครั้งเดียว) — repo `atomnet18-lang/Risk-Dashboard`

รันในโฟลเดอร์นี้ (ที่มี index.html):
```bash
git init
git add .
git commit -m "init RM dashboard 2569"
git branch -M main
git remote add origin https://github.com/atomnet18-lang/Risk-Dashboard.git
git push -u origin main
```
จากนั้น repo → **Settings → Pages → Deploy from a branch → main / (root) → Save**

---

## 🛠 เมื่อ "แผน" เปลี่ยน (นานๆ ครั้ง)

แก้ไฟล์ `ระบบสร้าง_Dashboard/แผนปฏิบัติการบริหารความเสี่ยง_2569.xlsx` แล้ว:
```bash
cd ระบบสร้าง_Dashboard
python3 parse_plan.py        # อ่านแผน → plan_data.json
python3 make_sample.py       # อัปเดตรหัส/เทมเพลต + MS_List_template
python3 make_sharepoint.py   # อัปเดตไฟล์ SharePoint (48 กิจกรรม)
python3 build_dashboard.py   # สร้าง index.html ใหม่
cd ..
git add -A && git commit -m "update plan" && git push
```

---

## เกณฑ์สถานะเทียบแผน
ตามแผน/เร็วกว่าแผน = ผลงาน ≥ เป้าหมาย · ล่าช้าเล็กน้อย = ต่ำกว่า < 10 จุด · ล่าช้า = ต่ำกว่า ≥ 10 จุด · ยังไม่ถึงกำหนด/รอรายงานผล = ยังไม่ถึงเดือน/ยังไม่กรอก

## ปรับน้ำหนักกิจกรรม
ค่าเริ่มต้นเฉลี่ยเท่ากัน (≈2.08%) — สร้าง `ระบบสร้าง_Dashboard/weights.csv` (`ID,weight`) แล้วรัน build ใหม่ ถ้าต้องการถ่วงตามความสำคัญ
