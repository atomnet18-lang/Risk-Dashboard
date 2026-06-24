#!/bin/bash
# ดับเบิลคลิกไฟล์นี้เพื่อแปลงข้อมูลจาก Excel -> data/rm_actual.csv และอัปเดต index.html
# จากนั้นเปิด GitHub Desktop เพื่อ Commit + Push
cd "$(dirname "$0")"
echo "==============================================="
echo " อัปเดตข้อมูล Dashboard บริหารความเสี่ยง (RM) 2569"
echo "==============================================="
echo ""

# ตรวจ/ติดตั้ง openpyxl ให้อัตโนมัติ (ครั้งแรกอาจใช้เวลาสักครู่)
python3 -c "import openpyxl" 2>/dev/null || {
  echo "กำลังติดตั้งโมดูล openpyxl ..."
  pip3 install --user openpyxl --quiet 2>/dev/null || pip3 install --break-system-packages openpyxl --quiet 2>/dev/null
}

echo "กำลังแปลงข้อมูลจากไฟล์ บันทึกผลการดำเนินงาน_RM_2569.xlsx ..."
echo ""
python3 "ระบบสร้าง_Dashboard/convert_to_csv.py"
status=$?
echo ""
if [ $status -eq 0 ]; then
  echo "✓ แปลงข้อมูล + อัปเดต index.html เรียบร้อย"
  echo "  ขั้นต่อไป: เปิด GitHub Desktop -> ใส่ข้อความ -> Commit to main -> Push origin"
  echo "  แล้วเว็บ https://atomnet18-lang.github.io/Risk-Dashboard/ จะอัปเดตใน ~2 นาที"
else
  echo "✗ เกิดข้อผิดพลาด"
  echo "  ตรวจสอบ:"
  echo "   1) บันทึก + ปิดไฟล์ บันทึกผลการดำเนินงาน_RM_2569.xlsx ก่อนรัน"
  echo "   2) ติดตั้ง Python 3 แล้ว (พิมพ์ python3 --version ใน Terminal)"
fi
echo ""
read -p "กด Enter เพื่อปิดหน้าต่างนี้..."
