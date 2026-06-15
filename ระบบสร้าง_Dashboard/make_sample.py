# -*- coding: utf-8 -*-
"""สร้าง items_master + เทมเพลต MS List (คำอธิบาย) + ข้อมูลตัวอย่าง (data/rm_actual.csv)
ใช้ชุดคอลัมน์มาตรฐานเดียวกับไฟล์ SharePoint (CANON)"""
import csv, random, os
import openpyxl
from openpyxl.styles import Font,PatternFill,Alignment,Border,Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter as col
import os
from common_items import load,CANON,DONE,DOING,NOT
HERE=os.path.dirname(os.path.abspath(__file__)); ROOT=os.path.dirname(HERE)
FORMDIR=os.path.join(ROOT,"ฟอร์ม_SharePoint"); os.makedirs(FORMDIR,exist_ok=True)

M,items=load()
NAVY="16275C"; HF=Font(color="FFFFFF",bold=True,size=11); HEADF=PatternFill("solid",fgColor=NAVY)
thin=Side(style="thin",color="BBBBBB"); B=Border(thin,thin,thin,thin)
ctr=Alignment(horizontal="center",vertical="center",wrap_text=True); wrap=Alignment(vertical="center",wrap_text=True)

# ---- items_master.csv (อ้างอิง + แผน%) ----
with open(os.path.join(HERE,"items_master.csv"),"w",encoding="utf-8-sig",newline="") as f:
    w=csv.writer(f)
    w.writerow(["รหัสกิจกรรม","แผนหลัก","แผนงานย่อย","รหัส","ชื่อกิจกรรม/มาตรการ","ผู้รับผิดชอบ","น้ำหนัก%","งบ(ลบ.)"]+[f"แผน%_{m}" for m in M])
    for r in items:
        pp=r["plan"] or [None]*12
        w.writerow([r["id"],r["sec"],r["sub"],r["code"],r["name"],r["resp"],r["weight"],""]+[("" if v is None else v) for v in pp])

# ---- actuals_template.csv (หัวเปล่า ชุด CANON) ----
with open(os.path.join(HERE,"actuals_template.csv"),"w",encoding="utf-8-sig",newline="") as f:
    csv.writer(f).writerow(CANON)

# ---- ตัวอย่างผลจริง (long format, สะสมถึง มิ.ย.) ----
random.seed(11); CUR=5; sample=[]
for r in items:
    pp=r["plan"] or [None]*12; last=0
    for mi in range(CUR+1):
        planv=None
        for j in range(mi+1):
            if pp[j] is not None: planv=pp[j]
        if planv is None: continue
        act=max(last,min(100,round(planv+random.choice([0,0,3,-4,-12,5,-2,-18,2])))); last=act
        st=DONE if act>=100 else (NOT if act==0 else DOING)
        note=f"{r['name'][:24]} — ดำเนินการเดือน {M[mi]} ({act}%)"
        issue="ติดขั้นตอนอนุมัติ ล่าช้ากว่าแผน" if (planv-act)>=12 else ""
        rep=r["resp"].split("/")[-1].strip() or "-"
        sample.append([r["id"],r["sec"],r["sub"],r["name"],r["resp"],M[mi],act,st,note,issue,rep,"2026-{:02d}-25".format(mi+1)])
for path in [os.path.join(HERE,"actuals_sample.csv")]:
    with open(path,"w",encoding="utf-8-sig",newline="") as f:
        wr=csv.writer(f); wr.writerow(CANON); wr.writerows(sample)

# ---- MS_List_template.xlsx (อ้างอิง 48 + ฟอร์ม CANON + คำอธิบาย) ----
wb=openpyxl.Workbook()
ws=wb.active; ws.title="กิจกรรม_Master"
h1=["รหัสกิจกรรม","แผนหลัก","แผนงานย่อย","รหัส","ชื่อกิจกรรม/มาตรการ","ผู้รับผิดชอบ","น้ำหนัก%","งบ(ลบ.)"]+[f"แผน%\n{m}" for m in M]
ws.append(h1)
for c in range(1,len(h1)+1):
    x=ws.cell(1,c); x.fill=HEADF; x.font=HF; x.alignment=ctr; x.border=B
for r in items:
    pp=r["plan"] or [None]*12
    ws.append([r["id"],r["sec"],r["sub"],r["code"],r["name"],r["resp"],r["weight"],""]+[("" if v is None else v) for v in pp])
for i,wd in enumerate([13,7,9,9,46,22,8,8]+[6]*12,1): ws.column_dimensions[col(i)].width=wd
for row in ws.iter_rows(min_row=2,max_row=ws.max_row):
    for c in row: c.border=B; c.alignment=wrap if c.column==5 else ctr
ws.freeze_panes="E2"; ws.row_dimensions[1].height=34

ws2=wb.create_sheet("ฟอร์มกรอกผล_CANON")
ws2.append(CANON)
for c in range(1,len(CANON)+1):
    x=ws2.cell(1,c); x.fill=HEADF; x.font=HF; x.alignment=ctr; x.border=B
ex=[items[15]["id"],items[15]["sec"],items[15]["sub"],items[15]["name"],items[15]["resp"],"มิ.ย.",30,DOING,"ทบทวนนโยบาย GRC เสนอผู้บริหารแล้ว","รออนุมัติคณะอนุกรรมการฯ","เป็ด","2026-06-25"]
ws2.append(ex)
for i,wd in enumerate([13,8,10,40,20,8,11,18,34,26,14,14],1): ws2.column_dimensions[col(i)].width=wd
dvs=DataValidation(type="list",formula1=f'"{DONE},{DOING},{NOT}"',allow_blank=True)
dvm=DataValidation(type="list",formula1='"'+",".join(M)+'"',allow_blank=True)
dvp=DataValidation(type="whole",operator="between",formula1=0,formula2=100,allow_blank=True)
ws2.add_data_validation(dvs);ws2.add_data_validation(dvm);ws2.add_data_validation(dvp)
dvm.add("F2:F5000");dvp.add("G2:G5000");dvs.add("H2:H5000")
for row in ws2.iter_rows(min_row=2,max_row=ws2.max_row):
    for c in row: c.border=B; c.alignment=wrap if c.column in(4,9,10) else ctr
ws2.freeze_panes="A2"

ws3=wb.create_sheet("คำอธิบาย")
desc=[["คอลัมน์","ชนิด (SharePoint)","คำอธิบาย / หมายเหตุ"],
 ["รหัสกิจกรรม","Single line of text","รหัส เช่น M_2.1.1, P_3.1 — ใช้เชื่อม Dashboard ห้ามแก้ (อย่าตั้งชื่อ 'ID')"],
 ["แผนหลัก","Single line of text","1. / 2. / 3. (เติมให้แล้วในไฟล์ SharePoint)"],
 ["แผนงานย่อย","Single line of text","เช่น 1.1, 2.4 (เติมให้แล้ว)"],
 ["ชื่อกิจกรรม/มาตรการ","Single line of text","ชื่อกิจกรรม (เติมให้แล้ว)"],
 ["ผู้รับผิดชอบ","Single line of text / Person","หน่วยงาน/ผู้รับผิดชอบ (เติมให้แล้ว)"],
 ["เดือน","Choice","เดือนที่รายงาน (ม.ค.–ธ.ค.) — กรอกตอนรายงาน"],
 ["percent_จริง","Number (0–100)","% ความคืบหน้าจริงสะสม — กรอกตอนรายงาน"],
 ["สถานะงาน","Choice",f"{DONE} / {DOING} / {NOT}"],
 ["รายละเอียดการดำเนินการ","Multiple lines of text","สิ่งที่ทำเดือนนั้น (โชว์ในป๊อปอัป)"],
 ["ปัญหาอุปสรรค","Multiple lines of text","ปัญหา/แนวทางแก้ไข (ถ้ามี)"],
 ["ผู้รายงาน","Person / Text","ผู้กรอกข้อมูลเดือนนั้น"],
 ["วันที่รายงาน","Date","วันที่กรอก"],
 ["—","—","คอลัมน์ 1–5 เติมให้แล้วในไฟล์ SharePoint · คอลัมน์ 6–12 ให้ผู้รับผิดชอบกรอกรายเดือน"],
 ["—","—","Dashboard อ่านเฉพาะ: รหัสกิจกรรม, เดือน, percent_จริง, สถานะงาน, รายละเอียด, ปัญหา (คอลัมน์อ้างอิงไม่กระทบ)"]]
for r in desc: ws3.append(r)
for c in range(1,4):
    x=ws3.cell(1,c); x.fill=HEADF; x.font=HF; x.alignment=ctr; x.border=B
for i,wd in enumerate([22,26,74],1): ws3.column_dimensions[col(i)].width=wd
for row in ws3.iter_rows(min_row=2,max_row=ws3.max_row):
    for c in row: c.border=B; c.alignment=wrap
wb.save(os.path.join(FORMDIR,"MS_List_template.xlsx"))
print("master items:",len(items),"| sample rows:",len(sample),"| CANON cols:",len(CANON))
