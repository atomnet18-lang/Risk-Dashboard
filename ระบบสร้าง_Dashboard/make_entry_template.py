# -*- coding: utf-8 -*-
"""เทมเพลตกรอกผล: 1 แถว = 1 กิจกรรมย่อย · มีคอลัมน์ "เป้าหมาย" (แก้เองได้) + "ผลจริง" รายเดือน
ออก -> ../บันทึกผลการดำเนินงาน_RM_2569.xlsx"""
import os, random
import openpyxl
from openpyxl.styles import Font,PatternFill,Alignment,Border,Side
from openpyxl.worksheet.table import Table,TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter as col
from common_items import build,leaves
HERE=os.path.dirname(os.path.abspath(__file__)); ROOT=os.path.dirname(HERE)
M,acts,secs=build(); lv=leaves(acts)
info={}
for a in acts:
    if a["subs"]:
        for s in a["subs"]:
            info[s["id"]]=("bits",[i for i in range(12) if s["bits"][i]=="1"])
    else: info[a["id"]]=("plan",a["plan"])
def target_cum(idd):
    kind,d=info[idd]
    if kind=="bits":
        if not d: return [None]*12
        raw=[ (round(100*sum(1 for x in d if x<=m)/len(d)) if any(x<=m for x in d) else None) for m in range(12)]
    else:
        raw=[(v if v is not None else None) for v in (d or [None]*12)]
    out=[]; last=None
    for v in raw:
        if v is not None: last=v
        out.append(last)   # carry-forward (None ก่อนเริ่ม)
    return out

NAVY="16275C"; HF=Font(color="FFFFFF",bold=True,size=10); HEADF=PatternFill("solid",fgColor=NAVY)
TGTH=PatternFill("solid",fgColor="8a6d1f"); ACTH=PatternFill("solid",fgColor="1e3a8a")
REF=PatternFill("solid",fgColor="EEF2F7"); TGF=PatternFill("solid",fgColor="FBF3DD"); ACF=PatternFill("solid",fgColor="EAF1FB")
thin=Side(style="thin",color="C7CED6"); B=Border(thin,thin,thin,thin)
ctr=Alignment(horizontal="center",vertical="center",wrap_text=True); wrap=Alignment(vertical="center",wrap_text=True)
title=Font(bold=True,size=13,color=NAVY)

wb=openpyxl.Workbook(); ws=wb.active; ws.title="บันทึกผล"
HEAD=["รหัสกิจกรรม","แผนหลัก","แผนงานย่อย","กิจกรรมหลัก","ชื่อกิจกรรมย่อย","ผู้รับผิดชอบ"]\
     +[f"เป้า {m}" for m in M]+[f"ผล {m}" for m in M]+["รายละเอียดการดำเนินการ (ล่าสุด)","ปัญหาอุปสรรค","ผู้รายงาน"]
ws.append(HEAD)
random.seed(7); CUR=5
for l in lv:
    tg=target_cum(l["id"])
    tcells=[("" if v is None else v) for v in tg]
    # prefill actual (ตัวอย่างถึง มิ.ย.)
    acells=[""]*12; last=0; latest=-1
    for mi in range(CUR+1):
        t=tg[mi]
        if t is None: continue
        act=max(last,min(100,round(t+random.choice([0,0,5,-6,-15,4,-3,-20])))); last=act; latest=mi
        acells[mi]=act
    detail=f'ดำเนินการ {l["name"][:20]} (ถึง {M[latest]})' if latest>=0 else ""
    issue="ล่าช้ากว่าแผน ติดขั้นตอนอนุมัติ" if (latest>=0 and tg[latest] is not None and (tg[latest]-acells[latest])>=15) else ""
    rep=l["resp"].split("/")[-1].strip() or "-"
    ws.append([l["id"],l["sec"],l["subplan"],f'{l["parent_code"]} {l["parent_name"]}',l["name"],l["resp"]]+tcells+acells+[detail,issue,rep])
# styling
NC=len(HEAD)
for c in range(1,NC+1):
    x=ws.cell(1,c); x.font=HF; x.alignment=ctr; x.border=B
    x.fill=TGTH if 7<=c<=18 else ACTH if 19<=c<=30 else HEADF
for r in range(2,ws.max_row+1):
    for c in range(1,NC+1):
        cell=ws.cell(r,c); cell.border=B; cell.alignment=wrap if c in (4,5,31,32) else ctr
        if c<=6: cell.fill=REF
        elif 7<=c<=18: cell.fill=TGF
        elif 19<=c<=30: cell.fill=ACF
widths=[15,7,9,28,36,16]+[7]*24+[32,24,13]
for i,w in enumerate(widths,1): ws.column_dimensions[col(i)].width=w
ws.row_dimensions[1].height=30; ws.freeze_panes="G2"
tbl=Table(displayName="RMEntry",ref=f"A1:{col(NC)}{ws.max_row}")
tbl.tableStyleInfo=TableStyleInfo(name="TableStyleLight9",showRowStripes=False); ws.add_table(tbl)
dvp=DataValidation(type="whole",operator="between",formula1=0,formula2=100,allow_blank=True)
ws.add_data_validation(dvp); dvp.add(f"G2:AD{ws.max_row}")

# Sheet 2: วิธีใช้
w3=wb.create_sheet("วิธีใช้")
steps=[["วิธีกรอกและนำเข้า Dashboard",""],["",""],
 ["1)","ชีต \"บันทึกผล\" — แต่ละแถวคือกิจกรรมย่อย 1 ข้อ (คอลัมน์ 1-6 เติมให้แล้ว ห้ามแก้รหัส)"],
 ["2)","กลุ่ม \"เป้า ม.ค.–ธ.ค.\" (พื้นเหลือง) = ค่าเป้าหมาย % สะสม — พรีฟิลตามแผนให้แล้ว แก้ได้ตามต้องการ"],
 ["3)","กลุ่ม \"ผล ม.ค.–ธ.ค.\" (พื้นฟ้า) = % ผลจริงสะสม กรอกเฉพาะเดือนที่อัปเดต"],
 ["","   เดือนที่เว้นว่าง ระบบถือว่าคงค่าล่าสุด (ทั้งเป้าหมายและผลจริง)"],
 ["4)","รายละเอียด/ปัญหา/ผู้รายงาน = ของเดือนล่าสุด"],
 ["",""],
 ["แปลงเข้า GitHub:",""],
 ["5)","python3 ระบบสร้าง_Dashboard/convert_to_csv.py   → สร้าง data/rm_actual.csv + อัปเดต index.html"],
 ["6)","git add -A && git commit -m \"update\" && git push"],
 ["",""],
 ["หมายเหตุ","Dashboard ใช้ค่า \"เป้าหมาย\" จากไฟล์นี้โดยตรง (ถ้าแก้เป้าหมาย กราฟ/แถบ/สถานะ จะขยับตาม)"],
 ["","สถานะงานคำนวณจากผลจริง: 100=เสร็จสิ้น · 1-99=อยู่ระหว่างดำเนินการ · 0/ว่าง=ยังไม่ดำเนินการ"]]
for r in steps: w3.append(r)
w3["A1"].font=title
for r in range(1,w3.max_row+1):
    w3.cell(r,1).alignment=Alignment(vertical="top",horizontal="right")
    w3.cell(r,2).alignment=Alignment(vertical="top",wrap_text=True)
w3["A9"].font=title
w3.column_dimensions["A"].width=10; w3.column_dimensions["B"].width=92

out=os.path.join(ROOT,"บันทึกผลการดำเนินงาน_RM_2569.xlsx")
wb.save(out)
print("saved:",os.path.basename(out),"| cols=",NC,"| rows=",len(lv))
