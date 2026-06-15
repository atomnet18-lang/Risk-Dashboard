# -*- coding: utf-8 -*-
"""ไฟล์ Excel ชีตเดียว พร้อม import SharePoint List — ระดับกิจกรรมย่อย (126 รายการ)"""
import os
import openpyxl
from openpyxl.styles import Font,PatternFill,Alignment,Border,Side
from openpyxl.worksheet.table import Table,TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter as col
from common_items import build,leaves,CANON,DONE,DOING,NOT
HERE=os.path.dirname(os.path.abspath(__file__)); ROOT=os.path.dirname(HERE)
FORMDIR=os.path.join(ROOT,"ฟอร์ม_SharePoint"); os.makedirs(FORMDIR,exist_ok=True)

M,acts,secs=build(); lv=leaves(acts)
NAVY="16275C"; HF=Font(color="FFFFFF",bold=True,size=11); HEADF=PatternFill("solid",fgColor=NAVY)
REF=PatternFill("solid",fgColor="EEF2F7")
thin=Side(style="thin",color="C7CED6"); B=Border(thin,thin,thin,thin)
ctr=Alignment(horizontal="center",vertical="center",wrap_text=True); wrap=Alignment(vertical="center",wrap_text=True)

wb=openpyxl.Workbook(); ws=wb.active; ws.title="ผลการดำเนินงาน"
ws.append(CANON)
for l in lv:
    ws.append([l["id"],l["sec"],l["subplan"],f'{l["parent_code"]} {l["parent_name"]}',l["name"],l["resp"],"","","","","","",""])
for c in range(1,len(CANON)+1):
    x=ws.cell(1,c); x.fill=HEADF; x.font=HF; x.alignment=ctr; x.border=B
for r in range(2,ws.max_row+1):
    for c in range(1,len(CANON)+1):
        cell=ws.cell(r,c); cell.border=B; cell.alignment=wrap if c in (4,5,10,11) else ctr
        if c<=6: cell.fill=REF
        if c==13: cell.number_format="yyyy-mm-dd"
widths=[15,8,10,34,40,20,9,12,20,30,24,15,14]
for i,w in enumerate(widths,1): ws.column_dimensions[col(i)].width=w
ws.row_dimensions[1].height=30; ws.freeze_panes="A2"
ref=f"A1:{col(len(CANON))}{ws.max_row}"
tbl=Table(displayName="RMActual",ref=ref); tbl.tableStyleInfo=TableStyleInfo(name="TableStyleMedium2",showRowStripes=True)
ws.add_table(tbl)
dvm=DataValidation(type="list",formula1='"'+",".join(M)+'"',allow_blank=True)
dvp=DataValidation(type="whole",operator="between",formula1=0,formula2=100,allow_blank=True)
dvs=DataValidation(type="list",formula1=f'"{DONE},{DOING},{NOT}"',allow_blank=True)
ws.add_data_validation(dvm);ws.add_data_validation(dvp);ws.add_data_validation(dvs)
dvm.add("G2:G5000");dvp.add("H2:H5000");dvs.add("I2:I5000")
wb.save(os.path.join(FORMDIR,"SharePoint_List_RM_2569.xlsx"))
print(f"saved SharePoint (1 sheet, table RMActual) | leaf rows={len(lv)} cols={len(CANON)}")
