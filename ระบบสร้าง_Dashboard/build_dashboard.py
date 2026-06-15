# -*- coding: utf-8 -*-
"""
build_dashboard.py  (v2 — สไตล์ BCM Dashboard)
ฝังโครงสร้างแผน (ACTIVITIES) ลงใน index.html
ผลงานจริงโหลดแบบ runtime จาก data/rm_actual.csv (MS Lists -> GitHub) ไม่ต้อง rebuild เมื่อข้อมูลเปลี่ยน
rebuild เฉพาะเมื่อ "แผน" เปลี่ยน:  python3 parse_plan.py <excel> && python3 build_dashboard.py
"""
import json, csv, sys, datetime, os

HERE=os.path.dirname(os.path.abspath(__file__)); ROOT=os.path.dirname(HERE)
PLAN=os.path.join(HERE,"plan_data.json"); TPL=os.path.join(HERE,"template.html"); OUT=os.path.join(ROOT,"index.html")
DATACSV=os.path.join(ROOT,"data","rm_actual.csv")
EMBED_SAMPLE = sys.argv[1] if len(sys.argv)>1 else DATACSV

d=json.load(open(PLAN,encoding="utf-8"))
MONTHS=d["months"]

# ---- รวบรวมกิจกรรมที่ track (leaf): sec1/3=subplan, sec2=measure ----
rows=[]
for s in d["sections"]:
    for sub in s["children"]:
        if sub["type"]=="measure":
            rows.append((s["code"], sub, None))
        elif sub["type"]=="subplan":
            if sub["children"]:
                for m in sub["children"]:
                    rows.append((s["code"], m, sub))
            else:
                rows.append((s["code"], sub, None))

N=len(rows)
base_w=round(100.0/N,2)

# อ่านน้ำหนักที่ผู้ใช้กำหนดเอง (ถ้ามีไฟล์ weights.csv: ID,weight)
wmap={}
if os.path.exists("weights.csv"):
    for r in csv.DictReader(open("weights.csv",encoding="utf-8-sig")):
        try: wmap[r["ID"].strip()]=float(r["weight"])
        except: pass

def bits(sched):
    return "".join("1" if i in (sched or []) else "0" for i in range(12))

acts=[]
for sc,node,parent in rows:
    iid=node["id"]
    tasks=[{"name":t["name"],"bits":bits(t.get("sched"))} for t in node.get("tasks",[])]
    acts.append({
        "id":iid,"section":sc,
        "subplan":(parent["code"] if parent else (node["code"].rsplit('.',1)[0] if node["type"]=="measure" else node["code"])),
        "code":node["code"],"name":node["name"],
        "output":(node.get("indicator") or (parent.get("indicator") if parent else None) or ""),
        "resp":(node.get("resp") or (parent.get("resp") if parent else None) or ""),
        "budget":(node.get("budget") if isinstance(node.get("budget"),(int,float)) else ""),
        "weight":wmap.get(iid, base_w),
        "plan":[(v if v is not None else None) for v in (node.get("plan_pct") or [None]*12)],
        "tasks":tasks,
    })

sections=[{"code":s["code"],"name":s["name"]} for s in d["sections"]]

RAW={
  "year":d["year"],"year_be":int(d["year"]),"months":MONTHS,
  "sections":sections,"activities":acts,
  "built":datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
}

# ---- ฝังตัวอย่างผลจริง (fallback เวลาเปิดไฟล์ local ที่ fetch ไม่ได้) ----
embed_rows=[]
if os.path.exists(EMBED_SAMPLE):
    with open(EMBED_SAMPLE,encoding="utf-8-sig") as f:
        embed_rows=list(csv.reader(f))

tpl=open(TPL,encoding="utf-8").read()
html=tpl.replace("/*__RAW__*/", json.dumps(RAW,ensure_ascii=False))
html=html.replace("/*__EMBED_ACTUAL__*/", json.dumps(embed_rows,ensure_ascii=False))
open(OUT,"w",encoding="utf-8").write(html)

# (ผลจริงอยู่ที่ ../data/rm_actual.csv — อัปเดตจาก SharePoint export โดยตรง ไม่เขียนทับที่นี่)
os.makedirs(os.path.join(ROOT,"data"),exist_ok=True)

print(f"OK -> {OUT} ({len(html):,} bytes) | activities={N} weight/each≈{base_w}% | embed_rows={len(embed_rows)-1 if embed_rows else 0}")
