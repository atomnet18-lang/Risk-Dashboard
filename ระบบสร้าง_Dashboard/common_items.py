# -*- coding: utf-8 -*-
"""โหลดรายการกิจกรรม 48 รายการ (ใช้ร่วมกันทุกสคริปต์) + คอลัมน์มาตรฐาน"""
import json, os
HERE=os.path.dirname(os.path.abspath(__file__))
CANON=["รหัสกิจกรรม","แผนหลัก","แผนงานย่อย","ชื่อกิจกรรม/มาตรการ","ผู้รับผิดชอบ",
       "เดือน","percent_จริง","สถานะงาน","รายละเอียดการดำเนินการ","ปัญหาอุปสรรค","ผู้รายงาน","วันที่รายงาน"]
DONE,DOING,NOT="เสร็จสิ้น","อยู่ระหว่างดำเนินการ","ยังไม่ดำเนินการ"
def load(plan=None):
    plan=plan or os.path.join(HERE,"plan_data.json")
    d=json.load(open(plan,encoding="utf-8")); M=d["months"]; rows=[]
    for s in d["sections"]:
        for sub in s["children"]:
            if sub["type"]=="measure": rows.append((s["code"],sub,None))
            elif sub["type"]=="subplan":
                if sub["children"]:
                    for m in sub["children"]: rows.append((s["code"],m,sub))
                else: rows.append((s["code"],sub,None))
    items=[]
    for sc,node,parent in rows:
        items.append({"id":node["id"],"sec":sc,
            "sub":(parent["code"] if parent else node["code"]),
            "code":node["code"],"name":node["name"],
            "resp":node.get("resp") or (parent.get("resp") if parent else None) or "",
            "plan":node.get("plan_pct") or [None]*12})
    N=len(items); base_w=round(100.0/N,2)
    import csv as _csv; wmap={}
    wpath=os.path.join(HERE,"weights.csv")
    if os.path.exists(wpath):
        for r in _csv.DictReader(open(wpath,encoding="utf-8-sig")):
            try: wmap[r["ID"].strip()]=float(r["weight"])
            except: pass
    for it in items: it["weight"]=wmap.get(it["id"],base_w)
    return M,items
