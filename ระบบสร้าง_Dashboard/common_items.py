# -*- coding: utf-8 -*-
"""โมเดลข้อมูลกลาง — ติดตามระดับ "กิจกรรมย่อย" (sub-task)
ลำดับชั้น: แผนหลัก(3) -> กิจกรรม/มาตรการ(48) -> กิจกรรมย่อย(leaf)
leaf = กิจกรรมย่อยของแต่ละกิจกรรม ถ้ากิจกรรมไหนไม่มีย่อย (แผน 3) ใช้ตัวกิจกรรมเองเป็น leaf
"""
import json, os
HERE=os.path.dirname(os.path.abspath(__file__))

# คอลัมน์มาตรฐานสำหรับ SharePoint/CSV (ระดับกิจกรรมย่อย)
CANON=["รหัสกิจกรรม","แผนหลัก","แผนงานย่อย","กิจกรรมหลัก","ชื่อกิจกรรมย่อย","ผู้รับผิดชอบ",
       "เดือน","percent_เป้าหมาย","percent_จริง","สถานะงาน","รายละเอียดการดำเนินการ","ปัญหาอุปสรรค","ผู้รายงาน","วันที่รายงาน"]
DONE,DOING,NOT="เสร็จสิ้น","อยู่ระหว่างดำเนินการ","ยังไม่ดำเนินการ"

def _bits(sched):
    return "".join("1" if i in (sched or []) else "0" for i in range(12))

def build(plan=None):
    """คืน (MONTHS, activities, sections)
       activities[i] = กิจกรรมหลัก 48 ตัว แต่ละตัวมี subs[] (กิจกรรมย่อย)"""
    plan=plan or os.path.join(HERE,"plan_data.json")
    d=json.load(open(plan,encoding="utf-8")); M=d["months"]
    rows=[]
    for s in d["sections"]:
        for sub in s["children"]:
            if sub["type"]=="measure": rows.append((s["code"],sub,None))
            elif sub["type"]=="subplan":
                if sub["children"]:
                    for m in sub["children"]: rows.append((s["code"],m,sub))
                else: rows.append((s["code"],sub,None))
    base_w=round(100.0/len(rows),2)
    # weights override
    import csv as _csv; wmap={}
    wp=os.path.join(HERE,"weights.csv")
    if os.path.exists(wp):
        for r in _csv.DictReader(open(wp,encoding="utf-8-sig")):
            try: wmap[r["ID"].strip()]=float(r["weight"])
            except: pass
    acts=[]
    for sc,node,parent in rows:
        subs=[]  # ตัดระดับกิจกรรมย่อยออก — ติดตามที่ระดับกิจกรรม (48 หัวข้อ)
        acts.append({
            "id":node["id"],"section":sc,
            "subplan":(parent["code"] if parent else node["code"]),
            "code":node["code"],"name":node["name"],
            "output":(node.get("indicator") or (parent.get("indicator") if parent else None) or ""),
            "resp":(node.get("resp") or (parent.get("resp") if parent else None) or ""),
            "budget":(node.get("budget") if isinstance(node.get("budget"),(int,float)) else ""),
            "weight":wmap.get(node["id"],base_w),
            "plan":[(v if v is not None else None) for v in (node.get("plan_pct") or [None]*12)],
            "subs":subs,
        })
    sections=[{"code":s["code"],"name":s["name"],"plan":[(v if v is not None else None) for v in (s.get("plan") or [None]*12)]} for s in d["sections"]]
    return M,acts,sections

def leaves(acts):
    """รายการ leaf ที่กรอกผลได้ (กิจกรรมย่อย หรือ กิจกรรมที่ไม่มีย่อย)"""
    out=[]
    for a in acts:
        if a["subs"]:
            for s in a["subs"]:
                out.append({"id":s["id"],"sec":a["section"],"subplan":a["subplan"],
                    "parent_code":a["code"],"parent_name":a["name"],"name":s["name"],"resp":a["resp"]})
        else:
            out.append({"id":a["id"],"sec":a["section"],"subplan":a["subplan"],
                "parent_code":a["code"],"parent_name":a["name"],"name":a["name"],"resp":a["resp"]})
    return out
