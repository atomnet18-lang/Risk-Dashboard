# -*- coding: utf-8 -*-
"""build_dashboard.py — สร้าง ../index.html (ติดตามระดับกิจกรรมย่อย)
ผลจริงโหลด runtime จาก ../data/rm_actual.csv  | rebuild เฉพาะตอนแผนเปลี่ยน"""
import json, csv, sys, datetime, os
from common_items import build

HERE=os.path.dirname(os.path.abspath(__file__)); ROOT=os.path.dirname(HERE)
TPL=os.path.join(HERE,"template.html"); OUT=os.path.join(ROOT,"index.html")
DATACSV=os.path.join(ROOT,"data","rm_actual.csv")
EMBED_SAMPLE = sys.argv[1] if len(sys.argv)>1 else DATACSV

M,acts,sections=build()
RAW={"year":"2569","year_be":2569,"months":M,"sections":sections,"activities":acts,
     "built":datetime.datetime.now().strftime("%Y-%m-%d %H:%M")}

embed_rows=[]
if os.path.exists(EMBED_SAMPLE):
    with open(EMBED_SAMPLE,encoding="utf-8-sig") as f:
        embed_rows=list(csv.reader(f))

tpl=open(TPL,encoding="utf-8").read()
html=tpl.replace("/*__RAW__*/", json.dumps(RAW,ensure_ascii=False))
html=html.replace("/*__EMBED_ACTUAL__*/", json.dumps(embed_rows,ensure_ascii=False))
open(OUT,"w",encoding="utf-8").write(html)
os.makedirs(os.path.join(ROOT,"data"),exist_ok=True)
nsub=sum(len(a["subs"]) for a in acts)
print(f"OK -> {OUT} ({len(html):,} chars) | activities={len(acts)} sub-tasks={nsub} embed_rows={len(embed_rows)-1 if embed_rows else 0}")
