#!/usr/bin/env python3
"""SeptemarDump makrosunun _dump.tsv ciktisini uygulamanin okuyacagi <kod>-sw.json'a cevirir.
Kullanim: python3 sw_dump_to_json.py dump.tsv data/68-109-sw.json"""
import sys,json,re
src,dst=sys.argv[1:3]
raw=open(src,'rb').read()
try:txt=raw.decode('utf-8')
except UnicodeDecodeError:txt=raw.decode('cp1254')
P={}
for l in txt.splitlines():
    r=l.split('\t')
    if r[0]=='PROP' and r[2].startswith(('CFG','FILE')):P.setdefault(r[1],{})[r[3]]=r[4]
def g(p,k):
    v=p.get(k);return '' if v is None else v.strip()
def num(v):
    try:return float(v)
    except:return 0
out={}
for path,p in P.items():
    if not g(p,'SWOODCP_PanelLength') and not g(p,'SWOODCP_CoreStockMaterial'):continue
    leaf=path.split('/')[-1];leaf=re.sub(r'-\d+$','',leaf)
    leaf=re.sub(r'[^\x20-\x7e]','?',leaf.strip()).upper()
    e={'top':g(p,'SWOODCP_TopStockMaterial'),'bot':g(p,'SWOODCP_BottomStockMaterial'),'core':g(p,'SWOODCP_CoreStockMaterial'),
       'desen':g(p,'Desen_Takibi'),'parkur':g(p,'Makina_Parkuru'),
       'boy':num(g(p,'SWOODCP_PanelLength')),'en':num(g(p,'SWOODCP_PanelWidth')),'kal':num(g(p,'SWOODCP_PanelThickness')),
       'kenar':{s:{'ad':g(p,'SWOODCP_Edge%sName'%s),'mal':g(p,'SWOODCP_Edge%sMaterial'%s),'kal':num(g(p,'SWOODCP_Edge%sThickness'%s)),'boy':num(g(p,'SWOODCP_Edge%sLength'%s))} for s in('Left','Front','Right','Back')}}
    out[leaf]=e
json.dump(out,open(dst,'w',encoding='utf-8'),ensure_ascii=False,separators=(',',':'))
print(len(out),'panel')
