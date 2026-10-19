#!/usr/bin/env python3
"""GLB kucultme: tekstur gorsellerini JPEG'e cevirip boyutlandirir. Kullanim: glb_compress.py in.glb out.glb [maxpx=1024] [kalite=78]"""
import sys,json,struct,io
from PIL import Image
src,dst=sys.argv[1:3];mx=int(sys.argv[3]) if len(sys.argv)>3 else 1024;q=int(sys.argv[4]) if len(sys.argv)>4 else 78
b=open(src,'rb').read()
jl=struct.unpack('<I',b[12:16])[0];j=json.loads(b[20:20+jl]);o=20+jl+8
binl=struct.unpack('<I',b[20+jl:24+jl])[0];bin_=b[o:o+binl]
bv=j['bufferViews'];imgs=j.get('images',[])
newimg={}
for i,im in enumerate(imgs):
    v=bv[im['bufferView']];raw=bin_[v.get('byteOffset',0):v.get('byteOffset',0)+v['byteLength']]
    p=Image.open(io.BytesIO(raw))
    alpha=p.mode in('RGBA','LA') and p.getchannel('A').getextrema()[0]<250
    if alpha:
        p=p.copy();p.thumbnail((mx,mx));buf=io.BytesIO();p.save(buf,'PNG',optimize=True);mt='image/png'
    else:
        p=p.convert('RGB');p.thumbnail((mx,mx));buf=io.BytesIO();p.save(buf,'JPEG',quality=q,optimize=True);mt='image/jpeg'
    newimg[im['bufferView']]=(buf.getvalue(),mt);im['mimeType']=mt
out=bytearray();nbv=[]
for i,v in enumerate(bv):
    data=newimg[i][0] if i in newimg else bin_[v.get('byteOffset',0):v.get('byteOffset',0)+v['byteLength']]
    while len(out)%4:out.append(0)
    nv=dict(v);nv['byteOffset']=len(out);nv['byteLength']=len(data);nbv.append(nv);out+=data
while len(out)%4:out.append(0)
j['bufferViews']=nbv;j['buffers']=[{'byteLength':len(out)}]
js=json.dumps(j,separators=(',',':')).encode();js+=b' '*((4-len(js)%4)%4)
tot=12+8+len(js)+8+len(out)
open(dst,'wb').write(struct.pack('<III',0x46546C67,2,tot)+struct.pack('<II',len(js),0x4E4F534A)+js+struct.pack('<II',len(out),0x004E4942)+bytes(out))
print(len(imgs),'doku',len(b)//1024//1024,'MB ->',tot//1024//1024,'MB')
