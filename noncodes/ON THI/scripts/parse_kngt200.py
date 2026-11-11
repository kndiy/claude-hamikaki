import sys, re, json, zipfile
from xml.etree import ElementTree as ET
sys.stdout.reconfigure(encoding='utf-8')
W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
x=ET.fromstring(zipfile.ZipFile(sys.argv[1]).read('word/document.xml'))
bai=None; items=[]; cur=None; mode=None
for p in x.iter(W+'p'):
    t=''.join(e.text or '' for e in p.iter(W+'t')).strip()
    if not t: continue
    hl=''.join(''.join(e.text or '' for e in r.iter(W+'t')) for r in p.iter(W+'r') if r.find(W+'rPr/'+W+'highlight') is not None).strip()
    if re.match(r'^BÀI \d+\.',t): bai=t; continue
    m=re.match(r'^Câu (\d+)$',t)
    if m: cur={'bai':bai,'n':int(m.group(1)),'q':'','opts':{},'hl':''}; items.append(cur); mode='q'; continue
    if cur is None: continue
    if mode=='q': cur['q']=t; mode='o'; continue
    if mode=='o':
        parts=re.split(r'(?:(?<=^)|(?<=[^\s]))\s*([A-D])\.\s',' '+t if False else t)
        # parts: ['', 'A', text, 'B', text,...]
        it=iter(parts[1:]); cur['opts']={k:v.strip() for k,v in zip(it,it)}
        cur['hl']=hl; mode=None
for c in items:
    h=c['hl']; m=re.match(r'^([A-D])\.',h)
    if m: L=m.group(1)
    else:
        L=[k for k,v in c['opts'].items() if v.startswith(h) or h.startswith(v)]; L=L[0] if len(L)==1 else None
    c['L']=L; c['a']=c['opts'].get(L,'')
    c['a']=re.sub(r'\s*[֐-׿]+','',c['a']).strip()
    if len(c['opts'])!=4 or not L: print('CHECK',c['n'],c['opts'],repr(h))
json.dump(items,open(sys.argv[2],'w',encoding='utf-8'),ensure_ascii=False,indent=0)
print(len(items), sorted(set(i['bai'] for i in items)))
for i in items:
    hlt=re.sub(r'^[A-D]\.\s*','',i['hl']).strip()
    if hlt and not i['a'].startswith(hlt[:10]): print('MISMATCH',i['n'],hlt,'|',i['a'])
