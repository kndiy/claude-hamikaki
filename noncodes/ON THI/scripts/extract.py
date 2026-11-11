import sys, os, re, zipfile, subprocess
sys.stdout.reconfigure(encoding="utf-8")
from xml.etree import ElementTree as ET
src, out = sys.argv[1], sys.argv[2]
W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
A='{http://schemas.openxmlformats.org/drawingml/2006/main}'
def docx(p):
    z=zipfile.ZipFile(p); root=ET.fromstring(z.read('word/document.xml'))
    lines=[]
    for para in root.iter(W+'p'):
        t=''.join(x.text or '' for x in para.iter(W+'t'))
        if t.strip(): lines.append(t)
    return '\n'.join(lines)
def pptx(p):
    z=zipfile.ZipFile(p)
    names=sorted([n for n in z.namelist() if re.match(r'ppt/slides/slide\d+\.xml$',n)],key=lambda n:int(re.findall(r'\d+',n)[0]))
    res=[]
    for n in names:
        root=ET.fromstring(z.read(n)); ls=[]
        for para in root.iter(A+'p'):
            t=''.join(x.text or '' for x in para.iter(A+'t'))
            if t.strip(): ls.append(t)
        res.append(f'--- Slide {re.findall(r"[0-9]+",n)[0]} ---\n'+'\n'.join(ls))
    return '\n'.join(res)
def pdf(p):
    import pymupdf
    d=pymupdf.open(p); return '\n'.join(f'--- Page {i+1} ---\n'+pg.get_text() for i,pg in enumerate(d))
def doc(p):
    return subprocess.run(['antiword','-m','UTF-8.txt',p],capture_output=True).stdout.decode('utf-8','replace')
skip=('Tham khao',)
for dp,_,fs in os.walk(src):
    if any(s in dp for s in skip): continue
    for f in fs:
        p=os.path.join(dp,f); ext=f.lower().rsplit('.',1)[-1]
        fn={'docx':docx,'pptx':pptx,'pdf':pdf,'doc':doc}.get(ext)
        if not fn or 'điểm danh' in f.lower(): continue
        try: t=fn(p)
        except Exception as e: t=f'ERROR {e}'
        rel=os.path.relpath(p,src).replace(os.sep,'__')
        open(os.path.join(out,rel+'.txt'),'w',encoding='utf-8').write(t)
        print(len(t), rel)
