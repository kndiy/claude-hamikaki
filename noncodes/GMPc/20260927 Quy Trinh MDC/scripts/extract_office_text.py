import zipfile, sys, os, re
import xml.etree.ElementTree as ET
W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
def ptext(p):
    t=''.join(x.text or '' for x in p.iter() if x.tag in (W+'t',W+'tab') )
    st=p.find(f'{W}pPr/{W}pStyle')
    num=p.find(f'{W}pPr/{W}numPr')
    pre=''
    if st is not None: pre='['+st.get(W+'val')+'] '
    if num is not None: pre+='• '
    return pre+t
def cell(tc):
    return ' / '.join(ptext(p) for p in tc.iter(W+'p') if ptext(p).strip())
def body(el,out,depth=0):
    for c in el:
        if c.tag==W+'p':
            s=ptext(c)
            if s.strip(): out.append(s)
        elif c.tag==W+'tbl':
            out.append('<<TABLE>>')
            for tr in c.findall(W+'tr'):
                out.append('| '+' | '.join(cell(tc) for tc in tr.findall(W+'tc'))+' |')
            out.append('<</TABLE>>')
        elif c.tag==W+'sdt':
            body(c.find(W+'sdtContent') or [],out)
def docx(path):
    z=zipfile.ZipFile(path); out=[]
    for n in sorted(z.namelist()):
        if re.match(r'word/header\d*\.xml',n):
            out.append('== '+n); body(ET.fromstring(z.read(n)),out)
    out.append('== BODY')
    body(ET.fromstring(z.read('word/document.xml')).find(W+'body'),out)
    for n in z.namelist():
        if n in('word/comments.xml',):
            out.append('== COMMENTS'); body(ET.fromstring(z.read(n)),out)
    imgs=[n for n in z.namelist() if n.startswith('word/media/')]
    out.append('== MEDIA: '+', '.join(f'{n}({z.getinfo(n).file_size})' for n in imgs))
    return '\n'.join(out)
S='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
R='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
def xlsx(path):
    z=zipfile.ZipFile(path); out=[]
    ss=[]
    if 'xl/sharedStrings.xml' in z.namelist():
        for si in ET.fromstring(z.read('xl/sharedStrings.xml')).findall(S+'si'):
            ss.append(''.join(t.text or '' for t in si.iter(S+'t')))
    wb=ET.fromstring(z.read('xl/workbook.xml'))
    rels=ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    rmap={r.get('Id'):r.get('Target') for r in rels}
    for sh in wb.find(S+'sheets'):
        tgt=rmap[sh.get(R+'id')]; tgt=tgt.lstrip('/'); 
        if not tgt.startswith('xl/'): tgt='xl/'+tgt
        out.append(f'== SHEET {sh.get("name")} state={sh.get("state")}')
        x=ET.fromstring(z.read(tgt))
        for row in x.iter(S+'row'):
            vals=[]
            for c in row.findall(S+'c'):
                v=c.find(S+'v'); f=c.find(S+'f'); t=c.get('t')
                if t=='s' and v is not None: val=ss[int(v.text)]
                elif t=='inlineStr': val=''.join(tt.text or '' for tt in c.iter(S+'t'))
                else: val=v.text if v is not None else ''
                if f is not None: val=f'{val} {{={f.text}}}'
                if val: vals.append(f'{c.get("r")}:{val}')
            if vals: out.append(' ; '.join(vals))
        dv=x.find(S+'dataValidations')
        if dv is not None:
            for d in dv: out.append('DV '+d.get('sqref','')+' '+''.join(e.text or '' for e in d))
    return '\n'.join(out)
for p in sys.argv[2:]:
    name=os.path.basename(p)
    try:
        txt=docx(p) if p.endswith('.docx') else xlsx(p)
    except Exception as e: txt='ERR '+repr(e)
    open(os.path.join(sys.argv[1],name+'.txt'),'w',encoding='utf-8').write(txt)
    print(name, len(txt))
