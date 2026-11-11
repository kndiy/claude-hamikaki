import sys, json, re, zipfile, unicodedata
from xml.sax.saxutils import escape
sys.stdout.reconfigure(encoding='utf-8')
data=json.load(open(sys.argv[1],encoding='utf-8'))
# fix the one inline "?" answer and stray leading dots
for d in data:
    if not d['q'] and '?' in d['a']:
        q,a=d['a'].rsplit('?',1); d['q'],d['a']=q.strip()+'?',a.strip()
    d['a']=d['a'].lstrip('. ').strip()
TYPO=[('tạo đều kiện','tạo điều kiện'),('Những đều kiện','Những điều kiện'),('nhu cẩu','nhu cầu'),('Kết hơp','Kết hợp'),
      ('cán bô y tế','cán bộ y tế'),('cách ứng sử','cách ứng xử'),('Qúa trình','Quá trình'),('lập lại','lặp lại'),
      ('Khi tập kế hoạch','Khi lập kế hoạch'),('địa Phương','địa phương'),('trên thế giới','trên thế giới')]
def fix(s):
    for a,b in TYPO: s=s.replace(a,b)
    return s
def key(s): return re.sub(r'[^\w]','',unicodedata.normalize('NFC',fix(s)).lower())
merged={}; order=[]
for d in data:
    k=key(d['q']); src=f"Đ{d['set']}-C{d['n']}"
    if k in merged:
        m=merged[k]; m['src'].append(src)
        if key(m['a'])!=key(d['a']): print('ANSWER CONFLICT',m['q'],'|',m['a'],'|',d['a'])
    else:
        merged[k]={'q':fix(d['q']),'a':fix(d['a']),'src':[src]}; order.append(k)
items=[merged[k] for k in order]
dups=[i for i in items if len(i['src'])>1]
print('total in',len(data),'unique',len(items),'dup groups',len(dups))
for i in dups: print(' ', '+'.join(i['src']), i['q'][:70])
# ---- docx ----
def run(t,b=False,sz=None,color=None):
    pr=('<w:b/>' if b else '')+(f'<w:color w:val="{color}"/>' if color else '')+(f'<w:sz w:val="{sz}"/>' if sz else '')
    return f'<w:r><w:rPr>{pr}</w:rPr><w:t xml:space="preserve">{escape(t)}</w:t></w:r>'
def p(runs,jc=None,after=60):
    return f'<w:p><w:pPr>{f"<w:jc w:val=\"{jc}\"/>" if jc else ""}<w:spacing w:after="{after}"/></w:pPr>{runs}</w:p>'
def cell(content,w,shade=None):
    sh=f'<w:shd w:val="clear" w:color="auto" w:fill="{shade}"/>' if shade else ''
    return f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/>{sh}</w:tcPr>{content}</w:tc>'
W=[700,5000,3900]
rows=[f'<w:tr><w:trPr><w:tblHeader/></w:trPr>'+cell(p(run('STT',True),'center',0),W[0],'D9E2F3')+cell(p(run('Câu hỏi',True),'center',0),W[1],'D9E2F3')+cell(p(run('Đáp án',True),'center',0),W[2],'D9E2F3')+'</w:tr>']
for n,i in enumerate(items,1):
    rows.append('<w:tr><w:trPr><w:cantSplit/></w:trPr>'+cell(p(run(str(n)),'center',0),W[0])
      +cell(p(run(i['q']),None,0),W[1])
      +cell(p(run(i['a'],True),None,0),W[2])+'</w:tr>')
border=''.join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="999999"/>' for s in ['top','left','bottom','right','insideH','insideV'])
tbl=f'<w:tbl><w:tblPr><w:tblW w:w="{sum(W)}" w:type="dxa"/><w:tblBorders>{border}</w:tblBorders><w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="40" w:type="dxa"/><w:left w:w="80" w:type="dxa"/><w:bottom w:w="40" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>'+''.join(f'<w:gridCol w:w="{w}"/>' for w in W)+'</w:tblGrid>'+''.join(rows)+'</w:tbl>'
body=(p(run('ĐỀ CƯƠNG ÔN MÔN KNGT - TT-GDSK TRONG THỰC HÀNH ĐD',True,32),'center',80)
     +p(run(f'Bản tổng hợp từ 2 đề (Đề 1: 50 câu, Đề 2: 49 câu), đã bỏ câu trùng: còn {len(items)} câu.',sz=20),'center',160)
     +tbl+'<w:p/>')
doc=f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1000" w:right="1000" w:bottom="1000" w:left="1000" w:header="500" w:footer="500" w:gutter="0"/></w:sectPr></w:body></w:document>'
styles='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman" w:eastAsia="Times New Roman"/><w:sz w:val="24"/><w:lang w:val="vi-VN"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>'
ct='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>'
rels='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
drels='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
with zipfile.ZipFile(sys.argv[2],'w',zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml',ct); z.writestr('_rels/.rels',rels)
    z.writestr('word/document.xml',doc); z.writestr('word/styles.xml',styles); z.writestr('word/_rels/document.xml.rels',drels)
print('wrote',sys.argv[2])
