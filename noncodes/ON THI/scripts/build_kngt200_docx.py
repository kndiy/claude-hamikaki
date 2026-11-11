import sys, json, zipfile
from xml.sax.saxutils import escape
sys.stdout.reconfigure(encoding='utf-8')
items=json.load(open(sys.argv[1],encoding='utf-8'))
def run(t,b=False,sz=None):
    pr=('<w:b/>' if b else '')+(f'<w:sz w:val="{sz}"/>' if sz else '')
    return f'<w:r><w:rPr>{pr}</w:rPr><w:t xml:space="preserve">{escape(t)}</w:t></w:r>'
def p(runs,jc=None,after=0,before=0,keep=False):
    return f'<w:p><w:pPr>{"<w:keepNext/>" if keep else ""}{f"<w:jc w:val=\"{jc}\"/>" if jc else ""}<w:spacing w:before="{before}" w:after="{after}"/></w:pPr>{runs}</w:p>'
def cell(c,w,shade=None):
    sh=f'<w:shd w:val="clear" w:color="auto" w:fill="{shade}"/>' if shade else ''
    return f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/>{sh}</w:tcPr>{c}</w:tc>'
W=[5500,4100]
border=''.join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="999999"/>' for s in ['top','left','bottom','right','insideH','insideV'])
def table(groups):
    head='<w:tr><w:trPr><w:tblHeader/></w:trPr>'+''.join(cell(p(run(h,True),'center'),w,'D9E2F3') for h,w in zip(['Câu hỏi','Đáp án'],W))+'</w:tr>'
    body=''
    for bai,rows in groups.items():
        if bai: body+=f'<w:tr><w:trPr><w:cantSplit/></w:trPr><w:tc><w:tcPr><w:tcW w:w="{sum(W)}" w:type="dxa"/><w:gridSpan w:val="2"/><w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/></w:tcPr>{p(run(bai,True))}</w:tc></w:tr>'
        body+=''.join('<w:tr><w:trPr><w:cantSplit/></w:trPr>'+cell(p(run(i['q'])),W[0])+cell(p(run(i['a'],True)),W[1])+'</w:tr>' for i in rows)
    return f'<w:tbl><w:tblPr><w:tblW w:w="{sum(W)}" w:type="dxa"/><w:tblBorders>{border}</w:tblBorders><w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="40" w:type="dxa"/><w:left w:w="80" w:type="dxa"/><w:bottom w:w="40" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>'+''.join(f'<w:gridCol w:w="{w}"/>' for w in W)+'</w:tblGrid>'+head+body+'</w:tbl>'
groups={}
for i in items: groups.setdefault(i.get('bai'),[]).append(i)
TITLE=sys.argv[3] if len(sys.argv)>3 else 'TRUYỀN THÔNG – GIÁO DỤC SỨC KHỎE'
SUB=sys.argv[4] if len(sys.argv)>4 else f'Bộ {len(items)} câu hỏi trắc nghiệm có đáp án, chia theo bài'
body=p(run(TITLE,True,32),'center',60)+p(run(SUB,sz=20),'center',120)
body+=table(groups)
body+='<w:p/>'
doc=f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1000" w:right="1000" w:bottom="1000" w:left="1000" w:header="500" w:footer="500" w:gutter="0"/></w:sectPr></w:body></w:document>'
styles='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman" w:eastAsia="Times New Roman"/><w:sz w:val="24"/><w:lang w:val="vi-VN"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>'
ct='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>'
rels='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
drels='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
with zipfile.ZipFile(sys.argv[2],'w',zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml',ct); z.writestr('_rels/.rels',rels)
    z.writestr('word/document.xml',doc); z.writestr('word/styles.xml',styles); z.writestr('word/_rels/document.xml.rels',drels)
print('wrote',sys.argv[2],{k:len(v) for k,v in groups.items()})
