"""Minimal Markdown -> .docx for study notes: #/## headings, bullets, numbered lists (nested), > notes, | tables |, **bold**."""
import sys, re, zipfile
from xml.sax.saxutils import escape
sys.stdout.reconfigure(encoding='utf-8')
src, dst = sys.argv[1], sys.argv[2]
CONTINUOUS = '--continuous' in sys.argv[3:]  # no page break before each '# ' section
FONT='Times New Roman'
def runs(text, size=None, color=None, bold_all=False):
    out=[]
    for i,part in enumerate(re.split(r'\*\*(.+?)\*\*', text)):
        if not part: continue
        b = bold_all or i%2==1
        pr=('<w:b/>' if b else '')+(f'<w:color w:val="{color}"/>' if color else '')+(f'<w:sz w:val="{size}"/>' if size else '')
        out.append(f'<w:r><w:rPr>{pr}</w:rPr><w:t xml:space="preserve">{escape(part)}</w:t></w:r>')
    return ''.join(out)
def para(content, ind=0, hang=0, before=0, after=60, jc=None, keep=False, shade=None, border=None):
    ppr=('<w:keepNext/>' if keep else '')
    if border: ppr+=f'<w:pBdr><w:left w:val="single" w:sz="18" w:space="6" w:color="{border}"/></w:pBdr>'
    if shade: ppr+=f'<w:shd w:val="clear" w:color="auto" w:fill="{shade}"/>'
    ppr+=f'<w:spacing w:before="{before}" w:after="{after}"/>'
    if ind or hang: ppr+=f'<w:ind w:left="{ind}" w:hanging="{hang}"/>'
    if jc: ppr+=f'<w:jc w:val="{jc}"/>'
    return f'<w:p><w:pPr>{ppr}</w:pPr>{content}</w:p>'
border=''.join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="999999"/>' for s in ['top','left','bottom','right','insideH','insideV'])
TOTAL=9900
def table(rows):
    hdr=[c.strip() for c in rows[0].strip('|').split('|')]
    body=[[c.strip() for c in r.strip('|').split('|')] for r in rows[2:]]
    n=len(hdr)
    # width heuristic from content length
    lens=[max(len(hdr[i]),*(len(r[i]) for r in body)) for i in range(n)]
    lens=[min(max(l,6),60) for l in lens]; W=[int(TOTAL*l/sum(lens)) for l in lens]
    def tr(cells,head):
        tcs=''.join(f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/>'+('<w:shd w:val="clear" w:color="auto" w:fill="D9E2F3"/>' if head else '')+f'</w:tcPr>{para(runs(c,bold_all=head,size=22),after=0)}</w:tc>' for c,w in zip(cells,W))
        return f'<w:tr><w:trPr>{"<w:tblHeader/>" if head else "<w:cantSplit/>"}</w:trPr>{tcs}</w:tr>'
    return (f'<w:tbl><w:tblPr><w:tblW w:w="{TOTAL}" w:type="dxa"/><w:tblBorders>{border}</w:tblBorders><w:tblLayout w:type="fixed"/>'
            f'<w:tblCellMar><w:top w:w="30" w:type="dxa"/><w:left w:w="80" w:type="dxa"/><w:bottom w:w="30" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tblCellMar></w:tblPr>'
            '<w:tblGrid>'+''.join(f'<w:gridCol w:w="{w}"/>' for w in W)+'</w:tblGrid>'+tr(hdr,True)+''.join(tr(r,False) for r in body)+'</w:tbl>'+para('',after=60))
lines=open(src,encoding='utf-8').read().split('\n')
out=[]; i=0; first_h1=True; first_part=True
while i<len(lines):
    l=lines[i]
    if l.startswith('|'):
        blk=[]
        while i<len(lines) and lines[i].startswith('|'): blk.append(lines[i]); i+=1
        out.append(table(blk)); continue
    s=l.strip()
    if not s: i+=1; continue
    if l.startswith('# '):
        if first_h1: out.append(para(runs(s[2:],size=32,bold_all=True,color='1F3864'),jc='center',after=120)); first_h1=False
        else:
            pb='' if (first_part or CONTINUOUS) else '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'; first_part=False
            out.append(pb+para(runs(s[2:],size=28,bold_all=True,color='1F3864'),after=160,keep=True,border='1F3864'))
    elif l.startswith('## '):
        out.append(para(runs(s[3:],size=26,bold_all=True,color='2E74B5'),before=200,after=80,keep=True))
    elif l.startswith('> '):
        out.append(para(runs(s[2:],size=22),shade='FFF2CC',border='BF9000',before=60,after=100,ind=120))
    else:
        lead=len(l)-len(l.lstrip(' '))
        lvl=0 if lead==0 else (1 if lead<=3 else 2)
        m=re.match(r'^(-|\d+\.)\s+(.*)$',s)
        if m:
            mark='•' if m.group(1)=='-' else m.group(1)
            if lvl==1 and mark=='•': mark='–'
            if lvl==2 and mark=='•': mark='·'
            ind=360+lvl*400
            out.append(para(runs(mark+'\t'+m.group(2)),ind=ind,hang=300,after=40))
        else:
            out.append(para(runs(s),after=80))
    i+=1
body=''.join(out)
doc=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
     +body+'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1000" w:right="1000" w:bottom="1000" w:left="1000" w:header="500" w:footer="500" w:gutter="0"/></w:sectPr></w:body></w:document>')
styles=f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:cs="{FONT}" w:eastAsia="{FONT}"/><w:sz w:val="24"/><w:lang w:val="vi-VN"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>'
ct='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>'
rels='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
drels='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
with zipfile.ZipFile(dst,'w',zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml',ct); z.writestr('_rels/.rels',rels)
    z.writestr('word/document.xml',doc); z.writestr('word/styles.xml',styles); z.writestr('word/_rels/document.xml.rels',drels)
print('wrote',dst)
