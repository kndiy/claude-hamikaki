"""Parse 200_cau_trac_nghiem_Giai_phau_Sinh_ly.docx text: the correct option is the one followed by an explanation sentence."""
import sys, re, json
sys.stdout.reconfigure(encoding='utf-8')
lines=[l.strip() for l in open(sys.argv[1],encoding='utf-8') if l.strip()]
items=[]; cur=None
for l in lines:
    m=re.match(r'^Câu (\d+)\.\s*(.*)$',l)
    if m: cur={'n':int(m.group(1)),'q':m.group(2),'opts':{}}; items.append(cur); continue
    m=re.match(r'^([A-D])\.\s*(.*)$',l)
    if m and cur: cur['opts'][m.group(1)]=m.group(2)
EXPL=re.compile(r'\.\s+(?=[A-ZĐÂĂÊÔƠƯÁÀẢÃẠẮẰẲẴẶẤẦẨẪẬÉÈẺẼẸẾỀỂỄỆÍÌỈĨỊÓÒỎÕỌỐỒỔỖỘỚỜỞỠỢÚÙỦŨỤỨỪỬỮỰÝ])')
bad=0
for c in items:
    hits=[(k,v) for k,v in c['opts'].items() if EXPL.search(v)]
    if len(hits)!=1 or len(c['opts'])!=4:
        bad+=1; print('CHECK',c['n'],c['opts']); continue
    k,v=hits[0]; a,expl=EXPL.split(v,1)
    c['L'],c['a'],c['expl']=k,a.strip(),expl.strip()
print(len(items),'parsed, bad',bad)
json.dump(items,open(sys.argv[2],'w',encoding='utf-8'),ensure_ascii=False,indent=0)
