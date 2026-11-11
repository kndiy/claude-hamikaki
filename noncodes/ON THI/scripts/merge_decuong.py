import sys, re, json
sys.stdout.reconfigure(encoding='utf-8')
lines=[l.strip() for l in open(sys.argv[1],encoding='utf-8').read().split('\n')]
lines=[l for l in lines if l and l not in ('..','HẾT')][1:]  # drop title
sets=[[]]; 
for l in lines:
    m=re.match(r'^Câu\s*(\d+)\s*[:.]\s*(.*)$',l)
    if m:
        n=int(m.group(1))
        if sets[-1] and n< sets[-1][-1]['n']: sets.append([])
        sets[-1].append({'n':n,'raw':m.group(2).strip(),'ans':[]})
    else: sets[-1][-1]['ans'].append(l)
out=[]
for si,s in enumerate(sets):
    for q in s:
        raw=q['raw']; ans=' '.join(q['ans']).strip()
        if not ans:  # answer inline after ':' 
            i=raw.rfind(':'); raw,ans=raw[:i+1],raw[i+1:]
        out.append({'set':si+1,'n':q['n'],'q':re.sub(r'\s+',' ',raw).strip(),'a':re.sub(r'\s+',' ',ans).strip().rstrip('.')})
json.dump(out,open(sys.argv[2],'w',encoding='utf-8'),ensure_ascii=False,indent=0)
for o in out: print(o['set'],o['n'],'|',o['q'],'=>',o['a'])
