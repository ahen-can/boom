import re,sys,json
TXT=sys.argv[1]
# file -> (FY26 block title line, FY25 block title line)
B={"ACKO_Life_Mar_26":(113,221),"ICICI_Life_insurance_co_-_q4fy26":(141,291),"Kotak_mahindra_life_insurance_co_-_q4fy26":(121,250),
"aegas-federal-life-q4fy26":(105,295),"amli-q4fy26":(144,315),"aviva-life-q4-website-pd":(82,165),"bajaj-life-insurance-q4fy26":(159,337),
"bandhan-q4fy26":(133,271),"bharti-life-q4fy26":(70,237),"canara-hsbc-q4fy26":(72,257),"creditAccess-life-insruance-q4fy26":(64,120),
"edelweiss-life-insurance":(164,373),"generali-central-q4fy26":(18,206),"go-digit-march-2026":(154,304),"hdfc_life_insurance_co_-_q4fy26":(183,395),
"indiaFirst-life-insurance-q4-2025-2026":(131,272),"sriram-life_q4_2026":(126,259),"star-union-daiichi-life-insurance":(211,353),"tata_aia_-_q4fy26":(143,303),"1782ec5f-pramerica-life-insurance-q4fy26":(138,284),"3ccf3c03-inlic-q4fy26":(229,311),"69e944fe-ABSLI_-_q4fy26":(154,341),"sbi":(236,331),"pnb":(213,285)}
tok=re.compile(r'\(\s*-?[\d,]*\.?\d*\s*\)|-?[\d,]+(?:\.\d+)?|(?<!\w)-(?!\w)')
def num(t):
    t=t.strip()
    neg=t.startswith('(')
    t=t.strip('()').strip()
    if t in('-',''): return 0.0
    v=float(t.replace(',',''))
    return -v if neg else v
out={}
for f,(a,b) in B.items():
    L=open(f"{TXT}/{f}.txt").read().split('\n')
    res=[]
    for s in (a,b):
        if s is None: res.append(None);continue
        for i in range(s,s+40):
            if re.search(r'\(a\)\s*Premium',L[i]):
                line=L[i]; j=i
                rest=re.split(r'L-?0?4',line,maxsplit=1)
                body=rest[1] if len(rest)>1 else line.split('Premium',1)[1]
                toks=[num(x) for x in tok.findall(body)]
                if not toks: # values on following line(s)
                    for k in range(i+1,i+4):
                        tt=[num(x) for x in tok.findall(L[k])]
                        if tt: toks=tt;break
                res.append(toks);break
        else: res.append('NOTFOUND')
    out[f]=res
json.dump(out,open(sys.argv[2],'w'))
for f,(c,p) in out.items():
    print(f"{f[:28]:28} n26={len(c) if isinstance(c,list) else c} n25={len(p) if isinstance(p,list) else p} | {c}")
