import json,sys
R=json.load(open(sys.argv[1]))
# segment keys: PL par life, PP par pension, PA par annuity, NL nonpar life, NA nonpar annuity, NP nonpar pension, LL linked life, LP linked pension, H health(any), V variable(any)
STD=['LL','LP','H','V',None,'PL','PA','PP','H','V',None,'NL','NA','NP','H','V',None,'GT']
LAY={
 'ACKO_Life_Mar_26':STD,'ICICI_Life_insurance_co_-_q4fy26':STD,'Kotak_mahindra_life_insurance_co_-_q4fy26':STD,'amli-q4fy26':STD,
 'bajaj-life-insurance-q4fy26':STD,'creditAccess-life-insruance-q4fy26':STD,'generali-central-q4fy26':STD,'go-digit-march-2026':STD,
 'hdfc_life_insurance_co_-_q4fy26':STD,'indiaFirst-life-insurance-q4-2025-2026':STD,'tata_aia_-_q4fy26':STD,
 'aegas-federal-life-q4fy26':['LL','LP',None,'PL',None,'NL','NA','NP','H','V',None,'GT'],
 'aviva-life-q4-website-pd':['LL','LP','PL','PP','NL','NP','NA','H','V','GT'],
 'bandhan-q4fy26':['LL','LP',None,'PL','PP',None,'NL','NA','H',None,'GT'],
 'bharti-life-q4fy26':['PL','PP','NL','NP','H','NL*G','LL','LP','LL*G','GT'],
 'canara-hsbc-q4fy26':['LL','LP',None,'PL',None,'NL','NA','NP','H','V','V',None,'GT'],
 'sriram-life_q4_2026':['LL','LP','PL','NL','V','H','NA','GT'],
 'star-union-daiichi-life-insurance':['LL','LP','LL*G',None,'PL',None,'NL*G','NA*G','NL','NA','NP','H',None,'GT'],
 '1782ec5f-pramerica-life-insurance-q4fy26':STD,
 '3ccf3c03-inlic-q4fy26':['LL','LP','LL*G',None,'PL','PP',None,'NL','V','NA','H','NL*G','V',None,'GT'],
 '69e944fe-ABSLI_-_q4fy26':['LL','LL*G','LP','LP*G','H','NL','NL*G','V','NP','NA','NA*G','NP*G','V','H','PL','GT'],
 'pnb':STD,
 'sbi':['LL','LP',None,'PL','PP','V',None,'NL','NA','NP','H','V',None,'GT'],
}
# Edelweiss handled manually (Total sub-row of L-4 premium): Total|IndParLife|IndParPension|IndNPLife|IndNPAnnuity|Health|GrpNPLife|GrpVarLife|IndLinkedLife|IndLinkedPension|GrpLinkedLife
EDEL_COLS=['GT','PL','PP','NL','NA','H','NL*G','V','LL','LP','LL*G']
EDEL={26:[223565,69043,95,107679,3825,155,501,0,40425,426,1417],25:[209299,65882,118,102338,1729,160,1004,0,37052,370,647]}
NAME={'ACKO_Life_Mar_26':'Acko Life','ICICI_Life_insurance_co_-_q4fy26':'ICICI Prudential Life','Kotak_mahindra_life_insurance_co_-_q4fy26':'Kotak Mahindra Life',
'amli-q4fy26':'Axis Max Life','bajaj-life-insurance-q4fy26':'Bajaj Life','creditAccess-life-insruance-q4fy26':'CreditAccess Life',
'generali-central-q4fy26':'Generali Central Life','go-digit-march-2026':'Go Digit Life','hdfc_life_insurance_co_-_q4fy26':'HDFC Life',
'indiaFirst-life-insurance-q4-2025-2026':'IndiaFirst Life','tata_aia_-_q4fy26':'Tata AIA Life','aegas-federal-life-q4fy26':'Ageas Federal Life',
'aviva-life-q4-website-pd':'Aviva Life','bandhan-q4fy26':'Bandhan Life','bharti-life-q4fy26':'Bharti Life','canara-hsbc-q4fy26':'Canara HSBC Life',
'sriram-life_q4_2026':'Shriram Life','star-union-daiichi-life-insurance':'Star Union Dai-ichi Life','edelweiss-life-insurance':'Edelweiss Life','1782ec5f-pramerica-life-insurance-q4fy26':'Pramerica Life','3ccf3c03-inlic-q4fy26':'IndusInd Nippon Life','69e944fe-ABSLI_-_q4fy26':'Aditya Birla Sun Life','sbi':'SBI Life','pnb':'PNB MetLife'}
# Group total premium FY26 from L-25(ii) YTD grand total (Rs lakh)
GRP={'ACKO_Life_Mar_26':18066,'ICICI_Life_insurance_co_-_q4fy26':1560887.24,'Kotak_mahindra_life_insurance_co_-_q4fy26':596678.01,'amli-q4fy26':241231,
'bajaj-life-insurance-q4fy26':622244,'creditAccess-life-insruance-q4fy26':62558,'generali-central-q4fy26':61784,'go-digit-march-2026':132988.2,
'hdfc_life_insurance_co_-_q4fy26':1844815,'indiaFirst-life-insurance-q4-2025-2026':161428.57,'tata_aia_-_q4fy26':178092,'aegas-federal-life-q4fy26':50589,
'aviva-life-q4-website-pd':24946,'bandhan-q4fy26':7221,'bharti-life-q4fy26':13138,'canara-hsbc-q4fy26':124826,'sriram-life_q4_2026':91120,
'star-union-daiichi-life-insurance':348751.80,'edelweiss-life-insurance':1917.14,'1782ec5f-pramerica-life-insurance-q4fy26':114058,'3ccf3c03-inlic-q4fy26':23521,'69e944fe-ABSLI_-_q4fy26':783002.42,'sbi':1520192,'pnb':279231}
def agg(vals,lay):
    d={k:0.0 for k in ['PL','PP','PA','NL','NA','NP','LL','LP','H','V','GT']}
    assert len(vals)==len(lay),(len(vals),len(lay))
    for v,k in zip(vals,lay):
        if k: d[k.split('*')[0]]+=v
    return d
out={}
for f,lay in list(LAY.items())+[('edelweiss-life-insurance',EDEL_COLS)]:
    if f=='edelweiss-life-insurance': c,p=EDEL[26],EDEL[25]
    else:
        c,p=R[f]
        if f=='aviva-life-q4-website-pd':
            c=c[:9]+[c[10]]; p=p[:9]+[p[10]]
        if f=='ACKO_Life_Mar_26': c=[abs(x) for x in c]; p=[abs(x) for x in p]
    d26=agg(c,lay); d25=agg(p,lay) if p else None
    out[NAME[f]]={'file':f,'fy26':d26,'fy25':d25,'group_fy26':GRP[f],'layout':lay}
# LIC: L-1-A-RA in Rs crore (no FY25 segment split supplied) -> lakh. CRAC 0.85 cr put with excluded 'V'.
L=[20835.67,2103.27,56.03,0,None,283265.35,1.67,319.48,0,0,None,107426.47,15825.84,106654.25,257.10,2.58,None,0.85,536748.56]
lay=STD[:17]+['V','GT']
d=agg([0 if x is None else x*100 for x in L],lay)
out['LIC']={'file':'4194e382-LIC-1A-_Revenue_Account_for_the_period_ended_31.03.2026 (Rs crore x100)','fy26':d,'fy25':None,'group_fy26':19660900,'layout':lay,
 'group_note':'LIC press release 21 May 2026: Total Group Business Premium Income Rs 1,96,609 cr (L-25(ii) not supplied)'}
json.dump(out,open(sys.argv[2],'w'),indent=1)
for n,o in sorted(out.items(),key=lambda x:-x[1]['fy26']['GT']):
    d=o['fy26']; s=sum(v for k,v in d.items() if k!='GT')
    print(f"{n:26} GT={d['GT']/100:9.0f}cr sumseg={s/100:9.0f} diff={(s-d['GT']):8.1f}L  grp={o['group_fy26']/100:7.0f}cr ({o['group_fy26']/d['GT']*100:4.0f}%)  fy25GT={(o['fy25'] or {}).get('GT',0)/100:7.0f}")
