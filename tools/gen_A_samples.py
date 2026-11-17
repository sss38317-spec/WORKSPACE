import json, openpyxl, math
S='/tmp/claude-0/-home-user-WORKSPACE/a6bca765-432b-53ba-b314-c602b6c54736/scratchpad'
hdr=json.load(open(f'{S}/sample_hdr.json')); rows=json.load(open(f'{S}/sample_rows.json'))
num=lambda v: float(v) if v not in ("",None) else 0.0
# ---- 매출 (선박선) 제품군별 월별 합계
G={}
sales=[]
for r in rows:
    g=r[3]; d=dict(biz=r[1],g=g,det=r[4],cust=r[5],cur=r[12],dept=r[2])
    d['q']=[num(r[19+16*m]) for m in range(12)]; d['a']=[num(r[23+16*m]) for m in range(12)]
    d['tw']=[num(r[21+16*m]) for m in range(12)]; d['th']=[num(r[22+16*m]) for m in range(12)]
    sales.append(d)
    if d['biz']!='선박선': continue
    x=G.setdefault(g,{'q':[0]*12,'a':[0]*12})
    for m in range(12): x['q'][m]+=d['q'][m]; x['a'][m]+=d['a'][m]
OEMR={'상선용':0.10,'해양용':0.05}
P={}
for g,x in G.items():
    oq=[round(v*OEMR.get(g,0),3) for v in x['q']]; oa=[round(v*OEMR.get(g,0),3) for v in x['a']]
    P[g]={'자체':{'q':[round(x['q'][m]-oq[m],3) for m in range(12)],'a':[round(x['a'][m]-oa[m],3) for m in range(12)]},'OEM':{'q':oq,'a':oa}}
pre={'상선용':(60,65,1800,1950),'해양용':(5,6,150,180)}  # 26.11/26.12 자체 동량·금액
# ---- 생산계획 TSV (A~X)
def prow(b,c,d,e,f,g11,g12,months):
    tot=(f or 0)+(g11 or 0)+(g12 or 0)
    return ["",b,c,d,e,f,g11,g12,tot]+months+[sum(months),"",""]
H2=["","생산\n구분","사업장","제 품 분 류","","2026년 예상 실적","","",""]+[f"{m}월" for m in range(1,13)]+["년간계","톤당단가",""]
H3=["","","","","","1~10월실적","11월 예상","12월 예상","합 계"]+[22,19,24,23,21,24,24,22,21,22,24,24]+["","26년","27년"]
prod=[H2,H3]
first=True
for g in ('상선용','해양용'):
    q11,q12,a11,a12=pre[g]
    prod.append(prow("자체" if first else "", "입 장" if first else "", g,"동량",500,q11,q12,P[g]['자체']['q']))
    prod.append(prow("","","","금액",15000,a11,a12,P[g]['자체']['a'])); first=False
for g in ('이동용','전력용'):
    prod.append(prow("","",g,"동량",0,0,0,[0]*12)); prod.append(prow("","","","금액",0,0,0,[0]*12))
sq=[sum(P[g]['자체']['q'][m] for g in P) for m in range(12)]; sa=[sum(P[g]['자체']['a'][m] for g in P) for m in range(12)]
prod.append(prow("","","입장 계","동량",1000,pre['상선용'][0]+pre['해양용'][0],pre['상선용'][1]+pre['해양용'][1],sq))
prod.append(prow("","","","금액",30000,pre['상선용'][2]+pre['해양용'][2],pre['상선용'][3]+pre['해양용'][3],sa))
prod.append(prow("","북  면","상선용","동량",0,0,0,[0]*12)); prod.append(prow("","","","금액",0,0,0,[0]*12))
prod.append(prow("","","북면 계","동량",0,0,0,[0]*12)); prod.append(prow("","","","금액",0,0,0,[0]*12))
prod.append(prow("","소  계","","동량",1000,0,0,sq)); prod.append(prow("","","","금액",30000,0,0,sa))
first=True
for g in ('상선용','해양용'):
    prod.append(prow("생산\nOEM" if first else "","입장\n북면" if first else "",g,"동량",50,0,0,P[g]['OEM']['q']))
    prod.append(prow("","","","금액",1500,0,0,P[g]['OEM']['a'])); first=False
prod.append(prow("","소계","","동량",0,0,0,[0]*12)); prod.append(prow("","","","금액",0,0,0,[0]*12))
prod.append(prow("합  계","","","동량",0,0,0,[0]*12)); prod.append(prow("","","","금액",0,0,0,[0]*12))
tot14=[pre['상선용'][0]+pre['해양용'][0],pre['상선용'][1]+pre['해양용'][1]]+sq
# ---- 재료비 기준 (재료비동가) TSV: 2행(목표) + 4행~
T=500
sqv=[round(t*0.4,3) for t in tot14]; jqv=[round(t*0.1,3) for t in tot14]
sl=[12800,13000,13000,14000,14000,14500,13000,12000,13500,13800,14000,14200,14300,14400]
sp=[520,520,520,510,510,515,512,500,520,520,520,520,520,520]
jl=[12900,13100,12000,14100,13900,14500,13100,12200,13600,13900,14100,14300,14400,14500]
jp=[450,450,450,450,430,430,450,450,450,450,450,450,450,450]
nl=[13200,13300]+[14500]*12; npv=[500,500]+[520]*12; fx=[1380,1380]+[1360]*12
def buy(i):
    t=tot14[i]; nq=t-sqv[i]-jqv[i]
    al=(sqv[i]*sl[i]+jqv[i]*jl[i]+nq*nl[i])/t
    return (al+T)*fx[i]/1000
bp=[buy(i) for i in range(14)]
MON=["26년 11월","26년 12월"]+[f"{m}월" for m in range(1,13)]
bas=[["","목표 프리미엄 ($/톤, 연간 1개)","","",T,"← 계획 적용 프리미엄 (모든 달에 적용)"],
     [""]+[""]*3+["2026년 (예상)","","2027년 (사업계획)"],
     ["","구분","항목","단위"]+MON+["비고"]]
def br(g,it,u,vals): return ["",g,it,u]+["" if v is None else v for v in vals]+[""]
bas+= [br("제품","제품 총동량","톤",tot14), br("SCR","선물잔량","톤",sqv), br("","선물LME (가중평균)","$/톤",sl), br("","프리미엄 (가중평균)","$/톤",sp),
       br("","잔량 단가 (LME+프리미엄)","$/톤",[a+b for a,b in zip(sl,sp)]),
       br("JCR","선물잔량","톤",jqv), br("","선물LME (가중평균)","$/톤",jl), br("","프리미엄 (가중평균)","$/톤",jp),
       br("","잔량 단가 (LME+프리미엄)","$/톤",[a+b for a,b in zip(jl,jp)]),
       br("비선물","비선물 동량","톤",[t-a-b for t,a,b in zip(tot14,sqv,jqv)]), br("","LME","$/톤",nl), br("","프리미엄","$/톤",npv),
       br("","비선물 단가","$/톤",[a+b for a,b in zip(nl,npv)]),
       br("동가","매입동가","원/kg",[round(v,4) for v in bp]), br("","적용동가 (전월 매입동가)","원/kg",[None]+[round(v,4) for v in bp[:-1]])]
bas[-2][2]="매입동가"
# 환율 행 (매입동가 앞)
bas.insert(len(bas)-2, br("","환율","원/$",fx))
# ---- 상품재료비율
# 상품재료비율: 상세제품군별 1~9월 상품 매출·원가 실적 (사용자 제공 표)
RTAB=[("상선용","JIS",12904483,15931311733,13996932140),("","XLPE",1054333,5884064473,5602298999),("해양용","NEK",109326,906656760,713545292)]
RTOT=tuple(sum(r[i] for r in RTAB) for i in (2,3,4))
RATd={}; _g=""
for g_,d_,q_,a_,c_ in RTAB:
    _g=g_ or _g; RATd[(_g,d_)]=c_/a_
RTOTAL=RTOT[2]/RTOT[1]
def ratOf(g,det): return RATd.get((g,det),RTOTAL)
PACKW=0.303089   # 포장비 계수 (백만원/톤) – 사용자 기준양식처럼 전 행 일괄값
# 상품재료비율 (사용자 수정 양식: 3행 머리글, 4행 매출액·재료비 보조 머리글, 5행부터 값)
rat=[["","품목자산분류","상품"],["","* 1~9월 매출 및 원가실적 "],["","","","값"],["","제품군","상세","합계 : 기준단위수량","합계 : 원화판매금액","합계 : 상품매출원가",""]]+[["",g_,d_,q_,a_,c_,f"{c_/a_*100:.1f}%"] for g_,d_,q_,a_,c_ in RTAB]+[["","총합계","",*RTOT,f"{RTOTAL*100:.1f}%"]]
# ---- 동량지수_부재료비 (사용자 수정 양식: B~R 영업 동량지수 구조, T~V 부재료비 계수, W 포장비 계수)
ix=openpyxl.load_workbook(f'{S}/src/sales.xlsx',data_only=False)['동량지수_1~9월 매출품목']
COEF={'JIS':(4.2137,2.6058),'XLPE':(3.9124,2.4381),'NEK':(4.5212,2.9067),'TYPE P':(4.0453,2.5179),'MIL':(5.0318,3.0246)}
coef=[]
for r in range(10,ix.max_row+1):
    row=[]
    for c in range(1,24):
        v=ix.cell(r,c).value if c<=18 else None
        if c==2 and r>=12: v=f"{ix.cell(r,3).value}{ix.cell(r,4).value}{ix.cell(r,5).value}{ix.cell(r,6).value}" if ix.cell(r,3).value else ""
        if isinstance(v,str) and v.startswith("=+"): v=ix[f"{v[2]}{int(v[3:])}"].value if r>=12 else v
        row.append("" if v is None else v)
    if r==10: row[19]="부재료비 계수 (백만원/톤)"; row[22]="포장비 계수"
    if r==11: row[19]="TOTAL"; row[20]="세선"; row[21]="태선"
    if r>=12 and ix.cell(r,3).value:
        cs,ct=COEF.get(str(ix.cell(r,4).value),(4.0123,2.5321)); row[20]=cs; row[21]=ct; row[19]=round((cs+ct)/2,3); row[22]=PACKW
    if any(str(x).strip() for x in row): coef.append(row)
def coefOf(g,det,cust,cur):
    for row in coef[2:]:
        if (row[2],row[3],row[4],row[5])==(g,det,cust,cur): return row[20],row[21]
    return None
# ---- 기대값 (Python 독립 계산)
apply=[bp[m+1] for m in range(12)]
tot={'q':0,'a':0,'cu':0,'sub':0,'pq':0,'mq':0,'ma':0,'msub':0}
for d in sales:
    if d['biz']!='선박선': continue
    cs,ct=coefOf(d['g'],d['det'],d['cust'],d['cur'])
    for m in range(12):
        q,a=d['q'][m],d['a'][m]; g=d['g']
        Gq,Ga=G[g]['q'][m],G[g]['a'][m]
        oq=P[g]['OEM']['q'][m]*q/Gq if Gq else 0; oa=P[g]['OEM']['a'][m]*a/Ga if Ga else 0
        pq=q-oq; unit=d['tw'][m]*cs+d['th'][m]*ct+PACKW
        cu=pq*apply[m]/1000; sub=pq*unit+oa*ratOf(g,d['det'])
        tot['q']+=q; tot['a']+=a; tot['cu']+=cu; tot['sub']+=sub; tot['pq']+=pq; tot['mq']+=oq; tot['ma']+=oa; tot['msub']+=oa*ratOf(g,d['det'])
tsv=lambda rr: "\n".join("\t".join(('"'+str(x).replace('"','""')+'"') if any(ch in str(x) for ch in '\t\n"') else str(x) for x in r) for r in rr)
out={'sales':tsv([*hdr,*rows]),'prod':tsv(prod),'basis':tsv(bas),'ratio':tsv(rat),'coef':tsv(coef)}
json.dump(out,open(f'{S}/A_samples.json','w'),ensure_ascii=False)
json.dump({'tot':tot,'buy':bp,'apply':apply},open(f'{S}/A_expected.json','w'),ensure_ascii=False)
print(json.dumps(tot,ensure_ascii=False)); print([round(x,1) for x in apply])
