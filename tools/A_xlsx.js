/* ===== 엑셀(.xlsx) 내보내기: 서식 + 수식 (외부 라이브러리 없음) ===== */
const XLSX_W=(()=>{
  const enc=new TextEncoder();
  const CRC=(()=>{const t=new Uint32Array(256);for(let n=0;n<256;n++){let c=n;for(let k=0;k<8;k++)c=c&1?0xEDB88320^(c>>>1):c>>>1;t[n]=c>>>0;}return t;})();
  const crc32=b=>{let c=0xFFFFFFFF;for(let i=0;i<b.length;i++)c=CRC[(c^b[i])&255]^(c>>>8);return (c^0xFFFFFFFF)>>>0;};
  function zip(files){ // [{name, text}] → Blob (무압축)
    const parts=[],cent=[];let off=0;
    files.forEach(f=>{const name=enc.encode(f.name),data=enc.encode(f.text),crc=crc32(data);
      const h=new DataView(new ArrayBuffer(30));h.setUint32(0,0x04034b50,true);h.setUint16(4,20,true);h.setUint16(6,0x0800,true);h.setUint16(8,0,true);
      h.setUint32(14,crc,true);h.setUint32(18,data.length,true);h.setUint32(22,data.length,true);h.setUint16(26,name.length,true);
      parts.push(new Uint8Array(h.buffer),name,data);
      const c=new DataView(new ArrayBuffer(46));c.setUint32(0,0x02014b50,true);c.setUint16(4,20,true);c.setUint16(6,20,true);c.setUint16(8,0x0800,true);
      c.setUint32(16,crc,true);c.setUint32(20,data.length,true);c.setUint32(24,data.length,true);c.setUint16(28,name.length,true);c.setUint32(42,off,true);
      cent.push(new Uint8Array(c.buffer),name);off+=30+name.length+data.length;});
    const csz=cent.reduce((s,x)=>s+x.length,0);const e=new DataView(new ArrayBuffer(22));e.setUint32(0,0x06054b50,true);
    e.setUint16(8,files.length,true);e.setUint16(10,files.length,true);e.setUint32(12,csz,true);e.setUint32(16,off,true);
    return new Blob([...parts,...cent,new Uint8Array(e.buffer)],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
  }
  const x=s=>String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
  // 서식 등록
  const FILL={hdr:"D9E1F2",prod:"EAF2FB",prodH:"DDEBF7",merch:"FDF1E6",merchH:"FCE4D6",sub:"E2EFDA",grand:"C6E0B4",inp:"FFF2CC",key:"FFE699"};
  const FMT={n0:'#,##0;(#,##0);"-"',n1:'#,##0.0;(#,##0.0);"-"',n2:'#,##0.00;(#,##0.00);"-"',n3:'#,##0.000;(#,##0.000);"-"',pct:'0.0%;(0.0%);"-"'};
  function Styles(){
    const fonts=['<font><sz val="9"/><name val="맑은 고딕"/></font>','<font><b/><sz val="9"/><name val="맑은 고딕"/></font>',
      '<font><b/><sz val="12"/><color rgb="FF1F3864"/><name val="맑은 고딕"/></font>','<font><sz val="9"/><color rgb="FF595959"/><name val="맑은 고딕"/></font>'];
    const fk=Object.keys(FILL), fmtK=Object.keys(FMT), xfs=['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'], reg={};
    const get=(o={})=>{const key=JSON.stringify(o);if(reg[key]!=null)return reg[key];
      const font={n:0,b:1,t:2,g:3}[o.font||"n"], fill=o.fill?2+fk.indexOf(o.fill):0, border=o.border==="none"?0:(o.top?2:1), fmt=o.fmt?164+fmtK.indexOf(o.fmt):0;
      const al=o.al?`<alignment horizontal="${o.al}" vertical="center"${o.wrap?' wrapText="1"':''}/>`:(o.wrap?'<alignment vertical="center" wrapText="1"/>':'');
      xfs.push(`<xf numFmtId="${fmt}" fontId="${font}" fillId="${fill}" borderId="${border}" xfId="0"${fmt?' applyNumberFormat="1"':''}${font?' applyFont="1"':''}${fill?' applyFill="1"':''}${border?' applyBorder="1"':''}${al?' applyAlignment="1">'+al+'</xf>':'/>'}`);
      return reg[key]=xfs.length-1;};
    const xml=()=>`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">`+
      `<numFmts count="${fmtK.length}">${fmtK.map((k,i)=>`<numFmt numFmtId="${164+i}" formatCode="${x(FMT[k])}"/>`).join("")}</numFmts>`+
      `<fonts count="${fonts.length}">${fonts.join("")}</fonts>`+
      `<fills count="${2+fk.length}"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>${fk.map(k=>`<fill><patternFill patternType="solid"><fgColor rgb="FF${FILL[k]}"/><bgColor indexed="64"/></patternFill></fill>`).join("")}</fills>`+
      `<borders count="3"><border><left/><right/><top/><bottom/><diagonal/></border>`+
      `<border><left style="thin"><color rgb="FFBFBFBF"/></left><right style="thin"><color rgb="FFBFBFBF"/></right><top style="thin"><color rgb="FFBFBFBF"/></top><bottom style="thin"><color rgb="FFBFBFBF"/></bottom><diagonal/></border>`+
      `<border><left style="thin"><color rgb="FFBFBFBF"/></left><right style="thin"><color rgb="FFBFBFBF"/></right><top style="medium"><color rgb="FF808080"/></top><bottom style="thin"><color rgb="FFBFBFBF"/></bottom><diagonal/></border></borders>`+
      `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${xfs.length}">${xfs.join("")}</cellXfs>`+
      `<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
    return {get,xml};
  }
  // 시트: cells[r][c] = {v, f, s}  (r,c 1부터)
  function Sheet(name){const S={name,rows:{},merges:[],cols:{},freeze:null,filter:null,heights:{}};
    S.set=(r,c,v,s,f)=>{(S.rows[r]=S.rows[r]||{})[c]={v,s,f};};
    return S;}
  function sheetXml(S){
    const rs=Object.keys(S.rows).map(Number).sort((a,b)=>a-b);
    let maxC=1;rs.forEach(r=>Object.keys(S.rows[r]).forEach(c=>{if(+c>maxC)maxC=+c;}));
    let o=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`;
    o+=`<dimension ref="A1:${colName(maxC)}${rs.length?rs[rs.length-1]:1}"/>`;
    if(S.freeze){const [fr,fc]=S.freeze;o+=`<sheetViews><sheetView workbookViewId="0" zoomScale="85"><pane${fc>1?` xSplit="${fc-1}"`:""}${fr>1?` ySplit="${fr-1}"`:""} topLeftCell="${colName(fc)}${fr}" activePane="bottomRight" state="frozen"/></sheetView></sheetViews>`;}
    else o+=`<sheetViews><sheetView workbookViewId="0" zoomScale="85"/></sheetViews>`;
    o+=`<sheetFormatPr defaultRowHeight="15"/>`;
    const cs=Object.keys(S.cols).map(Number).sort((a,b)=>a-b);
    if(cs.length) o+="<cols>"+cs.map(c=>`<col min="${c}" max="${c}" width="${S.cols[c]}" customWidth="1"/>`).join("")+"</cols>";
    o+="<sheetData>";
    rs.forEach(r=>{const row=S.rows[r];o+=`<row r="${r}"${S.heights[r]?` ht="${S.heights[r]}" customHeight="1"`:""}>`;
      Object.keys(row).map(Number).sort((a,b)=>a-b).forEach(c=>{const {v,s,f}=row[c];const ref=colName(c)+r;const sa=s?` s="${s}"`:"";
        if(f!=null){const num=typeof v==="number"&&isFinite(v);o+=`<c r="${ref}"${sa}${num?"":(v==null||v===""?"":' t="str"')}><f>${x(f)}</f>${num?`<v>${v}</v>`:(v==null||v===""?"":`<v>${x(v)}</v>`)}</c>`;}
        else if(typeof v==="number"&&isFinite(v)) o+=`<c r="${ref}"${sa}><v>${v}</v></c>`;
        else if(v==null||v==="") o+=`<c r="${ref}"${sa}/>`;
        else o+=`<c r="${ref}"${sa} t="inlineStr"><is><t xml:space="preserve">${x(v)}</t></is></c>`;});
      o+="</row>";});
    o+="</sheetData>";
    if(S.filter) o+=`<autoFilter ref="${S.filter}"/>`;
    if(S.merges.length) o+=`<mergeCells count="${S.merges.length}">${S.merges.map(m=>`<mergeCell ref="${m}"/>`).join("")}</mergeCells>`;
    o+=`<pageMargins left="0.5" right="0.5" top="0.5" bottom="0.5" header="0.3" footer="0.3"/></worksheet>`;
    return o;
  }
  function book(sheets,styles){
    const files=[];
    files.push({name:"[Content_Types].xml",text:`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${sheets.map((s,i)=>`<Override PartName="/xl/worksheets/sheet${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("")}</Types>`});
    files.push({name:"_rels/.rels",text:`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`});
    const dn=sheets.map((s,i)=>s.filter?`<definedName name="_xlnm._FilterDatabase" localSheetId="${i}" hidden="1">'${s.name}'!$${s.filter.replace(/([A-Z]+)(\d+)/g,"$1$$$2").replace(":",":$")}</definedName>`:"").join("");
    files.push({name:"xl/workbook.xml",text:`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets>${sheets.map((s,i)=>`<sheet name="${x(s.name)}" sheetId="${i+1}" r:id="rId${i+1}"/>`).join("")}</sheets>${dn?`<definedNames>${dn}</definedNames>`:""}<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>`});
    files.push({name:"xl/_rels/workbook.xml.rels",text:`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((s,i)=>`<Relationship Id="rId${i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i+1}.xml"/>`).join("")}<Relationship Id="rId${sheets.length+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`});
    files.push({name:"xl/styles.xml",text:styles.xml()});
    sheets.forEach((s,i)=>files.push({name:`xl/worksheets/sheet${i+1}.xml`,text:sheetXml(s)}));
    return zip(files);
  }
  return {Styles,Sheet,book};
})();

function buildWorkbook(){
  const R=RES, ST=XLSX_W.Styles(), st=ST.get, C=colName;
  const H=st({font:"b",fill:"hdr",al:"center",wrap:1}), HP=st({font:"b",fill:"prodH",al:"center",wrap:1}), HM=st({font:"b",fill:"merchH",al:"center",wrap:1});
  const T=st({font:"t",border:"none"}), NOTE=st({font:"g",border:"none"}), TX=st({}), TXB=st({font:"b"});
  const NUM=(fmt,o={})=>st({fmt,...o});
  const sheets=[];
  // ---------- 1) 적용동가 ----------
  const A=XLSX_W.Sheet("적용동가"); sheets.push(A);
  A.set(1,1,"적용동가 계산 (26년 11월 ~ 27년 12월)",T); A.set(2,1,"목표 프리미엄 ($/톤, 연간)",TXB); A.set(2,2,B.target,NUM("n0",{fill:"inp"}));
  A.set(4,1,"항목",H);A.set(4,2,"단위",H); M14.forEach((m,i)=>A.set(4,3+i,m,H));
  const AR={}; const arows=[["t","제품 총동량 (생산계획 자체)","톤","v",x=>x.t,"n1"],["sq","SCR 선물잔량","톤","v",x=>x.sq,"n1"],["sl","SCR 선물LME","$/톤","v",x=>x.sl,"n0"],
    ["jq","JCR 선물잔량","톤","v",x=>x.jq,"n1"],["jl","JCR 선물LME","$/톤","v",x=>x.jl,"n0"],["nq","비선물 동량","톤","f",null,"n1"],["nl","비선물 LME","$/톤","v",x=>x.nl,"n0"],
    ["al","평균 LME (프리미엄 제외)","$/톤","f",null,"n1"],["tp","계획 적용 프리미엄 (목표)","$/톤","f",null,"n0"],["ll","평균 LME (프리미엄 포함)","$/톤","f",null,"n1"],
    ["fx","환율","원/$","v",x=>x.fx,"n2"],["bp","매입동가","원/kg","f",null,"n1"],["ap","적용동가 (= 전월 매입동가)","원/kg","f",null,"n1"]];
  arows.forEach(([k],i)=>AR[k]=5+i);
  arows.forEach(([k,lab,u,kind,fn,fmt])=>{const r=AR[k];A.set(r,1,lab,k==="bp"||k==="ap"?TXB:TX);A.set(r,2,u,TX);
    R.cu.forEach((cx,i)=>{const c=3+i,L=C(c),ref=kk=>L+AR[kk];let f=null,v=null;
      if(kind==="v"){v=fn(cx);A.set(r,c,v==null?"":v,NUM(fmt,{fill:"inp"}));return;}
      if(k==="nq"){f=`${ref("t")}-${ref("sq")}-${ref("jq")}`;v=cx.nq;}
      if(k==="al"){f=`IFERROR((${ref("sq")}*${ref("sl")}+${ref("jq")}*${ref("jl")}+${ref("nq")}*${ref("nl")})/${ref("t")},0)`;v=cx.al||0;}
      if(k==="tp"){f="$B$2";v=cx.tp;}
      if(k==="ll"){f=`${ref("al")}+${ref("tp")}`;v=(cx.al||0)+cx.tp;}
      if(k==="bp"){f=`${ref("ll")}*${ref("fx")}/1000`;v=cx.bp||0;}
      if(k==="ap"){ if(i===0){A.set(r,c,"",NUM(fmt,{fill:"key",font:"b"}));return;} f=`${C(c-1)}${AR.bp}`;v=R.cu[i-1].bp||0;}
      A.set(r,c,v,NUM(fmt,{fill:(k==="bp"||k==="ap")?"key":undefined,font:(k==="bp"||k==="ap")?"b":"n"}),f);});});
  A.cols={1:28,2:7}; for(let i=0;i<14;i++)A.cols[3+i]=11; A.freeze=[5,3];
  A.set(AR.ap+2,1,"※ 노란 칸 = 입력값(HTML에 붙여넣은 값), 나머지는 수식. 상세 시트의 적용동가는 이 시트 적용동가 행을 참조합니다.",NOTE);
  const apRef=m=>`'적용동가'!$${C(3+m+2)}$${AR.ap}`;   // 27년 m월 → 적용동가 시트 열: 26.11(C), 26.12(D), 1월(E)...
  // ---------- 2) 상세 (수식) ----------
  const D=XLSX_W.Sheet("상세");
  const FR=[["엑셀행",null],...FRONT.map(f=>[f[1],f[0]])]; // A~R
  const cS=FR.length+1, cT=cS+1, cU=cS+2, MB=cS+3, MWD=16;   // 계수 세선(S), 태선(T), 상품재료비율(U), 월 블록 시작
  const DK=["동량","매출액","세선비중","태선비중","상품동량","상품매출","제품동량","제품매출","적용동가","부재료비단가","동재료비","제품부재료비","상품재료비","부재료비","재료비합","재료비율"];
  const mc=(m,k)=>MB+m*MWD+DK.indexOf(k);
  const YK=["동량","매출액","상품동량","상품매출","제품동량","제품매출","동재료비","제품부재료비","상품재료비","부재료비","재료비합","재료비율"];
  const YB=MB+12*MWD, yc=k=>YB+YK.indexOf(k);
  const r0=5, rN=r0+R.rows.length-1;
  D.set(1,1,"재료비 상세 (선박선, 영업 매출계획 행 기준) – 노란 칸 입력값, 나머지 수식",T);
  FR.forEach((f,i)=>{D.set(3,1+i,"",H);D.set(4,1+i,f[0],H);D.merges.push(`${C(1+i)}3:${C(1+i)}4`);});
  [[cS,"세선 부재료비 계수"],[cT,"태선 부재료비 계수"],[cU,"상품재료비율"]].forEach(([c,l])=>{D.set(3,c,"",H);D.set(4,c,l,H);D.merges.push(`${C(c)}3:${C(c)}4`);});
  for(let m=0;m<12;m++){D.set(3,MB+m*MWD,(m+1)+"월",H);D.merges.push(`${C(MB+m*MWD)}3:${C(MB+m*MWD+MWD-1)}3`);
    DK.forEach((k,j)=>{D.set(4,MB+m*MWD+j,k,["상품동량","상품매출","상품재료비"].includes(k)?HM:(["제품동량","제품매출","제품부재료비"].includes(k)?HP:H));if(j)D.set(3,MB+m*MWD+j,"",H);});}
  D.set(3,YB,"연간",H);D.merges.push(`${C(YB)}3:${C(YB+YK.length-1)}3`);YK.forEach((k,j)=>{D.set(4,YB+j,k,H);if(j)D.set(3,YB+j,"",H);});
  const rng=c=>`'상세'!$${C(c)}$${r0}:$${C(c)}$${Math.max(rN,r0)}`;
  const PB="'상품배분'";
  R.rows.forEach((rw,i)=>{const r=r0+i, c=rw.row.cells;
    D.set(r,1,rw.row.excelRow,TX); FRONT.forEach((f,j)=>{const v=c[COL[f[0]]];const nv=parseNum(v);D.set(r,2+j,(nv!=null&&!isNaN(nv)&&j>=8)?nv:norm(v),j>=8?NUM(j>=12&&j<=15?"n3":"n2"):TX);});
    D.set(r,cS,rw.cf?rw.cf.s:"",NUM("n3",{fill:"inp"}));D.set(r,cT,rw.cf?rw.cf.t:"",NUM("n3",{fill:"inp"}));
    D.set(r,cU,rw.rat||0,NUM("pct"),`SUMIFS(${PB}!$C:$C,${PB}!$A:$A,$D${r},${PB}!$B:$B,"상품재료비율")`);
    rw.mon.forEach((x,m)=>{const ref=k=>C(mc(m,k))+r, pc=`${PB}!${C(3+m)}:${C(3+m)}`, crit=t=>`${PB}!$A:$A,$D${r},${PB}!$B:$B,"${t}"`;
      D.set(r,mc(m,"동량"),x.q,NUM("n2",{fill:"inp"}));D.set(r,mc(m,"매출액"),x.a,NUM("n1",{fill:"inp"}));
      D.set(r,mc(m,"세선비중"),x.tw,NUM("pct",{fill:"inp"}));D.set(r,mc(m,"태선비중"),x.th,NUM("pct",{fill:"inp"}));
      D.set(r,mc(m,"상품동량"),x.mq,NUM("n2",{fill:"merch"}),`IFERROR(SUMIFS(${pc},${crit("OEM 동량")})*${ref("동량")}/SUMIFS(${pc},${crit("매출 동량")}),0)`);
      D.set(r,mc(m,"상품매출"),x.ma,NUM("n1",{fill:"merch"}),`IFERROR(SUMIFS(${pc},${crit("OEM 금액")})*${ref("매출액")}/SUMIFS(${pc},${crit("매출 금액")}),0)`);
      D.set(r,mc(m,"제품동량"),x.pq,NUM("n2",{fill:"prod"}),`${ref("동량")}-${ref("상품동량")}`);
      D.set(r,mc(m,"제품매출"),x.pa,NUM("n1",{fill:"prod"}),`${ref("매출액")}-${ref("상품매출")}`);
      D.set(r,mc(m,"적용동가"),x.price||0,NUM("n1"),apRef(m));
      D.set(r,mc(m,"부재료비단가"),x.unit,NUM("n3"),`${ref("세선비중")}*$${C(cS)}${r}+${ref("태선비중")}*$${C(cT)}${r}`);
      D.set(r,mc(m,"동재료비"),x.pcu,NUM("n1"),`${ref("제품동량")}*${ref("적용동가")}/1000`);
      D.set(r,mc(m,"제품부재료비"),x.psub,NUM("n1",{fill:"prod"}),`${ref("제품동량")}*${ref("부재료비단가")}`);
      D.set(r,mc(m,"상품재료비"),x.msub,NUM("n1",{fill:"merch"}),`${ref("상품매출")}*$${C(cU)}${r}`);
      D.set(r,mc(m,"부재료비"),x.psub+x.msub,NUM("n1"),`${ref("제품부재료비")}+${ref("상품재료비")}`);
      D.set(r,mc(m,"재료비합"),x.pcu+x.psub+x.msub,NUM("n1",{font:"b"}),`${ref("동재료비")}+${ref("부재료비")}`);
      D.set(r,mc(m,"재료비율"),x.a?(x.pcu+x.psub+x.msub)/x.a:0,NUM("pct"),`IFERROR(${ref("재료비합")}/${ref("매출액")},0)`);});
    const y=yearRow(rw), yv={동량:y.q,매출액:y.a,상품동량:y.mq,상품매출:y.ma,제품동량:y.pq,제품매출:y.pa,동재료비:y.cuT,제품부재료비:y.psub,상품재료비:y.subT-y.psub,부재료비:y.subT,재료비합:y.cuT+y.subT};
    YK.forEach(k=>{const c2=yc(k);if(k==="재료비율"){D.set(r,c2,y.a?(y.cuT+y.subT)/y.a:0,NUM("pct",{fill:"sub"}),`IFERROR(${C(yc("재료비합"))}${r}/${C(yc("매출액"))}${r},0)`);return;}
      D.set(r,c2,yv[k],NUM(["동량","상품동량","제품동량"].includes(k)?"n2":"n1",{fill:"sub"}),MONTHS.map((_,m)=>C(mc(m,k))+r).join("+"));});
  });
  D.cols={1:6,2:7,3:6,4:7,5:9,6:14,7:8,8:16}; for(let c=9;c<=YB+YK.length;c++) if(!D.cols[c]) D.cols[c]=10;
  D.freeze=[5,7]; D.filter=`A4:${C(YB+YK.length-1)}${Math.max(rN,4)}`;
  // ---------- 3) 상품배분 ----------
  const PBs=XLSX_W.Sheet("상품배분");
  PBs.set(1,1,"상품(OEM) 배분 기준 – 제품군별 월 합계",T);
  PBs.set(3,1,"제품군",H);PBs.set(3,2,"구분",H);MONTHS.forEach((m,i)=>PBs.set(3,3+i,m,H));PBs.set(3,15,"연간",H);
  let pr=4; const gl=[...new Set([...Object.values(R.G).map(g=>g.name)])];
  const Pg=P&&!P.error?P.g:{}, ratM=RT&&!RT.error?RT.m:{};
  gl.forEach(g=>{const gk=nk(g), G=R.G[gk], p=Pg[gk];
    [["매출 동량",m=>G.q[m],"n2",k=>`SUMIF(${rng(4)},$A${pr},${rng(mc(k,"동량"))})`],["매출 금액",m=>G.a[m],"n1",k=>`SUMIF(${rng(4)},$A${pr+1},${rng(mc(k,"매출액"))})`],
     ["OEM 동량",m=>p?p.OEM.q[m+2]:0,"n2",null],["OEM 금액",m=>p?p.OEM.a[m+2]:0,"n1",null]].forEach(([lab,fn,fmt,ff],j)=>{const r=pr+j;
      PBs.set(r,1,g,j===0?TXB:TX);PBs.set(r,2,lab,TX);
      MONTHS.forEach((_,m)=>PBs.set(r,3+m,fn(m),NUM(fmt,ff?{}:{fill:"inp"}),ff?ff(m):null));
      PBs.set(r,15,sumA(MONTHS.map((_,m)=>fn(m))),NUM(fmt,{fill:"sub"}),`SUM(C${r}:N${r})`);});
    PBs.set(pr+4,1,g,TX);PBs.set(pr+4,2,"상품재료비율",TX);PBs.set(pr+4,3,ratM[gk]==null?"":ratM[gk],NUM("pct",{fill:"inp"}));
    pr+=6;});
  PBs.cols={1:10,2:14};for(let c=3;c<=15;c++)PBs.cols[c]=10;PBs.freeze=[4,3];
  PBs.set(pr,1,"※ 매출 동량·금액 = 상세 시트 SUMIF(수식), OEM 동량·금액 = 생산계획 생산/OEM(입력값), 상품재료비율 = 재료비 기준(입력값)",NOTE);
  // ---------- 4) 보고양식 (SUMIFS 수식) ----------
  const Rp=XLSX_W.Sheet("보고양식"); const rep=buildReport(R);
  Rp.set(1,1,"연간 매출 / 재료비 TOTAL"+(state.verName?" – "+state.verName:""),T); Rp.set(2,1,"(단위: 금액·재료비 백만원 / 동량 톤 / 톤당단가 백만원/톤)",NOTE);
  const SEC=[["T","매출 전체",H],["P","제품 (동재료비·부재료비)",HP],["M","상품 (상품재료비율 → 부재료비)",HM]], W=13;
  ["제품군","상세분류","거래처","부서","구분"].forEach((l,i)=>{Rp.set(3,1+i,l,H);Rp.set(4,1+i,"",H);Rp.merges.push(`${C(1+i)}3:${C(1+i)}4`);});
  SEC.forEach(([k,l,hs],si)=>{const c0=6+si*W;Rp.set(3,c0,l,hs);for(let j=1;j<W;j++)Rp.set(3,c0+j,"",hs);Rp.merges.push(`${C(c0)}3:${C(c0+W-1)}3`);
    MONTHS.forEach((m,i)=>Rp.set(4,c0+i,m,hs));Rp.set(4,c0+12,"연간",hs);});
  // 상세 열 매핑: 영역별 금액/동량/동재료비/부재료비
  const SRC={T:{a:"매출액",q:"동량",cu:"동재료비",sub:"부재료비"},P:{a:"제품매출",q:"제품동량",cu:"동재료비",sub:"제품부재료비"},M:{a:"상품매출",q:"상품동량",cu:null,sub:"상품재료비"}};
  let rr=5;
  rep.forEach(g=>{const lvl=g.type, crit=[];
    if(lvl==="cust") crit.push([4,g.k[0]],[5,g.k[1]],[6,g.k[2]],[3,g.k[3]]);
    else if(lvl==="sub") crit.push([4,g.k[0]],[5,g.k[1].replace(/ 계$/,"")]);
    else if(lvl==="gsub") crit.push([4,g.k[0].replace(/ 계$/,"")]);
    const fill=lvl==="cust"?undefined:(lvl==="grand"?"grand":"sub"), font=lvl==="cust"?"n":"b";
    const base=rr;
    LINES.forEach(([lab,key],li)=>{const r=base+li, top=li===0;
      Rp.set(r,1,li===0?g.k[0]:"",st({fill,font,top}));Rp.set(r,2,li===0?g.k[1]:"",st({fill,font,top}));
      Rp.set(r,3,lvl==="cust"?g.k[2]:(li===0?g.k[2]:""),st({fill,font,top}));Rp.set(r,4,lvl==="cust"?g.k[3]:"",st({fill,font,top}));Rp.set(r,5,lab,st({fill,font,top}));
      SEC.forEach(([sk],si)=>{const c0=6+si*W, b=g[sk];
        for(let m=0;m<=12;m++){const c=c0+m, L=C(c), yr=m===12, v=lineVal(b,key,yr?"Y":m);
          const fmt=key==="rate"?"pct":(key==="q"||key==="up"?"n2":"n1"), sfill=fill||(sk==="P"?"prod":(sk==="M"?"merch":undefined));
          let f=null;
          if(key==="up") f=`IFERROR(${L}${base}/${L}${base+1},0)`;
          else if(key==="sum") f=`${L}${base+3}+${L}${base+4}`;
          else if(key==="rate") f=`IFERROR(${L}${base+5}/${L}${base},0)`;
          else if(yr) f=`SUM(${C(c0)}${r}:${C(c0+11)}${r})`;
          else {const srcK=SRC[sk][key==="a"?"a":(key==="q"?"q":(key==="cu"?"cu":"sub"))];
            if(!srcK) f="0"; else {const sc=rng(mc(m,srcK)); f=crit.length?`SUMIFS(${sc},${crit.map(([cc,val])=>`${rng(cc)},"${String(val).replace(/"/g,'""')}"`).join(",")})`:`SUM(${sc})`;}}
          Rp.set(r,c,v,NUM(fmt,{fill:sfill,font,top}),f);}});});
    rr+=LINES.length;});
  Rp.cols={1:9,2:10,3:14,4:6,5:9};for(let c=6;c<6+3*W;c++)Rp.cols[c]=10;
  Rp.freeze=[5,6]; Rp.filter=`A4:${C(5+3*W)}${rr-1}`;
  Rp.set(rr+1,1,"※ 금액·동량·동재료비·부재료비 = 상세 시트 SUMIFS(제품군·상세·거래처·부서), 톤당단가·재료비합·재료비율·연간 = 수식",NOTE);
  sheets.unshift(Rp); sheets.push(D,PBs);
  // 시트 순서: 보고양식, 적용동가, 상세, 상품배분
  return XLSX_W.book(sheets,ST);
}
function htmlReport(){ // 서식 유지 복사용 HTML (값)
  const ms=[...MONTHS.map((_,m)=>m),"Y"], ml=m=>m==="Y"?"연간":(m+1)+"월";
  const css={hdr:"#D9E1F2",P:"#EAF2FB",M:"#FDF1E6",PH:"#DDEBF7",MH:"#FCE4D6",sub:"#E2EFDA",grand:"#C6E0B4"};
  const bd="border:.5pt solid #BFBFBF;font-family:'맑은 고딕';font-size:9pt;";
  const nf=key=>key==="rate"?"0.0%":(key==="q"||key==="up"?"#,##0.00":"#,##0.0");
  let h=`<table style="border-collapse:collapse"><tr>${["제품군","상세분류","거래처","부서","구분"].map(l=>`<th rowspan="2" style="${bd}background:${css.hdr}">${l}</th>`).join("")}`+
    `<th colspan="13" style="${bd}background:${css.hdr}">매출 전체</th><th colspan="13" style="${bd}background:${css.PH}">제품 (동재료비·부재료비)</th><th colspan="13" style="${bd}background:${css.MH}">상품 (상품재료비율 → 부재료비)</th></tr><tr>`+
    [css.hdr,css.PH,css.MH].map(bg=>ms.map(m=>`<th style="${bd}background:${bg}">${ml(m)}</th>`).join("")).join("")+"</tr>";
  REP.forEach(x=>LINES.forEach(([lab,key],li)=>{const bg=x.type==="cust"?"":(x.type==="grand"?css.grand:css.sub), fw=x.type==="cust"?"":"font-weight:bold;", top=li===0?"border-top:1.5pt solid #808080;":"";
    const td=(v,b2)=>`<td style="${bd}${top}${fw}background:${bg||b2||"#FFFFFF"}">${esc(v)}</td>`;
    const tn=(v,b2)=>`<td style="${bd}${top}${fw}background:${bg||b2||"#FFFFFF"};mso-number-format:'${nf(key)}';text-align:right">${isFinite(v)?Math.round(v*1e6)/1e6:""}</td>`;
    h+="<tr>"+td(li===0?x.k[0]:"")+td(li===0?x.k[1]:"")+td(x.type==="cust"?x.k[2]:(li===0?x.k[2]:""))+td(x.type==="cust"?x.k[3]:"")+td(lab)+
      ms.map(m=>tn(lineVal(x.T,key,m))).join("")+ms.map(m=>tn(lineVal(x.P,key,m),css.P)).join("")+ms.map(m=>tn(lineVal(x.M,key,m),css.M)).join("")+"</tr>";}));
  return h+"</table>";
}
async function copyHtml(html,text){
  try{ if(window.ClipboardItem&&navigator.clipboard&&navigator.clipboard.write){
      await navigator.clipboard.write([new ClipboardItem({"text/html":new Blob([html],{type:"text/html"}),"text/plain":new Blob([text],{type:"text/plain"})})]);
      toast("서식 포함 복사 완료 – 엑셀에 붙여넣으세요");return;} }catch(e){}
  const d=document.createElement("div");d.contentEditable="true";d.style.cssText="position:fixed;left:-9999px;top:0";d.innerHTML=html;document.body.appendChild(d);
  const rg=document.createRange();rg.selectNodeContents(d);const s=getSelection();s.removeAllRanges();s.addRange(rg);document.execCommand("copy");s.removeAllRanges();d.remove();
  toast("서식 포함 복사 완료 – 엑셀에 붙여넣으세요");
}
$("xlsxRep").onclick=$("xlsxDet").onclick=()=>{if(!needOk())return;
  saveBlob(buildWorkbook(),"재료비_A단계_결과"+(state.verName?"_"+state.verName.replace(/[\\/:*?"<>|]/g,""):"")+".xlsx");toast("엑셀 파일을 내려받았습니다 (서식 + 수식)");};
$("cpRep").onclick=()=>{if(!needOk())return;copyHtml(htmlReport(),reportRows().map(r=>r.map(tsvCell).join("\t")).join("\n"));};
