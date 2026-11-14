"""재료비계획_자동화.xlsx 생성 스크립트 (예시 데이터 포함)."""
import json, sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.comments import Comment

EXT = sys.argv[1] if len(sys.argv) > 1 else None
OUT = sys.argv[2] if len(sys.argv) > 2 else "재료비계획_자동화.xlsx"

F = "맑은 고딕"
f_in = Font(name=F, size=10, color="0000FF")
f_calc = Font(name=F, size=10)
f_link = Font(name=F, size=10, color="008000")
f_b = Font(name=F, size=10, bold=True)
f_h = Font(name=F, size=10, bold=True, color="FFFFFF")
f_t = Font(name=F, size=14, bold=True, color="1F3864")
f_note = Font(name=F, size=9, color="595959")
fill_h = PatternFill("solid", fgColor="1F3864")
fill_h2 = PatternFill("solid", fgColor="D9E1F2")
fill_in = PatternFill("solid", fgColor="FFF2CC")
fill_help = PatternFill("solid", fgColor="F2F2F2")
fill_tot = PatternFill("solid", fgColor="E2EFDA")
thin = Side(style="thin", color="BFBFBF")
bd = Border(left=thin, right=thin, top=thin, bottom=thin)
ctr = Alignment(horizontal="center", vertical="center", wrap_text=True)
NUM = '#,##0;(#,##0);"-"'
NUM1 = '#,##0.0;(#,##0.0);"-"'
NUM2 = '#,##0.00;(#,##0.00);"-"'
PCT = '0.0%;(0.0%);"-"'

NS = 4            # 시나리오 수
NP = 20           # 제품군 최대 수
NR = 1000         # 매출 붙여넣기 최대 행
R0 = 6            # 매출 데이터 시작 행
RE = R0 + NR - 1

def hdr(c, v, fill=fill_h, font=f_h):
    c.value = v; c.fill = fill; c.font = font; c.alignment = ctr; c.border = bd

def cell(ws, ref, v, font=f_calc, fmt=None, fill=None, b=True):
    c = ws[ref]; c.value = v; c.font = font
    if fmt: c.number_format = fmt
    if fill: c.fill = fill
    if b: c.border = bd
    return c

wb = Workbook()

# ------------------------------------------------------------------ 사용안내
g = wb.active; g.title = "사용안내"
g.column_dimensions["A"].width = 2; g.column_dimensions["B"].width = 22; g.column_dimensions["C"].width = 100
g["B2"] = "재료비계획 자동 산출 템플릿"; g["B2"].font = f_t
rows = [
    ("구성", ""),
    ("① 조건설정", "계획 시작월·단위, 시나리오(최대 4개)별 LME·프리미엄·환율(월별), 적용 매출버전, 조정률을 입력합니다."),
    ("② 제품마스터", "제품군명(매출자료의 '제품군'과 똑같이), 기준 동량지수(매출액÷동량), 로스율, 기타재료비율을 입력합니다."),
    ("③ 매출_기준 / 매출_변경", "영업에서 받은 매출계획을 정해진 양식(노란 칸)에 '값 붙여넣기'만 하면 됩니다. 기존 데이터는 지우고 붙여넣으세요."),
    ("④ 계산", "자동 계산 영역(수정 금지). 시나리오×제품군×월별 매출·동량·동재료비·기타재료비·재료비."),
    ("⑤ 비교", "기간(연간/반기/분기/월)을 고르면 시나리오별 결과와 기준(S1) 대비 차이가 한 화면에 나옵니다."),
    ("", ""),
    ("반복 작업 방법", ""),
    ("매출이 바뀌었을 때", "새 매출자료를 매출_변경(또는 매출_기준) 시트에 붙여넣기 → 끝. 상단 '검증' 칸에서 미등록 제품군이 0인지 확인."),
    ("재료비 조건이 바뀌었을 때", "파일을 복사하지 말고 조건설정의 빈 시나리오(S2~S4)에 바뀐 조건만 입력 → 비교 시트에서 바로 차이 확인."),
    ("롤링 계획", "조건설정의 '계획 시작월'을 바꾸면(예: 2026-11) M1~M12 라벨이 자동으로 바뀝니다. 분기/반기는 M1 기준 3·6개월 단위입니다."),
    ("", ""),
    ("계산 로직", ""),
    ("동가(원/kg)", "= (LME $/톤 + 프리미엄 $/톤) × 환율(원/$) ÷ 1,000"),
    ("동량(톤)", "= 매출자료 동량 (비어 있으면 매출액 ÷ 제품마스터 기준 동량지수) × (1+매출조정률) × (1+동량조정률)"),
    ("동재료비(백만원)", "= 동량(톤) × (1+로스율) × 동가(원/kg) ÷ 1,000"),
    ("기타재료비(백만원)", "= 매출액 × 기타재료비율 × (1+기타재료비 증감률)"),
    ("재료비", "= 동재료비 + 기타재료비,  재료비율 = 재료비 ÷ 매출액"),
    ("", ""),
    ("셀 색상 규칙", ""),
    ("노란 칸 · 파란 글씨", "입력 칸 (직접 입력/붙여넣기)"),
    ("흰 칸 · 검정 글씨", "수식 (수정 금지)"),
    ("회색 칸", "매출 시트의 보조 계산 열 (수정 금지)"),
    ("", ""),
    ("예시 데이터", "27년 사업계획 Rev.1109 파일의 '영업팀' 버전(매출_기준)과 '최종' 버전(매출_변경), 27년 LME $14,240 + 프리미엄 $450, 동가 20,046원/kg 기준으로 채워져 있습니다. 실제 사용 시 덮어쓰세요."),
]
r = 4
for k, v in rows:
    g.cell(r, 2, k).font = f_b if v == "" or k.startswith(("①","②","③","④","⑤")) else f_calc
    if v == "" and k: g.cell(r, 2).font = Font(name=F, size=11, bold=True, color="1F3864")
    g.cell(r, 3, ("→ " + v[1:].strip()) if v.startswith("=") else v).font = f_calc; g.cell(r, 3).alignment = Alignment(wrap_text=True, vertical="top")
    r += 1
g["B24"].fill = fill_in; g["B24"].font = f_in

# ------------------------------------------------------------------ 조건설정
c = wb.create_sheet("조건설정")
c.column_dimensions["A"].width = 2; c.column_dimensions["B"].width = 26
for i in range(3, 3 + NS): c.column_dimensions[L(i)].width = 16
c.column_dimensions[L(3 + NS)].width = 50
c["B2"] = "조건설정"; c["B2"].font = f_t
cell(c, "B4", "계획 시작월 (YYYY-MM-01)", f_b)
cell(c, "C4", None, f_in, "yyyy-mm", fill_in)
c["C4"].value = "=DATE(2027,1,1)"; c["C4"].font = f_in
cell(c, "B5", "매출 단위", f_b); cell(c, "C5", "백만원", f_in, None, fill_in)
cell(c, "B6", "동량 단위", f_b); cell(c, "C6", "톤", f_in, None, fill_in)
cell(c, "D4", "← 롤링이면 시작월만 변경", f_note, b=False)

# 시나리오 헤더
hdr(c["B8"], "항목")
for s in range(NS): hdr(c.cell(8, 3 + s), f"S{s+1}")
hdr(c.cell(8, 3 + NS), "설명")
labels = [
    ("시나리오명", ["기준(영업팀 매출)", "수정안(최종 매출)", "수정안+LME 10%↑", "수정안+환율 1,300"], None, "비교 화면에 표시될 이름"),
    ("적용 매출버전", ["기준", "변경", "변경", "변경"], None, "기준 = 매출_기준 시트, 변경 = 매출_변경 시트"),
    ("매출 조정률", [0, 0, 0, 0], PCT, "매출·동량 일괄 증감 (예: -5%)"),
    ("동량 조정률", [0, 0, 0, 0], PCT, "동량만 증감 (동량 검증 반영 등)"),
    ("프리미엄 ($/톤)", [450, 450, 450, 450], NUM, "LME에 더해지는 동 프리미엄"),
    ("기타재료비 증감률", [0, 0, 0, 0], PCT, "제품마스터 기타재료비율에 곱해지는 증감"),
]
for i, (lab, vals, fmt, note) in enumerate(labels):
    rr = 9 + i
    cell(c, f"B{rr}", lab, f_b)
    for s in range(NS):
        cell(c, f"{L(3+s)}{rr}", vals[s], f_in, fmt, fill_in)
    cell(c, f"{L(3+NS)}{rr}", note, f_note)
dv = DataValidation(type="list", formula1='"기준,변경"', allow_blank=False)
c.add_data_validation(dv); dv.add(f"C10:{L(2+NS)}10")
ROW_NAME, ROW_VER, ROW_SADJ, ROW_QADJ, ROW_PREM, ROW_OADJ = 9, 10, 11, 12, 13, 14

# 월별 LME / 환율 / 동가
LME27, FX27 = 14239.703429565683, None
cu27 = 20046.318450491781
FX27 = round(cu27 / (LME27 + 450) * 1000, 2)
def block(top, title, unit, fmt, val_fn, is_input=True):
    cell(c, f"B{top}", f"{title} ({unit})", Font(name=F, size=11, bold=True, color="1F3864"), b=False)
    hdr(c[f"B{top+1}"], "월")
    for s in range(NS):
        hdr(c.cell(top + 1, 3 + s), None); c.cell(top + 1, 3 + s).value = f"={L(3+s)}${ROW_NAME}"
    for m in range(12):
        rr = top + 2 + m
        cell(c, f"B{rr}", f'="M{m+1} ("&TEXT(EDATE($C$4,{m}),"yyyy-mm")&")"', f_calc)
        for s in range(NS):
            v = val_fn(s, m, rr)
            if is_input:
                cell(c, f"{L(3+s)}{rr}", v, f_in, fmt, fill_in)
            else:
                cell(c, f"{L(3+s)}{rr}", v, f_calc, fmt)
    return top + 2

def lme(s, m, rr):
    if s == 2: return f"=C{rr}*1.1" if False else round(LME27 * 1.1, 1)
    return round(LME27, 1)
def fx(s, m, rr):
    return 1300 if s == 3 else FX27
LME0 = block(17, "LME", "$/톤", NUM1, lme)
FX0 = block(17 + 15, "환율", "원/$", NUM2, fx)
CU0 = block(17 + 30, "동가 = (LME+프리미엄)×환율÷1000", "원/kg, 자동계산", NUM,
            lambda s, m, rr: f"=({L(3+s)}{LME0+m}+{L(3+s)}${ROW_PREM})*{L(3+s)}{FX0+m}/1000", False)
cell(c, f"{L(3+NS)}{LME0}", "S3 예시: 수정안에서 LME만 10% 상승 / S4: 환율만 1,300", f_note, b=False)
cell(c, f"{L(3+NS)}{FX0}", "예시 환율은 27년 예상동가 20,046원/kg ÷ (LME+$450)로 역산한 값", f_note, b=False)
c.freeze_panes = "C9"

# ------------------------------------------------------------------ 제품마스터
p = wb.create_sheet("제품마스터")
p.column_dimensions["A"].width = 2
for col, w in zip("BCDEFGH", [6, 16, 16, 12, 14, 14, 40]): p.column_dimensions[col].width = w
p["B2"] = "제품마스터"; p["B2"].font = f_t
p["B3"] = "제품군명은 매출자료의 '제품군' 값과 글자까지 똑같아야 합니다. 행 추가는 빈 칸(최대 20개)에 입력."; p["B3"].font = f_note
for i, h in enumerate(["No", "제품군", "사업구분", "기준 동량지수\n(매출÷동량)", "로스율", "기타재료비율\n(매출 대비)", "비고"]):
    hdr(p.cell(5, 2 + i), h)
prods = [("JIS", "기존사업", 28.85), ("XLPE", "기존사업", 24.16), ("MIL", "기존사업", 49.76),
         ("NEK", "기존사업", 32.64), ("TYPE P", "기존사업", 33.13), ("고무선", "기존사업", 31.30),
         ("원전용", "기존사업", 35.03), ("전력용", "신규사업", 34.86)]
P0 = 6; PE = P0 + NP - 1
for i in range(NP):
    rr = P0 + i
    cell(p, f"B{rr}", i + 1, f_calc)
    nm, biz, idx = prods[i] if i < len(prods) else (None, None, None)
    cell(p, f"C{rr}", nm, f_in, None, fill_in)
    cell(p, f"D{rr}", biz, f_in, None, fill_in)
    cell(p, f"E{rr}", idx, f_in, NUM2, fill_in)
    cell(p, f"F{rr}", 0 if nm else None, f_in, PCT, fill_in)
    cell(p, f"G{rr}", 0 if nm else None, f_in, PCT, fill_in)
    cell(p, f"H{rr}", "동량지수: 26년(9월누계+4Q예상) 실적" if i == 0 else None, f_note)
dvb = DataValidation(type="list", formula1='"기존사업,신규사업"', allow_blank=True)
p.add_data_validation(dvb); dvb.add(f"D{P0}:D{PE}")
p["H7"] = "로스율·기타재료비율은 예시로 0% — 실제 값으로 입력"; p["H7"].font = f_note
p.freeze_panes = "C6"

# ------------------------------------------------------------------ 매출 시트 (기준/변경)
# 열: B No, C 제품군, D 거래처/프로젝트, E 구분, F~Q 매출 M1~M12, R~AC 동량 M1~M12,
# 보조: AE 등록여부, AF~AQ 산출동량 M1~M12
SALE_C0, QTY_C0, HELP_OK, HELP_C0 = 6, 18, 31, 32
def build_sales(name, title, data):
    ws = wb.create_sheet(name)
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 6; ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 18; ws.column_dimensions["E"].width = 10
    for i in range(SALE_C0, HELP_C0 + 12): ws.column_dimensions[L(i)].width = 9.5
    ws.column_dimensions[L(HELP_OK)].width = 8; ws.column_dimensions["AD"].width = 11
    ws["B1"] = title; ws["B1"].font = f_t
    ws["B2"] = "노란 영역(6행부터)에 '값 붙여넣기'. 동량을 모르면 비워 두면 매출÷기준동량지수로 산출."; ws["B2"].font = f_note
    # 검증
    cell(ws, "F2", "버전/메모", f_b); ws.merge_cells("G2:J2"); cell(ws, "G2", data["memo"], f_in, None, fill_in)
    cell(ws, "L2", "매출합계", f_b); cell(ws, "M2", f"=SUM(F{R0}:Q{RE})", f_calc, NUM)
    cell(ws, "N2", "미등록 제품군 행수", f_b)
    cell(ws, "P2", f'=COUNTIF({L(HELP_OK)}{R0}:{L(HELP_OK)}{RE},"미등록")', f_calc, "0")
    cell(ws, "R2", "미등록 매출액", f_b)
    cell(ws, "S2", f'=SUMIF({L(HELP_OK)}{R0}:{L(HELP_OK)}{RE},"미등록",{L(QTY_C0+12)}{R0}:{L(QTY_C0+12)}{RE})', f_calc, NUM)
    ws.conditional_formatting.add("P2", CellIsRule(operator="greaterThan", formula=["0"], fill=PatternFill("solid", fgColor="F8CBAD")))
    # 헤더
    ws.merge_cells(f"{L(SALE_C0)}4:{L(SALE_C0+11)}4"); hdr(ws.cell(4, SALE_C0), "매출액 (조건설정 단위)")
    ws.merge_cells(f"{L(QTY_C0)}4:{L(QTY_C0+11)}4"); hdr(ws.cell(4, QTY_C0), "동량 (톤) – 선택 입력")
    ws.merge_cells(f"{L(HELP_OK)}4:{L(HELP_C0+11)}4"); hdr(ws.cell(4, HELP_OK), "보조계산 (수정 금지): 등록여부 / 산출동량", PatternFill("solid", fgColor="7F7F7F"))
    for i, h in enumerate(["No", "제품군", "거래처/프로젝트", "구분"]): hdr(ws.cell(5, 2 + i), h)
    for m in range(12):
        lab = f'="M{m+1} "&TEXT(EDATE(조건설정!$C$4,{m}),"yy.mm")'
        for c0, fl in ((SALE_C0, fill_h), (QTY_C0, fill_h), (HELP_C0, PatternFill("solid", fgColor="7F7F7F"))):
            hdr(ws.cell(5, c0 + m), lab, fl)
    hdr(ws.cell(5, HELP_OK), "등록", PatternFill("solid", fgColor="7F7F7F"))
    hdr(ws.cell(5, QTY_C0 + 12), "행합계(매출)", PatternFill("solid", fgColor="7F7F7F"))
    rows = data["rows"]
    for k in range(NR):
        rr = R0 + k
        row = rows[k] if k < len(rows) else None
        vals = [k + 1 if row else None, row[0] if row else None, row[1] if row else None, row[2] if row else None]
        for i, v in enumerate(vals):
            ws.cell(rr, 2 + i, v).font = f_in
        for m in range(12):
            a = ws.cell(rr, SALE_C0 + m, row[3][m] if row else None); a.font = f_in; a.number_format = NUM
            q = ws.cell(rr, QTY_C0 + m, row[4][m] if row else None); q.font = f_in; q.number_format = NUM2
            h = ws.cell(rr, HELP_C0 + m,
                f'=IF(ISNUMBER({L(QTY_C0+m)}{rr}),{L(QTY_C0+m)}{rr},IFERROR({L(SALE_C0+m)}{rr}/INDEX(제품마스터!$E${P0}:$E${PE},MATCH($C{rr},제품마스터!$C${P0}:$C${PE},0)),0))')
            h.number_format = NUM2; h.font = f_calc; h.fill = fill_help
        ok = ws.cell(rr, HELP_OK, f'=IF($C{rr}="","",IF(ISNUMBER(MATCH($C{rr},제품마스터!$C${P0}:$C${PE},0)),"OK","미등록"))')
        ok.font = f_calc; ok.fill = fill_help; ok.alignment = ctr
        t = ws.cell(rr, QTY_C0 + 12, f"=SUM({L(SALE_C0)}{rr}:{L(SALE_C0+11)}{rr})"); t.number_format = NUM; t.fill = fill_help
    for rng in (f"B{R0}:{L(QTY_C0+11)}{RE}",):
        pass
    ws.conditional_formatting.add(f"{L(HELP_OK)}{R0}:{L(HELP_OK)}{RE}",
        CellIsRule(operator="equal", formula=['"미등록"'], fill=PatternFill("solid", fgColor="F8CBAD")))
    # 입력영역 표시 (노란 바탕, 대량 셀 서식 대신 조건부 아님 - 헤더 바로 아래 몇 행만 강조 X) -> 전 영역 배경 지정
    for rr in range(R0, RE + 1):
        for cc in list(range(2, 6)) + list(range(SALE_C0, QTY_C0 + 12)):
            ws.cell(rr, cc).fill = fill_in
    ws.freeze_panes = ws.cell(R0, SALE_C0)
    return ws

def load_rows(ext, sid):
    if not ext: return []
    d = ext[sid]; out = []
    from openpyxl.utils import column_index_from_string as ci
    qc = [L(ci("S") + 16 * i) for i in range(12)]; sc = [L(ci("W") + 16 * i) for i in range(12)]
    for r in sorted(d, key=int):
        rr = d[r]
        if not (15 <= int(r) <= 122) or not rr.get("E"): continue
        q = [round(float(rr[x]), 4) if rr.get(x) not in (None, "") else None for x in qc]
        s = [round(float(rr[x]), 4) if rr.get(x) not in (None, "") else None for x in sc]
        q = [v if v else None for v in q]; s = [v if v else None for v in s]
        if not any(q) and not any(s): continue
        # 매출 없이 동량만 있는 셀은 동량 그대로 사용
        out.append((rr["E"], f"원본 {r}행", "국내", s, q))
    return out

ext = json.load(open(EXT)) if EXT else None
build_sales("매출_기준", "매출계획 – 기준 버전", {"memo": "27년 사업계획 Rev.1109 (영업팀)", "rows": load_rows(ext, "6")})
build_sales("매출_변경", "매출계획 – 변경 버전", {"memo": "27년 사업계획 Rev.1109 (최종)", "rows": load_rows(ext, "0")})

# ------------------------------------------------------------------ 계산
k = wb.create_sheet("계산")
k["B1"] = "계산 (자동, 수정 금지)"; k["B1"].font = f_t
k.column_dimensions["A"].width = 2
for col, w in zip("BCDEFGHI", [5, 18, 12, 10, 10, 10, 10, 10]): k.column_dimensions[col].width = w
# 열 배치: B 시나리오#, C 시나리오명, D 제품군, E 매출버전키, F 로스율, G 기타율, H 등록여부(제품명 존재)
METRICS = [("매출액", NUM), ("동량", NUM1), ("동재료비", NUM), ("기타재료비", NUM), ("재료비", NUM)]
M0 = 10  # 첫 지표 시작열
def mcol(mi, m):  # 지표 mi, 월 m(0..11) -> 열번호 ; 각 지표 12개월 + 합계
    return M0 + mi * 13 + m
hdr(k.cell(4, 2), "S#"); hdr(k.cell(4, 3), "시나리오"); hdr(k.cell(4, 4), "제품군")
hdr(k.cell(4, 5), "매출버전"); hdr(k.cell(4, 6), "로스율"); hdr(k.cell(4, 7), "기타율"); hdr(k.cell(4, 8), "동량지수"); hdr(k.cell(4, 9), "사업구분")
for mi, (nm, fmt) in enumerate(METRICS):
    k.merge_cells(start_row=3, start_column=mcol(mi, 0), end_row=3, end_column=mcol(mi, 12))
    hdr(k.cell(3, mcol(mi, 0)), nm)
    for m in range(12):
        hdr(k.cell(4, mcol(mi, m)), f'="M{m+1} "&TEXT(EDATE(조건설정!$C$4,{m}),"yy.mm")', fill_h2, f_b)
        k.column_dimensions[L(mcol(mi, m))].width = 9.5
    hdr(k.cell(4, mcol(mi, 12)), "합계", fill_h2, f_b); k.column_dimensions[L(mcol(mi, 12))].width = 11
K0 = 5
def sales_sum(sheet, col, rr):
    return f"SUMIF({sheet}!$C${R0}:$C${RE},$D{rr},{sheet}!{col}${R0}:{col}${RE})"
for s in range(NS):
    sc = L(3 + s)
    for i in range(NP):
        rr = K0 + s * NP + i
        pr = P0 + i
        k.cell(rr, 2, s + 1)
        k.cell(rr, 3, f"=조건설정!${sc}${ROW_NAME}").font = f_link
        k.cell(rr, 4, f'=IF(제품마스터!$C${pr}="","",제품마스터!$C${pr})').font = f_link
        k.cell(rr, 5, f"=조건설정!${sc}${ROW_VER}").font = f_link
        k.cell(rr, 6, f"=N(제품마스터!$F${pr})").number_format = PCT
        k.cell(rr, 7, f"=N(제품마스터!$G${pr})").number_format = PCT
        k.cell(rr, 8, f"=N(제품마스터!$E${pr})").number_format = NUM2
        k.cell(rr, 9, f'=IF(제품마스터!$D${pr}="","",제품마스터!$D${pr})').font = f_link
        for m in range(12):
            scol, hcol = L(SALE_C0 + m), L(HELP_C0 + m)
            a = L(mcol(0, m)); q = L(mcol(1, m)); cu = L(mcol(2, m)); ot = L(mcol(3, m))
            k[f"{a}{rr}"] = (f'=IF($D{rr}="",0,IF($E{rr}="기준",{sales_sum("매출_기준", scol, rr)},'
                             f'{sales_sum("매출_변경", scol, rr)})*(1+조건설정!${sc}${ROW_SADJ}))')
            k[f"{q}{rr}"] = (f'=IF($D{rr}="",0,IF($E{rr}="기준",{sales_sum("매출_기준", hcol, rr)},'
                             f'{sales_sum("매출_변경", hcol, rr)})*(1+조건설정!${sc}${ROW_SADJ})*(1+조건설정!${sc}${ROW_QADJ}))')
            k[f"{cu}{rr}"] = f"={q}{rr}*(1+$F{rr})*조건설정!${sc}${CU0+m}/1000"
            k[f"{ot}{rr}"] = f"={a}{rr}*$G{rr}*(1+조건설정!${sc}${ROW_OADJ})"
            k[f"{L(mcol(4, m))}{rr}"] = f"={cu}{rr}+{ot}{rr}"
        for mi, (nm, fmt) in enumerate(METRICS):
            k.cell(rr, mcol(mi, 12), f"=SUM({L(mcol(mi,0))}{rr}:{L(mcol(mi,11))}{rr})").fill = fill_tot
            for m in range(13):
                k.cell(rr, mcol(mi, m)).number_format = fmt
KE = K0 + NS * NP - 1
k.freeze_panes = k.cell(K0, M0)
for row in k.iter_rows(min_row=K0, max_row=KE, min_col=2, max_col=mcol(4, 12)):
    for x in row:
        if x.font == Font(): x.font = f_calc

# ------------------------------------------------------------------ 비교
v = wb.create_sheet("비교")
v.column_dimensions["A"].width = 2; v.column_dimensions["B"].width = 14
v["B2"] = "시나리오 비교"; v["B2"].font = f_t
# 기간 정의 표 (숨김 영역 AZ:BB)
periods = [("연간", 1, 12), ("상반기", 1, 6), ("하반기", 7, 12), ("1Q", 1, 3), ("2Q", 4, 6), ("3Q", 7, 9), ("4Q", 10, 12)] + \
          [(f"M{i}", i, i) for i in range(1, 13)]
PZ = 60  # 열 번호(BH)
hdr(v.cell(3, PZ), "기간"); hdr(v.cell(3, PZ + 1), "시작"); hdr(v.cell(3, PZ + 2), "종료")
for i, (nm, a, b) in enumerate(periods):
    v.cell(4 + i, PZ, nm); v.cell(4 + i, PZ + 1, a); v.cell(4 + i, PZ + 2, b)
PL = f"${L(PZ)}$4:${L(PZ)}${3+len(periods)}"
cell(v, "B4", "기간 선택", f_b); cell(v, "C4", "연간", f_in, None, fill_in)
dvp = DataValidation(type="list", formula1=PL, allow_blank=False); v.add_data_validation(dvp); dvp.add("C4")
cell(v, "D4", f'=INDEX(${L(PZ+1)}$4:${L(PZ+1)}$30,MATCH($C$4,{PL},0))', f_calc, "0")
cell(v, "E4", f'=INDEX(${L(PZ+2)}$4:${L(PZ+2)}$30,MATCH($C$4,{PL},0))', f_calc, "0")
cell(v, "F4", '=TEXT(EDATE(조건설정!$C$4,$D$4-1),"yyyy.mm")&" ~ "&TEXT(EDATE(조건설정!$C$4,$E$4-1),"yyyy.mm")', f_calc)
v.merge_cells("F4:H4")
cell(v, "J4", "단위: 매출·재료비 = ", f_note, b=False); v["L4"] = '=조건설정!$C$5&", 동량 = "&조건설정!$C$6'; v["L4"].font = f_note
for col in range(PZ, PZ + 3): v.column_dimensions[L(col)].hidden = True

# 요약표 (시나리오 x 지표)
def period_sum(mi, crit_s, crit_p=None, biz=None):
    """계산 시트에서 지표 mi 의 선택기간 합 (SUMPRODUCT)."""
    rng = f"계산!${L(mcol(mi,0))}${K0}:${L(mcol(mi,11))}${KE}"
    mon = f"(COLUMN(계산!${L(mcol(mi,0))}$4:${L(mcol(mi,11))}$4)-COLUMN(계산!${L(mcol(mi,0))}$4)+1)"
    cond = f"(계산!$B${K0}:$B${KE}={crit_s})"
    if crit_p: cond += f"*(계산!$D${K0}:$D${KE}={crit_p})"
    if biz: cond += f"*(계산!$I${K0}:$I${KE}={biz})"
    return f"SUMPRODUCT({cond}*({mon}>=$D$4)*({mon}<=$E$4)*{rng})"

r = 6
cell(v, f"B{r}", "① 시나리오 요약", Font(name=F, size=11, bold=True, color="1F3864"), b=False)
r += 1
heads = ["시나리오", "매출버전", "매출액", "동량", "평균동가(원/kg)", "동재료비", "기타재료비", "재료비", "재료비율", "S1 대비 재료비", "S1 대비 %"]
for i, h in enumerate(heads): hdr(v.cell(r, 2 + i), h)
SUM0 = r + 1
for s in range(NS):
    rr = SUM0 + s
    cell(v, f"B{rr}", f"=조건설정!{L(3+s)}${ROW_NAME}", f_link)
    cell(v, f"C{rr}", f"=조건설정!{L(3+s)}${ROW_VER}", f_link)
    for mi, col in zip([0, 1, 2, 3, 4], ["D", "E", "G", "H", "I"]):
        cell(v, f"{col}{rr}", "=" + period_sum(mi, s + 1), f_calc, NUM1 if mi == 1 else NUM)
    cell(v, f"F{rr}", f"=IFERROR(G{rr}/E{rr}/(1+0)*1000/1000,0)", f_calc, NUM)
    v[f"F{rr}"].value = f"=IFERROR(G{rr}*1000/E{rr},0)"
    v[f"F{rr}"].comment = Comment("동재료비÷동량 (로스율 포함 실효 동가)", "template")
    cell(v, f"J{rr}", f"=IFERROR(I{rr}/D{rr},0)", f_calc, PCT)
    cell(v, f"K{rr}", f"=I{rr}-$I${SUM0}", f_calc, NUM)
    cell(v, f"L{rr}", f"=IFERROR(I{rr}/$I${SUM0}-1,0)", f_calc, PCT)
for col, w in zip("CDEFGHIJKL", [10, 13, 11, 13, 13, 12, 13, 10, 13, 10]): v.column_dimensions[col].width = w

# 제품군별 비교 (지표 선택)
r = SUM0 + NS + 2
cell(v, f"B{r}", "② 제품군별 비교", Font(name=F, size=11, bold=True, color="1F3864"), b=False)
cell(v, f"E{r}", "지표 선택", f_b); cell(v, f"F{r}", "재료비", f_in, None, fill_in)
dvm = DataValidation(type="list", formula1='"매출액,동량,동재료비,기타재료비,재료비"', allow_blank=False)
v.add_data_validation(dvm); dvm.add(f"F{r}")
SEL = f"$F${r}"
r += 1
hdr(v.cell(r, 2), "제품군"); hdr(v.cell(r, 3), "사업구분")
for s in range(NS): hdr(v.cell(r, 4 + s), f"=조건설정!{L(3+s)}${ROW_NAME}")
for s in range(1, NS): hdr(v.cell(r, 4 + NS + s - 1), f'="S{s+1}-S1"', fill_h2, f_b)
PR0 = r + 1
def sel_sum(s, rr_prod=None, biz=None):
    parts = [period_sum(mi, s + 1, rr_prod, biz) for mi in range(5)]
    return f'=CHOOSE(MATCH({SEL},{{"매출액","동량","동재료비","기타재료비","재료비"}},0),' + ",".join(parts) + ")"
for i in range(NP):
    rr = PR0 + i
    cell(v, f"B{rr}", f'=IF(제품마스터!$C${P0+i}="","",제품마스터!$C${P0+i})', f_link)
    cell(v, f"C{rr}", f'=IF(제품마스터!$D${P0+i}="","",제품마스터!$D${P0+i})', f_link)
    for s in range(NS):
        cell(v, f"{L(4+s)}{rr}", f'=IF($B{rr}="","",' + sel_sum(s, f"$B{rr}")[1:] + ")", f_calc, NUM)
    for s in range(1, NS):
        cell(v, f"{L(4+NS+s-1)}{rr}", f'=IF($B{rr}="","",{L(4+s)}{rr}-$D{rr})', f_calc, NUM)
PRE = PR0 + NP - 1
sub = [("기존사업 소계", '"기존사업"'), ("신규사업 소계", '"신규사업"')]
for j, (nm, crit) in enumerate(sub):
    rr = PRE + 1 + j
    cell(v, f"B{rr}", nm, f_b, None, fill_tot); cell(v, f"C{rr}", None, f_calc, None, fill_tot)
    for s in range(NS):
        cell(v, f"{L(4+s)}{rr}", f'=SUMIF($C${PR0}:$C${PRE},{crit},{L(4+s)}${PR0}:{L(4+s)}${PRE})', f_b, NUM, fill_tot)
    for s in range(1, NS):
        cell(v, f"{L(4+NS+s-1)}{rr}", f"={L(4+s)}{rr}-$D{rr}", f_b, NUM, fill_tot)
TOT = PRE + 3
cell(v, f"B{TOT}", "TOTAL", f_b, None, fill_tot); cell(v, f"C{TOT}", None, f_calc, None, fill_tot)
for s in range(NS):
    cell(v, f"{L(4+s)}{TOT}", f"=SUM({L(4+s)}{PR0}:{L(4+s)}{PRE})", f_b, NUM, fill_tot)
for s in range(1, NS):
    cell(v, f"{L(4+NS+s-1)}{TOT}", f"={L(4+s)}{TOT}-$D{TOT}", f_b, NUM, fill_tot)
cell(v, f"B{TOT+1}", "검증: ①요약과 일치", f_note, b=False)
cell(v, f"D{TOT+1}", f'=IF(ABS(D{TOT}-CHOOSE(MATCH({SEL},{{"매출액","동량","동재료비","기타재료비","재료비"}},0),D{SUM0},E{SUM0},G{SUM0},H{SUM0},I{SUM0}))<0.5,"OK","확인")', f_note, b=False)

# 월별 재료비 추이
r = TOT + 3
cell(v, f"B{r}", "③ 월별 재료비 추이 (전체 12개월)", Font(name=F, size=11, bold=True, color="1F3864"), b=False)
r += 1
hdr(v.cell(r, 2), "월")
for s in range(NS): hdr(v.cell(r, 3 + s), f"=조건설정!{L(3+s)}${ROW_NAME}")
MN0 = r + 1
for m in range(12):
    rr = MN0 + m
    cell(v, f"B{rr}", f'=TEXT(EDATE(조건설정!$C$4,{m}),"yy.mm")', f_calc)
    for s in range(NS):
        col = L(mcol(4, m))
        cell(v, f"{L(3+s)}{rr}", f"=SUMIF(계산!$B${K0}:$B${KE},{s+1},계산!{col}${K0}:{col}${KE})", f_calc, NUM)
ch = LineChart(); ch.title = "월별 재료비 (시나리오별)"; ch.height = 8; ch.width = 18
ch.add_data(Reference(v, min_col=3, max_col=2 + NS, min_row=MN0 - 1, max_row=MN0 + 11), titles_from_data=True)
ch.set_categories(Reference(v, min_col=2, min_row=MN0, max_row=MN0 + 11))
ch.roundedCorners = False; ch.legend.position = "b"; ch.legend.overlay = False; ch.title.overlay = False
ch.x_axis.delete = False; ch.y_axis.delete = False; ch.y_axis.numFmt = "#,##0"
v.add_chart(ch, f"H{MN0-1}")
v.freeze_panes = "B6"

# 시트 순서/탭 색
order = ["사용안내", "조건설정", "제품마스터", "매출_기준", "매출_변경", "비교", "계산"]
wb._sheets = [wb[n] for n in order]
for n, col in [("조건설정", "FFC000"), ("제품마스터", "FFC000"), ("매출_기준", "FFC000"), ("매출_변경", "FFC000"), ("비교", "70AD47"), ("계산", "A6A6A6")]:
    wb[n].sheet_properties.tabColor = col
wb.active = 5
wb.save(OUT)
print("saved", OUT, "rows", len(load_rows(ext, "6")), len(load_rows(ext, "0")))
