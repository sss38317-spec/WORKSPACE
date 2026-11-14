"""0단계: 영업팀 매출계획 양식 (VBA 없이 '지침' 입력 방식) 생성.

사용: python build_sales_template.py <원본 예시양식.xlsx> <출력.xlsx>
- 원본에서 머리글(1~14행), 예시 데이터(15~30행), 동량지수 시트를 가져온다.
- LME(AA)·동적용환율(AD) 칸은 숫자 또는 "지침"만 허용.
"""
import sys
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.utils import get_column_letter as L, column_index_from_string as CI
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.comments import Comment

SRC, OUT = sys.argv[1], sys.argv[2]
src = openpyxl.load_workbook(SRC)
s_main = src["27년 매출월별상세(최종)"]
s_idx = src["동량지수_1~9월 매출품목"]

FONT = "맑은 고딕"
f_base = Font(name=FONT, size=9)
f_in = Font(name=FONT, size=9, color="0000FF")
f_b = Font(name=FONT, size=9, bold=True)
f_title = Font(name=FONT, size=12, bold=True, color="1F3864")
f_note = Font(name=FONT, size=9, color="595959")
fill_hd = PatternFill("solid", fgColor="D9E1F2")
fill_hd2 = PatternFill("solid", fgColor="1F3864")
fill_in = PatternFill("solid", fgColor="FFF2CC")
fill_calc = PatternFill("solid", fgColor="F2F2F2")
fill_sum = PatternFill("solid", fgColor="E2EFDA")
thin = Side(style="thin", color="BFBFBF")
bd = Border(left=thin, right=thin, top=thin, bottom=thin)
wrap = Alignment(horizontal="center", vertical="center", wrap_text=True)
UNLOCK = Protection(locked=False)
NUM0 = '#,##0;(#,##0);"-"'
NUM1 = '#,##0.0;(#,##0.0);"-"'
NUM2 = '#,##0.00;(#,##0.00);"-"'
PCT = '0.0%;(0.0%);"-"'

R1, NROWS = 15, 200            # 데이터 15행 ~ 214행
RE = R1 + NROWS - 1
M0, MW = CI("S"), 16           # 1월 시작열, 월 블록 폭
Y0 = M0 + 12 * MW              # HC
ADJ = CI("IW")                 # 동량지수 조정률(예외)
OFF = dict(S=0, T=1, U=2, V=3, W=4, X=5, Y=6, Z=7, AA=8, AB=9, AC=10, AD=11, AE=12, AF=13, AG=14, AH=15)
def mc(m, k): return L(M0 + m * MW + OFF[k])

G = "사업계획지침"
def g_lme(m): return f"{G}!${L(3+m)}$11"
def g_fx(m): return f"{G}!${L(3+m)}$14"
G_PRE = f"{G}!$C$12"

wb = Workbook()

# =============================================================== 사업계획지침
g = wb.active; g.title = G
g.column_dimensions["A"].width = 2; g.column_dimensions["B"].width = 26
for c in range(3, 16): g.column_dimensions[L(c)].width = 10
g["B2"] = "사업계획 기준 지침"; g["B2"].font = f_title
g["B4"], g["C4"], g["D4"], g["E4"], g["F4"], g["G4"] = "사업계획", "KRW", "USD", "JPY", "EUR", "동가"
g["B5"], g["C5"], g["D5"], g["E5"], g["F5"] = "환율 (원/외화 1단위)", 1, 1360, 8.9, 1600
g["B6"] = "적용동가 (연평균, 원/kg)"; g["G6"] = "=AVERAGE(C15:N15)"
for c in range(2, 8):
    g.cell(4, c).font = f_b; g.cell(4, c).fill = fill_hd; g.cell(4, c).border = bd; g.cell(4, c).alignment = wrap
    g.cell(5, c).border = bd; g.cell(6, c).border = bd
for c in range(3, 7): g.cell(5, c).font = f_in; g.cell(5, c).fill = fill_in; g.cell(5, c).number_format = "#,##0.00"
g["B5"].font = f_b; g["B6"].font = f_b; g["G6"].number_format = NUM0
g["B8"] = "※ 월별 지침 (영업 양식에서 \"지침\"이라고 입력한 셀에 적용)"; g["B8"].font = f_b
g["B10"] = "구분"
for m in range(12): g.cell(10, 3 + m, f"{m+1}월")
g["O10"] = "연간평균"
rows = [(11, "LME ($/Ton)", [14500] * 12, "#,##0"),
        (12, "동프리미엄 ($/Ton, 연간 일괄)", None, "#,##0"),
        (13, "적용LME = LME + 프리미엄", None, "#,##0"),
        (14, "동적용환율 (원/$)", [1360] * 12, "#,##0.00"),
        (15, "적용동가 (원/kg)", None, "#,##0")]
for r, lab, vals, fmt in rows:
    g.cell(r, 2, lab).font = f_b; g.cell(r, 2).border = bd
    for m in range(12):
        c = g.cell(r, 3 + m); c.border = bd; c.number_format = fmt
        if vals: c.value = vals[m]; c.font = f_in; c.fill = fill_in
        elif r == 13: c.value = f"={L(3+m)}11+$C$12"
        elif r == 15: c.value = f"={L(3+m)}13*{L(3+m)}14/1000"
    if r != 12:
        g.cell(r, 15, f"=AVERAGE(C{r}:N{r})").number_format = fmt
    g.cell(r, 15).border = bd
g["C12"] = 520; g["C12"].font = f_in; g["C12"].fill = fill_in
g["D12"] = "← 연간 1개 값 (D~N 칸은 비워 둠)"; g["D12"].font = f_note
for c in range(2, 16): g.cell(10, c).font = f_b; g.cell(10, c).fill = fill_hd; g.cell(10, c).border = bd; g.cell(10, c).alignment = wrap
notes = ["※ 노란 칸만 수정합니다. 동적용환율 기본값은 USD 사업계획환율(D5)과 같게 입력해 두었습니다.",
         "※ 프리미엄은 원래 $450 → 현재 $520 (변경 조건). 이 칸 하나만 바꾸면 전체 행에 반영됩니다.",
         "※ 영업 매출계획 시트: LME·동적용환율 칸에 숫자(직접입력) 또는 \"지침\"을 입력. 빈칸은 입력 누락으로 빨간색 표시."]
for i, t in enumerate(notes): g.cell(17 + i, 2, t).font = f_note

# =============================================================== 동량지수 시트 (원본 복사)
ix = wb.create_sheet("동량지수_1~9월 매출품목")
for row in s_idx.iter_rows(min_row=1, max_row=s_idx.max_row):
    for c in row:
        if c.value is None: continue
        n = ix.cell(c.row, c.column, c.value); n.font = f_base
        if c.row >= 12 and c.column >= 3 and not (isinstance(c.value, str) and c.value.startswith("=")):
            n.font = f_in
for m in s_idx.merged_cells.ranges: ix.merge_cells(str(m))
for k, d in s_idx.column_dimensions.items(): ix.column_dimensions[k].width = d.width
for c in range(2, 19):
    for r in (10, 11): ix.cell(r, c).font = f_b; ix.cell(r, c).fill = fill_hd; ix.cell(r, c).alignment = wrap
ix["T10"] = "※ 행을 추가하면 B열 CONCATENATE 수식도 같이 복사하세요 (매출계획 N~R열이 이 표를 VLOOKUP 합니다)."; ix["T10"].font = f_note

# =============================================================== 매출계획 시트
ws = wb.create_sheet("27년 매출월별상세", 0)
# 머리글 1~14행 복사 (값 + 병합)
for r in range(1, 15):
    for c in range(1, 257):
        v = s_main.cell(r, c).value
        if v is not None: ws.cell(r, c, v)
for m in s_main.merged_cells.ranges:
    if m.max_row <= 14: ws.merge_cells(str(m))
for k, d in s_main.column_dimensions.items():
    if d.width: ws.column_dimensions[k].width = d.width
for r in range(1, 15):
    if s_main.row_dimensions[r].height: ws.row_dimensions[r].height = s_main.row_dimensions[r].height
ws.row_dimensions[12].height = 120
# 머리글 스타일
for r in range(11, 15):
    for c in range(2, Y0 + 13):
        x = ws.cell(r, c); x.font = f_b; x.alignment = wrap; x.border = bd; x.fill = fill_hd
for r in range(1, 11):
    for c in range(1, 60):
        if ws.cell(r, c).value is not None: ws.cell(r, c).font = f_note
ws["B9"].font = f_title
# 작성 안내 추가 (S6~S8)
ws["S6"] = "* LME·동적용환율: 숫자(직접입력) 또는 \"지침\" 입력 → \"지침\"이면 사업계획지침 시트 값 자동 적용 (빈칸 = 입력 누락, 빨간색)"
ws["S7"] = "* 동프리미엄·사업계획환율은 사업계획지침 시트에서 일괄 적용 (입력 불필요)"
ws["S8"] = "* IW열 '동량지수 조정률(예외)': 일부 거래선만 동량지수를 낮출 때 입력 (예: 3.4%)"
for k in ("S6", "S7", "S8"): ws[k].font = Font(name=FONT, size=9, bold=True, color="C00000")
# 월 블록 머리글 보정
for m in range(12):
    ws[f"{mc(m,'AA')}14"] = "*LME\n(숫자 또는 \"지침\")"
    ws[f"{mc(m,'AD')}14"] = "*동적용환율\n(숫자 또는 \"지침\")"
    ws[f"{mc(m,'AB')}14"] = "동프리미엄\n(지침 일괄)"
    ws[f"{mc(m,'AE')}14"] = "사업계획환율\n(지침 일괄)"
ws[f"{L(ADJ)}12"] = "동량지수 조정률\n(예외)"; ws[f"{L(ADJ)}13"] = "%"; ws[f"{L(ADJ)}14"] = "일부 거래선만\n입력 (예: 3.4%)"
for r in (12, 13, 14):
    x = ws.cell(r, ADJ); x.font = f_b; x.alignment = wrap; x.border = bd; x.fill = fill_hd
ws.column_dimensions[L(ADJ)].width = 12
# HQ~IV 머리글 (원본 값 유지됨) + 스타일
for r in range(12, 15):
    for c in range(Y0 + 13, ADJ):
        x = ws.cell(r, c)
        if x.value is not None: x.font = f_b; x.alignment = wrap; x.fill = fill_hd; x.border = bd

# ---- 예시 데이터 읽기 (원본 15~30행)
example = []
for r in range(15, 31):
    if s_main.cell(r, 2).value is None: continue
    t = str(s_main.cell(r, CI("T")).value or "")
    adj = 0.034 if "3.4%" in t else None
    d = {"base": [s_main.cell(r, c).value for c in range(2, 14)], "adj": adj, "m": []}
    for m in range(12):
        def v(k): return s_main.cell(r, M0 + m * MW + OFF[k]).value
        lme, fx = v("AA"), v("AD")
        d["m"].append(dict(W=v("W"), AA=("지침" if lme == 14500 else lme), AD=("지침" if fx == 1360 else fx), AG=v("AG")))
    example.append(d)

def eff(col, r, gref):
    return f'IF({col}{r}="지침",{gref},N({col}{r}))'

for i in range(NROWS):
    r = R1 + i
    ex = example[i] if i < len(example) else None
    # A블록 (B~M) 입력
    for k, c in enumerate(range(2, 14)):
        x = ws.cell(r, c, ex["base"][k] if ex else None); x.font = f_in; x.protection = UNLOCK; x.fill = fill_in
    # B블록 (N~R) VLOOKUP
    key = f"CONCATENATE($D{r},$E{r},$F{r},$M{r})"
    rng = "'동량지수_1~9월 매출품목'!$B$12:$R$300"
    for col, n in (("N", 10), ("O", 13), ("P", 14), ("Q", 17), ("R", 8)):
        ws[f"{col}{r}"] = f'=IF($E{r}="","",VLOOKUP({key},{rng},{n},FALSE))'
        ws[f"{col}{r}"].number_format = "0.0##" if col in "NP" else ("General" if col == "R" else "0.00##")
    for m in range(12):
        c = {k: mc(m, k) for k in OFF}
        lme, fx = eff(c["AA"], r, g_lme(m)), eff(c["AD"], r, g_fx(m))
        ws[f"{c['S']}{r}"] = f"=IFERROR(ROUND({c['W']}{r}/{c['T']}{r},2),0)"
        ws[f"{c['T']}{r}"] = f"=IF($E{r}=\"\",0,(($O{r}*{c['U']}{r}+$Q{r}*{c['V']}{r})+{c['X']}{r})*(1-N(${L(ADJ)}{r})))"
        ws[f"{c['U']}{r}"] = f"=$N{r}"
        ws[f"{c['V']}{r}"] = f"=$P{r}"
        ws[f"{c['X']}{r}"] = f"=ROUNDDOWN(IF({c['AA']}{r}=\"\",0,({lme}-N($R{r}))*{fx}/1000000),2)"
        ws[f"{c['Y']}{r}"] = f"={c['S']}{r}*{c['Z']}{r}/1000"
        ws[f"{c['Z']}{r}"] = f"=IF({c['AC']}{r}=\"\",0,{c['AC']}{r}*{fx}/1000)"
        ws[f"{c['AB']}{r}"] = f"={G_PRE}"
        ws[f"{c['AC']}{r}"] = f"=IF({c['AA']}{r}=\"\",\"\",{lme}+{c['AB']}{r})"
        ws[f"{c['AE']}{r}"] = f"=IF($M{r}=\"\",\"\",HLOOKUP($M{r},{G}!$B$4:$G$6,2,FALSE))"
        ws[f"{c['AF']}{r}"] = f"=IF(OR($M{r}=\"KRW\",$M{r}=\"\"),0,IFERROR({c['W']}{r}/{c['AE']}{r},0))"
        ws[f"{c['AH']}{r}"] = f"={c['W']}{r}*{c['AG']}{r}"
        em = ex["m"][m] if ex else {}
        for k in ("W", "AA", "AD", "AG"):
            x = ws[f"{c[k]}{r}"]; x.value = em.get(k); x.font = f_in; x.protection = UNLOCK; x.fill = fill_in
        for k in ("U", "V"):  # 세/태선 비중: 기본 수식, 영업이 덮어쓰기 가능
            ws[f"{c[k]}{r}"].protection = UNLOCK
        fm = {"S": NUM2, "T": NUM2, "U": "0.0##%", "V": "0.0##%", "W": "General", "X": NUM2, "Y": NUM1, "Z": NUM0,
              "AA": "General", "AB": "#,##0", "AC": "#,##0", "AD": "General", "AE": "#,##0.00", "AF": NUM2, "AG": "0.0##%", "AH": NUM1}
        for k, f in fm.items(): ws[f"{c[k]}{r}"].number_format = f
    # 조정률
    x = ws.cell(r, ADJ, ex["adj"] if ex else None); x.font = f_in; x.protection = UNLOCK; x.fill = fill_in; x.number_format = "0.0##%"
    # 연간 (HC~HO)
    col = lambda k: [f"{mc(m,k)}{r}" for m in range(12)]
    S_, W_, Y_, Z_, AB_, AF_, AH_ = (col(k) for k in ("S", "W", "Y", "Z", "AB", "AF", "AH"))
    lme_terms = [f"{mc(m,'S')}{r}*IF({mc(m,'AA')}{r}=\"지침\",{g_lme(m)},N({mc(m,'AA')}{r}))" for m in range(12)]
    a = lambda k: f"{L(Y0 + k)}{r}"
    ws[a(0)] = "=" + "+".join(S_)
    ws[a(1)] = f"=IFERROR({a(2)}/{a(0)},0)"
    ws[a(2)] = "=" + "+".join(W_)
    ws[a(3)] = "=" + "+".join(Y_)
    ws[a(4)] = f"=IFERROR((" + "+".join(f"{s}*{z}" for s, z in zip(S_, Z_)) + f")/{a(0)},0)"
    ws[a(5)] = f"=IFERROR((" + "+".join(lme_terms) + f")/{a(0)},0)"
    ws[a(6)] = "=AVERAGE(" + ",".join(AB_) + ")"
    ws[a(7)] = f"={a(5)}+{a(6)}"
    ws[a(8)] = f"=IFERROR({a(4)}*1000/{a(7)},0)"      # 수정: 프리미엄 포함 LME로 환율 산출
    ws[a(9)] = f"=IFERROR({a(2)}/{a(10)},0)"
    ws[a(10)] = "=" + "+".join(AF_)
    ws[a(11)] = f"=IFERROR({a(12)}/{a(2)},0)"
    ws[a(12)] = "=" + "+".join(AH_)
    for k, f in enumerate([NUM2, NUM2, NUM0, NUM1, NUM0, "#,##0", "#,##0", "#,##0", "#,##0.00", "#,##0.00", NUM2, "0.0%", NUM1]):
        ws[a(k)].number_format = f; ws[a(k)].fill = fill_sum
    # HP, 판매수수료(HQ~IC), 해외운송비(IE~IQ), IU, IV
    ws[f"HP{r}"] = f"={a(5)}*{a(0)}"
    for m in range(12):
        ws[f"{L(CI('HQ')+m)}{r}"] = f"={mc(m,'W')}{r}*N($J{r})"
        ws[f"{L(CI('IE')+m)}{r}"] = f"={mc(m,'W')}{r}*N($K{r})"
    ws[f"IC{r}"] = f"=SUM(HQ{r}:IB{r})"; ws[f"IQ{r}"] = f"=SUM(IE{r}:IP{r})"
    ws[f"IU{r}"] = f"={a(4)}*{a(0)}/1000"; ws[f"IV{r}"] = f"=IU{r}-{a(3)}"
    for c in range(1, ADJ + 1): ws.cell(r, c).border = bd
    for c in range(14, ADJ):
        x = ws.cell(r, c)
        if x.font != f_in: x.font = f_base

# ---- 데이터 유효성 검사 / 조건부 서식 (LME·동적용환율)
red = PatternFill("solid", fgColor="F8CBAD"); blue = PatternFill("solid", fgColor="DDEBF7")
for m in range(12):
    for k, nm in (("AA", "LME"), ("AD", "동적용환율")):
        cl = mc(m, k); rng = f"{cl}{R1}:{cl}{RE}"
        dv = DataValidation(type="custom", formula1=f'OR({cl}{R1}="지침",AND(ISNUMBER({cl}{R1}),{cl}{R1}>0))',
                            allow_blank=True, showErrorMessage=True, showInputMessage=True,
                            errorTitle="입력 오류", error=f'{nm}은 숫자 또는 "지침"만 입력할 수 있습니다.',
                            promptTitle=nm, prompt='숫자(직접입력) 또는 "지침" 입력')
        ws.add_data_validation(dv); dv.add(rng)
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'AND({cl}{R1}="",N({mc(m,"W")}{R1})<>0)'], fill=red))
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{cl}{R1}="지침"'], fill=blue, font=Font(color="1F4E79", bold=True)))
ws.conditional_formatting.add(f"N{R1}:R{RE}", FormulaRule(formula=[f"ISNA(N{R1})"], fill=red))

# ---- 하단 소계 (원본 34~55행 구조)
SB = RE + 2
def sum_cols():
    out = []
    for m in range(12):
        out += [(mc(m, "S"), "sum"), (mc(m, "T"), ("ratio", mc(m, "W"), mc(m, "S"))), (mc(m, "W"), "sum"), (mc(m, "Y"), "sum"),
                (mc(m, "Z"), ("ratio1000", mc(m, "Y"), mc(m, "S"))), (mc(m, "AG"), ("ratio", mc(m, "AH"), mc(m, "W"))), (mc(m, "AH"), "sum")]
    A = lambda k: L(Y0 + k)
    out += [(A(0), "sum"), (A(1), ("ratio", A(2), A(0))), (A(2), "sum"), (A(3), "sum"), (A(4), ("ratio1000", A(3), A(0))),
            (A(11), ("ratio", A(12), A(2))), (A(12), "sum"), ("IC", "sum"), ("IQ", "sum")]
    return out
COLS = sum_cols()
def write_row(r, label, kind):
    ws.cell(r, 3, label).font = f_b
    for c in range(2, ADJ): ws.cell(r, c).fill = fill_sum
    for colL, how in COLS:
        if how == "sum":
            if kind[0] == "subtotal": f = f"=SUBTOTAL(9,{colL}{R1}:{colL}{RE})"
            elif kind[0] == "sumif": f = f"=SUMIF(${kind[1]}${R1}:${kind[1]}${RE},$C{r},{colL}${R1}:{colL}${RE})"
            else: f = "=" + "+".join(f"{colL}{x}" for x in kind[1])
        else:
            t, num, den = how
            f = f"=IFERROR({num}{r}/{den}{r},0)" if t == "ratio" else f"=IFERROR({num}{r}*1000/{den}{r},0)"
        ws[f"{colL}{r}"] = f; ws[f"{colL}{r}"].font = f_b
        ws[f"{colL}{r}"].number_format = PCT if colL in [mc(m, "AG") for m in range(12)] + [L(Y0 + 11)] else NUM1
write_row(SB, "합계 (필터 시 보이는 행)", ("subtotal",))
r = SB + 2; ws.cell(r, 3, "팀 별").font = f_title
team = {}
for nm in ("국내", "해외", "신사업"):
    r += 1; write_row(r, nm, ("sumif", "C")); team[nm] = r
r += 1; write_row(r, "영업 TOTAL", ("rows", list(team.values()))); TEAM_TOT = r
r += 2; ws.cell(r, 3, "제품군별").font = f_title
def prod(nm): global r; r += 1; write_row(r, nm, ("sumif", "E")); return r
def grp(nm, rows_): global r; r += 1; write_row(r, nm, ("rows", rows_)); return r
a1, a2 = prod("JIS"), prod("XLPE"); g1 = grp("상선용 계", [a1, a2])
b1, b2 = prod("NEK"), prod("TYPE P"); g2 = grp("해양용 계", [b1, b2])
g3 = grp("선박선 계", [g1, g2])
c1 = prod("고무선"); g4 = grp("고무선 계", [c1])
d1, d2 = prod("원전용"), prod("전력용"); g5 = grp("전력용 계", [d1, d2])
PT = grp("영업 TOTAL", [g5, g4, g3])
r += 2; ws.cell(r, 3, "검증 (팀별 = 제품군별)").font = f_b
for colL, how in COLS:
    if how == "sum":
        ws[f"{colL}{r}"] = f'=IF(ROUND({colL}{PT}-{colL}{TEAM_TOT},6)=0,"TRUE","FALSE")'
        ws.conditional_formatting.add(f"{colL}{r}", FormulaRule(formula=[f'{colL}{r}="FALSE"'], fill=red))
ws.cell(r + 1, 3, "※ FALSE이면 상세제품군 목록(제품군별)에 없는 품목이 있습니다. 소계 행을 추가하세요.").font = f_note

# ---- 시트 보호 / 보기 설정
ws.freeze_panes = "S15"
ws.sheet_view.zoomScale = 85
ws.protection.sheet = True
ws.protection.insertColumns = True     # 열 삽입 금지
ws.protection.deleteColumns = True
ws.protection.insertRows = False       # 행 삽입 허용
ws.protection.deleteRows = False
ws.protection.formatColumns = False
ws.protection.formatRows = False
ws.protection.autoFilter = False
ws.protection.sort = False
ws.auto_filter.ref = f"B14:{L(ADJ)}{RE}"
ws.sheet_properties.tabColor = "1F3864"
g.sheet_properties.tabColor = "FFC000"
wb.active = 0
wb.save(OUT)
print("saved", OUT, "example rows", len(example))
