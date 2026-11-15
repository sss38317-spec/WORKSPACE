"""재료비 기준 양식: ① 재료비동가 (26년 11월 ~ 27년 12월), ② 상품재료비율 (제품군별 1개).

사용: python build_basis_template.py <출력.xlsx> [영업 예시양식.xlsx (동량지수 시트 예시 복사용)]
기존 '재료비 동가 예시' 산식을 따르되: 월별 지침 값, 연간 목표 프리미엄 1개, 26년 11·12월 예상 값, 전월 매입동가 = 당월 적용동가.
"""
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.formatting.rule import FormulaRule

OUT = sys.argv[1]
SRC = sys.argv[2] if len(sys.argv) > 2 else None
F = "맑은 고딕"
f_b = Font(name=F, size=9, bold=True); f_in = Font(name=F, size=9, color="0000FF"); f_c = Font(name=F, size=9)
f_t = Font(name=F, size=12, bold=True, color="1F3864"); f_n = Font(name=F, size=9, color="595959")
hd = PatternFill("solid", fgColor="D9E1F2"); yin = PatternFill("solid", fgColor="FFF2CC"); key = PatternFill("solid", fgColor="E2EFDA")
gry = PatternFill("solid", fgColor="F2F2F2")
th = Side(style="thin", color="BFBFBF"); bd = Border(left=th, right=th, top=th, bottom=th)
ctr = Alignment(horizontal="center", vertical="center", wrap_text=True)

wb = Workbook(); ws = wb.active; ws.title = "재료비동가"
ws.column_dimensions["A"].width = 2; ws.column_dimensions["B"].width = 12; ws.column_dimensions["C"].width = 24; ws.column_dimensions["D"].width = 7
MON = ["26년 11월", "26년 12월"] + [f"{m}월" for m in range(1, 13)]
C0 = 5                                   # E열 = 26년 11월, F = 26년 12월, G~R = 27년 1~12월
for i in range(14): ws.column_dimensions[L(C0 + i)].width = 10.5
ws.column_dimensions[L(C0 + 14)].width = 46
ws["B1"] = "재료비 적용동가 기준"; ws["B1"].font = f_t
ws["B2"] = "목표 프리미엄 ($/톤, 연간 1개)"; ws["B2"].font = f_b
ws["E2"] = 500; ws["E2"].font = f_in; ws["E2"].fill = yin; ws["E2"].border = bd; ws["E2"].number_format = "#,##0"
ws["F2"] = "← 계획 적용 프리미엄 (모든 달에 적용)"; ws["F2"].font = f_n
hdr = ["구분", "항목", "단위"] + MON + ["비고"]
for i, h in enumerate(hdr):
    c = ws.cell(4, 2 + i, h); c.font = f_b; c.fill = hd; c.border = bd; c.alignment = ctr
ws.merge_cells("E3:F3"); ws["E3"] = "2026년 (예상)"; ws.merge_cells("G3:R3"); ws["G3"] = "2027년 (사업계획)"
for c in ("E3", "G3"): ws[c].font = f_b; ws[c].fill = hd; ws[c].alignment = ctr
for c in range(5, 19): ws.cell(3, c).border = bd

# 예시 값 (2027년은 '재료비 동가 예시' 파일 값, 2026년 11·12월은 임의 예시)
ex = {
 "tot": [750, 800, 800, 800, 900, 1000, 1200, 1000, 1200, 700, 1000, 1500, 800, 700],
 "sq":  [500, 550, 600, 500, 700, 900, 500, 400, 200, 0, 0, 0, 0, 0],
 "sl":  [12800, 13000, 13000, 14000, 14000, 14500, 13000, 12000, 13500, None, None, None, None, None],
 "sp":  [520, 520, 520, 510, 510, 515, 512, 500, 520, None, None, None, None, None],
 "jq":  [100, 100, 100, 100, 100, 100, 100, 100, 100, 0, 0, 0, 0, 0],
 "jl":  [12900, 13100, 12000, 14100, 13900, 14500, 13100, 12200, 13600, None, None, None, None, None],
 "jp":  [450, 450, 450, 450, 430, 430, 450, 450, 450, None, None, None, None, None],
 "nl":  [13200, 13300] + [14500] * 12,
 "np":  [500, 500] + [520] * 12,
 "fx":  [1380, 1380] + [1360] * 12,
}
R = {}
rows = [
 ("tot", "제품", "제품 총동량", "톤", "in", "생산계획 '생산/자체 소계' 동량 (HTML은 생산계획 값을 사용하고 이 행과 대조)"),
 ("sq", "SCR", "선물잔량", "톤", "in", ""),
 ("sl", "", "선물LME (가중평균)", "$/톤", "in", ""),
 ("sp", "", "프리미엄 (가중평균)", "$/톤", "in", ""),
 ("su", "", "잔량 단가 (LME+프리미엄)", "$/톤", "f", ""),
 ("jq", "JCR", "선물잔량", "톤", "in", ""),
 ("jl", "", "선물LME (가중평균)", "$/톤", "in", ""),
 ("jp", "", "프리미엄 (가중평균)", "$/톤", "in", ""),
 ("ju", "", "잔량 단가 (LME+프리미엄)", "$/톤", "f", ""),
 ("nq", "비선물", "비선물 동량", "톤", "f", "제품 총동량 − SCR − JCR (음수면 빨간색)"),
 ("nl", "", "LME", "$/톤", "in", "26년 11·12월: 예상 LME / 27년: 사업계획 지침 LME"),
 ("np", "", "프리미엄", "$/톤", "in", "26년 11·12월: 예상 프리미엄 / 27년: 지침 프리미엄"),
 ("nu", "", "비선물 단가", "$/톤", "f", ""),
 ("al", "평균", "평균 LME (프리미엄 제외)", "$/톤", "f", "(SCR량×LME + JCR량×LME + 비선물량×LME) ÷ 총동량"),
 ("ap", "", "평균 프리미엄 (참고)", "$/톤", "f", "물량 가중평균 – 참고용"),
 ("tp", "", "계획 적용 프리미엄", "$/톤", "f", "목표 프리미엄(E2)과 같은 값"),
 ("ll", "", "평균 LME (프리미엄 포함)", "$/톤", "f", ""),
 ("fx", "", "환율", "원/$", "in", "사업계획 지침 환율 (헷징 환율은 사용하지 않음)"),
 ("bp", "동가", "매입동가", "원/kg", "f", "평균 LME(프리미엄 포함) × 환율 ÷ 1,000"),
 ("ap2", "", "적용동가 (전월 매입동가)", "원/kg", "f", "당월 적용동가 = 전월 매입동가 (27년 1월 = 26년 12월)"),
]
r0 = 5
for i, (k, g, name, unit, kind, note) in enumerate(rows): R[k] = r0 + i
for k, g, name, unit, kind, note in rows:
    r = R[k]
    for c, v in ((2, g), (3, name), (4, unit)):
        x = ws.cell(r, c, v); x.font = f_b if c == 2 else f_c; x.border = bd
    ws.cell(r, 4).alignment = ctr
    ws.cell(r, 19, note).font = f_n
    for i in range(14):
        col = L(C0 + i); x = ws.cell(r, C0 + i); x.border = bd
        x.number_format = "#,##0.0" if unit == "톤" else "#,##0"
        if kind == "in":
            x.value = ex[k][i]; x.font = f_in; x.fill = yin
            continue
        rr = lambda kk: f"{col}{R[kk]}"
        f = {
         "su": f"=IF(N({rr('sq')})=0,0,{rr('sl')}+{rr('sp')})",
         "ju": f"=IF(N({rr('jq')})=0,0,{rr('jl')}+{rr('jp')})",
         "nq": f"={rr('tot')}-N({rr('sq')})-N({rr('jq')})",
         "nu": f"={rr('nl')}+{rr('np')}",
         "al": f"=IFERROR((N({rr('sq')})*N({rr('sl')})+N({rr('jq')})*N({rr('jl')})+{rr('nq')}*{rr('nl')})/{rr('tot')},0)",
         "ap": f"=IFERROR((N({rr('sq')})*N({rr('sp')})+N({rr('jq')})*N({rr('jp')})+{rr('nq')}*{rr('np')})/{rr('tot')},0)",
         "tp": "=$E$2",
         "ll": f"={rr('al')}+{rr('tp')}",
         "bp": f"={rr('ll')}*{rr('fx')}/1000",
         "ap2": (f"={L(C0 + i - 1)}{R['bp']}" if i >= 1 else None),
        }[k]
        x.value = f; x.font = f_c
        if k in ("bp", "ap2"): x.fill = key; x.font = Font(name=F, size=9, bold=True)
        if k == "ap2" and i == 0: x.fill = gry
    if kind == "in": pass
ws.cell(R["ap2"], C0).value = None
ws.cell(R["ap2"], 19).value = "당월 적용동가 = 전월 매입동가 (26년 11월 적용동가는 사용하지 않음)"
red = PatternFill("solid", fgColor="F8CBAD")
ws.conditional_formatting.add(f"E{R['nq']}:R{R['nq']}", FormulaRule(formula=[f"E{R['nq']}<0"], fill=red))
ws.freeze_panes = "E5"
n0 = R["ap2"] + 2
for i, t in enumerate([
    "※ 노란 칸만 입력합니다. 2026년 11·12월은 예상 값(비선물: 예상 LME·예상 프리미엄), 2027년은 선물잔량·지침 값.",
    "※ 재료비 HTML에는 4행(머리글)부터 마지막 행까지 붙여넣습니다. 항목 이름(C열)으로 행을 찾으므로 행 순서가 바뀌어도 됩니다.",
    "※ 예시 값: 2027년 SCR·JCR·총동량은 '재료비 동가 예시' 파일 값, 2026년 11·12월은 임의 예시입니다.",
    "※ 재고 반영은 이후 단계에서 추가합니다."]):
    ws.cell(n0 + i, 2, t).font = Font(name=F, size=9, color="C00000")

# ---- 상품재료비율
g = wb.create_sheet("상품재료비율")
g.column_dimensions["A"].width = 2
for col, w in zip("BCDE", [14, 16, 26, 40]): g.column_dimensions[col].width = w
g["B1"] = "상품재료비율 (제품군별 1개)"; g["B1"].font = f_t
for i, h in enumerate(["제품군", "상품재료비율", "산출 기준 (예: 26년 1~10월 실적)", "비고"]):
    c = g.cell(3, 2 + i, h); c.font = f_b; c.fill = hd; c.border = bd; c.alignment = ctr
for i, nm in enumerate(["상선용", "해양용", "이동용", "전력용"]):
    r = 4 + i
    g.cell(r, 2, nm).font = f_c
    for c in range(2, 6): g.cell(r, c).border = bd
    x = g.cell(r, 3); x.fill = yin; x.font = f_in; x.number_format = "0.0%"
    y = g.cell(r, 4); y.fill = yin; y.font = f_in
g["B9"] = "※ 상품(생산/OEM) 재료비 = 상품 매출액 × 상품재료비율 → 전액 부재료비로 표시"; g["B9"].font = Font(name=F, size=9, color="C00000")
g["B10"] = "※ 제품군 이름은 영업 매출계획 D열·생산계획 제품분류와 같아야 합니다. 제품군을 추가하면 행을 늘려 입력하세요."; g["B10"].font = Font(name=F, size=9, color="C00000")
# ---- 동량지수_부재료비 (B~R = 영업 양식 동량지수 시트와 같은 구조, U·V = 부재료비 계수)
ix = wb.create_sheet("동량지수_부재료비")
if SRC:
    import openpyxl
    s_idx = openpyxl.load_workbook(SRC)["동량지수_1~9월 매출품목"]
    for row in s_idx.iter_rows(min_row=1, max_row=s_idx.max_row):
        for c in row:
            if c.value is None or c.column > 18: continue
            n = ix.cell(c.row, c.column, c.value); n.font = f_c
            if c.row >= 12 and c.column >= 3 and not (isinstance(c.value, str) and c.value.startswith("=")): n.font = f_in
    for m in s_idx.merged_cells.ranges: ix.merge_cells(str(m))
    for k, d in s_idx.column_dimensions.items():
        if d.width and d.min:
            for cc in range(d.min, (d.max or d.min) + 1):
                if cc <= 18: ix.column_dimensions[L(cc)].width = d.width
for c in range(2, 19):
    for r in (10, 11): ix.cell(r, c).font = f_b; ix.cell(r, c).fill = hd; ix.cell(r, c).alignment = ctr
ix.merge_cells("U10:V10"); ix["U10"] = "부재료비 계수 (백만원/톤)"
ix["U11"] = "세선"; ix["V11"] = "태선"
for col in ("U", "V"):
    ix.column_dimensions[col].width = 11
    for r in (10, 11): ix[f"{col}{r}"].font = f_b; ix[f"{col}{r}"].fill = PatternFill("solid", fgColor="FFE699"); ix[f"{col}{r}"].alignment = ctr; ix[f"{col}{r}"].border = bd
last = max([r for r in range(12, ix.max_row + 1) if ix.cell(r, 3).value is not None] + [11])
for r in range(12, max(last, 40) + 1):
    for col in ("U", "V"):
        x = ix[f"{col}{r}"]; x.border = bd; x.number_format = "0.000"
        k = ix[f"K{r}"].value   # 윗행을 참조하는 행(예: USD 행 =+K12)은 계수도 윗행 참조
        if isinstance(k, str) and k.startswith("=+K"): x.value = f"=+{col}{k[3:]}"; x.font = f_c
        else: x.font = f_in; x.fill = yin
ix["B2"] = "동량지수 + 부재료비 계수"; ix["B2"].font = f_t
notes = ["※ B~R열: 영업 매출계획 양식의 '동량지수_1~9월 매출품목' 시트를 그대로 복사해 붙여넣습니다 (같은 구조).",
         "※ U·V열: 세선·태선 부재료비 계수(백만원/톤)를 입력합니다. B~R열을 다시 붙여넣어도 U·V열은 지워지지 않습니다 (행 순서가 바뀌면 계수도 같이 옮기세요).",
         "※ 거래처 행의 부재료비 단가 = 세선비중 × 세선 계수 + 태선비중 × 태선 계수 (키: 제품군+상세+거래선+통화, B열)"]
for i, t in enumerate(notes): ix.cell(4 + i, 2, t).font = Font(name=F, size=9, color="C00000")
wb.save(OUT); print("saved", OUT)
