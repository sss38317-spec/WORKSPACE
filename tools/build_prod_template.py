"""생산계획 양식: 26년 '11~12월예상'을 11월·12월 예상으로 분리 (원본 G열 → G·H열).

사용: python build_prod_template.py <원본 생산계획 예시양식.xlsx> <출력.xlsx>
"""
import re, sys
from copy import copy
import openpyxl
from openpyxl.utils import get_column_letter as L, column_index_from_string as CI
from openpyxl.styles import Font, PatternFill

SRC, OUT = sys.argv[1], sys.argv[2]
src = openpyxl.load_workbook(SRC); s = src.active
wb = openpyxl.Workbook(); ws = wb.active; ws.title = "생산계획"
SH = lambda c: c + 1 if c >= 8 else c          # 원본 H열 이후 한 칸 이동 (G → G·H 분리)

def shift_formula(f):
    return re.sub(r"(\$?)([A-Z]{1,2})(\$?\d+)", lambda m: m.group(1) + L(SH(CI(m.group(2)))) + m.group(3), f)

for row in s.iter_rows(min_row=1, max_row=s.max_row):
    for c in row:
        n = ws.cell(c.row, SH(c.column))
        v = c.value
        if isinstance(v, str) and v.startswith("="): v = shift_formula(v)
        n.value = v
        if c.has_style:
            n.font = copy(c.font); n.fill = copy(c.fill); n.border = copy(c.border)
            n.number_format = c.number_format; n.alignment = copy(c.alignment)
        if c.column == 7:                          # 원본 G(11~12월예상) → G(11월)·H(12월)
            h = ws.cell(c.row, 8)
            if c.has_style:
                h.font = copy(c.font); h.fill = copy(c.fill); h.border = copy(c.border)
                h.number_format = c.number_format; h.alignment = copy(c.alignment)
            if isinstance(c.value, (int, float)):  # 예시 값: 반씩 나눔
                n.value = round(c.value / 2, 3); h.value = round(c.value - n.value, 3)
            elif isinstance(c.value, str) and c.value.startswith("="):
                h.value = re.sub(r"\bG(\d+)", r"H\1", shift_formula(c.value))
for m in s.merged_cells.ranges:
    a, b = m.min_col, m.max_col
    a = a + 1 if a >= 8 else a
    b = b + 1 if b >= 7 else b
    ws.merge_cells(start_row=m.min_row, end_row=m.max_row, start_column=a, end_column=b)
ws["G3"] = "11월 예상"; ws["H3"] = "12월 예상"; ws["I3"] = "합 계"
# 합계(I) = 1~10월 + 11월 + 12월
for r in range(4, s.max_row + 1):
    if ws.cell(r, 6).value is not None and ws.cell(r, 5).value in ("동량", "금액"):
        f = s.cell(r, 8).value
        if isinstance(f, str) and f.replace(" ", "") in (f"=+F{r}+G{r}", f"=F{r}+G{r}"):
            ws.cell(r, 9).value = f"=+F{r}+G{r}+H{r}"
for k, d in s.column_dimensions.items():
    if d.width and d.min:
        for c in range(d.min, (d.max or d.min) + 1): ws.column_dimensions[L(SH(c))].width = d.width
ws.column_dimensions["H"].width = ws.column_dimensions["G"].width or 12
ws["B31"] = "※ 26년 11월·12월 예상을 나눠 입력합니다 (27년 1월 적용동가 = 26년 12월 매입동가). G·H열 예시 값은 원본 11~12월 합계를 반씩 나눈 값입니다."
ws["B32"] = "※ 생산/자체 = 제품, 생산/OEM = 상품. 재료비 계산 HTML은 이 시트를 머리글(2~3행)부터 붙여넣어 사용합니다."
for k in ("B31", "B32"): ws[k].font = Font(name="맑은 고딕", size=9, color="C00000")
wb.save(OUT)
print("saved", OUT)
