"""재료비_기준양식(사용자 수정본)·생산계획_양식: 표시 서식만 변경 (값·수식 그대로).
- 생산계획: 동량·금액 숫자 칸 → 소수 3자리
- 재료비동가: 톤 단위 행(제품 총동량·선물잔량·비선물 동량) → 소수 3자리
- 동량지수_부재료비: 부재료비 계수 T~V열, 포장비 계수 W열 → 소수 6자리
"""
import sys, openpyxl
W = '/home/user/WORKSPACE/'
Q3 = '#,##0.000_);[Red]\\(#,##0.000\\)'
def prod_fmt(ws, r1, r2, c1, c2):
    n = 0
    for r in range(r1, r2 + 1):
        if ws.cell(r, 5).value not in ("동량", "금액"): continue
        for c in range(c1, c2 + 1):
            x = ws.cell(r, c)
            if isinstance(x.value, (int, float)) or (isinstance(x.value, str) and x.value.startswith("=")) or x.value is None:
                x.number_format = Q3; n += 1
    return n

# 1) 재료비_기준양식 (사용자 수정본)
f = W + '재료비_기준양식.xlsx'
wb = openpyxl.load_workbook(f)
n1 = prod_fmt(wb['생산계획'], 4, 29, 6, 22)
b = wb['재료비동가']; n2 = 0
for r in range(5, 25):
    if b.cell(r, 4).value == "톤":
        for c in range(5, 20): b.cell(r, c).number_format = '#,##0.000'; n2 += 1
x = wb['동량지수_부재료비']; n3 = 0
for r in range(12, 71):
    for c in range(20, 24):
        if x.cell(r, c).value is not None: x.cell(r, c).number_format = '0.000000'; n3 += 1
wb.save(f)
# 2) 생산계획_양식
g = W + '생산계획_양식.xlsx'
wp = openpyxl.load_workbook(g)
n4 = prod_fmt(wp['생산계획'], 4, 29, 6, 22)
wp.save(g)
print('서식 변경 칸 수 – 생산계획(기준양식):', n1, '재료비동가:', n2, '동량지수 계수:', n3, '생산계획_양식:', n4)
