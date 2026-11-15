"""A단계 HTML 조립: tools/A_template.html + 1단계 HTML의 영업 매출계획 파서 + tools/A_xlsx.js + 예시 데이터.

사용 (저장소 루트에서): python tools/build_A_html.py <예시 JSON (gen_A_samples.py 출력)>
"""
import json, sys

S = sys.argv[1]
one = open('재료비계획_1단계.html', encoding='utf-8').read()
a = one.index('const COL = {'); b = one.index('const LS_KEY')
c = one.index('/* ---------------- 유틸 ---------------- */'); e = one.index('/* ---------------- 계산 ---------------- */')
shared = one[a:b] + "\n" + one[c:e]

t = open('tools/A_template.html', encoding='utf-8').read()
for dup in ('const $=id=>document.getElementById(id);\n',
            'function toast(t){const e=$("toast");e.textContent=t;e.classList.add("on");setTimeout(()=>e.classList.remove("on"),1800);}\n',
            'async function copyText(t){try{await navigator.clipboard.writeText(t);}catch(e){const a=document.createElement("textarea");a.value=t;document.body.appendChild(a);a.select();document.execCommand("copy");a.remove();}toast("복사했습니다. 엑셀에 붙여넣으세요");}\n'):
    t = t.replace(dup, '')
pit = '''function pasteInfoText(){
  const ck=parsed.check;
  return `데이터 ${parsed.rows.length}행 · 제외 ${parsed.skipped}행 · `+(parsed.header?`머리글 인식 ✓ (1월 동량 ${colName(actCol(M0))}열)`:"머리글 없음 → 기본 열 위치")+
    ` · 열 위치 검사 ${ck.checked?(ck.bad?"✗ 불일치":"✓ "+(ck.checked-ck.fails.length)+"/"+ck.checked):"-"}`;
}
'''
if 'function pasteInfoText' not in t:
    t = t.replace('/* --- 렌더 --- */', '/* --- 렌더 --- */\n' + pit)
t = t.replace('__SHARED__', shared).replace('__XLSX__', open('tools/A_xlsx.js', encoding='utf-8').read())
t = t.replace('__SAMPLES__', json.dumps(json.load(open(S)), ensure_ascii=False))
open('재료비계획_A단계.html', 'w', encoding='utf-8').write(t)
print('built', len(t))
