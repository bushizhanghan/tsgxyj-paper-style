# -*- coding: utf-8 -*-
"""xref_verify.py —— 验证交叉引用域转换结果

检查四件事：
  1. 文本零变化（与转换前备份对比，可选）
  2. 域/书签结构配平（fldChar begin=separate=end；每个被引编号都有书签）
  3. 正文 [n] 全部为上标；文末编号保持非上标
  4. 书签只包住 [n] 本身，不含后续文字

用法:
    python xref_verify.py <转换后.docx> [转换前备份.docx]
（第二个参数可选——提供时做全文文本零变化对比）
"""
import re, sys
from docx import Document
from docx.oxml.ns import qn

# Windows 控制台默认 GBK 无法输出部分 UTF-8 字符，统一走 UTF-8
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

if len(sys.argv) < 2:
    sys.exit("用法: python xref_verify.py <转换后.docx> [转换前备份.docx]")
NEW = sys.argv[1]
BAK = sys.argv[2] if len(sys.argv) > 2 else None


def full_text(path):
    d = Document(path)
    parts = [p.text for p in d.paragraphs]
    for t in d.tables:
        for row in t.rows:
            for c in row.cells:
                parts += [p.text for p in c.paragraphs]
    return "\n".join(parts)


if BAK:
    a, b = full_text(BAK), full_text(NEW)
    same = a == b
    print("全文文本是否一致:", same, "| 长度", len(a), len(b))
    if not same:
        for i, (x, y) in enumerate(zip(a.split("\n"), b.split("\n"))):
            if x != y:
                print("首个差异行", i)
                print(" 旧:", repr(x[:120]))
                print(" 新:", repr(y[:120]))
                break

d = Document(NEW)
xml = d.element.xml
n_begin = xml.count('w:fldCharType="begin"')
n_end = xml.count('w:fldCharType="end"')
n_sep = xml.count('w:fldCharType="separate"')
print("fldChar begin/separate/end:", n_begin, n_sep, n_end, "| 配平:", n_begin == n_end == n_sep)

instrs = re.findall(r"<w:instrText[^>]*>([^<]+)</w:instrText>", xml)
print("域数量:", len(instrs), "| 样例:", instrs[:3])
refs = sorted({int(m.group(1)) for s in instrs for m in [re.search(r"_Ref(\d+)", s)] if m})
print("被引用编号:", refs)

bms = re.findall(r'<w:bookmarkStart w:id="(\d+)" w:name="(_Ref\d+)"', xml)
names = sorted({n for _, n in bms})
print("书签名数量:", len(names))
print("引用编号是否都有书签:", set(refs) <= set(int(n[4:]) for n in names))

# 定位参考文献标题（上标检查只覆盖正文段；文末编号本就应非上标）
REF_TITLE = None
for i, p in enumerate(d.paragraphs):
    if p.text.strip() in ("参考文献", "参考文献：", "参考文献:"):
        REF_TITLE = i
        break
if REF_TITLE is None:
    print("[警告] 未找到参考文献标题段，上标检查将覆盖全文（含文献表，可能误报）")

# 域结果文本与上标检查（仅正文）
ok_sup = 0
bad = []
for p in d.paragraphs[:REF_TITLE if REF_TITLE is not None else len(d.paragraphs)]:
    runs = p.runs
    txt = "".join(r.text or "" for r in runs)
    for m in re.finditer(r"\[(\d{1,2})\]", txt):
        n = int(m.group(1))
        pos = 0
        for r in runs:
            t = r.text or ""
            if pos <= m.start() < pos + len(t):
                rpr = r._r.find(qn('w:rPr'))
                va = rpr.find(qn('w:vertAlign')) if rpr is not None else None
                sup = va is not None and va.get(qn('w:val')) == 'superscript'
                if sup:
                    ok_sup += 1
                else:
                    bad.append((m.group(0), p.text[:30]))
                break
            pos += len(t)
print("正文上标 [n] 计数:", ok_sup, "| 非上标:", bad[:5] if bad else "无")

# 文末编号应保持非上标（仅文献表段）
tail_base = 0
if REF_TITLE is not None:
    for p in d.paragraphs[REF_TITLE + 1:]:
        m = re.match(r"\[(\d{1,2})\]", p.text.strip())
        if not m:
            continue
        r = p.runs[0]
        rpr = r._r.find(qn('w:rPr'))
        va = rpr.find(qn('w:vertAlign')) if rpr is not None else None
        if va is None or va.get(qn('w:val')) != 'superscript':
            tail_base += 1
print("文末编号保持非上标:", tail_base)

# 书签只包住 [n]
bm_xml = re.findall(r'<w:bookmarkStart w:id="\d+" w:name="_Ref\d+"/>.{0,80}?<w:bookmarkEnd w:id="\d+"/>', xml, re.S)
print("书签-结束配对样例:", bm_xml[0][:160] if bm_xml else "无")
print()
print("结论判定：文本一致 + begin/end 配平 + 引用均有书签 + 正文全上标 + 文末非上标 = 转换合格")
