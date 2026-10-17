# -*- coding: utf-8 -*-
"""unfield_refs.py -- 把正文 REF 域还原为纯文本 [n]（保留上标），移除 _Ref 书签

用途：编辑带引用的段落之前先"退域"为纯文本，改完后再用 build_xref.py 重建域。
这样可避免在域结构上直接编辑导致的域损坏/嵌套。

用法:
    python unfield_refs.py <论文.docx> [--backup]
        --backup   保存前先复制一份 <文件名>.bak.docx

依赖: python-docx
纪律:
   1. 运行前务必备份（或用 --backup）
   2. 退域后正文变为纯文本 [n]——此时可安全增删改文字
   3. 改完务必用 build_xref.py 重建域（只能跑一次）
"""
import copy
import sys
import os
import shutil
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Windows 控制台默认 GBK 无法输出部分 UTF-8 字符，统一走 UTF-8
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

_args = [a for a in sys.argv[1:] if not a.startswith("--")]
_flags = {a for a in sys.argv[1:] if a.startswith("--")}
if not _args:
    sys.exit(__doc__)
PATH = _args[0]
if not os.path.isfile(PATH):
    sys.exit(f"找不到文件：{PATH}")
if "--backup" in _flags:
    bak = os.path.splitext(PATH)[0] + ".bak.docx"
    shutil.copy2(PATH, bak)
    print(f"已备份: {bak}")
doc = Document(PATH)


def fld_type(run):
    fc = run._r.find(qn('w:fldChar'))
    if fc is None:
        return None
    return fc.get(qn('w:fldCharType'))


def run_text(run):
    t = run._r.find(qn('w:t'))
    return t.text if t is not None else ''


def rebuild_paragraph(p):
    """把域结构折叠成单个文本 run，其余 run 原样保留"""
    runs = p.runs
    i = 0
    out = []  # (text, rpr_element_or_None)
    touched = 0
    while i < len(runs):
        r = runs[i]
        if fld_type(r) == 'begin':
            # 收集：begin -> instrText -> separate -> text(runs) -> end
            j = i + 1
            texts = []
            rpr = r._r.find(qn('w:rPr'))
            separated = False
            while j < len(runs):
                ft = fld_type(runs[j])
                if ft == 'separate':
                    separated = True
                elif ft == 'end':
                    break
                elif separated and fld_type(runs[j]) is None:
                    texts.append(run_text(runs[j]))
                j += 1
            out.append((''.join(texts), rpr))
            touched += 1
            i = j + 1
        else:
            out.append((r.text or '', r._r.find(qn('w:rPr'))))
            i += 1
    if touched == 0:
        return 0

    parent = p._p
    for r in list(p.runs):
        parent.remove(r._r)
    for txt, rpr in out:
        if txt == '':
            continue
        nr = OxmlElement('w:r')
        if rpr is not None:
            nr.append(copy.deepcopy(rpr))
        t = OxmlElement('w:t')
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        t.text = txt
        nr.append(t)
        parent.append(nr)
    return touched


total = 0
for p in doc.paragraphs:
    total += rebuild_paragraph(p)
print("折叠域数:", total)

# 移除书签
xml_root = doc.element.body
bm_s = xml_root.findall('.//' + qn('w:bookmarkStart'))
bm_e = xml_root.findall('.//' + qn('w:bookmarkEnd'))
cnt = 0
for bm in list(bm_s) + list(bm_e):
    name = bm.get(qn('w:name')) or ''
    if name.startswith('_Ref'):
        bm.getparent().remove(bm)
        cnt += 1
print("移除书签元素:", cnt)

doc.save(PATH)
print("已保存（纯文本状态）——改完文字后请用 build_xref.py 重建域")
