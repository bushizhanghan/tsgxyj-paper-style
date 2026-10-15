# -*- coding: utf-8 -*-
"""build_xref.py —— 把正文手打的 [n] 替换为 Word 交叉引用域（REF _Ref<n> \\h）

机制：文末每条参考文献条目的 [n] 加隐藏书签（_Ref<n>）作为跳转目标，
正文中的 [n] 替换为 REF 域（外观不变、保持上标），Ctrl+点击可跳转到对应条目。

用法:
    python build_xref.py <论文.docx>

依赖: python-docx
纪律:
  1. 运行前务必备份原文件
  2. 本脚本对"纯文本 [n] 状态"的文档运行——若文档已含 REF 域，先跑 unfield_refs.py 退域
  3. rebuild 只能跑一次：重复运行会导致域嵌套；修复方法 = 先 unfield 再单次重建
  4. 运行后用 xref_verify.py 核验（文本零变化 + 域/书签配平 + 上标）
"""
import re, copy, sys
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

if len(sys.argv) < 2:
    sys.exit("用法: python build_xref.py <论文.docx>")
PATH = sys.argv[1]
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'
pat = re.compile(r"\[(\d{1,2})\]")


def make_run(rpr, text=None, instr=None, fld=None):
    r = OxmlElement('w:r')
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    if fld is not None:
        fc = OxmlElement('w:fldChar')
        fc.set(qn('w:fldCharType'), fld)
        r.append(fc)
    if instr is not None:
        it = OxmlElement('w:instrText')
        it.set(XML_SPACE, 'preserve')
        it.text = instr
        r.append(it)
    if text is not None:
        t = OxmlElement('w:t')
        t.set(XML_SPACE, 'preserve')
        t.text = text
        r.append(t)
    return r


def field_runs(rpr, n, label):
    """生成 REF 域的 5 个 run：begin / instrText / separate / 缓存文本 / end"""
    return [
        make_run(rpr, fld='begin'),
        make_run(rpr, instr=' REF _Ref%d \\h ' % n),
        make_run(rpr, fld='separate'),
        make_run(rpr, text=label),
        make_run(rpr, fld='end'),
    ]


def set_superscript(rpr):
    va = rpr.find(qn('w:vertAlign'))
    if va is None:
        va = OxmlElement('w:vertAlign')
        rpr.append(va)
    va.set(qn('w:val'), 'superscript')


doc = Document(PATH)
paras = doc.paragraphs

# 定位参考文献标题
REF_TITLE = None
for i, p in enumerate(paras):
    if p.text.strip() in ("参考文献", "参考文献：", "参考文献:"):
        REF_TITLE = i
        break
assert REF_TITLE is not None, "未找到参考文献标题"

# ---------- 1. 文末：给每条 [n] 加书签 ----------
bookmark_ids = {}
bid = 1000
ref_made = []
for i in range(REF_TITLE + 1, len(paras)):
    p = paras[i]
    m = pat.match(p.text.strip())
    if not m:
        continue
    n = int(m.group(1))
    run = p.runs[0]
    txt = run.text
    assert pat.match(txt), (i, txt[:30])
    end = m.end()
    head, tail = txt[:end], txt[end:]
    rpr = run._r.find(qn('w:rPr'))

    bm_start = OxmlElement('w:bookmarkStart')
    bm_start.set(qn('w:id'), str(bid))
    bm_start.set(qn('w:name'), '_Ref%d' % n)
    bm_end = OxmlElement('w:bookmarkEnd')
    bm_end.set(qn('w:id'), str(bid))
    bookmark_ids[n] = bid
    bid += 1

    new_head = make_run(rpr, text=head)
    new_tail = make_run(rpr, text=tail) if tail else None
    parent = run._r.getparent()
    parent.insert(parent.index(run._r), bm_start)
    run._r.addprevious(new_head)
    new_head.addnext(bm_end)
    if new_tail is not None:
        bm_end.addnext(new_tail)
    parent.remove(run._r)
    ref_made.append(n)

print("文末书签:", ref_made)

# ---------- 2. 正文：把 [n] 换成 REF 域 ----------
body_count = 0
for i in range(0, REF_TITLE):
    p = paras[i]
    if '[' not in p.text:
        continue
    for run in list(p.runs):
        txt = run.text or ''
        if '[' not in txt:
            continue
        hits = list(pat.finditer(txt))
        if not hits:
            continue
        rpr = run._r.find(qn('w:rPr'))
        new_rpr = copy.deepcopy(rpr) if rpr is not None else OxmlElement('w:rPr')
        set_superscript(new_rpr)

        pieces = []
        pos = 0
        for m in hits:
            if m.start() > pos:
                pieces.append(('t', txt[pos:m.start()]))
            pieces.append(('f', int(m.group(1))))
            pos = m.end()
        if pos < len(txt):
            pieces.append(('t', txt[pos:]))

        parent = run._r.getparent()
        anchor = run._r
        for kind, val in pieces:
            if kind == 't':
                els = [make_run(rpr, text=val)]
            else:
                assert val in bookmark_ids, "正文引用 [%d] 但文末无该条" % val
                els = field_runs(new_rpr, val, '[%d]' % val)
                body_count += 1
            for el in els:
                anchor.addprevious(el)
        parent.remove(run._r)

print("正文替换域数:", body_count)
doc.save(PATH)
print("已保存:", PATH)
print("下一步: 用 xref_verify.py 核验（文本零变化 + 域/书签配平 + 上标）")
