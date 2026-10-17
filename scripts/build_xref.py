# -*- coding: utf-8 -*-
"""build_xref.py -- 把正文手打的 [n] 替换为 Word 交叉引用域（REF _Ref<n> \\h）

机制：文末每条参考文献条目的 [n] 加隐藏书签（_Ref<n>）作为跳转目标，
正文中的 [n] 替换为 REF 域（外观不变、保持上标），Ctrl+点击可跳转到对应条目。

用法:
    python build_xref.py <论文.docx> [--dry-run] [--backup]
        --dry-run   只报告将替换/将跳过的标注，不写文件
        --backup    保存前先复制一份 <文件名>.bak.docx

依赖: python-docx
纪律:
   1. 运行前务必备份原文件（或用 --backup）
   2. 本脚本对"纯文本 [n] 状态"的文档运行——若文档已含 REF 域，先跑 unfield_refs.py 退域
   3. rebuild 只能跑一次：重复运行会导致域嵌套；修复方法 = 先 unfield 再单次重建
   4. 运行后用 xref_verify.py 核验（文本零变化 + 域/书签配平 + 上标）

已知边界（运行时会显式报告，不静默跳过）：
   * 只替换单个编号 [n]；范围式 [2-5] / 列举式 [2,3] 不会转换——
     Word 一个 REF 域只能指向一个书签，范围式需拆成 [2][3][4][5] 多个域
     （该刊正文标注以范围式为主流，见 references/citation-format.md；
      若为保体例可整体不做域化）
   * 编号上限 99（正则 \\d{1,2}）；超 99 条的文献表需先改 pat
"""
import re
import copy
import sys
import os
import shutil

# Windows 控制台默认 GBK 无法输出部分 UTF-8 字符，统一走 UTF-8
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

if len(sys.argv) < 2:
    sys.exit(__doc__)
_args = [a for a in sys.argv[1:] if not a.startswith("--")]
_flags = {a for a in sys.argv[1:] if a.startswith("--")}
PATH = _args[0]
if not os.path.isfile(PATH):
    sys.exit(f"找不到文件：{PATH}")
DRY = "--dry-run" in _flags
if not DRY and "--backup" in _flags:
    bak = os.path.splitext(PATH)[0] + ".bak.docx"
    shutil.copy2(PATH, bak)
    print(f"已备份: {bak}")

XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'
pat = re.compile(r"\[(\d{1,2})\]")
# 范围式 / 列举式（仅用于报告，不转换）
pat_range = re.compile(r"\[(\d{1,2})\s*[-–—]\s*\d{1,2}\]")
pat_list = re.compile(r"\[\d{1,2}(?:\s*[,，]\s*\d{1,2})+\]")


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
if REF_TITLE is None:
    sys.exit("未找到参考文献标题（需要独立的「参考文献」段落）")

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
    if not pat.match(txt):
        print(f"[警告] 段 {i + 1} 编号与文本错位，跳过：{txt[:30]!r}")
        continue
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
skipped_range = []
skipped_list = []
for i in range(0, REF_TITLE):
    p = paras[i]
    if '[' not in p.text:
        continue
    for m_ in pat_range.finditer(p.text):
        skipped_range.append(m_.group(0))
    for m_ in pat_list.finditer(p.text):
        skipped_list.append(m_.group(0))
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
                if val not in bookmark_ids:
                    print(f"[警告] 正文引用 [{val}] 但文末无该条，保持原样")
                    els = [make_run(rpr, text='[%d]' % val)]
                else:
                    els = field_runs(new_rpr, val, '[%d]' % val)
                    body_count += 1
            for el in els:
                anchor.addprevious(el)
        parent.remove(run._r)

print("正文替换域数:", body_count)
if skipped_range:
    print(f"[跳过] 范围式标注 {len(skipped_range)} 处：{skipped_range[:6]}"
          f"{' …' if len(skipped_range) > 6 else ''}")
    print("        （一个 REF 域只能指向一个书签；范围式需拆成多个域才可域化。")
    print("         该刊正文以范围式为主流——不为体例可整体不域化，或只对单引编号生效）")
if skipped_list:
    print(f"[跳过] 列举式标注 {len(skipped_list)} 处：{skipped_list[:6]}"
          f"{' …' if len(skipped_list) > 6 else ''}")

if DRY:
    print("（--dry-run：未写文件）")
else:
    doc.save(PATH)
    print("已保存:", PATH)
    print("下一步: 用 xref_verify.py 核验（文本零变化 + 域/书签配平 + 上标）")
