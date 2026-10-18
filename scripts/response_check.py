# -*- coding: utf-8 -*-
"""response_check.py —— 审稿回复逐条回应完整性核对器（模式 F 返修阶段）

返修阶段最怕的不是改得不好，而是**漏回**：审稿人提了 5 条意见，回复信只
回应了 4 条；或者每条都写「感谢指正，已修改」却不写改在哪儿——审稿人
无法核对，只能怀疑作者没改。这类问题在 word 里人工点数是极易漏掉的。
本脚本把「审稿意见文件」与「回复信文件」对照，核验三件事：

  · 条数对齐：意见条数 vs 回应区块数。按**编号集合**比对（不只是计数），
    直接报出漏回了哪几条、回复信多出哪几条；
  · 三件套齐全：每个回应区块是否含「意见（复述）/ 回应 / 修改位置」
    ——对应 references/response-template.md 规定的逐条回应制格式；
  · 位置精确：修改位置是否精确到节/段（第 X 节第 X 段、2.1 节、表 3…），
    只写「已修改」「见修改稿」的报待核对，让审稿人能核对。

输入支持 .docx（段落 + 表格，表格按表头「审稿意见/回复/修改位置」分列）
与 .md / .txt（含 markdown 表格与引用块）。

用法:
    python response_check.py --reviews <审稿意见.docx|md|txt> --response <回复信.docx|md|txt>
    python response_check.py --reviews 意见.docx --response 回复.docx --top 20 --json
退出码：0 = 无硬伤（漏回/缺件/位置缺失）；1 = 存在硬伤；
        2 = 参数不全或文件不存在。

局限（词面与结构启发式，不是语义理解）：
  · 条目切分依赖「意见 N / 问题 N」编号、表格行或「第 N 条」标记；
    自由书写的回复信只能报「未识别到逐条结构」，交人工核对；
  · 位置精确度按关键词判断（第 X 节/段/页/表/图），**不验证该位置是否
    真的改了**——「改在哪儿说得准不准」要打开修改稿核对；
  · 主题词重叠只提示「疑似答非所问」，不是语义判定。
"""
import os
import re
import sys
import json

# Windows 控制台 GBK 保护
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

# ---------------- 常量 ----------------
NUM = r"[0-9０-９一二三四五六七八九十]"
CN_DIGITS = {"零": 0, "一": 1, "二": 2, "两": 2, "三": 3, "四": 4, "五": 5,
             "六": 6, "七": 7, "八": 8, "九": 9, "十": 10}
_FW2HW = {ord("０") + i: ord("0") + i for i in range(10)}

# 条目起始标记：「意见 1：」「问题2」「审稿人1，意见3」「Q1」「第一条：」
HEAD_LABELS = r"意见|问题|点评|疑问|质疑|建议|回复|comment|question|point|item|q"
ITEM_HEAD_RE = re.compile(
    r"^(?:审稿人?\s*\d+\s*[，,、]\s*)?(?:" + HEAD_LABELS + r")\s*"
    r"(" + NUM + r"{1,3})\s*(?=[：:（(（【\[]|\s*[，,。.\s]|$)", re.I)
ORD_HEAD_RE = re.compile(r"^第\s*(" + NUM + r"{1,3})\s*[条点个项]\s*(?=[：:])")
GENERIC_HEAD_RE = re.compile(
    r"^(?:[（(]\s*(" + NUM + r"{1,3})\s*[)）]|(" + NUM + r"{1,3})\s*[、.．)])\s*")

# 三件套字段标签（锚定行首匹配，长模式在前，避免「位置」吃掉「修改位置」）
LABEL_RES = [
    ("location", r"(?:修改|改动|变更)?(?:位置|之处|处所|所在)"
                 r"|(?:修改|改动)(?:段落|章节|部分|地方)"),
    ("response", r"(?:作者)?(?:回应|回复|答复|说明|情况|处理)(?:如下)?"
                 r"|(?:修改|改动)(?:说明|情况|回复|回应)|答"),
    ("opinion", r"(?:审稿)?(?:意见|问题|点评|原文)"),
]
LABEL_TEXT = {"opinion": "意见", "response": "回应", "location": "修改位置"}
FIELD_NAME = {"opinion": "意见（复述）", "response": "回应", "location": "修改位置"}

# 回应线索（无标签信件的内容判定用）：出现这些词说明作者确实作了回应
RESPONSE_CUES = (r"感谢(?:指正|建议|意见|专家|审稿)|已(?:予以|经|作)?(?:修改|完善|补充|调整|更正|删除|压缩|改写|核实|核对|修改)"
                r"|同意|部分同意|不(?:同意|赞同)|商榷|据此(?:修改|调整|补充)|修订后|改为|调整为|现(?:已|将)|补充了|增加了")

# 位置标注：精确（到节/段/页/表/图）与笼统两档
LOC_SPECIFIC = [
    (r"第\s*[0-9０-９一二三四五六七八九十百]+\s*[节章段页部分]", "第X节/章/段/页"),
    (r"(?:第|见|在|详见|参见|修改后见|现见于)\s*\d+(?:\s*\.\s*\d+)+\s*[节章部分]?", "2.1 节号"),
    (r"\d+(?:\s*\.\s*\d+)+\s*[节章部分]", "2.1 节号"),
    (r"[表图]\s*[0-9０-９一二三四五六七八九十百]+", "表/图号"),
    (r"摘\s*要|引\s*言|结\s*语|参考文献|关\s*键\s*词|研究设计|对象与方法|对策", "篇章名"),
    (r"[（(]\s*[0-9０-９]{1,3}\s*[)）]", "编号项"),
]
LOC_VAGUE = [
    (r"已(?:予以|经)?(?:修改|完善|补充|调整|更正|删除|改写|压缩)", "已修改（无位置）"),
    (r"见(?:修改|修订)?(?:稿|后稿)", "见修改稿（无位置）"),
    (r"全文|通篇|全篇|各(?:个)?(?:章节|部分|段落)|相关(?:章节|部分|内容|表述)|相应(?:章节|部分)",
     "笼统范围"),
]

# 收尾信号（从最后一个区块中拆出，不当成条目内容）
CLOSING_RE = re.compile(r"^(此致|敬礼|再次感谢|再次致谢|特此(?:回复|说明|函复)|盼复|以上说明|再次恳请)")
# 审稿人分组头（仅作提示，不参与切分）
REVIEWER_HEAD_RE = re.compile(
    r"^(?:审稿人|评审专家|外审专家|专家|复审|一审|二审)\s*[0-9０-９一二三四五六七八九十]*\s*[：:意见]?$")

# 主题词重叠：答非所问提示用（词面启发式）
OVERLAP_STOP = set("的了在和是为对从于以而中上下后前内外等与被将把让向往由这那之也都又再更最很较比越显"
                   "感谢指正建议修改补充已经本文我们笔者研究文章论文如下部分方面问题情况内容")
OVERLAP_MIN_HIT = 2   # 共享 n-gram 少于该值 → 提示「疑似答非所问」


def to_int(tok):
    """编号 token → int（支持阿拉伯/全角/中文数字）；无法解析返回 None。"""
    if not tok:
        return None
    s = str(tok).strip().translate(_FW2HW)
    if re.fullmatch(r"\d{1,4}", s):
        return int(s)
    if re.fullmatch(r"[零一二两三四五六七八九十]{1,7}", s):
        if "十" in s:
            left, _, right = s.partition("十")
            hi = CN_DIGITS.get(left, 1) if left else 1
            lo = CN_DIGITS.get(right, 0) if right else 0
            return hi * 10 + lo
        total = 0
        for ch in s:
            v = CN_DIGITS.get(ch)
            if v is None:
                return None
            total = total * 10 + v
        return total
    return None


def norm_line(line):
    """行归一：去 markdown 强调符/引用符/首尾空白（保留可见文本）。"""
    t = (line or "").replace("*", "").replace("`", "").replace("~~", "")
    t = re.sub(r"^[>\s]+", "", t)
    return t.strip()


# ---------------- 文件读取 ----------------
def _dedup_cells(cells):
    """合并单元格会整行重复，去掉与前一格相同的连续重复。"""
    out = []
    for c in cells:
        if out and out[-1] == c:
            continue
        out.append(c)
    return out


def _header_cols(header_cells):
    """表头 → {列号: 字段}；另识别「序号」列。返回 (cols, no_col)。"""
    cols, no_col = {}, None
    for ci, c in enumerate(header_cells):
        t = re.sub(r"[\s*（）()【】\[\]]", "", c or "")
        if not t:
            continue
        if re.fullmatch(r"(序号|编号|序|条号|no\.?|number)", t, re.I):
            no_col = ci
            continue
        for field, pat in LABEL_RES:
            if field in cols:
                continue
            if re.search(r"^" + pat + r"$", t) or (len(t) <= 8 and re.search(pat, t)):
                cols[ci] = field
                break
    return cols, no_col


def _row_lines(cells, cols, no_col, row_no):
    """把一个表格数据行转成带字段标签的行；返回 (lines, head_no)。"""
    cells = _dedup_cells([re.sub(r"\s+", " ", (c or "").strip()) for c in cells])
    lines, head_no = [], None
    if no_col is not None and no_col < len(cells):
        head_no = to_int(cells[no_col])
    if cols and len(cols) >= 2:
        for ci, field in sorted(cols.items()):
            if ci < len(cells) and cells[ci]:
                lines.append(f"{LABEL_TEXT[field]}：{cells[ci]}")
    else:
        joined = " ".join(c for c in cells if c)
        if joined:
            lines.append(joined)
    if not lines:
        return [], None
    if head_no is None:
        head_no = row_no
    return lines, head_no


def _read_md_table_block(block_lines):
    """markdown 表格块 → (lines, head_nos)。"""
    rows = []
    for ln in block_lines:
        raw = [c.strip() for c in ln.strip().strip("|").split("|")]
        rows.append(raw)
    if len(rows) < 2:
        return [], []
    cols, no_col = _header_cols(rows[0])
    lines, heads = [], []
    for ri, cells in enumerate(rows[1:], start=1):
        if all(re.fullmatch(r":?-{2,}:?", c or "-") for c in cells if c):
            continue  # 分隔行
        ls, hn = _row_lines(cells, cols, no_col, ri)
        lines.extend(ls)
        heads.extend([hn] + [None] * (len(ls) - 1))
    return lines, heads


def read_paragraphs(path):
    """读取文件 → (lines, head_nos)。

    lines: 归一化后的文本行；head_nos: 与 lines 等长，非 None 表示该行是
    一个条目的起始（表格数据行），否则由 ITEM_HEAD_RE 判定。
    支持 .docx（正文顺序：段落与表格交错）与 .md/.txt（含 markdown 表格）。
    """
    ext = os.path.splitext(path)[1].lower()
    lines, heads = [], []
    if ext in (".docx", ".doc"):
        from docx import Document
        from docx.text.paragraph import Paragraph
        from docx.table import Table
        doc = Document(path)
        for child in doc.element.body.iterchildren():
            tag = child.tag.split("}")[-1]
            if tag == "p":
                t = norm_line(Paragraph(child, doc).text)
                if t:
                    lines.append(t)
                    heads.append(None)
            elif tag == "tbl":
                tbl = Table(child, doc)
                if not tbl.rows:
                    continue
                rows = [[c.text for c in r.cells] for r in tbl.rows]
                cols, no_col = _header_cols(rows[0])
                for ri, cells in enumerate(rows[1:], start=1):
                    ls, hn = _row_lines(cells, cols, no_col, ri)
                    lines.extend(ls)
                    heads.extend([hn] + [None] * (len(ls) - 1))
        return lines, heads
    # 文本类：md / txt
    raw = None
    for enc in ("utf-8-sig", "utf-8", "gbk"):
        try:
            with open(path, encoding=enc) as f:
                raw = f.read()
            break
        except (UnicodeDecodeError, LookupError):
            continue
    if raw is None:
        raise ValueError(f"无法解码文件：{path}")
    buf = []
    for line in raw.splitlines():
        t = norm_line(line)
        if not t:
            if buf:
                _flush_md_table(buf, lines, heads)
                buf = []
            continue
        if t.startswith("|"):
            buf.append(t)
            continue
        if buf:
            _flush_md_table(buf, lines, heads)
            buf = []
        lines.append(t)
        heads.append(None)
    if buf:
        _flush_md_table(buf, lines, heads)
    return lines, heads


def _flush_md_table(buf, lines, heads):
    ls, hs = _read_md_table_block(buf)
    lines.extend(ls)
    heads.extend(hs)


# ---------------- 条目切分 ----------------
def explicit_head_no(t):
    """显式条目头（意见 N / 问题 N / 第 N 条）→ 编号，否则 None。"""
    m = ITEM_HEAD_RE.match(t)
    if m:
        return to_int(m.group(1))
    m = ORD_HEAD_RE.match(t)
    if m:
        return to_int(m.group(1))
    return None


def generic_head_no(t):
    """通用编号头（1. / 1、/ （1））→ 编号，否则 None。"""
    m = GENERIC_HEAD_RE.match(t)
    if not m:
        return None
    return to_int(m.group(1) or m.group(2))


def split_items(lines, heads, kind="opinions"):
    """把行切成条目。返回 (items, prelude, mode)。

    items: [{"no": int|None, "par": 起始行号, "lines": [str]}]
    mode: "explicit"（编号/表格行）/ "generic"（通用编号）/ "para"（按段）/
          "none"（未识别到结构）
    """
    idxs = []
    for i, hn in enumerate(heads or []):
        if hn is not None:
            idxs.append((i, hn))
    mode = "explicit"
    if not idxs:
        for i, t in enumerate(lines):
            no = explicit_head_no(t)
            if no is not None:
                idxs.append((i, no))
    if not idxs and kind == "opinions":
        # 意见文件无编号：按段落切分（粘贴的审稿邮件常见），需人工确认
        for i, t in enumerate(lines):
            if generic_head_no(t) is not None and len(t) >= 12:
                idxs.append((i, generic_head_no(t)))
        mode = "generic" if idxs else "para"
        if mode == "para":
            idxs = [(i, None) for i, t in enumerate(lines) if len(t) >= 15]
    if not idxs:
        return [], list(lines), "none"

    items, prelude = [], list(lines[:idxs[0][0]])
    for k, (start, no) in enumerate(idxs):
        end = idxs[k + 1][0] if k + 1 < len(idxs) else len(lines)
        block = list(lines[start:end])
        # 收尾段从最后一个条目中拆出
        if k == len(idxs) - 1 and len(block) > 1:
            cut = len(block)
            for j in range(len(block) - 1, 0, -1):
                if CLOSING_RE.match(block[j]):
                    cut = j
                else:
                    break
            if cut < len(block):
                prelude.extend(block[cut:])
                block = block[:cut]
        items.append({"no": no, "par": start + 1, "lines": block})
    return items, prelude, mode


# ---------------- 字段识别 ----------------
def leading_label(t):
    """行首字段标签 → (field, 剩余文本)；无标签返回 (None, t)。"""
    for field, pat in LABEL_RES:
        m = re.match(r"^(?:" + pat + r")\s*[：:]", t)
        if m:
            return field, t[m.end():].strip()
    return None, t


def parse_block(block_lines):
    """把一个回应区块解析成三件套字段。"""
    fields = {"opinion": "", "response": "", "location": ""}
    cur = "opinion"
    for t in block_lines:
        field, rest = leading_label(t)
        if field:
            cur = field
            fields[cur] = (fields[cur] + " " + rest).strip() if fields[cur] else rest
            continue
        # 无标签行：默认续接上一字段；若含回应线索且回应字段还空着，判为回应
        if cur == "opinion" and not fields["response"] and re.search(RESPONSE_CUES, t):
            cur = "response"
        fields[cur] = (fields[cur] + " " + t).strip() if fields[cur] else t
    # 末轮兜底：三件套全空但整块有回应线索 → 视为回应
    if not fields["response"] and re.search(RESPONSE_CUES, " ".join(block_lines)):
        fields["response"] = " ".join(block_lines)
    # 末轮兜底：区块有内容却未归入任何字段（表格列不齐/自由书写）→
    # 按回应线索判定归属：有线索并入回应，否则并入意见（复述）
    if not fields["opinion"] and not fields["response"]:
        body = " ".join(block_lines)
        if re.search(RESPONSE_CUES, body):
            fields["response"] = body
        else:
            fields["opinion"] = body
    # 表格行常见：区块只有「意见 + 修改位置」两列而无回应列——只要出现回应
    # 线索词（已补充/同意…）就还原出回应，避免把「缺回应」误判为硬伤
    if not fields["response"]:
        for t in block_lines:
            if leading_label(t)[0] is None and re.search(RESPONSE_CUES, t):
                fields["response"] = (fields["response"] + " " + t).strip()
    return fields


def precision_of(text):
    """位置标注精确度 → ("precise"|"vague"|"missing", 命中证据)。"""
    t = (text or "").strip()
    if not t:
        return "missing", "位置标注为空"
    for pat, name in LOC_SPECIFIC:
        m = re.search(pat, t)
        if m:
            return "precise", f"{name}「{m.group(0).strip()}」"
    for pat, name in LOC_VAGUE:
        m = re.search(pat, t)
        if m:
            return "vague", f"{name}「{m.group(0).strip()}」"
    return "vague", "未出现节/段/页/表/图定位词"


def ngrams(s, ns=(2, 3)):
    """内容 n-gram 集合（去标点，过滤含停用字的片段）。"""
    s = re.sub(r"[^0-9A-Za-z\u4e00-\u9fff]", "", s or "")
    out = set()
    for i in range(len(s)):
        for n in ns:
            if i + n <= len(s):
                w = s[i:i + n]
                if any(c in OVERLAP_STOP for c in w):
                    continue
                out.add(w)
    return out


def topic_overlap(a, b):
    return len(ngrams(a) & ngrams(b))


def trunc(s, n):
    return s if len(s) <= n else s[:n] + "…"


def align_items(op_items, rs_items):
    """意见条目 ↔ 回应区块对齐。返回 (pairs, missing, extra, by_number)。"""
    op_nos = [it["no"] for it in op_items]
    rs_nos = [it["no"] for it in rs_items]
    by_number = (all(n is not None for n in op_nos) and len(set(op_nos)) == len(op_nos)
                 and all(n is not None for n in rs_nos) and len(set(rs_nos)) == len(rs_nos))
    if by_number:
        rs_by_no = {it["no"]: it for it in rs_items}
        pairs = [(op, rs_by_no[op["no"]]) for op in op_items if op["no"] in rs_by_no]
        missing = [op for op in op_items if op["no"] not in rs_by_no]
        extra = [rs for rs in rs_items if rs["no"] not in {o["no"] for o in op_items}]
        return pairs, missing, extra, True
    n = min(len(op_items), len(rs_items))
    return list(zip(op_items[:n], rs_items[:n])), op_items[n:], rs_items[n:], False


# ---------------- 主核对 ----------------
def check(reviews_path, response_path):
    """执行核对，返回 {"hard": [...], "todo": [...], "ok": [...], "info": {...}}。"""
    hard, todo, ok = [], [], []
    info = {}
    r_lines, r_heads = read_paragraphs(reviews_path)
    s_lines, s_heads = read_paragraphs(response_path)
    op_items, op_pre, op_mode = split_items(r_lines, r_heads, "opinions")
    rs_items, rs_pre, rs_mode = split_items(s_lines, s_heads, "response")
    info["n_opinions"] = len(op_items)
    info["n_blocks"] = len(rs_items)
    info["op_mode"] = op_mode
    info["rs_mode"] = rs_mode
    info["reviewers"] = sum(1 for t in r_lines if REVIEWER_HEAD_RE.match(t))

    # 结构识别失败：只报提示，不做逐条硬判（避免把「解析不出」当「漏回」）
    if not op_items or not rs_items:
        todo.append({
            "code": "structure", "title": "未识别到逐条结构，无法逐条核对",
            "detail": (f"意见条目 {len(op_items)} 个、回应区块 {len(rs_items)} 个"
                       f"（意见切分={op_mode}，回应切分={rs_mode}）"),
            "hint": "请按「意见 N：／问题 N」编号或表格行组织两份文件后重跑",
        })
        if not rs_items:
            whole = " ".join(s_lines)
            if not re.search(r"回应|回复|答复|答\s*[：:]", whole):
                hard.append({
                    "code": "no_response_field", "title": "回复信未检出任何回应字段",
                    "detail": "全文无「回应/回复/答复」标签，也无回应线索词",
                    "hint": "按 response-template.md 补「回应：」字段",
                })
            level, ev = precision_of(whole)
            if level == "missing":
                hard.append({
                    "code": "no_location_global", "title": "回复信全文无位置标注",
                    "detail": "未检出「第 X 节/段/页/表/图」类定位",
                    "hint": "逐条补「修改位置：第 X 节第 X 段」",
                })
        return {"hard": hard, "todo": todo, "ok": ok, "info": info}

    # ① 条数对齐（按编号集合）
    pairs, missing, extra, by_number = align_items(op_items, rs_items)
    info["aligned_by_number"] = by_number
    if missing:
        nos = "、".join(str(it["no"]) if it["no"] is not None else f"段{it['par']}"
                        for it in missing)
        hard.append({
            "code": "missing_response", "title": f"漏回 {len(missing)} 条意见",
            "detail": f"意见 {nos} 在回复信中无对应区块（意见 {info['n_opinions']} 条 "
                      f"vs 回应区块 {info['n_blocks']} 个）",
            "hint": "每条意见都必须回应，哪怕只是「已按要求补做」",
        })
    if len(op_items) > len(rs_items) and not missing:
        # 无编号或编号对不齐时的兜底：按条数差额报漏回
        hard.append({
            "code": "count_mismatch", "title": "意见条数多于回应区块数",
            "detail": f"意见 {info['n_opinions']} 条 vs 回应区块 {info['n_blocks']} 个"
                      + ("" if by_number else "（编号未对齐，按序号顺序配对）"),
            "hint": "核对是否有意见被合并/跳过；每条意见一个区块",
        })
    if extra:
        nos = "、".join(str(it["no"]) if it["no"] is not None else f"段{it['par']}"
                        for it in extra)
        todo.append({
            "code": "extra_block", "title": f"回复信多出 {len(extra)} 个区块",
            "detail": f"区块 {nos} 在审稿意见中无对应条目",
            "hint": "确认是编号错位还是合并/拆分了意见",
        })
    if op_mode == "para":
        todo.append({
            "code": "op_mode", "title": "意见文件无编号，已按段落切分",
            "detail": f"切成 {len(op_items)} 条，可能与审稿人原意不符",
            "hint": "人工确认意见条数后再看本报告",
        })
    if op_mode == "generic":
        todo.append({
            "code": "op_mode", "title": "意见文件按通用编号（1. / （1））切分",
            "detail": f"切成 {len(op_items)} 条",
            "hint": "人工确认切分是否符合审稿人原意",
        })

    # ② 三件套 + ③ 位置精确度 + 主题词重叠
    uses_resp_label = any(leading_label(t)[0] == "response" for it in rs_items for t in it["lines"])
    uses_loc_label = any(leading_label(t)[0] == "location" for it in rs_items for t in it["lines"])
    for op, rs in pairs:
        label = f"意见 {op['no']}" if op["no"] is not None else f"意见（段{op['par']}）"
        f = parse_block(rs["lines"])
        issues = []
        # 回应
        if not f["response"]:
            if uses_resp_label:
                issues.append(("hard", "缺「回应」字段"))
            else:
                issues.append(("todo", "未识别到「回应」字段（疑似混排，需人工确认）"))
        # 意见复述
        if not f["opinion"]:
            if uses_resp_label or uses_loc_label:
                issues.append(("todo", "缺「意见（复述）」字段"))
        # 修改位置
        if not f["location"]:
            if uses_loc_label:
                issues.append(("hard", "缺「修改位置」字段"))
            else:
                level, ev = precision_of(f["response"] or " ".join(rs["lines"]))
                if level == "missing":
                    issues.append(("hard", "无位置标注（全文找不到节/段/页定位）"))
                elif level == "vague":
                    issues.append(("todo", f"位置标注不精确（{ev}）"))
        else:
            level, ev = precision_of(f["location"])
            if level == "precise":
                pass
            elif level == "vague":
                issues.append(("todo", f"位置标注不精确（{ev}）"))
            else:
                issues.append(("hard", "「修改位置」字段为空"))
        # 主题词重叠：疑似答非所问
        op_text = " ".join(op["lines"])
        ov = topic_overlap(op_text, f["opinion"] or " ".join(rs["lines"]))
        if ov < OVERLAP_MIN_HIT and f["response"]:
            issues.append(("todo", f"回应与意见主题词重叠仅 {ov}（疑似答非所问或未复述意见）"))
        if not issues:
            ok.append({"label": label, "par": rs["par"]})
        else:
            for lv, msg in issues:
                rec = {"label": label, "par": rs["par"], "msg": msg,
                       "detail": trunc(f["opinion"] or " ".join(rs["lines"]), 60)}
                (hard if lv == "hard" else todo).append(rec)

    # 模板结构提示（非阻断）
    if len(rs_items) >= 2:
        pre_text = " ".join(rs_pre)
        if not re.search(r"感谢|致谢|谢", pre_text):
            todo.append({
                "code": "no_prelude", "title": "未检出开场总述段",
                "detail": "模板第 1 项：感谢编辑与审稿专家 + 概述本轮修改方向",
                "hint": "补 2–3 句开场总述；感谢语克制，不堆「非常宝贵」",
            })
    return {"hard": hard, "todo": todo, "ok": ok, "info": info}


# ---------------- 输出 ----------------
def main():
    args = sys.argv[1:]
    flags = set(a for a in args if a.startswith("--"))
    if not args or "-h" in args or "--help" in args:
        print(__doc__)
        sys.exit(0)

    def _opt(name):
        if name in args:
            i = args.index(name)
            return args[i + 1] if i + 1 < len(args) else None
        return None

    reviews = _opt("--reviews")
    response = _opt("--response")
    if not reviews or not response:
        print("用法: python response_check.py --reviews <审稿意见文件> --response <回复信文件> "
              "[--top 20] [--json]")
        print("     支持 .docx（段落+表格）/ .md / .txt")
        sys.exit(2)
    for p in (reviews, response):
        if not os.path.isfile(p):
            print(f"[错误] 找不到文件：{p}")
            sys.exit(2)
    TOP = 20
    if "--top" in args:
        i = args.index("--top")
        try:
            TOP = int(args[i + 1])
        except (IndexError, ValueError):
            pass

    try:
        res = check(reviews, response)
    except ValueError as e:
        print(f"[错误] {e}")
        sys.exit(2)

    info = res["info"]
    print("=" * 70)
    print("审稿回复逐条回应完整性核对（模式 F 返修阶段）")
    print(f"审稿意见：{os.path.basename(reviews)}（{info['n_opinions']} 条）")
    print(f"回复信：{os.path.basename(response)}（{info['n_blocks']} 个回应区块）")
    if info.get("reviewers"):
        print(f"检测到 {info['reviewers']} 位审稿人分组头")
    print("=" * 70)

    print(f"\n【核对结果】🔴 硬伤 {len(res['hard'])} 项 ｜ 🟡 待人工核对 {len(res['todo'])} 项 ｜ "
          f"✅ 完整 {len(res['ok'])} 条")
    print("-" * 70)
    if res["hard"]:
        print(f"🔴 审稿回复硬伤 {len(res['hard'])} 项（返修提交前必须处理）：")
        for r in res["hard"][:TOP]:
            _show(r, hard=True)
        if len(res["hard"]) > TOP:
            print(f"   …另 {len(res['hard']) - TOP} 项，用 --top N 调整显示条数")
    else:
        print("🔴 审稿回复硬伤 0 项")

    if res["todo"]:
        print(f"\n🟡 回复待人工核对 {len(res['todo'])} 项（不阻断，建议处理）：")
        for r in res["todo"][:TOP]:
            _show(r, hard=False)
        if len(res["todo"]) > TOP:
            print(f"   …另 {len(res['todo']) - TOP} 项，用 --top N 调整显示条数")

    if not res["hard"] and not res["todo"]:
        print("✅ 意见逐条有回应，三件套齐全，位置标注到节/段。")

    if "--json" in flags:
        jf = os.path.splitext(response)[0] + ".response_check.json"
        with open(jf, "w", encoding="utf-8") as f:
            json.dump({"reviews": os.path.basename(reviews),
                       "response": os.path.basename(response),
                       "info": info,
                       "hard": res["hard"], "todo": res["todo"],
                       "ok_n": len(res["ok"])},
                      f, ensure_ascii=False, indent=1)
        print(f"\nJSON 已写入：{jf}")

    sys.exit(1 if res["hard"] else 0)


def _show(r, hard=True):
    mark = "🔴" if hard else "🟡"
    head = f"   · {mark} [{r.get('label') or r.get('code') or '结构'}"
    if r.get("par"):
        head += f"·块{r['par']}"
    head += f"] {r.get('title') or r.get('msg')}"
    print(head)
    if r.get("detail"):
        print(f"       {r['detail']}")
    if r.get("hint"):
        print(f"       ↳ {r['hint']}")


if __name__ == "__main__":
    main()
