# -*- coding: utf-8 -*-
"""benchmark_paper.py —— 论文对标评估：把一篇论文放进《图书馆学研究》语料库分布中定位
对比维度：段长 / 句长 / 引用密度 / 摘要 / 关键词 / 文献量 / 图表 / 反模式 / 结构完备度
另附：同主题参照组（扩展区 10 篇 + 主库相关 8 篇）三项对比

用法: python benchmark_paper.py <论文.docx>
说明: 需与语料库（corpus_stats.json / paragraphs.json）同目录。
     本工具评估的是【形式与结构完成度】，不是学术质量排名。
"""
import re, sys, os, json, statistics as st

if len(sys.argv) < 2:
    sys.exit("用法: python benchmark_paper.py <论文.docx>")
D = os.path.dirname(os.path.abspath(__file__))
P = sys.argv[1]

def clean(s): return re.sub(r"\s+", "", s)

# ---------- 论文端提取 ----------
from docx import Document
doc = Document(P)
paras = [(p.style.name if p.style else "", p.text.strip()) for p in doc.paragraphs]
abs_t = kw_t = ""
refs, body, titles = [], [], []
in_ref = False
for style, t in paras:
    if not t: continue
    if t.startswith("参考文献"): in_ref = True; continue
    if in_ref:
        if re.match(r"^\[?\d{1,3}\]", t) or re.match(r"^[\[［]\d+[\]］]", t): refs.append(t)
        continue
    if re.search(r"摘\s*要", t[:10]) and not abs_t: abs_t = t; continue
    if re.search(r"关\s*键\s*词", t[:10]) and not kw_t: kw_t = t; continue
    if re.match(r"^(0|1|2|3|4|5|6|7)\s*[^\d]", t) and len(t) < 40 or style.startswith("Heading") and len(t) < 40:
        titles.append(t); continue
    if len(t) < 45 and not re.search(r"[。；，]", t) and len(body) > 0: continue
    if len(t) > 45 or (not body):
        if "摘" in t[:6] or "Abstract" in t[:12]: continue
    body.append(t)

bk = [p for p in body if len(clean(p)) >= 30]
lens = [len(clean(p)) for p in bk]
sents = []
for p in bk: sents += [s for s in re.split(r"[。？！]", clean(p)) if len(s) >= 6]
alltxt = clean("".join(bk))
cites = len(re.findall(r"\[\d+([-,，]\s*\d+)*\]", alltxt))
benwen = alltxt.count("本文")
abs_len = len(clean(re.sub(r"^【?摘\s*要】?[：:]?", "", abs_t)))
kw_n = len([k for k in re.split(r"[\s;；,，、]+", re.sub(r"^【?关\s*键\s*词】?[：:]?", "", kw_t)) if k])
ref_n = len(refs)
figs = [int(m) for m in re.findall(r"图\s*(\d{1,2})", alltxt)]
tabs = [int(m) for m in re.findall(r"表\s*(\d{1,2})", alltxt)]
fig_n = max([x for x in figs if x <= 20], default=0)
tab_n = max([x for x in tabs if x <= 20], default=0)
apa = {}
for name, pat in [("三连×化", r"[\u4e00-\u9fff]{1,4}化[、,，][\u4e00-\u9fff]{1,4}化[、,，][\u4e00-\u9fff]{1,4}化"),
                  ("万能句", r"具有重要意义|值得深入探讨|提供有益参考|意义重大"),
                  ("机械序数", r"(首先|其次|再次)[，,]"),
                  ("不仅…而且", r"不仅[^。]{2,50}(而且|更)")]:
    apa[name] = len(re.findall(pat, alltxt))
apa_total = sum(apa.values())

intro_chars = 0; cur = ""
for _, t in paras:
    if re.match(r"^0\s", t.strip()) and len(t) < 40: cur = "intro"; continue
    if re.match(r"^1\s", t.strip()) and len(t) < 40: cur = ""; continue
    if cur == "intro": intro_chars += len(clean(t))
total_body = len(alltxt)
intro_pct = intro_chars / max(total_body, 1) * 100
h1_hit = sum(1 for pat in [r"^0?\s*引\s*言", r"^1\s", r"^[2-4]\s", r"^4\.6|^5\s.*策略|^5\s.*对策|^4\s.*策略", r"^[45]\s*结\s*语"]
             for t in titles if re.search(pat, t))

# ---------- 语料库端 ----------
base = json.load(open(os.path.join(D, "corpus_stats.json"), encoding="utf-8"))
ok = [r for r in base if r.get("has_indent", True) and r.get("max_para", 0) <= 1500]

def pct(v, arr):
    arr = sorted(x for x in arr if x is not None)
    if not arr: return 50
    return sum(1 for x in arr if x < v) / len(arr) * 100

def line(metric, v, arr, kind="pos", unit="", fmt="{:.0f}"):
    p = pct(v, arr)
    if kind == "pos":
        tag = "▲ 强" if p >= 75 else ("▼ 弱" if p <= 25 else "≈ 中")
    elif kind == "neutral":
        q25, q75 = sorted(arr)[len(arr)//4], sorted(arr)[len(arr)*3//4]
        tag = "偏短" if v < q25 else ("偏长" if v > q75 else "居中")
        p = None
    else:
        tag = "—"
    pstr = f"{p:.0f}%" if p is not None else "—"
    med = st.median(arr)
    print(f"  {metric:<12s} {fmt.format(v):>7s}{unit}   刊中位 {med:.1f} | 百分位 {pstr:>5s}  {tag}")

print("=" * 64)
print(f"论文对标评估：{os.path.basename(P)[:44]}")
print(f"基准：《图书馆学研究》108 篇（2026 年 1—9 期）")
print("=" * 64)
print("\n【形式维度对标】（百分位 = 高于该%的刊文）")
line("段长中位", st.median(lens) if lens else 0, [r["para_med"] for r in ok], "neutral")
line("句均字数", st.mean([len(s) for s in sents]) if sents else 0, [r["sent_mean"] for r in ok if r["sent_mean"] > 0], "neutral")
line("引用密度", cites / max(len(alltxt), 1) * 1000, [r["cite_rate"] for r in ok], "pos", "/千字", "{:.2f}")
line("摘要长度", abs_len, [r["abstract_len"] for r in ok if 50 < r["abstract_len"] < 800], "neutral")
line("关键词数", kw_n, [r["kw_n"] for r in ok if r["kw_n"] > 0], "pos", "个")
line("文献条数", ref_n, [r["ref_cnt"] for r in ok if r["ref_cnt"] > 0], "pos", "条")
line("图数", fig_n, [r["fig_n"] for r in ok], "pos", "张")
line("表数", tab_n, [r["tab_n"] for r in ok], "pos", "张")
line("正文字数", total_body, [r["body_chars"] for r in ok], "neutral")

print(f"\n【净度】反模式命中：{apa_total} 处 {apa}")
print(f"  （刊文水平：每万字体量下三项合计约 1.5 处；0 处 = 优于全部刊文）")

print(f"\n【结构】章节标题 {len(titles)} 个 | 引言占正文 {intro_pct:.1f}%（刊单列引言章中位 7.1%）")
print(f"  一级标题序列：{' › '.join(t[:14] for t in titles if re.match(r'^[0-7]\s', t))[:90]}")

# ---------- 同主题参照组 ----------
print("\n【同主题参照组】（无障碍/老年/特殊群体方向 18 篇）")
data = json.load(open(os.path.join(D, "paragraphs.json"), encoding="utf-8"))
topic_files = ["011", "039", "047", "051", "057", "085", "086", "092"]
topic = [x for x in data if x["no"] in topic_files or (x.get("no") and int(x["no"]) >= 109)]
tl, ts = [], []
for x in topic:
    tp = [p for p in x["paras"] if len(p) >= 100]
    tl += [len(p) for p in tp]
    for p in tp[:40]:
        ts += [len(s) for s in re.split(r"[。？！]", p) if len(s) >= 6]
if tl:
    print(f"  段长中位：你 {st.median(lens):.0f} vs 主题组 {st.median(tl):.0f} 字")
    print(f"  句均字数：你 {st.mean([len(s) for s in sents]):.0f} vs 主题组 {st.mean(ts):.0f} 字")

print("\n【边界声明】本评估为形式与结构完成度的统计对标，非学术质量排名；")
print("创新性、研究价值以编辑部与审稿专家判断为准。")
