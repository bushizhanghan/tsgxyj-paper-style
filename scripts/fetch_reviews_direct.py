# -*- coding: utf-8 -*-
"""fetch_reviews_direct.py -- 直连抓取该刊投稿者点评并结构化（扩充样本用）

本工具用于**维护数据源**（不是日常审稿工具）。抓取结果供
references/submission-reality.md 更新使用。

用法:
    python fetch_reviews_direct.py                 # 抓 1–6 页（默认）
    python fetch_reviews_direct.py --pages 1 12    # 抓全部有效页（该源共 12 页）
    python fetch_reviews_direct.py --analyze-only  # 只分析已抓的缓存

设计说明：
    · 用标准库 urllib 直连（web_fetch 工具在本机曾因出口问题失败，本脚本不受此影响）
    · 抓取后缓存到系统临时目录，可反复分析而不重复请求
    · 解析结果含审稿时间 / 录用结果 / 稿件字数 / 查重口径 / 点评原文
    · 抓取带礼貌间隔，避免给站点造成压力
"""
import os
import re
import sys
import json
import time
import glob
import urllib.request
import urllib.error
from collections import Counter, defaultdict
import statistics as st

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

BASE = "https://eshukan.com/sci/scijdianping.aspx?pg={pg}&jid=9814"
CACHE = os.path.join(os.environ.get("TEMP", "/tmp"), "dianping")
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/120.0 Safari/537.36")


def strip_tags(s):
    s = re.sub(r"<br\s*/?>", "\n", s)
    s = re.sub(r"<[^>]+>", "", s)
    s = (s.replace("&nbsp;", " ").replace("&amp;", "&")
         .replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"'))
    return re.sub(r"[ \t]+", " ", s).strip()


def fetch_page(pg):
    url = BASE.format(pg=pg)
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=25) as r:
        return r.read().decode("utf-8", errors="replace")


def parse_html(html):
    """从一页 HTML 解析出点评记录列表。"""
    out = []
    for b in re.split(r'<div class="onedpinfo"', html)[1:]:
        rec = {}
        m = re.search(r"在\s*(\d{4}/\d{1,2}/\d{1,2})\s*[\d:]+\s*进行点评", b)
        if m:
            rec["date"] = m.group(1)
        for k, v in re.findall(r"<li>\s*([^：<]{2,8})：\s*<span>([^<]*)</span>", b):
            rec[k.strip()] = v.strip()
        m = re.search(r'<span>我的点评：</span>(.*?)(?:<div class="huiyinginfo"|'
                      r'<div class="onedpinfo"|\Z)', b, re.S)
        if m:
            t = strip_tags(m.group(1))
            if len(t) > 8:
                rec["comment"] = t[:800]
        m = re.search(r"投稿主题：</span>([^<]{2,40})", b)
        if m:
            rec["topic"] = m.group(1).strip()
        replies = [strip_tags(x)[:220] for x in re.findall(r"<p>(.*?)</p>", b, re.S)
                   if 8 < len(strip_tags(x)) < 320]
        if replies:
            rec["replies"] = replies[:4]
        if rec.get("comment") or rec.get("审稿时间"):
            out.append(rec)
    return out


def parse_days(t):
    if not t or t == "-":
        return None
    m = re.search(r"(\d+)\s*日", t)
    if m:
        return int(m.group(1))
    m = re.search(r"(\d+)\s*个?月", t)
    if m:
        return int(m.group(1)) * 30
    return None


def analyze(all_recs):
    print("=" * 70)
    print(f"《图书馆学研究》投稿者反馈分析（样本 {len(all_recs)} 条）")
    print("=" * 70)

    # 时间
    timed = [(parse_days(r.get("审稿时间", "")), r.get("是否录用", ""))
             for r in all_recs]
    timed = [(d, res) for d, res in timed if d]
    if timed:
        ns = [d for d, _ in timed]
        print(f"\n【审稿周期】可量化 {len(ns)} 条：中位≈{st.median(ns):.0f} 天  "
              f"均值≈{st.mean(ns):.0f} 天  范围 {min(ns)}–{max(ns)} 天")
        dist = Counter()
        for d in ns:
            if d <= 10:
                dist["10日内"] += 1
            elif d <= 15:
                dist["15日内"] += 1
            elif d <= 30:
                dist["1个月"] += 1
            elif d <= 60:
                dist["2个月"] += 1
            else:
                dist["3个月+"] += 1
        for k in ["10日内", "15日内", "1个月", "2个月", "3个月+"]:
            if dist[k]:
                print(f"  {k}: {dist[k]} 条")

        print("\n【结果 × 周期】（关键：退稿集中在短周期）")
        by_res = defaultdict(list)
        for d, res in timed:
            if res:
                by_res[res].append(d)
        for res, lst in sorted(by_res.items(), key=lambda x: -len(x[1])):
            print(f"  {res:<8} n={len(lst):<2} 中位≈{st.median(lst):.0f} 天  "
                  f"范围 {min(lst)}–{max(lst)} 天")

    # 结果分布
    print("\n【录用结果分布】")
    for k, v in Counter(r.get("是否录用", "") for r in all_recs if r.get("是否录用")).most_common():
        print(f"  {k}: {v} 条")

    # 字数
    print("\n【稿件字数】")
    wc = Counter(r.get("稿件字数", "") for r in all_recs if r.get("稿件字数"))
    for k, v in wc.most_common(8):
        print(f"  {k}: {v} 条")

    # 关注点
    KEY = {
        "审稿慢/进度不透明": r"进度不透明|没有消息|没消息|等.*久|长很多|漏审",
        "退稿": r"退稿|毙了|拒稿",
        "查重": r"查重|重复率|复制比",
        "格式/标点严": r"标点|格式.*严|校稿.*严|看得很仔细",
        "修改1–2轮": r"修改了?[一两]次|两轮|第二轮|返修|再审",
        "拟录用": r"拟录用",
        "不收版面费": r"不收版面费|不用.*版面费|没有让我交版面费",
        "见刊快": r"见刊.*快|排期.*快|非常迅速",
        "主题对位重要": r"喜好的主题|主题.*录用概率",
        "不建议催问": r"不建议.*投|不敢.*问|怕.*催",
    }
    print("\n【关注点频次】")
    for name, pat in KEY.items():
        n = sum(1 for r in all_recs if re.search(pat, json.dumps(r, ensure_ascii=False)))
        if n:
            print(f"  {name}: {n} 条")

    # 最有信息量的原文
    print("\n" + "=" * 70)
    print("高信息量原文（含流程/查重/建议）")
    print("=" * 70)
    shown = 0
    for r in all_recs:
        c = r.get("comment", "") or ""
        if re.search(r"查重|拟录用|初审|外审|喜好的主题|三审三校", c) and len(c) > 60:
            print(f"\n[{r.get('date', '?')}] 周期={r.get('审稿时间', '-')} "
                  f"结果={r.get('是否录用', '-')}")
            print(f"  {c[:400]}")
            shown += 1
            if shown >= 12:
                break


def main():
    args = sys.argv[1:]
    pages = range(1, 7)
    if "--pages" in args:
        i = args.index("--pages")
        try:
            a, b = int(args[i + 1]), int(args[i + 2])
            pages = range(a, b + 1)
        except (ValueError, IndexError):
            print("用法: --pages <起> <止>")
            return 2

    os.makedirs(CACHE, exist_ok=True)

    if "--analyze-only" not in args:
        for pg in pages:
            path = os.path.join(CACHE, f"pg{pg}.html")
            if os.path.isfile(path) and os.path.getsize(path) > 1000:
                print(f"pg{pg} 已缓存，跳过")
                continue
            try:
                html = fetch_page(pg)
                with open(path, "w", encoding="utf-8") as f:
                    f.write(html)
                print(f"pg{pg} 抓取成功（{len(html)} 字节）")
            except Exception as e:
                print(f"pg{pg} 失败：{e}")
            time.sleep(1.0)   # 礼貌间隔

    all_recs = []
    for fn in sorted(glob.glob(os.path.join(CACHE, "pg*.html"))):
        html = open(fn, encoding="utf-8", errors="ignore").read()
        all_recs.extend(parse_html(html))

    # 去重（同日期+同点评前 40 字视为同一条）
    uniq, seen = [], set()
    for r in all_recs:
        key = (r.get("date", ""), (r.get("comment", "") or "")[:40])
        if key in seen:
            continue
        seen.add(key)
        uniq.append(r)

    print(f"\n解析 {len(all_recs)} 条，去重后 {len(uniq)} 条\n")
    out = os.path.join(CACHE, "records.json")
    with open(out, "w", encoding="utf-8") as f:
        json.dump(uniq, f, ensure_ascii=False, indent=1)
    print(f"已保存：{out}\n")
    analyze(uniq)
    return 0


if __name__ == "__main__":
    sys.exit(main())
