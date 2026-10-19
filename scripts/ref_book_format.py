# -*- coding: utf-8 -*-
"""ref_book_format.py —— 专著与报纸文献格式核验（官方 §7 著作 / 报纸格式串）

用途：
    专门核 [M] 专著与 [N] 报纸两类条目是否符合官方投稿须知的格式串。
    「出版地：出版者」缺失是专著条目最常见的问题（作者直接写
    「出版社名, 年」或只写「上海, 2019」），报纸条目则常漏「：版面」
    或日期月日不补 0。

依据（references/official-guidelines.md §7，本 skill 唯一官方口径来源）：
    | 著作     | 作者．书名[M]．出版地：出版者，出版年：参考页码． |
    | 报纸文章 | 作者．题名[N]．报名,年-月-日(月日为单数前面加0)：版面． |
  由该格式串可机械派生的规则：
    B1 专著必须有**出版地**：格式串为「出版地：出版者」，即类型标识后
       紧跟一个「城市名：」结构；缺「北京：」式出版地即报；
    B2 专著必须有**出版者**：「：」之后为出版者（出版社/公司名），
       缺失即报；
    B3 专著必须有**出版年**：出版者后的 4 位年份；
    B4 专著应有**参考页码**（格式串「：参考页码」）——缺失作提示级
       （官方写了但未标注必填强度，故不判硬伤）；
    B5 报纸条目**月日补 0**（官方括注明文「月日为单数前面加0」）；
    B6 报纸条目应有**版面**（格式串「：版面．」收尾）；
    B7 专著若带网址/引用日期，应标 [M/OL]（官方电子文献格式串）。

阈值/规则出处：B1–B7 全部来自 official-guidelines.md §7 著作与报纸文章
两行格式串的文字，无自造阈值。

用法:
    python ref_book_format.py <论文.docx> [--json]

退出码:
    0 = M/N 条目均符合官方格式（含「本文无专著与报纸条目」）
    1 = 发现不合规项（缺出版地/出版者/出版年、报纸缺日期或版面）
    2 = 文件不存在 / 未找到参考文献段 / 未解析到条目
"""
import os
import re
import sys
import json

# Windows 控制台 GBK 保护
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import diagnose_paper as dp  # noqa: E402  （复用 extract_paras）
import ref_format as rf      # noqa: E402  （复用 parse_item / TYPE_PAT）

CITY = (r"北京|上海|天津|重庆|长春|吉林|沈阳|大连|哈尔滨|南京|苏州|杭州|宁波|"
        r"合肥|福州|厦门|南昌|济南|青岛|郑州|武汉|长沙|广州|深圳|珠海|南宁|"
        r"海口|成都|贵阳|昆明|西安|兰州|西宁|银川|乌鲁木齐|拉萨|石家庄|太原|"
        r"呼和浩特|香港|澳门|台北|伦敦|纽约|波士顿|芝加哥|柏林|巴黎|东京|"
        r"新加坡|New York|London|Boston|Chicago|Berlin|Paris|Tokyo|Singapore")

# 出版地：出版者（官方格式串「出版地：出版者」）
PLACE_PUB = re.compile(CITY + r"\s*[：:]\s*[\u4e00-\u9fffA-Za-z&·]{2,40}")
# 出版者信号（无出版地时仍可能是「出版社名」写法）
PUB_SIGN = re.compile(r"出版社|出版公司|书局|印书馆|大学出版社|Press|"
                      r"Publishing|Publisher|ACM|IEEE", re.I)
YEAR_RE = re.compile(r"(?<!\d)((?:19|20)\d{2})(?!\d)")
URL_RE = re.compile(r"https?://|www\.", re.I)
# 报纸日期与版面
N_DATE = re.compile(r"[,，]\s*(\d{4})\s*-\s*(\d{1,2})\s*-\s*(\d{1,2})")
N_PAGE = re.compile(r"[：:]\s*(?:第?\s*)?[A-Z0-9]{1,6}\s*版?"
                    r"(?:[-\u2013\u2014]\s*[A-Z0-9]{1,6}\s*版?)?\s*[。．.]?\s*$")


def ref_num(text):
    m = re.match(r"^[\[［]?\s*(\d{1,3})", text or "")
    return int(m.group(1)) if m else None


def check_monograph(raw):
    """核 [M] 专著条目。返回 [(级别, 规则号, 说明)]。"""
    f = []
    tm = rf.TYPE_PAT.search(raw)
    tag = tm.group(1).upper() if tm else ""
    after = raw[tm.end():] if tm else raw
    # 去掉网址/DOI/方括号访问日期后再核出版信息
    core = re.sub(r"\[\s*\d{4}-\d{2}-\d{2}\s*\]", " ", after)
    core = URL_RE.sub(" ", core)

    # B7：带网址的专著应标 /OL
    if URL_RE.search(after) and "/OL" not in tag:
        f.append(("red", "B7",
                  f"专著条目含网址但类型标识为 [{tag}]——按官方电子文献格式"
                  f"串应标 [M/OL]"))

    pm = PLACE_PUB.search(core)
    if pm:
        pass  # B1+B2 同时满足
    else:
        # B1 出版地
        if not re.search(CITY, core):
            f.append(("red", "B1",
                      "缺出版地——官方格式串为「作者．书名[M]．出版地：出版者，"
                      "出版年：参考页码．」，缺「北京：」式出版地最常见"))
        # B2 出版者
        if not PUB_SIGN.search(core):
            f.append(("red", "B2",
                      "缺出版者——「出版地：」之后须为出版社/出版公司名"))

    # B3 出版年
    if not YEAR_RE.search(core):
        f.append(("red", "B3",
                  "缺出版年——官方格式串「出版者，出版年」"))
    else:
        # B4 参考页码（官方写了「：参考页码」；缺失作提示）
        if not re.search(r"\d{4}\s*[,，]\s*[\u4e00-\u9fffA-Za-z&·]{2,40}\s*[,，]?\s*"
                         r"(?:19|20)\d{2}\s*[：:]\s*\d", core):
            f.append(("notice", "B4",
                      "专著条目未见「出版年：参考页码」形态——官方著作格式串"
                      "以「：参考页码．」收尾，建议补齐"))
    return f


def check_newspaper(raw):
    """核 [N] 报纸条目。返回 [(级别, 规则号, 说明)]。"""
    f = []
    m = N_DATE.search(raw)
    if m:
        y, mm, dd = m.group(1), m.group(2), m.group(3)
        # B5 月日补 0
        if len(mm) != 2 or len(dd) != 2:
            f.append(("red", "B5",
                      f"报纸日期月日未补 0：{y}-{mm}-{dd}——官方括注"
                      f"「月日为单数前面加0」，应写 {y}-{mm.zfill(2)}"
                      f"-{dd.zfill(2)}"))
    else:
        f.append(("red", "B5",
                  "N 类条目未见「年-月-日」——官方报纸文章格式串为"
                  "「作者．题名[N]．报名,年-月-日(月日为单数前面加0)：版面．」"))
    # B6 版面
    if not N_PAGE.search(raw.rstrip()):
        f.append(("red", "B6",
                  "缺版面号——官方报纸文章格式串以「：版面．」收尾"
                  "（如「：05」或「：A05」）"))
    return f


def analyze(doc):
    paras, refs, abs_text, kw_text, titles, body = dp.extract_paras(doc)
    findings = []
    n_m = n_n = 0
    for raw in refs:
        tm = rf.TYPE_PAT.search(raw)
        tag = tm.group(1).upper() if tm else ""
        base = tag.split("/")[0]
        if base == "M" and "//" not in raw:   # 论文集析出 [M]// 不算专著
            n_m += 1
            for lvl, rule, msg in check_monograph(raw):
                findings.append({"num": ref_num(raw), "level": lvl,
                                 "rule": rule, "msg": msg, "raw": raw[:70]})
        elif tag == "N":
            n_n += 1
            for lvl, rule, msg in check_newspaper(raw):
                findings.append({"num": ref_num(raw), "level": lvl,
                                 "rule": rule, "msg": msg, "raw": raw[:70]})

    reds = [x for x in findings if x["level"] == "red"]
    notices = [x for x in findings if x["level"] == "notice"]

    def nums(items):
        return "".join(f"[{x['num']}]" for x in items[:24])

    by_rule = {}
    for x in findings:
        by_rule.setdefault(x["rule"], []).append(x)

    return {"n_refs": len(refs), "n_m": n_m, "n_n": n_n,
            "reds": reds, "notices": notices, "by_rule": by_rule,
            "nums": nums, "findings": findings}


def main():
    args = sys.argv[1:]
    if not args or args[0] in ("-h", "--help"):
        print(__doc__)
        return 0
    _rest = [a for a in args if not a.startswith("--")]
    _flags = {a for a in args if a.startswith("--")}
    if not _rest:
        print(__doc__)
        return 2
    DOC = _rest[0]
    if not os.path.isfile(DOC):
        print(f"找不到文件：{DOC}")
        return 2

    try:
        res = analyze(DOC)
    except Exception as e:
        print(f"解析失败（未找到参考文献段或条目？）：{e}")
        return 2
    if res["n_refs"] == 0:
        print("未解析到文献条目")
        return 2

    print("=" * 74)
    print("专著与报纸文献格式核验（官方 §7 著作 / 报纸文章格式串）")
    print(f"稿件：{os.path.basename(DOC)}")
    print(f"文献 {res['n_refs']} 条：专著 [M] {res['n_m']} 条"
          f"／报纸 [N] {res['n_n']} 条")
    print("=" * 74)
    print("【官方格式基准】references/official-guidelines.md §7")
    print("  著作    ：作者．书名[M]．出版地：出版者，出版年：参考页码．")
    print("  报纸文章：作者．题名[N]．报名,年-月-日"
          "(月日为单数前面加0)：版面．")
    print("  B1 出版地（「北京：」式）｜B2 出版者｜B3 出版年｜B4 参考页码｜"
          "B5 报纸月日补 0｜B6 版面号｜B7 带网址专著标 [M/OL]")

    if res["n_m"] == 0 and res["n_n"] == 0:
        print("\n（本文无专著与报纸条目，官方格式串不适用）")
        return 0

    print(f"\n【汇总】不合规 {len(res['reds'])} 处（硬伤）／"
          f"提示 {len(res['notices'])} 处")
    labels = {
        "B1": "专著缺出版地（「北京：」式）",
        "B2": "专著缺出版者",
        "B3": "专著缺出版年",
        "B4": "专著未见「出版年：参考页码」（提示）",
        "B5": "报纸日期月日未补 0 / 缺日期段",
        "B6": "报纸缺版面号",
        "B7": "专著带网址却未标 [M/OL]",
    }
    for rule, items in sorted(res["by_rule"].items(),
                              key=lambda kv: -len(kv[1])):
        lvl = "🔴" if any(i["level"] == "red" for i in items) else "🟡"
        print(f"  {lvl} {labels.get(rule, rule)}：{len(items)} 处  "
              f"{res['nums'](items)}")

    if res["reds"]:
        print("\n【明细】逐条定位（前 10 条）")
        seen = set()
        shown = 0
        for x in res["reds"]:
            if x["num"] in seen:
                continue
            seen.add(x["num"])
            shown += 1
            if shown > 10:
                break
            print(f"  [{x['num']}] {x['raw'][:58]}")
            for y in [z for z in res["reds"] if z["num"] == x["num"]][:3]:
                print(f"       · {y['msg']}")
        print("\n【修法】专著改为「作者．书名[M]．出版地：出版者，出版年："
              "参考页码．」；报纸改为「作者．题名[N]．报名,2023-06-05："
              "版面．」")
    else:
        print("\n✅ 专著与报纸条目均符合官方 §7 格式。")

    if res["notices"]:
        print("\n【提示】")
        for x in res["notices"][:8]:
            print(f"  · [{x['num']}] {x['msg']}")

    print("\n【边界】只核格式要素；[M]// 论文集析出条目按「论文集析出」格式串"
          "另行核对，不计入专著。文献真实性需人工核验。")

    if "--json" in _flags:
        jf = os.path.splitext(DOC)[0] + ".ref_book.json"
        with open(jf, "w", encoding="utf-8") as fh:
            json.dump({k: v for k, v in res.items() if k != "nums"},
                      fh, ensure_ascii=False, indent=1)
        print(f"\nJSON 已写入：{jf}")

    return 1 if res["reds"] else 0


if __name__ == "__main__":
    sys.exit(main())
