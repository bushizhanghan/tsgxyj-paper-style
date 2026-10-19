# -*- coding: utf-8 -*-
"""ref_online_format.py —— 电子文献格式核验（官方 §7 电子文献格式串）

用途：
    专门核电子文献条目（[EB/OL] 及带 /OL 的 [J/OL][M/OL][D/OL][C/OL]
    [N/OL][S/OL][Z/OL] 等）是否符合官方投稿须知的电子文献格式，以及
    日期是否按官方规定补 0。

依据（references/official-guidelines.md §7，本 skill 唯一官方口径来源）：
    | 电子文献 | 作者．题名[电子文献类型/电子文献载体]．[年-月-日]
                (月日为单数前面加0)．网址． |
    | 报纸文章 | 作者．题名[N]．报名,年-月-日(月日为单数前面加0)：版面． |
  由该格式串可机械派生的规则：
    E1 引用日期必须带**方括号** `[年-月-日]`（官方格式串写「[年-月-日]」），
       仅有圆括号 `(年-月-日)` 不合规；方括号缺失或改用圆括号即报；
    E2 引用日期的**月、日须为两位数**（单数前加 0）：2023-06-05 合规，
       2023-6-5 不合规（官方括注明文「月日为单数前面加0」）；
    E3 必须含**网址**（格式串末段「网址」）——`http(s)://` 或 `www.`；
    E4 电子文献类型标识应为 `类型/OL` 形态（如 EB/OL、J/OL、Z/OL、S/OL）；
       仅 [EB] 或 [OL] 单独出现即报；
    E5 各段分隔点用全角「．」（与官方格式串一致；标点细节可另跑
       ref_format_style.py，此处只在电子条目上做复核）。

另核 N 类报纸条目的日期补 0 与「：版面」要素（官方 §7 报纸文章格式串
同样明文「月日为单数前面加0」）。

阈值/规则出处：E1–E5 全部来自 official-guidelines.md §7 电子文献与报纸
文章两行格式串的文字，无自造阈值。

用法:
    python ref_online_format.py <论文.docx> [--json]

退出码:
    0 = 电子文献条目均符合官方格式（含「本文无电子文献条目」）
    1 = 发现不合规项（缺方括号日期 / 月日未补 0 / 缺网址 / 标识形态错）
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
import ref_format as rf      # noqa: E402  （复用 parse_item）

# ---- 官方 §7 电子文献格式串派生规则 ----
URL_RE = re.compile(r"https?://\S+|www\.[\w.-]+", re.I)
# 方括号引用日期：[2023-06-05]
DATE_BRACKET = re.compile(r"\[\s*(\d{4})\s*-\s*(\d{1,2})\s*-\s*(\d{1,2})\s*\]")
# 圆括号日期（官方要求方括号）
DATE_PAREN = re.compile(r"\(\s*(\d{4})\s*-\s*(\d{1,2})\s*-\s*(\d{1,2})\s*\)")
# 电子文献类型标识（官方：电子文献类型/电子文献载体）
ONLINE_TYPE = re.compile(r"\[([A-Za-z]{1,3})/OL\]")
# 报纸类
NEWSPAPER_TYPE = re.compile(r"\[N\]")
# 报纸日期：报名,年-月-日：版面
N_DATE = re.compile(r"[,，]\s*(\d{4})\s*-\s*(\d{1,2})\s*-\s*(\d{1,2})")
N_PAGE = re.compile(r"[：:]\s*\d{1,3}(?:[-\u2013\u2014]\d{1,3})?\s*[。．.]?\s*$")


def ref_num(text):
    m = re.match(r"^[\[［]?\s*(\d{1,3})", text or "")
    return int(m.group(1)) if m else None


def date_ok(y, m, d):
    """月日补 0 判定：官方括注「月日为单数前面加0」。"""
    return len(m) == 2 and len(d) == 2


def check_online(raw):
    """核单条电子文献。返回 findings: [(级别, 规则号, 说明)]。"""
    f = []
    tm = rf.TYPE_PAT.search(raw)
    tag = tm.group(1).upper() if tm else ""

    # E4：电子文献标识形态
    if "/OL" not in tag:
        if re.fullmatch(r"EB|OL", tag, re.I):
            f.append(("red", "E4",
                      f"电子文献类型标识为 [{tm.group(1)}]——官方格式串为"
                      f"「[电子文献类型/电子文献载体]」，应为 [EB/OL]"))
    # 电子文献强信号但未标 /OL（含网址、方括号访问日期）
    if URL_RE.search(raw) and "/OL" not in tag and tag not in ("EB/OL",):
        if re.search(r"\[\s*\d{4}-\d{1,2}-\d{1,2}\s*\]", raw):
            f.append(("red", "E4",
                      f"条目含网址与引用日期，类型标识却是 [{tag or '无'}]"
                      f"——官方电子文献格式应为 [EB/OL]"))

    # E1：引用日期须方括号
    db = list(DATE_BRACKET.finditer(raw))
    if db:
        for m in db:
            # E2：月日补 0
            if not date_ok(m.group(1), m.group(2), m.group(3)):
                f.append(("red", "E2",
                          f"引用日期月日未补 0：[{m.group(1)}-{m.group(2)}"
                          f"-{m.group(3)}]——官方括注「月日为单数前面加0」，"
                          f"应写 [{m.group(1)}-{m.group(2).zfill(2)}"
                          f"-{m.group(3).zfill(2)}]"))
        # E2 反向：补 0 过度（月日出现 0 开头但 >12/31 之类）
        for m in db:
            mm, dd = m.group(2).zfill(2), m.group(3).zfill(2)
            if int(mm) > 12 or int(dd) > 31:
                f.append(("red", "E2",
                          f"引用日期越界：[{m.group(1)}-{m.group(2)}"
                          f"-{m.group(3)}]"))
    else:
        dp_ = list(DATE_PAREN.finditer(raw))
        if dp_:
            m = dp_[0]
            f.append(("red", "E1",
                      f"引用日期用圆括号 ({m.group(1)}-{m.group(2)}"
                      f"-{m.group(3)})——官方格式串为方括号"
                      f"「[年-月-日]」，应改为 [{m.group(1)}-{m.group(2).zfill(2)}"
                      f"-{m.group(3).zfill(2)}]"))
        else:
            f.append(("red", "E1",
                      "缺引用日期——官方电子文献格式串为"
                      "「…[年-月-日]．网址．」，无引用日期不合规"))

    # E3：网址
    if not URL_RE.search(raw):
        f.append(("red", "E3",
                  "缺网址——官方电子文献格式串末段为「网址」，"
                  "电子文献必须给出可访问地址"))

    # E5：电子条目分隔点全角（与官方格式串一致）
    for m in re.finditer(r"\]([.．])", raw):
        if m.group(1) == ".":
            f.append(("notice", "E5",
                      "类型标识后用半角「.」——官方电子文献格式串分隔点为"
                      "全角「．」（可另跑 ref_format_style.py 统一处理）"))
    return f


def check_newspaper(raw):
    """核 N 类报纸条目：日期补 0 + 版面要素。"""
    f = []
    m = N_DATE.search(raw)
    if m:
        if not date_ok(m.group(1), m.group(2), m.group(3)):
            f.append(("red", "E2",
                      f"报纸日期月日未补 0：{m.group(1)}-{m.group(2)}"
                      f"-{m.group(3)}——官方括注「月日为单数前面加0」，"
                      f"应写 {m.group(1)}-{m.group(2).zfill(2)}"
                      f"-{m.group(3).zfill(2)}"))
    else:
        f.append(("red", "E1",
                  "N 类条目未见「年-月-日」日期段——官方报纸文章格式串为"
                  "「作者．题名[N]．报名,年-月-日(月日为单数前面加0)：版面．」"))
    if not N_PAGE.search(raw.rstrip()):
        f.append(("red", "E3",
                  "N 类条目末尾未见「：版面」——官方报纸文章格式串以"
                  "「：版面．」收尾"))
    return f


def analyze(doc):
    paras, refs, abs_text, kw_text, titles, body = dp.extract_paras(doc)
    findings = []
    n_online = n_n = 0
    for raw in refs:
        tm = rf.TYPE_PAT.search(raw)
        tag = tm.group(1).upper() if tm else ""
        if "/OL" in tag or (URL_RE.search(raw) and
                            re.search(r"\[\s*\d{4}-\d{1,2}-\d{1,2}\s*\]", raw)):
            n_online += 1
            for lvl, rule, msg in check_online(raw):
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

    return {"n_refs": len(refs), "n_online": n_online, "n_n": n_n,
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
    print("电子文献格式核验（官方 §7 电子文献 / 报纸文章格式串）")
    print(f"稿件：{os.path.basename(DOC)}")
    print(f"文献 {res['n_refs']} 条：电子文献 {res['n_online']} 条"
          f"（EB/OL 及 /OL 各型）／报纸 [N] {res['n_n']} 条")
    print("=" * 74)
    print("【官方格式基准】references/official-guidelines.md §7")
    print("  电子文献：作者．题名[电子文献类型/电子文献载体]．"
          "[年-月-日](月日为单数前面加0)．网址．")
    print("  报纸文章：作者．题名[N]．报名,年-月-日"
          "(月日为单数前面加0)：版面．")
    print("  E1 引用日期方括号｜E2 月日补 0（2023-06-05 而非 2023-6-5）｜"
          "E3 网址/版面｜E4 类型标识「类型/OL」｜E5 分隔点全角")

    if res["n_online"] == 0 and res["n_n"] == 0:
        print("\n（本文无电子文献与报纸条目，官方格式串不适用；"
              "网络资源应标 [EB/OL] 并给引用日期与网址）")
        return 0

    print(f"\n【汇总】不合规 {len(res['reds'])} 处（硬伤）／"
          f"提示 {len(res['notices'])} 处")
    labels = {
        "E1": "缺方括号引用日期 / 误用圆括号 / N 类缺日期段",
        "E2": "引用日期或报纸日期月日未补 0",
        "E3": "缺网址 / 报纸缺版面",
        "E4": "电子文献类型标识形态错",
        "E5": "分隔点半角（提示）",
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
        print("\n【修法】改为官方形态：作者．题名[EB/OL]．"
              "[2023-06-05]．网址．（月日单数前补 0，方括号，网址必给）")
    else:
        print("\n✅ 电子文献条目均符合官方 §7 格式。")

    if res["notices"]:
        print("\n【提示】")
        for x in res["notices"][:8]:
            print(f"  · [{x['num']}] {x['msg']}")

    print("\n【边界】只核电子/报纸条目的格式与日期补 0；文献真实性与可访问性"
          "需人工点击核验。")

    if "--json" in _flags:
        jf = os.path.splitext(DOC)[0] + ".ref_online.json"
        with open(jf, "w", encoding="utf-8") as fh:
            json.dump({k: v for k, v in res.items() if k != "nums"},
                      fh, ensure_ascii=False, indent=1)
        print(f"\nJSON 已写入：{jf}")

    return 1 if res["reds"] else 0


if __name__ == "__main__":
    sys.exit(main())
