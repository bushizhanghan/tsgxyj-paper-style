# -*- coding: utf-8 -*-
"""ref_format_style.py —— 参考文献标点形态核验（对照官方 §7 明文标点）

用途：
    只核「标点形态」这一件事：中文文献条目里作者/题名/刊名/年(期)/页码
    之间的分隔点与括号形态，是否与《图书馆学研究》官方投稿须知 §7 给出的
    格式串一致。要素完整性、类型标识、年份越界等由 ref_format.py /
    ref_type_check.py / ref_year_check.py 负责，本脚本不重复。

依据（references/official-guidelines.md §7，本 skill 唯一官方口径来源）：
    | 中文期刊论文 | 作者．题名[J]．刊名，年(年内期数)：页码． |
    | 英文期刊论文 | 作者．题名[J]．刊名，总期数(年内期数)：页码． |
    | 著作        | 作者．书名[M]．出版地：出版者，出版年：参考页码． |
    | 报纸文章    | 作者．题名[N]．报名,年-月-日(月日为单数前面加0)：版面． |
    | 电子文献    | 作者．题名[电子文献类型/电子文献载体]．[年-月-日](月日为单数前面加0)．网址． |
  由上述格式串可机械派生的标点规则：
    R1 分隔点一律用**全角「．」(U+FF0E)**——作者后、题名（类型标识）后、
       条目末尾，均为「．」；半角「.」(U+002E) 与中文句号「。」均不合规；
    R2 「年(期内期数)」的圆括号用**半角** `(` `)`（格式串原文即半角）；
    R3 页码区间用**半角连接号「-」**（格式串「页码」；该刊与 GB/T 7714
       通行形态为 `92-99`，不用 – — ～ － 等变体）；
    R4 作者姓名后的分隔点同 R1（「姓名后的分隔点」是最常见的漏改点）。

阈值/规则出处：以上 R1–R4 全部直接抄自 official-guidelines.md §7 的格式串
文字，无自造阈值。外文条目（作者段为拉丁字母）按 GB/T 7714-2015 用半角点，
不做全角要求，故外文条目只统计、不判红。

用法:
    python ref_format_style.py <论文.docx> [--json]

退出码:
    0 = 未发现标点形态问题（含「无中文条目」）
    1 = 发现标点形态问题（任一中文条目命中 R1–R4）
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
import ref_format as rf      # noqa: E402  （复用 TYPE_PAT / parse_item）

# ---- 官方 §7 标点基准（字符级常量，便于核对）----
FULL_STOP = "\uff0e"        # ． 全角句点（官方格式串的分隔点）
HALF_STOP = "."             # .  半角点（最常见误用）
CN_STOP = "\u3002"          # 。 中文句号（同样不合规）
HALF_PAREN = ("(", ")")     # 年(期) 官方用半角圆括号
FULL_PAREN = ("\uff08", "\uff09")
DASH_OK = "-"               # 页码区间连接号（半角）
DASH_BAD = "\u2013\u2014\uff5e\uff0d"   # – — ～ －
CJK = r"\u4e00-\u9fff"

# 屏蔽噪声：URL / DOI / 邮箱 / 小数与版本号 / 标准号——其中的半角点不算分隔点
NOISE_PATS = [
    re.compile(r"https?://\S+"),
    re.compile(r"DOI\s*[:：]?\s*\S+", re.I),
    re.compile(r"[\w.+-]+@[\w-]+\.[\w.]+"),
    re.compile(r"\d+\.\d+"),                                   # 2.1 / 14.11
    re.compile(r"(?:GB/T|GB|ISO|IEC)\s*[\d.]+"),               # GB/T 37668—2019
    re.compile(r"\[\s*\d{4}-\d{2}-\d{2}\s*\]"),                # [2026-09-01]
    re.compile(r"\(\s*\d{4}-\d{2}-\d{2}\s*\)"),                # (2023-06-28)
]


def mask_noise(text):
    """把非分隔点语境里的半角点屏蔽掉，避免 DOI/URL/小数误报。"""
    t = text
    for p in NOISE_PATS:
        t = p.sub(" ", t)
    return t


def is_cjk_item(raw):
    """中文条目判定：条目前 40 字符内出现汉字即按中文条目核全角标点。

    外文条目（KOTIS K, ANGOURA E. …）按 GB/T 7714 用半角点，不适用官方
    §7 的全角「．」要求，故排除。
    """
    head = raw[:40]
    return bool(re.search(r"[" + CJK + r"]", head))


def half_dots(masked):
    """在半角点候选位里挑出「分隔点」：前一字符为中文/]/数字/字母，
    后一字符为中文/[/字母/空白/结尾。返回下标列表。"""
    hits = []
    for m in re.finditer(r"\.", masked):
        i = m.start()
        prev = masked[i - 1] if i > 0 else ""
        nxt = masked[i + 1] if i + 1 < len(masked) else ""
        if re.match(r"[" + CJK + r"\]\)0-9A-Za-z]", prev or " "):
            if nxt == "" or re.match(r"[" + CJK + r"\[A-Za-z\s]", nxt):
                hits.append(i)
    return hits


def author_dot_pos(raw):
    """作者段结束点（姓名后的分隔点）在 raw 中的下标。

    复用 ref_format.parse_item 的切分口径：去掉序号后第一个「.」之前是
    作者段。条目无作者段（如标准号开头「信息技术 …: GB/T …[S].」）时返回
    None——那种条目的第一个点是类型标识后的分隔点，不归 R4。
    """
    m = re.match(r"^[\[［]?\s*\d{1,3}\s*[\]］]?\s*[:：.]?\s*(.+)$", raw, re.S)
    txt = m.group(1) if m else raw
    am = re.match(r"^([^.]{2,60}?)\.", txt)
    if not am:
        return None
    # txt 在 raw 中的起点
    start = raw.find(txt[:20])
    if start < 0:
        start = len(raw) - len(txt)
    return start + am.end(1)


def check_item(raw):
    """核单条。返回 findings: [(级别, 规则号, 说明)]，级别 red/notice。"""
    f = []
    num = rf.parse_item(raw)["num"]
    tag = f"[{num}]" if num else "[?]"

    if not is_cjk_item(raw):
        return f, tag, False

    masked = mask_noise(raw)
    dots = half_dots(masked)
    adot = author_dot_pos(raw)
    if dots:
        first = dots[0]
        seg = raw[max(0, first - 6):first + 6]
        if adot is not None and abs(first - adot) <= 1:
            # R4/R1：作者姓名后的分隔点
            f.append(("red", "R1/R4",
                      f"作者姓名后用半角「.」：…{seg}… 应为全角「．」"
                      f"（官方：作者．题名）"))
        else:
            f.append(("red", "R1",
                      f"分隔点用半角「.」：…{seg}… 应为全角「．」"))
        for i in dots[1:]:
            seg = raw[max(0, i - 8):i + 8]
            if adot is not None and abs(i - adot) <= 1:
                f.append(("red", "R1/R4",
                          f"作者姓名后用半角「.」：…{seg}… 应为全角「．」"
                          f"（官方：作者．题名）"))
            else:
                f.append(("red", "R1",
                          f"分隔点用半角「.」：…{seg}… 应为全角「．」"))
    # R1 变体：误用中文句号「。」
    for m in re.finditer(r"\u3002", masked):
        seg = raw[max(0, m.start() - 8):m.start() + 8]
        f.append(("red", "R1",
                  f"误用中文句号「。」：…{seg}… 应为全角间隔号「．」"))
    # R1：条目末尾点
    tail = raw.rstrip()
    if tail and tail[-1] in (HALF_STOP, CN_STOP):
        f.append(("red", "R1",
                  f"条目末尾用「{tail[-1]}」，官方格式串以全角「．」收尾"))
    # R2：年(期) 括号须半角
    for m in re.finditer(r"(\d{4})\s*(\()\s*(\d{1,3})\s*(\))", raw):
        f.append(("red", "R2",
                  f"年(期) 用全角括号「{m.group(2)}{m.group(4)}」："
                  f"{m.group(1)}{m.group(2)}{m.group(3)}{m.group(4)}，"
                  f"官方格式串为半角「{m.group(1)}({m.group(3)})」"))
    # R3：页码连接号须半角「-」
    for m in re.finditer(r"[：:]\s*(\d{1,4})\s*([" + DASH_BAD + r"])\s*(\d{1,4})",
                         raw):
        if m.group(2) != DASH_OK:
            f.append(("red", "R3",
                      f"页码区间用「{m.group(2)}」：{m.group(1)}{m.group(2)}"
                      f"{m.group(3)}，应统一半角「-」"))
    # R2 提示：中文期刊条目带「年, 卷(期)」——官方中文期刊格式串未列卷号
    tm = rf.TYPE_PAT.search(raw)
    if tm and tm.group(1).upper() == "J" and \
            re.search(r"\d{4}\s*[,，]\s*\d{1,3}\s*[(\uff08]\s*\d{1,3}", raw):
        f.append(("notice", "R2",
                  "中文期刊条目为「年, 卷(期)」形态；官方 §7 中文期刊格式串为"
                  "「年(期内期数)」未列卷号——GB/T 7714 允许，仅提示对位"))

    return f, tag, True


def analyze(doc):
    paras, refs, abs_text, kw_text, titles, body = dp.extract_paras(doc)
    all_f = []
    n_cjk = n_lat = 0
    for raw in refs:
        f, tag, cjk = check_item(raw)
        if cjk:
            n_cjk += 1
        else:
            n_lat += 1
        for lvl, rule, msg in f:
            all_f.append({"num": rf.parse_item(raw)["num"], "level": lvl,
                          "rule": rule, "msg": msg, "raw": raw[:70]})

    reds = [x for x in all_f if x["level"] == "red"]
    notices = [x for x in all_f if x["level"] == "notice"]

    def nums(items):
        return "".join(f"[{x['num']}]" for x in items[:24])

    by_rule = {}
    for x in all_f:
        by_rule.setdefault(x["rule"], []).append(x)

    return {"n_refs": len(refs), "n_cjk": n_cjk, "n_lat": n_lat,
            "reds": reds, "notices": notices, "by_rule": by_rule,
            "nums": nums, "findings": all_f}


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
    except Exception as e:  # 文献段缺失等
        print(f"解析失败（未找到参考文献段或条目？）：{e}")
        return 2
    if res["n_refs"] == 0:
        print("未解析到文献条目")
        return 2

    print("=" * 74)
    print("参考文献标点形态核验（官方 §7 明文标点）")
    print(f"稿件：{os.path.basename(DOC)}")
    print(f"文献 {res['n_refs']} 条：中文条目 {res['n_cjk']} 条（核全角标点）"
          f"／外文条目 {res['n_lat']} 条（GB/T 半角，不判红）")
    print("=" * 74)
    print("【官方标点基准】references/official-guidelines.md §7")
    print("  中文期刊论文：作者．题名[J]．刊名，年(年内期数)：页码．")
    print("  R1 分隔点全角「．」(U+FF0E)｜R2 年(期) 半角括号｜"
          "R3 页码区间半角「-」｜R4 姓名后分隔点同 R1")

    print(f"\n【汇总】中文条目标点问题 {len(res['reds'])} 处"
          f"（硬伤）／提示 {len(res['notices'])} 处")
    labels = {
        "R1": "分隔点/末尾点用半角「.」或「。」",
        "R1/R4": "作者姓名后分隔点用半角「.」",
        "R2": "年(期) 括号全角 / 中文期刊带卷号",
        "R3": "页码连接号非半角「-」",
    }
    for rule, items in sorted(res["by_rule"].items(),
                              key=lambda kv: -len(kv[1])):
        lvl = "🔴" if any(i["level"] == "red" for i in items) else "🟡"
        print(f"  {lvl} {labels.get(rule, rule)}：{len(items)} 处  "
              f"{res['nums'](items)}")

    if res["reds"]:
        print("\n【明细】中文条目逐条定位（前 12 条）")
        shown = 0
        seen = set()
        for x in res["reds"]:
            if x["num"] in seen:
                continue
            seen.add(x["num"])
            shown += 1
            if shown > 12:
                break
            cnt = sum(1 for y in res["reds"] if y["num"] == x["num"])
            print(f"  [{x['num']}] {cnt} 处｜{x['raw'][:52]}")
            for y in [z for z in res["reds"] if z["num"] == x["num"]][:3]:
                print(f"       · {y['msg']}")
        print("\n【修法】把中文条目的分隔点统一替换为全角「．」(U+FF0E)："
              "作者．题名[J]．刊名，年(期)：页码．")
    else:
        print("\n✅ 中文条目标点形态均符合官方 §7。")

    if res["notices"]:
        print("\n【提示】")
        for x in res["notices"][:8]:
            print(f"  · [{x['num']}] {x['msg']}")

    print("\n【边界】只核标点形态；文献真实性、要素完整性另见 ref_format.py／"
          "verify_refs.py。外文条目按 GB/T 7714 半角点，未判红。")

    if "--json" in _flags:
        jf = os.path.splitext(DOC)[0] + ".ref_punct.json"
        with open(jf, "w", encoding="utf-8") as fh:
            json.dump({k: v for k, v in res.items() if k != "nums"},
                      fh, ensure_ascii=False, indent=1)
        print(f"\nJSON 已写入：{jf}")

    return 1 if res["reds"] else 0


if __name__ == "__main__":
    sys.exit(main())
