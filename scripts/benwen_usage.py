# -*- coding: utf-8 -*-
"""benwen_usage.py —— 「本文」使用率核验（克制使用专查）

官方须知 §5 明文：「语气用**第三人称**。不要出现「**我们**」「**本文**」等字样。」
该刊实测（108 篇）：正文「本文」中位 2.1 次/万字、P75=7.3、P90=12.7，
且 **32.4% 的篇目完全不用「本文」**——主流是用「本研究」或直接陈述。

diagnose_paper 只报「本文」的绝对次数与密度，不判**是否超限**，也不区分
**出现位置**：摘要里出现是官方明文禁止（红线），正文里少量使用是该刊常态，
但超过 P75（7.3/万字）就显啰嗦。

本脚本按位置分级核验：
  · 摘要含「本文」→ 🔴 红线（官方明文禁止）；
  · 正文密度 > P75（7.3/万字）→ 🟡 提示改为「本研究」或删；
  · 正文密度在 P25–P75 之间 → ✅ 正常（该刊主流区间）；
  · 同时报「我们」的出现（官方同样点名）。

用法:
    python benwen_usage.py <论文.docx> [--json]
退出码：1 = 摘要含「本文」（红线级）。
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

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import diagnose_paper as dp  # noqa: E402

# 该刊 108 篇实测：「本文」次/万字
P25, P50, P75, P90 = 0.0, 2.1, 7.3, 12.7
ZERO_RATE = 0.324          # 完全不用「本文」的篇目占比


def analyze(doc):
    paras, refs, abs_text, kw_text, titles, body = dp.extract_paras(doc)
    abs_c = dp.clean(abs_text or "")
    body_c = dp.clean("\n".join(body))
    n_chars = max(len(body_c), 1)

    n_abs = abs_c.count("本文") + abs_c.count("我们")
    n_body_bw = body_c.count("本文")
    n_body_wm = body_c.count("我们")
    rate = n_body_bw / n_chars * 10000

    issues, notices = [], []
    if n_abs:
        issues.append(f"摘要含「本文/我们」×{n_abs}——官方须知明文禁止"
                      f"（「不要出现『我们』『本文』等字样」），必须改「本研究」或删")
    if rate > P75:
        notices.append(f"正文「本文」密度 {rate:.1f}/万字，超过该刊 P75（{P75}）——"
                       f"建议改「本研究」或删（该刊中位 {P50}，"
                       f"{ZERO_RATE:.0%} 篇目完全不用）")
    elif rate > P90:
        issues.append(f"正文「本文」密度 {rate:.1f}/万字，超过该刊 P90（{P90}）——"
                      f"显啰嗦，需精简")
    if n_body_wm:
        notices.append(f"正文含「我们」×{n_body_wm}——官方要求第三人称，建议改「本研究」")

    return {"abs_hits": n_abs, "body_bw": n_body_bw, "body_wm": n_body_wm,
            "rate": round(rate, 2), "n_chars": n_chars,
            "issues": issues, "notices": notices}


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    flags = set(a for a in sys.argv[1:] if a.startswith("--"))
    if not args or args[0] in ("-h", "--help"):
        print(__doc__)
        sys.exit(0)
    DOC = args[0]
    if not os.path.isfile(DOC):
        sys.exit(f"找不到文件：{DOC}")

    res = analyze(DOC)
    print("=" * 74)
    print("「本文」使用率核验（官方第三人称要求 + 该刊实测密度）")
    print(f"稿件：{os.path.basename(DOC)}")
    print(f"摘要「本文/我们」×{res['abs_hits']} ｜ 正文「本文」×{res['body_bw']}"
          f"（{res['rate']}/万字，正文 {res['n_chars']} 字）｜「我们」×{res['body_wm']}")
    print(f"该刊基线：中位 {P50}/万字 ｜ P75 {P75} ｜ P90 {P90} ｜ "
          f"{ZERO_RATE:.0%} 篇目完全不用")
    print("=" * 74)

    if res["issues"]:
        print(f"🔴 发现 {len(res['issues'])} 项问题：")
        for it in res["issues"]:
            print(f"   · {it}")
    else:
        print("✅ 「本文」使用符合官方要求与刊内密度。")
    if res["notices"]:
        print(f"\n🟡 建议核对 {len(res['notices'])} 项：")
        for it in res["notices"]:
            print(f"   · {it}")

    if "--json" in flags:
        jf = os.path.splitext(DOC)[0] + ".benwen_usage.json"
        with open(jf, "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=1)
        print(f"\nJSON 已写入：{jf}")

    sys.exit(1 if res["issues"] else 0)


if __name__ == "__main__":
    main()
