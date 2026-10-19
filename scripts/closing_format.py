# -*- coding: utf-8 -*-
"""closing_format.py —— 文末序列核验（官方 §4 格式序列后半段）

官方投稿须知 §4「格式顺序（官方明文）」给出的文末序列为：

    正文 → 结语 → 注释 → 参考文献 → 作者简介 →
    邮编及详细通信地址、电话、电子邮箱、定稿日期

并特别注明：「官方序列里末章写的是『**结语**』，且正文后是『**注释**』栏
（注释与参考文献同列问句号）」。已有工具链只查这些元素「是否存在」
（diagnose_paper 的官方合规维度报「文末未见作者简介/联系方式/定稿日期」
三条，submission_pack 查要素齐备），不查**先后次序与完整性**：

  · 次序：作者简介出现在参考文献之前＝顺序错；结语之后直接跳到
    参考文献而跳过「注释」栏＝缺栏（须提示）；
  · 完整性：注释栏、参考文献、作者简介、通信方式、定稿日期五段是否
    都在文末；
  · 「注释」缺失但有脚注式注释：读 docx 的 footnotes.xml / endnotes.xml，
    正文挂脚注而文末无「注释」栏时提示该并入注释栏（官方「注释与参考
    文献同列问句号」）；
  · 结语命名（官方序列写「结语」而非「结论」）与参考文献标题形态。

用法:
    python closing_format.py <论文.docx> [--tail 45] [--json]
退出码：0 = 文末序列次序与完整；1 = 检出次序错误或缺文末必备段
        （结语/参考文献/作者简介/通信方式/定稿日期之一）。

说明：只核**文末段落顺序与存在性**；作者简介块内部的要素齐备度见
       submission_pack.py，形态与一致性见 author_meta.py。
"""
import os
import re
import sys
import json
import zipfile

# Windows 控制台 GBK 保护
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import diagnose_paper as dp  # noqa: E402

DEFAULT_TAIL = 45          # 文末观察窗（官方序列后半段，段落数不大）

# 官方 §4 文末序列键序
TAIL_SEQ = ["jieyu", "zhushi", "refs", "bio", "contact", "date"]
TAIL_LABEL = {
    "jieyu": "结语", "zhushi": "注释栏", "refs": "参考文献",
    "bio": "作者简介", "contact": "邮编及通信地址/电话/电子邮箱",
    "date": "定稿日期",
}
# 必备段（缺任一即 🔴）
REQUIRED = ["jieyu", "refs", "bio", "contact", "date"]

CJK_RE = re.compile(r"[\u4e00-\u9fff]")


def _notes_from_xml(path):
    """读 docx 包内 footnotes.xml / endnotes.xml 的实际注释文本。

    返回 {source: [注释文本, ...]}；无部件或为空时返回空表。
    """
    out = {"footnotes": [], "endnotes": []}
    try:
        z = zipfile.ZipFile(path)
    except Exception:
        return out
    for key, part in (("footnotes", "word/footnotes.xml"),
                      ("endnotes", "word/endnotes.xml")):
        try:
            xml = z.read(part).decode("utf-8", "ignore")
        except KeyError:
            continue
        # 逐 <w:footnote>/<w:endnote> 节点取文本
        for node in re.findall(r"<w:(?:footnote|endnote)\b.*?</w:(?:footnote|"
                               r"endnote)>", xml, re.S):
            txt = "".join(re.findall(r"<w:t[^>]*>(.*?)</w:t>", node, re.S))
            txt = txt.strip()
            # 剔除分隔符/续页符等系统注释（通常无文本或极短）
            if txt and len(txt) >= 2:
                out[key].append(txt)
    return out


def find_tail_elements(paras, doc_path="", tail_n=DEFAULT_TAIL):
    """在文末观察窗内定位官方序列各元素，返回 (found, notes, tail_paras)。

    found: {键: (段号, 原文)}；notes: XML 注释计数。

    作者简介/通信方式/定稿日期只在**参考文献标题之后**的区域认定——官方
    序列为「…注释 → 参考文献 → 作者简介 → 邮编及详细通信地址…」；若观察
    窗内未见参考文献标题，则退回整窗认定（防漏报）。
    """
    idxs = [(i, t) for i, (s, t) in enumerate(paras) if (t or "").strip()]
    tail = idxs[-tail_n:] if len(idxs) > tail_n else idxs
    found = {}
    notes = _notes_from_xml(doc_path) if doc_path else {"footnotes": [],
                                                        "endnotes": []}

    # 参考文献标题在观察窗内的位置
    ref_pos = None
    for i, (pi, raw) in enumerate(tail):
        t = (raw or "").strip()
        if re.match(r"^\s*参\s*考\s*文\s*献\s*$", t[:16]):
            ref_pos = i
            found["refs"] = (pi, t)
            break

    for i, (pi, raw) in enumerate(tail):
        t = (raw or "").strip()
        head = t[:16]
        # 结语（末章）
        if "jieyu" not in found and re.match(r"^\s*\d*\s*结\s*语\s*$", head):
            found["jieyu"] = (pi, t)
        # 注释栏
        if "zhushi" not in found and re.match(r"^\s*(?:注\s*释|【注\s*释】|"
                                              r"Notes?)\s*[：:]?\s*$", head):
            found["zhushi"] = (pi, t)
        # 作者简介
        if "bio" not in found and re.search(r"作者简介|作者信息|【作者简介】",
                                            head):
            found["bio"] = (pi, t)
        # 通信方式与定稿日期：只在参考文献标题之后的区域认定
        if ref_pos is None or i > ref_pos:
            if "contact" not in found and re.search(
                    r"邮(?:政编码|编)|通信地址|通讯地址|地址[：:]|电话|手机|"
                    r"@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}", t):
                found["contact"] = (pi, t)
            if "date" not in found and re.search(r"(?:定稿日期|收稿日期)\s*[：:]?"
                                                 r"\s*\d{4}", t):
                found["date"] = (pi, t)
    return found, notes, tail


def check_order(found):
    """次序检查：相邻必备段的段号必须递增。"""
    issues, notices = [], []
    seq = [k for k in TAIL_SEQ if k in found]
    for a, b in zip(seq, seq[1:]):
        pa, pb = found[a][0], found[b][0]
        if pb < pa:
            issues.append(
                f"次序错误：{TAIL_LABEL[b]}（段{pb}）排在"
                f"{TAIL_LABEL[a]}（段{pa}）之前——官方 §4 序列要求"
                f"「{TAIL_LABEL[a]}」在前")
    # 作者简介与参考文献的相对位置（最常见的错）
    if "bio" in found and "refs" in found and found["bio"][0] < found["refs"][0]:
        issues.append("作者简介出现在参考文献之前——官方序列为"
                      "「参考文献 → 作者简介」，请后移")
    # 结语应在参考文献之前
    if "jieyu" in found and "refs" in found and found["refs"][0] < \
            found["jieyu"][0]:
        issues.append("参考文献出现在结语之前——官方序列为"
                      "「正文 → 结语 → 注释 → 参考文献」")
    return issues, notices


def analyze(doc):
    paras, refs, abs_text, kw_text, titles, body = dp.extract_paras(doc)
    tail_n = DEFAULT_TAIL
    found, notes, tail = find_tail_elements(paras, doc, tail_n)
    issues, notices = check_order(found)

    # ---------- 完整性 ----------
    missing = [k for k in REQUIRED if k not in found]
    for k in missing:
        issues.append(f"文末缺「{TAIL_LABEL[k]}」——官方 §4 格式序列要求")

    # ---------- 注释栏缺失 vs 脚注式注释 ----------
    n_fn = len(notes["footnotes"])
    n_en = len(notes["endnotes"])
    if "zhushi" not in found and (n_fn or n_en):
        notices.append(
            f"文末未见「注释」栏，但稿件含 {n_fn} 条脚注 / {n_en} 条尾注——"
            f"官方 §4 序列在正文后有「注释」栏（注释与参考文献同列问句号），"
            f"脚注式注释请确认刊物是否接受，或整理为文末注释栏")
    elif "zhushi" in found and not (n_fn or n_en):
        notices.append("文末有「注释」栏 ✓（与参考文献同列，符合官方序列）")

    # ---------- 末章命名 ----------
    last_title = ""
    for s, t in paras:
        tt = (t or "").strip()
        if tt and (re.match(r"^\d+\s", tt) or re.match(r"^\s*结\s*语", tt)):
            if not re.match(r"^\s*参\s*考\s*文\s*献", tt):
                last_title = tt
    if last_title and re.search(r"结\s*论|结束语|余\s*论", last_title) \
            and not re.search(r"结\s*语", last_title):
        issues.append(f"末章名为「{last_title[:20]}」——官方 §4 序列写"
                      f"「结语」，建议对齐")

    stat = {"found": {k: {"para": v[0], "text": v[1][:40]}
                      for k, v in found.items()},
            "notes": notes, "missing": missing,
            "tail_n": len(tail), "last_title": last_title}
    return {"issues": issues, "notices": notices, "stat": stat}


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    flags = set(a for a in sys.argv[1:] if a.startswith("--"))
    if not args or args[0] in ("-h", "--help"):
        print(__doc__)
        sys.exit(0)
    DOC = args[0]
    if not os.path.isfile(DOC):
        sys.exit(f"找不到文件：{DOC}")

    tail_n = DEFAULT_TAIL
    for i, a in enumerate(sys.argv):
        if a == "--tail" and i + 1 < len(sys.argv):
            try:
                tail_n = max(10, int(sys.argv[i + 1]))
            except ValueError:
                pass

    res = analyze(DOC)
    st = res["stat"]
    print("=" * 72)
    print("文末序列核验（官方须知 §4 格式顺序后半段）")
    print(f"稿件：{os.path.basename(DOC)}")
    print(f"观察窗：文末 {st['tail_n']} 个非空段落")
    print("=" * 72)

    print("\n【官方序列实测】")
    for k in TAIL_SEQ:
        if k in st["found"]:
            v = st["found"][k]
            print(f"  ✓ {TAIL_LABEL[k]:<28} 段{v['para']:>3}  {v['text'][:36]}")
        else:
            mark = "🔴 缺失" if k in REQUIRED else "— 未见"
            print(f"  ✗ {TAIL_LABEL[k]:<28} {mark}")
    nf, ne = len(st["notes"]["footnotes"]), len(st["notes"]["endnotes"])
    print(f"  · 稿件内脚注 {nf} 条 ｜ 尾注 {ne} 条（XML 层实读）")

    print("\n" + "-" * 72)
    if res["issues"]:
        print(f"🔴 发现 {len(res['issues'])} 项文末序列问题：")
        for it in res["issues"]:
            print(f"   · {it}")
    else:
        print("✅ 文末序列次序与完整性符合官方 §4 格式顺序。")
    if res["notices"]:
        print(f"\n🟡 建议核对 {len(res['notices'])} 项：")
        for it in res["notices"]:
            print(f"   · {it}")

    if "--json" in flags:
        jf = os.path.splitext(DOC)[0] + ".closing_format.json"
        with open(jf, "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=1)
        print(f"\nJSON 已写入：{jf}")

    sys.exit(1 if res["issues"] else 0)


if __name__ == "__main__":
    main()
