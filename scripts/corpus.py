# -*- coding: utf-8 -*-
"""corpus.py —— 《图书馆学研究》语料库自动定位与加载模块

被 scripts/ 下的 diagnose_paper.py / search.py / verify_refs.py /
benchmark_paper.py 共用。职责：

1. 定位语料库目录（含 corpus_stats.json / paragraphs.json / index.md 的目录）
2. 加载统计基准（corpus_stats.json，108 篇全刊统计）
3. 加载段落库（paragraphs.json，118 篇段落级检索）

定位优先级（从高到低）：
    A. 命令行 --corpus <dir>           （最明确）
    B. 环境变量 TSGXYJ_CORPUS=<dir>
    C. 当前工作目录（cwd）直接含语料文件
    D. 脚本所在目录向上逐级寻找（技能目录→上级→...）
    E. 用户登记文件 ~/.tsgxyj_corpus（每行一个路径）
    F. 技能仓库内的 ../corpus 目录

找不到语料库时：诊断工具降级为内置静态基线（风格统计表中的常数值），
检索 / 溯源工具明确报错提示，不静默产出错误结果。
"""
import os
import json
import sys

# Windows 控制台默认 GBK 无法输出部分 UTF-8 字符，统一走 UTF-8
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# 语料库标志文件：存在其一即视为语料库目录
_MARKERS = ("corpus_stats.json", "paragraphs.json")

# 用户级配置文件（每行一个候选路径，# 开头为注释）。与 skill 一起分发时不含任何
# 本机私有路径，用户可自行登记自己的语料库位置，或改用 TSGXYJ_CORPUS 环境变量。
_USER_CONFIG = os.path.join(os.path.expanduser("~"), ".tsgxyj_corpus")

# 通用候选位置（不含任何作者本机路径）
_CANDIDATES = [
    # 技能仓库内的默认位置（随技能分发，可选）
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "corpus"),
]


def _user_candidates():
    """读取 ~/.tsgxyj_corpus 中登记的候选路径（每行一个，支持 ~ 展开）。"""
    if not os.path.isfile(_USER_CONFIG):
        return []
    out = []
    try:
        with open(_USER_CONFIG, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                out.append(os.path.expanduser(line))
    except Exception:
        return []
    return out

_STATIC_BASE = None  # 内置静态基线，懒加载


def _is_corpus_dir(path):
    """判断目录是否含语料标志文件。"""
    if not path or not os.path.isdir(path):
        return False
    return any(os.path.isfile(os.path.join(path, m)) for m in _MARKERS)


def find_corpus_dir(cli_corpus=None, env="TSGXYJ_CORPUS"):
    """定位语料库目录，返回绝对路径；找不到返回 None。"""
    # A. 命令行
    if cli_corpus:
        p = os.path.abspath(cli_corpus)
        if _is_corpus_dir(p):
            return p
        raise FileNotFoundError(f"--corpus 指定的目录不是语料库：{p}（缺 corpus_stats.json / paragraphs.json）")
    # B. 环境变量
    if env and os.environ.get(env):
        p = os.path.abspath(os.environ[env])
        if _is_corpus_dir(p):
            return p
        raise FileNotFoundError(f"环境变量 {env} 指向的目录不是语料库：{p}")
    # C. 当前工作目录
    if _is_corpus_dir(os.getcwd()):
        return os.path.abspath(os.getcwd())
    # D. 脚本目录向上逐级
    cur = os.path.dirname(os.path.abspath(__file__))
    while True:
        if _is_corpus_dir(cur):
            return cur
        parent = os.path.dirname(cur)
        if parent == cur:
            break
        cur = parent
    # E. 用户登记的位置（~/.tsgxyj_corpus）
    for c in _user_candidates():
        if _is_corpus_dir(c):
            return os.path.abspath(c)
    # F. 技能仓库内的通用候选
    for c in _CANDIDATES:
        if _is_corpus_dir(c):
            return os.path.abspath(c)
    return None


def load_corpus_stats(corpus_dir=None):
    """加载 corpus_stats.json（108 篇逐篇统计明细）。返回 list[dict]。"""
    d = corpus_dir or find_corpus_dir()
    path = os.path.join(d, "corpus_stats.json") if d else None
    if not path or not os.path.isfile(path):
        return []
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    return data if isinstance(data, list) else []


def load_paragraphs(corpus_dir=None):
    """加载 paragraphs.json（118 篇段落库）。返回 list[dict]。"""
    d = corpus_dir or find_corpus_dir()
    path = os.path.join(d, "paragraphs.json") if d else None
    if not path or not os.path.isfile(path):
        return []
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    return data if isinstance(data, list) else []


def load_reference_stats(corpus_dir=None):
    """从语料库统计中汇总「该刊基准」（各维度的中位/四分位/均值）。

    返回 dict：
        paras:  段长统计（中位 / P25 / P75）
        sents:  句均统计
        cites:  引用密度统计（/千字）
        refs:   参考文献条数统计
        abs:    摘要长度统计
        kws:    关键词数统计
        figs/tabs: 图表数统计
        conn:   连接词频率基线（次/千字）
        benwen: 「本文」频次基线（次/万字）
    每项含 med / q25 / q75 / mean / n。
    语料库缺失时回退到 style_stats.md 记录的内置静态值（标记 "static": true）。
    """
    rows = load_corpus_stats(corpus_dir)
    if rows:
        return _from_rows(rows)
    return _static_base()


def _pct(arr, p):
    arr = sorted(x for x in arr if x is not None)
    if not arr:
        return 0
    i = min(len(arr) - 1, int(len(arr) * p))
    return arr[i]


def _stat(name, arr):
    if not arr:
        return {"med": 0, "q25": 0, "q75": 0, "mean": 0.0, "n": 0}
    return {
        "med": _pct(arr, 0.5),
        "q25": _pct(arr, 0.25),
        "q75": _pct(arr, 0.75),
        "mean": sum(arr) / len(arr),
        "n": len(arr),
    }


def _from_rows(rows):
    ok = [r for r in rows if r.get("has_indent", True) and r.get("max_para", 0) <= 1500]
    result = {
        "paras": _stat("段长", [r["para_med"] for r in ok if r.get("para_med")]),
        "sents": _stat("句均", [r["sent_mean"] for r in ok if (r.get("sent_mean") or 0) > 0]),
        "cites": _stat("引用密度", [r["cite_rate"] for r in ok if r.get("cite_rate") is not None]),
        "refs": _stat("文献条数", [r["ref_cnt"] for r in rows if r.get("ref_cnt")]),
        "abs": _stat("摘要长度", [r["abstract_len"] for r in rows if 50 < (r.get("abstract_len") or 0) < 800]),
        "kws": _stat("关键词数", [r["kw_n"] for r in rows if r.get("kw_n")]),
        "figs": _stat("图数", [r["fig_n"] for r in rows if r.get("fig_n") is not None]),
        "tabs": _stat("表数", [r["tab_n"] for r in rows if r.get("tab_n") is not None]),
        "benwen": _stat("「本文」", [r["benwen"] / max(r.get("body_chars", 1), 1) * 10000
                                     for r in rows if r.get("body_chars")]),
        "static": False,
        "source": f"语料库 {len(rows)} 篇",
    }
    # 连接词频率（次/千字）——语料库 conn 是次计数，需要正文字数换算
    conn = {}
    total_chars = sum(r.get("body_chars", 0) for r in rows)
    if total_chars:
        for r in rows:
            for k, v in (r.get("conn") or {}).items():
                conn[k] = conn.get(k, 0) + v
        conn = {k: v / total_chars * 1000 for k, v in sorted(conn.items(), key=lambda x: -x[1])}
    result["conn"] = conn
    return result


def _static_base():
    """style_stats.md 记录的内置静态基准（语料库缺失时使用）。"""
    return {
        "paras": {"med": 206, "q25": 129, "q75": 326, "mean": 0, "n": 104},
        "sents": {"med": 63, "q25": 0, "q75": 0, "mean": 65.0, "n": 108},
        "cites": {"med": 3.23, "q25": 2.22, "q75": 4.48, "mean": 0, "n": 108},
        "refs": {"med": 40, "q25": 32, "q75": 57, "mean": 44.6, "n": 107},
        "abs": {"med": 262, "q25": 0, "q75": 0, "mean": 269.0, "n": 107},
        "kws": {"med": 4.4, "q25": 4, "q75": 5, "mean": 4.4, "n": 107},
        "figs": {"med": 2, "q25": 0, "q75": 0, "mean": 2.4, "n": 108},
        "tabs": {"med": 4, "q25": 0, "q75": 0, "mean": 4.3, "n": 108},
        "benwen": {"med": 2.04, "q25": 0, "q75": 0, "mean": 4.57, "n": 108},
        "conn": {"同时": 0.56, "因此": 0.36, "一方面": 0.26, "从而": 0.23,
                 "此外": 0.19, "另一方面": 0.14, "进而": 0.13, "然而": 0.08},
        "static": True,
        "source": "style_stats.md 内置静态值（语料库缺失）",
    }


def find_corpus_or_exit(cli_corpus=None, need="统计/检索"):
    """供命令行工具使用的便捷入口：找不到语料库时打印提示并退出。"""
    d = find_corpus_dir(cli_corpus)
    if d is None:
        print(f"[提示] 未找到语料库（{need} 需要 corpus_stats.json / paragraphs.json）。")
        print("       可任选其一指定：")
        print("         · 命令行加 --corpus <语料库目录>")
        print("         · 设置环境变量 TSGXYJ_CORPUS=<语料库目录>")
        print("         · 在语料库目录下直接运行本脚本")
        print(f"         · 在 {_USER_CONFIG} 中登记路径（每行一个）")
        sys.exit(2)
    return d
