"""引文佐证校验：判断一条引文是否真的记载了某条人物关系。

背景（2026-09-28 实测）
----------------------
夜间轮回源核查（``research/content-upgrade-overnight-2026-09-06``）记录的 ``quote``
多数是按页码定位截取的固定窗口，而不是真正记载该关系的那一句。48 条被判定为
``evidence_support=support`` 且人工裁决成立的关系中，只有 17 条的引文同时出现两位当事人；
28 条的引文里找不到当事人（但同一本地原文中确实存在同时记载两人的段落，属引文窗口捕错）；
3 条在本地原文中找不到任何同窗共现。

因此 ``evidence_support=support`` 不能单独作为公开依据，必须再过本模块的**双方佐证门**：
引文（空白归一后）须同时含甲方与乙方的姓名或别名。鲁迅日记为日记体，作者即甲方，
只需引文含乙方（``diary_author_implicit``）。

本模块只做只读判定，不改写任何数据，也不降低 ``relation_risk_level``。
判定从严：只做精确字符串匹配，不做 OCR 异体字容错——漏判进入重捕队列，不误判为已佐证。
"""

from __future__ import annotations

import re
from pathlib import Path

DEFAULT_WINDOW = 260
NAME_SPLIT = re.compile(r"[;；、,，/|]")


def normalized_bundle(path: Path) -> tuple[str, str, list[int]]:
    """返回 (原文, 去空白文本, 去空白位置 -> 原文位置 的索引表)。

    OCR 全文在每个字之间插入空格，直接检索会全部落空，必须先去空白再定位。
    """
    text = Path(path).read_text(encoding="utf-8", errors="replace")
    buf: list[str] = []
    idx: list[int] = []
    for i, ch in enumerate(text):
        if not ch.isspace():
            buf.append(ch)
            idx.append(i)
    return text, "".join(buf), idx


def name_candidates(person_row: dict[str, str]) -> list[str]:
    """人物在原文中可能出现的写法：标准名 + 别名（过滤单字，避免误命中）。"""
    out = [str(person_row.get("standard_name", "")).strip()]
    for alias in NAME_SPLIT.split(str(person_row.get("aliases", "") or "")):
        alias = alias.strip()
        if alias:
            out.append(alias)
    seen: list[str] = []
    for name in out:
        if len(name) >= 2 and name not in seen:
            seen.append(name)
    return seen


def quote_attests_pair(
    quote: str,
    names_a: list[str],
    names_b: list[str],
    diary_author_implicit: bool = False,
) -> tuple[bool, str]:
    """引文是否同时记载双方。返回 (是否佐证, 判定依据)。"""
    flat = re.sub(r"\s+", "", str(quote or ""))
    if not flat:
        return False, "empty_quote"
    hit_a = next((n for n in names_a if n in flat), "")
    hit_b = next((n for n in names_b if n in flat), "")
    if hit_b and (hit_a or diary_author_implicit):
        basis = f"A={hit_a or 'diary_author_implicit'};B={hit_b}"
        return True, basis
    if not hit_b and not hit_a:
        return False, "neither_party_in_quote"
    return False, f"only_{'A' if hit_a else 'B'}_in_quote"


def find_cooccurrence(
    bundle: tuple[str, str, list[int]],
    names_a: list[str],
    names_b: list[str],
    window: int = DEFAULT_WINDOW,
    limit: int = 5,
) -> list[dict[str, object]]:
    """在原文中找同时含双方的最窄窗口，供重捕引文使用（只读，不写生产层）。"""
    text, flat, idx = bundle
    found: list[dict[str, object]] = []
    for a in names_a:
        start = 0
        while True:
            pos = flat.find(a, start)
            if pos < 0:
                break
            start = pos + 1
            lo = max(0, pos - window)
            hi = pos + len(a) + window
            segment = flat[lo:hi]
            for b in names_b:
                bpos = segment.find(b)
                if bpos < 0:
                    continue
                distance = abs((lo + bpos) - pos)
                origin = idx[min(lo + bpos, len(idx) - 1)]
                excerpt = re.sub(r"\s+", "", text[max(0, origin - 170): origin + 240])
                found.append({
                    "name_a": a, "name_b": b, "distance": distance,
                    "normalized_pos": pos, "origin_pos": origin, "excerpt": excerpt,
                })
            if len(found) >= limit * max(1, len(names_a)):
                break
    found.sort(key=lambda item: (int(item["distance"]), int(item["normalized_pos"])))
    return found[:limit]


def looks_like_name_list(excerpt: str) -> bool:
    """粗判窗口是否只是顿号人名罗列（共现级），而非叙述性互动记载。

    仅用于给重捕队列打提示标签，不作为最终裁决；人名密集且缺少常见互动动词即视为罗列。
    """
    flat = re.sub(r"\s+", "", str(excerpt or ""))
    if not flat:
        return False
    separators = flat.count("、") + flat.count(",") + flat.count("，")
    verbs = ("拜访", "见面", "联名", "合编", "介绍", "通信", "来信", "复信", "寄", "访",
             "约", "同往", "参加", "发起", "编辑", "探", "会晤", "谈", "赠", "邀", "署名", "签名")
    density = separators / max(1, len(flat))
    has_verb = any(v in flat for v in verbs)
    return density > 0.06 and not has_verb
