"""来源层级：作品(work) / 引文(passage) / 来源族(family)。

- source_works.csv: 文献作品（去重后的独立作品）
- source_passages.csv: 具体引文（每条 sources.csv 对应一条 passage，保留 source_id 兼容映射）
- sources.csv 新增 source_family 列：同一作品的不同引文共享同一 family，
  同一史料的不同载体（如鲁迅日记纸本转录与维基文库）共享同一 family。
"""
from __future__ import annotations

import hashlib
import re
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[2]


def source_family_for(title: object, source_path: object, source_url: object) -> str:
    t = str(title or "").strip()
    u = str(source_url or "").strip()
    # 鲁迅日记各载体同族
    if "鲁迅日记" in t or "日记" in t and ("鲁迅" in t or "wikisource" in u):
        return "luxun_diary"
    if t == "鲁迅日记" or t.startswith("鲁迅日记·"):
        return "luxun_diary"
    if t in ("左联词典",):
        return "zuolian_cidian"
    if t in ("左联史",):
        return "zuolian_shi"
    if t in ("左联回忆录",) or t.startswith("左联回忆录"):
        return "zuolian_memoir"
    if t == "《左联相关档案资源目录》原始表格":
        return "raw_workbook"
    if "shhk.gov.cn" in u:
        return "shhk_gov"
    if "cpc.people.com.cn" in u or "dangshi.people.com.cn" in u:
        return "party_history"
    if "chinawriter.com.cn" in u:
        return "chinawriter"
    if "thepaper.cn" in u:
        return "thepaper"
    if "wikipedia.org" in u or "baike.baidu.com" in u:
        return "encyclopedia"
    if "shu.edu.cn" in u:
        return "shu_archive"
    if "ccphistory.org.cn" in u:
        return "ccp_history"
    if u:
        # 其余每个域名自成一族（域名级），避免不同作品被合并
        m = re.search(r"https?://([^/]+)/", u)
        if m:
            return "web_" + m.group(1).replace(".", "_")
        return "web_other"
    # 无 URL 的本地引文按标题分族
    if t in ("未分类引文", ""):
        return "unclassified"
    return "local_" + re.sub(r"\W+", "_", t)[:40].strip("_") or "local_unknown"


def portable_source_path(raw: object) -> str:
    """仓库内绝对路径改为仓库相对路径；仓库外路径原样保留；空值保持空。"""
    s = str(raw or "").strip()
    if not s:
        return ""
    try:
        p = Path(s)
        # Windows 绝对路径且位于仓库内 -> 相对路径（posix 风格，保证可移植）
        try:
            rel = p.resolve().relative_to(PROJECT_ROOT.resolve())
            return rel.as_posix()
        except Exception:
            pass
        # 已经是相对路径形态（含 research/ 前缀）则规范为 posix
        if "research/" in s.replace("\\", "/") or s.startswith("research"):
            return s.replace("\\", "/")
        # 仓库外的绝对路径不得伪造相对路径，原样返回
        return s
    except Exception:
        return s


def file_sha256(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()
