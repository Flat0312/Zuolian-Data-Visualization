"""来源层级：作品(work) / 引文(passage) / 来源族(family) + 统一注册/同步入口。

- source_works.csv: 文献作品（去重后的独立作品，work_id 稳定不重排）
- source_passages.csv: 具体引文（每条 sources.csv 记录有且只有一条 passage 映射）
- sources.csv 的 source_family 列：同一作品的不同引文共享同一 family。

sync_source_layer 是**唯一**的层级维护入口：
- 任何脚本新增/修改 sources.csv 后必须调用；
- 已有 work_id / passage_id 永不重排，新 ID 只在末尾追加；
- citation_count 一律按该 work 实际 passage 数量重算；
- 函数幂等：输入不变时输出字节不变。
"""
from __future__ import annotations

import hashlib
import re
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[2]

WORK_COLUMNS = (
    "work_id",
    "title",
    "author",
    "version",
    "publication_info",
    "source_category",
    "source_family",
    "citation_count",
)
PASSAGE_COLUMNS = ("passage_id", "work_id", "source_id", "locator", "citation", "file_hash", "source_url")


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


def _read_csv_strings(path: Path) -> list[dict[str, str]]:
    import csv

    if not path.exists():
        return []
    with open(path, encoding="utf-8-sig", newline="") as fh:
        return [dict((k, v if v is not None else "") for k, v in row.items()) for row in csv.DictReader(fh)]


def _write_csv(path: Path, columns: tuple[str, ...], rows: list[dict[str, str]]) -> None:
    import csv

    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=list(columns), extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def _next_numeric_id(prefix: str, existing_ids: list[str]) -> str:
    max_n = 0
    width = 5 if prefix == "PSGN" else 4
    for value in existing_ids:
        m = re.fullmatch(rf"{re.escape(prefix)}-(\d+)", str(value).strip())
        if m:
            max_n = max(max_n, int(m.group(1)))
            width = max(width, len(m.group(1)))
    # WORK 沿用 4 位、PSGN 沿用 5 位；已存在更宽时沿用更宽，保证字典序即数字序、不重排既有 ID。
    if prefix == "WORK":
        width = max(4, width)
    else:
        width = max(5, width)
    return f"{prefix}-{max_n + 1:0{width}d}"


def sync_source_layer(processed_dir: Path) -> dict[str, int]:
    """将 source_works / source_passages 与 sources.csv 完全对齐（统一注册入口）。

    规则：
    1. 每条 source 有且只有一条 passage（缺失补齐、重复去重保首个）；
    2. passage 引用有效 work_id / source_id，无效引用按作品键重建；
    3. 新 work / passage 的 ID 只追加不重排；
    4. citation_count 按各 work 实际 passage 数量重算；
    5. 无 passage 的孤儿 work 移除（保持引用闭合）。
    幂等：相同输入产生字节相同的输出。
    """
    processed_dir = Path(processed_dir)
    sources = _read_csv_strings(processed_dir / "sources.csv")
    works_rows = _read_csv_strings(processed_dir / "source_works.csv")
    passages_rows = _read_csv_strings(processed_dir / "source_passages.csv")

    source_by_id: dict[str, dict[str, str]] = {}
    for row in sources:
        sid = str(row.get("source_id", "")).strip()
        if sid:
            source_by_id[sid] = row

    work_row_by_id: dict[str, dict[str, str]] = {}
    for row in works_rows:
        wid = str(row.get("work_id", "")).strip()
        if not wid:
            continue
        merged = {col: str(row.get(col, "") or "") for col in WORK_COLUMNS}
        work_row_by_id[wid] = merged

    # 作品表不含 path/url 列，跨轮匹配以标题为主键；
    # 同题多作品时用既有 passage 的 source_url 消歧（确定性）。
    work_ids_by_title: dict[str, list[str]] = {}
    for wid, row in work_row_by_id.items():
        work_ids_by_title.setdefault(row["title"].strip(), []).append(wid)
    urls_by_work: dict[str, set[str]] = {}
    for row in passages_rows:
        wid = str(row.get("work_id", "")).strip()
        url = str(row.get("source_url", "")).strip()
        if wid:
            urls_by_work.setdefault(wid, set()).add(url)

    def _resolve_work_id(title: str, surl: str) -> str:
        candidates = sorted(work_ids_by_title.get(title, []))
        if not candidates:
            return ""
        if len(candidates) == 1:
            return candidates[0]
        with_url = [c for c in candidates if surl and surl in urls_by_work.get(c, set())]
        if with_url:
            return with_url[0]
        if not surl:
            no_url = [c for c in candidates if not urls_by_work.get(c, set())]
            if no_url:
                return no_url[0]
        return candidates[0]

    passage_by_source: dict[str, dict[str, str]] = {}
    for row in passages_rows:
        sid = str(row.get("source_id", "")).strip()
        if sid and sid not in passage_by_source:
            passage_by_source[sid] = {col: str(row.get(col, "") or "") for col in PASSAGE_COLUMNS}

    existing_work_ids = [str(row.get("work_id", "")).strip() for row in works_rows]
    existing_passage_ids = [str(row.get("passage_id", "")).strip() for row in passages_rows]
    taken_passage_ids: set[str] = set()
    hash_cache: dict[str, str] = {}

    def _hash_for(source_path: str) -> str:
        p = source_path.strip()
        if not p:
            return ""
        if p in hash_cache:
            return hash_cache[p]
        cand = Path(p) if Path(p).is_absolute() else PROJECT_ROOT / p
        value = ""
        if cand.is_file():
            try:
                value = file_sha256(cand)
            except Exception:
                value = ""
        hash_cache[p] = value
        return value

    new_passages: list[dict[str, str]] = []
    used_work_ids: set[str] = set()
    added_works = added_passages = 0
    for sid in sorted(source_by_id):
        src = source_by_id[sid]
        title = str(src.get("title", "")).strip()
        spath = str(src.get("source_path", "")).strip()
        surl = str(src.get("source_url", "")).strip()
        citation = str(src.get("citation", "")).strip()

        old_pas = passage_by_source.get(sid)
        wid = ""
        if old_pas:
            wid = str(old_pas.get("work_id", "")).strip()
            if wid not in work_row_by_id:
                wid = ""
        if not wid:
            wid = _resolve_work_id(title, surl)
            if not wid:
                wid = _next_numeric_id("WORK", existing_work_ids)
                existing_work_ids.append(wid)
                _fam = str(src.get("source_family", "")).strip()
                if not _fam:
                    _fam = source_family_for(title, spath, surl)
                work_row_by_id[wid] = {
                    "work_id": wid,
                    "title": title,
                    "author": "鲁迅" if "鲁迅日记" in title else "",
                    "version": "",
                    "publication_info": "待人工核录",
                    "source_category": str(src.get("evidence_type", "")).strip(),
                    "source_family": _fam,
                    "citation_count": "0",
                }
                work_ids_by_title.setdefault(title, []).append(wid)
                urls_by_work.setdefault(wid, set()).add(surl)
                added_works += 1
        used_work_ids.add(wid)

        pid = str(old_pas.get("passage_id", "")).strip() if old_pas else ""
        if not pid or pid in taken_passage_ids:
            pid = _next_numeric_id("PSGN", existing_passage_ids)
            existing_passage_ids.append(pid)
            added_passages += 1
        taken_passage_ids.add(pid)
        new_passages.append(
            {
                "passage_id": pid,
                "work_id": wid,
                "source_id": sid,
                "locator": citation[:500],
                "citation": citation[:800],
                "file_hash": _hash_for(spath),
                "source_url": surl,
            }
        )

    # citation_count 按实际 passage 数量重算；空 family 回填（旧基线无 family 列时）。
    count_by_work: dict[str, int] = {}
    for pas in new_passages:
        count_by_work[pas["work_id"]] = count_by_work.get(pas["work_id"], 0) + 1
    # 用于回填的标题->family 映射（取首个非空来源 family，否则按规则计算）
    family_by_title: dict[str, str] = {}
    for sid in sorted(source_by_id):
        src = source_by_id[sid]
        title = str(src.get("title", "")).strip()
        fam = str(src.get("source_family", "")).strip()
        if title and fam and title not in family_by_title:
            family_by_title[title] = fam
    final_works = []
    for wid in sorted(work_row_by_id):
        if wid not in used_work_ids:
            continue
        row = dict(work_row_by_id[wid])
        if not row.get("source_family", "").strip():
            _t = row.get("title", "").strip()
            row["source_family"] = family_by_title.get(_t, "") or source_family_for(
                _t, "", ""
            )
        row["citation_count"] = str(count_by_work.get(wid, 0))
        final_works.append(row)

    _write_csv(processed_dir / "source_works.csv", WORK_COLUMNS, final_works)
    _write_csv(processed_dir / "source_passages.csv", PASSAGE_COLUMNS, new_passages)
    return {
        "sources": len(source_by_id),
        "works": len(final_works),
        "passages": len(new_passages),
        "added_works": added_works,
        "added_passages": added_passages,
    }
