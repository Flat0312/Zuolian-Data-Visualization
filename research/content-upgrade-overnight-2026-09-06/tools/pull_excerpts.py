# -*- coding: utf-8 -*-
"""按关系ID批量提取被引证据摘录，写入 R/work/batch_*.md 供逐条预审。只读生产数据与原文。"""
import csv, json, re, os, sys
sys.path.insert(0, os.path.dirname(__file__))
import corpus

R = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
ROOT = corpus.ROOT
WORK = os.path.join(R, "work")
os.makedirs(WORK, exist_ok=True)

PERSONS = {r["person_id"]: r for r in csv.DictReader(open(os.path.join(ROOT, "data", "processed", "persons.csv"), encoding="utf-8-sig"))}
RELS = {r["relation_id"]: r for r in csv.DictReader(open(os.path.join(ROOT, "data", "processed", "person_relations.csv"), encoding="utf-8-sig"))}

DP = corpus.dict_pages()
HP = corpus.hist_pages()
DI = corpus.diary_index()
MM = corpus.memoir()

def name_and_aliases(pid):
    p = PERSONS.get(pid, {})
    names = [p.get("standard_name", "")]
    al = (p.get("aliases") or "").replace("；", ";").replace("、", ";").replace(",", ";").split(";")
    names += [a.strip() for a in al if a.strip()]
    return [n for n in dict.fromkeys(names) if n]

def diary_loc(ys, ms, ds):
    e = DI.get((int(ys), int(ms), int(ds)))
    if e is None:
        return None
    return "鲁迅日记 %s-%02d-%02d条：%s" % (ys, int(ms), int(ds), e[:220])

def page_with_names(pages, src, no, names):
    t = pages.get(int(no))
    if t is None:
        return "%s 第%s页：[页码未命中OCR页索引]" % (src, no)
    # focus window around first name hit if any
    pos = -1
    for n in names:
        if n and n in t:
            pos = t.find(n); break
    if pos >= 0:
        lo = max(0, pos - 100); hi = min(len(t), pos + 300)
        frag = t[lo:hi].replace("\n", " ")
    else:
        frag = t[:300].replace("\n", " ")
    return "%s 第%s页：%s" % (src, no, frag)

def excerpts_for(rid):
    r = RELS[rid]
    names = name_and_aliases(r["source_person_id"]) + name_and_aliases(r["target_person_id"])
    out = []
    for c in (r["evidence_ref"] or "").split(";"):
        c = c.strip()
        if not c:
            continue
        m = re.match(r"鲁迅日记\s*(\d{4})年(\d{1,2})月(\d{1,2})日", c)
        if m:
            d = diary_loc(m.group(1), m.group(2), m.group(3))
            out.append(d or "鲁迅日记 %s-%s-%s：[本地未命中该日期条目]" % m.groups())
            continue
        m = re.match(r"左联词典\s*第(\d+)页", c)
        if m:
            out.append(page_with_names(DP, "左联词典", m.group(1), names)); continue
        m = re.match(r"左联史\s*第(\d+)页", c)
        if m:
            out.append(page_with_names(HP, "左联史", m.group(1), names)); continue
        m = re.match(r"左联回忆录\s*第(\d+)页", c)
        if m:
            t = MM.get(m.group(1))
            out.append("左联回忆录 第%s页：%s" % (m.group(1), (t or "[OCR JSON无该页]")[:300].replace("\n", " ")))
            continue
        out.append("[未解析定位] " + c)
    return r, out

def main(ids, outname):
    lines = []
    for rid in ids:
        r, exs = excerpts_for(rid)
        sp = PERSONS.get(r["source_person_id"], {}).get("standard_name", r["source_person_id"])
        tp = PERSONS.get(r["target_person_id"], {}).get("standard_name", r["target_person_id"])
        lines.append("## %s | %s(%s) - %s(%s) | 类型:%s" % (
            rid, sp, r["source_person_id"], tp, r["target_person_id"],
            r["final_relation_type"] or r["original_relation_type"]))
        ctx = (r["context"] or "").replace("\n", " ")[:260]
        lines.append("生产context: " + ctx)
        for e in exs:
            lines.append("  * " + e)
        lines.append("")
    p = os.path.join(WORK, outname)
    open(p, "w", encoding="utf-8").write("\n".join(lines))
    print("wrote", p, len(ids), "relations")

if __name__ == "__main__":
    b = json.load(open(os.path.join(ROOT, "research", "briefs", "content-upgrade-overnight-2026-09-06", "baseline.json"), encoding="utf-8"))
    s = b["selection"]
    if sys.argv[1] == "ids":
        main(sys.argv[2:], "batch_custom.md")
    else:
        start, end = int(sys.argv[1]), int(sys.argv[2])
        ids = s["relation_sample_ids"][start:end]
        main(ids, "batch_%03d_%03d.md" % (start, end))
