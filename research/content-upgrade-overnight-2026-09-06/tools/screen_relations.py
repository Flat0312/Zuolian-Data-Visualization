# -*- coding: utf-8 -*-
"""半自动预筛：对每个关系按引用定位提取摘录并做名字匹配预分类。
输出 work/auto_screen.md，AI 执行者逐条人工复核后写 relations.jsonl。自动结果不是最终结论。"""
import csv, json, re, os, sys
sys.path.insert(0, os.path.dirname(__file__))
import corpus

R = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
ROOT = corpus.ROOT

PERSONS = {r["person_id"]: r for r in csv.DictReader(open(os.path.join(ROOT, "data", "processed", "persons.csv"), encoding="utf-8-sig"))}
RELS = {r["relation_id"]: r for r in csv.DictReader(open(os.path.join(ROOT, "data", "processed", "person_relations.csv"), encoding="utf-8-sig"))}
DP, HP, DI, MM = corpus.dict_pages(), corpus.hist_pages(), corpus.diary_index(), corpus.memoir()

# 日记中常用简称（手工核对过的映射，仅用于检索，不自动判定身份）
SHORT = {"ZLH-005": ["雪峰"], "ZLH-024": ["许广平", "广平"], "ZLH-124": ["小峰"], "ZLH-128": ["林语堂", "语堂"],
         "ZLH-015": ["郁达夫", "达夫"], "ZLH-016": ["柔石"], "ZLH-018": ["冯铿"], "ZLH-021": ["丁玲"],
         "ZLH-003": ["瞿秋白", "秋白", "何凝", "维宁"], "ZLH-002": ["茅盾", "玄珠", "明甫"], "ZLH-006": ["潘汉年"],
         "ZLH-004": ["夏衍", "沈端先"], "ZLH-008": ["田汉", "寿昌"], "ZLH-010": ["洪深"], "ZLH-014": ["冯乃超"],
         "ZLH-017": ["胡也频"], "ZLH-025": ["沙汀"], "ZLH-026": ["胡风", "张光人"], "ZLH-032": ["徐懋庸", "懋庸"],
         "ZLH-036": ["穆木天"], "ZLH-037": ["楼适夷", "适夷"], "ZLH-039": ["欧阳山"], "ZLH-041": ["骆宾基"],
         "ZLH-048": ["王任叔", "任叔"], "ZLH-091": ["叶以群", "以群"], "ZLH-123": ["邹韬奋", "韬奋"],
         "ZLH-133": ["斯诺", "Snow"], "ZLH-141": ["高尔基"]}

def forms(pid):
    p = PERSONS.get(pid, {})
    out = [p.get("standard_name", "")]
    al = (p.get("aliases") or "").replace("；", ";").replace("、", ";").replace(",", ";").replace("/", ";").split(";")
    out += [a.strip() for a in al if a.strip()]
    out += SHORT.get(pid, [])
    n = p.get("standard_name", "")
    if len(n) >= 3 and n[:2] in ("欧阳", "司徒", "司马"):
        out.append(n[2:])
    elif len(n) >= 2:
        out.append(n[1:])  # 日记惯用去姓称
    seen, res = set(), []
    for x in out:
        if x and x not in seen and len(x) >= 2:
            seen.add(x); res.append(x)
    return res

def hits_in(text, forms_list):
    for f in forms_list:
        if f in text:
            return f
    return None

def locate(c, names):
    """返回 (locator_str, matched_bool, excerpt)"""
    m = re.match(r"鲁迅日记\s*(\d{4})年(\d{1,2})月(\d{1,2})日", c)
    if m:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
        e = DI.get((y, mo, d))
        if e is None:
            return c, False, "[日记本地未命中该日期条目]"
        f = hits_in(e, names)
        return c, bool(f), e[:200].replace("\n", " ")
    for src, pages in (("左联词典", DP), ("左联史", HP)):
        m = re.match(src + r"\s*第(\d+)页", c)
        if m:
            t = pages.get(int(m.group(1)))
            if t is None:
                return c, False, "[%s页索引未命中]" % src
            compact = t.replace(" ", "").replace("\n", "")
            f = hits_in(compact, names)
            if f:
                pos = compact.find(f)
                return c, True, compact[max(0, pos - 120):pos + 200]
            return c, False, compact[:150]
    m = re.match(r"左联回忆录\s*第(\d+)页", c)
    if m:
        t = MM.get(m.group(1), "")
        compact = (t or "").replace("\n", "")
        return c, bool(hits_in(compact, names)), compact[:200]
    return c, False, "[未解析定位]"

def screen(rid):
    r = RELS[rid]
    type_ = r["final_relation_type"] or r["original_relation_type"]
    sf, tf = forms(r["source_person_id"]), forms(r["target_person_id"])
    locs = [c.strip() for c in (r["evidence_ref"] or "").split(";") if c.strip()]
    results = []
    both_direct = None  # 日记条目同时或分别直接记录双方互动
    for c in locs:
        loc, ok, ex = locate(c, tf if "鲁迅日记" in c else tf + sf)
        results.append((loc, ok, ex))
        if "鲁迅日记" in c and ok and ("鲁迅日记" in c):
            both_direct = both_direct or (loc, ex)
    return r, type_, sf, tf, results, both_direct

def run(ids, out):
    lines = []
    for rid in ids:
        r, type_, sf, tf, results, dd = screen(rid)
        sp = PERSONS.get(r["source_person_id"], {}).get("standard_name")
        tp = PERSONS.get(r["target_person_id"], {}).get("standard_name")
        lines.append("## %s | %s - %s | %s" % (rid, sp, tp, type_))
        lines.append("ctx: " + (r["context"] or "").replace("\n", " ")[:200])
        for loc, ok, ex in results[:6]:
            lines.append(("  [HIT] " if ok else "  [   ] ") + loc + " ⇒ " + ex[:220])
        lines.append("")
    open(os.path.join(R, "work", out), "w", encoding="utf-8").write("\n".join(lines))
    print("wrote", out, len(ids))

if __name__ == "__main__":
    b = json.load(open(os.path.join(ROOT, "research", "briefs", "content-upgrade-overnight-2026-09-06", "baseline.json"), encoding="utf-8"))
    s = b["selection"]
    a, z = int(sys.argv[1]), int(sys.argv[2])
    run(s["relation_sample_ids"][a:z], "auto_%03d_%03d.md" % (a, z))
