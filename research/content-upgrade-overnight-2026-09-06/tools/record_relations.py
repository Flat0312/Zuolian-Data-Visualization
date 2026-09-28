# -*- coding: utf-8 -*-
"""记录关系预审结论。输入: verdicts file (每行: rid|verdict|quote_key|reason|note)
verdict in support/associated/conflict/insufficient
quote_key: 用于复现证据摘录的定位串（原样记录evidence_ref里的定位）
reason: 判定理由正文。note: 附加说明可空
仅人工(AI执行者)复核过定位摘录的关系才可判support。
生成 relations.jsonl / evidence.jsonl / search_log.jsonl 增量记录。"""
import csv, json, re, os, sys, hashlib, datetime
sys.path.insert(0, os.path.dirname(__file__))
import corpus

R = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
ROOT = corpus.ROOT
NOW = "2026-09-06"  # 实际运行日期；具体时刻由调用方传入或默认批次时间

PERSONS = {r["person_id"]: r for r in csv.DictReader(open(os.path.join(ROOT, "data", "processed", "persons.csv"), encoding="utf-8-sig"))}
RELS = {r["relation_id"]: r for r in csv.DictReader(open(os.path.join(ROOT, "data", "processed", "person_relations.csv"), encoding="utf-8-sig"))}
DP, HP, DI, MM = corpus.dict_pages(), corpus.hist_pages(), corpus.diary_index(), corpus.memoir()

LOCAL_PATHS = {"鲁迅日记": corpus.DIARY, "左联词典": corpus.DICT, "左联史": corpus.HIST}
FAMILY_OF = {"鲁迅日记": "鲁迅日记（公开转录，本地逐字核对）", "左联词典": "左联词典（OCR）", "左联史": "左联史（OCR）", "左联回忆录": "左联回忆录（OCR节选）"}

def quote_for(locator):
    """按定位串回原文提取引文（连续短摘录）。返回 (source_title, quote, hash_basis)"""
    m = re.match(r"鲁迅日记\s*(\d{4})年(\d{1,2})月(\d{1,2})日", locator)
    if m:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
        e = DI.get((y, mo, d))
        if e is None:
            return None
        return ("鲁迅日记 %d年" % y, e.replace("\n", " ")[:160], "local_file_bytes(OCR全文,引文为其中连续片段)")
    for src, pages in (("左联词典", DP), ("左联史", HP)):
        m = re.match(src + r"\s*第(\d+)页", locator)
        if m:
            t = pages.get(int(m.group(1)))
            if t is None:
                return None
            compact = t.replace(" ", "").replace("\n", "")
            focus = quote_for.focus
            pos = -1
            for n in focus or []:
                p = compact.find(n)
                if p >= 0:
                    pos = p; break
            if pos >= 0:
                frag = compact[max(0, pos - 60):pos + 140]
            else:
                frag = compact[:160]
            return (src, frag, "local_file_bytes(OCR全文,引文为其中连续片段)")
    m = re.match(r"左联回忆录\s*第(\d+)页", locator)
    if m and MM.get(m.group(1)):
        return ("左联回忆录", MM[m.group(1)].replace("\n", "")[:160], "local_file_bytes(OCR全文,引文为其中连续片段)")
    return None

def evidence_id_for(locator):
    h = hashlib.sha256(locator.encode("utf-8")).hexdigest()[:12].upper()
    return "EVI-N1-%s" % h

quote_for.focus = []

def name_forms_for(r):
    out = []
    for pid in (r["source_person_id"], r["target_person_id"]):
        p = PERSONS.get(pid, {})
        n = p.get("standard_name", "")
        if n:
            out.append(n)
            if len(n) >= 3:
                out.append(n[1:])
    return out

def main(verdict_file, ts):
    rel_out, ev_out, log_out = [], [], []
    seen_locators = {}
    for line in open(verdict_file, encoding="utf-8"):
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        parts = line.split("|")
        rid, verdict, locator, reason = parts[0].strip(), parts[1].strip(), parts[2].strip(), parts[3].strip()
        note = parts[4].strip() if len(parts) > 4 else ""
        r = RELS[rid]
        proposed = parts[5].strip() if len(parts) > 5 and parts[5].strip() else (r["final_relation_type"] or r["original_relation_type"])
        grp = json.load(open(os.path.join(ROOT, "research", "briefs", "content-upgrade-overnight-2026-09-06", "baseline.json"), encoding="utf-8"))["selection"]
        groups = [g for g, c in (("phase7", rid in grp["phase7_relation_ids"]), ("sample400", rid in grp["relation_sample_ids"])) if c]
        ev_ids, log_ids = [], []
        if locator and locator != "-":
            if locator not in seen_locators:
                q = quote_for(locator)
                if q is None:
                    print("!! quote failed for", rid, locator)
                    continue
                src_title, quote, hash_basis = q
                eid = evidence_id_for(locator)
                fam = next((f for k, f in FAMILY_OF.items() if locator.startswith(k)), "unknown")
                seen_locators[locator] = eid
                ev_out.append({
                    "evidence_id": eid, "source_title": src_title, "author_or_institution": "unknown" if src_title != "鲁迅日记" else "鲁迅",
                    "version": "OCR本地全文（项目既有raw_texts/runtime_sources，未复制）",
                    "source_family": fam.split("（")[0], "source_level": "B（公开转录或OCR工具书；本地逐字核对）",
                    "source_type": "diary" if src_title.startswith("鲁迅日记") else "reference_work",
                    "source_url": "", "local_path": LOCAL_PATHS.get(fam.split("（")[0], "").replace("\\", "/"),
                    "accessed_at": NOW + "T" + ts, "locator": locator, "quote": quote,
                    "locator_context": locator, "content_hash": corpus.sha256_file(LOCAL_PATHS.get(fam.split("（")[0], corpus.DIARY)) if fam.split("（")[0] in LOCAL_PATHS else "",
                    "hash_basis": hash_basis, "quote_sha256": hashlib.sha256(quote.encode("utf-8")).hexdigest(),
                    "retrieval_status": "local_checked", "review_status": "pending_human_review"})
            ev_ids.append(seen_locators[locator])
            log_out.append({"search_id": "SRCH-%s" % rid, "claim_or_relation_ids": [rid], "round": 1,
                            "query_or_local_probe": "本地原文核对：" + locator, "source_url_or_path": locator,
                            "attempted_at": NOW + "T" + ts, "result": "local_checked", "failure_detail": "", "receipt_path": ""})
        rel_out.append({
            "relation_id": rid, "groups": groups,
            "source_person_id": r["source_person_id"], "target_person_id": r["target_person_id"],
            "original_relation_type": r["final_relation_type"] or r["original_relation_type"],
            "proposed_relation_type": proposed,
            "evidence_support": verdict, "reason": reason + (("；" + note) if note else ""),
            "evidence_ids": ev_ids, "search_log_ids": [l["search_id"] for l in log_out if rid in l["claim_or_relation_ids"]],
            "review_status": "pending_human_review", "ai_executor": "ZCode GLM-5.3-Flash (AI executor)",
            "checked_at": NOW + "T" + ts, "human_verdict": "", "human_reviewer": "", "human_reviewed_at": "",
            "execution_status": "checked"})
        quote_for.focus = name_forms_for(r)
    with open(os.path.join(R, "relations.jsonl"), "a", encoding="utf-8") as f:
        for x in rel_out:
            f.write(json.dumps(x, ensure_ascii=False) + "\n")
    with open(os.path.join(R, "evidence.jsonl"), "a", encoding="utf-8") as f:
        for x in ev_out:
            f.write(json.dumps(x, ensure_ascii=False) + "\n")
    with open(os.path.join(R, "search_log.jsonl"), "a", encoding="utf-8") as f:
        for x in log_out:
            f.write(json.dumps(x, ensure_ascii=False) + "\n")
    print("recorded", len(rel_out), "relations;", len(ev_out), "new evidence")

if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else "02:00:00")
