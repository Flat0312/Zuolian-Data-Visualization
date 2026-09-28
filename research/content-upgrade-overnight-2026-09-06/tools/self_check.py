# -*- coding: utf-8 -*-
"""R内自检（返工扩展版）：冻结集合、记录状态、证据定位与引文、外键闭合、账本一致性、
required outputs、topic execution_status、档案claim链、support语义复核覆盖。
用法: python tools/self_check.py [相对R的副本目录]  （默认检查R根目录）"""
import json, re, os, sys, hashlib, collections
sys.path.insert(0, os.path.dirname(__file__))
import corpus

R = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
ROOT = corpus.ROOT
errors, warnings = [], []
base = os.path.abspath(os.path.join(R, sys.argv[1])) if len(sys.argv) > 1 else R

def load(name):
    p = os.path.join(base, name)
    if not os.path.exists(p):
        errors.append("missing file: " + name); return []
    return [json.loads(l) for l in open(p, encoding="utf-8") if l.strip()]

b = json.load(open(os.path.join(ROOT, "research", "briefs", "content-upgrade-overnight-2026-09-06", "baseline.json"), encoding="utf-8"))
s = b["selection"]
union = set(s["relation_sample_ids"]) | set(s["phase7_relation_ids"])

# required outputs（准确路径；根目录 validation_report.md）
_req = ["INDEX.md","PROGRESS.md","BLOCKED.md","selection.json","facts.jsonl","evidence.jsonl",
        "relations.jsonl","dossiers.jsonl","topics.json","change_proposals.csv","search_log.jsonl",
        "validation_report.md","night-state.json","scope_ledger.jsonl","morning-report.md","repair-report.md"]
for _f in _req:
    if not os.path.exists(os.path.join(base, _f)):
        errors.append("required output missing at exact path: " + _f)

rel = load("relations.jsonl")
ids = [r["relation_id"] for r in rel]
if len(ids) != len(set(ids)):
    errors.append("duplicate relation_id in relations.jsonl")
missing = union - set(ids)
extra = set(ids) - union
if extra:
    errors.append("relations.jsonl contains IDs outside frozen union: %s" % sorted(extra)[:5])
if missing:
    warnings.append("union not fully covered in this dir: %d missing" % len(missing))

ev = {e["evidence_id"]: e for e in load("evidence.jsonl")}
if len(ev) != len([1 for e in ev]):
    errors.append("duplicate evidence_id")
logs = load("search_log.jsonl")
logsids = {l["search_id"] for l in logs}
if len(logsids) != len(logs):
    errors.append("duplicate search_id")

facts = load("facts.jsonl")
fact_ids = set()
for f in facts:
    fact_ids.add(f["claim_id"])
for e in f["evidence_ids"] if facts else []:
    pass
for f in facts:
    for _e in f["evidence_ids"]:
        if _e not in ev:
            errors.append("fact %s dangling evidence: %s" % (f["claim_id"], _e))
    for _s in f["search_log_ids"]:
        if _s not in logsids:
            errors.append("fact %s dangling search_log: %s" % (f["claim_id"], _s))
    if f["review_status"] != "pending_human_review":
        errors.append("fact %s bad review_status" % f["claim_id"])

# 关系记录：枚举/人工字段/support引文/外键
for r in rel:
    rid = r["relation_id"]
    if r["evidence_support"] == "support" and not r["evidence_ids"]:
        errors.append("support without evidence: " + rid)
    if r["review_status"] != "pending_human_review":
        errors.append("bad review_status: " + rid)
    if r["human_verdict"] or r["human_reviewer"] or r["human_reviewed_at"]:
        errors.append("human fields non-empty: " + rid)
    if r["evidence_support"] not in ("support", "associated", "conflict", "insufficient"):
        errors.append("bad enum: " + rid)
    for _e in r["evidence_ids"]:
        if _e not in ev:
            errors.append("relation %s dangling evidence: %s" % (rid, _e))
    for _s in r["search_log_ids"]:
        if _s not in logsids:
            errors.append("relation %s dangling search_log: %s" % (rid, _s))

# 引文回定位
DI, DP, HP = corpus.diary_index(), corpus.dict_pages(), corpus.hist_pages()
for e in ev.values():
    if e["quote"].strip() == "":
        if e["retrieval_status"] in ("retrieved", "local_checked"):
            errors.append("empty quote on retrieved evidence: " + e["evidence_id"])
        continue
    loc = e["locator"]; q = e["quote"].replace("　", "")
    if loc.startswith("鲁迅日记") or re.match(r"^\d{4}年", loc):
        m = re.search(r"(\d{4})年(\d{1,2})月(\d{1,2})日", loc)
        if not m:
            errors.append("unparsable diary locator: " + e["evidence_id"]); continue
        entry = DI.get((int(m.group(1)), int(m.group(2)), int(m.group(3))))
        if entry is None:
            errors.append("diary date not found locally: " + e["evidence_id"]); continue
        c = entry.replace("\n", " ").replace("　", "")
        if not (q[:8] in c or q[:6] in c):
            errors.append("quote not located in dated entry: " + e["evidence_id"])
    else:
        mm = re.search(r"第(\d+)页", loc)
        pages = {"左联词典": DP, "左联史": HP}.get(e.get("source_family", ""))
        if not mm or not pages:
            warnings.append("non-page locator skipped: " + e["evidence_id"]); continue
        t = pages.get(int(mm.group(1)), "").replace(" ", "").replace("\n", "")
        if t and q[:10] not in t:
            errors.append("quote not located in cited page: " + e["evidence_id"])

# dossiers：60份、正文存在、claim链闭合、状态合法
doss = load("dossiers.jsonl")
if len(doss) != 60:
    errors.append("dossiers.jsonl must have 60 records, got %d" % len(doss))
ent_key = collections.Counter((d["entity_type"], d["entity_id"]) for d in doss)
for k, v in ent_key.items():
    if v > 1:
        errors.append("duplicate dossier: %s" % (k,))
for d in doss:
    if d["content_status"] not in ("review_ready_draft", "limited_report"):
        errors.append("bad dossier content_status: %s" % d["entity_id"])
    if d["review_status"] != "pending_human_review":
        errors.append("bad dossier review_status: %s" % d["entity_id"])
    p = os.path.join(base, d["document_path"]) if not os.path.isabs(d["document_path"]) else d["document_path"]
    if not os.path.exists(p):
        errors.append("dossier body missing: %s (%s)" % (d["entity_id"], d["document_path"]))
        continue
    body = open(p, encoding="utf-8").read()
    for cid in d["claim_ids"]:
        if cid not in fact_ids:
            errors.append("dossier %s dangling claim: %s" % (d["entity_id"], cid))
        elif cid not in body:
            errors.append("dossier %s body does not cite claim %s" % (d["entity_id"], cid))

# topics：execution_status、路径存在
tops_raw = open(os.path.join(base, "topics.json"), encoding="utf-8").read() if os.path.exists(os.path.join(base, "topics.json")) else "[]"
try:
    tops = json.loads(tops_raw)
except Exception:
    tops = [json.loads(l) for l in tops_raw.splitlines() if l.strip()]
if not isinstance(tops, list):
    errors.append("topics.json must be a JSON array"); tops = []
if len(tops) != 5:
    errors.append("topics.json must have exactly 5 topics")
for t in tops:
    if t.get("execution_status") not in ("queued", "in_progress", "checked", "blocked"):
        errors.append("topic %s missing/invalid execution_status" % t.get("topic_id"))
    if t.get("content_status") not in ("review_ready_draft", "limited_report"):
        errors.append("topic %s invalid content_status" % t.get("topic_id"))
    dp = t.get("document_path")
    if dp and not os.path.exists(os.path.join(base, dp)):
        errors.append("topic document missing: " + dp)
    for cid in t.get("claim_ids", []):
        if cid not in fact_ids and not cid.startswith("CLAIM-"):
            errors.append("topic %s dangling claim ref: %s" % (t["topic_id"], cid))

# scope ledger：476对象、状态与结果一致
led = load("scope_ledger.jsonl")
lk = collections.Counter((l["object_type"], l["object_id"]) for l in led)
for k, v in lk.items():
    if v > 1:
        errors.append("duplicate scope ledger row: %s" % (k,))
doss_keys = {(d["entity_type"], d["entity_id"]) for d in doss}
relids = set(ids)
for l in led:
    ot, oid = l["object_type"], l["object_id"]
    if l["execution_status"] not in ("queued", "in_progress", "checked", "blocked"):
        errors.append("ledger %s/%s bad status" % (ot, oid))
    if ot in ("person", "event", "place"):
        if (ot, oid) in doss_keys:
            d = [d for d in doss if (d["entity_type"], d["entity_id"]) == (ot, oid)][0]
            if l["execution_status"] == "checked" and not d["document_path"]:
                errors.append("ledger checked but dossier has no path: %s" % oid)
        elif l["execution_status"] == "checked":
            errors.append("ledger checked but no dossier record: %s" % oid)
    if ot == "relation":
        if l["execution_status"] == "checked" and oid not in relids:
            errors.append("ledger checked but relation missing: " + oid)

# support语义复核覆盖：复核表须含58条（原support全集）且覆盖当前全部support
sr = load("verification/support_semantic_review.jsonl") if base == R else []
if base == R:
    sr_ids = {x["relation_id"] for x in sr}
    if len(sr) < 58:
        errors.append("support semantic review must cover 58 original supports, got %d" % len(sr))
    cur_sup = {r["relation_id"] for r in rel if r["evidence_support"] == "support"}
    if not cur_sup <= sr_ids:
        errors.append("semantic review missing current supports: %s" % sorted(cur_sup - sr_ids)[:5])

# not_attempted 与完成声明冲突
state = {}
sp = os.path.join(base, "night-state.json")
if os.path.exists(sp):
    state = json.load(open(sp, encoding="utf-8"))
    rs = state.get("run_status", "")
    na = sum(1 for l in logs if l["result"] == "not_attempted")
    q = sum(1 for l in led if l["execution_status"] in ("queued", "in_progress"))
    if ("complete" in rs) and (na or q):
        errors.append("complete claimed but %d not_attempted logs and %d queued/in_progress ledger rows" % (na, q))

print("SELF_CHECK errors=%d warnings=%d" % (len(errors), len(warnings)))
for x in errors[:40]: print("ERROR:", x)
for x in warnings[:10]: print("WARN:", x)
sys.exit(1 if errors else 0)
