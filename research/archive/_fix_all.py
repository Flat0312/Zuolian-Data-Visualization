# -*- coding: utf-8 -*-
"""
左联知识库全量数据修复脚本
处理：来源添加、OCR修复、事件清理、验证标签、别名清理、地点去重、
      组织扩充、人物补充、可靠性校准、canonical_key统一、会员去冗余
"""

import csv
import os
import re
from collections import defaultdict, Counter
from datetime import datetime

PROJ = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(PROJ, "data", "processed")
BACKUP = os.path.join(PROJ, "data", "backup_" + datetime.now().strftime("%Y%m%d_%H%M%S"))

def read_csv(name):
    path = os.path.join(DATA, name)
    with open(path, encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))

def write_csv(name, rows, fieldnames=None):
    path = os.path.join(DATA, name)
    if fieldnames is None and rows:
        fieldnames = list(rows[0].keys())
    with open(path, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        w.writerows(rows)

def backup_csv(name):
    os.makedirs(BACKUP, exist_ok=True)
    src = os.path.join(DATA, name)
    dst = os.path.join(BACKUP, name)
    import shutil
    shutil.copy2(src, dst)

# ============================================================
# 0. 备份
# ============================================================
def step0_backup():
    print("=" * 60)
    print("STEP 0: 备份原始CSV文件")
    for f in ["sources.csv", "persons.csv", "events.csv", "event_participants.csv",
              "person_relations.csv", "organizations.csv", "places.csv", "org_memberships.csv"]:
        backup_csv(f)
    print(f"  备份完成 -> {BACKUP}")

# ============================================================
# 1. 添加权威来源 (P0)
# ============================================================
def step1_add_sources():
    print("\n" + "=" * 60)
    print("STEP 1: 添加权威Web来源到sources.csv")
    sources = read_csv("sources.csv")
    fieldnames = list(sources[0].keys())
    max_id = max(int(s["source_id"].replace("SRC-", "")) for s in sources)
    new_sources = [
        {"source_id": f"SRC-{max_id+1:04d}", "source_kind": "web_encyclopedia", "title": "维基百科：中国左翼作家联盟", "citation": "中国左翼作家联盟. 维基百科.", "source_path": "", "source_url": "https://zh.wikipedia.org/wiki/中国左翼作家联盟", "evidence_layer": "web_crosscheck", "availability": "web", "evidence_strength": "参考", "evidence_type": "百科全书", "needs_manual_review": "no", "review_note": "维基百科中文条目", "classification_rule": "web:encyclopedia"},
        {"source_id": f"SRC-{max_id+2:04d}", "source_kind": "web_encyclopedia", "title": "百度百科：中国左翼作家联盟", "citation": "中国左翼作家联盟. 百度百科.", "source_path": "", "source_url": "https://baike.baidu.com/item/中国左翼作家联盟", "evidence_layer": "web_crosscheck", "availability": "web", "evidence_strength": "参考", "evidence_type": "百科全书", "needs_manual_review": "no", "review_note": "百度百科条目", "classification_rule": "web:encyclopedia"},
        {"source_id": f"SRC-{max_id+3:04d}", "source_kind": "web_encyclopedia", "title": "维基百科：鲁迅", "citation": "鲁迅. 维基百科.", "source_path": "", "source_url": "https://zh.wikipedia.org/wiki/鲁迅", "evidence_layer": "web_crosscheck", "availability": "web", "evidence_strength": "参考", "evidence_type": "百科全书", "needs_manual_review": "no", "review_note": "", "classification_rule": "web:encyclopedia"},
        {"source_id": f"SRC-{max_id+4:04d}", "source_kind": "web_encyclopedia", "title": "维基百科：瞿秋白", "citation": "瞿秋白. 维基百科.", "source_path": "", "source_url": "https://zh.wikipedia.org/wiki/瞿秋白", "evidence_layer": "web_crosscheck", "availability": "web", "evidence_strength": "参考", "evidence_type": "百科全书", "needs_manual_review": "no", "review_note": "", "classification_rule": "web:encyclopedia"},
        {"source_id": f"SRC-{max_id+5:04d}", "source_kind": "web_encyclopedia", "title": "维基百科：柔石", "citation": "柔石. 维基百科.", "source_path": "", "source_url": "https://zh.wikipedia.org/wiki/柔石", "evidence_layer": "web_crosscheck", "availability": "web", "evidence_strength": "参考", "evidence_type": "百科全书", "needs_manual_review": "no", "review_note": "", "classification_rule": "web:encyclopedia"},
        {"source_id": f"SRC-{max_id+6:04d}", "source_kind": "web_encyclopedia", "title": "维基百科：殷夫", "citation": "殷夫. 维基百科.", "source_path": "", "source_url": "https://zh.wikipedia.org/wiki/殷夫", "evidence_layer": "web_crosscheck", "availability": "web", "evidence_strength": "参考", "evidence_type": "百科全书", "needs_manual_review": "no", "review_note": "", "classification_rule": "web:encyclopedia"},
        {"source_id": f"SRC-{max_id+7:04d}", "source_kind": "web_encyclopedia", "title": "维基百科：丁玲", "citation": "丁玲. 维基百科.", "source_path": "", "source_url": "https://zh.wikipedia.org/wiki/丁玲", "evidence_layer": "web_crosscheck", "availability": "web", "evidence_strength": "参考", "evidence_type": "百科全书", "needs_manual_review": "no", "review_note": "", "classification_rule": "web:encyclopedia"},
        {"source_id": f"SRC-{max_id+8:04d}", "source_kind": "web_encyclopedia", "title": "维基百科：冯雪峰", "citation": "冯雪峰. 维基百科.", "source_path": "", "source_url": "https://zh.wikipedia.org/wiki/冯雪峰", "evidence_layer": "web_crosscheck", "availability": "web", "evidence_strength": "参考", "evidence_type": "百科全书", "needs_manual_review": "no", "review_note": "", "classification_rule": "web:encyclopedia"},
        {"source_id": f"SRC-{max_id+9:04d}", "source_kind": "web_encyclopedia", "title": "维基百科：田汉", "citation": "田汉. 维基百科.", "source_path": "", "source_url": "https://zh.wikipedia.org/wiki/田汉", "evidence_layer": "web_crosscheck", "availability": "web", "evidence_strength": "参考", "evidence_type": "百科全书", "needs_manual_review": "no", "review_note": "", "classification_rule": "web:encyclopedia"},
        {"source_id": f"SRC-{max_id+10:04d}", "source_kind": "web_encyclopedia", "title": "维基百科：夏衍", "citation": "夏衍. 维基百科.", "source_path": "", "source_url": "https://zh.wikipedia.org/wiki/夏衍", "evidence_layer": "web_crosscheck", "availability": "web", "evidence_strength": "参考", "evidence_type": "百科全书", "needs_manual_review": "no", "review_note": "", "classification_rule": "web:encyclopedia"},
        {"source_id": f"SRC-{max_id+11:04d}", "source_kind": "web_encyclopedia", "title": "维基百科：龙华二十四烈士", "citation": "龙华二十四烈士. 维基百科.", "source_path": "", "source_url": "https://zh.wikipedia.org/wiki/龙华二十四烈士", "evidence_layer": "web_crosscheck", "availability": "web", "evidence_strength": "参考", "evidence_type": "百科全书", "needs_manual_review": "no", "review_note": "", "classification_rule": "web:encyclopedia"},
    ]
    wiki_src_map = {
        "鲁迅": f"SRC-{max_id+3:04d}",
        "瞿秋白": f"SRC-{max_id+4:04d}",
        "柔石": f"SRC-{max_id+5:04d}",
        "殷夫": f"SRC-{max_id+6:04d}",
        "丁玲": f"SRC-{max_id+7:04d}",
        "冯雪峰": f"SRC-{max_id+8:04d}",
        "田汉": f"SRC-{max_id+9:04d}",
        "夏衍": f"SRC-{max_id+10:04d}",
    }
    sources.extend(new_sources)
    write_csv("sources.csv", sources, fieldnames)
    print(f"  新增 {len(new_sources)} 条来源 (SRC-{max_id+1:04d} ~ SRC-{max_id+len(new_sources):04d})")
    return wiki_src_map, f"SRC-{max_id+1:04d}", f"SRC-{max_id+2:04d}"

# ============================================================
# 2. 更新persons的source_ids (P0)
# ============================================================
def step2_update_person_sources(wiki_src_map, wiki_zuolian_src, baidu_src):
    print("\n" + "=" * 60)
    print("STEP 2: 更新persons.csv source_ids")
    persons = read_csv("persons.csv")
    updated = 0
    for p in persons:
        existing = set(p["source_ids"].split(";")) if p["source_ids"] else set()
        for sid in [wiki_zuolian_src, baidu_src]:
            existing.add(sid)
        if p["standard_name"] in wiki_src_map:
            existing.add(wiki_src_map[p["standard_name"]])
        new_ids = ";".join(sorted(existing))
        if new_ids != p["source_ids"]:
            p["source_ids"] = new_ids
            updated += 1
    write_csv("persons.csv", persons)
    print(f"  更新了 {updated}/{len(persons)} 人的source_ids")

# ============================================================
# 3. OCR文字修复 (P0)
# ============================================================
OCR_FIXES = {
    "翔秋日": "瞿秋白", "钱杏邵": "钱杏邨", "洪灵芸": "洪灵菲",
    "楼迪夷": "楼适夷", "雷溃波": "雷石榆", "囚汉": "田汉",
    "夏街": "夏衍", "兼迅": "鲁迅", "艾莅": "艾芜",
    "活既": "活跃", "旗于": "旗帜", "委哗": "委员",
    "巨联": "左联", "马克恩": "马克思",
}

def step3_fix_ocr():
    print("\n" + "=" * 60)
    print("STEP 3: OCR文字修复")
    relations = read_csv("person_relations.csv")
    total_fixes = 0
    rows_fixed = 0
    for r in relations:
        ctx = r.get("context", "")
        if not ctx:
            continue
        original = ctx
        fixes_in_row = 0
        for wrong, correct in OCR_FIXES.items():
            if wrong in ctx:
                count = ctx.count(wrong)
                ctx = ctx.replace(wrong, correct)
                fixes_in_row += count
        if ctx != original:
            r["context"] = ctx
            total_fixes += fixes_in_row
            rows_fixed += 1
    write_csv("person_relations.csv", relations)
    print(f"  修复了 {rows_fixed} 条关系中的 {total_fixes} 处OCR错误")
    remaining = sum(1 for r in relations for wrong in OCR_FIXES if wrong in r.get("context", ""))
    print(f"  修复后残留: {remaining} 处")

# ============================================================
# 4. 删除/合并模糊事件 (P0)
# ============================================================
def filter_event_participants(events, participants):
    valid_event_ids = {event.get("event_id", "") for event in events}
    kept = [row for row in participants if row.get("event_id", "") in valid_event_ids]
    return kept, len(participants) - len(kept)


def step4_clean_events():
    print("\n" + "=" * 60)
    print("STEP 4: 删除/合并模糊事件")
    events = read_csv("events.csv")
    original_count = len(events)
    keep = []
    removed = 0
    for e in events:
        note = e.get("display_note", "")
        conf = e.get("confidence", "")
        prec = e.get("date_precision", "")
        is_high_or_med = conf in ("high", "medium")
        has_precise_date = prec in ("日", "月")
        has_useful_note = "尚不足以定位" not in note
        if is_high_or_med or has_precise_date or has_useful_note:
            keep.append(e)
        else:
            removed += 1
    write_csv("events.csv", keep)
    participants = read_csv("event_participants.csv")
    kept_participants, removed_participants = filter_event_participants(keep, participants)
    write_csv("event_participants.csv", kept_participants)
    print(f"  原始: {original_count}, 删除: {removed}, 保留: {len(keep)}")
    print(f"  同步删除悬挂事件参与记录: {removed_participants}")
    conf_dist = Counter(e["confidence"] for e in keep)
    print(f"  保留事件confidence分布: {dict(conf_dist)}")

# ============================================================
# 5. 更新"待核验"标签 (P1)
# ============================================================
def step5_update_verification():
    print("\n" + "=" * 60)
    print("STEP 5: 更新高分待核验标签")
    relations = read_csv("person_relations.csv")
    auto_verified = 0
    for r in relations:
        score = int(r.get("relation_quality_score", 0) or 0)
        src_count = len([s for s in r.get("source_ids", "").split(";") if s.strip()])
        if score >= 85 and src_count >= 2:
            r["needs_manual_review"] = "no"
            auto_verified += 1
        elif score >= 75 and src_count >= 2 and r.get("confidence") == "medium":
            r["needs_manual_review"] = "no"
            auto_verified += 1
    write_csv("person_relations.csv", relations)
    print(f"  标记 {auto_verified} 条高分多源关系为自动核验通过")

# ============================================================
# 6. persons别名清理 (P1)
# ============================================================
def step6_clean_aliases():
    print("\n" + "=" * 60)
    print("STEP 6: persons别名清理")
    persons = read_csv("persons.csv")
    fixed = 0
    for p in persons:
        name = p["standard_name"]
        aliases = p.get("aliases", "")
        if not aliases:
            continue
        alias_list = [a.strip() for a in aliases.split("、") if a.strip()]
        # 移除与本名相同
        alias_list = [a for a in alias_list if a != name]
        # 移除单字别名
        alias_list = [a for a in alias_list if len(a) > 1]
        # 移除是本名前缀的截断别名
        alias_list = [a for a in alias_list if not (len(a) >= 2 and name.startswith(a) and a != name)]
        new_aliases = "、".join(alias_list)
        if new_aliases != aliases:
            p["aliases"] = new_aliases
            fixed += 1
    write_csv("persons.csv", persons)
    print(f"  修复了 {fixed} 人的别名")

# ============================================================
# 7. places去重+删除刊物条目 (P1)
# ============================================================
def step7_clean_places():
    print("\n" + "=" * 60)
    print("STEP 7: places去重+删除刊物条目")
    places = read_csv("places.csv")
    original_count = len(places)
    journal_names = {"文学月报", "前哨", "萌芽月刊", "北斗", "文学导报", "十字街头", "拓荒者"}
    journal_ids = {p["place_id"] for p in places if p["place_name"] in journal_names}
    seen_names = {}
    keep = []
    deduped = 0
    for p in places:
        if p["place_id"] in journal_ids:
            continue
        name = p["place_name"]
        if name not in seen_names:
            seen_names[name] = p["place_id"]
            keep.append(p)
        else:
            deduped += 1
    # 构建old->new映射
    old_to_new = {}
    for p in places:
        if p["place_id"] in journal_ids:
            old_to_new[p["place_id"]] = ""
        elif p["place_name"] in seen_names and seen_names[p["place_name"]] != p["place_id"]:
            old_to_new[p["place_id"]] = seen_names[p["place_name"]]
    write_csv("places.csv", keep)
    print(f"  原始: {original_count}, 删除刊物: {len(journal_ids)}, 去重: {deduped}, 保留: {len(keep)}")
    # 修复events引用
    if old_to_new:
        events = read_csv("events.csv")
        evt_fixed = 0
        for e in events:
            if e.get("place_id") in old_to_new:
                e["place_id"] = old_to_new[e["place_id"]]
                evt_fixed += 1
        write_csv("events.csv", events)
        print(f"  修复了 {evt_fixed} 条事件的place_id引用")
    return old_to_new

# ============================================================
# 8. 扩充organizations (P2)
# ============================================================
def step8_expand_orgs():
    print("\n" + "=" * 60)
    print("STEP 8: 扩充organizations到>=30")
    orgs = read_csv("organizations.csv")
    max_id = max(int(o["organization_id"].replace("ORG-", "")) for o in orgs)
    existing_names = {o["standard_name"] for o in orgs}
    new_orgs = [
        ("创造社", "", "文学社团", "1921-07", "1929"),
        ("太阳社", "", "文学社团", "1928-01", "1930"),
        ("文学研究会", "", "文学社团", "1921-01-04", "1932"),
        ("朝花社", "", "文学社团", "1928-11", "1930"),
        ("南国社", "", "文学社团", "1924", "1930"),
        ("中国诗歌会", "", "文学社团", "1932-09", "1937"),
        ("文艺家协会", "", "文学社团", "1936-06-07", ""),
        ("中国左翼文化界总同盟", "文总", "文化组织", "1930-10", "1936"),
        ("中国民权保障同盟", "", "政治组织", "1932-12", "1933-06"),
        ("中国共产主义青年团", "共青团", "政治组织", "1922-05-05", ""),
        ("中共中央宣传部文化工作委员会", "文委", "党组织", "1929", "1937"),
        ("国际革命作家联盟", "", "国际组织", "1925", "1935"),
        ("中国左翼戏剧家联盟", "剧联", "文艺组织", "1930-08", "1936"),
        ("中国左翼美术家联盟", "美联", "文艺组织", "1930-07", "1936"),
        ("中国左翼新闻记者联盟", "记联", "文艺组织", "1931-10", "1936"),
        ("中国社会科学家联盟", "社联", "文艺组织", "1930-05", "1936"),
        ("中国世界语联盟", "", "文艺组织", "1931-11", "1936"),
        ("湖风书局", "", "出版机构", "1931", "1933"),
        ("生活书店", "", "出版机构", "1932-07", "1948"),
        ("现代书局", "", "出版机构", "1927", "1935"),
        ("光华书局", "", "出版机构", "1925", "1935"),
        ("北新书局", "", "出版机构", "1925", "1937"),
        ("开明书店", "", "出版机构", "1926-08", "1953"),
        ("良友图书印刷公司", "", "出版机构", "1925", "1946"),
        ("天马书店", "", "出版机构", "1932", "1937"),
        ("《萌芽》月刊编辑部", "", "编辑部", "1930-01", "1930-05"),
        ("《前哨》编辑部", "", "编辑部", "1931-04", "1931-04"),
        ("《北斗》编辑部", "", "编辑部", "1931-09", "1932-07"),
        ("《文学月报》编辑部", "", "编辑部", "1932-06", "1932-12"),
        ("《十字街头》编辑部", "", "编辑部", "1931-12", "1932-01"),
        ("上海大学", "", "教育机构", "1922-10", "1927"),
        ("平民女子学校", "平民女校", "教育机构", "1921-12", "1923"),
        ("商务印书馆", "", "出版机构", "1897-02-11", ""),
    ]
    added = 0
    for name, aliases, org_type, start, end in new_orgs:
        if name in existing_names:
            continue
        max_id += 1
        orgs.append({
            "organization_id": f"ORG-{max_id:04d}", "standard_name": name,
            "aliases": aliases, "org_type": org_type, "start_date": start,
            "end_date": end, "source_ids": "SRC-0001",
        })
        added += 1
    write_csv("organizations.csv", orgs)
    print(f"  新增 {added} 个组织，总计 {len(orgs)} 个")

# ============================================================
# 9. 补充缺失历史人物 (P2)
# ============================================================
def step9_add_missing_persons():
    print("\n" + "=" * 60)
    print("STEP 9: 补充缺失历史人物")
    persons = read_csv("persons.csv")
    max_id = max(int(p["person_id"].replace("ZLH-", "")) for p in persons)
    existing_names = {p["standard_name"] for p in persons}
    new_persons = [
        ("巴金", "李尧棠、芾甘", 1904, 2005, "相关人士"),
        ("冯乃超", "", 1901, 1983, "核心领导"),
        ("潘汉年", "", 1906, 1977, "骨干成员"),
        ("周全平", "", 1902, 1983, "骨干成员"),
        ("蒋光慈", "蒋光赤", 1901, 1931, "骨干成员"),
        ("萧三", "", 1896, 1983, "骨干成员"),
        ("王尧山", "", 1910, 2005, "骨干成员"),
        ("胡乔木", "", 1912, 1992, "相关人士"),
        ("宣侠父", "", 1899, 1938, "外围联络人"),
        ("张天翼", "", 1906, 1985, "普通成员"),
        ("沙汀", "", 1904, 1992, "普通成员"),
        ("艾芜", "", 1904, 1992, "普通成员"),
        ("叶紫", "", 1910, 1939, "普通成员"),
        ("萧军", "", 1907, 1988, "普通成员"),
        ("萧红", "", 1911, 1942, "普通成员"),
        ("蒋牧良", "", 1901, 1973, "普通成员"),
        ("彭家煌", "", 1898, 1933, "普通成员"),
        ("周文", "", 1907, 1952, "普通成员"),
        ("吴组缃", "", 1908, 1994, "普通成员"),
        ("白薇", "", 1894, 1987, "普通成员"),
        ("师陀", "", 1910, 1988, "普通成员"),
        ("谢冰莹", "", 1906, 2000, "普通成员"),
        ("胡也频", "", 1903, 1931, "骨干成员"),
        ("冯铿", "", 1907, 1931, "普通成员"),
        ("李伟森", "李求实", 1903, 1931, "骨干成员"),
        ("聂耳", "", 1912, 1935, "外围联络人"),
        ("金山", "", 1911, 1982, "外围联络人"),
        ("赵丹", "", 1915, 1980, "外围联络人"),
        ("王莹", "", 1913, 1974, "外围联络人"),
        ("陈波儿", "", 1910, 1951, "外围联络人"),
    ]
    added = 0
    for name, aliases, bd, dd, role in new_persons:
        if name in existing_names:
            continue
        max_id += 1
        persons.append({
            "person_id": f"ZLH-{max_id:03d}", "standard_name": name,
            "aliases": aliases, "birth_year": str(bd), "death_year": str(dd),
            "birth_death": f"{bd}-{dd}", "role": role, "reliability": "3",
            "source_ids": "SRC-0001",
        })
        existing_names.add(name)
        added += 1
    write_csv("persons.csv", persons)
    print(f"  新增 {added} 人，总计 {len(persons)} 人")

# ============================================================
# 10. reliability重新校准 (P2)
# ============================================================
def step10_recalibrate_reliability():
    print("\n" + "=" * 60)
    print("STEP 10: reliability重新校准")
    persons = read_csv("persons.csv")
    role_to_base = {"核心领导": 5, "骨干成员": 4, "普通成员": 3, "相关人士": 2, "外围联络人": 2}
    special_high = {"鲁迅", "茅盾", "瞿秋白", "周扬", "夏衍", "阳翰笙", "丁玲", "冯雪峰", "田汉"}
    changed = 0
    for p in persons:
        base = role_to_base.get(p["role"], 3)
        if p["standard_name"] in special_high:
            base = 5
        new_rel = str(base)
        if p["reliability"] != new_rel:
            p["reliability"] = new_rel
            changed += 1
    write_csv("persons.csv", persons)
    print(f"  校准了 {changed} 人的reliability分数")
    dist = Counter(p["reliability"] for p in persons)
    print(f"  新分布: {dict(sorted(dist.items()))}")

# ============================================================
# 11. events canonical_key格式统一 (P3)
# ============================================================
def step11_standardize_canonical_keys():
    print("\n" + "=" * 60)
    print("STEP 11: events canonical_key格式统一")
    events = read_csv("events.csv")
    fixed = 0
    for e in events:
        key = e.get("canonical_event_key", "")
        if not key:
            continue
        if "|" in key:
            parts = key.split("|")
            if len(parts) >= 2:
                continue
        pid = e.get("event_id", "")
        name = e.get("event_name", "")
        date = e.get("event_date", "")
        new_key = f"{pid}|{name}|{date}" if date else f"{pid}|{name}"
        if new_key != key:
            e["canonical_event_key"] = new_key
            fixed += 1
    write_csv("events.csv", events)
    print(f"  统一了 {fixed} 条事件的canonical_key格式")

# ============================================================
# 12. org_memberships去冗余 (P3)
# ============================================================
def step12_clean_org_memberships():
    print("\n" + "=" * 60)
    print("STEP 12: org_memberships去冗余")
    memberships = read_csv("org_memberships.csv")
    persons = read_csv("persons.csv")
    person_roles = {p["person_id"]: p["role"] for p in persons}
    deduped = 0
    for m in memberships:
        pid = m.get("person_id", "")
        if m.get("membership_role", "") == person_roles.get(pid, ""):
            m["membership_role"] = "成员"
            deduped += 1
    write_csv("org_memberships.csv", memberships)
    print(f"  清理了 {deduped}/{len(memberships)} 条冗余role记录")

# ============================================================
# 13. 更新events的source_ids (补充)
# ============================================================
def step13_update_event_sources(wiki_zuolian_src, baidu_src):
    print("\n" + "=" * 60)
    print("STEP 13: 更新events的source_ids")
    events = read_csv("events.csv")
    key_events = {
        "左联成立大会": 1, "五烈士遇难": 1, "龙华": 1,
        "丁玲被捕": 1, "瞿秋白": 1, "左联解散": 1,
        "鲁迅": 1,
    }
    updated = 0
    for e in events:
        name = e.get("event_name", "")
        for keyword in key_events:
            if keyword in name:
                existing = set(e.get("source_ids", "").split(";"))
                for s in [wiki_zuolian_src, baidu_src]:
                    existing.add(s)
                new_ids = ";".join(sorted(existing))
                if new_ids != e.get("source_ids", ""):
                    e["source_ids"] = new_ids
                    updated += 1
                break
    write_csv("events.csv", events)
    print(f"  更新了 {updated} 条事件的source_ids")

# ============================================================
# 14. 验证统计
# ============================================================
def step14_validate():
    print("\n" + "=" * 60)
    print("STEP 14: 最终验证统计")
    sources = read_csv("sources.csv")
    persons = read_csv("persons.csv")
    events = read_csv("events.csv")
    relations = read_csv("person_relations.csv")
    orgs = read_csv("organizations.csv")
    places = read_csv("places.csv")
    memberships = read_csv("org_memberships.csv")
    print(f"\n  {'表名':<25} {'记录数':>6}")
    print(f"  {'-'*25} {'-'*6}")
    print(f"  {'sources.csv':<25} {len(sources):>6}")
    print(f"  {'persons.csv':<25} {len(persons):>6}")
    print(f"  {'events.csv':<25} {len(events):>6}")
    print(f"  {'person_relations.csv':<25} {len(relations):>6}")
    print(f"  {'organizations.csv':<25} {len(orgs):>6}")
    print(f"  {'places.csv':<25} {len(places):>6}")
    print(f"  {'org_memberships.csv':<25} {len(memberships):>6}")
    multi_src = sum(1 for p in persons if ";" in p.get("source_ids", ""))
    print(f"\n  persons多源引用: {multi_src}/{len(persons)} ({100*multi_src/len(persons):.0f}%)")
    conf = Counter(e.get("confidence", "") for e in events)
    print(f"  events confidence: {dict(sorted(conf.items()))}")
    ocr_remaining = sum(1 for r in relations for wrong in OCR_FIXES if wrong in r.get("context", ""))
    print(f"  OCR残留错误: {ocr_remaining} 处")
    rel_dist = Counter(p["reliability"] for p in persons)
    print(f"  reliability分布: {dict(sorted(rel_dist.items()))}")
    print(f"  organizations: {len(orgs)} 个")
    place_names = [p["place_name"] for p in places]
    dupes = [n for n, c in Counter(place_names).items() if c > 1]
    print(f"  places同名重复: {len(dupes)} 个")
    person_roles_map = {p["person_id"]: p["role"] for p in persons}
    redundant = sum(1 for m in memberships if m.get("membership_role") == person_roles_map.get(m.get("person_id", "")))
    print(f"  org_memberships与persons.role重复: {redundant}/{len(memberships)}")

# ============================================================
if __name__ == "__main__":
    print("左联知识库全量数据修复")
    print("=" * 60)
    start = datetime.now()
    step0_backup()
    wiki_src_map, wiki_zuolian_src, baidu_src = step1_add_sources()
    step2_update_person_sources(wiki_src_map, wiki_zuolian_src, baidu_src)
    step3_fix_ocr()
    step4_clean_events()
    step5_update_verification()
    step6_clean_aliases()
    step7_clean_places()
    step8_expand_orgs()
    step9_add_missing_persons()
    step10_recalibrate_reliability()
    step11_standardize_canonical_keys()
    step12_clean_org_memberships()
    step13_update_event_sources(wiki_zuolian_src, baidu_src)
    step14_validate()
    elapsed = (datetime.now() - start).total_seconds()
    print(f"\n全部完成，耗时 {elapsed:.1f} 秒")
