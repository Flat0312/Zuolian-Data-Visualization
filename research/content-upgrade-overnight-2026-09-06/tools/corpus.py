# -*- coding: utf-8 -*-
"""本轮专用只读语料索引：词典/史按页、日记按日期。只读原文件，输出到 R。"""
import json, re, os, hashlib

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))
DICT = os.path.join(ROOT, "data", "processed", "runtime_sources", "左联词典.txt")
HIST = os.path.join(ROOT, "data", "processed", "runtime_sources", "左联史.txt")
DIARY = os.path.join(ROOT, "research", "raw_texts", "日记全编：全2册 (鲁迅 著) (Z-Library).txt")
MEMOIR = os.path.join(ROOT, "data", "processed", "runtime_sources", "左联回忆录_ocr_text.json")

_CN_NUM = {"〇":0,"一":1,"二":2,"三":3,"四":4,"五":5,"六":6,"七":7,"八":8,"九":9,"十":10,
           "二十":20,"三十":30,"三十一":31,"二十一":21,"二十二":22,"二十三":23,"二十四":24,
           "二十五":25,"二十六":26,"二十七":27,"二十八":28,"二十九":29}

def _pages(path):
    t = open(path, encoding="utf-8").read()
    parts = re.split(r"[─]{10,}\s*\n\s*第\s*(\d+)\s*页\s*\n\s*[─]{10,}", t)
    out = {}
    for i in range(1, len(parts) - 1, 2):
        out[int(parts[i])] = parts[i + 1].strip()
    return out

def dict_pages():
    return _pages(DICT)

def hist_pages():
    return _pages(HIST)

def memoir():
    try:
        return json.load(open(MEMOIR, encoding="utf-8"))
    except Exception:
        return {}

def sha256_file(path):
    return hashlib.sha256(open(path, "rb").read()).hexdigest()

def diary_index():
    """返回 {(year,month,day): entry_text}。日记全编结构：日记十X（YYYY年）→ 一月…十二月 → N日　…"""
    t = open(DIARY, encoding="utf-8").read()
    yr = None
    mo = None
    idx = {}
    # split by year header
    for ym in re.finditer(r"日记十[一二三四五六七八九]{0,2}[（(](\d{4})年[)）]", t):
        pass
    year_headers = [(m.start(), int(m.group(1))) for m in re.finditer(r"日记[^\n（(]{0,6}[（(](\d{4})年[)）]", t)]
    for yi, (ystart, year) in enumerate(year_headers):
        yend = year_headers[yi + 1][0] if yi + 1 < len(year_headers) else len(t)
        body = t[ystart:yend]
        # months inside year; month headers appear as bare "一月" lines (both TOC and body; take last occurrence group)
        month_headers = [(m.start(), m.group(1)) for m in re.finditer(r"\n+(一月|二月|三月|四月|五月|六月|七月|八月|九月|十月|十一月|十二月)\n", body)]
        _CNM = {"一月":1,"二月":2,"三月":3,"四月":4,"五月":5,"六月":6,"七月":7,"八月":8,"九月":9,"十月":10,"十一月":11,"十二月":12}
        # TOC months come first in a compact block; body months precede entry text.
        # A body month is followed within 200 chars by a day entry.
        months = []
        for mi, (mstart, mname) in enumerate(month_headers):
            mend = month_headers[mi + 1][0] if mi + 1 < len(month_headers) else len(body)
            seg = body[mstart:mend]
            if re.search(r"\n+\s*[一二三四五六七八九十]{1,3}日[\s\u3000]", seg[:400]):
                months.append((mstart, mend, _CNM[mname]))
        for msi, (mstart, mend, month) in enumerate(months):
            seg = body[mstart:mend]
            day_pat = list(re.finditer(r"\n+\s*([一二三四五六七八九十]{1,3})日[\s\u3000]", seg))
            for di, dm in enumerate(day_pat):
                day = _cn_to_int(dm.group(1))
                dend = day_pat[di + 1].start() if di + 1 < len(day_pat) else len(seg)
                entry = seg[dm.end():dend].strip()
                key = (year, month, day)
                if key not in idx:
                    idx[key] = entry
                else:
                    idx[key] += "\n" + entry
    return idx

def _cn_to_int(s):
    s = s.strip()
    if s == "十": return 10
    if s == "二十": return 20
    if s == "三十": return 30
    if "十" in s:
        a, b = s.split("十")
        tens = _cn_digit(a) if a else 1
        ones = _cn_digit(b) if b else 0
        return tens * 10 + ones
    return _cn_digit(s)

def _cn_digit(c):
    return {"一":1,"二":2,"三":3,"四":4,"五":5,"六":6,"七":7,"八":8,"九":9,"〇":0,"０":0}.get(c, 0)

if __name__ == "__main__":
    dp = dict_pages(); hp = hist_pages(); di = diary_index()
    print("dict pages", len(dp), "hist pages", len(hp), "diary entries", len(di))
    ys = sorted({k[0] for k in di})
    print("diary years", ys[:5], "...", ys[-3:])
    e = di.get((1928, 11, 24), "")
    print("1928-11-24:", e[:120])
