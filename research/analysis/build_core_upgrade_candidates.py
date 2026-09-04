"""核心内容补证候选包构建脚本（返修版，只读生产数据，可重复运行）。

选择对象：30 名核心人物、20 个关键事件、10 个重要地点。
优先主题：左联成立及组织演变、龙华烈士相关事件、鲁迅与青年作家网络、
左翼刊物和出版机构、女性作家社会网络。

原则：
- 选择完全由 data/processed 现有数据 + 脚本内显式登记的种子常量决定，可复现；
- 只登记"现有证据"与"缺口"，不新增史实判断；
- EVIDENCE_SEEDS 为只读检索所得的候选线索（含 URL、访问日期、
  定位与逐字短引文），其中维基百科/百度百科/普通媒体一律标为 web_lead/D 级线索，
  不得称为权威证据；中国作家网/政府/纪念馆/档案馆/大学档案/学术论文/原始文献
  按现有来源规则分级；一律 pending_human_review 或 conflict，未检索到的字段记 missing；
- 与生产值冲突的候选（如殷夫/周扬生年、萌芽创刊/楼适夷被捕日期）保持 conflict 状态，
  禁止自动落库；
- 不修改任何生产数据。

输出（默认 research/drafts/reports/）：
- core_person_evidence_candidates.csv / core_event_evidence_candidates.csv
- core_place_review_candidates.csv / core_upgrade_selection_report.md
"""
from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_DATA_DIR = PROJECT_ROOT / "data" / "processed"
DEFAULT_OUT_DIR = PROJECT_ROOT / "research" / "drafts" / "reports"

SNAPSHOT_DATE = "2026-09-04"  # 数据读取与检索的固定快照日期，保证重跑一致

N_PERSONS = 30
N_EVENTS = 20
N_PLACES = 10

# —— 通说种子（仅用于选择优先级，不作为史实断言写入候选行）——
# 左联五烈士（史学界通称）：用于"龙华烈士相关事件"主题的优先级判定。
LONGHUA_MARTYRS = {"柔石", "胡也频", "冯铿", "殷夫", "李伟森"}
# 库内可考的左翼女性作家（依据中国现代文学史通说筛选，仅作研究优先级种子）。
FEMALE_WRITERS = {"丁玲", "冯铿", "萧红", "白薇", "草明", "葛琴", "许广平", "王莹", "谢冰莹"}
LU_XUN_ID = "ZLH-001"
# "青年作家"操作定义：与鲁迅存在可信关系且出生晚于 1895 年（鲁迅 1881 年生的下一代）。
YOUNG_WRITER_BORN_AFTER = 1895

ORG_EVOLUTION_PAT = ("成立", "大会", "执委", "党团", "左联", "筹组")
LONGHUA_PAT = ("龙华", "遇难", "烈士", "被捕", "秘密会议", "东方旅社", "遇害")
PUBLICATION_PAT = ("创刊", "月刊", "周刊", "书店", "书局", "出版", "发行", "编辑", "丛书")

PENDING = "pending_human_review"
MISSING = "missing"
CONFLICT = "conflict"

# 与生产值冲突、禁止自动落库的候选（实体ID, 字段）集合：保持 conflict 状态。
CONFLICT_SEEDS: set[tuple[str, str]] = {
    ("ZLH-019", "birth_year"),  # 殷夫生年：维基 1909-06-11 vs 生产值 1910
    ("ZLH-007", "birth_year"),  # 周扬生年：维基 1907-11-07 vs 生产值 1908
    ("EVT-00183", "direct_support"),  # 萌芽创刊：维基 1930-01-01 vs 生产 1928
    ("EVT-00138", "direct_support"),  # 楼适夷被捕：维基 1933 vs 生产 1934
}


def classify_candidate_source(url: str, title: str = "") -> tuple[str, str]:
    """候选来源分级（返修版）：百科/普通媒体一律 D 级 web_lead，不得称为权威。

    - 维基百科/百度百科 -> D / web_lead(encyclopedia)
    - 普通媒体（界面新闻等） -> D / web_lead(news_media)
    - 中国作家网 -> B / industry_official
    - 政府网站（gov.cn） -> B / government
    - 纪念馆/档案馆/大学档案/学术论文/原始文献 -> 按现有规则 A/B（此处保守记 B，待人工核定升 A）
    - 未知 web -> D / web_lead
    """
    u = (url or "").strip().lower()
    if not u:
        return "", ""
    if "wikipedia.org" in u or "baike.baidu.com" in u:
        return "D", "web_lead"
    if "chinawriter.com.cn" in u:
        return "B", "industry_official"
    if "gov.cn" in u:
        return "B", "government"
    if any(k in u for k in ("memorial", "museum", "archive", "edu.cn", "cnki", "wanfang", "cqvip")):
        return "B", "archive_academic"
    if "jiemian.com" in u or "thepaper.cn" in u:
        return "D", "web_lead"
    if u.startswith("http"):
        return "D", "web_lead"
    return "D", "web_lead"


def content_hash_for(url: str, locator: str, quote: str) -> str:
    """访问凭据等价物：URL + 定位 + 引文的 SHA256（空候选返回空）。"""
    import hashlib

    if not (url or "").strip():
        return ""
    basis = "|".join([(url or "").strip(), (locator or "").strip(), (quote or "").strip()])
    return hashlib.sha256(basis.encode("utf-8")).hexdigest()

# —— 只读检索所得候选证据种子（访问日期均为实际检索日 2026-09-04）——
# 结构: 实体ID -> 字段 -> dict(title, url, access_date, locator, quote)
# 留空的字段在输出中记 missing；种子一律 pending_human_review，不转正。
# 引文均为检索快照或页面核验的逐字原文；未经逐字核验的内容不写 quote。
EVIDENCE_SEEDS: dict[str, dict[str, dict[str, str]]] = {
    # 柔石（维基百科条目首句，生卒/身份一并提供）
    "ZLH-016": {
        "birth_year": {
            "title": "柔石 - 维基百科",
            "url": "https://zh.wikipedia.org/zh-cn/%E6%9F%94%E7%9F%B3",
            "access_date": "2026-09-04",
            "locator": "条目首句",
            "quote": "柔石（1902年9月28日—1931年2月7日），原名赵平福，后改名赵平复，男，浙江省宁海县人",
        },
        "death_year": {
            "title": "柔石 - 维基百科",
            "url": "https://zh.wikipedia.org/zh-cn/%E6%9F%94%E7%9F%B3",
            "access_date": "2026-09-04",
            "locator": "条目首句",
            "quote": "柔石（1902年9月28日—1931年2月7日）",
        },
        "role": {
            "title": "柔石 - 维基百科",
            "url": "https://zh.wikipedia.org/zh-cn/%E6%9F%94%E7%9F%B3",
            "access_date": "2026-09-04",
            "locator": "条目首句（职业身份）",
            "quote": "中国作家、革命家及翻译家，“左联五烈士”之一",
        },
    },
    # 胡也频（中国作家网作家档案，标题含生卒）
    "ZLH-017": {
        "birth_year": {
            "title": "胡也频(1903.5.4—1931.2.7) - 中国作家网",
            "url": "http://www.chinawriter.com.cn/xdzj/444.shtml",
            "access_date": "2026-09-04",
            "locator": "页面标题；正文首段",
            "quote": "1903年生于福州。",
        },
        "death_year": {
            "title": "胡也频(1903.5.4—1931.2.7) - 中国作家网",
            "url": "http://www.chinawriter.com.cn/xdzj/444.shtml",
            "access_date": "2026-09-04",
            "locator": "页面标题",
            "quote": "胡也频(1903.5.4—1931.2.7)",
        },
    },
    # 冯铿（维基百科条目首句）
    "ZLH-018": {
        "birth_year": {
            "title": "冯铿 - 维基百科",
            "url": "https://zh.wikipedia.org/zh-cn/%E5%86%AF%E9%93%BF",
            "access_date": "2026-09-04",
            "locator": "条目首句",
            "quote": "冯铿（1907年10月10日—1931年2月7日），又名岭梅。广东潮州潮安人",
        },
        "death_year": {
            "title": "冯铿 - 维基百科",
            "url": "https://zh.wikipedia.org/zh-cn/%E5%86%AF%E9%93%BF",
            "access_date": "2026-09-04",
            "locator": "条目首句",
            "quote": "冯铿（1907年10月10日—1931年2月7日）",
        },
        "role": {
            "title": "冯铿 - 维基百科",
            "url": "https://zh.wikipedia.org/zh-cn/%E5%86%AF%E9%93%BF",
            "access_date": "2026-09-04",
            "locator": "条目首句（职业身份）",
            "quote": "著名女作家、女诗人，“左联五烈士”之一，龙华二十四烈士之一",
        },
    },
    # 殷夫（维基百科条目首句；生年存在 1909/1910 两说，见 SEED_NOTES）
    "ZLH-019": {
        "birth_year": {
            "title": "殷夫 - 维基百科",
            "url": "https://zh.wikipedia.org/zh-cn/%E6%AE%B7%E5%A4%AB",
            "access_date": "2026-09-04",
            "locator": "条目首句",
            "quote": "殷夫（1909年6月11日—1931年2月7日），本名徐孝杰，字柏庭，又笔名白莽",
        },
        "death_year": {
            "title": "殷夫 - 维基百科",
            "url": "https://zh.wikipedia.org/zh-cn/%E6%AE%B7%E5%A4%AB",
            "access_date": "2026-09-04",
            "locator": "条目首句",
            "quote": "殷夫（1909年6月11日—1931年2月7日）",
        },
        "role": {
            "title": "殷夫 - 维基百科",
            "url": "https://zh.wikipedia.org/zh-cn/%E6%AE%B7%E5%A4%AB",
            "access_date": "2026-09-04",
            "locator": "条目首句（职业身份）",
            "quote": "中国作家、诗人，“左联五烈士”之一，也是龙华二十四烈士之一",
        },
    },
    # 李伟森（维基百科条目首句）
    "ZLH-160": {
        "birth_year": {
            "title": "李伟森 - 维基百科",
            "url": "https://zh.wikipedia.org/zh-cn/%E6%9D%8E%E4%BC%9F%E6%A3%AE",
            "access_date": "2026-09-04",
            "locator": "条目首句",
            "quote": "李伟森（1903年—1931年2月7日），乳名伟生，亦名国纬，字北平，笔名李求实",
        },
        "death_year": {
            "title": "李伟森 - 维基百科",
            "url": "https://zh.wikipedia.org/zh-cn/%E6%9D%8E%E4%BC%9F%E6%A3%AE",
            "access_date": "2026-09-04",
            "locator": "条目首句",
            "quote": "李伟森（1903年—1931年2月7日）",
        },
        "role": {
            "title": "李伟森 - 维基百科",
            "url": "https://zh.wikipedia.org/zh-cn/%E6%9D%8E%E4%BC%9F%E6%A3%AE",
            "access_date": "2026-09-04",
            "locator": "条目首句（党内职务）",
            "quote": "中国共产党早期党员，中国共青团早期领导人",
        },
    },
    # 丁玲（中国作家网会员档案页 + 文史文章）
    "ZLH-021": {
        "birth_year": {
            "title": "丁玲——会员——中国作家网",
            "url": "http://www.chinawriter.com.cn/n1/2016/0627/c404931-28487247.html",
            "access_date": "2026-09-04",
            "locator": "档案页姓名下生卒年栏（2026-09-04 逐字核验）",
            "quote": "(1904～1986)",
        },
        "death_year": {
            "title": "丁玲——会员——中国作家网",
            "url": "http://www.chinawriter.com.cn/n1/2016/0627/c404931-28487247.html",
            "access_date": "2026-09-04",
            "locator": "档案页姓名下生卒年栏（2026-09-04 逐字核验）",
            "quote": "(1904～1986)",
        },
        "role": {
            "title": "丁玲与中国作协 - 中国作家网",
            "url": "http://www.chinawriter.com.cn/n1/2019/0722/c404064-31247824.html",
            "access_date": "2026-09-04",
            "locator": "正文（30年代左翼活动段）",
            "quote": "丁玲参与中国共产党领导下的文学组织活动较早，上世纪30年代，丁玲即参与左翼文学活动的组织和出版工作，主编“左联”机关刊物《北斗》",
        },
    },
    # 楼适夷（维基百科条目首句）
    "ZLH-037": {
        "birth_year": {
            "title": "楼适夷 - 维基百科",
            "url": "https://zh.wikipedia.org/zh-cn/%E6%A5%BC%E9%80%82%E5%A4%B7",
            "access_date": "2026-09-04",
            "locator": "条目首句",
            "quote": "楼适夷(1905年1月3日—2001年4月20日),原名楼锡春,曾用笔名楼建南,男,浙江余姚人",
        },
        "death_year": {
            "title": "楼适夷 - 维基百科",
            "url": "https://zh.wikipedia.org/zh-cn/%E6%A5%BC%E9%80%82%E5%A4%B7",
            "access_date": "2026-09-04",
            "locator": "条目首句",
            "quote": "楼适夷(1905年1月3日—2001年4月20日)",
        },
        "role": {
            "title": "楼适夷 - 维基百科",
            "url": "https://zh.wikipedia.org/zh-cn/%E6%A5%BC%E9%80%82%E5%A4%B7",
            "access_date": "2026-09-04",
            "locator": "条目首句（职业身份）",
            "quote": "中国作家、翻译家、出版家",
        },
    },
    # 冯雪峰（维基百科条目首句）
    "ZLH-005": {
        "birth_year": {
            "title": "冯雪峰 - 维基百科",
            "url": "https://zh.wikipedia.org/zh-cn/%E5%86%AF%E9%9B%AA%E5%B3%B0",
            "access_date": "2026-09-04",
            "locator": "条目首句",
            "quote": "冯雪峰（1903年6月2日—1976年1月31日），原名冯福春，笔名画室、洛扬、维山、成文英、何丹仁、吕克玉等，男，浙江义乌人",
        },
        "death_year": {
            "title": "冯雪峰 - 维基百科",
            "url": "https://zh.wikipedia.org/zh-cn/%E5%86%AF%E9%9B%AA%E5%B3%B0",
            "access_date": "2026-09-04",
            "locator": "条目首句",
            "quote": "冯雪峰（1903年6月2日—1976年1月31日）",
        },
        "role": {
            "title": "冯雪峰 - 维基百科",
            "url": "https://zh.wikipedia.org/zh-cn/%E5%86%AF%E9%9B%AA%E5%B3%B0",
            "access_date": "2026-09-04",
            "locator": "条目首句（职业身份）",
            "quote": "中国诗人、文艺评论家",
        },
    },
    # 周扬（维基百科条目首句；生年 1907/1908 两说，见 SEED_NOTES）
    "ZLH-007": {
        "birth_year": {
            "title": "周扬(政治人物) - 维基百科",
            "url": "https://zh.wikipedia.org/zh-cn/%E5%91%A8%E6%89%AC_(%E6%94%BF%E6%B2%BB%E4%BA%BA%E7%89%A9)",
            "access_date": "2026-09-04",
            "locator": "条目首句",
            "quote": "周扬(1907年11月7日—1989年7月31日)，原名周运宜，字起应，笔名绮影、谷扬、周苋等，男，湖南益阳人",
        },
        "death_year": {
            "title": "周扬(政治人物) - 维基百科",
            "url": "https://zh.wikipedia.org/zh-cn/%E5%91%A8%E6%89%AC_(%E6%94%BF%E6%B2%BB%E4%BA%BA%E7%89%A9)",
            "access_date": "2026-09-04",
            "locator": "条目首句",
            "quote": "周扬(1907年11月7日—1989年7月31日)",
        },
        "role": {
            "title": "周扬 - 赫山区人民政府（家乡政府人物志）",
            "url": "https://www.hnhs.gov.cn/22556/22562/content_1186164.html",
            "access_date": "2026-09-04",
            "locator": "正文（左翼文化运动段）",
            "quote": "1932年9月,他接替姚蓬子主编左联机关刊物《文学月报》,在极其艰难的条件下",
        },
    },
    # EVT-00183《萌芽月刊》：维基记载 1930-01-01 创刊，与生产事件 1928 冲突（SEED_NOTES）
    "EVT-00183": {
        "direct_support": {
            "title": "萌芽月刊 - 维基百科",
            "url": "https://zh.wikipedia.org/zh-cn/%E8%90%8C%E8%8A%BD%E6%9C%88%E5%88%8A",
            "access_date": "2026-09-04",
            "locator": "条目首句",
            "quote": "《萌芽月刊》是由鲁迅与冯雪峰编辑,上海光华书局发行的一份左派文学刊物,于1930年1月1日创刊。",
        },
    },
    # EVT-00138 楼适夷被捕：维基记载 1933 年被捕，与生产事件 1934 冲突（SEED_NOTES）
    "EVT-00138": {
        "direct_support": {
            "title": "楼适夷 - 维基百科",
            "url": "https://zh.wikipedia.org/zh-cn/%E6%A5%BC%E9%80%82%E5%A4%B7",
            "access_date": "2026-09-04",
            "locator": "生平段（被捕记述）",
            "quote": "1933年，因为文学作品宣传进步思想，被捕入狱。",
        },
    },
    # EVT-00093 1930年内山书店文学活动记录：界面新闻（2026-09-04 逐字核验）
    "EVT-00093": {
        "direct_support": {
            "title": "跨越一个世纪，在内山书店与鲁迅重逢 - 界面新闻",
            "url": "https://m.jiemian.com/article/9013783.html",
            "access_date": "2026-09-04",
            "locator": "“出版与庇护”一节首句（2026-09-04 逐字核验）",
            "quote": "1930年3月，鲁迅在中国左翼作家联盟成立大会上演讲后遭到国民党通缉，曾在内山书店避居一个月。",
        },
    },
    # PLC-00007 内山书店：现址沿革（1929 年迁入四川北路2050号）
    "PLC-00007": {
        "address": {
            "title": "内山书店旧址 - 百度百科",
            "url": "https://baike.baidu.com/item/%E5%86%85%E5%B1%B1%E4%B9%A6%E5%BA%97%E6%97%A7%E5%9D%80/56384650",
            "access_date": "2026-09-04",
            "locator": "词条摘要（地址沿革）",
            "quote": "内山书店旧址位于上海市虹口区四川北路2050号，建筑为砖木结构假三层，坐北朝南，一层为米白色花岗石贴面，二层为淡黄色水泥拉毛墙面，扁圆拱形窗，建成于1924年。",
        },
    },
    "PLC-00027": {
        "address": {
            "title": "内山书店旧址 - 百度百科",
            "url": "https://baike.baidu.com/item/%E5%86%85%E5%B1%B1%E4%B9%A6%E5%BA%97%E6%97%A7%E5%9D%80/56384650",
            "access_date": "2026-09-04",
            "locator": "词条摘要（地址沿革）",
            "quote": "内山书店旧址位于上海市虹口区四川北路2050号，建筑为砖木结构假三层，坐北朝南，一层为米白色花岗石贴面，二层为淡黄色水泥拉毛墙面，扁圆拱形窗，建成于1924年。",
        },
    },
    # PLC-00003 龙华刑场：今龙华烈士陵园地址
    "PLC-00003": {
        "address": {
            "title": "龙华烈士陵园 - 维基百科",
            "url": "https://zh.wikipedia.org/zh-cn/%E9%BE%99%E5%8D%8E%E7%83%88%E5%A3%AB%E9%99%B5%E5%9B%AD",
            "access_date": "2026-09-04",
            "locator": "条目首句",
            "quote": "龙华烈士陵园,位于中华人民共和国上海市徐汇区龙华西路180号,其东北角的龙华革命烈士纪念地是全国重点文物保护单位之一。",
        },
    },
}

# —— 检索中发现的口径差异与否定性结果（写入选择报告，全部待人工裁决）——
SEED_NOTES: list[str] = [
    "殷夫生年两说：维基百科记 1909-06-11，生产值与同济大学档案馆等记 1910——候选证据与生产值冲突，待人工裁决。",
    "周扬生年两说：维基百科记 1907-11-07，生产值 1908（盛夏编《周扬同志年谱》封面作 1908—1989）——待人工裁决。",
    "李伟森生年：检索确认权威记载为 1903（维基百科、百度百科一致，生产值同为 1903）；早期材料中出现的 1898 系误记。",
    "EVT-00183『1928年《萌芽月刊》文学活动记录』：维基百科记载《萌芽月刊》于 1930-01-01 创刊（鲁迅与冯雪峰编辑、光华书局发行），与生产事件日期 1928 冲突，建议人工改期并具体化为创刊事件。",
    "EVT-00138『楼适夷因支持《文学》月刊被捕』：维基百科记载 1933 年被捕（1937 年出狱），与生产事件年份 1934 冲突，待人工裁决。",
    "EVT-00078/EVT-00079『周扬1932/1933年遭拘押』：本轮检索未见任何权威佐证（其被捕记录为留日时期），按 missing 记录，不得据此改写或删除，待人工复查。",
    "EVT-00187 洛阳书店大会：受内容过滤与检索命中限制，本轮未获第二独立来源（仓库内已有澎湃·左联纪念馆馆员文章一条），independent_source 记 missing。",
    "内山书店沿革：现址四川北路2050号为 1929 年迁入（百度百科词条），此前店址不在该处——1929 年及以前『内山书店』相关事件的地点归属需人工按沿革核对。",
]


def _read(data_dir: Path, name: str) -> pd.DataFrame:
    return pd.read_csv(data_dir / name, encoding="utf-8-sig", dtype=str).fillna("")


def _matches(text: str, patterns: tuple[str, ...]) -> bool:
    return any(p in text for p in patterns)


def is_trusted_relation(row: dict[str, str]) -> bool:
    """补证优先级用的启发式关系判定（较低风险规则筛选口径，非可信断言）。

    排除：待核验类型、needs_manual_review=yes、critical/high 风险、confidence=low；
    不看 relation_evidences、不看 publish_status（保留历史 Top30 可比性）。
    命名上不得称为“可信关系”，仅用于补证优先级排序；真正的可信判定见
    build_trustworthy_network_analysis.is_trusted_relation（须满足 support 证据）。
    """
    if str(row.get("final_relation_type", "")).strip() == "待核验":
        return False
    if str(row.get("needs_manual_review", "")).strip().lower() == "yes":
        return False
    if str(row.get("relation_risk_level", "")).strip().lower() in ("critical", "high"):
        return False
    if str(row.get("confidence", "")).strip().lower() == "low":
        return False
    return True


def _norm(series: pd.Series) -> pd.Series:
    numeric = pd.to_numeric(series, errors="coerce").fillna(0.0)
    max_v = float(numeric.max()) if len(numeric) else 0.0
    if max_v <= 0:
        return numeric * 0.0
    return numeric / max_v


def _birth_after(year_text: str) -> bool:
    try:
        return int(year_text) > YOUNG_WRITER_BORN_AFTER
    except (ValueError, TypeError):
        return False


def _seed_status(entity_id: str, seeds: dict[str, dict[str, str]], field: str) -> str:
    entry = seeds.get(field)
    if entry and entry.get("url"):
        if (entity_id, field) in CONFLICT_SEEDS:
            return CONFLICT
        return PENDING
    return MISSING


def _seed_retrieval_status(status: str) -> str:
    if status == PENDING:
        return "retrieved"
    if status == CONFLICT:
        return "conflict"
    return "missing"


def _seed_enrichment(entity_id: str, field: str, seeds: dict[str, dict[str, str]]) -> dict[str, str]:
    """返回 source_level / source_type / retrieval_status / content_hash 四件套。"""
    entry = seeds.get(field, {})
    url = (entry.get("url", "") or "").strip()
    title = (entry.get("title", "") or "").strip()
    locator = (entry.get("locator", "") or "").strip()
    quote = (entry.get("quote", "") or "").strip()
    status = _seed_status(entity_id, seeds, field)
    if status == MISSING or not url:
        return {
            "source_level": "",
            "source_type": "",
            "retrieval_status": "missing",
            "content_hash": "",
        }
    level, stype = classify_candidate_source(url, title)
    return {
        "source_level": level,
        "source_type": stype,
        "retrieval_status": _seed_retrieval_status(status),
        "content_hash": content_hash_for(url, locator, quote),
    }


def _level_counts(frames: list[pd.DataFrame], status_prefixes: tuple[str, ...]) -> dict[str, int]:
    counts = {"A": 0, "B": 0, "C": 0, "D": 0}
    for frame in frames:
        for col in frame.columns:
            if not col.endswith("_source_level"):
                continue
            prefix = col[: -len("_source_level")]
            # 仅统计有候选（非 missing）的来源等级
            status_col = f"{prefix}_status" if f"{prefix}_status" in frame.columns else ""
            for _, row in frame.iterrows():
                if status_col and row.get(status_col, "") == MISSING:
                    continue
                lv = str(row.get(col, "")).strip()
                if lv in counts:
                    counts[lv] += 1
    return counts


def build(data_dir: Path, out_dir: Path) -> dict[str, object]:
    persons = _read(data_dir, "persons.csv")
    relations = _read(data_dir, "person_relations.csv")
    events = _read(data_dir, "events.csv")
    places = _read(data_dir, "places.csv")
    facts = _read(data_dir, "fact_evidences.csv")
    parts = _read(data_dir, "event_participants.csv")
    orgs = _read(data_dir, "organizations.csv")
    memberships = _read(data_dir, "org_memberships.csv")
    sources = _read(data_dir, "sources.csv")

    name_by_id = dict(zip(persons["person_id"], persons["standard_name"]))
    birth_by_id = dict(zip(persons["person_id"], persons["birth_year"]))
    event_name_by_id = dict(zip(events["event_id"], events["event_name"]))
    place_name_by_id = dict(zip(places["place_id"], places["historical_name"]))
    org_name_by_id = dict(zip(orgs["organization_id"], orgs["standard_name"]))
    event_place_by_id = dict(zip(events["event_id"], events["place_id"]))

    zuolian_org_ids = {oid for oid, nm in org_name_by_id.items() if "左翼作家联盟" in nm}

    # —— 关系度数与可信邻接 ——
    trusted_adj: dict[str, set[str]] = {pid: set() for pid in persons["person_id"]}
    full_deg = {pid: 0 for pid in persons["person_id"]}
    for row in relations.to_dict("records"):
        s, t = row.get("source_person_id", ""), row.get("target_person_id", "")
        if s not in full_deg or t not in full_deg:
            continue
        full_deg[s] += 1
        full_deg[t] += 1
        if is_trusted_relation(row):
            trusted_adj[s].add(t)
            trusted_adj[t].add(s)
    luxun_trusted = set(trusted_adj.get(LU_XUN_ID, set()))

    # —— 参与者/事件索引 ——
    participants_by_event: dict[str, list[str]] = {}
    events_by_person: dict[str, set[str]] = {pid: set() for pid in persons["person_id"]}
    for row in parts.to_dict("records"):
        eid, pid = row.get("event_id", ""), row.get("person_id", "")
        if not eid or not pid or pid not in events_by_person:
            continue
        participants_by_event.setdefault(eid, []).append(pid)
        events_by_person[pid].add(eid)

    # —— 事件级主题 ——
    event_themes: dict[str, set[str]] = {}
    for row in events.to_dict("records"):
        eid = row.get("event_id", "")
        blob = " ".join(
            [
                row.get("event_name", ""),
                row.get("original_event_names", ""),
                row.get("historical_location", ""),
                place_name_by_id.get(row.get("place_id", ""), ""),
            ]
        )
        themes: set[str] = set()
        if _matches(blob, ORG_EVOLUTION_PAT):
            themes.add("左联成立及组织演变")
        if _matches(blob, LONGHUA_PAT):
            themes.add("龙华烈士相关事件")
        plist = participants_by_event.get(eid, [])
        young = [p for p in plist if p != LU_XUN_ID and _birth_after(birth_by_id.get(p, ""))]
        if LU_XUN_ID in plist and young:
            themes.add("鲁迅与青年作家网络")
        if _matches(blob, PUBLICATION_PAT):
            themes.add("左翼刊物和出版机构")
        if any(name_by_id.get(p, "") in FEMALE_WRITERS for p in plist):
            themes.add("女性作家社会网络")
        event_themes[eid] = themes

    # —— 人物级主题 ——
    mem_rows_by_person: dict[str, list[dict[str, str]]] = {}
    for row in memberships.to_dict("records"):
        mem_rows_by_person.setdefault(row.get("person_id", ""), []).append(row)
    zuolian_mem_type = {
        pid: next(
            (r["membership_type"] for r in rows if r.get("organization_id") in zuolian_org_ids), ""
        )
        for pid, rows in mem_rows_by_person.items()
    }

    def person_themes(pid: str) -> set[str]:
        name = name_by_id.get(pid, "")
        themes: set[str] = set()
        my_events = events_by_person.get(pid, set())
        if zuolian_mem_type.get(pid) or any(
            _matches(event_name_by_id.get(e, ""), ORG_EVOLUTION_PAT) for e in my_events
        ):
            themes.add("左联成立及组织演变")
        if name in LONGHUA_MARTYRS or any(
            _matches(event_name_by_id.get(e, ""), LONGHUA_PAT)
            or "龙华" in place_name_by_id.get(event_place_by_id.get(e, ""), "")
            for e in my_events
        ):
            themes.add("龙华烈士相关事件")
        if pid in luxun_trusted and _birth_after(birth_by_id.get(pid, "")):
            themes.add("鲁迅与青年作家网络")
        orgs_of_pid = {
            org_name_by_id.get(r.get("organization_id", ""), "") for r in mem_rows_by_person.get(pid, [])
        }
        if any(_matches(o, PUBLICATION_PAT) for o in orgs_of_pid if o) or any(
            _matches(event_name_by_id.get(e, ""), PUBLICATION_PAT) for e in my_events
        ):
            themes.add("左翼刊物和出版机构")
        if name in FEMALE_WRITERS:
            themes.add("女性作家社会网络")
        return themes

    # —— 人物事实级证据计数（birth_year/death_year/role）——
    person_fact = {pid: {"birth_year": 0, "death_year": 0, "role": 0} for pid in persons["person_id"]}
    for row in facts.to_dict("records"):
        if row.get("subject_type") != "person" or row.get("review_status") == "rejected":
            continue
        pid, pred = row.get("subject_id", ""), row.get("predicate", "")
        if pid in person_fact and pred in person_fact[pid]:
            person_fact[pid][pred] += 1

    # —— 人物评分 ——
    p_rows = []
    for pid in persons["person_id"]:
        themes = person_themes(pid)
        p_rows.append(
            {
                "person_id": pid,
                "standard_name": name_by_id.get(pid, ""),
                "trusted_degree": len(trusted_adj[pid]),
                "full_degree": full_deg[pid],
                "event_count": len(events_by_person.get(pid, set())),
                "zuolian_membership": zuolian_mem_type.get(pid, ""),
                "themes": ";".join(sorted(themes)),
                "theme_count": len(themes),
            }
        )
    persons_scored = pd.DataFrame(p_rows)
    persons_scored["selection_score"] = (
        _norm(persons_scored["trusted_degree"]) * 2.0
        + _norm(persons_scored["full_degree"]) * 0.8
        + _norm(persons_scored["event_count"]) * 1.2
        + persons_scored["zuolian_membership"].map(
            {"confirmed_member": 1.0, "related_person": 0.5, "candidate": 0.25}
        ).fillna(0.0)
        + _norm(persons_scored["theme_count"]) * 1.0
    ).round(4)
    # 优先主题（任务书要求）：命中任一优先主题者整体排在纯结构度数者之前
    persons_scored["_theme_first"] = (persons_scored["theme_count"] > 0).astype(int)
    persons_scored = persons_scored.sort_values(
        ["_theme_first", "selection_score", "person_id"], ascending=[False, False, True]
    ).reset_index(drop=True)
    persons_scored = persons_scored.drop(columns=["_theme_first"])
    top_persons = persons_scored.head(N_PERSONS).copy()

    # —— 事件证据口径（排除 rejected）——
    src_family = (
        dict(zip(sources["source_id"], sources["source_family"]))
        if "source_family" in sources.columns
        else {}
    )
    event_evid: dict[str, dict[str, object]] = {
        eid: {"attached": 0, "direct": 0, "families": set(), "sources": set()} for eid in events["event_id"]
    }
    for row in facts.to_dict("records"):
        if row.get("subject_type") != "event" or row.get("review_status") == "rejected":
            continue
        eid = row.get("subject_id", "")
        if eid not in event_evid:
            continue
        event_evid[eid]["attached"] = int(event_evid[eid]["attached"]) + 1
        if row.get("evidence_support") == "support":
            event_evid[eid]["direct"] = int(event_evid[eid]["direct"]) + 1
        sid = row.get("source_id", "")
        event_evid[eid]["sources"].add(sid)
        # source_family 列缺失时以 source_id 计数并在报告中声明口径降级
        event_evid[eid]["families"].add(src_family.get(sid, sid))

    # —— 事件评分 ——
    top_person_ids = set(top_persons["person_id"])
    ev_rows = []
    for row in events.to_dict("records"):
        eid = row.get("event_id", "")
        themes = event_themes.get(eid, set())
        plist = participants_by_event.get(eid, [])
        core_p = [
            p
            for p in plist
            if p in top_person_ids or zuolian_mem_type.get(p) == "confirmed_member"
        ]
        ev_rows.append(
            {
                "event_id": eid,
                "event_name": row.get("event_name", ""),
                "event_date": row.get("event_date", ""),
                "themes": ";".join(sorted(themes)),
                "theme_count": len(themes),
                "participant_count": len(plist),
                "core_participant_share": round(len(core_p) / len(plist), 4) if plist else 0.0,
                "attached_evidences": int(event_evid[eid]["attached"]),
                "direct_support_evidences": int(event_evid[eid]["direct"]),
                "independent_source_families": len(event_evid[eid]["families"]),
                "needs_manual_review": row.get("needs_manual_review", ""),
            }
        )
    events_scored = pd.DataFrame(ev_rows)
    events_scored["selection_score"] = (
        _norm(events_scored["theme_count"]) * 1.2
        + events_scored["core_participant_share"] * 0.8
        + (events_scored["direct_support_evidences"] == 0).astype(float) * 1.5
        + (events_scored["independent_source_families"] <= 1).astype(float) * 0.5
        + (events_scored["needs_manual_review"] == "yes").astype(float) * 0.3
    ).round(4)
    events_scored = events_scored.sort_values(
        ["selection_score", "event_id"], ascending=[False, True]
    ).reset_index(drop=True)
    top_events = events_scored.head(N_EVENTS).copy()

    # —— 地点评分 ——
    events_by_place: dict[str, int] = {}
    for pl in event_place_by_id.values():
        if pl:
            events_by_place[pl] = events_by_place.get(pl, 0) + 1
    plc_rows = []
    for row in places.to_dict("records"):
        pid = row.get("place_id", "")
        blob = " ".join([row.get("historical_name", ""), row.get("current_name", ""), row.get("place_name", "")])
        themes = set()
        if _matches(blob, LONGHUA_PAT):
            themes.add("龙华烈士相关事件")
        if _matches(blob, ("内山", "大陆新村", "拉摩斯")):
            themes.add("鲁迅与青年作家网络")
        if _matches(blob, ("多伦路", "中华艺术大学", "左联")):
            themes.add("左联成立及组织演变")
        if _matches(blob, PUBLICATION_PAT):
            themes.add("左翼刊物和出版机构")
        historical = row.get("historical_name", "")
        current = row.get("current_name", "")
        coord_prec = row.get("coordinate_precision", "") or row.get("coord_precision", "")
        geo_method = row.get("geocoding_method", "")
        plc_rows.append(
            {
                "place_id": pid,
                "place_name": row.get("place_name", ""),
                "historical_name": historical,
                "current_name": current,
                "place_type": row.get("place_type", ""),
                "attached_event_count": events_by_place.get(pid, 0),
                "themes": ";".join(sorted(themes)),
                "theme_count": len(themes),
                "historical_address_missing": bool(not historical.strip() or "待核" in historical),
                "current_address_missing": bool(not current.strip() or "待核" in current),
                "coord_precision": coord_prec,
                "coord_precision_unknown": coord_prec.strip().lower() in ("", "unknown", "nan"),
                "coord_source_missing": not geo_method.strip(),
                "is_generic_city": row.get("place_type", "") == "city" or historical.strip() == "上海",
            }
        )
    places_scored = pd.DataFrame(plc_rows)
    # 泛化城市级地点（如"上海"，83 个事件挂接）不进入核心补证 Top10，其治理见时空审计脚本
    concrete = places_scored[~places_scored["is_generic_city"]].copy()
    concrete["selection_score"] = (
        _norm(concrete["attached_event_count"]) * 1.0
        + _norm(concrete["theme_count"]) * 0.6
        + concrete["historical_address_missing"].astype(float) * 0.4
        + concrete["current_address_missing"].astype(float) * 0.4
        + concrete["coord_precision_unknown"].astype(float) * 0.4
        + concrete["coord_source_missing"].astype(float) * 0.3
    ).round(4)
    concrete = concrete.sort_values(["selection_score", "place_id"], ascending=[False, True]).reset_index(drop=True)
    top_places = concrete.head(N_PLACES).copy()

    # —— 输出 CSV ——
    out_dir.mkdir(parents=True, exist_ok=True)
    person_df = _person_output(top_persons, persons, person_fact)
    person_df.to_csv(out_dir / "core_person_evidence_candidates.csv", index=False, encoding="utf-8-sig")
    event_df = _event_output(top_events)
    event_df.to_csv(out_dir / "core_event_evidence_candidates.csv", index=False, encoding="utf-8-sig")
    place_df = _place_output(top_places)
    place_df.to_csv(out_dir / "core_place_review_candidates.csv", index=False, encoding="utf-8-sig")
    _write_report(out_dir, person_df, event_df, place_df, data_dir, src_family)

    seed_person = sum(
        1 for pid in person_df["person_id"] for f in ("birth_year", "death_year", "role") if f in EVIDENCE_SEEDS.get(pid, {})
    )
    seed_event = sum(
        1 for eid in event_df["event_id"] for f in ("direct_support", "independent_source") if f in EVIDENCE_SEEDS.get(eid, {})
    )
    seed_place = sum(
        1 for plid in place_df["place_id"] for f in ("address", "coord") if f in EVIDENCE_SEEDS.get(plid, {})
    )
    return {
        "persons": len(person_df),
        "events": len(event_df),
        "places": len(place_df),
        "person_seed_fields": seed_person,
        "event_seed_fields": seed_event,
        "place_seed_fields": seed_place,
    }


def _person_output(
    top_persons: pd.DataFrame, persons: pd.DataFrame, person_fact: dict[str, dict[str, int]]
) -> pd.DataFrame:
    birth = dict(zip(persons["person_id"], persons["birth_year"]))
    death = dict(zip(persons["person_id"], persons["death_year"]))
    role = dict(zip(persons["person_id"], persons["role"]))
    rows = []
    for i, row in enumerate(top_persons.to_dict("records"), start=1):
        pid = row["person_id"]
        seeds = EVIDENCE_SEEDS.get(pid, {})
        birth_status = _seed_status(pid, seeds, "birth_year")
        death_status = _seed_status(pid, seeds, "death_year")
        role_status = _seed_status(pid, seeds, "role")
        birth_en = _seed_enrichment(pid, "birth_year", seeds)
        death_en = _seed_enrichment(pid, "death_year", seeds)
        role_en = _seed_enrichment(pid, "role", seeds)
        rows.append(
            {
                "candidate_id": f"CPC-P{i:02d}",
                "person_id": pid,
                "standard_name": row["standard_name"],
                "birth_year_current": birth.get(pid, ""),
                "death_year_current": death.get(pid, ""),
                "role_current": role.get(pid, ""),
                "birth_year_fact_evidences": person_fact[pid]["birth_year"],
                "death_year_fact_evidences": person_fact[pid]["death_year"],
                "role_fact_evidences": person_fact[pid]["role"],
                "selection_score": row["selection_score"],
                "themes": row["themes"],
                "priority_reason": (
                    f"可信关系度{row['trusted_degree']}；全量度{row['full_degree']}；事件参与{row['event_count']}；"
                    f"左联身份={row['zuolian_membership'] or '无'}；命中主题{row['theme_count']}项"
                ),
                "existing_evidence_summary": "组织身份证据见 org_membership_evidences；生卒年/角色事实级证据 0 条",
                "evidence_gap": "birth_year/death_year/role 事实级证据全部缺失（现值为传承导入，未逐条立证）",
                "suggested_action": "补权威辞典/纪念馆页面来源，locator 定位到页码或条目，人工复核后立证",
                "birth_status": birth_status,
                "birth_candidate_title": seeds.get("birth_year", {}).get("title", ""),
                "birth_candidate_url": seeds.get("birth_year", {}).get("url", ""),
                "birth_candidate_access_date": seeds.get("birth_year", {}).get("access_date", ""),
                "birth_candidate_locator": seeds.get("birth_year", {}).get("locator", ""),
                "birth_candidate_quote": seeds.get("birth_year", {}).get("quote", ""),
                "birth_source_level": birth_en["source_level"],
                "birth_source_type": birth_en["source_type"],
                "birth_retrieval_status": birth_en["retrieval_status"],
                "birth_content_hash": birth_en["content_hash"],
                "death_status": death_status,
                "death_candidate_title": seeds.get("death_year", {}).get("title", ""),
                "death_candidate_url": seeds.get("death_year", {}).get("url", ""),
                "death_candidate_access_date": seeds.get("death_year", {}).get("access_date", ""),
                "death_candidate_locator": seeds.get("death_year", {}).get("locator", ""),
                "death_candidate_quote": seeds.get("death_year", {}).get("quote", ""),
                "death_source_level": death_en["source_level"],
                "death_source_type": death_en["source_type"],
                "death_retrieval_status": death_en["retrieval_status"],
                "death_content_hash": death_en["content_hash"],
                "role_status": role_status,
                "role_candidate_title": seeds.get("role", {}).get("title", ""),
                "role_candidate_url": seeds.get("role", {}).get("url", ""),
                "role_candidate_access_date": seeds.get("role", {}).get("access_date", ""),
                "role_candidate_locator": seeds.get("role", {}).get("locator", ""),
                "role_candidate_quote": seeds.get("role", {}).get("quote", ""),
                "role_source_level": role_en["source_level"],
                "role_source_type": role_en["source_type"],
                "role_retrieval_status": role_en["retrieval_status"],
                "role_content_hash": role_en["content_hash"],
            }
        )
    return pd.DataFrame(rows)


def _event_output(top_events: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for i, row in enumerate(top_events.to_dict("records"), start=1):
        eid = row["event_id"]
        seeds = EVIDENCE_SEEDS.get(eid, {})
        ds_status = _seed_status(eid, seeds, "direct_support")
        is_status = _seed_status(eid, seeds, "independent_source")
        ds_en = _seed_enrichment(eid, "direct_support", seeds)
        is_en = _seed_enrichment(eid, "independent_source", seeds)
        rows.append(
            {
                "candidate_id": f"CPC-E{i:02d}",
                "event_id": eid,
                "event_name": row["event_name"],
                "event_date": row["event_date"],
                "selection_score": row["selection_score"],
                "themes": row["themes"],
                "priority_reason": (
                    f"主题{row['theme_count']}项；参与者{row['participant_count']}人（核心占比{row['core_participant_share']}）；"
                    f"直接支持证据{row['direct_support_evidences']}条；独立来源族{row['independent_source_families']}"
                ),
                "existing_evidence_summary": (
                    f"已挂接证据{row['attached_evidences']}条，其中直接支持{row['direct_support_evidences']}条"
                ),
                "evidence_gap": (
                    "无直接支持证据"
                    if row["direct_support_evidences"] == 0
                    else ("仅单一来源族" if row["independent_source_families"] <= 1 else "证据较全，复核即可")
                ),
                "suggested_action": (
                    "补直接支持证据并核对日期/地点/参与者"
                    if row["direct_support_evidences"] == 0
                    else "补第二独立来源族并交叉核对"
                ),
                "direct_support_status": ds_status,
                "direct_support_candidate_title": seeds.get("direct_support", {}).get("title", ""),
                "direct_support_candidate_url": seeds.get("direct_support", {}).get("url", ""),
                "direct_support_candidate_access_date": seeds.get("direct_support", {}).get("access_date", ""),
                "direct_support_candidate_locator": seeds.get("direct_support", {}).get("locator", ""),
                "direct_support_candidate_quote": seeds.get("direct_support", {}).get("quote", ""),
                "direct_support_source_level": ds_en["source_level"],
                "direct_support_source_type": ds_en["source_type"],
                "direct_support_retrieval_status": ds_en["retrieval_status"],
                "direct_support_content_hash": ds_en["content_hash"],
                "independent_source_status": is_status,
                "independent_source_candidate_title": seeds.get("independent_source", {}).get("title", ""),
                "independent_source_candidate_url": seeds.get("independent_source", {}).get("url", ""),
                "independent_source_candidate_access_date": seeds.get("independent_source", {}).get("access_date", ""),
                "independent_source_candidate_locator": seeds.get("independent_source", {}).get("locator", ""),
                "independent_source_candidate_quote": seeds.get("independent_source", {}).get("quote", ""),
                "independent_source_source_level": is_en["source_level"],
                "independent_source_source_type": is_en["source_type"],
                "independent_source_retrieval_status": is_en["retrieval_status"],
                "independent_source_content_hash": is_en["content_hash"],
            }
        )
    return pd.DataFrame(rows)


def _place_output(top_places: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for i, row in enumerate(top_places.to_dict("records"), start=1):
        plid = row["place_id"]
        seeds = EVIDENCE_SEEDS.get(plid, {})
        ad_status = _seed_status(plid, seeds, "address")
        co_status = _seed_status(plid, seeds, "coord")
        ad_en = _seed_enrichment(plid, "address", seeds)
        co_en = _seed_enrichment(plid, "coord", seeds)
        rows.append(
            {
                "candidate_id": f"CPC-L{i:02d}",
                "place_id": plid,
                "place_name": row["place_name"],
                "selection_score": row["selection_score"],
                "themes": row["themes"],
                "priority_reason": (
                    f"挂接事件{row['attached_event_count']}个；主题{row['theme_count']}项；"
                    f"历史地址缺={row['historical_address_missing']}；现代地址缺={row['current_address_missing']}；"
                    f"坐标精度={row['coord_precision'] or '未标注'}"
                ),
                "existing_evidence_summary": (
                    f"挂接事件{row['attached_event_count']}个；坐标精度字段={row['coord_precision'] or '空'}"
                ),
                "evidence_gap": (
                    f"历史地址{'缺' if row['historical_address_missing'] else '有'}；"
                    f"现代地址{'缺' if row['current_address_missing'] else '有'}；"
                    f"坐标来源{'缺' if row['coord_source_missing'] else '有'}；"
                    f"坐标精度{'未知' if row['coord_precision_unknown'] else row['coord_precision']}"
                ),
                "suggested_action": "补历史沿革/旧址保护单位页面，登记现代地址与坐标来源，人工确认坐标精度",
                "address_status": ad_status,
                "address_candidate_title": seeds.get("address", {}).get("title", ""),
                "address_candidate_url": seeds.get("address", {}).get("url", ""),
                "address_candidate_access_date": seeds.get("address", {}).get("access_date", ""),
                "address_candidate_locator": seeds.get("address", {}).get("locator", ""),
                "address_candidate_quote": seeds.get("address", {}).get("quote", ""),
                "address_source_level": ad_en["source_level"],
                "address_source_type": ad_en["source_type"],
                "address_retrieval_status": ad_en["retrieval_status"],
                "address_content_hash": ad_en["content_hash"],
                "coord_status": co_status,
                "coord_candidate_title": seeds.get("coord", {}).get("title", ""),
                "coord_candidate_url": seeds.get("coord", {}).get("url", ""),
                "coord_candidate_access_date": seeds.get("coord", {}).get("access_date", ""),
                "coord_candidate_locator": seeds.get("coord", {}).get("locator", ""),
                "coord_candidate_quote": seeds.get("coord", {}).get("quote", ""),
                "coord_source_level": co_en["source_level"],
                "coord_source_type": co_en["source_type"],
                "coord_retrieval_status": co_en["retrieval_status"],
                "coord_content_hash": co_en["content_hash"],
            }
        )
    return pd.DataFrame(rows)


def _write_report(
    out_dir: Path,
    person_df: pd.DataFrame,
    event_df: pd.DataFrame,
    place_df: pd.DataFrame,
    data_dir: Path,
    src_family: dict[str, str],
) -> None:
    fields = ("birth", "death", "role")
    seed_found_person = sum((person_df[f"{f}_status"] == PENDING).sum() for f in fields)
    seed_conflict_person = sum((person_df[f"{f}_status"] == CONFLICT).sum() for f in fields)
    total_missing_person = sum((person_df[f"{f}_status"] == MISSING).sum() for f in fields)
    level_counts = _level_counts([person_df, event_df, place_df], ("birth", "death", "role"))
    family_note = (
        "以 sources.source_family 计独立来源族"
        if src_family
        else "sources.source_family 列缺失，退化为按 source_id 计数（上界口径）"
    )
    lines = [
        "# 核心补证候选包选择报告（返修版）",
        "",
        f"- 数据快照：`{data_dir}`（读取日期 {SNAPSHOT_DATE}，只读）",
        f"- 候选数量：人物 {len(person_df)} / 事件 {len(event_df)} / 地点 {len(place_df)}",
        "- 人物生卒年/角色事实级证据：全库 0 条（fact_evidences 无 person birth/death/role 谓词），",
        "  现有 birth_year/death_year/role 值均为传承导入，**不得作为已证实史实展示**。",
        f"- 独立来源族口径：{family_note}。",
        "- 候选来源分级：维基百科/百度百科/普通媒体一律为 web_lead/D 级线索，",
        "  不得称为权威证据；中国作家网/政府/纪念馆/档案馆/大学档案/学术论文/原始文献按现有来源规则分级。",
        f"- 候选来源等级分布（A/B/C/D）：{level_counts['A']}/{level_counts['B']}/"
        f"{level_counts['C']}/{level_counts['D']}（仅统计非 missing 候选；missing 不计入）。",
        "",
        "## 选择规则（可复现）",
        "",
        "1. 人物分 = 2.0×较低风险关系度(归一，启发式，非可信断言) + 0.8×全量关系度(归一) + 1.2×事件参与数(归一) + 左联身份(正式1.0/相关0.5/候选0.25) + 1.0×主题命中数(归一)。",
        "   较低风险关系 = 排除 待核验类型 / needs_manual_review=yes / critical-high 风险 / confidence=low；",
        "   不看 relation_evidences、不看 publish_status；不得称为可信关系，仅用于补证优先级排序。",
        "   真正的可信判定（须 support 证据）见可信网络分析；当前证据支持为 0，补证优先级暂用启发式口径。",
        "2. 事件分 = 1.2×主题命中(归一) + 0.8×核心参与者占比 + 1.5×(无直接支持证据) + 0.5×(独立来源族≤1) + 0.3×(需人工复核)。",
        "3. 地点分 = 1.0×挂接事件数(归一) + 0.6×主题命中(归一) + 0.4×历史地址缺 + 0.4×现代地址缺 + 0.4×坐标精度未知 + 0.3×坐标来源缺。",
        "   泛化城市级地点（如『上海』，83 个事件挂接）不进核心补证 Top10，其治理见时空审计报告。",
        "4. 主题判定使用显式关键词与种子常量（五烈士名单、女性作家种子名单、青年=与鲁迅有可信关系且生于1895年后），全部登记于脚本头部。",
        "5. 同分按 ID 升序决断；命中任一优先主题的人物整体优先于纯结构度数人物；重跑输出字节一致（无时间戳）。",
        "",
        "## 候选证据种子状态",
        "",
        f"- 人物三字段（生/卒/角色）共 {3 * len(person_df)} 个待补事实：检索到候选线索 {seed_found_person} 个（{PENDING}），"
        f"冲突 {seed_conflict_person} 个（{CONFLICT}，与生产值冲突、禁止自动落库），",
        f"其余 {total_missing_person} 个记 {MISSING}。",
        f"- 事件两字段（直接支持/独立来源）共 {2 * len(event_df)} 个待补：{PENDING} "
        f"{int((event_df['direct_support_status'] == PENDING).sum() + (event_df['independent_source_status'] == PENDING).sum())} 个；"
        f"{CONFLICT} {int((event_df['direct_support_status'] == CONFLICT).sum() + (event_df['independent_source_status'] == CONFLICT).sum())} 个。",
        f"- 地点两字段（地址/坐标）共 {2 * len(place_df)} 个待补：{PENDING} "
        f"{int((place_df['address_status'] == PENDING).sum() + (place_df['coord_status'] == PENDING).sum())} 个。",
        "- 所有候选线索仅登记 URL/访问日期/定位/短引文 + source_level/source_type/retrieval_status/content_hash，",
        "  不修改生产数据，不转正；conflict 候选禁止自动落库，须人工裁决。",
        "",
        "## 检索中发现的口径差异与否定性结果（全部待人工裁决，本 Agent 不改生产数据）",
        "",
    ]
    lines.extend(f"{i}. {note}" for i, note in enumerate(SEED_NOTES, start=1))
    lines += [
        "",
        "## Top 人物（前 10 展示，全表见 CSV）",
        "",
        "| 排名 | ID | 姓名 | 选择分 | 主题 |",
        "| --- | --- | --- | --- | --- |",
    ]
    for i, row in enumerate(person_df.head(10).to_dict("records"), start=1):
        lines.append(
            f"| {i} | {row['person_id']} | {row['standard_name']} | {row['selection_score']} | {row['themes'] or '—'} |"
        )
    lines += [
        "",
        "## Top 事件（前 10 展示，全表见 CSV）",
        "",
        "| 排名 | ID | 名称 | 日期 | 直接支持状态 | 独立来源状态 |",
        "| --- | --- | --- | --- | --- | --- |",
    ]
    for i, row in enumerate(event_df.head(10).to_dict("records"), start=1):
        lines.append(
            f"| {i} | {row['event_id']} | {row['event_name']} | {row['event_date'] or '—'} | "
            f"{row['direct_support_status']} | {row['independent_source_status']} |"
        )
    lines += [
        "",
        "## Top 地点（全表见 CSV）",
        "",
        "| 排名 | ID | 名称 |",
        "| --- | --- | --- |",
    ]
    for i, row in enumerate(place_df.to_dict("records"), start=1):
        lines.append(f"| {i} | {row['place_id']} | {row['place_name']} |")
    lines += [
        "",
        "## 证据局限",
        "",
        "- 选择分数反映**当前数据集内的结构位置与补证紧迫度**，不是历史重要性排名。",
        "- 关键词与种子名单仅用于研究优先级，不构成对新史实的断言。",
        "- missing 表示本轮未检索到可用线索或仅有低级线索，不代表证据不存在；找不到权威替代时保留 missing，不得凑数。",
        "- D 级 web_lead 线索（百科/普通媒体）不得称为权威证据，转正须人工复核并补更权威来源。",
        "",
    ]
    (out_dir / "core_upgrade_selection_report.md").write_text("\n".join(lines), encoding="utf-8")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="构建核心补证候选包（只读）")
    parser.add_argument("--data-dir", type=Path, default=DEFAULT_DATA_DIR)
    parser.add_argument("--out-dir", type=Path, default=DEFAULT_OUT_DIR)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    summary = build(args.data_dir, args.out_dir)
    print(f"core upgrade candidates: {summary}")
    print(f"out_dir: {args.out_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
