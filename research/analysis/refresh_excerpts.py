from __future__ import annotations

import csv
import json
import re
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path

import fitz
from rapidocr_onnxruntime import RapidOCR

PAGE_SOURCES = {"左联史", "左联词典"}

RELATION_HINTS = {
    "通信": ("信", "致", "复", "函"),
    "交往": ("来", "访", "会", "见", "同"),
    "交游": ("来", "访", "会", "见", "同"),
    "同属组织": ("左联", "盟员", "成员", "发起人", "参加"),
    "组织隶属": ("左联", "盟员", "成员", "发起人", "参加"),
}

MONTH_MAP = {
    "正": 1,
    "元": 1,
    "一": 1,
    "二": 2,
    "三": 3,
    "四": 4,
    "五": 5,
    "六": 6,
    "七": 7,
    "八": 8,
    "九": 9,
    "十": 10,
    "十一": 11,
    "冬": 11,
    "十二": 12,
    "腊": 12,
}

DAY_DIGITS = {
    "零": 0,
    "〇": 0,
    "○": 0,
    "一": 1,
    "二": 2,
    "三": 3,
    "四": 4,
    "五": 5,
    "六": 6,
    "七": 7,
    "八": 8,
    "九": 9,
}

DAY_PATTERN = re.compile(r"^([一二三四五六七八九十廿卅〇○零元正]+)日[　 ]?(.*)$")
YEAR_PATTERN = re.compile(r"^日记.*[（(](\d{4})年[）)]$")
PAGE_PATTERN = re.compile(r"第(\d+)页")
DATE_PATTERN = re.compile(r"(\d{4})年(\d{1,2})月(\d{1,2})日")


@dataclass(frozen=True)
class Citation:
    source: str
    locator: str


@dataclass
class ProjectPaths:
    root: Path
    kb_dir: Path
    app_dir: Path
    raw_text_dir: Path
    extracted_dir: Path
    person_relations_csv: Path
    event_evidences_json: Path
    relation_evidence_mock_json: Path
    persons_csv: Path
    event_participants_csv: Path
    events_csv: Path
    diary_txt: Path
    shi_ocr_json: Path
    cidian_ocr_json: Path
    memoir_pdf: Path
    memoir_cache_json: Path

    @classmethod
    def discover(cls, root: Path) -> ProjectPaths:
        research_dir = root / "research"
        raw_text_dir = research_dir / "raw_texts"
        extracted_dir = research_dir / "intermediate" / "extracted"
        kb_dir = root / "data" / "processed"
        app_dir = root / "app" / "frontend"
        parent_pdfs = list(root.parent.glob("*.pdf"))

        return cls(
            root=root,
            kb_dir=kb_dir,
            app_dir=app_dir,
            raw_text_dir=raw_text_dir,
            extracted_dir=extracted_dir,
            person_relations_csv=kb_dir / "person_relations.csv",
            event_evidences_json=kb_dir / "event_evidences.json",
            relation_evidence_mock_json=app_dir / "relation_evidence_mock.json",
            persons_csv=kb_dir / "persons.csv",
            event_participants_csv=kb_dir / "event_participants.csv",
            events_csv=kb_dir / "events.csv",
            diary_txt=next(p for p in raw_text_dir.iterdir() if p.name.startswith("日记全编")),
            shi_ocr_json=next(p for p in extracted_dir.glob("*_ocr_text.json") if "左联史" in p.name),
            cidian_ocr_json=next(p for p in extracted_dir.glob("*_ocr_text.json") if "左联词典" in p.name),
            memoir_pdf=next(p for p in parent_pdfs if "回忆录" in p.name),
            memoir_cache_json=extracted_dir / "左联回忆录_ocr_text.json",
        )


def unique_preserve_order(values: list[str]) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []
    for value in values:
        value = value.strip()
        if not value or value in seen:
            continue
        seen.add(value)
        result.append(value)
    return result


def split_aliases(value: str) -> list[str]:
    if not value:
        return []
    parts = re.split(r"[、,，;；/ ]+", value)
    return [part.strip() for part in parts if len(part.strip()) >= 2]


def chinese_number_to_int(token: str) -> int:
    if token in DAY_DIGITS:
        return DAY_DIGITS[token]
    if token == "十":
        return 10
    if token.startswith("十"):
        return 10 + DAY_DIGITS.get(token[1:], 0)
    if token.startswith("廿"):
        suffix = token[1:]
        return 20 + DAY_DIGITS.get(suffix, 0)
    if token.startswith("卅"):
        suffix = token[1:]
        return 30 + DAY_DIGITS.get(suffix, 0)
    if "十" in token:
        head, tail = token.split("十", 1)
        return DAY_DIGITS.get(head, 0) * 10 + DAY_DIGITS.get(tail, 0)
    return 0


def parse_month(line: str) -> int | None:
    if not line.endswith("月"):
        return None
    return MONTH_MAP.get(line[:-1])


def clean_diary_entry(text: str) -> str:
    text = re.sub(r"\s+", "", text)
    return text.strip("。；，, ")


def parse_diary_entries(path: Path) -> dict[str, str]:
    lines = path.read_text(encoding="utf-8").splitlines()
    entries: dict[str, str] = {}
    current_year: int | None = None
    current_month: int | None = None
    current_key: str | None = None
    current_lines: list[str] = []
    skip_account = False

    def flush() -> None:
        nonlocal current_key, current_lines
        if current_key and current_lines:
            entries[current_key] = clean_diary_entry("".join(current_lines))
        current_key = None
        current_lines = []

    for raw_line in lines:
        line = raw_line.strip()
        if not line:
            continue
        year_match = YEAR_PATTERN.match(line)
        if year_match:
            flush()
            current_year = int(year_match.group(1))
            current_month = None
            skip_account = False
            continue
        if line in {"书帐", "居帐", "西牖书钞"}:
            flush()
            skip_account = True
            continue
        month = parse_month(line)
        if month is not None:
            flush()
            current_month = month
            skip_account = False
            continue
        if skip_account or current_year is None or current_month is None:
            continue
        day_match = DAY_PATTERN.match(line)
        if day_match:
            flush()
            day = chinese_number_to_int(day_match.group(1))
            if day:
                current_key = f"{current_year}年{current_month}月{day}日"
                current_lines = [day_match.group(2).strip()]
            continue
        if current_key:
            current_lines.append(line)
    flush()
    return entries


def normalize_ocr_line(line: str) -> str:
    line = line.replace("\u3000", " ")
    line = line.replace("_", "").replace("`", "")
    line = re.sub(r"[─═—]{2,}", " ", line)
    line = re.sub(r"(?<=[\u4e00-\u9fff0-9])\s+(?=[\u4e00-\u9fff0-9])", "", line)
    line = re.sub(r"(?<=[\u4e00-\u9fff0-9])\s+(?=[，。、“”‘’；：！？（）《》〈〉、,.!?:;])", "", line)
    line = re.sub(r"(?<=[，。、“”‘’；：！？（）《》〈〉、,.!?:;])\s+(?=[\u4e00-\u9fff0-9])", "", line)
    line = re.sub(r"\s+", " ", line).strip()
    return line


def is_noise_line(line: str) -> bool:
    compact = line.replace(" ", "")
    if compact in {"封面", "书名", "版权", "前言", "目录"}:
        return True
    if re.fullmatch(r"[0-9]+", compact):
        return True
    if any(title in compact for title in ("左联词典", "亡联词典", "左联史")) and len(compact) <= 18:
        return True
    if "第" in compact and "章" in compact and len(compact) <= 28 and "。" not in compact:
        return True
    if len(compact) <= 2:
        return True
    return False


def normalize_ocr_text(text: str) -> list[str]:
    lines: list[str] = []
    for raw_line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        line = normalize_ocr_line(raw_line)
        if not line or is_noise_line(line):
            continue
        lines.append(line)
    return lines


def trim_excerpt(text: str, max_chars: int) -> str:
    text = text.strip("；;，, ")
    if len(text) <= max_chars:
        return text
    return text[: max_chars - 1].rstrip("；;，, ") + "…"


def count_bad_chars(text: str) -> int:
    return len(re.findall(r"[_~`^|\\[\]{}<>]+", text))


def build_windows(lines: list[str], centers: list[int]) -> list[str]:
    candidates: list[str] = []
    seen: set[str] = set()
    if not lines:
        return candidates
    if not centers:
        centers = list(range(min(len(lines), 8)))
    for center in centers:
        for before, after in ((0, 0), (0, 1), (1, 1), (1, 2), (2, 2)):
            start = max(0, center - before)
            end = min(len(lines), center + after + 1)
            snippet = "".join(lines[start:end])
            snippet = trim_excerpt(snippet, 220)
            if snippet and snippet not in seen:
                seen.add(snippet)
                candidates.append(snippet)
    return candidates


def score_snippet(
    text: str,
    key_groups: list[list[str]],
    bonus_keywords: list[str],
    hint_words: list[str],
) -> int:
    score = 0
    matched_groups = 0
    matched_bonus = 0
    for group in key_groups:
        group_hits = sum(1 for keyword in group if keyword and keyword in text)
        if group_hits:
            matched_groups += 1
            score += 12 + min(group_hits, 3) * 2
    if matched_groups > 1:
        score += 12
    for keyword in bonus_keywords:
        if keyword and keyword in text:
            matched_bonus += 1
    score += min(matched_bonus, 6) * 3
    for hint in hint_words:
        if hint in text:
            score += 2
    if 20 <= len(text) <= 160:
        score += 6
    elif len(text) <= 220:
        score += 2
    score -= count_bad_chars(text) * 2
    return score


def choose_excerpt_with_score(
    lines: list[str],
    key_groups: list[list[str]],
    bonus_keywords: list[str],
    hint_words: list[str],
    max_chars: int,
) -> tuple[str, int]:
    keyword_pool = unique_preserve_order(
        [keyword for group in key_groups for keyword in group] + bonus_keywords + hint_words
    )
    centers = [index for index, line in enumerate(lines) if any(keyword in line for keyword in keyword_pool)]
    candidates = build_windows(lines, centers)
    if not candidates:
        candidates = build_windows(lines, [])
    if not candidates:
        return "", 0
    filtered = [snippet for snippet in candidates if count_bad_chars(snippet) < 3]
    if filtered:
        candidates = filtered
    scored = [
        (trim_excerpt(snippet, max_chars), score_snippet(snippet, key_groups, bonus_keywords, hint_words))
        for snippet in candidates
    ]
    best_snippet, best_score = max(scored, key=lambda item: (item[1], -len(item[0])))
    return best_snippet, best_score


def parse_citation_ref(value: str) -> Citation | None:
    value = value.strip()
    if not value:
        return None
    parts = value.split(maxsplit=1)
    if len(parts) != 2:
        return None
    return Citation(source=parts[0], locator=parts[1].strip())


class MemoirOcrCache:
    def __init__(self, pdf_path: Path, cache_path: Path) -> None:
        self.pdf_path = pdf_path
        self.cache_path = cache_path
        self.cache = json.loads(cache_path.read_text(encoding="utf-8")) if cache_path.exists() else {}
        self._doc: fitz.Document | None = None
        self._ocr: RapidOCR | None = None

    def get_page_text(self, page_number: int) -> str:
        key = str(page_number)
        if key in self.cache:
            return self.cache[key]
        if self._doc is None:
            self._doc = fitz.open(self.pdf_path)
        if self._ocr is None:
            self._ocr = RapidOCR()
        page = self._doc.load_page(page_number - 1)
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
        result, _ = self._ocr(pix.tobytes("png"))
        text = "\n".join(line[1] for line in result) if result else ""
        self.cache[key] = text
        return text

    def save(self) -> None:
        self.cache_path.write_text(json.dumps(self.cache, ensure_ascii=False, indent=2), encoding="utf-8")


class CitationLibrary:
    def __init__(self, paths: ProjectPaths) -> None:
        self.paths = paths
        self.diary_entries = parse_diary_entries(paths.diary_txt)
        self.book_pages = {
            "左联史": json.loads(paths.shi_ocr_json.read_text(encoding="utf-8")),
            "左联词典": json.loads(paths.cidian_ocr_json.read_text(encoding="utf-8")),
        }
        self.memoir_cache = MemoirOcrCache(paths.memoir_pdf, paths.memoir_cache_json)
        self._line_cache: dict[tuple[str, str], list[str]] = {}

    def save(self) -> None:
        self.memoir_cache.save()

    def get_lines(self, citation: Citation) -> list[str]:
        cache_key = (citation.source, citation.locator)
        if cache_key in self._line_cache:
            return self._line_cache[cache_key]
        if citation.source == "鲁迅日记":
            entry = self.diary_entries.get(citation.locator, "")
            lines = [entry] if entry else []
        elif citation.source in PAGE_SOURCES:
            page_match = PAGE_PATTERN.search(citation.locator)
            if not page_match:
                lines = []
            else:
                page_number = int(page_match.group(1))
                lines = normalize_ocr_text(self.book_pages[citation.source].get(str(page_number), ""))
        elif citation.source == "左联回忆录":
            page_match = PAGE_PATTERN.search(citation.locator)
            if not page_match:
                lines = []
            else:
                page_number = int(page_match.group(1))
                lines = normalize_ocr_text(self.memoir_cache.get_page_text(page_number))
        else:
            lines = []
        self._line_cache[cache_key] = lines
        return lines


def load_person_keywords(path: Path) -> dict[str, list[str]]:
    keywords: dict[str, list[str]] = {}
    with path.open("r", encoding="utf-8-sig", newline="") as file:
        for row in csv.DictReader(file):
            names = [row["standard_name"]] + split_aliases(row.get("aliases", ""))
            keywords[row["person_id"]] = unique_preserve_order(names)
    return keywords


def load_event_metadata(events_csv: Path, participants_csv: Path) -> tuple[dict[str, dict[str, str]], dict[str, list[str]]]:
    events: dict[str, dict[str, str]] = {}
    with events_csv.open("r", encoding="utf-8-sig", newline="") as file:
        for row in csv.DictReader(file):
            events[row["event_id"]] = row

    participants: dict[str, list[str]] = defaultdict(list)
    with participants_csv.open("r", encoding="utf-8-sig", newline="") as file:
        for row in csv.DictReader(file):
            participants[row["event_id"]].append(row["participant_name"])
    for event_id, values in participants.items():
        participants[event_id] = unique_preserve_order(values)
    return events, participants


def relation_keywords(row: dict[str, str], person_keywords: dict[str, list[str]]) -> tuple[list[list[str]], list[str], list[str]]:
    source_names = person_keywords.get(row["source_person_id"], [])
    target_names = person_keywords.get(row["target_person_id"], [])
    key_groups = [source_names[:4], target_names[:4]]
    bonus_keywords = unique_preserve_order(source_names[4:] + target_names[4:])
    hint_words = list(RELATION_HINTS.get(row.get("final_relation_type", ""), ()))
    return key_groups, bonus_keywords, hint_words


def event_keywords(
    event_id: str,
    events: dict[str, dict[str, str]],
    participants: dict[str, list[str]],
) -> tuple[list[list[str]], list[str], list[str]]:
    event = events.get(event_id, {})
    event_name = event.get("event_name", "")
    event_variants = [event_name]
    for suffix in ("事件", "活动", "大会", "斗争", "成立", "创办", "建立", "被捕", "营救"):
        if event_name.endswith(suffix) and len(event_name) > len(suffix) + 1:
            event_variants.append(event_name[: -len(suffix)])
    for phrase in ("中国左翼作家联盟", "左翼作家联盟", "左联", "成立大会", "会面", "成立"):
        if phrase in event_name:
            event_variants.append(phrase)
    if "与" in event_name:
        event_variants.extend(part for part in event_name.split("与") if len(part) >= 2)
    if "左联" in event_name and "左联" not in event_variants:
        event_variants.append("左联")
    participant_names = participants.get(event_id, [])[:10]
    hint_words: list[str] = []
    full_date = event.get("event_date", "")
    date_match = DATE_PATTERN.search(full_date)
    if date_match:
        year, month, day = date_match.groups()
        hint_words.extend([f"{year}年", f"{year}年{int(month)}月{int(day)}日", f"{int(month)}月{int(day)}日"])
    elif re.match(r"(\d{4})", full_date):
        hint_words.append(f"{full_date[:4]}年")
    return [unique_preserve_order(event_variants)], participant_names, hint_words


def build_relation_context(
    row: dict[str, str],
    person_keywords: dict[str, list[str]],
    library: CitationLibrary,
) -> str:
    key_groups, bonus_keywords, hint_words = relation_keywords(row, person_keywords)
    scored_snippets: list[tuple[int, int, str]] = []
    for index, chunk in enumerate(row["evidence_ref"].split(";")):
        citation = parse_citation_ref(chunk)
        if not citation:
            continue
        lines = library.get_lines(citation)
        if not lines:
            continue
        if citation.source == "鲁迅日记":
            snippet = trim_excerpt(lines[0], 120)
            score = score_snippet(snippet, key_groups, bonus_keywords, hint_words) + 10
        else:
            snippet, score = choose_excerpt_with_score(lines, key_groups, bonus_keywords, hint_words, 120)
        if snippet:
            scored_snippets.append((score, index, snippet))
    if not scored_snippets:
        return ""
    scored_snippets.sort(key=lambda item: (-item[0], item[1]))
    selected: list[tuple[int, str]] = []
    seen: set[str] = set()
    min_score = max(scored_snippets[0][0] - 10, 10)
    for score, index, snippet in scored_snippets:
        if snippet in seen or score < min_score:
            continue
        seen.add(snippet)
        selected.append((index, snippet))
        if len(selected) >= 3:
            break
    if not selected:
        selected = [(scored_snippets[0][1], scored_snippets[0][2])]
    selected.sort(key=lambda item: item[0])
    return "；".join(snippet for _, snippet in selected)


def build_event_quote(
    item: dict[str, str],
    events: dict[str, dict[str, str]],
    participants: dict[str, list[str]],
    library: CitationLibrary,
) -> str:
    citation = Citation(source=item["source"], locator=item["source_loc"])
    lines = library.get_lines(citation)
    if not lines:
        return ""
    key_groups, bonus_keywords, hint_words = event_keywords(item["event_id"], events, participants)
    if citation.source == "鲁迅日记":
        return trim_excerpt(lines[0], 180)
    snippet, score = choose_excerpt_with_score(lines, key_groups, bonus_keywords, hint_words, 180)
    if score >= 12:
        return snippet
    fallback_snippet, _ = choose_excerpt_with_score(lines, [bonus_keywords[:6]], bonus_keywords, hint_words, 180)
    return fallback_snippet


def build_mock_excerpt(
    evidence: dict[str, str],
    related_names: list[str],
    library: CitationLibrary,
) -> str:
    citation = parse_citation_ref(evidence.get("citation_ref", ""))
    if not citation:
        return ""
    lines = library.get_lines(citation)
    if not lines:
        return ""
    if citation.source == "鲁迅日记":
        return trim_excerpt(lines[0], 120)
    key_groups = [related_names]
    snippet, _ = choose_excerpt_with_score(lines, key_groups, related_names, [], 120)
    return snippet


def main() -> int:
    root = Path(__file__).resolve().parents[2]
    paths = ProjectPaths.discover(root)
    library = CitationLibrary(paths)
    person_keywords = load_person_keywords(paths.persons_csv)
    events, participants = load_event_metadata(paths.events_csv, paths.event_participants_csv)

    changed_relation_rows = 0
    relation_rows: list[dict[str, str]] = []
    with paths.person_relations_csv.open("r", encoding="utf-8-sig", newline="") as file:
        for row in csv.DictReader(file):
            new_context = build_relation_context(row, person_keywords, library)
            if new_context and row["context"] != new_context:
                row["context"] = new_context
                changed_relation_rows += 1
            relation_rows.append(row)

    with paths.person_relations_csv.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=list(relation_rows[0].keys()))
        writer.writeheader()
        writer.writerows(relation_rows)

    changed_event_rows = 0
    event_evidences = json.loads(paths.event_evidences_json.read_text(encoding="utf-8"))
    for item in event_evidences:
        new_quote = build_event_quote(item, events, participants, library)
        if new_quote and item.get("quote") != new_quote:
            item["quote"] = new_quote
            changed_event_rows += 1
    paths.event_evidences_json.write_text(
        json.dumps(event_evidences, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    changed_mock_rows = 0
    relation_mock = json.loads(paths.relation_evidence_mock_json.read_text(encoding="utf-8"))
    for item in relation_mock:
        related_names = unique_preserve_order(
            person_keywords.get(item.get("person_a_id", ""), [])[:4]
            + person_keywords.get(item.get("person_b_id", ""), [])[:4]
        )
        for evidence in item.get("evidence_samples", []):
            new_excerpt = build_mock_excerpt(evidence, related_names, library)
            if new_excerpt and evidence.get("excerpt") != new_excerpt:
                evidence["excerpt"] = new_excerpt
                changed_mock_rows += 1
    paths.relation_evidence_mock_json.write_text(
        json.dumps(relation_mock, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    library.save()

    print(f"Updated relation contexts: {changed_relation_rows}")
    print(f"Updated event quotes: {changed_event_rows}")
    print(f"Updated mock excerpts: {changed_mock_rows}")
    print(f"Diary entries loaded: {len(library.diary_entries)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
