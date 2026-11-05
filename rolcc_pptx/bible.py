# -*- coding: utf-8 -*-
"""Fetch NIV (English) and CUV (Traditional Chinese) verse text."""
from __future__ import annotations

import re
import time
from dataclasses import dataclass
from typing import Callable, Dict, List, Optional, Tuple
from urllib.parse import quote

import requests
from bs4 import BeautifulSoup

UA = {"User-Agent": "Mozilla/5.0 ROLCC-PPTX/1.0"}
BG_UA = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/122.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
}
_BG_VERSE_CLASS = re.compile(r"^.+-(\d+)-(\d+)$")

# Chinese short book names / abbreviations → (English name, bolls book id)
BOOK_MAP = {
    "創": ("Genesis", 1),
    "創世記": ("Genesis", 1),
    "出": ("Exodus", 2),
    "出埃及記": ("Exodus", 2),
    "利": ("Leviticus", 3),
    "民": ("Numbers", 4),
    "申": ("Deuteronomy", 5),
    "書": ("Joshua", 6),
    "約書亞記": ("Joshua", 6),
    "士": ("Judges", 7),
    "得": ("Ruth", 8),
    "撒上": ("1 Samuel", 9),
    "撒下": ("2 Samuel", 10),
    "王上": ("1 Kings", 11),
    "王下": ("2 Kings", 12),
    "代上": ("1 Chronicles", 13),
    "代下": ("2 Chronicles", 14),
    "拉": ("Ezra", 15),
    "尼": ("Nehemiah", 16),
    "斯": ("Esther", 17),
    "伯": ("Job", 18),
    "詩": ("Psalms", 19),
    "詩篇": ("Psalms", 19),
    "箴": ("Proverbs", 20),
    "傳": ("Ecclesiastes", 21),
    "歌": ("Song of Solomon", 22),
    "賽": ("Isaiah", 23),
    "以賽亞書": ("Isaiah", 23),
    "耶": ("Jeremiah", 24),
    "哀": ("Lamentations", 25),
    "結": ("Ezekiel", 26),
    "但": ("Daniel", 27),
    "何": ("Hosea", 28),
    "珥": ("Joel", 29),
    "摩": ("Amos", 30),
    "俄": ("Obadiah", 31),
    "拿": ("Jonah", 32),
    "彌": ("Micah", 33),
    "鴻": ("Nahum", 34),
    "哈": ("Habakkuk", 35),
    "番": ("Zephaniah", 36),
    "該": ("Haggai", 37),
    "亞": ("Zechariah", 38),
    "瑪": ("Malachi", 39),
    "太": ("Matthew", 40),
    "馬太福音": ("Matthew", 40),
    "可": ("Mark", 41),
    "馬可": ("Mark", 41),
    "馬可福音": ("Mark", 41),
    "路": ("Luke", 42),
    "路加福音": ("Luke", 42),
    "約": ("John", 43),
    "約翰福音": ("John", 43),
    "徒": ("Acts", 44),
    "羅": ("Romans", 45),
    "羅馬書": ("Romans", 45),
    "林前": ("1 Corinthians", 46),
    "哥林多前書": ("1 Corinthians", 46),
    "林後": ("2 Corinthians", 47),
    "加": ("Galatians", 48),
    "弗": ("Ephesians", 49),
    "以弗所書": ("Ephesians", 49),
    "腓": ("Philippians", 50),
    "腓立比書": ("Philippians", 50),
    "西": ("Colossians", 51),
    "帖前": ("1 Thessalonians", 52),
    "帖後": ("2 Thessalonians", 53),
    "提前": ("1 Timothy", 54),
    "提後": ("2 Timothy", 55),
    "多": ("Titus", 56),
    "門": ("Philemon", 57),
    "來": ("Hebrews", 58),
    "雅": ("James", 59),
    "雅各書": ("James", 59),
    "彼前": ("1 Peter", 60),
    "彼後": ("2 Peter", 61),
    "約壹": ("1 John", 62),
    "約貳": ("2 John", 63),
    "約叁": ("3 John", 64),
    "猶": ("Jude", 65),
    "啟": ("Revelation", 66),
}


@dataclass
class VerseRef:
    book_en: str
    book_id: int
    chapter: int
    start: int
    end: int
    raw: str

    def display_en(self) -> str:
        if self.start == self.end:
            return f"{self.book_en} {self.chapter}:{self.start}"
        return f"{self.book_en} {self.chapter}:{self.start}–{self.end}"

    def verse_numbers(self) -> List[int]:
        return list(range(self.start, self.end + 1))


def parse_scripture_ref(text: str) -> Optional[VerseRef]:
    """Parse refs like 可7：24-37 / 羅13:1-2 / 詩115:1 / 可8:34."""
    if not text:
        return None
    t = text.strip()
    t = t.replace("經文：", "").replace("經文:", "").strip()
    # Prefer longest book key match
    book_keys = sorted(BOOK_MAP.keys(), key=len, reverse=True)
    book = None
    rest = None
    for key in book_keys:
        if t.startswith(key):
            book = key
            rest = t[len(key) :].lstrip()
            break
    if not book:
        m = re.search(
            r"(可|羅|詩|詩篇|林前|林後|腓|弗|雅|賽|書|太|路|約|徒|西|加|來)(\d+)",
            t,
        )
        if not m:
            return None
        book = m.group(1)
        rest = t[m.start(2) :]
    book_en, book_id = BOOK_MAP[book]
    m = re.search(
        r"(\d+)\s*[:：]\s*(\d+)(?:\s*[-–—~至到]\s*(\d+))?",
        rest or t,
    )
    if not m:
        return None
    chapter = int(m.group(1))
    start = int(m.group(2))
    end = int(m.group(3) or start)
    return VerseRef(book_en, book_id, chapter, start, end, text.strip())


def _strip_html(s: str) -> str:
    s = re.sub(r"<br\s*/?>", " ", s, flags=re.I)
    s = re.sub(r"<[^>]+>", "", s)
    return re.sub(r"\s+", " ", s).strip()


def _strip_niv_heading(text: str) -> str:
    """Remove section titles glued onto verse text (bolls.life NIV)."""
    openers = (
        "Jesus", "Then", "When", "During", "Now", "After", "So", "But", "And",
        "He", "She", "They", "The", "A", "An", "In", "On", "As", "For",
        "While", "His", "Her", "At", "From", "One", "Some", "There",
    )
    # Prefer cutting right before a narrative opener when a short title precedes it
    # e.g. "Jesus Feeds the Four Thousand During those days..."
    for opener in openers:
        m = re.match(
            rf"^([A-Z][A-Za-z0-9'’:&;\-]*(?:\s+[A-Za-z0-9'’:&;\-]+){{1,10}})\s+({re.escape(opener)}\b.*)$",
            text,
        )
        if m:
            heading = m.group(1).strip()
            rest = m.group(2).strip()
            if "." not in heading and 3 <= len(heading) <= 70 and heading.split()[-1] != opener:
                return rest
    # Glued camel-case boundary: WomanJesus / ManThen
    m2 = re.match(r"^(.*?[a-z])([A-Z][a-z].+)$", text)
    if m2 and " " not in m2.group(1)[-12:]:
        return m2.group(2)
    return text


def _fetch_chapter(translation: str, book_id: int, chapter: int) -> Dict[int, str]:
    url = f"https://bolls.life/get-text/{translation}/{book_id}/{chapter}/"
    r = requests.get(url, headers=UA, timeout=45)
    r.raise_for_status()
    data = r.json()
    out: Dict[int, str] = {}
    for item in data:
        v = int(item["verse"])
        text = _strip_html(item.get("text", ""))
        if translation == "NIV":
            text = _strip_niv_heading(text)
        out[v] = text
    return out


def _clean_cuv(text: str) -> str:
    # Remove spaces between CJK / CJK punctuation
    text = re.sub(r"(?<=[\u4e00-\u9fff])\s+(?=[\u4e00-\u9fff])", "", text)
    text = re.sub(r"(?<=[\u4e00-\u9fff])\s+(?=[，。！？、：；）】」])", "", text)
    text = re.sub(r"(?<=[（【「，。！？、：；])\s+(?=[\u4e00-\u9fff])", "", text)
    text = re.sub(r"\s+([，。！？、：；）】」])", r"\1", text)
    text = re.sub(r"([（【「])\s+", r"\1", text)
    return text.strip()


def _niv_unusable(text: str) -> bool:
    """True when a verse is blank or bolls.life's NIV copyright refusal."""
    cleaned = (text or "").strip()
    if not cleaned:
        return True
    lowered = cleaned.lower()
    return "prohibited me from using the niv" in lowered or "biblica, inc." in lowered


def _missing_verse_nums(texts: Dict[int, str], ref: VerseRef) -> List[int]:
    return [n for n in ref.verse_numbers() if _niv_unusable(texts.get(n, ""))]


def _biblegateway_query(ref: VerseRef) -> str:
    if ref.start == ref.end:
        return f"{ref.book_en} {ref.chapter}:{ref.start}"
    return f"{ref.book_en} {ref.chapter}:{ref.start}-{ref.end}"


def _biblegateway_passage_url(ref: VerseRef) -> str:
    """Passage page linked from the NIV version page on Bible Gateway."""
    return (
        "https://www.biblegateway.com/passage/?search="
        + quote(_biblegateway_query(ref), safe=":")
        + "&version=NIV"
    )


def _biblegateway_span_text(span) -> str:
    """Passage wording as shown on the page, without footnote or cross-reference marks.

    NIV shows Mark 9:44 and 9:46 as bracketed numbers only. Keep that marker.
    """
    markers = [
        node.get_text(" ", strip=True)
        for node in span.select("sup.versenum")
    ]
    for node in span.select(
        "sup.footnote, sup.crossreference, sup.versenum, span.chapternum"
    ):
        node.decompose()
    text = re.sub(r"\s+", " ", span.get_text(" ", strip=True)).strip()
    if text:
        return text
    bracketed = [m for m in markers if m.startswith("[") and m.endswith("]")]
    return " ".join(bracketed)


def _fetch_niv_biblegateway(ref: VerseRef) -> Dict[int, str]:
    """NIV verses from the Bible Gateway passage page for this reference."""
    query = _biblegateway_query(ref)
    url = _biblegateway_passage_url(ref)
    response = None
    last_error: Optional[BaseException] = None
    for attempt in range(3):
        try:
            response = requests.get(url, headers=BG_UA, timeout=45)
            response.raise_for_status()
            break
        except requests.HTTPError as exc:
            status = exc.response.status_code if exc.response is not None else 0
            if status and status < 500:
                raise
            last_error = exc
            response = None
        except requests.RequestException as exc:
            last_error = exc
            response = None
        if attempt < 2:
            time.sleep(0.6)
    if response is None:
        raise RuntimeError(f"Bible Gateway request failed for {query}") from last_error
    soup = BeautifulSoup(response.text, "lxml")
    root = soup.select_one("div.version-NIV.result-text-style-normal")
    box = root.select_one(".std-text") if root is not None else None
    if box is None:
        raise RuntimeError(f"Bible Gateway did not return NIV text for {query}")

    parts: Dict[int, List[str]] = {}
    for span in box.select("span.text"):
        if span.find_parent(["h1", "h2", "h3", "h4", "h5"]):
            continue
        if span.find_parent("span", class_="text"):
            continue
        verse_no = None
        chapter_no = None
        for cls in span.get("class") or []:
            match = _BG_VERSE_CLASS.match(cls)
            if match:
                chapter_no = int(match.group(1))
                verse_no = int(match.group(2))
                break
        if verse_no is None or chapter_no != ref.chapter:
            continue
        if verse_no < ref.start or verse_no > ref.end:
            continue
        text = _biblegateway_span_text(span)
        if text:
            parts.setdefault(verse_no, []).append(text)

    out = {n: " ".join(chunks).strip() for n, chunks in parts.items()}
    if not out:
        raise RuntimeError(f"Bible Gateway NIV passage was empty for {query}")
    return out


def fetch_verses(
    ref: VerseRef,
    note: Optional[Callable[[str], None]] = None,
) -> Tuple[Dict[int, str], Dict[int, str]]:
    """Return (niv_en, cuv_zh) maps verse_num -> text.

    NIV is read from the Bible Gateway passage page for the reference
    (the same pages linked from the NIV version list). Chinese Union Version
    still comes from bolls.life.
    """
    def say(msg: str) -> None:
        if note:
            note(msg)

    passage_url = _biblegateway_passage_url(ref)
    say(f"Fetching NIV from Bible Gateway…\n{passage_url}")
    niv: Dict[int, str] = {}
    gateway_error: Optional[BaseException] = None
    try:
        niv = _fetch_niv_biblegateway(ref)
    except Exception as exc:  # noqa: BLE001 — fall back if the passage page fails
        gateway_error = exc
        say(f"Bible Gateway NIV page failed ({exc}). Trying bolls.life…")
        try:
            niv = _fetch_chapter("NIV", ref.book_id, ref.chapter)
        except Exception as boll_exc:
            raise RuntimeError(
                "Could not fetch NIV text from Bible Gateway "
                f"({exc}) or bolls.life ({boll_exc})."
            ) from boll_exc

    missing = _missing_verse_nums(niv, ref)
    if missing and gateway_error is None:
        say(
            "Bible Gateway is missing verses "
            f"{_verse_span(missing)}. Trying bolls.life…"
        )
        try:
            boll = _fetch_chapter("NIV", ref.book_id, ref.chapter)
        except Exception as exc:  # noqa: BLE001
            say(f"bolls.life NIV failed ({exc}).")
            boll = {}
        for n in list(missing):
            if not _niv_unusable(boll.get(n, "")):
                niv[n] = boll[n]
        missing = _missing_verse_nums(niv, ref)
    if missing:
        source = "Bible Gateway and bolls.life" if gateway_error is None else "bolls.life"
        raise RuntimeError(
            f"NIV text still missing for verses {_verse_span(missing)} after {source}."
        )
    if gateway_error is None:
        say("Using NIV text from Bible Gateway.")

    cuv = _fetch_chapter("CUV", ref.book_id, ref.chapter)
    cuv = {k: _clean_cuv(v) for k, v in cuv.items()}
    en = {n: niv.get(n, "") for n in ref.verse_numbers()}
    zh = {n: cuv.get(n, "") for n in ref.verse_numbers()}
    return en, zh


def _verse_span(nums: List[int]) -> str:
    if not nums:
        return ""
    if nums == list(range(nums[0], nums[-1] + 1)):
        if nums[0] == nums[-1]:
            return str(nums[0])
        return f"{nums[0]}–{nums[-1]}"
    return ", ".join(str(n) for n in nums)


def chunk_verses(verse_nums: List[int], max_per_slide: int = 7) -> List[List[int]]:
    """Split verses into 2–4 slides depending on length (fallback when no author sections)."""
    n = len(verse_nums)
    if n <= 0:
        return []
    if n <= max_per_slide:
        return [verse_nums]
    # Prefer 2 slides when possible, up to 4
    slides = 2 if n <= max_per_slide * 2 else (3 if n <= max_per_slide * 3 else 4)
    size = (n + slides - 1) // slides
    chunks = []
    for i in range(0, n, size):
        chunks.append(verse_nums[i : i + size])
    return chunks


def chunk_by_sections(
    verse_nums: List[int],
    sections: list,
    max_per_slide: int = 10,
) -> List[Tuple[List[int], str, str]]:
    """Return list of (verse_nums, title_zh, title_en) using author sections when available.

    If a section is longer than max_per_slide, split it further while keeping the section title.
    """
    if not verse_nums:
        return []
    verse_set = set(verse_nums)
    if sections:
        out: List[Tuple[List[int], str, str]] = []
        covered = set()
        for sec in sections:
            nums = [n for n in sec.verse_numbers() if n in verse_set]
            if not nums:
                continue
            covered.update(nums)
            title_zh = getattr(sec, "title_zh", "") or ""
            title_en = getattr(sec, "title_en", "") or ""
            if len(nums) <= max_per_slide:
                out.append((nums, title_zh, title_en))
            else:
                for i in range(0, len(nums), max_per_slide):
                    out.append((nums[i : i + max_per_slide], title_zh, title_en))
        # Any verses not covered by sections (append at end)
        leftover = [n for n in verse_nums if n not in covered]
        if leftover:
            for chunk in chunk_verses(leftover, max_per_slide=max_per_slide):
                out.append((chunk, "", ""))
        return out

    return [(chunk, "", "") for chunk in chunk_verses(verse_nums, max_per_slide=max_per_slide)]
