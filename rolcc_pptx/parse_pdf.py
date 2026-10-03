# -*- coding: utf-8 -*-
"""Parse ROLCC celldoc PDF into structured study content (2-column aware)."""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

import pdfplumber

from .bible import VerseRef, parse_scripture_ref


@dataclass
class DiscussionQuestion:
    number: int
    chinese: str
    english: str = ""
    short_label: str = ""
    section_index: int = -1  # 0-based author section; -1 = unassigned


@dataclass
class StudySection:
    """Author section such as 一、…（馬可8：1-10）."""
    title_zh: str
    start: int
    end: int
    title_en: str = ""

    def verse_numbers(self) -> List[int]:
        return list(range(self.start, self.end + 1))


@dataclass
class StudyContent:
    title_zh: str
    date_text: str
    scripture_raw: str
    ref: Optional[VerseRef]
    commentaries_zh: Dict[int, str] = field(default_factory=dict)
    commentaries_en: Dict[int, str] = field(default_factory=dict)
    questions: List[DiscussionQuestion] = field(default_factory=list)
    sections: List[StudySection] = field(default_factory=list)
    raw_text: str = ""
    title_en: str = ""


def _dedupe_chars(s: str) -> str:
    """Collapse 4× repeated glyphs used in some ROLCC PDFs."""
    return re.sub(r"(.)\1{3}", r"\1", s)


def _words_to_text(words: list) -> str:
    if not words:
        return ""
    words = sorted(words, key=lambda w: (round(w["top"] / 2.5) * 2.5, w["x0"]))
    lines: List[str] = []
    cur_y = None
    cur: List[str] = []

    def flush():
        nonlocal cur
        if not cur:
            return
        joined_nospace = "".join(cur)
        if any("\u4e00" <= c <= "\u9fff" for c in joined_nospace):
            # CJK-heavy line: drop spaces between CJK, keep spaces around ASCII
            s = ""
            for tok in cur:
                if not s:
                    s = tok
                elif (
                    ("\u4e00" <= s[-1] <= "\u9fff" or s[-1] in "，。！？、：；）】」")
                    and ("\u4e00" <= tok[0] <= "\u9fff" or tok[0] in "，。！？、：；（【「")
                ):
                    s += tok
                else:
                    s += tok if tok[:1] in ",.!?;:" or s[-1:] in ",.!?;:" else (" " + tok if tok.isascii() and s[-1:].isascii() else tok)
            lines.append(s)
        else:
            lines.append(" ".join(cur))
        cur = []

    for w in words:
        y = round(w["top"] / 2.5) * 2.5
        if cur_y is None or abs(y - cur_y) <= 2.5:
            cur.append(w["text"])
            cur_y = y if cur_y is None else cur_y
        else:
            flush()
            cur = [w["text"]]
            cur_y = y
    flush()
    return _dedupe_chars("\n".join(lines))


def extract_column_texts(path: str) -> Tuple[List[str], List[str], str]:
    """Return (left_cols, right_cols, full_flat) for each page."""
    lefts: List[str] = []
    rights: List[str] = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            mid = page.width / 2
            words = page.extract_words(use_text_flow=False) or []
            left_w = [w for w in words if w["x0"] < mid - 8]
            right_w = [w for w in words if w["x0"] >= mid - 8]
            lefts.append(_words_to_text(left_w))
            rights.append(_words_to_text(right_w))
    full = "\n".join(x for pair in zip(lefts, rights) for x in pair)
    return lefts, rights, full


def _is_study_column(text: str) -> bool:
    if not text or len(text) < 40:
        return False
    score = 0
    if re.search(r"[（(]\d{1,3}[）)]", text):
        score += 3
    if "討論問題" in text:
        score += 2
    if re.search(r"[一二三四五]、", text):
        score += 2
    if re.search(r"[A-D]\.\s*", text) or re.search(r"[ＡＢＣＤ]\.", text):
        score += 1
    if "第一時段" in text or "第二時段" in text:
        score -= 2
    if "宣教單位" in text:
        score -= 2
    return score >= 2


def _study_body(lefts: List[str], rights: List[str]) -> str:
    parts: List[str] = []
    for left, right in zip(lefts, rights):
        # Keep reading order: left then right, but only study-like columns.
        # On page 1, right is usually study; left is admin + theme.
        # On later pages both may be study.
        for col in (left, right):
            if _is_study_column(col):
                parts.append(col)
            elif "本週主題" in col or "經文：" in col or "經文:" in col:
                # keep metadata-bearing non-study column separately via meta extractor
                pass
    return "\n".join(parts)


def _meta_text(lefts: List[str], rights: List[str]) -> str:
    chunks = []
    for left, right in zip(lefts, rights):
        for col in (left, right):
            if any(k in col for k in ("本週主題", "經文：", "經文:", "小組聚會材料", "系列")):
                chunks.append(col)
    return "\n".join(chunks)


def _clean_title(title: str) -> str:
    title = title.strip()
    # Cut at first bleed of commentary / punctuation run-on
    title = re.split(r"[。！？\n]", title)[0]
    title = re.sub(r"\s+", " ", title)
    # Keep up to ~40 CJK chars
    m = re.match(r"^([\u4e00-\u9fffA-Za-z0-9系列生活查經篇－\-–—]{2,60})", title)
    if m:
        title = m.group(1)
    # Trim trailing incomplete words from column bleed
    title = re.sub(r"(般猶太人.*|出去給.*|主耶穌在.*)$", "", title)
    return title.strip(" ：:-")


def _find_title_date(meta: str, full: str) -> Tuple[str, str]:
    text = meta + "\n" + full
    date = ""
    m = re.search(r"(20\d{2}\s*年[^\n]{0,24}日)", text)
    if m:
        date = re.sub(r"\s+", "", m.group(1))
    title = ""
    m = re.search(r"本週主題[:：]\s*([^\n]+)", text)
    if m:
        title = _clean_title(m.group(1))
    if not title:
        m = re.search(
            r"((?:馬可福音|腓立比書|哥林多前書|生活查經|詩篇)\S{0,40})",
            text,
        )
        if m:
            title = _clean_title(m.group(1))
    return title, date


def _find_scripture(meta: str, body: str) -> str:
    for text in (meta, body):
        m = re.search(r"經文[:：]\s*([^\n]+)", text)
        if m:
            raw = m.group(1).strip()
            # Keep only the ref token
            m2 = re.search(
                r"((?:可|羅|詩篇|詩|林前|林後|腓|弗|雅|賽|書|太|路|約|徒|西|加|來|馬可福音)\s*\d+\s*[:：]\s*\d+(?:\s*[-–—~至到]\s*\d+)?)",
                raw,
            )
            return m2.group(1) if m2 else raw.split()[0]
    m = re.search(
        r"[（(]((?:可|羅|詩|林前|腓|弗|雅|賽|太|路|約|徒|西|加|來)\d+[:：]\d+[-–—]?\d*)[）)]",
        body,
    )
    return m.group(1) if m else ""


def _clean_commentary_chunk(chunk: str) -> str:
    chunk = re.sub(r"^.*再讀下文[）)]", "", chunk)
    chunk = re.sub(r"^[A-DＡ-Ｄ]\.\s*[^：:]{0,60}[:：]", "", chunk)
    chunk = re.sub(r"^[一二三四五]、\s*[^：:]{0,80}[:：]", "", chunk)
    chunk = re.sub(r"^\d{1,3}\s*[-–—]\s*\d{1,3}\s*[）)]\s*[:：]?", "", chunk)
    chunk = re.sub(r"^（?觀察[^）]*）", "", chunk)
    chunk = re.sub(r"^觀察題）", "", chunk)
    # If a section heading remains before narrative, keep text after last fullwidth/half colon
    # only when a long heading prefix is present.
    if "：" in chunk[:80] or ":" in chunk[:80]:
        # Prefer text after heading colon when left side looks like a title
        m = re.search(r"[：:]([^：:]{12,})$", chunk)
        left = chunk.split("：")[0] if "：" in chunk else chunk.split(":")[0]
        if m and len(left) < 50 and not left.endswith(("她", "他", "說", "的", "了")):
            # only if left looks like heading (contains 婦人/耶穌醫治 etc. short label)
            if re.search(r"[懇求謙卑痊癒醫治傳開帶]", left) and len(left) < 40:
                chunk = m.group(1)
    return chunk.strip(" ：:，,“\"")


def _extract_commentaries(body: str) -> Dict[int, str]:
    flat = re.sub(r"\n+", "", body)
    # Drop memory-verse / prayer / schedule tails for marker search region
    flat = re.split(r"禱告[:：]|第四時段", flat)[0]
    # Split on sentence ends but keep verse markers tied to preceding clause.
    # Use a more permissive capture that can cross 。 when needed by scanning markers.
    hits = re.findall(r"([^。！？]{6,500}?)[（(](\d{1,3})[）)]", flat)
    by_verse: Dict[int, List[str]] = {}
    for chunk, num_s in hits:
        num = int(num_s)
        chunk = _clean_commentary_chunk(chunk.strip(" ：:，,"))
        if any(k in chunk for k in ("背誦經文", "第一時段", "第二時段", "宣教單位", "生活查經", "馬可福音8")):
            continue
        if len(chunk) < 8:
            continue
        by_verse.setdefault(num, []).append(chunk)
    return {k: " ".join(dict.fromkeys(v)) for k, v in sorted(by_verse.items())}


def _extract_sections(body: str, ref: Optional[VerseRef]) -> List[StudySection]:
    """Parse author major sections: 一、標題（馬可8：1-10） / （可7：24-30）."""
    flat = re.sub(r"\s+", "", body)
    sections: List[StudySection] = []
    # 一、耶穌餵飽四千人（馬可8：1-10） or （可7：24-30）
    pattern = re.compile(
        r"([一二三四五六七八九十]+)、"
        r"([^（(]{2,40}?)"
        r"[（(]"
        r"(?:馬可|可|羅|詩|林前|林後|腓|弗|雅|賽|太|路|約|徒|西|加|來)?"
        r"(?:福音)?"
        r"(\d+)?[：:．.]?"
        r"(\d+)\s*[-–—~至到]\s*(\d+)"
        r"[）)]"
    )
    for m in pattern.finditer(flat):
        marker, title, chap_s, start_s, end_s = m.groups()
        start, end = int(start_s), int(end_s)
        if end < start:
            continue
        # Prefer chapter from heading; else use study ref chapter
        chapter = int(chap_s) if chap_s else (ref.chapter if ref else 0)
        if ref and chapter and chapter != ref.chapter:
            # Skip ranges that clearly belong to another chapter (e.g. next-week preview)
            continue
        if ref and (end < ref.start or start > ref.end):
            continue
        title_zh = f"{marker}、{title.strip(' ：:')}"
        sections.append(StudySection(title_zh=title_zh, start=start, end=end))

    # Deduplicate overlapping identical ranges
    uniq: Dict[Tuple[int, int], StudySection] = {}
    for sec in sections:
        uniq.setdefault((sec.start, sec.end), sec)
    out = [uniq[k] for k in sorted(uniq)]
    return out


def _extract_questions(body: str, meta: str) -> List[DiscussionQuestion]:
    text = body + "\n" + meta
    text = re.sub(r"\r", "", text)
    qs: List[DiscussionQuestion] = []
    for m in re.finditer(
        r"討論問題\s*(\d+)\s*[:：]\s*(.+?)(?=討論問題\s*\d+|禱告[:：]|第四時段|\d{1,2}/\d{2}|$)",
        text,
        flags=re.S,
    ):
        num = int(m.group(1))
        body_q = re.sub(r"\s+", "", m.group(2))
        body_q = re.sub(r"[（(]觀察.*?[）)]", "", body_q)
        body_q = body_q.strip(" ：:。")
        if not body_q:
            continue
        if num == 1 and ("背誦" in body_q or "上主日" in body_q or "以賽亞" in body_q):
            continue
        # Trim bleed after full question (heuristic: cut at section markers)
        body_q = re.split(r"[一二三]、|[A-DＡ-Ｄ]\.|此段經文|親愛的阿爸", body_q)[0]
        qs.append(DiscussionQuestion(number=num, chinese=body_q))
    seen = {}
    for q in qs:
        seen.setdefault(q.number, q)
    return [seen[k] for k in sorted(seen)]


def _assign_questions_to_sections(
    body: str, sections: List[StudySection], questions: List[DiscussionQuestion]
) -> None:
    """Set question.section_index from document order (question belongs to prior section)."""
    if not sections or not questions:
        return
    flat = re.sub(r"\s+", "", body)

    # Locate each major section heading in flattened body
    sec_pos: List[Tuple[int, int]] = []  # (start_pos, section_index)
    for i, sec in enumerate(sections):
        # Match 一、標題…（…start-end）
        marker = sec.title_zh.split("、", 1)[0] if "、" in sec.title_zh else ""
        title_core = sec.title_zh.split("、", 1)[-1]
        pat = (
            re.escape(marker + "、" + title_core[:8])
            if marker
            else re.escape(title_core[:12])
        )
        m = re.search(pat, flat)
        if m:
            sec_pos.append((m.start(), i))
        else:
            # Fallback: search by verse range in heading
            m2 = re.search(
                rf"[一二三四五六七八九十]+、[^（(]{{2,40}}[（(][^）)]*{sec.start}\s*[-–—]\s*{sec.end}",
                flat,
            )
            if m2:
                sec_pos.append((m2.start(), i))
    sec_pos.sort()

    for q in questions:
        m = re.search(rf"討論問題\s*{q.number}\s*[:：]", flat)
        if not m:
            # try without spaces already stripped
            m = re.search(rf"討論問題{q.number}[:：]", flat)
        if not m:
            continue
        qpos = m.start()
        # Belong to the latest section whose heading appears before this question
        assigned = -1
        for spos, sidx in sec_pos:
            if spos < qpos:
                assigned = sidx
            else:
                break
        q.section_index = assigned


def parse_celldoc_pdf(path: str) -> StudyContent:
    lefts, rights, full = extract_column_texts(path)
    meta = _meta_text(lefts, rights)
    body = _study_body(lefts, rights)
    title, date = _find_title_date(meta, full)
    scripture_raw = _find_scripture(meta, body)
    ref = parse_scripture_ref(scripture_raw.replace("馬可福音", "可"))
    commentaries = _extract_commentaries(body)
    questions = _extract_questions(body, meta)
    sections = _extract_sections(body, ref)
    _assign_questions_to_sections(body, sections, questions)
    return StudyContent(
        title_zh=title,
        date_text=date,
        scripture_raw=scripture_raw,
        ref=ref,
        commentaries_zh=commentaries,
        questions=questions,
        sections=sections,
        raw_text=full,
    )
