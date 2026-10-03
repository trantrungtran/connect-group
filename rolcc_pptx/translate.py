# -*- coding: utf-8 -*-
"""Translate Chinese study text to English (with caching + fallbacks)."""
from __future__ import annotations

import hashlib
import json
import re
import time
from pathlib import Path
from typing import Dict

from .parse_pdf import StudyContent

CACHE_PATH = Path(r"G:\My Drive\Agent\Connect Group\downloads\translation_cache.json")


def _load_cache() -> Dict[str, str]:
    if CACHE_PATH.exists():
        try:
            return json.loads(CACHE_PATH.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def _save_cache(cache: Dict[str, str]) -> None:
    CACHE_PATH.parent.mkdir(parents=True, exist_ok=True)
    CACHE_PATH.write_text(json.dumps(cache, ensure_ascii=False, indent=1), encoding="utf-8")


def _key(text: str) -> str:
    return hashlib.sha1(text.encode("utf-8")).hexdigest()


def _translate_one(text: str, cache: Dict[str, str]) -> str:
    text = (text or "").strip()
    if not text:
        return ""
    if len(re.findall(r"[\u4e00-\u9fff]", text)) < 2:
        return text
    k = _key(text)
    if k in cache:
        return cache[k]

    translated = ""
    # Prefer Google, fall back to MyMemory
    try:
        from deep_translator import GoogleTranslator

        translator = GoogleTranslator(source="zh-TW", target="en")
        if len(text) <= 4000:
            translated = translator.translate(text)
        else:
            parts = []
            buf = ""
            for sent in re.split(r"(?<=[。！？.!?])", text):
                if len(buf) + len(sent) > 3800:
                    parts.append(translator.translate(buf))
                    time.sleep(0.8)
                    buf = sent
                else:
                    buf += sent
            if buf:
                parts.append(translator.translate(buf))
            translated = "".join(parts)
        time.sleep(0.6)
    except Exception:
        try:
            from deep_translator import MyMemoryTranslator

            translated = MyMemoryTranslator(source="zh-TW", target="en-US").translate(text[:500])
            time.sleep(0.4)
        except Exception:
            translated = ""

    if not translated or translated.startswith("[Translation unavailable"):
        # Keep Chinese rather than injecting error text into slides
        return text

    cache[k] = translated
    return translated


def _short_label(en: str, fallback: str) -> str:
    en = re.sub(r"\s+", " ", en).strip()
    keywords = [
        ("attitude", "Attitude"),
        ("faith", "Faith"),
        ("humility", "Humility"),
        ("humble", "Humility"),
        ("method", "Method"),
        ("heal", "Healing"),
        ("possess", "Possession"),
        ("fear", "Fear"),
        ("grace", "Grace"),
        ("prayer", "Prayer"),
        ("authority", "Authority"),
        ("follow", "Follow"),
        ("deny", "Deny Self"),
        ("bread", "Bread"),
        ("feed", "Feeding"),
        ("territory", "Race"),
        ("race", "Race"),
        ("gentile territory", "Race"),
    ]
    low = en.lower()
    for needle, label in keywords:
        if needle in low:
            return label
    words = [w for w in re.findall(r"[A-Za-z]+", en) if len(w) > 3][:2]
    if words:
        return " ".join(w.capitalize() for w in words)[:18]
    return fallback


def enrich_with_english(study: StudyContent) -> StudyContent:
    cache = _load_cache()
    if study.title_zh and not study.title_en:
        study.title_en = _translate_one(study.title_zh, cache)

    new_en: Dict[int, str] = {}
    for num, zh in study.commentaries_zh.items():
        if study.commentaries_en.get(num):
            new_en[num] = study.commentaries_en[num]
        else:
            new_en[num] = _translate_one(zh, cache)
    study.commentaries_en = new_en

    for q in study.questions:
        if not q.english:
            q.english = _translate_one(q.chinese, cache)
        # If still Chinese (failed translate), label by number
        q.short_label = _short_label(q.english, f"Q{q.number}")
        if re.search(r"[\u4e00-\u9fff]", q.short_label):
            q.short_label = f"Q{q.number}"

    for sec in getattr(study, "sections", []) or []:
        if sec.title_zh and not sec.title_en:
            # Drop leading 一、 marker for cleaner English title
            src = re.sub(r"^[一二三四五六七八九十]+、\s*", "", sec.title_zh)
            sec.title_en = _translate_one(src, cache)

    _save_cache(cache)
    return study
