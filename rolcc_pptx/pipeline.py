# -*- coding: utf-8 -*-
"""End-to-end: download → parse → translate → PowerPoint + PDFs."""
from __future__ import annotations

import re
import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Optional

from .bible import fetch_verses, parse_scripture_ref
from .catalog import StudyEntry, download_pdf
from .export_pdf import write_english_pdf
from .parse_pdf import parse_celldoc_pdf
from .pptx_builder import build_presentation
from .translate import enrich_with_english

DEFAULT_OUT_DIR = Path(r"G:\My Drive\Christianity\ROLCC")
CACHE_DIR = Path(r"G:\My Drive\Agent\Connect Group\downloads")


ProgressCb = Optional[Callable[[str], None]]


@dataclass
class ProcessResult:
    pptx: Path
    chinese_pdf: Path
    english_pdf: Path


def _log(cb: ProgressCb, msg: str):
    if cb:
        cb(msg)
    else:
        try:
            print(msg)
        except UnicodeEncodeError:
            print(msg.encode("ascii", "replace").decode("ascii"))


def _safe_filename(s: str) -> str:
    s = re.sub(r'[<>:"/\\|?*]', "", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s[:80]


def _stem(entry_date_key: str, study, book: str) -> str:
    return _safe_filename(
        f"{entry_date_key} {book} {study.ref.chapter}_{study.ref.start}-{study.ref.end}"
    )


def process_study(
    entry: StudyEntry,
    out_dir: Path = DEFAULT_OUT_DIR,
    prefer_traditional: bool = True,
    progress: ProgressCb = None,
    source_pdf: Optional[Path] = None,
) -> ProcessResult:
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    out_dir.mkdir(parents=True, exist_ok=True)

    if source_pdf is not None:
        pdf_path = Path(source_pdf)
        if not pdf_path.exists():
            raise FileNotFoundError(pdf_path)
        _log(progress, f"Using local PDF: {pdf_path}")
    else:
        url = entry.trad_url if prefer_traditional and entry.trad_url else entry.simp_url
        if not url:
            url = entry.trad_url or entry.simp_url
        if not url:
            raise RuntimeError("No download URL for this study.")
        pdf_path = CACHE_DIR / f"{entry.date_key}-celldoc.pdf"
        _log(progress, f"Downloading Chinese PDF…\n{url}")
        download_pdf(url, str(pdf_path))
        _log(progress, f"Cached PDF: {pdf_path}")

    _log(progress, "Parsing celldoc…")
    study = parse_celldoc_pdf(str(pdf_path))
    if study.ref is None and entry.scripture:
        study.scripture_raw = entry.scripture
        study.ref = parse_scripture_ref(entry.scripture)
    if study.ref is None:
        raise RuntimeError(
            f"Could not parse scripture reference from '{study.scripture_raw or entry.scripture}'."
        )
    if not study.title_zh:
        study.title_zh = entry.title

    _log(progress, f"Scripture: {study.ref.display_en()}")
    _log(
        progress,
        f"Found {len(study.commentaries_zh)} verse commentaries, "
        f"{len(study.questions)} discussion questions.",
    )

    book = study.ref.book_en.replace(" ", "")
    stem = _stem(entry.date_key, study, book)

    # Same stem as PPTX / EN PDF: "{date} Mark 8_1-21 CH.pdf"
    chinese_pdf = out_dir / f"{stem} CH.pdf"
    shutil.copy2(pdf_path, chinese_pdf)
    _log(progress, f"Saved Chinese PDF: {chinese_pdf}")
    # Remove legacy naming if present
    legacy = out_dir / f"{entry.date_key}-celldoc-zh.pdf"
    if legacy.exists() and legacy.resolve() != chinese_pdf.resolve():
        try:
            legacy.unlink()
        except OSError:
            pass

    _log(progress, "Fetching NIV + Chinese Union Version text...")
    niv, cuv = fetch_verses(study.ref)

    _log(progress, "Translating commentary & questions to English…")
    enrich_with_english(study)

    english_pdf = out_dir / f"{stem} EN.pdf"
    _log(progress, "Writing English translated PDF…")
    write_english_pdf(study, niv, english_pdf)
    _log(progress, f"Saved English PDF: {english_pdf}")

    out_path = out_dir / f"{stem}.pptx"
    _log(progress, "Building PowerPoint…")
    build_presentation(study, niv, cuv, str(out_path))
    _log(progress, f"Done: {out_path}")

    return ProcessResult(pptx=out_path, chinese_pdf=chinese_pdf, english_pdf=english_pdf)
