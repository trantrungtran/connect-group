# -*- coding: utf-8 -*-
"""Export an English-translated study PDF next to the Chinese source."""
from __future__ import annotations

from pathlib import Path
from typing import Dict

from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer

from .parse_pdf import StudyContent


def _p(text: str, style) -> Paragraph:
    # Escape XML special chars for reportlab Paragraph
    safe = (
        (text or "")
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace("\n", "<br/>")
    )
    return Paragraph(safe, style)


def write_english_pdf(
    study: StudyContent,
    niv: Dict[int, str],
    out_path: Path,
) -> Path:
    """Write a readable English study PDF (scripture, commentary, questions)."""
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "TitleEn",
        parent=styles["Heading1"],
        fontName="Helvetica-Bold",
        fontSize=16,
        spaceAfter=8,
        textColor="#1B2A4A",
    )
    h2 = ParagraphStyle(
        "H2En",
        parent=styles["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=12,
        spaceBefore=14,
        spaceAfter=6,
        textColor="#1B2A4A",
    )
    body = ParagraphStyle(
        "BodyEn",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=10.5,
        leading=14,
        alignment=TA_LEFT,
        spaceAfter=6,
    )
    verse = ParagraphStyle(
        "VerseEn",
        parent=body,
        leftIndent=8,
        spaceAfter=8,
    )
    meta = ParagraphStyle(
        "MetaEn",
        parent=body,
        fontSize=10,
        textColor="#3A455C",
        spaceAfter=10,
    )

    doc = SimpleDocTemplate(
        str(out_path),
        pagesize=letter,
        leftMargin=0.75 * inch,
        rightMargin=0.75 * inch,
        topMargin=0.7 * inch,
        bottomMargin=0.7 * inch,
        title=study.title_en or study.title_zh or "ROLCC Study",
    )

    story = []
    title = study.title_en or study.title_zh or "ROLCC Connect Group Study"
    story.append(_p(title, title_style))
    if study.ref:
        story.append(_p(f"Scripture: {study.ref.display_en()} (NIV)", meta))
    if study.date_text:
        story.append(_p(f"Date: {study.date_text}", meta))
    story.append(_p("River of Life Christian Church · Connect Group", meta))
    story.append(Spacer(1, 6))

    story.append(_p("Scripture (NIV)", h2))
    if study.ref:
        for n in study.ref.verse_numbers():
            t = (niv.get(n) or "").strip()
            if t:
                story.append(_p(f"<b>{n}</b>  {t}", verse))
            else:
                story.append(_p(f"<b>{n}</b>", verse))

    story.append(_p("Commentary", h2))
    verses = sorted(set(study.commentaries_en) | set(study.commentaries_zh))
    if not verses and study.ref:
        verses = study.ref.verse_numbers()
    for n in verses:
        en = (study.commentaries_en.get(n) or "").strip()
        zh = (study.commentaries_zh.get(n) or "").strip()
        text = en or zh
        if not text:
            continue
        story.append(_p(f"<b>Verse {n}</b>", body))
        story.append(_p(text, verse))

    if study.questions:
        story.append(_p("Discussion Questions", h2))
        for q in study.questions:
            qtext = q.english or q.chinese
            label = q.short_label or f"Q{q.number}"
            story.append(_p(f"<b>{q.number}. {label}</b> — {qtext}", verse))

    doc.build(story)
    return out_path
