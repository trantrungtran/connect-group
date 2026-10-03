# -*- coding: utf-8 -*-
"""Fetch the ROLCC celldoc study catalog."""
from __future__ import annotations

import re
from dataclasses import dataclass
from typing import List, Optional
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup

CATALOG_URL = "https://rolcc.net/rolcc2/zones-and-cells/celldoc/"
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) ROLCC-PPTX/1.0"}


@dataclass
class StudyEntry:
    date: str  # YYYY.MM.DD
    title: str
    scripture: str
    trad_url: str
    simp_url: str

    @property
    def date_key(self) -> str:
        return self.date.replace(".", "")

    def label(self) -> str:
        return f"{self.date}  |  {self.title}  |  {self.scripture}"


def fetch_catalog(session: Optional[requests.Session] = None) -> List[StudyEntry]:
    sess = session or requests.Session()
    r = sess.get(CATALOG_URL, headers=UA, timeout=45)
    r.raise_for_status()
    r.encoding = r.apparent_encoding or "utf-8"
    soup = BeautifulSoup(r.text, "html.parser")

    entries: List[StudyEntry] = []
    for tr in soup.find_all("tr"):
        tds = tr.find_all("td")
        if len(tds) < 4:
            continue
        date = tds[0].get_text(" ", strip=True)
        if not re.match(r"^\d{4}\.\d{2}\.\d{2}$", date):
            continue
        title = tds[1].get_text(" ", strip=True)
        scripture = tds[2].get_text(" ", strip=True)
        links = []
        for td in tds[3:]:
            a = td.find("a", href=True)
            links.append(urljoin(CATALOG_URL, a["href"]) if a else "")
        trad = links[0] if links else ""
        simp = links[1] if len(links) > 1 else ""
        if not trad and not simp:
            continue
        entries.append(
            StudyEntry(
                date=date,
                title=title,
                scripture=scripture,
                trad_url=trad,
                simp_url=simp,
            )
        )
    return entries


def download_pdf(url: str, dest_path: str, session: Optional[requests.Session] = None) -> str:
    sess = session or requests.Session()
    r = sess.get(url, headers=UA, timeout=90)
    r.raise_for_status()
    with open(dest_path, "wb") as f:
        f.write(r.content)
    return dest_path
