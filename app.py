# -*- coding: utf-8 -*-
"""Launch ROLCC Celldoc → PowerPoint GUI (or CLI)."""
from __future__ import annotations

import argparse
import sys
from pathlib import Path


def main(argv=None):
    parser = argparse.ArgumentParser(description="ROLCC celldoc to interactive PowerPoint")
    parser.add_argument("--cli", action="store_true", help="Run without GUI")
    parser.add_argument("--date", help="Study date YYYY.MM.DD or YYYYMMDD (CLI)")
    parser.add_argument("--pdf", help="Local celldoc PDF path (CLI, skips download)")
    parser.add_argument(
        "--out",
        default=r"G:\My Drive\Christianity\ROLCC",
        help="Output folder",
    )
    args = parser.parse_args(argv)

    if not args.cli:
        from rolcc_pptx.gui import main as gui_main

        gui_main()
        return 0

    from rolcc_pptx.catalog import StudyEntry, fetch_catalog
    from rolcc_pptx.pipeline import process_study

    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)

    if args.pdf:
        date = args.date or Path(args.pdf).stem[:8]
        # Normalize date_key digits
        date_key = "".join(ch for ch in date if ch.isdigit()) or "study"
        # Build a minimal StudyEntry for naming / scripture fallback
        entry = StudyEntry(
            date=f"{date_key[:4]}.{date_key[4:6]}.{date_key[6:8]}" if len(date_key) >= 8 else date_key,
            title="",
            scripture="",
            trad_url="",
            simp_url="",
        )
        # Ensure date_key property matches
        if len(date_key) >= 8:
            entry.date = f"{date_key[:4]}.{date_key[4:6]}.{date_key[6:8]}"
        result = process_study(entry, out_dir=out_dir, source_pdf=Path(args.pdf))
        print("Saved PPTX:", result.pptx)
        print("Saved Chinese PDF:", result.chinese_pdf)
        print("Saved English PDF:", result.english_pdf)
        return 0

    entries = fetch_catalog()
    if not entries:
        raise SystemExit("No studies found.")
    target = None
    if args.date:
        key = args.date.replace(".", "")
        for e in entries:
            if e.date_key == key or e.date == args.date:
                target = e
                break
        if not target:
            raise SystemExit(f"No study for date {args.date}")
    else:
        target = entries[0]
        for e in entries:
            if e.trad_url or e.simp_url:
                target = e
                break

    result = process_study(target, out_dir=out_dir)
    print("Saved PPTX:", result.pptx)
    print("Saved Chinese PDF:", result.chinese_pdf)
    print("Saved English PDF:", result.english_pdf)
    return 0


if __name__ == "__main__":
    sys.exit(main())
