# -*- coding: utf-8 -*-
"""Tkinter GUI: pick a ROLCC study → download → translate → PowerPoint."""
from __future__ import annotations

import threading
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk
from typing import List, Optional

from .catalog import StudyEntry, fetch_catalog
from .pipeline import DEFAULT_OUT_DIR, process_study


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("ROLCC Celldoc → PowerPoint")
        self.geometry("980x640")
        self.minsize(820, 520)

        self.entries: List[StudyEntry] = []
        self.out_dir = tk.StringVar(value=str(DEFAULT_OUT_DIR))
        self.status = tk.StringVar(value="Click Refresh to load studies from rolcc.net")
        self._busy = False

        self._build()

    def _build(self):
        top = ttk.Frame(self, padding=10)
        top.pack(fill=tk.X)

        ttk.Label(top, text="ROLCC Connect Group Bible Study → PPTX", font=("Segoe UI", 12, "bold")).pack(
            side=tk.LEFT
        )
        ttk.Button(top, text="Refresh list", command=self.refresh).pack(side=tk.RIGHT, padx=4)

        mid = ttk.Frame(self, padding=(10, 0, 10, 0))
        mid.pack(fill=tk.BOTH, expand=True)

        cols = ("date", "title", "scripture")
        self.tree = ttk.Treeview(mid, columns=cols, show="headings", selectmode="browse")
        self.tree.heading("date", text="Date")
        self.tree.heading("title", text="Title")
        self.tree.heading("scripture", text="Scripture")
        self.tree.column("date", width=110, stretch=False)
        self.tree.column("title", width=560)
        self.tree.column("scripture", width=140, stretch=False)
        scroll = ttk.Scrollbar(mid, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)

        opts = ttk.Frame(self, padding=10)
        opts.pack(fill=tk.X)
        ttk.Label(opts, text="Save to:").pack(side=tk.LEFT)
        ttk.Entry(opts, textvariable=self.out_dir, width=70).pack(side=tk.LEFT, padx=6)

        actions = ttk.Frame(self, padding=(10, 0, 10, 10))
        actions.pack(fill=tk.X)
        self.btn = ttk.Button(actions, text="Download · Translate · Build PowerPoint", command=self.run)
        self.btn.pack(side=tk.LEFT)
        ttk.Label(actions, textvariable=self.status, wraplength=700).pack(side=tk.LEFT, padx=12)

        log_frame = ttk.LabelFrame(self, text="Log", padding=8)
        log_frame.pack(fill=tk.BOTH, expand=False, padx=10, pady=(0, 10))
        self.log = tk.Text(log_frame, height=10, wrap=tk.WORD, font=("Consolas", 9))
        self.log.pack(fill=tk.BOTH, expand=True)

    def _append_log(self, msg: str):
        self.log.insert(tk.END, msg + "\n")
        self.log.see(tk.END)
        self.status.set(msg.split("\n")[0][:120])
        self.update_idletasks()

    def refresh(self):
        if self._busy:
            return
        self._busy = True
        self.btn.configure(state=tk.DISABLED)
        self._append_log("Loading catalog…")

        def work():
            try:
                entries = fetch_catalog()
                self.after(0, lambda: self._on_catalog(entries, None))
            except Exception as e:  # noqa: BLE001
                self.after(0, lambda: self._on_catalog([], e))

        threading.Thread(target=work, daemon=True).start()

    def _on_catalog(self, entries: List[StudyEntry], err: Optional[Exception]):
        self._busy = False
        self.btn.configure(state=tk.NORMAL)
        self.tree.delete(*self.tree.get_children())
        if err:
            self._append_log(f"Error: {err}")
            messagebox.showerror("Catalog error", str(err))
            return
        self.entries = entries
        for i, e in enumerate(entries):
            self.tree.insert("", tk.END, iid=str(i), values=(e.date, e.title, e.scripture))
        self._append_log(f"Loaded {len(entries)} studies.")
        if entries:
            self.tree.selection_set("0")

    def _selected(self) -> Optional[StudyEntry]:
        sel = self.tree.selection()
        if not sel:
            return None
        return self.entries[int(sel[0])]

    def run(self):
        if self._busy:
            return
        entry = self._selected()
        if not entry:
            messagebox.showinfo("Select a study", "Please select a study from the list.")
            return
        out_dir = Path(self.out_dir.get().strip() or str(DEFAULT_OUT_DIR))
        self._busy = True
        self.btn.configure(state=tk.DISABLED)
        self._append_log(f"Starting: {entry.label()}")

        def work():
            try:
                result = process_study(
                    entry,
                    out_dir=out_dir,
                    progress=lambda m: self.after(0, lambda msg=m: self._append_log(msg)),
                )
                self.after(0, lambda: self._on_done(result, None))
            except Exception as e:  # noqa: BLE001
                self.after(0, lambda: self._on_done(None, e))

        threading.Thread(target=work, daemon=True).start()

    def _on_done(self, result, err: Optional[Exception]):
        self._busy = False
        self.btn.configure(state=tk.NORMAL)
        if err:
            self._append_log(f"FAILED: {err}")
            messagebox.showerror("Build failed", str(err))
            return
        self._append_log(f"SUCCESS PPTX: {result.pptx}")
        self._append_log(f"SUCCESS Chinese PDF: {result.chinese_pdf}")
        self._append_log(f"SUCCESS English PDF: {result.english_pdf}")
        messagebox.showinfo(
            "Done",
            "Saved to Google Drive:\n\n"
            f"PowerPoint:\n{result.pptx}\n\n"
            f"Chinese PDF:\n{result.chinese_pdf}\n\n"
            f"English PDF:\n{result.english_pdf}",
        )


def main():
    app = App()
    app.after(200, app.refresh)
    app.mainloop()


if __name__ == "__main__":
    main()
