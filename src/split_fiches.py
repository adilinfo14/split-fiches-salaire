import argparse
import csv
import logging
from logging.handlers import RotatingFileHandler
import os
import re
import threading
import webbrowser
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Callable, Optional

from pypdf import PdfReader, PdfWriter

import tkinter as tk
from tkinter import filedialog, messagebox, ttk


# ---------------- Parsing rules ----------------

CIVILITY_RE = re.compile(r"\b(Madame|Monsieur)\b\s+([A-ZÀ-ÖØ-Ý][A-ZÀ-ÖØ-Ý\s\-']+)", re.UNICODE)
PERIOD_RE = re.compile(r"Période\s*:\s*(\d{2})\.(\d{4})", re.UNICODE)

UPPER_NAME_RE = re.compile(r"^[A-ZÀ-ÖØ-Ý]{2,}(?:[\s\-'][A-ZÀ-ÖØ-Ý]{2,})+$", re.UNICODE)
UPPER_EXCLUDE = {"OSAD", "HELVETIA", "SARL", "DECOMPTE", "DÉCOMPTE", "SALAIRE"}


def clean_filename(text: str) -> str:
    text = text.strip()
    text = re.sub(r"[^\w\s\-]", "", text, flags=re.UNICODE)
    text = re.sub(r"\s+", "_", text)
    return text[:120] if len(text) > 120 else text


def extract_name(text: str) -> Optional[str]:
    text = text or ""
    m_civ = CIVILITY_RE.search(text)
    if m_civ:
        return m_civ.group(2)

    lines = [l.strip() for l in text.splitlines() if l.strip()]
    for line in lines[:25]:
        if UPPER_NAME_RE.match(line):
            upper_words = set(re.split(r"[\s\-']+", line))
            if upper_words & UPPER_EXCLUDE:
                continue
            return line
    return None


def extract_period(text: str) -> Optional[str]:
    m = PERIOD_RE.search(text or "")
    if not m:
        return None
    return f"{m.group(1)}-{m.group(2)}"


def is_new_payslip_page(text: str) -> bool:
    # Heuristic: a payslip "header" page contains BOTH name and period
    return (extract_name(text) is not None) and (extract_period(text) is not None)


# ---------------- Logging & OS helpers ----------------

def setup_logger(log_path: Path) -> logging.Logger:
    logger = logging.getLogger("split_payslips")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    ch = logging.StreamHandler()
    ch.setLevel(logging.INFO)
    ch.setFormatter(logging.Formatter("%(message)s"))

    fh = RotatingFileHandler(log_path, maxBytes=2_000_000, backupCount=5, encoding="utf-8")
    fh.setLevel(logging.INFO)
    fh.setFormatter(logging.Formatter("%(asctime)s | %(levelname)s | %(message)s"))

    logger.addHandler(ch)
    logger.addHandler(fh)
    return logger


def open_folder(path: Path):
    path = path.resolve()
    if os.name == "nt":
        os.startfile(str(path))  # type: ignore[attr-defined]
    else:
        webbrowser.open(path.as_uri())


def open_file(path: Path):
    path = path.resolve()
    if os.name == "nt":
        os.startfile(str(path))  # type: ignore[attr-defined]
    else:
        webbrowser.open(path.as_uri())


# ---------------- Project dirs (as you requested) ----------------

def project_root() -> Path:
    # split-fiches-salaire/
    return Path(__file__).resolve().parents[1]


def ensure_base_dirs(root: Path):
    (root / "input").mkdir(exist_ok=True)
    (root / "output").mkdir(exist_ok=True)
    (root / "logs").mkdir(exist_ok=True)
    (root / "errors").mkdir(exist_ok=True)


def make_dirs(root: Path, timestamp: str):
    ok_dir = root / "output" / f"split_{timestamp}"
    err_dir = root / "errors" / f"split_{timestamp}"
    logs_dir = root / "logs"
    ok_dir.mkdir(parents=True, exist_ok=True)
    err_dir.mkdir(parents=True, exist_ok=True)
    logs_dir.mkdir(parents=True, exist_ok=True)
    return ok_dir, err_dir, logs_dir


# ---------------- PDF writing helpers ----------------

def write_pages(reader: PdfReader, page_indices: list[int], out_path: Path):
    writer = PdfWriter()
    for idx in page_indices:
        writer.add_page(reader.pages[idx])
    with open(out_path, "wb") as f:
        writer.write(f)


# ---------------- Records + CSV ----------------

@dataclass
class Record:
    status: str                 # OK / FALLBACK / ERROR / ORPHAN
    name: str                   # employee name or "-"
    period: str                 # mm-yyyy or "-"
    pages: str                  # "1-2" etc
    output_file: str            # filename or "-"
    output_path: str            # absolute path or "-"
    note: str                   # extra info


def export_csv(records: list[Record], csv_path: Path):
    csv_path.parent.mkdir(parents=True, exist_ok=True)
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["status", "name", "period", "pages", "output_file", "output_path", "note"])
        for r in records:
            w.writerow([r.status, r.name, r.period, r.pages, r.output_file, r.output_path, r.note])


# ---------------- Core split logic (multipage + UI-friendly records) ----------------

def split_pdf(
    input_pdf: Path,
    ok_dir: Path,
    err_dir: Path,
    logger: logging.Logger,
    group_multipage: bool = True,
    progress_cb: Optional[Callable[[int, int], None]] = None,
):
    reader = PdfReader(str(input_pdf))
    total_pages = len(reader.pages)

    records: list[Record] = []

    logger.info(f"📄 Fichier: {input_pdf.name}")
    logger.info(f"📌 Pages: {total_pages}")
    logger.info(f"📁 OK: {ok_dir}")
    logger.info(f"🧨 ERRORS: {err_dir}")
    logger.info(f"🧩 Multi-pages: {'ON' if group_multipage else 'OFF'}")
    logger.info("—" * 72)

    ok_files = 0
    fallback_pages = 0
    errors = 0
    orphan_pages = 0

    # ---- Mode 1: 1 page = 1 file
    if not group_multipage:
        for i in range(total_pages):
            page_no = i + 1
            try:
                text = reader.pages[i].extract_text() or ""
                nm = extract_name(text)
                pr = extract_period(text)

                if nm and pr:
                    filename = f"{clean_filename(nm.title())}_{pr}.pdf"
                    out_path = ok_dir / filename
                    if out_path.exists():
                        out_path = ok_dir / f"{clean_filename(nm.title())}_{pr}_page{page_no:03d}.pdf"
                    write_pages(reader, [i], out_path)
                    ok_files += 1
                    logger.info(f"✅ Page {page_no}/{total_pages} -> OK -> {out_path.name}")

                    records.append(Record(
                        status="OK",
                        name=nm.title(),
                        period=pr,
                        pages=f"{page_no}",
                        output_file=out_path.name,
                        output_path=str(out_path.resolve()),
                        note="",
                    ))
                else:
                    out_path = err_dir / f"fiche_page_{page_no:03d}.pdf"
                    write_pages(reader, [i], out_path)
                    fallback_pages += 1
                    logger.warning(f"⚠️ Page {page_no}: nom/période non détectés -> errors -> {out_path.name}")

                    records.append(Record(
                        status="FALLBACK",
                        name=nm.title() if nm else "-",
                        period=pr if pr else "-",
                        pages=f"{page_no}",
                        output_file=out_path.name,
                        output_path=str(out_path.resolve()),
                        note="nom/période non détectés",
                    ))

            except Exception as e:
                errors += 1
                out_path = err_dir / f"error_page_{page_no:03d}.pdf"
                try:
                    write_pages(reader, [i], out_path)
                except Exception:
                    out_path = Path("-")
                logger.error(f"❌ Page {page_no}: {type(e).__name__} - {e}")

                records.append(Record(
                    status="ERROR",
                    name="-",
                    period="-",
                    pages=f"{page_no}",
                    output_file=out_path.name if out_path != Path("-") else "-",
                    output_path=str(out_path.resolve()) if out_path != Path("-") else "-",
                    note=f"{type(e).__name__}: {e}",
                ))

            if progress_cb:
                progress_cb(page_no, total_pages)

        return {
            "pages": total_pages,
            "ok_files": ok_files,
            "fallback_pages": fallback_pages,
            "errors": errors,
            "orphans": orphan_pages,
            "records": records
        }

    # ---- Mode 2: group multi-pages
    current_pages: list[int] = []
    current_name: Optional[str] = None
    current_period: Optional[str] = None
    current_start_page: Optional[int] = None

    def flush_current():
        nonlocal ok_files, fallback_pages, errors, current_pages, current_name, current_period, current_start_page
        if not current_pages:
            return

        start_page = current_start_page if current_start_page else (current_pages[0] + 1)
        end_page = current_pages[-1] + 1
        pages_str = f"{start_page}-{end_page}" if start_page != end_page else f"{start_page}"

        if current_name and current_period:
            base = f"{clean_filename(current_name.title())}_{current_period}.pdf"
            out_path = ok_dir / base
            if out_path.exists():
                out_path = ok_dir / f"{clean_filename(current_name.title())}_{current_period}_p{start_page:03d}.pdf"
            try:
                write_pages(reader, current_pages, out_path)
                ok_files += 1
                logger.info(f"✅ Fiche pages {pages_str} -> OK -> {out_path.name}")

                records.append(Record(
                    status="OK",
                    name=current_name.title(),
                    period=current_period,
                    pages=pages_str,
                    output_file=out_path.name,
                    output_path=str(out_path.resolve()),
                    note="",
                ))
            except Exception as e:
                errors += 1
                out_err = err_dir / f"error_slip_p{start_page:03d}.pdf"
                try:
                    write_pages(reader, current_pages, out_err)
                    out_path_str = str(out_err.resolve())
                    out_file = out_err.name
                except Exception:
                    out_path_str = "-"
                    out_file = "-"
                logger.error(f"❌ Fiche p{start_page:03d}: {type(e).__name__} - {e}")

                records.append(Record(
                    status="ERROR",
                    name=current_name.title(),
                    period=current_period,
                    pages=pages_str,
                    output_file=out_file,
                    output_path=out_path_str,
                    note=f"{type(e).__name__}: {e}",
                ))
        else:
            out_err = err_dir / f"unknown_slip_p{start_page:03d}.pdf"
            try:
                write_pages(reader, current_pages, out_err)
                fallback_pages += len(current_pages)
                logger.warning(f"⚠️ Fiche pages {pages_str}: nom/période non détectés -> errors -> {out_err.name}")

                records.append(Record(
                    status="FALLBACK",
                    name=current_name.title() if current_name else "-",
                    period=current_period if current_period else "-",
                    pages=pages_str,
                    output_file=out_err.name,
                    output_path=str(out_err.resolve()),
                    note="nom/période non détectés",
                ))
            except Exception as e:
                errors += 1
                logger.error(f"❌ Fiche inconnue p{start_page:03d}: {type(e).__name__} - {e}")

                records.append(Record(
                    status="ERROR",
                    name="-",
                    period="-",
                    pages=pages_str,
                    output_file="-",
                    output_path="-",
                    note=f"{type(e).__name__}: {e}",
                ))

        current_pages = []
        current_name = None
        current_period = None
        current_start_page = None

    for i in range(total_pages):
        page_no = i + 1
        try:
            text = reader.pages[i].extract_text() or ""

            if is_new_payslip_page(text):
                # new slip starts -> flush previous
                if current_pages:
                    flush_current()

                current_pages = [i]
                current_name = extract_name(text)
                current_period = extract_period(text)
                current_start_page = page_no
            else:
                if current_pages:
                    current_pages.append(i)
                else:
                    # orphan page before any slip header
                    out_err = err_dir / f"orphan_page_{page_no:03d}.pdf"
                    write_pages(reader, [i], out_err)
                    orphan_pages += 1
                    fallback_pages += 1
                    logger.warning(f"⚠️ Page {page_no}: page isolée (pas de début fiche) -> errors -> {out_err.name}")

                    records.append(Record(
                        status="ORPHAN",
                        name="-",
                        period="-",
                        pages=f"{page_no}",
                        output_file=out_err.name,
                        output_path=str(out_err.resolve()),
                        note="page isolée avant toute fiche",
                    ))

        except Exception as e:
            errors += 1
            out_err = err_dir / f"error_page_{page_no:03d}.pdf"
            try:
                write_pages(reader, [i], out_err)
                out_path_str = str(out_err.resolve())
                out_file = out_err.name
            except Exception:
                out_path_str = "-"
                out_file = "-"
            logger.error(f"❌ Page {page_no}: {type(e).__name__} - {e}")

            records.append(Record(
                status="ERROR",
                name="-",
                period="-",
                pages=f"{page_no}",
                output_file=out_file,
                output_path=out_path_str,
                note=f"{type(e).__name__}: {e}",
            ))

        if progress_cb:
            progress_cb(page_no, total_pages)

    flush_current()

    logger.info("—" * 72)
    logger.info(f"📦 Fichiers OK: {ok_files}")
    logger.info(f"⚠️ Pages fallback (errors/orphans): {fallback_pages}")
    logger.info(f"🧩 Orphans: {orphan_pages}")
    logger.info(f"❌ Erreurs techniques: {errors}")
    logger.info("🎉 Terminé.")

    return {
        "pages": total_pages,
        "ok_files": ok_files,
        "fallback_pages": fallback_pages,
        "orphans": orphan_pages,
        "errors": errors,
        "records": records
    }


# ---------------- Professional UI (table + CSV + open selected) ----------------

class AppUI(ttk.Frame):
    def __init__(self, master: tk.Tk, root: Path):
        super().__init__(master, padding=16)
        self.master = master
        self.root = root
        self.pack(fill="both", expand=True)

        self.pdf_var = tk.StringVar()
        self.multipage_var = tk.BooleanVar(value=True)

        self.timestamp: Optional[str] = None
        self.ok_dir: Optional[Path] = None
        self.err_dir: Optional[Path] = None
        self.log_path: Optional[Path] = None
        self.csv_path: Optional[Path] = None
        self.records: list[Record] = []

        self._style()
        self._build()

    def _style(self):
        self.master.title("Split fiches de salaire — Outil RH")
        self.master.geometry("980x620")
        self.master.resizable(False, False)

        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("Title.TLabel", font=("Segoe UI", 13, "bold"))
        style.configure("Hint.TLabel", foreground="#666")
        style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"))

    def _build(self):
        ttk.Label(self, text="Split fiches de salaire (PDF)", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            self,
            text="Sorties dans split-fiches-salaire/ : output/ logs/ errors/ + export CSV de contrôle.",
            style="Hint.TLabel",
        ).grid(row=1, column=0, columnspan=6, sticky="w", pady=(2, 12))

        # File picker
        file_frame = ttk.LabelFrame(self, text="Fichier source", padding=12)
        file_frame.grid(row=2, column=0, columnspan=6, sticky="we")

        entry = ttk.Entry(file_frame, textvariable=self.pdf_var, width=110)
        entry.grid(row=0, column=0, sticky="we")
        ttk.Button(file_frame, text="Choisir…", command=self.pick_pdf).grid(row=0, column=1, padx=(10, 0))
        file_frame.columnconfigure(0, weight=1)

        # Options
        opt_frame = ttk.LabelFrame(self, text="Options", padding=12)
        opt_frame.grid(row=3, column=0, columnspan=6, sticky="we", pady=(12, 0))

        ttk.Checkbutton(
            opt_frame,
            text="Grouper les fiches multi-pages (recommandé)",
            variable=self.multipage_var
        ).grid(row=0, column=0, sticky="w")

        # Actions row
        self.run_btn = ttk.Button(self, text="Lancer le traitement", style="Primary.TButton", command=self.run)
        self.run_btn.grid(row=4, column=0, sticky="w", pady=(12, 0))

        self.open_output_btn = ttk.Button(self, text="Ouvrir output", command=self.open_output, state="disabled")
        self.open_output_btn.grid(row=4, column=1, sticky="w", padx=(10, 0), pady=(12, 0))

        self.open_errors_btn = ttk.Button(self, text="Ouvrir errors", command=self.open_errors, state="disabled")
        self.open_errors_btn.grid(row=4, column=2, sticky="w", padx=(10, 0), pady=(12, 0))

        self.open_logs_btn = ttk.Button(self, text="Ouvrir logs", command=self.open_logs, state="disabled")
        self.open_logs_btn.grid(row=4, column=3, sticky="w", padx=(10, 0), pady=(12, 0))

        self.export_csv_btn = ttk.Button(self, text="Exporter CSV", command=self.export_csv_ui, state="disabled")
        self.export_csv_btn.grid(row=4, column=4, sticky="w", padx=(10, 0), pady=(12, 0))

        self.open_selected_btn = ttk.Button(self, text="Ouvrir sélection", command=self.open_selected, state="disabled")
        self.open_selected_btn.grid(row=4, column=5, sticky="w", padx=(10, 0), pady=(12, 0))

        # Progress
        self.status_var = tk.StringVar(value="Prêt.")
        ttk.Label(self, textvariable=self.status_var).grid(row=5, column=0, columnspan=6, sticky="w", pady=(12, 6))

        self.progress = ttk.Progressbar(self, mode="determinate", length=940)
        self.progress.grid(row=6, column=0, columnspan=6, sticky="we")

        # Table
        table_frame = ttk.LabelFrame(self, text="Résultats", padding=10)
        table_frame.grid(row=7, column=0, columnspan=6, sticky="we", pady=(14, 0))

        cols = ("status", "name", "period", "pages", "file", "note")
        self.tree = ttk.Treeview(table_frame, columns=cols, show="headings", height=14)
        self.tree.heading("status", text="Statut")
        self.tree.heading("name", text="Nom")
        self.tree.heading("period", text="Période")
        self.tree.heading("pages", text="Pages")
        self.tree.heading("file", text="Fichier")
        self.tree.heading("note", text="Note")

        self.tree.column("status", width=90, anchor="center")
        self.tree.column("name", width=210)
        self.tree.column("period", width=90, anchor="center")
        self.tree.column("pages", width=80, anchor="center")
        self.tree.column("file", width=260)
        self.tree.column("note", width=260)

        yscroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=yscroll.set)

        self.tree.grid(row=0, column=0, sticky="we")
        yscroll.grid(row=0, column=1, sticky="ns")

        table_frame.columnconfigure(0, weight=1)

        self.columnconfigure(0, weight=1)

    def pick_pdf(self):
        p = filedialog.askopenfilename(title="Choisir un PDF", filetypes=[("PDF", "*.pdf")])
        if p:
            self.pdf_var.set(p)

    def run(self):
        pdf = self.pdf_var.get().strip()
        if not pdf:
            messagebox.showwarning("PDF manquant", "Choisis un fichier PDF.")
            return

        input_pdf = Path(pdf).expanduser().resolve()
        if not input_pdf.exists():
            messagebox.showerror("Introuvable", f"Fichier introuvable:\n{input_pdf}")
            return

        # Reset table
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.records = []

        self.run_btn.config(state="disabled")
        self.open_output_btn.config(state="disabled")
        self.open_errors_btn.config(state="disabled")
        self.open_logs_btn.config(state="disabled")
        self.export_csv_btn.config(state="disabled")
        self.open_selected_btn.config(state="disabled")

        self.status_var.set("Traitement en cours…")
        self.progress["value"] = 0

        th = threading.Thread(target=self._do_work, args=(input_pdf,), daemon=True)
        th.start()

    def _do_work(self, input_pdf: Path):
        try:
            self.timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            ok_dir, err_dir, logs_dir = make_dirs(self.root, self.timestamp)
            self.ok_dir, self.err_dir = ok_dir, err_dir
            self.log_path = logs_dir / f"split_{self.timestamp}.log"
            self.csv_path = logs_dir / f"split_{self.timestamp}.csv"

            logger = setup_logger(self.log_path)

            def progress_cb(done, total):
                self.master.after(0, lambda: self._update_progress(done, total))

            result = split_pdf(
                input_pdf=input_pdf,
                ok_dir=ok_dir,
                err_dir=err_dir,
                logger=logger,
                group_multipage=self.multipage_var.get(),
                progress_cb=progress_cb,
            )

            self.records = result["records"]
            # Auto-export CSV
            export_csv(self.records, self.csv_path)

            self.master.after(0, lambda: self._finish(result))

        except Exception as e:
            self.master.after(0, lambda: self._fail(e))

    def _update_progress(self, done, total):
        self.progress["maximum"] = total
        self.progress["value"] = done
        self.status_var.set(f"Traitement… {done}/{total}")

    def _finish(self, result: dict):
        # Fill table
        for r in self.records:
            self.tree.insert("", "end", values=(r.status, r.name, r.period, r.pages, r.output_file, r.note))

        self.run_btn.config(state="normal")
        self.open_output_btn.config(state="normal")
        self.open_errors_btn.config(state="normal")
        self.open_logs_btn.config(state="normal")
        self.export_csv_btn.config(state="normal")
        self.open_selected_btn.config(state="normal")

        self.status_var.set(
            f"Terminé — Fichiers OK: {result['ok_files']} | Fallback pages: {result['fallback_pages']} | Orphans: {result['orphans']} | Erreurs: {result['errors']}"
        )

        messagebox.showinfo(
            "Terminé",
            f"Fichiers OK: {result['ok_files']}\n"
            f"Pages fallback (errors/orphans): {result['fallback_pages']}\n"
            f"Orphans: {result['orphans']}\n"
            f"Erreurs techniques: {result['errors']}\n\n"
            f"Output: {self.ok_dir}\n"
            f"Errors: {self.err_dir}\n"
            f"Log: {self.log_path}\n"
            f"CSV: {self.csv_path}"
        )

    def _fail(self, e: Exception):
        self.run_btn.config(state="normal")
        self.status_var.set("Erreur.")
        messagebox.showerror("Erreur", f"{type(e).__name__}: {e}")

    def open_output(self):
        if self.ok_dir:
            open_folder(self.ok_dir)

    def open_errors(self):
        if self.err_dir:
            open_folder(self.err_dir)

    def open_logs(self):
        open_folder(self.root / "logs")

    def export_csv_ui(self):
        if not self.records:
            messagebox.showwarning("Aucun résultat", "Lance un traitement avant d'exporter.")
            return
        default = str(self.csv_path) if self.csv_path else str((self.root / "logs" / "split_export.csv").resolve())
        dest = filedialog.asksaveasfilename(
            title="Enregistrer le CSV",
            defaultextension=".csv",
            initialfile=Path(default).name,
            filetypes=[("CSV", "*.csv")]
        )
        if not dest:
            return
        export_csv(self.records, Path(dest))
        messagebox.showinfo("CSV exporté", f"CSV enregistré:\n{dest}")

    def open_selected(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Sélection", "Sélectionne une ligne dans le tableau.")
            return
        values = self.tree.item(sel[0], "values")
        # columns: status, name, period, pages, file, note
        filename = values[4]
        # find record matching that filename + pages
        pages = values[3]
        rec = next((r for r in self.records if r.output_file == filename and r.pages == pages), None)
        if rec and rec.output_path != "-" and Path(rec.output_path).exists():
            open_file(Path(rec.output_path))
        else:
            messagebox.showwarning("Impossible", "Fichier introuvable (ou non enregistré).")


# ---------------- CLI ----------------

def run_cli(pdf_path: str, multipage: bool):
    root = project_root()
    ensure_base_dirs(root)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    ok_dir, err_dir, logs_dir = make_dirs(root, ts)
    log_path = logs_dir / f"split_{ts}.log"
    csv_path = logs_dir / f"split_{ts}.csv"

    logger = setup_logger(log_path)
    result = split_pdf(Path(pdf_path).expanduser().resolve(), ok_dir, err_dir, logger, group_multipage=multipage)
    export_csv(result["records"], csv_path)

    logger.info(f"🧾 Log: {log_path}")
    logger.info(f"📄 CSV: {csv_path}")


def main():
    parser = argparse.ArgumentParser(description="Split fiches de salaire (UI si pas d'argument).")
    parser.add_argument("pdf", nargs="?", help="Chemin du PDF (sinon UI)")
    parser.add_argument("--no-multipage", action="store_true", help="Désactive le regroupement multi-pages")
    args = parser.parse_args()

    root = project_root()
    ensure_base_dirs(root)

    if args.pdf:
        run_cli(args.pdf, multipage=not args.no_multipage)
    else:
        app = tk.Tk()
        AppUI(app, root)
        app.mainloop()


if __name__ == "__main__":
    main()
