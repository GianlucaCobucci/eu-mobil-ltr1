import os
import re
import sys
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import openpyxl
from openpyxl.styles import Alignment, Border, Font as XLFont, PatternFill, Side
import pdfplumber

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    _DND = True
except ImportError:
    _DND = False

# ── Field definitions ──────────────────────────────────────────────────────
FIELDS = [
    ("scheda",               "Scheda"),
    ("cognome",              "Cognome"),
    ("nome",                 "Nome"),
    ("data_nascita",         "Data di Nascita"),
    ("stato_membro",         "Stato Membro"),
    ("data_com_stato_estero","Data Com. Stato Estero"),
    ("numero_permesso",      "Numero Permesso"),
    ("data_inizio_validita", "Data Inizio Validità"),
    ("data_fine_validita",   "Data Fine Validità"),
]

REQUIRED = {"cognome", "nome", "numero_permesso", "data_fine_validita"}

# ── PDF extraction ─────────────────────────────────────────────────────────

def _get(text: str, pattern: str) -> str:
    """Return value after label pattern on same line, cleaned."""
    m = re.search(pattern + r"[:\s]+([^\n]+)", text, re.IGNORECASE)
    if not m:
        return ""
    val = m.group(1).strip()
    return re.split(r"\s{3,}", val)[0].strip()


def _make_scheda(cognome: str, nome: str, data_nascita: str, numero: str) -> str:
    parts = []
    if cognome:
        parts.append(cognome.upper())
    if nome:
        parts.append(nome.title())
    if data_nascita:
        parts.append(data_nascita)
    if numero:
        parts.append(f"- {numero}")
    return " ".join(parts)


def extract_fields(pdf_path: str):
    try:
        pages = []
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                pages.append(page.extract_text(x_tolerance=3, y_tolerance=3) or "")
    except Exception as exc:
        return None, str(exc)

    p1 = pages[0] if pages else ""
    p2 = pages[1] if len(pages) > 1 else ""

    cognome      = _get(p1, r"Surnames")
    nome         = _get(p1, r"First\s+names")
    stato_membro = _get(p1, r"Member\s+State")
    data_nascita = _get(p1, r"Date\s+of\s+birth")
    # "Date:" alone (not "Date of birth")
    data_com     = _get(p1, r"(?<!\w)Date(?!\s+of\s+birth)")
    # Number of authorisation from page 1 (1st MS permit)
    numero       = _get(p1, r"Number\s+of\s+authorisation(?!\s+in\s+2nd)")

    # Option A block on page 2
    m = re.search(r"Option\s+A:.*?(?=Option\s+B:|$)", p2, re.DOTALL | re.IGNORECASE)
    a = m.group(0) if m else p2
    data_inizio  = _get(a, r"Issuance\s+date\s*(?:\(optional\))?")
    data_fine    = _get(a, r"(?<!\w)Validity")

    scheda = _make_scheda(cognome, nome, data_nascita, numero)

    return {
        "scheda":                scheda,
        "cognome":               cognome,
        "nome":                  nome,
        "data_nascita":          data_nascita,
        "stato_membro":          stato_membro,
        "data_com_stato_estero": data_com,
        "numero_permesso":       numero,
        "data_inizio_validita":  data_inizio,
        "data_fine_validita":    data_fine,
        "_filename":             os.path.basename(pdf_path),
    }, None


# ── Excel export ───────────────────────────────────────────────────────────
_C_BLUE   = "1F4E79"
_C_LBLUE  = "D6E4F0"
_C_WHITE  = "FFFFFF"
_C_YELLOW = "FFF2CC"
_C_RED    = "FCE4D6"


def _bdr():
    s = Side(style="thin")
    return Border(left=s, right=s, top=s, bottom=s)


def open_file(path: str):
    if sys.platform == "darwin":
        subprocess.run(["open", path])
    elif sys.platform == "win32":
        os.startfile(path)
    else:
        subprocess.run(["xdg-open", path])


def save_excel(records: list, output_path: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "EU-MOBIL LTR1"

    headers = ["File PDF", "Note"] + [lbl for _, lbl in FIELDS]
    widths  = [30, 22, 42, 18, 15, 18, 20, 22, 18, 22, 22]

    for col, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=col, value=h)
        c.fill      = PatternFill("solid", fgColor=_C_BLUE)
        c.font      = XLFont(bold=True, color=_C_WHITE, size=10)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border    = _bdr()
    ws.row_dimensions[1].height = 30

    alt  = PatternFill("solid", fgColor=_C_LBLUE)
    yell = PatternFill("solid", fgColor=_C_YELLOW)
    red  = PatternFill("solid", fgColor=_C_RED)

    for ri, rec in enumerate(records, 2):
        missing = [k for k in REQUIRED if not rec.get(k)]
        note    = "Campo mancante: " + ", ".join(missing) if missing else ""
        base    = alt if ri % 2 == 0 else PatternFill()
        vals    = [rec.get("_filename", ""), note] + [rec.get(k, "") for k, _ in FIELDS]

        for ci, val in enumerate(vals, 1):
            c = ws.cell(row=ri, column=ci, value=val)
            c.alignment = Alignment(horizontal="left", vertical="center")
            c.border    = _bdr()
            if ci == 2 and val:
                c.fill = red
            elif ci > 2:
                key = FIELDS[ci - 3][0] if ci - 3 < len(FIELDS) else None
                c.fill = yell if (key in REQUIRED and not val) else base
            else:
                c.fill = base
        ws.row_dimensions[ri].height = 18

    for ci, w in enumerate(widths, 1):
        ws.column_dimensions[ws.cell(row=1, column=ci).column_letter].width = w
    ws.freeze_panes = "A2"
    wb.save(output_path)


# ── Stili ttk (tema clam = colori personalizzabili su Mac e Windows) ────────

def _setup_styles():
    s = ttk.Style()
    s.theme_use("clam")

    defs = {
        "Add.TButton":     ("#1F6FBF", "#164F8C"),
        "Folder.TButton":  ("#0288D1", "#0277BD"),
        "Remove.TButton":  ("#5C5C5C", "#444444"),
        "Clear.TButton":   ("#7B1FA2", "#6A1B9A"),
        "Preview.TButton": ("#BF4E00", "#A84300"),
        "Export.TButton":  ("#276221", "#1E5219"),
        "Save.TButton":    ("#2E7D32", "#1B5E20"),
    }
    for name, (bg, active) in defs.items():
        s.configure(name, background=bg, foreground="white",
                    font=("Helvetica", 10, "bold"),
                    padding=(12, 6), borderwidth=0, relief="flat",
                    focuscolor=bg)
        s.map(name,
              background=[("active", active), ("pressed", active)],
              foreground=[("active", "white"), ("pressed", "white")])


# ── Preview window ─────────────────────────────────────────────────────────

class PreviewWindow(tk.Toplevel):
    BG = "#F5F7FA"

    def __init__(self, parent, records: list, on_save):
        super().__init__(parent)
        self.title("Anteprima dati estratti")
        self.geometry("1200x440")
        self.configure(bg=self.BG)
        self.records = records
        self.on_save = on_save
        self._build()

    def _build(self):
        cols = ["File PDF"] + [lbl for _, lbl in FIELDS]

        frm = tk.Frame(self, bg=self.BG)
        frm.pack(fill="both", expand=True, padx=10, pady=8)

        xs = ttk.Scrollbar(frm, orient="horizontal")
        ys = ttk.Scrollbar(frm, orient="vertical")
        tree = ttk.Treeview(
            frm, columns=cols, show="headings",
            xscrollcommand=xs.set, yscrollcommand=ys.set, height=14,
        )
        col_widths = {"File PDF": 200, "Scheda": 260, "Cognome": 110,
                      "Nome": 110, "Data di Nascita": 110}
        for col in cols:
            tree.heading(col, text=col)
            tree.column(col, width=col_widths.get(col, 120), minwidth=80)

        for rec in self.records:
            has_missing = any(not rec.get(k) for k in REQUIRED)
            vals = [rec.get("_filename", "")] + [rec.get(k, "") for k, _ in FIELDS]
            tree.insert("", "end", values=vals, tags=("warn",) if has_missing else ("ok",))

        tree.tag_configure("warn", background="#FFF2CC")
        tree.tag_configure("ok",   background="#FFFFFF")

        xs.config(command=tree.xview)
        ys.config(command=tree.yview)
        tree.grid(row=0, column=0, sticky="nsew")
        ys.grid(row=0, column=1, sticky="ns")
        xs.grid(row=1, column=0, sticky="ew")
        frm.grid_rowconfigure(0, weight=1)
        frm.grid_columnconfigure(0, weight=1)

        tk.Label(self, text="  Giallo = campo obbligatorio mancante",
                 font=("Helvetica", 9), fg="#888", bg=self.BG,
                 anchor="w").pack(anchor="w", padx=12)

        ttk.Button(self, text="Salva Excel", command=self._save,
                   style="Save.TButton").pack(pady=10)

    def _save(self):
        path = filedialog.asksaveasfilename(
            title="Salva Excel come...",
            defaultextension=".xlsx",
            filetypes=[("File Excel", "*.xlsx")],
            initialfile="EU_MOBIL_LTR1_export.xlsx",
        )
        if path:
            self.on_save(path)
            self.destroy()


# ── Main application ───────────────────────────────────────────────────────

BG_WIN   = "#F5F7FA"   # sfondo finestra
BG_LIST  = "#FFFFFF"   # sfondo listbox
FG_TITLE = "#1A1A2E"   # titolo scuro


class App:
    MAX = 50

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("EU-MOBIL LTR1 — Estrattore Dati")
        self.root.geometry("660x490")
        self.root.resizable(False, False)
        self.root.configure(bg=BG_WIN)
        self.files: list[str] = []
        self._build()
        if _DND:
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind("<<Drop>>", self._on_drop)

    def _build(self):
        # ── Intestazione ──────────────────────────────────────────────────
        hdr = tk.Frame(self.root, bg="#1F4E79")
        hdr.pack(fill="x")
        tk.Label(hdr,
                 text="EU-MOBIL LTR1  —  Estrazione Dati PDF → Excel",
                 font=("Helvetica", 13, "bold"),
                 bg="#1F4E79", fg="white", pady=12).pack()

        # ── Area file ─────────────────────────────────────────────────────
        lf_title = f"File PDF (max {self.MAX})" + \
                   ("  —  puoi trascinare i PDF qui" if _DND else "")
        frm = tk.LabelFrame(self.root, text=lf_title,
                            bg=BG_WIN, fg="#333", font=("Helvetica", 9),
                            padx=8, pady=6)
        frm.pack(fill="both", expand=True, padx=18, pady=(10, 4))

        lf = tk.Frame(frm, bg=BG_WIN)
        lf.pack(fill="both", expand=True)
        sb = ttk.Scrollbar(lf)
        sb.pack(side="right", fill="y")
        self.lb = tk.Listbox(lf, font=("Helvetica", 9),
                             selectmode=tk.EXTENDED,
                             bg=BG_LIST, fg="#222",
                             selectbackground="#1F4E79",
                             selectforeground="white",
                             activestyle="none",
                             yscrollcommand=sb.set,
                             relief="solid", bd=1)
        self.lb.pack(fill="both", expand=True)
        sb.config(command=self.lb.yview)

        # ── Pulsanti riga 1 ───────────────────────────────────────────────
        r1 = tk.Frame(self.root, bg=BG_WIN)
        r1.pack(pady=(8, 2))
        ttk.Button(r1, text="+ Aggiungi PDF",    command=self.add_files,  style="Add.TButton"   ).pack(side="left", padx=4)
        ttk.Button(r1, text="Aggiungi Cartella", command=self.add_folder, style="Folder.TButton").pack(side="left", padx=4)
        ttk.Button(r1, text="Rimuovi",           command=self.remove,     style="Remove.TButton").pack(side="left", padx=4)
        ttk.Button(r1, text="Pulisci tutto",     command=self.clear,      style="Clear.TButton" ).pack(side="left", padx=4)

        # ── Progress bar ──────────────────────────────────────────────────
        self.prog_val = tk.DoubleVar()
        self.prog = ttk.Progressbar(self.root, variable=self.prog_val,
                                    maximum=100, length=500, mode="determinate")

        # ── Pulsanti riga 2 ───────────────────────────────────────────────
        r2 = tk.Frame(self.root, bg=BG_WIN)
        r2.pack(pady=(4, 2))
        ttk.Button(r2, text="Anteprima Dati", command=self.preview, style="Preview.TButton").pack(side="left", padx=6)
        ttk.Button(r2, text="Estrai → Excel", command=self.process, style="Export.TButton" ).pack(side="left", padx=6)

        # ── Barra stato ───────────────────────────────────────────────────
        self.status = tk.StringVar(value="Pronto. Aggiungi PDF o una cartella.")
        tk.Label(self.root, textvariable=self.status,
                 font=("Helvetica", 9), fg="#555", bg=BG_WIN).pack(pady=4)

    def _on_drop(self, event):
        for m in re.finditer(r'\{([^}]+)\}|(\S+)', event.data):
            path = m.group(1) or m.group(2)
            if path.lower().endswith(".pdf"):
                self._add(path)
        self.status.set(f"{len(self.files)} PDF caricati.")

    def _add(self, path: str) -> bool:
        if len(self.files) >= self.MAX or path in self.files:
            return False
        self.files.append(path)
        self.lb.insert(tk.END, os.path.basename(path))
        return True

    def _focus(self):
        """Porta finestra + dialog in primo piano su Mac."""
        self.root.lift()
        self.root.attributes("-topmost", True)
        self.root.update()
        # topmost viene spento DOPO che il chiamante chiude il dialog

    def _unfocus(self):
        self.root.attributes("-topmost", False)

    def add_files(self):
        if len(self.files) >= self.MAX:
            messagebox.showwarning("Limite", f"Massimo {self.MAX} PDF.")
            return
        self._focus()
        files = filedialog.askopenfilenames(title="Seleziona PDF",
                                            filetypes=[("PDF", "*.pdf")])
        self._unfocus()
        for f in files:
            self._add(f)
        self.status.set(f"{len(self.files)} PDF caricati.")

    def add_folder(self):
        self._focus()
        folder = filedialog.askdirectory(title="Seleziona cartella")
        self._unfocus()
        if not folder:
            return
        pdfs = sorted(os.path.join(folder, f)
                      for f in os.listdir(folder) if f.lower().endswith(".pdf"))
        added = sum(1 for p in pdfs if self._add(p))
        self.status.set(f"{len(self.files)} PDF caricati. ({added} aggiunti dalla cartella)")

    def remove(self):
        for idx in sorted(self.lb.curselection(), reverse=True):
            self.lb.delete(idx)
            self.files.pop(idx)
        self.status.set(f"{len(self.files)} PDF caricati.")

    def clear(self):
        self.lb.delete(0, tk.END)
        self.files.clear()
        self.status.set("Lista svuotata.")

    def _run(self) -> tuple[list, list]:
        records, errors = [], []
        n = len(self.files)
        self.prog_val.set(0)
        # Mostra progress bar sopra la status label
        self.prog.pack(pady=3)
        self.root.update()
        for i, path in enumerate(self.files, 1):
            self.status.set(f"Elaboro {i}/{n}: {os.path.basename(path)}")
            self.root.update()
            rec, err = extract_fields(path)
            (errors if err else records).append(err if err else rec)
            self.prog_val.set(i / n * 100)
            self.root.update()
        self.prog.pack_forget()
        return records, errors

    def preview(self):
        if not self.files:
            messagebox.showwarning("Nessun file", "Aggiungi almeno un PDF.")
            return
        try:
            records, errors = self._run()
        except Exception as e:
            messagebox.showerror("Errore elaborazione", str(e))
            return
        if not records:
            messagebox.showerror("Errore", "Nessun dato estratto.\n\n" + "\n".join(errors))
            return

        def on_save(path):
            try:
                save_excel(records, path)
                n = len(records)
                messagebox.showinfo("Completato", f"{n} record esportati in:\n{path}")
                self.status.set(f"Completato. {n} record esportati.")
                open_file(path)
            except Exception as e:
                messagebox.showerror("Errore", str(e))

        PreviewWindow(self.root, records, on_save)
        if errors:
            self.status.set(f"Attenzione: {len(errors)} file con errori.")

    def process(self):
        if not self.files:
            messagebox.showwarning("Nessun file", "Aggiungi almeno un PDF.")
            return
        self._focus()
        path = filedialog.asksaveasfilename(
            title="Salva Excel come...", defaultextension=".xlsx",
            filetypes=[("File Excel", "*.xlsx")],
            initialfile="EU_MOBIL_LTR1_export.xlsx",
        )
        self._unfocus()
        if not path:
            return
        try:
            records, errors = self._run()
        except Exception as e:
            messagebox.showerror("Errore elaborazione", str(e))
            return
        if not records:
            messagebox.showerror("Errore", "Nessun dato estratto.\n\n" + "\n".join(errors))
            return
        try:
            save_excel(records, path)
        except Exception as e:
            messagebox.showerror("Errore salvataggio", str(e))
            return
        msg = f"{len(records)} record esportati in:\n{path}"
        if errors:
            msg += f"\n\nAttenzione — {len(errors)} errori:\n" + "\n".join(errors)
        messagebox.showinfo("Completato", msg)
        self.status.set(f"Completato. {len(records)} record esportati.")
        open_file(path)
        self.status.set(f"Completato. {len(records)} record esportati.")
        open_file(path)


# ── Entry point ────────────────────────────────────────────────────────────

def main():
    root = TkinterDnD.Tk() if _DND else tk.Tk()
    _setup_styles()
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.user32.ShowWindow(
                ctypes.windll.kernel32.GetConsoleWindow(), 0)
        except Exception:
            pass
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
