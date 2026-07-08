# -*- coding: utf-8 -*-
"""
Rep_Infra_App.py
================
Aplicación de escritorio (Tkinter) para generar el Reporte de Infraestructura.
El usuario sube dos archivos Excel (Semáforo AB y Semáforo Interconexión),
la app procesa los datos y muestra/descarga el resultado como CSV.

Punto de entrada único: ejecutar este archivo o el .exe generado con PyInstaller.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import traceback
import os
import sys
import pandas as pd
from datetime import datetime

# Asegura que el directorio del ejecutable esté en el path
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

sys.path.insert(0, BASE_DIR)

import infra_logic  # noqa: E402


# ──────────────────────────────────────────────────────────────────────────────
#  Paleta de colores
# ──────────────────────────────────────────────────────────────────────────────
BG_DARK      = "#0f0f1a"
BG_CARD      = "#1a1a2e"
BG_SURFACE   = "#16213e"
ACCENT       = "#4361ee"
ACCENT_HOVER = "#3a0ca3"
ACCENT2      = "#7209b7"
TEXT_PRI     = "#e2e8f0"
TEXT_MUT     = "#94a3b8"
SUCCESS      = "#06d6a0"
SUCCESS_H    = "#05b88a"
ERR_COLOR    = "#ef233c"
ROW_ODD      = "#1e1e35"
ROW_EVEN     = "#252540"
HEADER_BG    = "#4361ee"
UPLOAD_BG    = "#1e2a4a"
UPLOAD_HOVER = "#253560"
UPLOAD_OK    = "#0d3320"
UPLOAD_OK_H  = "#145230"

FONT_TITLE  = ("Segoe UI", 26, "bold")
FONT_SUB    = ("Segoe UI", 12)
FONT_BTN    = ("Segoe UI", 13, "bold")
FONT_TABLE  = ("Segoe UI", 10)
FONT_HEADER = ("Segoe UI", 10, "bold")

REPORT_NAME = "Reporte_infraestructura_Red Noroeste_SEM_2026"


# ──────────────────────────────────────────────────────────────────────────────
#  Componente: botón con hover
# ──────────────────────────────────────────────────────────────────────────────
class HoverButton(tk.Button):
    def __init__(self, master, normal: str, hover: str, **kw):
        super().__init__(master, bg=normal, activebackground=hover,
                         relief="flat", bd=0, cursor="hand2", **kw)
        self._n, self._h = normal, hover
        self.bind("<Enter>", lambda _: self.config(bg=self._h))
        self.bind("<Leave>", lambda _: self.config(bg=self._n))

    def set_enabled(self, enabled: bool):
        if enabled:
            self.config(state="normal", bg=self._n, cursor="hand2")
        else:
            self.config(state="disabled", bg="#2a2a40", cursor="arrow")


# ──────────────────────────────────────────────────────────────────────────────
#  Frame: Upload (pantalla principal)
# ──────────────────────────────────────────────────────────────────────────────
class UploadFrame(tk.Frame):
    def __init__(self, master, on_generate):
        super().__init__(master, bg=BG_DARK)
        self._on_generate = on_generate
        self._path_ab = None
        self._path_intx = None
        self._build()

    def _build(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        # Canvas decorativo de fondo
        canvas = tk.Canvas(self, bg=BG_DARK, highlightthickness=0)
        canvas.grid(row=0, column=0, sticky="nsew")

        def _draw_bg(event=None):
            canvas.delete("bg")
            w, h = canvas.winfo_width(), canvas.winfo_height()
            canvas.create_oval(w * .55, -h * .15, w * 1.25, h * .45,
                               fill="#1a1a40", outline="", tags="bg")
            canvas.create_oval(-w * .2, h * .5, w * .35, h * 1.15,
                               fill="#16213e", outline="", tags="bg")
            canvas.create_oval(w * .7, h * .6, w * 1.1, h * 1.05,
                               fill="#1a1a40", outline="", tags="bg")

        canvas.bind("<Configure>", _draw_bg)

        # ── Tarjeta central ──────────────────────────────────────────────────
        card = tk.Frame(canvas, bg=BG_CARD, padx=50, pady=44)
        card.place(relx=0.5, rely=0.5, anchor="center")

        # Título
        tk.Label(card, text="\U0001f4ca", font=("Segoe UI", 48),
                 bg=BG_CARD, fg=ACCENT).pack(pady=(0, 4))
        tk.Label(card, text="Reporte Infraestructura", font=("Segoe UI", 24, "bold"),
                 bg=BG_CARD, fg=TEXT_PRI).pack()
        tk.Label(card, text="Red Noroeste — Generador de Reportes",
                 font=FONT_SUB, bg=BG_CARD, fg=TEXT_MUT).pack(pady=(2, 28))

        # ── Card Upload: Semáforo AB ─────────────────────────────────────────
        tk.Label(card, text="📄  Semáforo AB", font=("Segoe UI", 11, "bold"),
                 bg=BG_CARD, fg=TEXT_PRI, anchor="w").pack(fill="x", pady=(0, 4))

        self._frame_ab = tk.Frame(card, bg=UPLOAD_BG, cursor="hand2", padx=20, pady=16)
        self._frame_ab.pack(fill="x", pady=(0, 18))

        self._icon_ab = tk.Label(self._frame_ab, text="\U0001f4c2", font=("Segoe UI", 22),
                                  bg=UPLOAD_BG, fg=ACCENT)
        self._icon_ab.pack(side="left", padx=(0, 12))

        self._lbl_ab = tk.Label(self._frame_ab, text="Click para seleccionar archivo .xlsx",
                                 font=("Segoe UI", 11), bg=UPLOAD_BG, fg=TEXT_MUT, anchor="w")
        self._lbl_ab.pack(side="left", fill="x", expand=True)

        for w in (self._frame_ab, self._icon_ab, self._lbl_ab):
            w.bind("<Button-1>", lambda _: self._pick_ab())
            w.bind("<Enter>", lambda _: self._hover_ab(True))
            w.bind("<Leave>", lambda _: self._hover_ab(False))

        # ── Card Upload: Semáforo Interconexión ──────────────────────────────
        tk.Label(card, text="📄  Semáforo Interconexión", font=("Segoe UI", 11, "bold"),
                 bg=BG_CARD, fg=TEXT_PRI, anchor="w").pack(fill="x", pady=(0, 4))

        self._frame_intx = tk.Frame(card, bg=UPLOAD_BG, cursor="hand2", padx=20, pady=16)
        self._frame_intx.pack(fill="x", pady=(0, 28))

        self._icon_intx = tk.Label(self._frame_intx, text="\U0001f4c2", font=("Segoe UI", 22),
                                    bg=UPLOAD_BG, fg=ACCENT)
        self._icon_intx.pack(side="left", padx=(0, 12))

        self._lbl_intx = tk.Label(self._frame_intx, text="Click para seleccionar archivo .xlsx",
                                   font=("Segoe UI", 11), bg=UPLOAD_BG, fg=TEXT_MUT, anchor="w")
        self._lbl_intx.pack(side="left", fill="x", expand=True)

        for w in (self._frame_intx, self._icon_intx, self._lbl_intx):
            w.bind("<Button-1>", lambda _: self._pick_intx())
            w.bind("<Enter>", lambda _: self._hover_intx(True))
            w.bind("<Leave>", lambda _: self._hover_intx(False))

        # ── Botón Generar Reporte ────────────────────────────────────────────
        self._btn_gen = HoverButton(card, normal=ACCENT, hover=ACCENT2,
                                     text="  \u25B6  GENERAR REPORTE  ", font=FONT_BTN,
                                     fg="white", padx=20, pady=14,
                                     command=self._submit)
        self._btn_gen.pack(fill="x")
        self._btn_gen.set_enabled(False)

        self._lbl_err = tk.Label(card, text="", font=("Segoe UI", 10),
                                  bg=BG_CARD, fg=ERR_COLOR)
        self._lbl_err.pack(pady=(10, 0))

    # ── Hover effects ────────────────────────────────────────────────────────
    def _hover_ab(self, enter: bool):
        bg = UPLOAD_OK_H if self._path_ab else (UPLOAD_HOVER if enter else UPLOAD_BG)
        for w in (self._frame_ab, self._icon_ab, self._lbl_ab):
            w.config(bg=bg)

    def _hover_intx(self, enter: bool):
        bg = UPLOAD_OK_H if self._path_intx else (UPLOAD_HOVER if enter else UPLOAD_BG)
        for w in (self._frame_intx, self._icon_intx, self._lbl_intx):
            w.config(bg=bg)

    # ── File pickers ─────────────────────────────────────────────────────────
    def _pick_ab(self):
        path = filedialog.askopenfilename(
            title="Seleccionar Semáforo AB",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")],
        )
        if path:
            self._path_ab = path
            name = os.path.basename(path)
            self._lbl_ab.config(text=f"✓  {name}", fg=SUCCESS)
            self._icon_ab.config(text="✅", fg=SUCCESS)
            bg = UPLOAD_OK
            for w in (self._frame_ab, self._icon_ab, self._lbl_ab):
                w.config(bg=bg)
        self._check_ready()

    def _pick_intx(self):
        path = filedialog.askopenfilename(
            title="Seleccionar Semáforo Interconexión",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")],
        )
        if path:
            self._path_intx = path
            name = os.path.basename(path)
            self._lbl_intx.config(text=f"✓  {name}", fg=SUCCESS)
            self._icon_intx.config(text="✅", fg=SUCCESS)
            bg = UPLOAD_OK
            for w in (self._frame_intx, self._icon_intx, self._lbl_intx):
                w.config(bg=bg)
        self._check_ready()

    def _check_ready(self):
        ready = self._path_ab is not None and self._path_intx is not None
        self._btn_gen.set_enabled(ready)

    def _submit(self):
        if not self._path_ab or not self._path_intx:
            self._lbl_err.config(text="⚠  Selecciona ambos archivos.")
            return
        self._lbl_err.config(text="")
        self._on_generate(self._path_ab, self._path_intx)


# ──────────────────────────────────────────────────────────────────────────────
#  Frame: Carga (barra de progreso)
# ──────────────────────────────────────────────────────────────────────────────
class LoadingFrame(tk.Frame):
    def __init__(self, master):
        super().__init__(master, bg=BG_DARK)
        self._build()

    def _build(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        center = tk.Frame(self, bg=BG_DARK)
        center.place(relx=0.5, rely=0.5, anchor="center")

        tk.Label(center, text="⚙", font=("Segoe UI", 60),
                 bg=BG_DARK, fg=ACCENT).pack()
        tk.Label(center, text="Procesando información...",
                 font=("Segoe UI", 20, "bold"), bg=BG_DARK, fg=TEXT_PRI).pack(pady=(10, 4))

        self._lbl_status = tk.Label(center, text="Iniciando...",
                                    font=FONT_SUB, bg=BG_DARK, fg=TEXT_MUT)
        self._lbl_status.pack(pady=(0, 18))

        # Barra de progreso
        sty = ttk.Style()
        sty.theme_use("clam")
        sty.configure("App.Horizontal.TProgressbar",
                       troughcolor=BG_SURFACE, background=ACCENT,
                       lightcolor=ACCENT, darkcolor=ACCENT2,
                       bordercolor=BG_DARK, thickness=10)
        self._bar = ttk.Progressbar(center, mode="indeterminate", length=440,
                                     style="App.Horizontal.TProgressbar")
        self._bar.pack()

        # Log en tiempo real
        lf = tk.Frame(center, bg=BG_CARD, padx=2, pady=2)
        lf.pack(fill="x", pady=(20, 0))

        self._log = tk.Text(lf, height=12, width=68, bg=BG_SURFACE, fg=TEXT_MUT,
                             font=("Consolas", 9), relief="flat", state="disabled", wrap="word")
        sb = tk.Scrollbar(lf, command=self._log.yview, bg=BG_CARD, troughcolor=BG_SURFACE)
        self._log.config(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self._log.pack(side="left", fill="both", expand=True)

    # ── API pública ──────────────────────────────────────────────────────────
    def start(self):
        self._bar.start(12)

    def stop(self):
        self._bar.stop()

    def set_status(self, msg: str):
        self._lbl_status.config(text=msg)

    def append_log(self, line: str):
        self._log.config(state="normal")
        self._log.insert("end", line if line.endswith("\n") else line + "\n")
        self._log.see("end")
        self._log.config(state="disabled")


# ──────────────────────────────────────────────────────────────────────────────
#  Frame: Resultados
# ──────────────────────────────────────────────────────────────────────────────
class ResultsFrame(tk.Frame):
    def __init__(self, master, df: pd.DataFrame, on_new):
        super().__init__(master, bg=BG_DARK)
        self._df     = df
        self._on_new = on_new
        self._build()

    def _build(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        # ── Header ─────────────────────────────────────────────────────────────
        hdr = tk.Frame(self, bg=BG_CARD, pady=15, padx=20)
        hdr.grid(row=0, column=0, sticky="ew")
        hdr.columnconfigure(1, weight=1)

        HoverButton(hdr, normal=BG_SURFACE, hover=ACCENT,
                    text="←  Nuevo Reporte", font=("Segoe UI", 11),
                    fg=TEXT_PRI, padx=12, pady=6,
                    command=self._on_new).grid(row=0, column=0, padx=(0, 20))

        tk.Label(hdr, text=f"\U0001f4cb  {REPORT_NAME}",
                 font=("Segoe UI", 16, "bold"), bg=BG_CARD, fg=TEXT_PRI).grid(
                     row=0, column=1, sticky="w")

        HoverButton(hdr, normal=SUCCESS, hover=SUCCESS_H,
                    text="⬇  Descargar CSV", font=("Segoe UI", 11, "bold"),
                    fg="#000c0a", padx=14, pady=6,
                    command=self._download).grid(row=0, column=2, padx=(20, 0))

        # ── Info bar ──────────────────────────────────────────────────────────
        info = tk.Frame(self, bg=BG_SURFACE, pady=8, padx=20)
        info.grid(row=1, column=0, sticky="ew")

        tk.Label(info, text=f"✓  Reporte generado exitosamente  ·  {len(self._df)} registros  ·  {len(self._df.columns)} columnas",
                 font=("Segoe UI", 11), bg=BG_SURFACE, fg=SUCCESS).pack(side="left")

        ts = datetime.now().strftime("%d/%m/%Y  %H:%M")
        tk.Label(info, text=f"Generado: {ts}",
                 font=("Segoe UI", 10), bg=BG_SURFACE, fg=TEXT_MUT).pack(side="right")

        # ── Tabla preview ────────────────────────────────────────────────────
        tf = tk.Frame(self, bg=BG_DARK, padx=12, pady=8)
        tf.grid(row=2, column=0, sticky="nsew")
        tf.columnconfigure(0, weight=1)
        tf.rowconfigure(0, weight=1)

        sty = ttk.Style()
        sty.theme_use("clam")
        sty.configure("Res.Treeview",
                       background=ROW_ODD, foreground=TEXT_PRI,
                       rowheight=24, fieldbackground=ROW_ODD,
                       font=FONT_TABLE, borderwidth=0)
        sty.configure("Res.Treeview.Heading",
                       background=HEADER_BG, foreground="white",
                       font=FONT_HEADER, relief="flat", padding=(5, 3))
        sty.map("Res.Treeview", background=[("selected", ACCENT2)])
        sty.map("Res.Treeview.Heading", background=[("active", ACCENT2)])

        cols = list(self._df.columns)
        tree = ttk.Treeview(tf, columns=cols, show="headings", style="Res.Treeview")

        n_filas = len(self._df)
        for col in cols:
            tree.heading(col, text=col)
            max_w = int(max(
                len(str(col)) * 9,
                (self._df[col].astype(str).str.len().max() * 8)
                if n_filas else 80,
            ))
            tree.column(col, width=min(max_w, 170), anchor="center", minwidth=55)

        tree.tag_configure("odd",  background=ROW_ODD,  foreground=TEXT_PRI)
        tree.tag_configure("even", background=ROW_EVEN, foreground=TEXT_PRI)

        for idx, row in self._df.iterrows():
            tag = "odd" if idx % 2 == 0 else "even"
            tree.insert("", "end", values=list(row), tags=(tag,))

        vsb = ttk.Scrollbar(tf, orient="vertical",   command=tree.yview)
        hsb = ttk.Scrollbar(tf, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

    # ── Descarga CSV ─────────────────────────────────────────────────────────
    def _download(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv"), ("Excel", "*.xlsx")],
            initialfile=f"Reporte_infraestructura_Red Noroeste_SEM_2026.csv",
            title="Guardar reporte como...",
        )
        if not path:
            return
        try:
            if path.endswith(".xlsx"):
                self._df.to_excel(path, index=False, engine="openpyxl")
            else:
                self._df.to_csv(path, index=False, encoding="utf-8-sig")
            messagebox.showinfo("Guardado", f"Archivo guardado en:\n{path}")
        except Exception as exc:
            messagebox.showerror("Error al guardar", str(exc))


# ──────────────────────────────────────────────────────────────────────────────
#  Aplicación principal
# ──────────────────────────────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Rep_Infra — Reporte Infraestructura")
        self.geometry("1150x740")
        self.minsize(900, 620)
        self.configure(bg=BG_DARK)

        # Ícono opcional
        ico = os.path.join(BASE_DIR, "icon.ico")
        if os.path.exists(ico):
            try:
                self.iconbitmap(ico)
            except Exception:
                pass

        self._df_reporte: pd.DataFrame | None = None
        self._frame: tk.Frame | None = None

        self._goto(UploadFrame(self, on_generate=self._on_generate))

    # ── Navegación ──────────────────────────────────────────────────────────
    def _goto(self, frame: tk.Frame):
        if self._frame:
            self._frame.destroy()
        self._frame = frame
        frame.pack(fill="both", expand=True)

    # ── Handlers de flujo ────────────────────────────────────────────────────
    def _on_generate(self, path_ab: str, path_intx: str):
        loading = LoadingFrame(self)
        self._goto(loading)
        loading.start()

        def _worker():
            try:
                resultado = infra_logic.run(
                    path_ab, path_intx,
                    on_log=lambda msg: self.after(0, lambda m=msg: loading.append_log(m)),
                )
                self.after(0, lambda: self._on_ok(resultado, loading))
            except Exception as exc:
                tb = traceback.format_exc()
                self.after(0, lambda: self._on_error(str(exc), tb))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_ok(self, resultado: dict, loading: LoadingFrame):
        loading.stop()
        loading.set_status("¡Listo! ✓")
        loading.append_log("\n✅ Reporte generado exitosamente.")
        self._df_reporte = resultado["df_reporte"]
        # Pequeña pausa para que el usuario vea "Listo"
        self.after(800, self._show_results)

    def _on_error(self, msg: str, tb: str):
        print(tb)
        messagebox.showerror(
            "Error durante el procesamiento",
            f"{msg}\n\nRevisa la consola para más detalles.",
        )
        self._goto(UploadFrame(self, on_generate=self._on_generate))

    def _show_results(self):
        if self._df_reporte is None or self._df_reporte.empty:
            messagebox.showwarning("Sin datos", "El reporte está vacío. Verifica los archivos.")
            self._goto(UploadFrame(self, on_generate=self._on_generate))
            return
        self._goto(ResultsFrame(self, self._df_reporte, on_new=self._show_upload))

    def _show_upload(self):
        self._goto(UploadFrame(self, on_generate=self._on_generate))


# ──────────────────────────────────────────────────────────────────────────────
#  Entry point
# ──────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    App().mainloop()
