# -*- coding: utf-8 -*-
"""
Reportes.py
===========
Aplicación de escritorio (Tkinter) para el sistema de reportes.
Importa `scraper` para ejecutar la extracción y procesamiento de datos.

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
# (necesario cuando se corre como .exe compilado con PyInstaller)
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

sys.path.insert(0, BASE_DIR)

import scraper  # noqa: E402


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

FONT_TITLE  = ("Segoe UI", 26, "bold")
FONT_SUB    = ("Segoe UI", 12)
FONT_BTN    = ("Segoe UI", 13, "bold")
FONT_TABLE  = ("Segoe UI", 10)
FONT_HEADER = ("Segoe UI", 10, "bold")


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


# ──────────────────────────────────────────────────────────────────────────────
#  Frame: Login
# ──────────────────────────────────────────────────────────────────────────────
class LoginFrame(tk.Frame):
    def __init__(self, master, on_login):
        super().__init__(master, bg=BG_DARK)
        self._on_login = on_login
        self._build()

    # ── construcción ──────────────────────────────────────────────────────────
    def _build(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        # Canvas decorativo
        canvas = tk.Canvas(self, bg=BG_DARK, highlightthickness=0)
        canvas.grid(row=0, column=0, sticky="nsew")
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)

        def _draw_bg(event=None):
            canvas.delete("bg")
            w, h = canvas.winfo_width(), canvas.winfo_height()
            canvas.create_oval(w * .6, -h * .1, w * 1.2, h * .5,
                               fill="#1a1a40", outline="", tags="bg")
            canvas.create_oval(-w * .2, h * .5, w * .4, h * 1.1,
                               fill="#16213e", outline="", tags="bg")

        canvas.bind("<Configure>", _draw_bg)

        # ── Tarjeta central ───────────────────────────────────────────────────
        card = tk.Frame(canvas, bg=BG_CARD, padx=52, pady=48)
        card.place(relx=0.5, rely=0.5, anchor="center")

        tk.Label(card, text="\U0001f512", font=("Segoe UI", 52),
                 bg=BG_CARD, fg=ACCENT).pack(pady=(0, 4))
        tk.Label(card, text="Reportes", font=("Segoe UI", 28, "bold"),
                 bg=BG_CARD, fg=TEXT_PRI).pack()
        tk.Label(card, text="Ingresa tus credenciales para continuar",
                 font=FONT_SUB, bg=BG_CARD, fg=TEXT_MUT).pack(pady=(2, 26))

        # ── Campo Usuario ─────────────────────────────────────────────────────
        tk.Label(card, text="Usuario", font=("Segoe UI", 10, "bold"),
                 bg=BG_CARD, fg=TEXT_MUT, anchor="w").pack(fill="x")
        border_u = tk.Frame(card, bg=ACCENT, padx=2, pady=2)
        border_u.pack(fill="x", pady=(3, 16))
        self._ent_user = tk.Entry(
            border_u, font=("Segoe UI", 12),
            bg=BG_SURFACE, fg=TEXT_PRI,
            insertbackground=TEXT_PRI,
            relief="flat", highlightthickness=0, bd=0,
        )
        self._ent_user.pack(fill="x", ipady=8, padx=1, pady=1)

        # ── Campo Contrase\u00f1a ──────────────────────────────────────────────────
        tk.Label(card, text="Contrase\u00f1a", font=("Segoe UI", 10, "bold"),
                 bg=BG_CARD, fg=TEXT_MUT, anchor="w").pack(fill="x")
        border_p = tk.Frame(card, bg=ACCENT, padx=2, pady=2)
        border_p.pack(fill="x", pady=(3, 26))
        self._ent_pass = tk.Entry(
            border_p, font=("Segoe UI", 12),
            bg=BG_SURFACE, fg=TEXT_PRI,
            insertbackground=TEXT_PRI,
            relief="flat", highlightthickness=0, bd=0,
        )
        self._ent_pass.pack(fill="x", ipady=8, padx=1, pady=1)

        # ── Bot\u00f3n Entrar ──────────────────────────────────────────────────────
        HoverButton(card, normal=ACCENT, hover=ACCENT2,
                    text="  ENTRAR  \u2192", font=FONT_BTN,
                    fg="white", padx=20, pady=12,
                    command=self._submit).pack(fill="x")

        self._lbl_err = tk.Label(card, text="", font=("Segoe UI", 10),
                                  bg=BG_CARD, fg=ERR_COLOR)
        self._lbl_err.pack(pady=(12, 0))

        # Atajos de teclado
        self._ent_user.bind("<Return>", lambda _: self._ent_pass.focus())
        self._ent_pass.bind("<Return>", lambda _: self._submit())

    def _submit(self):
        user = self._ent_user.get().strip()
        pwd  = self._ent_pass.get()
        if not user or not pwd:
            self._lbl_err.config(text="\u26a0  Ingresa usuario y contraseña.")
            return
        self._lbl_err.config(text="")
        self.config(cursor="watch")
        self._on_login(user, pwd)


# ──────────────────────────────────────────────────────────────────────────────
#  Frame: Carga
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

        tk.Label(center, text="\u2699", font=("Segoe UI", 60),
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
        self._bar = ttk.Progressbar(center, mode="indeterminate", length=400,
                                     style="App.Horizontal.TProgressbar")
        self._bar.pack()

        # Log en tiempo real
        lf = tk.Frame(center, bg=BG_CARD, padx=2, pady=2)
        lf.pack(fill="x", pady=(20, 0))

        self._log = tk.Text(lf, height=11, width=62, bg=BG_SURFACE, fg=TEXT_MUT,
                             font=("Consolas", 9), relief="flat", state="disabled", wrap="word")
        sb = tk.Scrollbar(lf, command=self._log.yview, bg=BG_CARD, troughcolor=BG_SURFACE)
        self._log.config(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self._log.pack(side="left", fill="both", expand=True)

    # ── API pública ────────────────────────────────────────────────────────────
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
#  Frame: Home
# ──────────────────────────────────────────────────────────────────────────────
class HomeFrame(tk.Frame):
    def __init__(self, master, df_pivote: pd.DataFrame, on_reportes, on_actualizar):
        super().__init__(master, bg=BG_DARK)
        self._df_pivote     = df_pivote
        self._on_reportes   = on_reportes
        self._on_actualizar = on_actualizar
        self._build()

    def _build(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(3, weight=1)   # la tabla ocupa el espacio restante

        # ── Header ──────────────────────────────────────────────────────────
        hdr = tk.Frame(self, bg=BG_CARD, pady=16, padx=20)
        hdr.grid(row=0, column=0, sticky="ew")
        hdr.columnconfigure(1, weight=1)

        tk.Label(hdr, text="\U0001f4ca  Panel Principal",
                 font=("Segoe UI", 20, "bold"), bg=BG_CARD, fg=TEXT_PRI).grid(
                     row=0, column=0, sticky="w")
        tk.Label(hdr, text=f"Actualizado: {datetime.now():%d/%m/%Y  %H:%M}",
                 font=("Segoe UI", 10), bg=BG_CARD, fg=TEXT_MUT).grid(
                     row=1, column=0, sticky="w")

        HoverButton(hdr, normal="#1b4332", hover="#2d6a4f",
                    text="\U0001f504  Actualizar datos", font=("Segoe UI", 11, "bold"),
                    fg="#b7e4c7", padx=14, pady=8,
                    command=self._on_actualizar).grid(row=0, column=2, rowspan=2, padx=(10, 0))

        # ── Barra de módulos ─────────────────────────────────────────────────
        mod_bar = tk.Frame(self, bg=BG_SURFACE, pady=15, padx=25)
        mod_bar.grid(row=1, column=0, sticky="ew")

        self._module_card(mod_bar, "\U0001f4cb", "REPORTES",
                          "Visualiza y descarga\nreportes generados",
                          self._on_reportes, col=0)

        # placeholder futuro
        ph = tk.Frame(mod_bar, bg=BG_CARD, width=130, height=80, padx=10, pady=10)
        ph.grid(row=0, column=1, padx=(10, 0))
        ph.grid_propagate(False)
        tk.Label(ph, text="\u2795", font=("Segoe UI", 24), bg=BG_CARD, fg="#333355").pack(expand=True)
        tk.Label(ph, text="Próximamente", font=("Segoe UI", 9),
                 bg=BG_CARD, fg="#333355").pack()

        # ── Sub-header tabla Pivote ──────────────────────────────────────────
        tbl_hdr = tk.Frame(self, bg=BG_CARD, pady=8, padx=20)
        tbl_hdr.grid(row=2, column=0, sticky="ew")
        n_filas = len(self._df_pivote)
        n_cols  = len(self._df_pivote.columns)
        tk.Label(tbl_hdr,
                 text=f"\U0001f5c2  Base Pivote  \u2014  {n_filas} registros  \u00b7  {n_cols} columnas",
                 font=("Segoe UI", 11, "bold"), bg=BG_CARD, fg=TEXT_PRI).pack(side="left")

        HoverButton(tbl_hdr, normal=SUCCESS, hover=SUCCESS_H,
                    text="\u2b07  Descargar Excel", font=("Segoe UI", 11, "bold"),
                    fg="#000c0a", padx=14, pady=6,
                    command=self._download_pivote).pack(side="right", padx=(20, 0))

        # ── Tabla df_Pivote ──────────────────────────────────────────────────
        tf = tk.Frame(self, bg=BG_DARK, padx=12, pady=8)
        tf.grid(row=3, column=0, sticky="nsew")
        tf.columnconfigure(0, weight=1)
        tf.rowconfigure(0, weight=1)

        sty = ttk.Style()
        sty.theme_use("clam")
        sty.configure("Piv.Treeview",
                       background=ROW_ODD, foreground=TEXT_PRI,
                       rowheight=24, fieldbackground=ROW_ODD,
                       font=FONT_TABLE, borderwidth=0)
        sty.configure("Piv.Treeview.Heading",
                       background=HEADER_BG, foreground="white",
                       font=FONT_HEADER, relief="flat", padding=(5, 3))
        sty.map("Piv.Treeview", background=[("selected", ACCENT2)])
        sty.map("Piv.Treeview.Heading", background=[("active", ACCENT2)])

        cols = list(self._df_pivote.columns)
        tree = ttk.Treeview(tf, columns=cols, show="headings", style="Piv.Treeview")

        for col in cols:
            tree.heading(col, text=col)
            max_w = int(max(
                len(str(col)) * 9,
                (self._df_pivote[col].astype(str).str.len().max() * 8)
                if n_filas else 80,
            ))
            tree.column(col, width=min(max_w, 160), anchor="center", minwidth=50)

        tree.tag_configure("odd",  background=ROW_ODD,  foreground=TEXT_PRI)
        tree.tag_configure("even", background=ROW_EVEN, foreground=TEXT_PRI)

        for idx, row in self._df_pivote.iterrows():
            tag = "odd" if idx % 2 == 0 else "even"
            tree.insert("", "end", values=list(row), tags=(tag,))

        vsb = ttk.Scrollbar(tf, orient="vertical",   command=tree.yview)
        hsb = ttk.Scrollbar(tf, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

    def _download_pivote(self):
        ts   = datetime.now().strftime("%Y%m%d_%H%M")
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")],
            initialfile=f"Pivote_{ts}.xlsx",
            title="Guardar Base Pivote como...",
        )
        if not path:
            return
        try:
            if path.endswith(".csv"):
                self._df_pivote.to_csv(path, index=False, encoding="utf-8-sig")
            else:
                self._df_pivote.to_excel(path, index=False, engine="openpyxl")
            messagebox.showinfo("Guardado", f"Archivo guardado en:\n{path}")
        except Exception as exc:
            messagebox.showerror("Error al guardar", str(exc))

    def _module_card(self, parent, icon, title, desc, cmd, col):
        card = tk.Frame(parent, bg=BG_CARD, padx=16, pady=10, cursor="hand2",
                        width=130, height=80)
        card.grid(row=0, column=col)
        card.grid_propagate(False)

        li  = tk.Label(card, text=icon,  font=("Segoe UI", 26), bg=BG_CARD, fg=ACCENT)
        lt  = tk.Label(card, text=title, font=("Segoe UI", 11, "bold"), bg=BG_CARD, fg=TEXT_PRI)
        bar = tk.Frame(card, bg=ACCENT, height=3)
        li.pack(); lt.pack(pady=(2, 0)); bar.pack(fill="x", side="bottom")

        def _hover(bg):
            for w in (card, li, lt):
                w.config(bg=bg)

        for w in (card, li, lt, bar):
            w.bind("<Button-1>", lambda _: cmd())
            w.bind("<Enter>",    lambda _, b="#1e2a5a": _hover(b))
            w.bind("<Leave>",    lambda _, b=BG_CARD:   _hover(b))


# ──────────────────────────────────────────────────────────────────────────────
#  Frame: Reportes (muestra df_TEB)
# ──────────────────────────────────────────────────────────────────────────────
class ReportesFrame(tk.Frame):
    def __init__(self, master, df_teb: pd.DataFrame, on_back):
        super().__init__(master, bg=BG_DARK)
        self._df      = df_teb
        self._on_back = on_back
        self._build()

    def _build(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        # ── Header ─────────────────────────────────────────────────────────────
        hdr = tk.Frame(self, bg=BG_CARD, pady=15, padx=20)
        hdr.grid(row=0, column=0, sticky="ew")
        hdr.columnconfigure(1, weight=1)

        HoverButton(hdr, normal=BG_SURFACE, hover=ACCENT,
                    text="\u2190  Inicio", font=("Segoe UI", 11),
                    fg=TEXT_PRI, padx=12, pady=6,
                    command=self._on_back).grid(row=0, column=0, padx=(0, 20))

        tk.Label(hdr, text="\U0001f4cb  Reporte \u2014 Tabla Escalada B",
                 font=("Segoe UI", 18, "bold"), bg=BG_CARD, fg=TEXT_PRI).grid(
                     row=0, column=1, sticky="w")

        HoverButton(hdr, normal=SUCCESS, hover=SUCCESS_H,
                    text="\u2b07  Descargar Excel", font=("Segoe UI", 11, "bold"),
                    fg="#000c0a", padx=14, pady=6,
                    command=self._download).grid(row=0, column=2, padx=(20, 0))

        # ── Info bar ───────────────────────────────────────────────────────────
        info = tk.Frame(self, bg=BG_SURFACE, pady=6, padx=20)
        info.grid(row=1, column=0, sticky="ew")
        tk.Label(info, text=f"  {len(self._df)} filas  ·  {len(self._df.columns)} columnas",
                 font=("Segoe UI", 10), bg=BG_SURFACE, fg=TEXT_MUT).pack(side="left")

        # ── Tabla ─────────────────────────────────────────────────────────────
        tf = tk.Frame(self, bg=BG_DARK, padx=15, pady=15)
        tf.grid(row=2, column=0, sticky="nsew")
        tf.columnconfigure(0, weight=1)
        tf.rowconfigure(0, weight=1)

        sty = ttk.Style()
        sty.theme_use("clam")
        sty.configure("TEB.Treeview",
                       background=ROW_ODD, foreground=TEXT_PRI,
                       rowheight=26, fieldbackground=ROW_ODD,
                       font=FONT_TABLE, borderwidth=0)
        sty.configure("TEB.Treeview.Heading",
                       background=HEADER_BG, foreground="white",
                       font=FONT_HEADER, relief="flat", padding=(6, 4))
        sty.map("TEB.Treeview", background=[("selected", ACCENT2)])
        sty.map("TEB.Treeview.Heading", background=[("active", ACCENT2)])

        cols = list(self._df.columns)
        tree = ttk.Treeview(tf, columns=cols, show="headings", style="TEB.Treeview")

        for col in cols:
            tree.heading(col, text=col)
            max_w = int(max(
                len(str(col)) * 9,
                (self._df[col].astype(str).str.len().max() * 8) if len(self._df) else 80,
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

    # ── Descarga ───────────────────────────────────────────────────────────────
    def _download(self):
        ts = datetime.now().strftime("%Y%m%d_%H%M")
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")],
            initialfile=f"Reporte_TEB_{ts}.xlsx",
            title="Guardar reporte como...",
        )
        if not path:
            return
        try:
            if path.endswith(".csv"):
                self._df.to_csv(path, index=False, encoding="utf-8-sig")
            else:
                self._df.to_excel(path, index=False, engine="openpyxl")
            messagebox.showinfo("Guardado", f"Archivo guardado en:\n{path}")
        except Exception as exc:
            messagebox.showerror("Error al guardar", str(exc))

    def _download_pivote(self):
        ts = datetime.now().strftime("%Y%m%d_%H%M")
        path = filedialog.asksaveasfilename(    
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")],
            initialfile=f"BD_Pivote_{ts}.xlsx",
            title="Guardar df como...",
        )
        if not path:
            return
        try:
            if path.endswith(".csv"):
                self._df_pivote.to_csv(path, index=False, encoding="utf-8-sig")
            else:
                self._df_pivote.to_excel(path, index=False, engine="openpyxl")
            messagebox.showinfo("Guardado", f"Archivo guardado en:\n{path}")
        except Exception as exc:
            messagebox.showerror("Error al guardar", str(exc))


# ──────────────────────────────────────────────────────────────────────────────
#  Aplicación principal
# ──────────────────────────────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Reportes")
        self.geometry("1100x720")
        self.minsize(900, 600)
        self.configure(bg=BG_DARK)

        # Ícono opcional
        ico = os.path.join(BASE_DIR, "icon.ico")
        if os.path.exists(ico):
            try:
                self.iconbitmap(ico)
            except Exception:
                pass

        self._resultados: dict | None = None
        self._frame: tk.Frame | None  = None
        self._usuario:    str = ""
        self._contrasena: str = ""

        self._goto(LoginFrame(self, on_login=self._on_login))

    # ── Navegación ─────────────────────────────────────────────────────────────
    def _goto(self, frame: tk.Frame):
        if self._frame:
            self._frame.destroy()
        self._frame = frame
        frame.pack(fill="both", expand=True)

    # ── Handlers de flujo ─────────────────────────────────────────────────────
    def _on_login(self, usuario: str, contrasena: str):
        self._usuario    = usuario
        self._contrasena = contrasena
        self._iniciar_scraping()

    def _iniciar_scraping(self):
        loading = LoadingFrame(self)
        self._goto(loading)
        loading.start()

        usuario    = self._usuario
        contrasena = self._contrasena

        def _worker():
            try:
                resultado = scraper.run(
                    usuario, contrasena,
                    on_log=lambda msg: self.after(0, lambda m=msg: loading.append_log(m)),
                )
                self.after(0, lambda: self._on_scraping_ok(resultado, loading))
            except Exception as exc:
                tb = traceback.format_exc()
                self.after(0, lambda: self._on_scraping_error(str(exc), tb))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_scraping_ok(self, resultado: dict, loading: LoadingFrame):
        loading.stop()
        loading.set_status("Listo ✓")
        self._resultados = resultado
        self._show_home()

    def _on_scraping_error(self, msg: str, tb: str):
        print(tb)
        messagebox.showerror(
            "Error durante el procesamiento",
            f"{msg}\n\nRevisa la consola para más detalles.",
        )
        self._goto(LoginFrame(self, on_login=self._on_login))

    def _show_reportes(self):
        df = self._resultados.get("df_TEB", pd.DataFrame()) if self._resultados else pd.DataFrame()
        if df.empty:
            messagebox.showwarning("Sin datos", "df_TEB está vacío. Verifica el scraping.")
            return
        self._goto(ReportesFrame(self, df, on_back=self._show_home))

    def _show_home(self):
        df_piv = self._resultados.get("df_Pivote", pd.DataFrame()) if self._resultados else pd.DataFrame()
        self._goto(HomeFrame(
            self,
            df_pivote=df_piv,
            on_reportes=self._show_reportes,
            on_actualizar=self._iniciar_scraping,
        ))


# ──────────────────────────────────────────────────────────────────────────────
#  Entry point
# ──────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    App().mainloop()
