"""
SICAG · Sistema Integrado de Control de Avances Geológicos
Área de Geología (Chinalco)
App de escritorio: Excel de control de avances KPDI/MREDP -> Dashboard HTML, Word y PPT.

Uso:    python app_sicag.py
Prueba: python app_sicag.py --cli Control_de_Avances.xlsx [carpeta_salida] [dd/mm/aaaa]
"""
import base64
import io
import json
import os
import queue
import subprocess
import sys
import threading
import traceback
from datetime import datetime, timedelta, timezone
from pathlib import Path

# En el .exe sin consola, stdout/stderr no existen: se reemplazan para que print() no falle.
if sys.stdout is None:
    sys.stdout = io.StringIO()
if sys.stderr is None:
    sys.stderr = io.StringIO()

LOGO_B64 = "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"
APP = "SICAG"
VERSION = "1.0"
CONFIG = Path.home() / ".sicag_geologia.json"

AYUDA = """Cómo usar

1. Seleccionar Excel de control: el archivo Control_de_Avances_KPDI_MREDP (hoja «Data»).
2. Carpeta de salida: donde se guardan los reportes (por defecto, «Reportes_DDH» junto al Excel).
3. Fecha del reporte: aparece en el contenido y en el nombre de cada archivo (por defecto, hoy en Lima).
4. Generar: Dashboard (HTML), Reporte gerencial (Word), Presentación (PPT) o todo a la vez.

Cada corrida valida los conteos contra el resumen del propio Excel (columnas AL:AM de «Data»).

Hoja opcional «Leyes»: columnas DDHID, Desde, Hasta y una columna por elemento (Zn, Pb, Ag, Cu, Au).
Si existe, el dashboard y el Word incluyen la sección de resultados de leyes.
"""


def cargar_config():
    try:
        return json.loads(CONFIG.read_text(encoding="utf-8"))
    except Exception:
        return {}


def guardar_config(cfg):
    try:
        CONFIG.write_text(json.dumps(cfg, ensure_ascii=False), encoding="utf-8")
    except Exception:
        pass


def abrir_carpeta(ruta):
    try:
        if sys.platform.startswith("win"):
            os.startfile(str(ruta))
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(ruta)])
        else:
            subprocess.Popen(["xdg-open", str(ruta)])
    except Exception:
        pass


# ----------------------------------------------------------------------------- modo consola (pruebas)
def modo_cli(args):
    import motor_reportes as mr
    excel = args[0]
    carpeta = args[1] if len(args) > 1 else str(Path(excel).resolve().parent / "Reportes_DDH")
    fecha = args[2] if len(args) > 2 else None
    mr.generar_todo(excel, fecha=fecha, carpeta=carpeta)


# ----------------------------------------------------------------------------- interfaz
def lanzar_gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    try:                                    # nitidez en pantallas de alta resolución (Windows)
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    BG, SURF, INK, INK2 = "#F3F5F8", "#FFFFFF", "#15212D", "#53626F"
    BLUE, BLUE_D, LINE, SOFT = "#1F5E99", "#174B7C", "#CBD5E1", "#EAF0F6"
    FONT = "Segoe UI" if sys.platform.startswith("win") else "DejaVu Sans"

    root = tk.Tk()
    root.title("SICAG · Sistema Integrado de Control de Avances Geológicos")
    root.configure(bg=BG)
    root.geometry("640x760")
    root.minsize(600, 720)

    # --- logo
    logo_raw = base64.b64decode(LOGO_B64)
    try:
        from PIL import Image, ImageTk
        img = Image.open(io.BytesIO(logo_raw)).convert("RGBA")
        logo = ImageTk.PhotoImage(img.resize((84, int(84 * img.height / img.width)), Image.LANCZOS))
        icono = ImageTk.PhotoImage(img.resize((64, 64), Image.LANCZOS))
    except Exception:
        logo = tk.PhotoImage(data=LOGO_B64).subsample(3)
        icono = logo
    try:
        root.iconphoto(True, icono)
    except Exception:
        pass

    # --- estilos
    st = ttk.Style()
    st.theme_use("clam")
    st.configure("TButton", font=(FONT, 10), padding=(14, 9), background=SURF, foreground=INK,
                 bordercolor=LINE, lightcolor=SURF, darkcolor=SURF, focuscolor=SURF, relief="flat")
    st.map("TButton", background=[("active", SOFT), ("disabled", BG)], foreground=[("disabled", "#9AA8B7")],
           bordercolor=[("active", BLUE), ("disabled", "#E3E8EE")])
    st.configure("Primary.TButton", font=(FONT, 10, "bold"), background=BLUE, foreground="#FFFFFF",
                 bordercolor=BLUE, lightcolor=BLUE, darkcolor=BLUE, padding=(14, 11))
    st.map("Primary.TButton", background=[("active", BLUE_D), ("disabled", "#A9C0D8")],
           foreground=[("disabled", "#FFFFFF")], bordercolor=[("active", BLUE_D), ("disabled", "#A9C0D8")])
    st.configure("Help.TButton", font=(FONT, 10, "bold"), padding=(8, 3))
    st.configure("Horizontal.TProgressbar", troughcolor=SOFT, background=BLUE, bordercolor=SOFT,
                 lightcolor=BLUE, darkcolor=BLUE, thickness=6)
    st.configure("TCheckbutton", background=BG, foreground=INK2, font=(FONT, 9), indicatorbackground=SURF,
                 indicatorforeground=BLUE, upperbordercolor=LINE, lowerbordercolor=LINE, indicatorcolor=SURF)
    st.map("TCheckbutton", background=[("active", BG)], indicatorcolor=[("selected", SURF)],
           indicatorbackground=[("active", SOFT)])

    # --- cabecera
    head = tk.Frame(root, bg=SURF)
    head.pack(fill="x")
    tk.Label(head, image=logo, bg=SURF).pack(side="left", padx=(22, 14), pady=14)
    tx = tk.Frame(head, bg=SURF)
    tx.pack(side="left", pady=14)
    tk.Label(tx, text=APP, font=(FONT, 20, "bold"), fg=INK, bg=SURF).pack(anchor="w")
    tk.Label(tx, text="Sistema Integrado de Control de Avances Geológicos", font=(FONT, 9), fg=INK2, bg=SURF).pack(anchor="w")
    tk.Label(tx, text="Seguimiento MREDP 2026 · KPDI", font=(FONT, 9), fg=INK2, bg=SURF).pack(anchor="w")
    badge = tk.Label(head, text="ÁREA DE GEOLOGÍA", font=(FONT, 8, "bold"), fg="#FFFFFF", bg=BLUE, padx=10, pady=4)
    badge.pack(side="right", padx=22)
    tk.Frame(root, bg=BLUE, height=3).pack(fill="x")

    body = tk.Frame(root, bg=BG)
    body.pack(fill="both", expand=True, padx=26, pady=(10, 0))

    def seccion(titulo):
        f = tk.Frame(body, bg=BG)
        f.pack(fill="x", pady=(14, 6))
        tk.Label(f, text=titulo, font=(FONT, 11, "bold", "underline"), fg=INK, bg=BG).pack(anchor="w")
        return f

    estado = {"excel": None, "salida": None, "ultimo": None, "ocupado": False}
    cfg = cargar_config()

    def recorta(p, n=58):
        p = str(p)
        return p if len(p) <= n else "…" + p[-(n - 1):]

    # --- entradas
    seccion("ENTRADA")
    fila = tk.Frame(body, bg=BG); fila.pack(fill="x", pady=3)
    b_excel = ttk.Button(fila, text="Seleccionar Excel de control…", width=28)
    b_excel.pack(side="left")
    l_excel = tk.Label(fila, text="Ningún archivo seleccionado", font=(FONT, 9), fg=INK2, bg=BG, anchor="w")
    l_excel.pack(side="left", padx=12, fill="x", expand=True)

    fila = tk.Frame(body, bg=BG); fila.pack(fill="x", pady=3)
    b_salida = ttk.Button(fila, text="Carpeta de salida…", width=28)
    b_salida.pack(side="left")
    l_salida = tk.Label(fila, text="Se creará «Reportes_DDH» junto al Excel", font=(FONT, 9), fg=INK2, bg=BG, anchor="w")
    l_salida.pack(side="left", padx=12, fill="x", expand=True)

    fila = tk.Frame(body, bg=BG); fila.pack(fill="x", pady=3)
    tk.Label(fila, text="Fecha del reporte", font=(FONT, 10), fg=INK, bg=BG, width=22, anchor="w").pack(side="left")
    v_fecha = tk.StringVar(value=datetime.now(timezone(timedelta(hours=-5))).strftime("%d/%m/%Y"))
    e_fecha = tk.Entry(fila, textvariable=v_fecha, width=14, font=(FONT, 10), relief="flat", bg=SURF, fg=INK,
                       highlightthickness=1, highlightbackground=LINE, highlightcolor=BLUE, justify="center")
    e_fecha.pack(side="left", ipady=5)
    tk.Label(fila, text="dd/mm/aaaa (hoy, Lima)", font=(FONT, 9), fg=INK2, bg=BG).pack(side="left", padx=12)

    # --- generar
    seccion("GENERAR")
    g = tk.Frame(body, bg=BG); g.pack(fill="x")
    g.columnconfigure((0, 1, 2), weight=1, uniform="g")
    b_html = ttk.Button(g, text="Dashboard (HTML)")
    b_doc = ttk.Button(g, text="Reporte (Word)")
    b_ppt = ttk.Button(g, text="Presentación (PPT)")
    b_html.grid(row=0, column=0, sticky="ew", padx=(0, 5))
    b_doc.grid(row=0, column=1, sticky="ew", padx=5)
    b_ppt.grid(row=0, column=2, sticky="ew", padx=(5, 0))
    b_todo = ttk.Button(body, text="Generar todo  (HTML + Word + PPT)", style="Primary.TButton")
    b_todo.pack(fill="x", pady=(10, 0))
    abrir_auto = tk.BooleanVar(value=cfg.get("abrir", True))
    ttk.Checkbutton(body, text="Abrir la carpeta de salida al terminar", variable=abrir_auto).pack(anchor="w", pady=(8, 0))

    # --- registro
    seccion("REGISTRO")
    prog = ttk.Progressbar(body, mode="determinate", value=0, style="Horizontal.TProgressbar")
    prog.pack(fill="x", pady=(0, 6))
    cont = tk.Frame(body, bg=SURF, highlightthickness=1, highlightbackground=LINE)
    cont.pack(fill="both", expand=True)
    log = tk.Text(cont, height=9, font=("Consolas" if sys.platform.startswith("win") else "DejaVu Sans Mono", 9),
                  bg=SURF, fg=INK, relief="flat", wrap="char", state="disabled", padx=10, pady=8)
    sb = ttk.Scrollbar(cont, orient="vertical", command=log.yview)
    log.configure(yscrollcommand=sb.set)
    sb.pack(side="right", fill="y")
    log.pack(side="left", fill="both", expand=True)
    log.tag_configure("ok", foreground="#2A7B58")
    log.tag_configure("err", foreground="#B3372F")
    log.tag_configure("dim", foreground=INK2)

    def escribe(txt, tag=None):
        log.configure(state="normal")
        log.insert("end", txt, tag)
        log.see("end")
        log.configure(state="disabled")

    # --- pie
    pie = tk.Frame(root, bg=BG)
    pie.pack(fill="x", padx=26, pady=12)
    b_ayuda = ttk.Button(pie, text="?", style="Help.TButton", width=3)
    b_ayuda.pack(side="left")
    b_abrir = ttk.Button(pie, text="Abrir carpeta de salida")
    b_abrir.pack(side="left", padx=10)
    tk.Label(pie, text=f"SICAG v{VERSION} · Área de Geología · Chinalco", font=(FONT, 8), fg=INK2, bg=BG).pack(side="right")

    botones_gen = (b_html, b_doc, b_ppt, b_todo)

    def refresca():
        listo = estado["excel"] is not None and not estado["ocupado"]
        for b in botones_gen:
            b.state(["!disabled"] if listo else ["disabled"])
        b_excel.state(["disabled"] if estado["ocupado"] else ["!disabled"])
        b_salida.state(["disabled"] if estado["ocupado"] else ["!disabled"])
        b_abrir.state(["!disabled"] if (estado["salida"] or estado["ultimo"]) else ["disabled"])

    def carpeta_salida():
        if estado["salida"]:
            return Path(estado["salida"])
        return Path(estado["excel"]).resolve().parent / "Reportes_DDH"

    def elegir_excel():
        ini = cfg.get("dir_excel") or str(Path.home())
        p = filedialog.askopenfilename(title="Excel de control de avances", initialdir=ini,
                                       filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("Todos", "*.*")])
        if not p:
            return
        estado["excel"] = p
        l_excel.configure(text=recorta(p), fg=INK)
        if not estado["salida"]:
            l_salida.configure(text=recorta(carpeta_salida()), fg=INK)
        cfg["dir_excel"] = str(Path(p).parent)
        guardar_config(cfg)
        escribe(f"Excel: {Path(p).name}\n", "dim")
        refresca()

    def elegir_salida():
        p = filedialog.askdirectory(title="Carpeta de salida", initialdir=cfg.get("dir_salida") or str(Path.home()))
        if not p:
            return
        estado["salida"] = p
        l_salida.configure(text=recorta(p), fg=INK)
        cfg["dir_salida"] = p
        guardar_config(cfg)
        refresca()

    # --- ejecución en segundo plano
    q = queue.Queue()

    class Salida:
        def write(self, s):
            if s:
                q.put(("log", s))
        def flush(self):
            pass

    def trabajo(formatos, fecha, excel, carpeta):
        antes = sys.stdout
        sys.stdout = Salida()
        try:
            q.put(("log", "Cargando motor de reportes…\n"))
            import motor_reportes as mr
            rutas = mr.generar_todo(excel, fecha=fecha, carpeta=carpeta, formatos=formatos)
            q.put(("fin", rutas))
        except Exception:
            q.put(("error", traceback.format_exc()))
        finally:
            sys.stdout = antes

    def sondeo():
        try:
            while True:
                tipo, dato = q.get_nowait()
                if tipo in ("fin", "error"):
                    prog.stop()
                    prog.configure(mode="determinate", value=0)
                if tipo == "log":
                    escribe(dato)
                elif tipo == "fin":
                    estado["ocupado"] = False
                    carpeta = Path(next(iter(dato.values()))).parent
                    estado["ultimo"] = str(carpeta)
                    escribe("\n✔ Reportes generados en:\n", "ok")
                    escribe(f"  {carpeta}\n", "ok")
                    refresca()
                    if abrir_auto.get():
                        abrir_carpeta(carpeta)
                elif tipo == "error":
                    estado["ocupado"] = False
                    escribe("\n✖ No se pudo generar el reporte.\n", "err")
                    escribe(dato + "\n", "err")
                    refresca()
                    messagebox.showerror("SICAG", "No se pudo generar el reporte.\nRevisa el registro para ver el detalle.")
        except queue.Empty:
            pass
        root.after(120, sondeo)

    def generar(formatos):
        if estado["excel"] is None or estado["ocupado"]:
            return
        txt = v_fecha.get().strip()
        fecha = None
        if txt:
            try:
                fecha = datetime.strptime(txt, "%d/%m/%Y").strftime("%d/%m/%Y")
            except ValueError:
                messagebox.showwarning("Fecha", "Escribe la fecha como dd/mm/aaaa, por ejemplo 22/11/2026.")
                return
        cfg["abrir"] = abrir_auto.get()
        guardar_config(cfg)
        estado["ocupado"] = True
        refresca()
        prog.configure(mode="indeterminate")
        prog.start(12)
        escribe("\n— Nueva corrida —\n", "dim")
        threading.Thread(target=trabajo, args=(formatos, fecha, estado["excel"], str(carpeta_salida())), daemon=True).start()

    b_excel.configure(command=elegir_excel)
    b_salida.configure(command=elegir_salida)
    b_html.configure(command=lambda: generar(("html",)))
    b_doc.configure(command=lambda: generar(("docx",)))
    b_ppt.configure(command=lambda: generar(("pptx",)))
    b_todo.configure(command=lambda: generar(("html", "docx", "pptx")))
    b_abrir.configure(command=lambda: abrir_carpeta(estado["ultimo"] or carpeta_salida()))
    b_ayuda.configure(command=lambda: messagebox.showinfo("Ayuda · SICAG", AYUDA))

    escribe("Selecciona el Excel de control para comenzar.\n", "dim")
    refresca()
    root.after(120, sondeo)
    root.mainloop()


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--cli":
        modo_cli(sys.argv[2:])
    else:
        lanzar_gui()
