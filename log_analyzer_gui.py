"""
GUI para log_analyzer.py
Uso: python log_analyzer_gui.py
"""

import sys
import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Forzar UTF-8 en stdout
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def run_analysis(log_file, xlsx_file, output_file, status_var, btn_run, root):
    """Ejecuta el análisis en un hilo separado para no bloquear la GUI."""
    try:
        # Importar funciones del analizador
        from log_analyzer import load_events_from_xlsx, extract_log_lines, analyze_log, export_html, export_xlsx

        status_var.set("Cargando eventos...")
        root.update_idletasks()
        events = load_events_from_xlsx(xlsx_file)

        status_var.set(f"Leyendo log ({os.path.basename(log_file)})...")
        root.update_idletasks()
        log_lines = extract_log_lines(log_file)

        status_var.set("Analizando...")
        root.update_idletasks()
        results = analyze_log(log_lines, events)

        status_var.set("Exportando reporte...")
        root.update_idletasks()

        if output_file.lower().endswith(".xlsx"):
            export_xlsx(results, output_file)
        else:
            export_html(results, log_file, output_file)

        status_var.set(
            f"Listo. {len(results)} ocurrencias en {len(log_lines)} líneas."
        )
        messagebox.showinfo(
            "Análisis completado",
            f"Ocurrencias encontradas: {len(results)}\n"
            f"Líneas procesadas: {len(log_lines)}\n\n"
            f"Reporte guardado en:\n{output_file}",
        )
    except Exception as e:
        status_var.set(f"Error: {e}")
        messagebox.showerror("Error", str(e))
    finally:
        btn_run.config(state="normal")


def browse_log(var):
    path = filedialog.askopenfilename(
        title="Seleccionar archivo de log",
        filetypes=[("Todos los archivos", "*.*"), ("Texto", "*.txt"), ("Log", "*.log")],
    )
    if path:
        var.set(path)


def browse_xlsx(var):
    path = filedialog.askopenfilename(
        title="Seleccionar base de eventos (.xlsx)",
        filetypes=[("Excel", "*.xlsx")],
    )
    if path:
        var.set(path)


def browse_output(var):
    path = filedialog.asksaveasfilename(
        title="Guardar reporte como...",
        defaultextension=".html",
        filetypes=[("HTML", "*.html"), ("Excel", "*.xlsx")],
    )
    if path:
        var.set(path)


def main():
    root = tk.Tk()
    root.title("Log Analyzer VL550")
    root.resizable(False, False)

    pad = {"padx": 10, "pady": 6}

    # ── Variables ────────────────────────────────────────────────────────────
    log_var    = tk.StringVar()
    xlsx_var   = tk.StringVar(value="lista_eventos_vl550.xlsx")
    output_var = tk.StringVar()
    status_var = tk.StringVar(value="Listo.")

    # ── Frame principal ───────────────────────────────────────────────────────
    frm = ttk.Frame(root, padding=16)
    frm.grid(row=0, column=0, sticky="nsew")

    # Título
    ttk.Label(frm, text="Log Analyzer — VL550 / CGI", font=("Segoe UI", 12, "bold")).grid(
        row=0, column=0, columnspan=3, pady=(0, 14), sticky="w"
    )

    # ── Archivo de log ────────────────────────────────────────────────────────
    ttk.Label(frm, text="Archivo de log:").grid(row=1, column=0, sticky="w", **pad)
    ttk.Entry(frm, textvariable=log_var, width=52).grid(row=1, column=1, sticky="ew", **pad)
    ttk.Button(frm, text="Examinar…", command=lambda: browse_log(log_var)).grid(
        row=1, column=2, **pad
    )

    # ── Base de eventos ───────────────────────────────────────────────────────
    ttk.Label(frm, text="Base de eventos (.xlsx):").grid(row=2, column=0, sticky="w", **pad)
    ttk.Entry(frm, textvariable=xlsx_var, width=52).grid(row=2, column=1, sticky="ew", **pad)
    ttk.Button(frm, text="Examinar…", command=lambda: browse_xlsx(xlsx_var)).grid(
        row=2, column=2, **pad
    )

    # ── Archivo de salida ─────────────────────────────────────────────────────
    ttk.Label(frm, text="Guardar reporte en:").grid(row=3, column=0, sticky="w", **pad)
    ttk.Entry(frm, textvariable=output_var, width=52).grid(row=3, column=1, sticky="ew", **pad)
    ttk.Button(frm, text="Guardar como…", command=lambda: browse_output(output_var)).grid(
        row=3, column=2, **pad
    )

    ttk.Label(frm, text="(extensión .html o .xlsx determina el formato)", foreground="#666",
              font=("Segoe UI", 8)).grid(row=4, column=1, sticky="w", padx=10)

    # ── Separador ─────────────────────────────────────────────────────────────
    ttk.Separator(frm, orient="horizontal").grid(
        row=5, column=0, columnspan=3, sticky="ew", pady=10
    )

    # ── Botón Analizar ────────────────────────────────────────────────────────
    btn_run = ttk.Button(frm, text="▶  Analizar", style="Accent.TButton")

    def on_run():
        log_file    = log_var.get().strip()
        xlsx_file   = xlsx_var.get().strip()
        output_file = output_var.get().strip()

        if not log_file:
            messagebox.showwarning("Falta dato", "Seleccioná el archivo de log.")
            return
        if not os.path.exists(log_file):
            messagebox.showerror("Error", f"No se encontró el log:\n{log_file}")
            return
        if not xlsx_file:
            messagebox.showwarning("Falta dato", "Seleccioná la base de eventos (.xlsx).")
            return
        if not os.path.exists(xlsx_file):
            messagebox.showerror("Error", f"No se encontró el Excel:\n{xlsx_file}")
            return
        if not output_file:
            messagebox.showwarning("Falta dato", "Indicá dónde guardar el reporte.")
            return

        btn_run.config(state="disabled")
        status_var.set("Procesando…")
        threading.Thread(
            target=run_analysis,
            args=(log_file, xlsx_file, output_file, status_var, btn_run, root),
            daemon=True,
        ).start()

    btn_run.config(command=on_run)
    btn_run.grid(row=6, column=0, columnspan=3, pady=(4, 10))

    # ── Barra de estado ───────────────────────────────────────────────────────
    ttk.Separator(frm, orient="horizontal").grid(
        row=7, column=0, columnspan=3, sticky="ew"
    )
    ttk.Label(frm, textvariable=status_var, foreground="#444",
              font=("Segoe UI", 8)).grid(
        row=8, column=0, columnspan=3, sticky="w", padx=10, pady=(4, 0)
    )

    root.mainloop()


if __name__ == "__main__":
    main()
