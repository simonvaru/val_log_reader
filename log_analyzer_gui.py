"""
GUI para log_analyzer.py — soporta múltiples logs en un solo reporte
Uso: python log_analyzer_gui.py
"""

import sys
import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Forzar UTF-8 en stdout (puede ser None en ejecutables sin consola)
if sys.stdout is not None and hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def run_analysis(log_files, xlsx_file, output_file, status_var, btn_run, root):
    """Ejecuta el análisis sobre múltiples logs y genera un único reporte."""
    import tempfile, shutil
    tmp_xlsx = None
    try:
        from log_analyzer import (load_events_from_xlsx, extract_log_lines,
                                   analyze_log, export_html, export_xlsx)

        # Copiar xlsx a temporal para evitar errno 13 si está abierto en Excel
        tmp_xlsx = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        tmp_xlsx.close()
        shutil.copy2(xlsx_file, tmp_xlsx.name)

        status_var.set("Cargando eventos...")
        root.update_idletasks()
        events = load_events_from_xlsx(tmp_xlsx.name)

        all_results   = []
        total_lines   = 0
        log_label     = ", ".join(os.path.basename(f) for f in log_files)

        for i, lf in enumerate(log_files, 1):
            status_var.set(f"Leyendo log {i}/{len(log_files)}: {os.path.basename(lf)}…")
            root.update_idletasks()
            lines = extract_log_lines(lf)
            total_lines += len(lines)

            status_var.set(f"Analizando {i}/{len(log_files)}: {os.path.basename(lf)}…")
            root.update_idletasks()
            results = analyze_log(lines, events)

            # Anotar de qué archivo proviene cada resultado
            for r in results:
                r["_source"] = os.path.basename(lf)
            all_results.extend(results)

        status_var.set("Exportando reporte...")
        root.update_idletasks()

        if output_file.lower().endswith(".xlsx"):
            export_xlsx(all_results, output_file)
        else:
            export_html(all_results, log_label, output_file)

        status_var.set(
            f"Listo. {len(all_results)} ocurrencias en {total_lines} líneas "
            f"({len(log_files)} log(s))."
        )
        messagebox.showinfo(
            "Análisis completado",
            f"Logs analizados: {len(log_files)}\n"
            f"Ocurrencias encontradas: {len(all_results)}\n"
            f"Líneas procesadas: {total_lines}\n\n"
            f"Reporte guardado en:\n{output_file}",
        )
    except Exception as e:
        status_var.set(f"Error: {e}")
        messagebox.showerror("Error", str(e))
    finally:
        if tmp_xlsx and os.path.exists(tmp_xlsx.name):
            try:
                os.remove(tmp_xlsx.name)
            except Exception:
                pass
        btn_run.config(state="normal")


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
    root.resizable(True, True)

    pad = {"padx": 10, "pady": 5}

    # ── Variables ─────────────────────────────────────────────────────────────
    output_var = tk.StringVar()
    status_var = tk.StringVar(value="Listo.")

    _base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    _default_xlsx = os.path.join(_base, "lista_eventos_vl550.xlsx")
    if not os.path.exists(_default_xlsx):
        _default_xlsx = os.path.join(os.getcwd(), "lista_eventos_vl550.xlsx")
    xlsx_var = tk.StringVar(value=_default_xlsx if os.path.exists(_default_xlsx) else "")

    # ── Frame principal ───────────────────────────────────────────────────────
    frm = ttk.Frame(root, padding=16)
    frm.grid(row=0, column=0, sticky="nsew")
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)
    frm.columnconfigure(1, weight=1)

    # Título
    ttk.Label(frm, text="Log Analyzer — VL550 / CGI",
              font=("Segoe UI", 12, "bold")).grid(
        row=0, column=0, columnspan=3, pady=(0, 12), sticky="w"
    )

    # ── Lista de logs ─────────────────────────────────────────────────────────
    ttk.Label(frm, text="Archivos de log:").grid(row=1, column=0, sticky="nw", **pad)

    list_frame = ttk.Frame(frm)
    list_frame.grid(row=1, column=1, sticky="nsew", **pad)
    list_frame.columnconfigure(0, weight=1)
    frm.rowconfigure(1, weight=1)

    log_listbox = tk.Listbox(list_frame, height=6, selectmode=tk.EXTENDED,
                              font=("Consolas", 9), activestyle="dotbox")
    log_listbox.grid(row=0, column=0, sticky="nsew")

    sb = ttk.Scrollbar(list_frame, orient="vertical", command=log_listbox.yview)
    sb.grid(row=0, column=1, sticky="ns")
    log_listbox.configure(yscrollcommand=sb.set)

    # Botones de la lista
    btn_frame = ttk.Frame(frm)
    btn_frame.grid(row=1, column=2, sticky="n", padx=(0, 10), pady=5)

    def add_logs():
        paths = filedialog.askopenfilenames(
            title="Agregar archivos de log",
            filetypes=[("Todos los archivos", "*.*"),
                       ("Texto", "*.txt"), ("Log", "*.log")],
        )
        existing = list(log_listbox.get(0, tk.END))
        for p in paths:
            if p not in existing:
                log_listbox.insert(tk.END, p)

    def remove_selected():
        for i in reversed(log_listbox.curselection()):
            log_listbox.delete(i)

    def clear_list():
        log_listbox.delete(0, tk.END)

    ttk.Button(btn_frame, text="➕ Agregar", width=12, command=add_logs).pack(pady=2)
    ttk.Button(btn_frame, text="➖ Quitar",  width=12, command=remove_selected).pack(pady=2)
    ttk.Button(btn_frame, text="🗑 Limpiar", width=12, command=clear_list).pack(pady=2)

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
    ttk.Label(frm, text="(extensión .html o .xlsx determina el formato)",
              foreground="#666", font=("Segoe UI", 8)).grid(
        row=4, column=1, sticky="w", padx=10
    )

    # ── Separador ─────────────────────────────────────────────────────────────
    ttk.Separator(frm, orient="horizontal").grid(
        row=5, column=0, columnspan=3, sticky="ew", pady=10
    )

    # ── Botón Analizar ────────────────────────────────────────────────────────
    btn_run = ttk.Button(frm, text="▶  Analizar")

    def on_run():
        log_files   = list(log_listbox.get(0, tk.END))
        xlsx_file   = xlsx_var.get().strip()
        output_file = output_var.get().strip()

        if not log_files:
            messagebox.showwarning("Falta dato", "Agregá al menos un archivo de log.")
            return
        missing = [f for f in log_files if not os.path.exists(f)]
        if missing:
            messagebox.showerror("Error", "No se encontraron:\n" + "\n".join(missing))
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
            args=(log_files, xlsx_file, output_file, status_var, btn_run, root),
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
        row=8, column=0, columnspan=3, sticky="w", padx=10, pady=(4, 2)
    )

    root.mainloop()


if __name__ == "__main__":
    main()
