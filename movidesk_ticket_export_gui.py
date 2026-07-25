"""
Movidesk Ticket Export — Com interface gráfica (Tkinter)

Essa não é ferramenta utilizada em produção por mim. Eu diminui muito o código para poder não trazer dados sensíveis e apenas testar e estruturar a retirada de tickets.
dados sensíveis e apenas testar e estruturar a retirada de tickets.

Esse scrit utiliza o movidesk_ticket_export.py, porém ele adiciona apenas uma interface gráfica para facilitar a utilização da ferramenta. 

Caso for utilizar o script com a interface gráfica, você precisa ter os códigos na mesma pasta.

movidesk_ticket_export.py
movidesk_ticket_export_gui.py

obs: 

Para retirar os tickets, vai precisar de um token de API. Para obter o token, voce deve solicitar ao suporte da Movidesk.


Para mais informações, entre em contato comigo no linkedin: www.linkedin.com/in/mateus-gomes-279349218

Requirements:
    pip install requests pandas openpyxl tkcalendar

exemplos de uso:
    python movidesk_ticket_export.py --token SEU_TOKEN --start-date 2025-01-01 --end-date 2025-01-31
    python movidesk_ticket_export.py --token SEU_TOKEN --start-date 2025-01-01 --end-date 2025-01-31 --format csv
    set MOVIDESK_API_TOKEN=YOUR_TOKEN
    python movidesk_ticket_export.py --start-date 2025-01-01 --end-date 2025-01-31
"""
from __future__ import annotations

import os
import threading
import tkinter as tk
from datetime import datetime, timedelta
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Optional

from movidesk_ticket_export import (
    MovideskClient,
    fetch_tickets,
    save_dataframe,
    tickets_to_dataframe,
)

try:
    from tkcalendar import DateEntry
except ImportError:  # optional dependency
    DateEntry = None  # type: ignore[misc, assignment]


# ---------------------------------------------------------------------------
# Theme (neutral slate — good for portfolio screenshots)
# ---------------------------------------------------------------------------
COLORS = {
    "bg": "#F4F6F8",
    "panel": "#FFFFFF",
    "border": "#D8DEE6",
    "text": "#1F2933",
    "muted": "#62748A",
    "accent": "#0F766E",
    "accent_hover": "#0D9488",
    "accent_text": "#FFFFFF",
    "danger": "#B91C1C",
    "log_bg": "#0B1220",
    "log_fg": "#E2E8F0",
}


class ExportApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Movidesk Ticket Export")
        self.minsize(720, 560)
        self.configure(bg=COLORS["bg"])
        self._worker: Optional[threading.Thread] = None
        self._build_styles()
        self._build_ui()
        self._center_window(780, 620)

    def _center_window(self, width: int, height: int) -> None:
        self.update_idletasks()
        x = (self.winfo_screenwidth() - width) // 2
        y = (self.winfo_screenheight() - height) // 3
        self.geometry(f"{width}x{height}+{x}+{y}")

    def _build_styles(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure("Root.TFrame", background=COLORS["bg"])
        style.configure("Panel.TFrame", background=COLORS["panel"])
        style.configure(
            "Title.TLabel",
            background=COLORS["bg"],
            foreground=COLORS["text"],
            font=("Segoe UI", 20, "bold"),
        )
        style.configure(
            "Subtitle.TLabel",
            background=COLORS["bg"],
            foreground=COLORS["muted"],
            font=("Segoe UI", 10),
        )
        style.configure(
            "Field.TLabel",
            background=COLORS["panel"],
            foreground=COLORS["text"],
            font=("Segoe UI", 10),
        )
        style.configure(
            "Hint.TLabel",
            background=COLORS["panel"],
            foreground=COLORS["muted"],
            font=("Segoe UI", 9),
        )
        style.configure(
            "Status.TLabel",
            background=COLORS["bg"],
            foreground=COLORS["muted"],
            font=("Segoe UI", 9),
        )
        style.configure(
            "Panel.TLabelframe",
            background=COLORS["panel"],
            foreground=COLORS["text"],
            bordercolor=COLORS["border"],
            relief="solid",
        )
        style.configure(
            "Panel.TLabelframe.Label",
            background=COLORS["panel"],
            foreground=COLORS["text"],
            font=("Segoe UI", 10, "bold"),
        )
        style.configure(
            "Accent.TButton",
            background=COLORS["accent"],
            foreground=COLORS["accent_text"],
            font=("Segoe UI", 10, "bold"),
            padding=(16, 8),
            borderwidth=0,
        )
        style.map(
            "Accent.TButton",
            background=[
                ("active", COLORS["accent_hover"]),
                ("disabled", "#94A3B8"),
            ],
            foreground=[("disabled", "#F8FAFC")],
        )
        style.configure("TEntry", fieldbackground="#FFFFFF", padding=6)
        style.configure("TCombobox", padding=4)
        style.configure(
            "TCheckbutton",
            background=COLORS["panel"],
            foreground=COLORS["text"],
            font=("Segoe UI", 10),
        )

    def _build_ui(self) -> None:
        root = ttk.Frame(self, style="Root.TFrame", padding=24)
        root.pack(fill="both", expand=True)

        ttk.Label(root, text="Movidesk Ticket Export", style="Title.TLabel").pack(
            anchor="w"
        )
        ttk.Label(
            root,
            text="Consulta genérica à API pública — exporta tickets sem regras de negócio.",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(4, 18))

        form = ttk.LabelFrame(
            root, text="  Parâmetros  ", style="Panel.TLabelframe", padding=18
        )
        form.pack(fill="x")

        # Token
        ttk.Label(form, text="Token da API", style="Field.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        self.token_var = tk.StringVar(value=os.getenv("MOVIDESK_API_TOKEN", ""))
        self.token_entry = ttk.Entry(form, textvariable=self.token_var, show="•", width=56)
        self.token_entry.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(4, 2))
        self.show_token_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            form,
            text="Mostrar token",
            variable=self.show_token_var,
            command=self._toggle_token,
        ).grid(row=2, column=0, sticky="w", pady=(0, 12))

        # Dates
        today = datetime.now().date()
        default_end = today
        default_start = today - timedelta(days=7)

        ttk.Label(form, text="Data inicial", style="Field.TLabel").grid(
            row=3, column=0, sticky="w"
        )
        ttk.Label(form, text="Data final", style="Field.TLabel").grid(
            row=3, column=1, sticky="w", padx=(16, 0)
        )
        ttk.Label(form, text="Formato", style="Field.TLabel").grid(
            row=3, column=2, sticky="w", padx=(16, 0)
        )

        self.start_date = self._make_date_widget(form, default_start)
        self.start_date.grid(row=4, column=0, sticky="ew", pady=(4, 12))
        self.end_date = self._make_date_widget(form, default_end)
        self.end_date.grid(row=4, column=1, sticky="ew", padx=(16, 0), pady=(4, 12))

        self.format_var = tk.StringVar(value="xlsx")
        format_box = ttk.Combobox(
            form,
            textvariable=self.format_var,
            values=("xlsx", "csv", "json"),
            state="readonly",
            width=10,
        )
        format_box.grid(row=4, column=2, sticky="w", padx=(16, 0), pady=(4, 12))

        # Output path
        ttk.Label(form, text="Arquivo de saída (opcional)", style="Field.TLabel").grid(
            row=5, column=0, sticky="w"
        )
        out_row = ttk.Frame(form, style="Panel.TFrame")
        out_row.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(4, 12))
        self.output_var = tk.StringVar(value="")
        ttk.Entry(out_row, textvariable=self.output_var).pack(
            side="left", fill="x", expand=True
        )
        ttk.Button(out_row, text="Escolher…", command=self._browse_output).pack(
            side="left", padx=(8, 0)
        )

        # Options
        opts = ttk.Frame(form, style="Panel.TFrame")
        opts.grid(row=7, column=0, columnspan=3, sticky="w")
        self.auto_past_var = tk.BooleanVar(value=True)
        self.force_past_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            opts,
            text="Usar /tickets/past automaticamente (> 90 dias)",
            variable=self.auto_past_var,
        ).pack(anchor="w")
        ttk.Checkbutton(
            opts,
            text="Forçar endpoint histórico (/tickets/past)",
            variable=self.force_past_var,
        ).pack(anchor="w", pady=(4, 0))

        form.columnconfigure(0, weight=1)
        form.columnconfigure(1, weight=1)
        form.columnconfigure(2, weight=0)

        # Actions
        actions = ttk.Frame(root, style="Root.TFrame")
        actions.pack(fill="x", pady=(16, 8))
        self.export_btn = ttk.Button(
            actions,
            text="Exportar tickets",
            style="Accent.TButton",
            command=self._start_export,
        )
        self.export_btn.pack(side="left")
        self.status_var = tk.StringVar(value="Pronto.")
        ttk.Label(actions, textvariable=self.status_var, style="Status.TLabel").pack(
            side="left", padx=(16, 0)
        )

        # Log
        log_frame = ttk.LabelFrame(
            root, text="  Log  ", style="Panel.TLabelframe", padding=10
        )
        log_frame.pack(fill="both", expand=True, pady=(8, 0))
        self.log = tk.Text(
            log_frame,
            height=12,
            wrap="word",
            bg=COLORS["log_bg"],
            fg=COLORS["log_fg"],
            insertbackground=COLORS["log_fg"],
            font=("Consolas", 10),
            relief="flat",
            padx=10,
            pady=10,
        )
        self.log.pack(fill="both", expand=True)
        self.log.configure(state="disabled")
        self._log("Informe o token e o período, depois clique em Exportar.")
        self._log("Os dados são gravados exatamente como retornados pela API (achatados em colunas).")

    def _make_date_widget(self, parent: tk.Misc, default) -> tk.Widget:
        if DateEntry is not None:
            kwargs = {
                "width": 14,
                "background": COLORS["accent"],
                "foreground": "white",
                "borderwidth": 1,
                "date_pattern": "yyyy-mm-dd",
            }
            try:
                widget = DateEntry(parent, locale="pt_BR", **kwargs)
            except Exception:
                widget = DateEntry(parent, **kwargs)
            widget.set_date(default)
            return widget

        var = tk.StringVar(value=default.strftime("%Y-%m-%d"))
        entry = ttk.Entry(parent, textvariable=var, width=16)
        entry._date_var = var  # type: ignore[attr-defined]
        return entry

    def _get_date(self, widget: tk.Widget) -> str:
        if DateEntry is not None and isinstance(widget, DateEntry):
            return widget.get_date().strftime("%Y-%m-%d")
        var = getattr(widget, "_date_var", None)
        if var is not None:
            return str(var.get()).strip()
        return str(widget.get()).strip()  # type: ignore[attr-defined]

    def _toggle_token(self) -> None:
        self.token_entry.configure(show="" if self.show_token_var.get() else "•")

    def _browse_output(self) -> None:
        fmt = self.format_var.get()
        ext = {"xlsx": ".xlsx", "csv": ".csv", "json": ".json"}.get(fmt, ".xlsx")
        path = filedialog.asksaveasfilename(
            title="Salvar exportação",
            defaultextension=ext,
            filetypes=[
                ("Excel", "*.xlsx"),
                ("CSV", "*.csv"),
                ("JSON", "*.json"),
                ("Todos", "*.*"),
            ],
        )
        if path:
            self.output_var.set(path)

    def _log(self, message: str) -> None:
        stamp = datetime.now().strftime("%H:%M:%S")
        self.log.configure(state="normal")
        self.log.insert("end", f"[{stamp}] {message}\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _set_busy(self, busy: bool) -> None:
        self.export_btn.configure(state="disabled" if busy else "normal")
        self.status_var.set("Exportando…" if busy else "Pronto.")

    def _start_export(self) -> None:
        if self._worker and self._worker.is_alive():
            return

        token = self.token_var.get().strip()
        if not token:
            messagebox.showwarning("Token", "Informe o token da API Movidesk.")
            return

        try:
            start = self._get_date(self.start_date)
            end = self._get_date(self.end_date)
            datetime.strptime(start, "%Y-%m-%d")
            datetime.strptime(end, "%Y-%m-%d")
        except Exception:
            messagebox.showerror("Datas", "Use datas no formato YYYY-MM-DD.")
            return

        if start > end:
            messagebox.showerror("Datas", "A data inicial não pode ser maior que a final.")
            return

        fmt = self.format_var.get()
        output = self.output_var.get().strip()
        force_past: Optional[bool] = None
        if self.force_past_var.get():
            force_past = True
        elif not self.auto_past_var.get():
            force_past = False

        self._set_busy(True)
        self._log(f"Iniciando exportação: {start} → {end} ({fmt})")

        self._worker = threading.Thread(
            target=self._run_export,
            args=(token, start, end, fmt, output, force_past),
            daemon=True,
        )
        self._worker.start()

    def _run_export(
        self,
        token: str,
        start: str,
        end: str,
        fmt: str,
        output: str,
        force_past: Optional[bool],
    ) -> None:
        try:
            client = MovideskClient(token)

            def progress(msg: str) -> None:
                self.after(0, lambda m=msg: self._log(m))

            # Lightweight progress by wrapping prints from fetch would need refactor;
            # call fetch and report counts after.
            progress("Consultando API Movidesk…")
            tickets = fetch_tickets(
                client,
                start_date=start,
                end_date=end,
                force_past=force_past,
                on_progress=progress,
            )

            df = tickets_to_dataframe(tickets)
            if not output:
                stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                output = str(Path.cwd() / f"movidesk_tickets_{stamp}")

            saved = save_dataframe(df, output, fmt)
            self.after(0, lambda: self._on_success(saved, len(df)))
        except Exception as exc:
            self.after(0, lambda: self._on_error(str(exc)))

    def _on_success(self, path: str, rows: int) -> None:
        self._set_busy(False)
        self._log(f"Concluído: {rows} linha(s) → {path}")
        self.status_var.set(f"Exportado: {rows} tickets")
        messagebox.showinfo("Exportação concluída", f"{rows} ticket(s) salvos em:\n{path}")

    def _on_error(self, message: str) -> None:
        self._set_busy(False)
        self._log(f"ERRO: {message}")
        self.status_var.set("Falha na exportação.")
        messagebox.showerror("Erro", message)


def main() -> None:
    app = ExportApp()
    app.mainloop()


if __name__ == "__main__":
    main()
