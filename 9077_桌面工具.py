"""Local desktop interface for the Excel/CSV cleaner. No network access."""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path
import queue
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import webbrowser

ROOT = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location("cleaner9077", ROOT / "9077_表格清洗.py")
engine = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = engine
spec.loader.exec_module(engine)


class CleanerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Table Cleaner | Python Desktop Automation")
        self.geometry("1060x740")
        self.minsize(940, 650)
        self.configure(bg="#f3f6fa")
        self.events = queue.Queue()
        self.report = None
        self.input_path = tk.StringVar(value=str(ROOT / "9077_示例订单.csv"))
        self.config_path = tk.StringVar(value=str(ROOT / "9077_清洗规则.json"))
        self.output_path = tk.StringVar(value=str(ROOT / "9077_桌面运行结果"))
        self.status_text = tk.StringVar(value="Choose a file and rule set, then run locally.")
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TFrame", background="#f3f6fa")
        style.configure("TLabel", background="#f3f6fa", foreground="#172b46", font=("Segoe UI", 10))
        style.configure("Title.TLabel", font=("Segoe UI", 25, "bold"))
        style.configure("Sub.TLabel", foreground="#607087")
        style.configure("TButton", font=("Segoe UI", 10), padding=(12, 7))
        style.configure("Run.TButton", background="#157a68", foreground="white", font=("Segoe UI", 11, "bold"))
        style.configure("Treeview", font=("Segoe UI", 10), rowheight=30)
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))
        frame = ttk.Frame(self, padding=28)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Table Cleaner", style="Title.TLabel").pack(anchor="w")
        ttk.Label(frame, text="Excel and CSV automation with a row-by-row audit", style="Sub.TLabel").pack(anchor="w", pady=(4, 20))
        form = ttk.Frame(frame)
        form.pack(fill="x")
        form.columnconfigure(1, weight=1)
        for row, (label, value, action) in enumerate([
            ("Input .csv / .xlsx", self.input_path, self.pick_input),
            ("Rules .json", self.config_path, self.pick_config),
            ("New output folder", self.output_path, self.pick_output),
        ]):
            ttk.Label(form, text=label, width=19).grid(row=row, column=0, sticky="w", pady=5)
            ttk.Entry(form, textvariable=value).grid(row=row, column=1, sticky="ew", padx=(0, 10), pady=5)
            ttk.Button(form, text="Browse", command=action).grid(row=row, column=2)
        buttons = ttk.Frame(frame)
        buttons.pack(fill="x", pady=16)
        self.run_button = ttk.Button(buttons, text="Run cleaning", style="Run.TButton", command=self.start_run)
        self.run_button.pack(side="left")
        self.report_button = ttk.Button(buttons, text="Open audit report", command=self.open_report, state="disabled")
        self.report_button.pack(side="left", padx=10)
        ttk.Label(buttons, text="Source files stay unchanged.", style="Sub.TLabel").pack(side="right")
        self.metrics = ttk.Frame(frame)
        self.metrics.pack(fill="x", pady=(2, 14))
        self.metric_values = {}
        for label, key in [("PROCESSED", "processed"), ("KEPT", "valid"), ("INVALID", "invalid"), ("DUPLICATES", "duplicate"), ("FILTERED", "filtered")]:
            box = tk.Frame(self.metrics, bg="white", padx=18, pady=10, highlightbackground="#dce4ed", highlightthickness=1)
            box.pack(side="left", fill="x", expand=True, padx=(0, 7))
            value = tk.StringVar(value="--")
            self.metric_values[key] = value
            tk.Label(box, textvariable=value, bg="white", fg="#157a68" if key == "valid" else "#172b46", font=("Segoe UI", 23, "bold")).pack(anchor="w")
            tk.Label(box, text=label, bg="white", fg="#607087", font=("Segoe UI", 9)).pack(anchor="w")
        ttk.Label(frame, text="Processing log", font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(4, 7))
        columns = ("row", "status", "id", "details")
        self.table = ttk.Treeview(frame, columns=columns, show="headings", height=9)
        for key, label, width in [("row", "Source row", 85), ("status", "Status", 100), ("id", "Record ID", 110), ("details", "Details", 570)]:
            self.table.heading(key, text=label)
            self.table.column(key, width=width, stretch=key == "details")
        self.table.pack(fill="both", expand=True)
        self.table.tag_configure("invalid", foreground="#ad382e")
        self.table.tag_configure("valid", foreground="#157a68")
        ttk.Label(frame, textvariable=self.status_text, style="Sub.TLabel", wraplength=960).pack(anchor="w", pady=(12, 3))
        ttk.Label(frame, text="Independent portfolio project. All included sample records are synthetic.", style="Sub.TLabel").pack(anchor="w")
        self.after(100, self.poll_events)

    def pick_input(self):
        selected = filedialog.askopenfilename(filetypes=[("Tables", "*.csv *.xlsx")])
        if selected:
            self.input_path.set(selected)

    def pick_config(self):
        selected = filedialog.askopenfilename(filetypes=[("JSON rules", "*.json")])
        if selected:
            self.config_path.set(selected)

    def pick_output(self):
        selected = filedialog.askdirectory(title="Select a parent folder for a NEW output folder")
        if selected:
            self.output_path.set(str(Path(selected) / "9077_清洗结果"))

    def start_run(self):
        if not all(v.get().strip() for v in (self.input_path, self.config_path, self.output_path)):
            messagebox.showerror("Missing path", "Select an input file, rules and output folder.")
            return
        self.run_button.configure(state="disabled")
        self.report_button.configure(state="disabled")
        self.status_text.set("Processing locally...")
        args = (self.input_path.get(), self.config_path.get(), self.output_path.get())
        def work():
            try:
                self.events.put(("done", engine.run(*args), args[2]))
            except Exception as exc:
                self.events.put(("error", str(exc), None))
        threading.Thread(target=work, daemon=True).start()

    def poll_events(self):
        try:
            kind, result, output = self.events.get_nowait()
            self.run_button.configure(state="normal")
            if kind == "error":
                self.status_text.set("Stopped: " + result)
                messagebox.showerror("Cleaning stopped", result)
            else:
                output = Path(output)
                audit = json.loads((output / "9077_审计报告.json").read_text(encoding="utf-8"))
                counts = audit["counts"]
                for key, value in self.metric_values.items():
                    value.set(str(counts[key]))
                self.table.delete(*self.table.get_children())
                for item in audit["rows"]:
                    cleaned = item.get("cleaned") or {}
                    original = item.get("original") or {}
                    record_id = cleaned.get("order_id", original.get("order_id", ""))
                    self.table.insert("", "end", values=(item["source_row"], item["status"], record_id, "; ".join(item["reasons"]) or "Validated and kept"), tags=(item["status"],))
                self.report = output / "9077_审计报告.html"
                self.report_button.configure(state="normal")
                self.status_text.set(f"Complete. CSV, Excel and audit reports saved to {output.name}")
        except queue.Empty:
            pass
        self.after(100, self.poll_events)

    def open_report(self):
        if self.report and self.report.is_file():
            webbrowser.open(self.report.resolve().as_uri())


if __name__ == "__main__":
    CleanerApp().mainloop()
