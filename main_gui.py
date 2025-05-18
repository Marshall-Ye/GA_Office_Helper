# main_gui.py ─ GA Office Helper (CustomTkinter edition)
# ======================================================
from __future__ import annotations

import json
import os
import sys
import threading
from datetime import date
from pathlib import Path
from tkinter import messagebox

import customtkinter as ctk
from PIL import Image

from parse_rows import parse_ptt_rows_from_text, parse_bd_row
from fill_ptt import (
    FIRM_CHOICES,
    DEFAULT_FIRM,
    generate_ptt_for_records,
    open_output_folder,
)
from fill_bd import generate_bd_sheet
from mini_updater import check_and_update, __version__ as APP_VERSION


# ───────────────────────── helpers ─────────────────────────
def get_exe_folder() -> str:
    return os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.getcwd()


def get_resource(*parts) -> Path:
    if hasattr(sys, "_MEIPASS"):               # one-file exe
        base = Path(sys._MEIPASS)
    elif getattr(sys, "frozen", False):        # one-folder build
        base = Path(sys.executable).parent
    else:                                      # source run
        base = Path(__file__).parent
    return base / "_internal" / Path(*parts)


def show_toast(parent, msg: str, duration: int = 3000) -> None:
    try:
        toast = ctk.CTkToplevel(parent)
        toast.overrideredirect(True)
        toast.attributes("-topmost", True)
        toast.configure(fg_color="#2b2b2b")
        ctk.CTkLabel(toast, text=msg).pack(padx=14, pady=8)
        parent.update_idletasks()
        x = parent.winfo_x() + parent.winfo_width() - toast.winfo_reqwidth() - 24
        y = parent.winfo_y() + parent.winfo_height() - toast.winfo_reqheight() - 48
        toast.geometry(f"+{x}+{y}")
        toast.after(duration, toast.destroy)
    except Exception:
        pass


# ─────────────────── settings persistence ────────────────────
INTERNAL_DIR = Path(get_exe_folder()) / "_internal"
INTERNAL_DIR.mkdir(exist_ok=True)
SETTINGS_PATH = INTERNAL_DIR / "settings.json"


def load_settings() -> dict:
    if SETTINGS_PATH.exists():
        try:
            return json.loads(SETTINGS_PATH.read_text("utf-8"))
        except json.JSONDecodeError:
            pass
    return {}


def save_settings(data: dict) -> None:
    try:
        SETTINGS_PATH.write_text(json.dumps(data, indent=2), "utf-8")
    except Exception as e:
        print(f"[WARN] Could not save settings: {e}")


# ───────────────────────── main window ─────────────────────────
class GAOfficeHelper(ctk.CTk):
    # ──────────────── init / layout ────────────────
    def __init__(self) -> None:
        super().__init__()

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")
        self.title("GA Office Helper")
        self.geometry("1000x700")
        self.minsize(920, 630)
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.tabs = ctk.CTkTabview(self, width=960, height=540)
        self.tabs.pack(fill="both", expand=True, padx=10, pady=(10, 0))

        self._build_ptt_tab()
        self._build_bd_tab()
        self._build_status_bar()

    # ───── PTT TAB ─────
    def _build_ptt_tab(self) -> None:
        tab = self.tabs.add("PTT Generator")

        # instructions + firm selector
        top = ctk.CTkFrame(tab, fg_color="transparent", height=44)
        top.pack(fill="x", padx=10, pady=(10, 0))

        instr = ctk.CTkLabel(
            top,
            text="Paste multiple PTT rows here.\nThen click “Generate PTT Docs”.",
            justify="center",
            anchor="center",
        )
        instr.place(relx=0.5, rely=0.0, anchor="n")

        # right-side firm selector
        self.firm_var = ctk.StringVar(value=DEFAULT_FIRM)
        right = ctk.CTkFrame(top, fg_color="transparent")
        right.pack(side="right", padx=(0, 10))
        ctk.CTkLabel(right, text="Firm:").pack(side="left", padx=(0, 4))
        ctk.CTkOptionMenu(right, variable=self.firm_var, values=list(FIRM_CHOICES)).pack(
            side="left"
        )

        # paste box
        self.ptt_text = ctk.CTkTextbox(tab, width=940, height=260)
        self.ptt_text.pack(padx=10, pady=(10, 10), fill="both", expand=True)

        # progress widgets
        self.ptt_bar = ctk.CTkProgressBar(tab, width=940)
        self.ptt_bar.pack(padx=10, pady=4)
        self.ptt_bar.set(0)
        self.ptt_label = ctk.CTkLabel(tab, text="")
        self.ptt_label.pack()

        # action row
        action = ctk.CTkFrame(tab, fg_color="transparent")
        action.pack(fill="x", padx=10, pady=(10, 4))
        ctk.CTkLabel(action, text="").pack(side="left", expand=True)

        ctk.CTkButton(
            action,
            text="Generate PTT Docs",
            command=self._start_ptt_generation,
        ).pack(side="left", padx=188)

        # operator entry
        op_frame = ctk.CTkFrame(action, fg_color="transparent")
        op_frame.pack(side="right", padx=10)
        ctk.CTkLabel(op_frame, text="Operator:").pack(side="left", padx=(0, 4))
        self.opname_var = ctk.StringVar(value=load_settings().get("last_operator", ""))
        ctk.CTkEntry(op_frame, width=140, textvariable=self.opname_var).pack(side="left")

        ctk.CTkButton(tab, text="Open Output Folder", command=open_output_folder).pack(
            pady=(0, 10)
        )

    # ---- PTT workflow --------------------------------------------------
    def _start_ptt_generation(self) -> None:
        raw = self.ptt_text.get("1.0", "end").strip()
        if not raw:
            messagebox.showwarning("No Data", "Please paste some rows first.")
            return
        records = parse_ptt_rows_from_text(raw)
        if not records:
            messagebox.showwarning("No Data", "No valid PTT rows found.")
            return

        op_name = self.opname_var.get().strip()
        if not op_name:
            messagebox.showwarning(
                "Operator Required", "Please enter your name before generating."
            )
            return

        firm_key = self.firm_var.get()

        # UI feedback – spinner in the same tab
        self.status_var.set("Generating PTT documents…")
        self.ptt_bar.configure(mode="indeterminate")
        self.ptt_bar.start()
        self.ptt_label.configure(text="Working…")
        self.update_idletasks()

        threading.Thread(
            target=self._worker_ptt_generation,
            args=(records, firm_key, op_name),
            daemon=True,
        ).start()

    def _worker_ptt_generation(self, records, firm_key: str, op_name: str) -> None:
        pdfs = generate_ptt_for_records(records, firm_key, op_name)

        def _ui_done():
            self.ptt_bar.stop()
            self.ptt_bar.configure(mode="determinate")
            self.ptt_bar.set(0)
            self.ptt_label.configure(text="")
            self.ptt_text.delete("1.0", "end")

            self.status_var.set(f"PTT done — {len(pdfs)} PDF(s) saved.")
            save_settings({"last_operator": op_name})
            show_toast(self, f"Generated {len(pdfs)} PTT PDF(s)")
            messagebox.showinfo("PTT Finished", f"Generated {len(pdfs)} PDF file(s).")

        self.after(0, _ui_done)

    # ───── BD TAB ─────
    def _build_bd_tab(self) -> None:
        tab = self.tabs.add("B/D Sheet Generator")

        ctk.CTkLabel(
            tab,
            text="Paste exactly one row for B/D sheet.\nThen click “Generate BD Doc”.",
        ).pack(padx=10, pady=10)

        self.bd_text = ctk.CTkTextbox(tab, width=940, height=110)
        self.bd_text.pack(padx=10, pady=(0, 10))

        ctk.CTkButton(
            tab,
            text="Generate BD Doc",
            command=lambda: threading.Thread(
                target=self._generate_bd_doc, daemon=True
            ).start(),
        ).pack(pady=(10, 4))

        ctk.CTkButton(tab, text="Open Output Folder", command=open_output_folder).pack(
            pady=(0, 10)
        )

    def _generate_bd_doc(self) -> None:
        raw = self.bd_text.get("1.0", "end").strip()
        if not raw:
            messagebox.showwarning("No Data", "Please paste one row.")
            return

        record = parse_bd_row(raw)
        if not record.get("mawb"):
            messagebox.showwarning("No Data", "No valid BD row.")
            return

        out_path = generate_bd_sheet(record)
        if not out_path or not os.path.exists(out_path):
            show_toast(self, "BD doc not generated.")
            return

        self.bd_text.delete("1.0", "end")
        msg = [f"Clearance: {record['hold']}", f"MAWB: {record['mawb']}"]
        if record.get("carriers") == [("800-", ""), ("807-", ""), ("808-", "")]:
            msg.append("Please fill in details for USPS-CO")

        self.status_var.set(f"BD done — saved {os.path.basename(out_path)}")
        show_toast(self, "BD sheet generated.")
        messagebox.showinfo("B/D Done", "\n".join(msg))

    # ───── status bar / updater ─────
    def _build_status_bar(self) -> None:
        bar = ctk.CTkFrame(self, fg_color="#1f1f1f", height=46)
        bar.pack(fill="x", side="bottom")

        left = ctk.CTkFrame(bar, fg_color="transparent")
        left.pack(side="left", padx=6, pady=2)

        try:
            banner = ctk.CTkImage(Image.open(get_resource("GA_Logo.png")), size=(100, 30))
            ctk.CTkLabel(left, image=banner, text="").pack(anchor="w")
        except Exception as e:
            print("[WARN] banner load failed:", e)

        ctk.CTkLabel(left, text=f"GA Office Helper  v{APP_VERSION}", font=("", 12)).pack(
            anchor="w", pady=(2, 0)
        )

        self.status_var = ctk.StringVar(value="Ready")
        ctk.CTkLabel(bar, textvariable=self.status_var, anchor="w").pack(
            side="left", fill="x", expand=True
        )

        ctk.CTkButton(
            bar,
            text="Check for Update",
            width=140,
            command=lambda: threading.Thread(
                target=self._run_update, daemon=True
            ).start(),
        ).pack(side="right", padx=6, pady=6)

    def _run_update(self) -> None:
        self.status_var.set("Checking for updates…")
        res = check_and_update()
        if res == "latest":
            self.status_var.set("Already on latest version.")
            show_toast(self, "Up-to-date.")
        elif res.startswith("error:"):
            self.status_var.set("Update failed.")
            messagebox.showerror("Update Error", res.split(":", 1)[1])

    # ───── close ─────
    def on_closing(self) -> None:
        self.destroy()
        sys.exit(0)


# ───────────────────────── run ─────────────────────────
if __name__ == "__main__":
    app = GAOfficeHelper()
    app.mainloop()
