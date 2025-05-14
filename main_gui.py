# main_gui.py ─ GA Office Helper (CustomTkinter edition)
# ------------------------------------------------------
import os, sys, threading
from datetime import date, datetime
import customtkinter as ctk
import json
from pathlib import Path

from tkinter import messagebox
from PIL import Image

from parse_rows import parse_ptt_rows_from_text, parse_bd_row
from fill_ptt   import get_template_path, open_output_folder
from fill_bd    import generate_bd_sheet
from airline_map import AIRLINE_MAP
from docxtpl     import DocxTemplate
from docx2pdf    import convert
from mini_updater import check_and_update, __version__ as APP_VERSION


# ───────────────────────── helpers ─────────────────────────
def get_exe_folder() -> str:
    return os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.getcwd()


def get_output_folder() -> str:
    out_dir = os.path.join(get_exe_folder(), "GeneratedDocuments")
    os.makedirs(out_dir, exist_ok=True)
    return out_dir


def show_toast(parent, msg: str, duration: int = 3000) -> None:
    """Tiny toast bottom-right."""
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
SETTINGS_PATH = Path(get_exe_folder()) / "settings.json"

def load_settings() -> dict:
    if SETTINGS_PATH.exists():
        try:
            return json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            pass
    return {}

def save_settings(data: dict) -> None:
    try:
        SETTINGS_PATH.write_text(json.dumps(data, indent=2), encoding="utf-8")
    except Exception as e:
        print(f"[WARN] Could not save settings: {e}")


# ───────────────────────── main window ─────────────────────────
class GAOfficeHelper(ctk.CTk):
    def __init__(self):
        super().__init__()

        # ---- style ----
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")
        self.title("GA Office Helper")
        self.geometry("1000x700")
        self.minsize(920, 630)
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        # ---- tabs ----
        self.tabs = ctk.CTkTabview(self, width=960, height=540)
        self.tabs.pack(fill="both", expand=True, padx=10, pady=(10, 0))

        self._build_ptt_tab()
        self._build_bd_tab()
        self._build_status_bar()

    # ───── PTT TAB ─────
    def _build_ptt_tab(self):
        tab = self.tabs.add("PTT Generator")

        # ---- instructions + firm code ---------------------------------
        top = ctk.CTkFrame(tab, fg_color="transparent", height=44)
        top.pack(fill="x", padx=10, pady=(10, 0))

        # label (use place so it's truly centred)
        instr = ctk.CTkLabel(
            top,
            text="Paste multiple PTT rows here.\nThen click “Generate PTT Docs”.",
            justify="center", anchor="center"
        )

        NUDGE_PX = 0  # ← bump ± pixels if it’s a hair off
        instr.place(relx=0.5, rely=0.0, anchor="n", x=NUDGE_PX)

        # firm-code group, stays at far right
        self.firm_code_var = ctk.StringVar(value="WBS6")
        right = ctk.CTkFrame(top, fg_color="transparent")
        right.pack(side="right", padx=(0, 10))
        ctk.CTkLabel(right, text="Firm Code:").pack(side="left", padx=(0, 4))
        ctk.CTkOptionMenu(right, variable=self.firm_code_var, values=["WBS6", "W353"]) \
            .pack(side="left")

        # ── paste box ──
        self.ptt_text = ctk.CTkTextbox(tab, width=940, height=260)
        self.ptt_text.pack(padx=10, pady=(10, 10), fill="both", expand=True)

        # ── progress ──
        self.ptt_bar = ctk.CTkProgressBar(tab, width=940); self.ptt_bar.set(0)
        self.ptt_bar.pack(padx=10, pady=4)
        self.ptt_label = ctk.CTkLabel(tab, text=""); self.ptt_label.pack()

        # ── operator + generate row ──
        action = ctk.CTkFrame(tab, fg_color="transparent")
        action.pack(fill="x", padx=10, pady=(10, 4))

        # left spacer (expand=True keeps center centred)
        ctk.CTkLabel(action, text="").pack(side="left", expand=True)

        # center generate button
        ctk.CTkButton(action, text="Generate PTT Docs",
                      command=lambda: threading.Thread(
                          target=self._generate_ptt_docs, daemon=True).start()
                     ).pack(side="left", padx=188)

        # right operator entry
        op_frame = ctk.CTkFrame(action, fg_color="transparent")
        op_frame.pack(side="right", padx=10)
        ctk.CTkLabel(op_frame, text="Operator:").pack(side="left", padx=(0, 4))
        self.opname_var = ctk.StringVar()
        # preload saved operator name (if any)
        prev = load_settings().get("last_operator", "")
        self.opname_var.set(prev)

        ctk.CTkEntry(op_frame, width=140, textvariable=self.opname_var)\
            .pack(side="left")

        # ── open folder button ──
        ctk.CTkButton(tab, text="Open Output Folder",
                      command=open_output_folder).pack(pady=(0, 10))

    def _generate_ptt_docs(self):
        raw = self.ptt_text.get("1.0", "end").strip()
        if not raw:
            messagebox.showwarning("No Data", "Please paste some rows first."); return
        records = parse_ptt_rows_from_text(raw)
        if not records:
            messagebox.showwarning("No Data", "No valid PTT rows found."); return

        op_name = self.opname_var.get().strip()
        if not op_name:
            messagebox.showwarning("Operator Required", "Please enter your name before generating.")
            return

        folder, total = get_output_folder(), len(records)
        firm_code, op_name = self.firm_code_var.get(), self.opname_var.get().strip()
        today = date.today().strftime("%m/%d/%Y")

        self.status_var.set("Generating PTT documents…")
        self.ptt_bar.configure(mode="determinate"); self.ptt_bar.set(0)

        for idx, rec in enumerate(records, start=1):
            mawb, flt, pcs, wt = rec["mawb"], rec["flt"], rec["pieces"], rec["weight"]
            airline = AIRLINE_MAP.get(mawb[:3], "Unknown Airline")

            doc = DocxTemplate(get_template_path())
            doc.render({
                "MAWB": mawb, "PIECES": pcs, "WEIGHT": wt, "FLT": flt,
                "AirlineName": airline, "TODAYS_DATE": today,
                "FirmCode": firm_code, "OPName": op_name
            })
            docx_path = os.path.join(folder, f"PTT_{mawb}.docx")
            doc.save(docx_path)
            try:
                convert(docx_path); os.remove(docx_path)
            except Exception as e:
                print(f"[WARN] PDF conversion failed: {e}")

            self.ptt_bar.set(idx/total)
            self.ptt_label.configure(text=f"Processing {idx}/{total}")
            self.update_idletasks()

        self.ptt_text.delete("1.0", "end")
        self.ptt_bar.set(0); self.ptt_label.configure(text="")
        self.status_var.set(f"PTT done — {total} PDFs saved.")

        save_settings({"last_operator": op_name})

        show_toast(self, f"Generated {total} PTT PDF(s)")

    # ───── BD TAB ─────
    def _build_bd_tab(self):
        tab = self.tabs.add("B/D Sheet Generator")
        ctk.CTkLabel(tab, text="Paste exactly one row for B/D sheet.\nThen click “Generate BD Doc”.")\
            .pack(padx=10, pady=10)
        self.bd_text = ctk.CTkTextbox(tab, width=940, height=110)
        self.bd_text.pack(padx=10, pady=(0, 10))
        ctk.CTkButton(tab, text="Generate BD Doc",
                      command=lambda: threading.Thread(
                          target=self._generate_bd_doc, daemon=True).start()
                     ).pack(pady=(10, 4))
        ctk.CTkButton(tab, text="Open Output Folder",
                      command=open_output_folder).pack(pady=(0, 10))

    def _generate_bd_doc(self):
        raw = self.bd_text.get("1.0", "end").strip()
        if not raw:
            messagebox.showwarning("No Data", "Please paste one row."); return
        record = parse_bd_row(raw)
        if not record.get("mawb"):
            messagebox.showwarning("No Data", "No valid BD row."); return

        out_path = generate_bd_sheet(record)
        if not out_path or not os.path.exists(out_path):
            show_toast(self, "BD doc not generated."); return

        self.bd_text.delete("1.0", "end")
        msg = [f"Clearance: {record['hold']}", f"MAWB: {record['mawb']}"]
        if record.get("carriers") == [("800-", ""), ("807-", ""), ("808-", "")]:
            msg.append("Please fill in details for USPS-CO")
        self.status_var.set(f"BD done — saved {os.path.basename(out_path)}")
        show_toast(self, "BD sheet generated.")
        messagebox.showinfo("B/D Done", "\n".join(msg))

    # ───── status bar ─────
    def _build_status_bar(self):
        bar = ctk.CTkFrame(self, fg_color="#1f1f1f", height=28)
        bar.pack(fill="x", side="bottom")
        try:
            logo_small = ctk.CTkImage(Image.open(os.path.join(get_exe_folder(), "GA_Logo.png")),
                                      size=(64, 32))
            ctk.CTkLabel(bar, image=logo_small, text="").pack(side="left", padx=(6, 4))
        except Exception: pass
        ctk.CTkLabel(bar, text=f"GA Office Helper  v{APP_VERSION}")\
            .pack(side="left", padx=(0, 8))
        self.status_var = ctk.StringVar(value="Ready")
        ctk.CTkLabel(bar, textvariable=self.status_var, anchor="w")\
            .pack(side="left", fill="x", expand=True)
        ctk.CTkButton(bar, text="Check for Update", width=140,
                      command=lambda: threading.Thread(
                          target=self._run_update, daemon=True).start()
                     ).pack(side="right", padx=6, pady=2)

    def _run_update(self):
        self.status_var.set("Checking for updates…")
        res = check_and_update()
        if res == "latest":
            self.status_var.set("Already on latest version."); show_toast(self, "Up-to-date.")
        elif res.startswith("error:"):
            self.status_var.set("Update failed.")
            messagebox.showerror("Update Error", res.split(":", 1)[1])

    # ───── close ─────
    def on_closing(self):
        self.destroy(); sys.exit(0)


# ───────────────────────── run ─────────────────────────
if __name__ == "__main__":
    app = GAOfficeHelper()
    app.mainloop()
