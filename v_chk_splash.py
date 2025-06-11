import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk
import time
import os

from src.v_chk import VaultHealthCheck
from src.v_chk_wb_setup import WbDataDef
from src.v_chk_wb_tabs import NewWb
from src.v_chk_xl import ExcelExporter

SPLASH_BG = "#B01513"
LOGO_PATH = os.path.abspath(os.path.join("..", "img", "swenlogoicon.png"))


class SplashScreen(tk.Tk):
    def __init__(self, logo_path):
        super().__init__()
        self.overrideredirect(True)
        self.configure(bg=SPLASH_BG)
        self.logo_img = self.load_logo(logo_path)
        self.logo_label = tk.Label(self, image=self.logo_img, bg=SPLASH_BG)
        self.logo_label.pack(expand=True)

        self.status_var = tk.StringVar()
        self.status_label = tk.Label(self, textvariable=self.status_var, anchor="sw",
                                     bg=SPLASH_BG, fg="white", font=("Arial", 12))
        self.status_label.pack(side="bottom", anchor="sw", padx=10, pady=10, fill="x")

        self.progress = ttk.Progressbar(self, orient="horizontal", mode="determinate", length=300)
        self.progress.pack(side="bottom", pady=10)
        self.progress["value"] = 0
        self.progress["maximum"] = 100

        self.center_window(400, 300)

    def load_logo(self, path):
        img = Image.open(path)
        img = img.resize((128, 128), Image.LANCZOS)
        return ImageTk.PhotoImage(img)

    def center_window(self, w, h):
        self.update_idletasks()
        ws = self.winfo_screenwidth()
        hs = self.winfo_screenheight()
        x = (ws // 2) - (w // 2)
        y = (hs // 2) - (h // 2)
        self.geometry(f"{w}x{h}+{x}+{y}")

    def update_status(self, text, progress_value=None):
        self.status_var.set(text)
        if progress_value is not None:
            self.progress["value"] = progress_value
        self.update_idletasks()


def main():
    splash = SplashScreen(LOGO_PATH)
    splash.update_status("Starting Vault Health Check...", 0)
    splash.after(500, lambda: run_main(splash))
    splash.mainloop()


def run_main(splash):
    DBUG_LVL = 0
    splash.update_status("Initializing Vault Health Check...", 10)
    # vc_def = VaultHealthCheck(DBUG_LVL)
    # tabs = NewWb(DBUG_LVL)
    # exporter = ExcelExporter(DBUG_LVL)
    # exporter.export(DBUG_LVL)
    try:
        splash.update_status("Processing Vault...", 20)
        DBUG_LVL = -1
        vc_obj = VaultHealthCheck(DBUG_LVL)

        splash.update_status("Building workbook tab structure...", 70)
        wb_obj = NewWb(DBUG_LVL)

        splash.update_status("Exporting workbook...", 90)
        exporter = ExcelExporter(DBUG_LVL)
        exporter.export(DBUG_LVL)

        splash.update_status("Done. Launching application...", 100)
        time.sleep(1)
    except Exception as e:
        splash.update_status(f"Error: {e}")
        time.sleep(2)
    finally:
        splash.destroy()


if __name__ == "__main__":
    main()