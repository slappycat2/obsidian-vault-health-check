import os
import yaml
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk
from pathlib import Path
from datetime import datetime
import subprocess
import platform
import json
from typing import Tuple, List, Dict, Any

from src.v_chk_class_lib import ObsidianApp

class SetupScreen:
    def __init__(self, cfg: 'SysConfig'):
        self.cfg = cfg
        self.vaults = self.cfg.o_app.obs_vaults
        self.v_list = list(self.vaults.keys())
        self.vault_id = self.cfg.vault_id
        self.dir_vault = self.cfg.dir_vault
        self.root = tk.Tk()
        self.root.title("Obsidian Vault Health Check")
        self.root.geometry("720x540")
        self.root.resizable(True, False)
        self.root.attributes('-topmost', 1)
        self.root.iconbitmap('../img/swenlogo.ico')
        self.logo_image = Image.open('../img/SwenLogo2.png').resize((200, 200))
        self.frame_image = ImageTk.PhotoImage(self.logo_image, master=self.root)

        # Tkinter variables
        self.vault_name_var = tk.StringVar(value=self.cfg.vault_name)
        self.pn_wb_exec_var = tk.StringVar(value=self.cfg.pn_wb_exec)
        self.dirs_skip_rel_str_var = tk.StringVar(value=self.cfg.dirs_skip_rel_str)
        self.bool_shw_notes_var = tk.BooleanVar(value=self.cfg.bool_shw_notes)
        self.bool_rel_paths_var = tk.BooleanVar(value=self.cfg.bool_rel_paths)
        self.bool_summ_rows_var = tk.BooleanVar(value=self.cfg.bool_summ_rows)
        self.bool_unused_1_var = tk.BooleanVar(value=self.cfg.bool_unused_1)
        self.bool_unused_2_var = tk.BooleanVar(value=self.cfg.bool_unused_2)
        self.bool_unused_3_var = tk.BooleanVar(value=self.cfg.bool_unused_3)
        self.link_lim_vals_var = tk.StringVar(value=str(self.cfg.link_lim_vals))
        self.link_lim_tags_var = tk.StringVar(value=str(self.cfg.link_lim_tags))

        # self.vault_name_status = None
        self.wb_exec_status = None
        self.dirs_skip_rel_str_status = None
        self.link_lim_vals_label = None
        self.link_lim_tags_label = None
        self.save_button = None
        self.wb_col_max = 16300
        self.wb_col_help = f"0=Unlimited (or {self.wb_col_max} Wkbk Max)"

    def show(self):

        def update_links_help(*args):
            try:
                vals = int(self.link_lim_vals_var.get())
                self.link_lim_vals_help.config(text="(Unlimited)" if vals == 0 else self.wb_col_help)
            except ValueError:
                self.link_lim_vals_help.config(text="Invalid")
            try:
                tags = int(self.link_lim_tags_var.get())
                self.link_lim_tags_help.config(text="(Unlimited)" if tags == 0 else self.wb_col_help)
            except ValueError:
                self.link_lim_tags_help.config(text="Invalid")

        # Main App Frame ---------------------------------------------------------------------
        main_frame = ttk.Frame(self.root, padding="1", borderwidth=1, relief="ridge")
        main_frame.pack(fill="both", expand=True)
        main_frame.columnconfigure(0, weight=1)
        # main_frame.columnconfigure(1, weight=2)
        mf_row = 0
        mf_col = 0
        f1_1st_col = 0

        # Obsidian Frame ---------------------------------------------------------------------
        f1_row = 0  # f1 denotes frame nesting level one; re-used for each frame
        f1_col = f1_1st_col

        # Obsidian Vault Details Frame ---------------------------------------------------------------------
        obs_frame = ttk.LabelFrame(main_frame, text="Obsidian Vault Details ", padding="20", borderwidth=1, relief="ridge")
        obs_frame.grid(row=mf_row, column=0, sticky="nsew", pady=5, padx=(10, 10))
        obs_frame.columnconfigure(1, weight=1)

        # label
        ttk.Label(obs_frame, text="Vault Name:", width=15).grid(row=0, column=0, sticky="w", padx=5, pady=5)

        # vault name entry
        f1_col += 1
        vault_name_entry = ttk.Combobox(obs_frame, textvariable=self.vault_name_var, width=50)
        vault_name_entry['values'] = self.v_list
        vault_name_entry['state'] = 'readonly'
        vault_name_entry.current(0)
        vault_name_entry.columnconfigure(f1_col, minsize=30, weight=2)
        vault_name_entry.grid(row=f1_row, column=f1_col, sticky="ew", padx=(0, 5)) # padx=(10, 0), pady=(0,10))

        # Ignore Directories (dir_skip_rel)
        # label
        f1_row += 1
        f1_col = f1_1st_col
        ttk.Label(obs_frame, text="Directories to Ignore:\n(comma separated)").grid(row=f1_row, column=0, sticky="w", padx=5, pady=(20, 5)) # padx=15, pady=(15, 5))


        # entry
        f1_col += 1
        dirs_skip_rel_str_entry = ttk.Entry(obs_frame, textvariable=self.dirs_skip_rel_str_var, width=50)
        dirs_skip_rel_str_entry.columnconfigure((f1_col, f1_col + 1), minsize=30, weight=2)
        dirs_skip_rel_str_entry.grid(row=f1_row, column=f1_col, sticky="ew", padx=(0, 5)) #sticky="ew", pady=(15, 5), padx=(10, 0))

        # status
        f1_col += 3
        self.dirs_skip_rel_str_status = ttk.Label(obs_frame, text="", foreground="red")
        self.dirs_skip_rel_str_status.columnconfigure(f1_col, weight=1)
        self.dirs_skip_rel_str_status.grid(row=f1_row, column=f1_col, sticky="w", padx=10)

        # Options Frame ---------------------------------------------------------------------
        mf_col = 0
        mf_row = 4
        f1_row = 0  # f1 denotes frame nesting level one; re-used for each frame
        f1_col = f1_1st_col
        # opts_frame = ttk.Frame(main_frame)
        # opts_frame.grid(row=mf_row, column=mf_col, sticky="ew", pady=5, padx=(10, 0))

        opts_frame = ttk.LabelFrame(main_frame, text="Workbook Options  ", padding="20", borderwidth=1, relief="ridge")
        opts_frame.grid(row=mf_row, column=0, sticky="nsew", pady=5, padx=(10, 10)) # padx=(0, 0))
        opts_frame.columnconfigure(1, weight=1)
        # opts_frame.columnconfigure(1, weight=1)

        ck_notes = ttk.Checkbutton(opts_frame, text="Show Notes", variable=self.bool_shw_notes_var)
        ck_notes.grid(row=0, column=0, sticky="w", pady=5)
        ck_open1 = ttk.Checkbutton(opts_frame, text="For Future Use-1", variable=self.bool_unused_1_var, state='disabled')
        ck_open1.grid(row=0, column=1, sticky="w", pady=5)
        ck_fullp = ttk.Checkbutton(opts_frame, text="Use Full Paths in Links", variable=self.bool_rel_paths_var)
        ck_fullp.grid(row=1, column=0, sticky="w", pady=5)
        ck_open2 = ttk.Checkbutton(opts_frame, text="For Future Use-2", variable=self.bool_unused_2_var, state='disabled')
        ck_open2.grid(row=1, column=1, sticky="w", pady=5)

        mf_row += 1

        # Links Frame ---------------------------------------------------------------------
        # Displayed Links Maximums
        lnks_frame = ttk.LabelFrame(main_frame, text="Workbook Link Columns", padding="20", borderwidth=1, relief="ridge")
        lnks_frame.grid(row=mf_row, column=0, sticky="nsew", pady=5, padx=(10, 10))
        lnks_frame.columnconfigure(1, weight=1)

        # Label
        ttk.Label(lnks_frame, text="Values Tab Maximum Links:").grid(row=0
                                                                     , column=0
                                                                     , sticky="w"
                                                                     , pady=5
                                                                     , padx=(0, 10))

        vals_spinbox = ttk.Spinbox(lnks_frame, from_=0, to=self.wb_col_max
                                   , textvariable=self.link_lim_vals_var, width=8)
        vals_spinbox.grid(row=0, column=1, sticky="w", pady=5)

        self.link_lim_vals_help = ttk.Label(lnks_frame
                                        , text="(Unlimited)" if self.cfg.link_lim_vals == 0 else self.wb_col_help)
        self.link_lim_vals_help.grid(row=0, column=1, sticky="w", pady=5, padx=(80,0))


        ttk.Label(lnks_frame, text="Tags Tab Maximum Links:").grid(row=1
                                                                   , column=0
                                                                   , sticky="w"
                                                                   , pady=5
                                                                   , padx=(0, 10))
        tags_spinbox = ttk.Spinbox(lnks_frame, from_=0, to=self.wb_col_max
                                   , textvariable=self.link_lim_tags_var, width=8)
        tags_spinbox.grid(row=1, column=1, sticky="w", pady=5)

        self.link_lim_tags_help = ttk.Label(lnks_frame
                                        , text="(Unlimited)" if self.cfg.link_lim_tags == 0 else self.wb_col_help)
        self.link_lim_tags_help.grid(row=1, column=1, sticky="w", pady=5, padx=(80,0))


        self.link_lim_vals_var.trace('w', update_links_help)
        self.link_lim_tags_var.trace('w', update_links_help)
        mf_row += 1

        # Executable Path Frame ---------------------------------------------------------------------
        wbex_frame = ttk.LabelFrame(main_frame, text="Workbook Executable ", padding="20", borderwidth=1, relief="ridge")
        wbex_frame.grid(row=mf_row, column=0, sticky="nsew", pady=5, padx=(10, 10))
        wbex_frame.columnconfigure(1, weight=1)

        # label
        ttk.Label(wbex_frame, text="Full Path:").grid(row=0, column=0, sticky="w", padx=5, pady=5)

        # entry
        wb_exec_entry = ttk.Entry(wbex_frame, textvariable=self.pn_wb_exec_var)
        wb_exec_entry.grid(row=0, column=1, sticky="ew", padx=(0, 5))

        # browse  button
        ttk.Button(wbex_frame, text="Browse", command=self.browse_exec_path).grid(row=0, column=6, padx=(15, 0))

        # status
        self.wb_exec_status = ttk.Label(wbex_frame, text="", foreground="red")
        self.wb_exec_status.grid(row=0, column=2, sticky="nw", padx=(5, 0))
        mf_row += 1


        # logo

        mf_col = 2 # = 1
        logo_frame = ttk.Frame(main_frame)
        logo_frame.grid(row=0, column=mf_col, rowspan=5, sticky="new", pady=5, padx=(0, 0))
        logo_label = ttk.Label(logo_frame, image=self.frame_image)
        logo_label.grid(row=0, column=mf_col, sticky="ne", pady=1, padx=1)
        logo_label.columnconfigure(mf_col, weight=1)

        # s = ttk.Separator(main_frame, orient="horizontal").grid(row=6, column=mf_col, sticky="new", padx=10, pady=5)

        # Buttons - Save & Run, Cancel
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=6, column=mf_col, rowspan=2, pady=(5, 30))
        self.save_button = ttk.Button(button_frame, text="Save & Run", command=self.on_save_and_run)
        self.save_button.pack(side="top", pady=(5, 10))
        cancel_button = ttk.Button(button_frame, text="Cancel", command=self.on_cancel)
        cancel_button.pack(side="top")
        button_frame.columnconfigure(1, weight=1)
        # button_frame.pack(side=tk.TOP, pady=(5, 30))

        # Bind validation
        # self.vault_name_var.trace('w', lambda *args: self.validate_all_fields())
        self.pn_wb_exec_var.trace('w', lambda *args: self.validate_all_fields())
        self.dirs_skip_rel_str_var.trace('w', lambda *args: self.validate_all_fields())
        self.validate_all_fields()

        # Center window
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (self.root.winfo_width() // 2)
        y = (self.root.winfo_screenheight() // 2) - (self.root.winfo_height() // 2)
        self.root.geometry(f"+{x}+{y}")
        self.root.mainloop()

    def browse_exec_path(self):
        file_path = filedialog.askopenfilename(
            title="Select Spreadsheet Executable",
            initialdir=os.path.dirname(self.pn_wb_exec_var.get()) if self.pn_wb_exec_var.get() else "/",
            filetypes=[
                ("Executable files", "*.exe" if self.cfg.cfg_os == "Windows" else "*"),
                ("All files", "*.*")
            ]
        )
        if file_path:
            self.pn_wb_exec_var.set(file_path)
            self.validate_all_fields()

    def validate_all_fields(self):
        wb_exec_valid, wb_exec_msg = self.cfg.validate_pn_wb_exec(self.pn_wb_exec_var.get())
        dirs_skip_rel_str_valid, dirs_skip_rel_str_msg = self.cfg.validate_dirs_skip_rel_str(
            self.dirs_skip_rel_str_var.get(), self.dir_vault
        )
        self.wb_exec_status.config(
            text=wb_exec_msg if not wb_exec_valid else "✓",
            foreground="red" if not wb_exec_valid else "green"
        )
        self.dirs_skip_rel_str_status.config(
            text=dirs_skip_rel_str_msg if not dirs_skip_rel_str_valid else "✓" if self.dirs_skip_rel_str_var.get().strip() else "",
            foreground="red" if not dirs_skip_rel_str_valid else "green"
        )
        all_valid = wb_exec_valid and dirs_skip_rel_str_valid
        self.save_button.config(state="normal" if all_valid else "disabled")
        return all_valid

    def on_save_and_run(self):
        if self.validate_all_fields():
            self.cfg.c_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.cfg.vault_name = self.vault_name_var.get().strip()
            self.cfg.vault_id = self.vaults[self.cfg.vault_name][0]
            self.cfg.dir_vault = self.vaults[self.cfg.vault_name][1]
            self.cfg.pn_wb_exec = self.pn_wb_exec_var.get().strip()
            self.cfg.dirs_skip_rel_str = self.dirs_skip_rel_str_var.get().strip()
            self.cfg.bool_shw_notes = self.bool_shw_notes_var.get()
            self.cfg.bool_rel_paths = self.bool_rel_paths_var.get()
            self.cfg.bool_summ_rows = self.bool_summ_rows_var.get()
            self.cfg.bool_unused_1  = self.bool_unused_1_var.get()
            self.cfg.bool_unused_2  = self.bool_unused_2_var.get()
            self.cfg.bool_unused_3  = self.bool_unused_3_var.get()
            try:
                self.cfg.link_lim_vals = int(self.link_lim_vals_var.get())
                self.cfg.link_lim_tags = int(self.link_lim_tags_var.get())
            except ValueError:
                messagebox.showerror("Error", "Maximum links values must be valid numbers")
                return
            if self.cfg.save_config():
                self.root.quit()
                self.root.destroy()
                subprocess.run(["python", "v_chk.py"])
            else:
                messagebox.showerror("Error", "Failed to save configuration")

    def on_cancel(self):
        self.root.quit()
        self.root.destroy()

class SysConfig:
    def __init__(self, dbug_lvl=0):
        self.DBUG_LVL = dbug_lvl
        self.cfg_os = platform.system()
        self.o_app = ObsidianApp()
        self.vault_name = self.o_app.dflt_vault_name
        self.vault_id   = self.o_app.obs_vaults[self.vault_name][0]
        self.dir_vault  = self.o_app.obs_vaults[self.vault_name][1]
        self.pn_wb_exec = self.o_app.dflt_wb_exec
        self.obs_vaults = self.o_app.obs_vaults
        self.cfg_sys_id = 'v_chk'
        self.cfg_sys_ver = "0.9.2"
        self.c_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.sys_dir = f"{Path(__file__).parent.parent.absolute()}/"
        self.dir_batch = f"{self.sys_dir}data/batch_files/"
        self.dir_wbs = f"{self.sys_dir}data/workbooks/"
        self.tab_seq = [ 'pros', 'vals', 'tags', 'file', 'code', 'xyml', 'dups', 'tmpl', 'nest', 'plug', 'summ', 'ar51' ]
        self.dirs_dot = []
        self.dirs_skip_rel_str = ""
        self.dirs_skip_abs_lst = []
        self.dir_templates = ""
        self.ctot = [0] * 13
        self.pn_batch = ""
        self.pn_wbs = ""
        self.pn_cfg = ""
        self.bat_num = 0
        self.bool_shw_notes = True
        self.bool_rel_paths = True
        self.bool_summ_rows = True
        self.bool_unused_1  = False
        self.bool_unused_2  = False
        self.bool_unused_3  = False
        self.link_lim_vals = 0
        self.link_lim_tags = 0
        self.cfg = {}
        pname = Path(self.sys_dir).joinpath(f"CONFIG.yaml")
        self.pn_cfg = f"{pname}"
        data_dir = f"{self.sys_dir}data"
        self.mkdirs([data_dir, self.sys_dir, self.dir_batch])
        if os.path.exists(self.pn_cfg):
            self.load_config()
            if not self.chk_fields_on_load():
                SetupScreen(self).show()
        else:
            SetupScreen(self).show()

        print("x")

    def chk_fields_on_load(self):
        dir_vault_valid, _ = self.validate_dir_vault(self.dir_vault)
        wb_exec_valid, _ = self.validate_pn_wb_exec(self.pn_wb_exec)
        return dir_vault_valid and wb_exec_valid

    def mkdirs(self, path):
        if isinstance(path, list):
            for p in path:
                self.mkdirs(p)
            return
        if os.path.isdir(path):
            return
        os.makedirs(path)

    def get_templates_dir(self):
        template_cfg_file = f"{self.dir_vault}/.obsidian/plugins/templater-obsidian/data.json"
        try:
            if os.path.isfile(template_cfg_file):
                with open(template_cfg_file, 'r') as f:
                    template_cfg_json = f.read()
                template_cfg = json.loads(template_cfg_json)
                templates_path = Path(self.dir_vault).joinpath(template_cfg['templates_folder'])
                return f"{templates_path}"
        except (FileNotFoundError, KeyError):
            return None
        except Exception as e:
            raise Exception(f"ConfigSys: Error in get_templates_dir: {e}")
        else:
            return None

    def load_config(self):
        try:
            with open(self.pn_cfg, 'r') as file:
                self.cfg = yaml.safe_load(file)
                self.cfg_unpack()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load config: {str(e)}")

    def save_config(self):
        self.cfg_pack()
        try:
            with open(self.pn_cfg, 'w') as file:
                yaml.dump(self.cfg, file, default_flow_style=False)
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save config: {str(e)}")
            return False

    def cfg_pack(self):
        # path = Path(self.dir_templates.strip())
        # if not path.exists() or not path.is_dir() or self.dir_templates == '':
        self.dir_templates = self.get_templates_dir()
        self.dirs_dot = [f.name for f in os.scandir(self.dir_vault) if
                         f.is_dir() and f.path.startswith(f"{self.dir_vault}/.")]
        dirs = [d.strip() for d in self.dirs_skip_rel_str.split(',') if d.strip()]
        self.dirs_skip_abs_lst = []
        for dir_name in dirs:
            dname = Path(self.dir_vault).joinpath(dir_name)
            self.dirs_skip_abs_lst += [str(dname)]
        self.cfg = {
              'vault_id':           self.vault_id
            , 'vault_name':         self.vault_name
            , 'dir_vault':          self.dir_vault
            , 'pn_wb_exec':         self.pn_wb_exec
            , 'cfg_sys_id':         self.cfg_sys_id
            , 'cfg_sys_ver':        self.cfg_sys_ver
            , 'c_date':             self.c_date
            , 'dir_batch':          self.dir_batch
            , 'dir_wbs':            self.dir_wbs
            , 'dir_templates':      self.dir_templates
            , 'tab_seq':            self.tab_seq
            , 'ctot':               self.ctot
            , 'dirs_dot':           self.dirs_dot
            , 'dirs_skip_rel_str':  self.dirs_skip_rel_str
            , 'dirs_skip_abs_lst':  self.dirs_skip_abs_lst
            , 'pn_batch':           self.pn_batch
            , 'pn_wbs':             self.pn_wbs
            , 'pn_cfg':             self.pn_cfg
            , 'bat_num':            self.bat_num
            , 'bool_shw_notes':     self.bool_shw_notes
            , 'bool_rel_paths':     self.bool_rel_paths
            , 'bool_summ_rows':     self.bool_summ_rows
            , 'bool_unused_1':      self.bool_unused_1
            , 'bool_unused_2':      self.bool_unused_2
            , 'bool_unused_3':      self.bool_unused_3
            , 'link_lim_vals':      self.link_lim_vals
            , 'link_lim_tags':      self.link_lim_tags
            , 'obs_vaults':         self.obs_vaults
        }

    def cfg_unpack(self):
        self.vault_id           = self.cfg.get('vault_id', '')
        self.vault_name         = self.cfg.get('vault_name', '')
        self.dir_vault          = self.cfg.get('dir_vault', '')
        self.pn_wb_exec         = self.cfg.get('pn_wb_exec', '')
        self.cfg_sys_id         = self.cfg.get('cfg_sys_id', 'v_chk')
        self.cfg_sys_ver        = self.cfg.get('cfg_sys_ver', '0.7')
        self.c_date             = self.cfg.get('c_date', datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        self.dir_batch          = self.cfg.get('dir_batch', f"{self.sys_dir}data/batch_files/")
        self.dir_wbs            = self.cfg.get('dir_wbs', f"{self.sys_dir}data/workbooks/")
        self.dir_templates      = self.cfg.get('dir_templates', self.get_templates_dir())
        self.tab_seq            = self.cfg.get('tab_seq', '')
        self.ctot               = self.cfg.get('ctot', '')
        self.dirs_dot           = self.cfg.get('dirs_dot', '')
        self.dirs_skip_rel_str  = self.cfg.get('dirs_skip_rel_str', '')
        self.dirs_skip_abs_lst  = self.cfg.get('dirs_skip_abs_lst', '')
        self.pn_batch           = self.cfg.get('pn_batch', '')
        self.pn_wbs             = self.cfg.get('pn_wbs', '')
        self.pn_cfg             = self.cfg.get('pn_cfg', '')
        self.bat_num            = self.cfg.get('bat_num', 0)
        self.bool_shw_notes     = self.cfg.get('bool_shw_notes', True)
        self.bool_rel_paths     = self.cfg.get('bool_rel_paths', True)
        self.bool_summ_rows     = self.cfg.get('bool_summ_rows', True)
        self.bool_unused_1      = self.cfg.get('bool_unused_1', True)
        self.bool_unused_2      = self.cfg.get('bool_unused_2', True)
        self.bool_unused_3      = self.cfg.get('bool_unused_3', True)
        self.link_lim_vals      = self.cfg.get('link_lim_vals', 0)
        self.link_lim_tags      = self.cfg.get('link_lim_tags', 0)
        self.obs_vaults         = self.cfg.get('obs_vaults', [])

    @staticmethod
    def validate_vault_id(vault_id):
        if not vault_id or not vault_id.strip():
            return False, "Vault ID cannot be empty"
        if len(vault_id.strip()) < 1:
            return False, "Vault ID must be at least 1 character"
        return True, ""

    @staticmethod
    def validate_dir_vault(dir_vault):
        if not dir_vault or not dir_vault.strip():
            return False, "Vault path cannot be empty"
        path = Path(dir_vault.strip())
        if not path.exists():
            return False, "Vault path does not exist"
        if not path.is_dir():
            return False, "Vault path must be a directory"
        return True, ""

    @staticmethod
    def validate_dirs_skip_rel_str(dirs_skip_rel_str, dir_vault):
        if not dirs_skip_rel_str or not dirs_skip_rel_str.strip():
            return True, ""
        if not dir_vault or not dir_vault.strip():
            return False, "Vault path must be set first"
        dir_vault_obj = Path(dir_vault.strip())
        if not dir_vault_obj.exists():
            return False, "Vault path must be valid first"
        dirs = [d.strip() for d in dirs_skip_rel_str.split(',') if d.strip()]
        if not dirs:
            return True, ""
        for dir_name in dirs:
            found = False
            for root, dirs_list, files in os.walk(dir_vault_obj):
                if dir_name in dirs_list:
                    found = True
                    break
            if not found:
                return False, f"X"
        return True, ""

    @staticmethod
    def validate_pn_wb_exec(pn_wb_exec):
        if not pn_wb_exec or not pn_wb_exec.strip():
            return False, "Executable path cannot be empty"
        path = Path(pn_wb_exec.strip())
        if not path.exists():
            return False, "Executable file does not exist"
        if not path.is_file():
            return False, "Executable path must be a file"
        if not os.access(path, os.X_OK):
            return False, "File is not executable"
        return True, ""


if __name__ == "__main__":
    sys_cfg = SysConfig()
    setup_screen = SetupScreen(sys_cfg).show()
