import os
import yaml
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from datetime import datetime
import subprocess
import platform
import json

class SysConfig:
    def __init__(self, dbug_lvl=0):
        self.vault_id       = ""
        self.dir_vault     = ""
        self.pn_wb_exec   = ""
        self.DBUG_LVL = dbug_lvl
        self.cfg_sys_id = 'v_chk'
        self.cfg_sys_ver = "0.7"
        # Instantiate default values, presumably overridden in read_cfg_sys
        self.c_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.sys_dir = f"{Path(__file__).parent.parent.absolute()}\\"
        self.dir_batch = f"{self.sys_dir}data\\batch_files\\"
        self.dir_wbs = f"{self.sys_dir}data\\workbooks\\"
        self.tab_seq = [ 'pros'   # ONLY USED FOR NEW INSTALL! This is defined in CONFIG.yaml
                       , 'vals'
                       , 'tags'
                       , 'file'
                       , 'code'
                       , 'xyml'
                       , 'dups'
                       , 'tmpl'
                       , 'nest'
                       , 'plug'
                       , 'summ'
                       , 'ar51'
        ]


        self.dirs_dot       = []
        self.dirs_skip_rel_str    = ""
        self.dirs_skip_abs_lst      = []
        self.dir_templates        = ""
        self.ctot           = [0] * 13

        self.pn_batch = ""
        self.pn_wbs = ""
        self.pn_cfg = ""
        self.bat_num   = 0
        
        self.bool_shw_notes = True
        self.bool_rel_paths = True
        self.bool_summ_rows = True
        self.bool_unused_1  = False
        self.bool_unused_2  = False
        self.bool_unused_3  = False

        self.link_lim_vals = 0  # Values Tab Maximum Links
        self.link_lim_tags = 0  # Tags Tab Maximum Links

        self.cfg = {}

        # NOTE: this is the SYSTEM config, not the wb runtime config
        pname = Path(self.sys_dir).joinpath(f"CONFIG.yaml")
        self.pn_cfg = f"{pname}" # prefer a string, rather than the Path Object

        # self.dir_templates = self.get_templates_dir()

        # make sure directories exist
        data_dir = f"{self.sys_dir}data"
        self.mkdirs([data_dir, self.sys_dir, self.dir_batch])

        # Load existing config or create new one
        if os.path.exists(self.pn_cfg):
            self.load_config()

            if not self.chk_fields_on_load():
                self.show_config_gui()
        else:
            self.show_config_gui()

        x = 'debug'

    def chk_fields_on_load(self):
        """Check if fields are valid on load"""
        # vault_id_valid, vault_id_msg     = self.validate_vault_id(self.vault_id)
        dir_vault_valid, dir_vault_msg = self.validate_dir_vault(self.dir_vault)
        wb_exec_valid, wb_exec_msg       = self.validate_pn_wb_exec(self.pn_wb_exec)
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
        template_cfg_file = f"{self.dir_vault}\\.obsidian\\plugins\\templater-obsidian\\data.json"
        try:
            if os.path.isfile(template_cfg_file):
                with open(template_cfg_file, 'r') as f:
                    template_cfg_json = f.read()
                template_cfg = json.loads(template_cfg_json)
                templates_path = Path(self.dir_vault).joinpath(template_cfg['templates_folder'])
                return f"{templates_path}"


        except FileNotFoundError or KeyError:
            return None

        except Exception as e:
            raise Exception(f"ConfigSys: Error in get_templates_dir: {e}")
        else:
            return None

    def load_config(self):
        """Load configuration from YAML file"""
        try:
            with open(self.pn_cfg, 'r') as file:
                self.cfg = yaml.safe_load(file)
                self.cfg_unpack()

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load config: {str(e)}")

    def save_config(self):
        """Save configuration to YAML file"""
        self.cfg_pack()

        try:
            with open(self.pn_cfg, 'w') as file:
                yaml.dump(self.cfg, file, default_flow_style=False)
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save config: {str(e)}")
            return False

    def cfg_pack(self):
        path = Path(self.dir_templates.strip())
        if not path.exists() or not path.is_dir() or self.dir_templates == '':
            self.dir_templates = self.get_templates_dir()

        self.dirs_dot = [f.name for f in os.scandir(self.dir_vault) if
                         f.is_dir() and f.path.startswith(f"{self.dir_vault}\\.")]

        dirs = [d.strip() for d in self.dirs_skip_rel_str.split(',') if d.strip()]
        self.dirs_skip_abs_lst = []
        for dir_name in dirs:
            dname = Path(self.dir_vault).joinpath(dir_name)
            self.dirs_skip_abs_lst += [str(dname)]

        self.cfg = {
              'vault_id':           self.vault_id
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
        }

    def cfg_unpack(self):
        self.vault_id           = self.cfg.get('vault_id', '')
        self.dir_vault         = self.cfg.get('dir_vault', '')
        self.pn_wb_exec       = self.cfg.get('pn_wb_exec', '')
        self.cfg_sys_id         = self.cfg.get('cfg_sys_id', 'v_chk')
        self.cfg_sys_ver        = self.cfg.get('cfg_sys_ver', '0.7')
        self.c_date             = self.cfg.get('c_date', datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        self.dir_batch            = self.cfg.get('dir_batch', f"{self.sys_dir}data\\batch_files\\")
        self.dir_wbs            = self.cfg.get('dir_wbs', f"{self.sys_dir}data\\workbooks\\")
        self.dir_templates            = self.cfg.get('dir_templates', self.get_templates_dir())
        self.tab_seq            = self.cfg.get('tab_seq', '')
        self.ctot               = self.cfg.get('ctot', '')
        self.dirs_dot           = self.cfg.get('dirs_dot', '')
        self.dirs_skip_rel_str  = self.cfg.get('dirs_skip_rel_str', '')
        self.dirs_skip_abs_lst  = self.cfg.get('dirs_skip_abs_lst', '')
        self.pn_batch          = self.cfg.get('pn_batch', '')
        self.pn_wbs          = self.cfg.get('pn_wbs', '')
        self.pn_cfg          = self.cfg.get('pn_cfg', '')
        self.bat_num            = self.cfg.get('bat_num', 0)
        self.bool_shw_notes     = self.cfg.get('bool_shw_notes', True)
        self.bool_rel_paths     = self.cfg.get('bool_rel_paths', True)
        self.bool_summ_rows     = self.cfg.get('bool_summ_rows', True)
        self.bool_unused_1      = self.cfg.get('bool_unused_1', True)
        self.bool_unused_2      = self.cfg.get('bool_unused_2', True)
        self.bool_unused_3      = self.cfg.get('bool_unused_3', True)
        self.link_lim_vals = self.cfg.get('link_lim_vals', 0)
        self.link_lim_tags = self.cfg.get('link_lim_tags', 0)

    def validate_vault_id(self, vault_id):
        """Validate Obsidian Vault ID"""
        if not vault_id or not vault_id.strip():
            return False, "Vault ID cannot be empty"

        # Basic validation - should be alphanumeric with some special chars
        if len(vault_id.strip()) < 1:
            return False, "Vault ID must be at least 1 character"

        return True, ""

    def validate_dir_vault(self, dir_vault):
        """Validate Obsidian Vault File Path"""
        if not dir_vault or not dir_vault.strip():
            return False, "Vault path cannot be empty"

        path = Path(dir_vault.strip())
        if not path.exists():
            return False, "Vault path does not exist"

        if not path.is_dir():
            return False, "Vault path must be a directory"

        return True, ""

    def validate_dirs_skip_rel_str(self, dirs_skip_rel_str, dir_vault):
        """Validate directories to ignore"""
        if not dirs_skip_rel_str or not dirs_skip_rel_str.strip():
            return True, ""  # Empty is valid

        if not dir_vault or not dir_vault.strip():
            return False, "Vault path must be set first"

        dir_vault_obj = Path(dir_vault.strip())
        if not dir_vault_obj.exists():
            return False, "Vault path must be valid first"

        dirs = [d.strip() for d in dirs_skip_rel_str.split(',') if d.strip()]
        if not dirs:
            return True, ""

        for dir_name in dirs:
            # Check if directory exists anywhere under dir_vault
            found = False
            for root, dirs_list, files in os.walk(dir_vault_obj):
                if dir_name in dirs_list:
                    found = True
                    break

            if not found:
                return False, f"Directory '{dir_name}' not found under vault path"

        return True, ""

    def validate_pn_wb_exec(self, pn_wb_exec):
        """Validate spreadsheet executable path"""
        if not pn_wb_exec or not pn_wb_exec.strip():
            return False, "Executable path cannot be empty"

        path = Path(pn_wb_exec.strip())
        if not path.exists():
            return False, "Executable file does not exist"

        if not path.is_file():
            return False, "Executable path must be a file"

        # Check if file is executable
        if not os.access(path, os.X_OK):
            return False, "File is not executable"

        return True, ""

    def browse_dir_vault(self):
        """Open directory browser for vault path"""
        folder_path = filedialog.askdirectory(
            title="Select Obsidian Vault Directory",
            initialdir=self.dir_vault if self.dir_vault else "/"
        )
        if folder_path:
            self.dir_vault_var.set(folder_path)
            self.validate_all_fields()

    def browse_exec_path(self):
        """Open file browser for executable path"""
        file_path = filedialog.askopenfilename(
            title="Select Spreadsheet Executable",
            initialdir=os.path.dirname(self.pn_wb_exec) if self.pn_wb_exec else "/",
            filetypes=[
                ("Executable files", "*.exe" if platform.system() == "Windows" else "*"),
                ("All files", "*.*")
            ]
        )
        if file_path:
            self.pn_wb_exec_var.set(file_path)
            self.validate_all_fields()

    def validate_all_fields(self):
        """Validate all fields and update save button state"""
        # vault_id_valid, vault_id_msg = self.validate_vault_id(self.vault_id_var.get())
        dir_vault_valid, dir_vault_msg = self.validate_dir_vault(self.dir_vault_var.get())
        wb_exec_valid, wb_exec_msg = self.validate_pn_wb_exec(self.pn_wb_exec_var.get())
        dirs_skip_rel_str_valid, dirs_skip_rel_str_msg = self.validate_dirs_skip_rel_str(
            self.dirs_skip_rel_str_var.get(), self.dir_vault_var.get()
        )

        # Update status labels
        # self.vault_id_status.config(
        #     text=vault_id_msg if not vault_id_valid else "✓ Valid",
        #     foreground="red" if not vault_id_valid else "green"
        # )
        self.dir_vault_status.config(
            text=dir_vault_msg if not dir_vault_valid else "✓ Valid",
            foreground="red" if not dir_vault_valid else "green"
        )
        self.wb_exec_status.config(
            text=wb_exec_msg if not wb_exec_valid else "✓ Valid",
            foreground="red" if not wb_exec_valid else "green"
        )
        self.dirs_skip_rel_str_status.config(
            text=dirs_skip_rel_str_msg if not dirs_skip_rel_str_valid else "✓ Valid" if self.dirs_skip_rel_str_var.get().strip() else "",
            foreground="red" if not dirs_skip_rel_str_valid else "green"
        )

        # Enable/disable save button
        all_valid = dir_vault_valid and wb_exec_valid and dirs_skip_rel_str_valid
        self.save_button.config(state="normal" if all_valid else "disabled")

        return all_valid

    def on_save_and_run(self):
        """Handle Save & Run button click"""
        if self.validate_all_fields():
            # self.vault_id = self.vault_id_var.get().strip()
            self.dir_vault = self.dir_vault_var.get().strip()
            self.pn_wb_exec = self.pn_wb_exec_var.get().strip()
            self.dirs_skip_rel_str = self.dirs_skip_rel_str_var.get().strip()
            self.bool_shw_notes = self.bool_shw_notes_var.get()
            self.bool_rel_paths = self.bool_rel_paths_var.get()
            self.bool_summ_rows = self.bool_summ_rows_var.get()
            self.bool_unused_1  = self.bool_unused_1_var.get()
            self.bool_unused_2  = self.bool_unused_2_var.get()
            self.bool_unused_3  = self.bool_unused_3_var.get()

            try:
                self.link_lim_vals = int(self.link_lim_vals_var.get())
                self.link_lim_tags = int(self.link_lim_tags_var.get())
            except ValueError:
                messagebox.showerror("Error", "Maximum links values must be valid numbers")
                return

            if self.save_config():
                #  messagebox.showinfo("Success", "Configuration saved successfully!")
                self.root.quit()
                self.root.destroy()
                subprocess.run(["python", "v_chk.py"])
            else:
                messagebox.showerror("Error", "Failed to save configuration")

    def on_cancel(self):
        """Handle Cancel button click"""
        self.root.quit()
        self.root.destroy()

    def show_config_gui(self):
        """Show the configuration GUI"""
        self.root = tk.Tk()
        self.root.title("Obsidian Vault Health Check")
        self.root.geometry("700x600")
        self.root.resizable(True, True)
        self.root.iconbitmap('..\\img\\swenlogo.ico')

        # Create main frame with scrollbar
        canvas = tk.Canvas(self.root)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Main content frame
        main_frame = ttk.Frame(scrollable_frame, padding="20")
        main_frame.pack(fill="both", expand=True)
        main_frame.columnconfigure(1, weight=1)

        # Create StringVar and BooleanVar objects for form fields
        self.dir_vault_var         = tk.StringVar(value=self.dir_vault)
        self.pn_wb_exec_var       = tk.StringVar(value=self.pn_wb_exec)
        self.dirs_skip_rel_str_var  = tk.StringVar(value=self.dirs_skip_rel_str)
        self.bool_shw_notes_var     = tk.BooleanVar(value=self.bool_shw_notes)
        self.bool_rel_paths_var     = tk.BooleanVar(value=self.bool_rel_paths)
        self.bool_summ_rows_var     = tk.BooleanVar(value=self.bool_summ_rows)
        self.bool_unused_1_var      = tk.BooleanVar(value=self.bool_unused_1)
        self.bool_unused_2_var      = tk.BooleanVar(value=self.bool_unused_2)
        self.bool_unused_3_var      = tk.BooleanVar(value=self.bool_unused_3)

        row = 0

        # Vault Path field
        ttk.Label(main_frame, text="Obsidian Vault Path:").grid(row=row, column=0, sticky=tk.W, pady=5)
        dir_vault_frame = ttk.Frame(main_frame)
        dir_vault_frame.grid(row=row, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
        dir_vault_frame.columnconfigure(0, weight=1)

        dir_vault_entry = ttk.Entry(dir_vault_frame, textvariable=self.dir_vault_var)
        dir_vault_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        ttk.Button(dir_vault_frame, text="Browse", command=self.browse_dir_vault).grid(row=0, column=1)

        row += 1
        self.dir_vault_status = ttk.Label(main_frame, text="", foreground="red")
        self.dir_vault_status.grid(row=row, column=1, sticky=tk.W, padx=(10, 0))

        row += 1

        # Executable Path field
        ttk.Label(main_frame, text="Spreadsheet Executable:").grid(row=row, column=0, sticky=tk.W, pady=(15, 5))


        wb_exec_frame = ttk.Frame(main_frame)
        wb_exec_frame.grid(row=row, column=1, sticky=(tk.W, tk.E), pady=(15, 5), padx=(10, 0))
        wb_exec_frame.columnconfigure(0, weight=1)

        wb_exec_entry = ttk.Entry(wb_exec_frame, textvariable=self.pn_wb_exec_var)
        wb_exec_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        ttk.Button(wb_exec_frame, text="Browse", command=self.browse_exec_path).grid(row=0, column=1)

        row += 1
        self.wb_exec_status = ttk.Label(main_frame, text="", foreground="red")
        self.wb_exec_status.grid(row=row, column=1, sticky=tk.W, padx=(10, 0))

        row += 1

        # Ignore Directories field
        ttk.Label(main_frame, text="Directories to Ignore\n(comma separated):").grid(row=row, column=0, sticky=tk.W,
                                                                                     pady=(15, 5))

        dirs_skip_rel_str_entry = ttk.Entry(main_frame, textvariable=self.dirs_skip_rel_str_var, width=50)
        dirs_skip_rel_str_entry.grid(row=row, column=1, sticky=(tk.W, tk.E), pady=(15, 5), padx=(10, 0))

        row += 1
        self.dirs_skip_rel_str_status = ttk.Label(main_frame, text="", foreground="red")
        self.dirs_skip_rel_str_status.grid(row=row, column=1, sticky=tk.W, padx=(10, 0))

        row += 1

        # Options section
        ttk.Label(main_frame, text="Options:", font=("TkDefaultFont", 10, "bold")).grid(
            row=row, column=0, columnspan=2, sticky=tk.W, pady=(20, 10)
        )

        row += 1
        options_frame = ttk.Frame(main_frame)
        options_frame.grid(row=row, column=0, columnspan=2, sticky=(tk.W, tk.E), padx=(0, 0))
        options_frame.columnconfigure(0, weight=1)
        options_frame.columnconfigure(1, weight=1)

        # First row of checkboxes
        ttk.Checkbutton(options_frame, text="Show Notes", variable=self.bool_shw_notes_var).grid(
            row=0, column=0, sticky=tk.W, pady=5
        )
        ttk.Checkbutton(options_frame, text="For Future Use-1", variable=self.bool_unused_1_var, state='disabled').grid(
            row=0, column=1, sticky=tk.W, pady=5
        )

        # Second row of checkboxes
        ttk.Checkbutton(options_frame, text="Use Full Paths in Links", variable=self.bool_rel_paths_var).grid(
            row=1, column=0, sticky=tk.W, pady=5
        )
        ttk.Checkbutton(options_frame, text="For Future Use-2", variable=self.bool_unused_2_var, state='disabled').grid(
            row=1, column=1, sticky=tk.W, pady=5
        )

        # Third row of checkboxes
        ttk.Checkbutton(options_frame, text="Show Rows in Summary", variable=self.bool_summ_rows_var, state='disabled').grid(
            row=2, column=0, sticky=tk.W, pady=5
        )
        ttk.Checkbutton(options_frame, text="For Future Use-3", variable=self.bool_unused_3_var, state='disabled').grid(
            row=2, column=1, sticky=tk.W, pady=5
        )

        row += 1

        # Displayed Links Maximums section
        ttk.Label(main_frame, text="Displayed Links Maximums:", font=("TkDefaultFont", 10, "bold")).grid(
            row=row, column=0, columnspan=2, sticky=tk.W, pady=(20, 10)
        )

        row += 1
        links_frame = ttk.Frame(main_frame)
        links_frame.grid(row=row, column=0, columnspan=2, sticky=(tk.W, tk.E), padx=(0, 0))
        links_frame.columnconfigure(1, weight=1)
        links_frame.columnconfigure(3, weight=1)

        # Values Tab Maximum Links
        ttk.Label(links_frame, text="Values Tab Maximum Links:").grid(row=0, column=0, sticky=tk.W, pady=5, padx=(0, 10))
        self.link_lim_vals_var = tk.StringVar(value=str(self.link_lim_vals))
        vals_spinbox = ttk.Spinbox(links_frame, from_=0, to=16300, textvariable=self.link_lim_vals_var, width=10)
        vals_spinbox.grid(row=0, column=1, sticky=tk.W, pady=5)
        self.link_lim_vals_label = ttk.Label(links_frame, text="Unlimited" if self.link_lim_vals == 0 else str(self.link_lim_vals))
        self.link_lim_vals_label.grid(row=0, column=2, sticky=tk.W, pady=5, padx=10)

        # Tags Tab Maximum Links
        ttk.Label(links_frame, text="Tags Tab Maximum Links:").grid(row=1, column=0, sticky=tk.W, pady=5, padx=(0, 10))
        self.link_lim_tags_var = tk.StringVar(value=str(self.link_lim_tags))
        tags_spinbox = ttk.Spinbox(links_frame, from_=0, to=16300, textvariable=self.link_lim_tags_var, width=10)
        tags_spinbox.grid(row=1, column=1, sticky=tk.W, pady=5)
        self.link_lim_tags_label = ttk.Label(links_frame, text="Unlimited" if self.link_lim_tags == 0 else str(self.link_lim_tags))
        self.link_lim_tags_label.grid(row=1, column=2, sticky=tk.W, pady=5, padx=10)

        def update_links_label(*args):
            try:
                vals = int(self.link_lim_vals_var.get())
                self.link_lim_vals_label.config(text="Unlimited" if vals == 0 else str(vals))
            except ValueError:
                self.link_lim_vals_label.config(text="Invalid")

            try:
                tags = int(self.link_lim_tags_var.get())
                self.link_lim_tags_label.config(text="Unlimited" if tags == 0 else str(tags))
            except ValueError:
                self.link_lim_tags_label.config(text="Invalid")

        self.link_lim_vals_var.trace('w', update_links_label)
        self.link_lim_tags_var.trace('w', update_links_label)

        row += 1

        # Buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=row, column=0, columnspan=2, pady=(30, 0))

        self.save_button = ttk.Button(button_frame, text="Save & Run", command=self.on_save_and_run)
        self.save_button.pack(side=tk.LEFT, padx=(0, 10))

        cancel_button = ttk.Button(button_frame, text="Cancel", command=self.on_cancel)
        cancel_button.pack(side=tk.LEFT)

        # Bind validation to field changes
        self.dir_vault_var.trace('w', lambda *args: self.validate_all_fields())
        self.pn_wb_exec_var.trace('w', lambda *args: self.validate_all_fields())
        self.dirs_skip_rel_str_var.trace('w', lambda *args: self.validate_all_fields())

        # Initial validation
        self.validate_all_fields()

        # Center the window
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (self.root.winfo_width() // 2)
        y = (self.root.winfo_screenheight() // 2) - (self.root.winfo_height() // 2)
        self.root.geometry(f"+{x}+{y}")

        # Start the GUI
        # keep the window displaying, but first, if windows, make ensure not blurry
        # try:
        #     from ctypes import windll
        #     windll.shcore.SetProcessDpiAwareness(1)
        # finally:
        #     self.root.mainloop()
        self.root.mainloop()

    def get_config(self):
        """Return current configuration as dictionary"""
        return self.cfg

# Example usage
if __name__ == "__main__":
    # Create config manager instance
    config_manager = SysConfig()
    config_manager.show_config_gui()


    # Access configuration data
    config = config_manager.get_config()
    # print("Current configuration:")
    # print(f"Vault Path: {config['dir_vault']}")
    # print(f"Executable Path: {config['pn_wb_exec']}")
    # print(f"Ignore Directories: {config['dirs_skip_rel_str']}")
    # print(f"Show Notes: {config['show_notes']}")
    # print(f"Use Full Paths: {config['use_full_paths']}")
    # print(f"Show Rows Summary: {config['show_rows_summary']}")
    # print(f"Future Use 1: {config['future_use_1']}")
    # print(f"Future Use 2: {config['future_use_2']}")
    # print(f"Future Use 3: {config['future_use_3']}")