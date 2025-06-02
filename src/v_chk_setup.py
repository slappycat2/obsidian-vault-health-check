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
        self.vault_path     = ""
        self.wb_exec_path   = ""
        self.DBUG_LVL = dbug_lvl
        self.cfg_sys_id = 'v_chk'
        self.cfg_sys_ver = "0.7"
        # Instantiate default values, presumably overridden in read_cfg_sys
        self.c_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.sys_dir = f"{Path(__file__).parent.parent.absolute()}\\"
        self.bat_dir = f"{self.sys_dir}data\\batch_files\\"
        self.xls_dir = f"{self.sys_dir}data\\workbooks\\"
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
        self.tmpldir        = ""
        self.ctot           = [0] * 10

        self.bat_pname = ""
        self.xls_pname = ""
        self.cfg_pname = ""
        self.bat_num   = 0
        
        self.bool_shw_notes = True
        self.bool_rel_paths = True
        self.bool_summ_rows = True
        self.bool_unused_1  = False
        self.bool_unused_2  = False
        self.bool_unused_3  = False

        self.cfg = {}

        # NOTE: this is the SYSTEM config, not the wb runtime config
        pname = Path(self.sys_dir).joinpath(f"CONFIG.yaml")
        self.cfg_pname = f"{pname}" # prefer a string, rather than the Path Object

        # self.tmpldir = self.get_templates_dir()

        # make sure directories exist
        data_dir = f"{self.sys_dir}data"
        self.mkdirs([data_dir, self.sys_dir, self.bat_dir])

        # Load existing config or create new one
        if os.path.exists(self.cfg_pname):
            self.load_config()

            if not self.chk_fields_on_load():
                self.show_config_gui()
        else:
            self.show_config_gui()

        x = 'debug'

    def chk_fields_on_load(self):
        """Check if fields are valid on load"""
        # vault_id_valid, vault_id_msg     = self.validate_vault_id(self.vault_id)
        vault_path_valid, vault_path_msg = self.validate_vault_path(self.vault_path)
        wb_exec_valid, wb_exec_msg       = self.validate_wb_exec_path(self.wb_exec_path)
        return vault_path_valid and wb_exec_valid

    def mkdirs(self, path):
        if isinstance(path, list):
            for p in path:
                self.mkdirs(p)
            return

        if os.path.isdir(path):
            return

        os.makedirs(path)

    def get_templates_dir(self):
        template_cfg_file = f"{self.vault_path}\\.obsidian\\plugins\\templater-obsidian\\data.json"
        try:
            if os.path.isfile(template_cfg_file):
                with open(template_cfg_file, 'r') as f:
                    template_cfg_json = f.read()
                template_cfg = json.loads(template_cfg_json)
                templates_path = Path(self.vault_path).joinpath(template_cfg['templates_folder'])
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
            with open(self.cfg_pname, 'r') as file:
                self.cfg = yaml.safe_load(file)
                self.cfg_unpack()

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load config: {str(e)}")

    def save_config(self):
        """Save configuration to YAML file"""
        self.cfg_pack()

        try:
            with open(self.cfg_pname, 'w') as file:
                yaml.dump(self.cfg, file, default_flow_style=False)
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save config: {str(e)}")
            return False

    def cfg_pack(self):
        path = Path(self.tmpldir.strip())
        if not path.exists() or not path.is_dir() or self.tmpldir == '':
            self.tmpldir = self.get_templates_dir()

        self.dirs_dot = [f.name for f in os.scandir(self.vault_path) if
                         f.is_dir() and f.path.startswith(f"{self.vault_path}\\.")]

        dirs = [d.strip() for d in self.dirs_skip_rel_str.split(',') if d.strip()]
        self.dirs_skip_abs_lst = []
        for dir_name in dirs:
            dname = Path(self.vault_path).joinpath(dir_name)
            self.dirs_skip_abs_lst += [str(dname)]

        self.cfg = {
              'vault_id':           self.vault_id
            , 'vault_path':         self.vault_path
            , 'wb_exec_path':       self.wb_exec_path
            , 'cfg_sys_id':         self.cfg_sys_id
            , 'cfg_sys_ver':        self.cfg_sys_ver
            , 'c_date':             self.c_date
            , 'bat_dir':            self.bat_dir
            , 'xls_dir':            self.xls_dir
            , 'tmpldir':            self.tmpldir
            , 'tab_seq':            self.tab_seq
            , 'ctot':               self.ctot
            , 'dirs_dot':           self.dirs_dot
            , 'dirs_skip_rel_str':  self.dirs_skip_rel_str
            , 'dirs_skip_abs_lst':  self.dirs_skip_abs_lst
            , 'bat_pname':          self.bat_pname
            , 'xls_pname':          self.xls_pname
            , 'cfg_pname':          self.cfg_pname
            , 'bat_num':            self.bat_num
            , 'bool_shw_notes':     self.bool_shw_notes
            , 'bool_rel_paths':     self.bool_rel_paths
            , 'bool_summ_rows':     self.bool_summ_rows
            , 'bool_unused_1':      self.bool_unused_1
            , 'bool_unused_2':      self.bool_unused_2
            , 'bool_unused_3':      self.bool_unused_3
        }

    def cfg_unpack(self):
        self.vault_id           = self.cfg.get('vault_id', '')
        self.vault_path         = self.cfg.get('vault_path', '')
        self.wb_exec_path       = self.cfg.get('wb_exec_path', '')
        self.cfg_sys_id         = self.cfg.get('cfg_sys_id', 'v_chk')
        self.cfg_sys_ver        = self.cfg.get('cfg_sys_ver', '0.7')
        self.c_date             = self.cfg.get('c_date', datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        self.bat_dir            = self.cfg.get('bat_dir', f"{self.sys_dir}data\\batch_files\\")
        self.xls_dir            = self.cfg.get('xls_dir', f"{self.sys_dir}data\\workbooks\\")
        self.tmpldir            = self.cfg.get('tmpldir', self.get_templates_dir())
        self.tab_seq            = self.cfg.get('tab_seq', '')
        self.ctot               = self.cfg.get('ctot', '')
        self.dirs_dot           = self.cfg.get('dirs_dot', '')
        self.dirs_skip_rel_str  = self.cfg.get('dirs_skip_rel_str', '')
        self.dirs_skip_abs_lst  = self.cfg.get('dirs_skip_abs_lst', '')
        self.bat_pname          = self.cfg.get('bat_pname', '')
        self.xls_pname          = self.cfg.get('xls_pname', '')
        self.cfg_pname          = self.cfg.get('cfg_pname', '')
        self.bat_num            = self.cfg.get('bat_num', 0)
        self.bool_shw_notes     = self.cfg.get('bool_shw_notes', True)
        self.bool_rel_paths     = self.cfg.get('bool_rel_paths', True)
        self.bool_summ_rows     = self.cfg.get('bool_summ_rows', True)
        self.bool_unused_1      = self.cfg.get('bool_unused_1', True)
        self.bool_unused_2      = self.cfg.get('bool_unused_2', True)
        self.bool_unused_3      = self.cfg.get('bool_unused_3', True)

    def validate_vault_id(self, vault_id):
        """Validate Obsidian Vault ID"""
        if not vault_id or not vault_id.strip():
            return False, "Vault ID cannot be empty"

        # Basic validation - should be alphanumeric with some special chars
        if len(vault_id.strip()) < 1:
            return False, "Vault ID must be at least 1 character"

        return True, ""

    def validate_vault_path(self, vault_path):
        """Validate Obsidian Vault File Path"""
        if not vault_path or not vault_path.strip():
            return False, "Vault path cannot be empty"

        path = Path(vault_path.strip())
        if not path.exists():
            return False, "Vault path does not exist"

        if not path.is_dir():
            return False, "Vault path must be a directory"

        return True, ""

    def validate_dirs_skip_rel_str(self, dirs_skip_rel_str, vault_path):
        """Validate directories to ignore"""
        if not dirs_skip_rel_str or not dirs_skip_rel_str.strip():
            return True, ""  # Empty is valid

        if not vault_path or not vault_path.strip():
            return False, "Vault path must be set first"

        vault_path_obj = Path(vault_path.strip())
        if not vault_path_obj.exists():
            return False, "Vault path must be valid first"

        dirs = [d.strip() for d in dirs_skip_rel_str.split(',') if d.strip()]
        if not dirs:
            return True, ""

        for dir_name in dirs:
            # Check if directory exists anywhere under vault_path
            found = False
            for root, dirs_list, files in os.walk(vault_path_obj):
                if dir_name in dirs_list:
                    found = True
                    break

            if not found:
                return False, f"Directory '{dir_name}' not found under vault path"

        return True, ""

    def validate_wb_exec_path(self, wb_exec_path):
        """Validate spreadsheet executable path"""
        if not wb_exec_path or not wb_exec_path.strip():
            return False, "Executable path cannot be empty"

        path = Path(wb_exec_path.strip())
        if not path.exists():
            return False, "Executable file does not exist"

        if not path.is_file():
            return False, "Executable path must be a file"

        # Check if file is executable
        if not os.access(path, os.X_OK):
            return False, "File is not executable"

        return True, ""

    def browse_vault_path(self):
        """Open directory browser for vault path"""
        folder_path = filedialog.askdirectory(
            title="Select Obsidian Vault Directory",
            initialdir=self.vault_path if self.vault_path else "/"
        )
        if folder_path:
            self.vault_path_var.set(folder_path)
            self.validate_all_fields()

    def browse_exec_path(self):
        """Open file browser for executable path"""
        file_path = filedialog.askopenfilename(
            title="Select Spreadsheet Executable",
            initialdir=os.path.dirname(self.wb_exec_path) if self.wb_exec_path else "/",
            filetypes=[
                ("Executable files", "*.exe" if platform.system() == "Windows" else "*"),
                ("All files", "*.*")
            ]
        )
        if file_path:
            self.wb_exec_path_var.set(file_path)
            self.validate_all_fields()

    def validate_all_fields(self):
        """Validate all fields and update save button state"""
        # vault_id_valid, vault_id_msg = self.validate_vault_id(self.vault_id_var.get())
        vault_path_valid, vault_path_msg = self.validate_vault_path(self.vault_path_var.get())
        wb_exec_valid, wb_exec_msg = self.validate_wb_exec_path(self.wb_exec_path_var.get())
        dirs_skip_rel_str_valid, dirs_skip_rel_str_msg = self.validate_dirs_skip_rel_str(
            self.dirs_skip_rel_str_var.get(), self.vault_path_var.get()
        )

        # Update status labels
        # self.vault_id_status.config(
        #     text=vault_id_msg if not vault_id_valid else "✓ Valid",
        #     foreground="red" if not vault_id_valid else "green"
        # )
        self.vault_path_status.config(
            text=vault_path_msg if not vault_path_valid else "✓ Valid",
            foreground="red" if not vault_path_valid else "green"
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
        all_valid = vault_path_valid and wb_exec_valid and dirs_skip_rel_str_valid
        self.save_button.config(state="normal" if all_valid else "disabled")

        return all_valid

    def on_save_and_run(self):
        """Handle Save & Run button click"""
        if self.validate_all_fields():
            # self.vault_id = self.vault_id_var.get().strip()
            self.vault_path = self.vault_path_var.get().strip()
            self.wb_exec_path = self.wb_exec_path_var.get().strip()
            self.dirs_skip_rel_str = self.dirs_skip_rel_str_var.get().strip()
            self.bool_shw_notes = self.bool_shw_notes_var.get()
            self.bool_rel_paths = self.bool_rel_paths_var.get()
            self.bool_summ_rows = self.bool_summ_rows_var.get()
            self.bool_unused_1  = self.bool_unused_1_var.get()
            self.bool_unused_2  = self.bool_unused_2_var.get()
            self.bool_unused_3  = self.bool_unused_3_var.get()

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
        self.vault_path_var         = tk.StringVar(value=self.vault_path)
        self.wb_exec_path_var       = tk.StringVar(value=self.wb_exec_path)
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
        vault_path_frame = ttk.Frame(main_frame)
        vault_path_frame.grid(row=row, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
        vault_path_frame.columnconfigure(0, weight=1)

        vault_path_entry = ttk.Entry(vault_path_frame, textvariable=self.vault_path_var)
        vault_path_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        ttk.Button(vault_path_frame, text="Browse", command=self.browse_vault_path).grid(row=0, column=1)

        row += 1
        self.vault_path_status = ttk.Label(main_frame, text="", foreground="red")
        self.vault_path_status.grid(row=row, column=1, sticky=tk.W, padx=(10, 0))

        row += 1

        # Executable Path field
        ttk.Label(main_frame, text="Spreadsheet Executable:").grid(row=row, column=0, sticky=tk.W, pady=(15, 5))


        wb_exec_frame = ttk.Frame(main_frame)
        wb_exec_frame.grid(row=row, column=1, sticky=(tk.W, tk.E), pady=(15, 5), padx=(10, 0))
        wb_exec_frame.columnconfigure(0, weight=1)

        wb_exec_entry = ttk.Entry(wb_exec_frame, textvariable=self.wb_exec_path_var)
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
        ttk.Checkbutton(options_frame, text="For Future Use-1", variable=self.bool_unused_1_var).grid(
            row=0, column=1, sticky=tk.W, pady=5
        )

        # Second row of checkboxes
        ttk.Checkbutton(options_frame, text="Use Full Paths in Links", variable=self.bool_rel_paths_var).grid(
            row=1, column=0, sticky=tk.W, pady=5
        )
        ttk.Checkbutton(options_frame, text="For Future Use-2", variable=self.bool_unused_2_var).grid(
            row=1, column=1, sticky=tk.W, pady=5
        )

        # Third row of checkboxes
        ttk.Checkbutton(options_frame, text="Show Rows in Summary", variable=self.bool_summ_rows_var).grid(
            row=2, column=0, sticky=tk.W, pady=5
        )
        ttk.Checkbutton(options_frame, text="For Future Use-3", variable=self.bool_unused_3_var).grid(
            row=2, column=1, sticky=tk.W, pady=5
        )

        row += 1

        # Buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=row, column=0, columnspan=2, pady=(30, 0))

        self.save_button = ttk.Button(button_frame, text="Save & Run", command=self.on_save_and_run)
        self.save_button.pack(side=tk.LEFT, padx=(0, 10))

        cancel_button = ttk.Button(button_frame, text="Cancel", command=self.on_cancel)
        cancel_button.pack(side=tk.LEFT)

        # Bind validation to field changes
        self.vault_path_var.trace('w', lambda *args: self.validate_all_fields())
        self.wb_exec_path_var.trace('w', lambda *args: self.validate_all_fields())
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
    # print(f"Vault Path: {config['vault_path']}")
    # print(f"Executable Path: {config['wb_exec_path']}")
    # print(f"Ignore Directories: {config['dirs_skip_rel_str']}")
    # print(f"Show Notes: {config['show_notes']}")
    # print(f"Use Full Paths: {config['use_full_paths']}")
    # print(f"Show Rows Summary: {config['show_rows_summary']}")
    # print(f"Future Use 1: {config['future_use_1']}")
    # print(f"Future Use 2: {config['future_use_2']}")
    # print(f"Future Use 3: {config['future_use_3']}")