import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk
from datetime import datetime
import subprocess
import copy

class SetupScreen:
    def __init__(self, sys_obj):
        self.sys_obj = sys_obj
        self.sys_cfg = self.sys_obj.sys_cfg

        self.c_vlts  = self.sys_obj.cur_vlts  # make pointer, for easier reference
        self.v_list  = list(self.c_vlts.keys())
        self.last_vault_name   = self.sys_obj.vault_name

        self.root = tk.Tk()
        self.root.title("Obsidian Vault Health Check")
        self.root.geometry("720x540")
        self.root.resizable(True, False)
        self.root.attributes('-topmost', 1)
        self.root.iconbitmap('../img/swenlogo.ico')
        self.logo_image = Image.open('../img/SwenLogo2.png').resize((200, 200))
        self.frame_image = ImageTk.PhotoImage(self.logo_image, master=self.root)

        # Tkinter variables
        self.vault_name_var        = tk.StringVar(value=self.sys_obj.vault_name)
        self.sys_pn_wb_exec_var    = tk.StringVar(value=self.sys_obj.sys_pn_wb_exec)
        self.skip_rel_str_var      = tk.StringVar(value=self.sys_obj.skip_rel_str)
        self.bool_shw_notes_var    = tk.BooleanVar(value=self.sys_obj.bool_shw_notes)
        self.bool_rel_paths_var    = tk.BooleanVar(value=self.sys_obj.bool_rel_paths)
        self.bool_summ_rows_var    = tk.BooleanVar(value=self.sys_obj.bool_summ_rows)
        self.bool_unused_1_var     = tk.BooleanVar(value=self.sys_obj.bool_unused_1)
        self.bool_unused_2_var     = tk.BooleanVar(value=self.sys_obj.bool_unused_2)
        self.bool_unused_3_var     = tk.BooleanVar(value=self.sys_obj.bool_unused_3)
        self.link_lim_vals_var     = tk.StringVar(value=str(self.sys_obj.link_lim_vals))
        self.link_lim_tags_var     = tk.StringVar(value=str(self.sys_obj.link_lim_tags))

        # self.vault_name_status = None
        self.wb_exec_status = None
        self.skip_rel_str_status = None
        self.skip_rel_str_valid = None
        self.skip_rel_str_msg = 'X'
        self.link_lim_vals_label = None
        self.link_lim_tags_label = None
        self.link_lim_vals_help = None
        self.link_lim_tags_help = None
        self.save_button = None
        self.wb_col_max = 16300
        self.wb_col_help = f"0=Unlimited"

    # End of __init__ ==========================================================================================
    def upd_all_sys_objs_with_tk_vars(self, vk: str) -> None:
        """
        Get "_vars" values from tkinter variables and store in sys_obj and cur_vaults
        :return: None
        """
        self.c_vlts[vk]['skip_rel_str']      = self.sys_obj.skip_rel_str      = self.skip_rel_str_var.get().strip()
        self.c_vlts[vk]['bool_shw_notes']    = self.sys_obj.bool_shw_notes    = self.bool_shw_notes_var.get()
        self.c_vlts[vk]['bool_rel_paths']    = self.sys_obj.bool_rel_paths    = self.bool_rel_paths_var.get()
        self.c_vlts[vk]['bool_summ_rows']    = self.sys_obj.bool_summ_rows    = self.bool_summ_rows_var.get()
        self.c_vlts[vk]['bool_unused_1']     = self.sys_obj.bool_unused_1     = self.bool_unused_1_var.get()
        self.c_vlts[vk]['bool_unused_2']     = self.sys_obj.bool_unused_2     = self.bool_unused_2_var.get()
        self.c_vlts[vk]['bool_unused_3']     = self.sys_obj.bool_unused_3     = self.bool_unused_3_var.get()
        self.c_vlts[vk]['link_lim_vals']     = self.sys_obj.link_lim_vals     = int(self.link_lim_vals_var.get())
        self.c_vlts[vk]['link_lim_tags']     = self.sys_obj.link_lim_tags     = int(self.link_lim_tags_var.get())

        self.sys_obj.sys_pn_wb_exec     = self.sys_pn_wb_exec_var.get().strip()
        self.sys_obj.vault_id           = self.c_vlts[vk]['vault_id']
        self.sys_obj.dir_vault          = self.c_vlts[vk]['dir_vault']

    def upd_sys_objs_with_vaults(self, vk: str) -> None:
        self.sys_obj.vault_id           = self.c_vlts[vk]['vault_id']
        self.sys_obj.dir_vault          = self.c_vlts[vk]['dir_vault']

        self.sys_obj.vault_name         = self.c_vlts[vk]['vault_name']
        self.sys_obj.skip_rel_str       = self.c_vlts[vk]['skip_rel_str']
        self.sys_obj.bool_shw_notes     = self.c_vlts[vk]['bool_shw_notes']
        self.sys_obj.bool_rel_paths     = self.c_vlts[vk]['bool_rel_paths']
        self.sys_obj.bool_summ_rows     = self.c_vlts[vk]['bool_summ_rows']
        self.sys_obj.bool_unused_1      = self.c_vlts[vk]['bool_unused_1']
        self.sys_obj.bool_unused_2      = self.c_vlts[vk]['bool_unused_2']
        self.sys_obj.bool_unused_3      = self.c_vlts[vk]['bool_unused_3']
        self.sys_obj.link_lim_vals      = self.c_vlts[vk]['link_lim_vals']
        self.sys_obj.link_lim_tags      = self.c_vlts[vk]['link_lim_tags']

    def upd_tk_vars_with_sys_obj(self) -> None:
        """
        Set tk "vars" variables for tkinter from sys_obj
        :return: None
        """
        self.vault_name_var         = tk.StringVar(value=self.sys_obj.vault_name)
        self.skip_rel_str_var       = tk.StringVar(value=self.sys_obj.skip_rel_str)
        self.bool_shw_notes_var     = tk.BooleanVar(value=self.sys_obj.bool_shw_notes)
        self.bool_rel_paths_var     = tk.BooleanVar(value=self.sys_obj.bool_rel_paths)
        self.bool_summ_rows_var     = tk.BooleanVar(value=self.sys_obj.bool_summ_rows)
        self.bool_unused_1_var      = tk.BooleanVar(value=self.sys_obj.bool_unused_1)
        self.bool_unused_2_var      = tk.BooleanVar(value=self.sys_obj.bool_unused_2)
        self.bool_unused_3_var      = tk.BooleanVar(value=self.sys_obj.bool_unused_3)
        self.link_lim_vals_var      = tk.StringVar(value=str(self.sys_obj.link_lim_vals))
        self.link_lim_tags_var      = tk.StringVar(value=str(self.sys_obj.link_lim_tags))

    def show(self) -> None:
        def debug_print(*args) -> None:
            print(*args)
            print(
                f"last_vault_name      {self.last_vault_name}\n"
              , f"vault_name           {self.sys_obj.vault_name}\n\n"
              , f"skip_rel_str         {self.sys_obj.skip_rel_str}\n"
              , f"bool_shw_notes       {self.sys_obj.bool_shw_notes}\n"
              , f"bool_rel_paths       {self.sys_obj.bool_rel_paths}\n"
              , f"bool_summ_rows       {self.sys_obj.bool_summ_rows}\n"
              , f"bool_unused_1        {self.sys_obj.bool_unused_1}\n"
              , f"bool_unused_2        {self.sys_obj.bool_unused_2}\n"
              , f"bool_unused_3        {self.sys_obj.bool_unused_3}\n"
              , f"link_lim_vals        {self.sys_obj.link_lim_vals}\n"
              , f"link_lim_tags        {self.sys_obj.link_lim_tags}\n\n"
              , f"vault_id             {self.sys_obj.vault_id}\n"
              , f"dir_vault            {self.sys_obj.dir_vault}\n"
              , f"sys_pn_wb_exec       {self.sys_obj.sys_pn_wb_exec}\n"
              , f"------------------------------------------\n"
            )

        # noinspection PyUnusedLocal
        def vault_name_changed(*args) -> None:
            """
            handle the vault_name combobox changed event
             At this point, the only thing that has changed is the vault_name, so we start swapping...

             First, update cur_vaults (using the last_vault_name) w/the tk (screen) vars
             Next, update sys_obj w/cur_vaults (using the newly selected vault_name)
                   (NB: I know this overwrites sys_obj updates from step one, but that's ok)
             Finally, update the tk (screen) vars w/sys_objs

             NB: Step one will need to be re-done (using the current vault_name) at Save&Run
            """

            print("=========================================================== vault_name event")
            print("pre-swaps state")
            debug_print()

            print("pre-step1 tk         -> cur_vaults")
            self.upd_all_sys_objs_with_tk_vars(self.last_vault_name)
            debug_print()

            print("pre-step2 cur_vaults -> sys_obj")
            self.sys_obj.vault_name = self.vault_name_var.get().strip()
            self.upd_sys_objs_with_vaults(self.sys_obj.vault_name)
            debug_print()

            print("pre-step3 sys_objs   -> tk")
            self.upd_tk_vars_with_sys_obj()
            debug_print()

            self.last_vault_name = self.sys_obj.vault_name

            print("------------------------------------------------------\n\n")
            combx_vault_name.configure(textvariable=self.vault_name_var)
            chekbx_notes.configure(variable=self.bool_shw_notes_var)
            chekbx_fullp.configure(variable=self.bool_rel_paths_var)

            entry_skip_rel_str.configure(textvariable=self.skip_rel_str_var)
            self.skip_rel_str_status.config(
                text=self.skip_rel_str_msg if not self.skip_rel_str_valid else
                "✓" if self.skip_rel_str_var.get().strip() else "",
                foreground="red" if not self.skip_rel_str_valid else "green"
            )

            spinbx_vals.configure(textvariable=self.link_lim_vals_var)
            self.link_lim_vals_help = ttk.Label(lnks_frame,
                                                text="(Unlimited)    " if self.sys_obj.link_lim_vals == 0 else self.wb_col_help)
            self.link_lim_vals_help.grid(row=0, column=1, sticky="w", pady=5, padx=(80, 0))

            spinbx_tags.configure(textvariable=self.link_lim_tags_var)
            self.link_lim_tags_help = ttk.Label(lnks_frame,
                                                text="(Unlimited)    " if self.sys_obj.link_lim_tags == 0 else self.wb_col_help)
            self.link_lim_tags_help.grid(row=1, column=1, sticky="w", pady=5, padx=(80, 0))

            self.link_lim_vals_var.trace('w', update_links_help)
            self.link_lim_tags_var.trace('w', update_links_help)

            # Bind validation
            combx_vault_name.bind('<<ComboboxSelected>>', lambda event: vault_name_changed())
            self.skip_rel_str_var.trace('w', lambda *args: self.validate_all_fields())
            self.sys_pn_wb_exec_var.trace('w', lambda *args: self.validate_all_fields())
            self.validate_all_fields()
            update_links_help()
            print("------------------------------------------------------\n\n")

        # noinspection PyUnusedLocal
        def update_links_help(*args) -> None:
            try:
                vals = int(self.link_lim_vals_var.get())
                self.link_lim_vals_help.config(text="(Unlimited)    " if vals == 0 else self.wb_col_help)
            except ValueError:
                self.link_lim_vals_help.config(text=" Invalid!!!")
            try:
                tags = int(self.link_lim_tags_var.get())
                self.link_lim_tags_help.config(text="(Unlimited)    " if tags == 0 else self.wb_col_help)
            except ValueError:
                self.link_lim_tags_help.config(text=" Invalid!!!")

# ============ end of show() function defs - on with the show! ===================================================
        # Main App Frame ---------------------------------------------------------------------
        main_frame = ttk.Frame(self.root, padding="1", borderwidth=1, relief="ridge")
        main_frame.pack(fill="both", expand=True)
        main_frame.columnconfigure(0, weight=1)
        # main_frame.columnconfigure(1, weight=2)
        mf_row = 0
        f1_1st_col = 0

        # Obsidian Frame ---------------------------------------------------------------------
        f1_row = 0  # f1 denotes frame nesting level one; re-used for each frame
        f1_col = f1_1st_col

        # Obsidian Vault Details Frame ---------------------------------------------------------------------
        obs_frame = ttk.LabelFrame(main_frame, text="Obsidian Vault Details ",
                                   padding="20", borderwidth=1, relief="ridge")
        obs_frame.grid(row=mf_row, column=0, sticky="nsew", pady=5, padx=(10, 10))
        obs_frame.columnconfigure(1, weight=1)

        # label
        ttk.Label(obs_frame, text="Vault Name:", width=15).grid(row=0, column=0, sticky="w", padx=5, pady=5)

        # vault name entry
        f1_col += 1
        combx_vault_name = ttk.Combobox(obs_frame, textvariable=self.vault_name_var, width=50)
        combx_vault_name['values'] = self.v_list
        combx_vault_name['state'] = 'readonly'

        # combx_vault_name.current(0)
        combx_vault_name.columnconfigure(f1_col, minsize=30, weight=2)
        combx_vault_name.grid(row=f1_row, column=f1_col, sticky="ew", padx=(0, 5))

        # Ignore Directories (dir_skip_rel)
        # label
        f1_row += 1
        f1_col = f1_1st_col
        ttk.Label(obs_frame, text="Directories to Ignore:\n(comma separated)").grid(row=f1_row,
                            column=0, sticky="w", padx=5, pady=(20, 5))

        # entry Ignore Directories (dir_skip_rel)
        f1_col += 1
        entry_skip_rel_str = ttk.Entry(obs_frame, textvariable=self.skip_rel_str_var, width=50)
        entry_skip_rel_str.columnconfigure((f1_col, f1_col + 1), minsize=30, weight=2)
        entry_skip_rel_str.grid(row=f1_row, column=f1_col, sticky="ew", padx=(0, 5))

        # status Ignore Directories (dir_skip_rel)
        f1_col += 3
        self.skip_rel_str_status = ttk.Label(obs_frame, text="", foreground="red")
        self.skip_rel_str_status.columnconfigure(f1_col, weight=1)
        self.skip_rel_str_status.grid(row=f1_row, column=f1_col, sticky="w", padx=10)

        # Options Frame ---------------------------------------------------------------------
        mf_row = 4
        # opts_frame = ttk.Frame(main_frame)
        # opts_frame.grid(row=mf_row, column=mf_col, sticky="ew", pady=5, padx=(10, 0))

        opts_frame = ttk.LabelFrame(main_frame, text="Workbook Options  ", padding="20", borderwidth=1, relief="ridge")
        opts_frame.grid(row=mf_row, column=0, sticky="nsew", pady=5, padx=(10, 10)) # padx=(0, 0))
        opts_frame.columnconfigure(1, weight=1)
        # opts_frame.columnconfigure(1, weight=1)

        chekbx_notes = ttk.Checkbutton(opts_frame, text="Show Notes", variable=self.bool_shw_notes_var)
        chekbx_notes.grid(row=0, column=0, sticky="w", pady=5)
        ck_open1 = ttk.Checkbutton(opts_frame, text="For Future Use-1",
                                   variable=self.bool_unused_1_var, state='disabled')
        ck_open1.grid(row=0, column=1, sticky="w", pady=5)
        chekbx_fullp = ttk.Checkbutton(opts_frame, text="Use Full Paths in Links", variable=self.bool_rel_paths_var)
        chekbx_fullp.grid(row=1, column=0, sticky="w", pady=5)
        ck_open2 = ttk.Checkbutton(opts_frame, text="For Future Use-2",
                                   variable=self.bool_unused_2_var, state='disabled')
        ck_open2.grid(row=1, column=1, sticky="w", pady=5)

        mf_row += 1

        # Links Frame ---------------------------------------------------------------------
        # Displayed Links Maximums
        lnks_frame = ttk.LabelFrame(main_frame, text="Workbook Link Columns",
                                    padding="20", borderwidth=1, relief="ridge")
        lnks_frame.grid(row=mf_row, column=0, sticky="nsew", pady=5, padx=(10, 10))
        lnks_frame.columnconfigure(1, weight=1)

        # Label
        ttk.Label(lnks_frame, text="Values Tab Maximum Links:").grid(row=0
                                                                     , column=0
                                                                     , sticky="w"
                                                                     , pady=5
                                                                     , padx=(0, 10))

        spinbx_vals = ttk.Spinbox(lnks_frame, from_=0, to=self.wb_col_max
                                   , textvariable=self.link_lim_vals_var, width=8)
        spinbx_vals.grid(row=0, column=1, sticky="w", pady=5)

        self.link_lim_vals_help = ttk.Label(lnks_frame
                                        , text="(Unlimited)    " if self.sys_obj.link_lim_vals == 0 else self.wb_col_help)
        self.link_lim_vals_help.grid(row=0, column=1, sticky="w", pady=5, padx=(80,0))


        ttk.Label(lnks_frame, text="Tags Tab Maximum Links:").grid(row=1
                                                                   , column=0
                                                                   , sticky="w"
                                                                   , pady=5
                                                                   , padx=(0, 10))
        spinbx_tags = ttk.Spinbox(lnks_frame, from_=0, to=self.wb_col_max
                                   , textvariable=self.link_lim_tags_var, width=8)
        spinbx_tags.grid(row=1, column=1, sticky="w", pady=5)

        self.link_lim_tags_help = ttk.Label(lnks_frame
                                        , text="(Unlimited)    " if self.sys_obj.link_lim_tags == 0 else self.wb_col_help)
        self.link_lim_tags_help.grid(row=1, column=1, sticky="w", pady=5, padx=(80,0))

        self.link_lim_vals_var.trace('w', update_links_help)
        self.link_lim_tags_var.trace('w', update_links_help)
        mf_row += 1

        # Executable Path Frame ---------------------------------------------------------------------
        wbex_frame = ttk.LabelFrame(main_frame, text="Workbook Executable ",
                                    padding="20", borderwidth=1, relief="ridge")
        wbex_frame.grid(row=mf_row, column=0, sticky="nsew", pady=5, padx=(10, 10))
        wbex_frame.columnconfigure(1, weight=1)

        # label
        ttk.Label(wbex_frame, text="Full Path:").grid(row=0, column=0, sticky="w", padx=5, pady=5)

        # entry
        wb_exec_entry = ttk.Entry(wbex_frame, textvariable=self.sys_pn_wb_exec_var)
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
        # noinspection PyTypeChecker
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
        combx_vault_name.bind('<<ComboboxSelected>>', lambda event: vault_name_changed())
        self.skip_rel_str_var.trace('w', lambda *args: self.validate_all_fields())
        self.sys_pn_wb_exec_var.trace('w', lambda *args: self.validate_all_fields())
        self.validate_all_fields()

        # Center window
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (self.root.winfo_width() // 2)
        y = (self.root.winfo_screenheight() // 2) - (self.root.winfo_height() // 2)
        self.root.geometry(f"+{x}+{y}")
        self.root.mainloop()

    def browse_exec_path(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Select Spreadsheet Executable",
            initialdir=os.path.dirname(self.sys_pn_wb_exec_var.get()) if self.sys_pn_wb_exec_var.get() else "/",
            filetypes=[
                ("Executable files", "*.exe" if self.sys_obj.sys_cfg_os == "Windows" else "*"),
                ("All files", "*.*")
            ]
        )
        if file_path:
            self.sys_pn_wb_exec_var.set(file_path)
            self.validate_all_fields()

    def validate_all_fields(self) -> None:
        wb_exec_valid, wb_exec_msg = self.sys_obj.validate_sys_pn_wb_exec(self.sys_pn_wb_exec_var.get())

        self.skip_rel_str_valid, self.skip_rel_str_msg = self.sys_obj.validate_skip_rel_str(
                                                      self.skip_rel_str_var.get()
                                                    , self.sys_obj.dir_vault
                                                    )
        self.wb_exec_status.config(
            text=wb_exec_msg if not wb_exec_valid else "✓",
            foreground="red" if not wb_exec_valid else "green"
            )
        self.skip_rel_str_status.config(
            text=self.skip_rel_str_msg if not self.skip_rel_str_valid else
                                        "✓" if self.skip_rel_str_var.get().strip() else "",
            foreground="red" if not self.skip_rel_str_valid else "green"
            )
        all_valid = wb_exec_valid and self.skip_rel_str_valid
        self.save_button.config(state="normal" if all_valid else "disabled")
        return all_valid

    def on_save_and_run(self) -> None:
        if self.validate_all_fields():
            self.upd_all_sys_objs_with_tk_vars(self.sys_obj.vault_name)
            self.sys_obj.v_chk_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            if self.sys_obj.save_config(self.sys_obj.sys_pn_cfg):
                self.root.quit()
                self.root.destroy()
                subprocess.run(["python", "v_chk.py"])
            else:
                messagebox.showerror("Error", "Failed to save configuration")

    def on_cancel(self) -> None:
        self.root.quit()
        self.root.destroy()

def main() -> None:
    print(f'Cannot run this script ("v_chk_setupscreen"), directly. Running "v_chk_setup" instead.')
    # sys.exit(0)
    subprocess.run(["python", "v_chk_setup.py"])

if __name__ == '__main__':
    main()


