import os
import yaml
from tkinter import messagebox
from pathlib import Path
from datetime import datetime
import platform
import json
from dataclasses import dataclass, field

import src.v_chk_class_lib
from src.v_chk_class_lib import ObsidianApp
from src.v_chk_setupscreen import SetupScreen

@dataclass
class SysConfig:
    sys_cfg:                     dict = field(default_factory=dict)
    sys_id:                  str = 'v_chk'
    sys_ver:                 str = '0.2.9'
    sys_dir_sys:             str = f"{Path(__name__).parent.parent.absolute()}/"
    sys_dir_data:            str = f"{Path(__name__).parent.parent.absolute()}/data/"
    sys_dir_batch:           str = f"{Path(__name__).parent.parent.absolute()}/data/batch_files/"
    sys_dir_wbs:             str = f"{Path(__name__).parent.parent.absolute()}/data/workbooks/"
    sys_pn_cfg:              str = f"{Path(__name__).parent.parent.absolute()}/CONFIG.yaml"
    sys_pn_batch:            str = ''
    sys_pn_wbs:              str = ''
    sys_pn_wb_exec:          str = ''
    sys_obs_vaults:          dict = field(default_factory=dict)
    sys_tab_seq:             tuple = ('pros', 'vals', 'tags', 'file',
                                      'code', 'xyml', 'dups', 'tmpl',
                                      'nest', 'plug', 'summ', 'ar51')
    sys_cfg_os:              str = platform.system()
    vault_name:              str = ''
    vault_id:                str = ''
    dir_vault:               str = ''
    dir_templates:           str = ''
    dirs_skip_rel_str:       str = ''
    dirs_skip_abs_lst:       list = field(default_factory=list)
    dirs_dot:                list = field(default_factory=list)
    ctot:                    list = field(default_factory=list)
    bool_shw_notes:          bool = True
    bool_rel_paths:          bool = True
    bool_summ_rows:          bool = True
    bool_unused_1:           bool = False
    bool_unused_2:           bool = False
    bool_unused_3:           bool = False
    link_lim_vals:           int = 0
    link_lim_tags:           int = 0
    v_chk_date:              str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def __post_init__(self):
        # make sure the necessary directories exist
        self.o_app = ObsidianApp()
        self.vault_name = self.o_app.dflt_vault_name
        self.vault_id   = self.o_app.cur_obs_vaults[self.vault_name][0]
        self.dir_vault  = self.o_app.cur_obs_vaults[self.vault_name][1]
        self.pn_wb_exec = self.set_dflt_wb_exec(self.sys_cfg_os)
        self.obs_vaults = self.o_app.obs_vaults
        self.cfg = {}
        pname = Path(self.sys_dir).joinpath(f"CONFIG.yaml")
        self.pn_cfg = f"{pname}"
        data_dir = f"{self.sys_dir}data"

        self.chk_dirs([self.sys_dir_data, self.sys_dir_batch, self.sys_dir_wbs])

        self.load_config(self.sys_pn_cfg)

    def chk_dirs(self, path_lst: str or list) -> None:
        """
            Checks if the given directory or list of directories exists, and creates them if they do not.
            :param path_lst: A single path as a string or a list of directory paths.
            :return: None
        """
        if isinstance(path_lst, list):
            for p in path_lst:
                self.chk_dirs(p)
            return
        if os.path.isdir(path_lst):
            return
        os.makedirs(path_lst)

    def set_dflt_wb_exec(self, cfg_os):
        # Todo - Needs testing on all platforms
        wb_exec = ""
        common_execs = {
              'Linux': ['scalc']
            , 'Darwin': ['open -a Numbers.app ']
            , 'Windows': ['C:/Program Files/Microsoft Office/root/Office16/EXCEL.EXE', 'C:/Program Files/LibreOffice/program/scalc.exe']
         }

        for wb_app in common_execs[cfg_os]:
            if cfg_os != 'Windows':
                wb_exec = wb_app
                break
            elif Path(wb_app).exists():
                wb_exec = wb_app
                break

        return wb_exec

    def chk_fields_on_load(self) -> bool:
        dir_vault_valid, _ = self.validate_dir_vault(self.dir_vault)
        wb_exec_valid, _ = self.validate_sys_pn_wb_exec(self.sys_pn_wb_exec)
        return dir_vault_valid and wb_exec_valid

    def get_templates_dir(self) -> str or None:
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

    def read_config(self, pn_file: str) -> dict:
        cfg_data = {}
        try:
            with open(pn_file, 'r') as file:
                cfg_data = yaml.safe_load(file)
        except FileNotFoundError:
                pass
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read {pn_file}: {str(e)}")
            raise Exception(f"Failed to read {pn_file}: {str(e)}")

        cfg_data[self.vault_name] = self.sys_cfg
        self.write_config(pn_file, cfg_data)

        return cfg_data

    def load_config(self, pn_file:str) -> None:
        self.cfg = self.read_config(pn_file)
        self.cfg_unpack()

    def write_config(self, pn_file, cfg_data):
        try:
            with open(pn_file, 'w') as file:
                yaml.dump(cfg_data, file, default_flow_style=False)
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save {pn_file}: {str(e)}")
            return False

    def save_config(self, pn_file):
        self.cfg_pack()
        self.write_config(self.sys_pn_cfg, self.cfg)

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
              'sys_cfg':            self.sys_cfg
            , 'sys_id':             self.sys_id
            , 'sys_ver':            self.sys_ver
            , 'sys_dir_sys':        self.sys_dir_sys
            , 'sys_dir_data':       self.sys_dir_data
            , 'sys_dir_batch':      self.sys_dir_batch
            , 'sys_dir_wbs':        self.sys_dir_wbs
            , 'sys_pn_cfg':         self.sys_pn_cfg
            , 'sys_pn_batch':       self.sys_pn_batch
            , 'sys_pn_wbs':         self.sys_pn_wbs
            , 'sys_tab_seq':        self.sys_tab_seq
            , 'sys_cfg_os':         self.sys_cfg_os
            , 'sys_obs_vaults':     self.sys_obs_vaults
            , 'vlt_cfg':            self.vlt_cfg
            , 'vault_name':         self.vault_name
            , 'vault_id':           self.vault_id
            , 'dir_vault':          self.dir_vault
            , 'dir_templates':      self.dir_templates
            , 'dirs_skip_rel_str':  self.dirs_skip_rel_str
            , 'dirs_skip_abs_lst':  self.dirs_skip_abs_lst
            , 'dirs_dot':           self.dirs_dot
            , 'ctot':               self.ctot
            , 'bool_shw_notes':     self.bool_shw_notes
            , 'bool_rel_paths':     self.bool_rel_paths
            , 'bool_summ_rows':     self.bool_summ_rows
            , 'bool_unused_1':      self.bool_unused_1
            , 'bool_unused_2':      self.bool_unused_2
            , 'bool_unused_3':      self.bool_unused_3
            , 'link_lim_vals':      self.link_lim_vals
            , 'link_lim_tags':      self.link_lim_tags
            , 'v_chk_date':         self.v_chk_date
        }

    def cfg_unpack(self):
        self.cfg                = self.cfg.get('cfg', {})
        self.sys_id             = self.cfg.get('sys_id', 'v_chk')
        self.sys_ver            = self.cfg.get('sys_ver', '0.2.9')
        self.sys_dir_sys        = self.cfg.get('sys_dir_sys', f"{Path(__name__).parent.parent.absolute()}/")
        self.sys_dir_data       = self.cfg.get('sys_dir_data', f"{Path(__name__).parent.parent.absolute()}/data/")
        self.sys_dir_batch      = self.cfg.get('sys_dir_batch', f"{Path(__name__).parent.parent.absolute()}/data/batch_files/")
        self.sys_dir_wbs        = self.cfg.get('sys_dir_wbs', f"{Path(__name__).parent.parent.absolute()}/data/workbooks/")
        self.sys_pn_cfg         = self.cfg.get('sys_pn_cfg', f"{Path(__name__).parent.parent.absolute()}/CONFIG.yaml")
        self.sys_pn_batch       = self.cfg.get('sys_pn_batch', '')
        self.sys_pn_wbs         = self.cfg.get('sys_pn_wbs', '')
        self.sys_tab_seq        = self.cfg.get('sys_tab_seq', ('pros', 'vals', 'tags', 'file',
                                                               'code', 'xyml', 'dups', 'tmpl',
                                                               'nest', 'plug', 'summ', 'ar51'))
        self.sys_cfg_os         = self.cfg.get('sys_cfg_os', platform.system())
        self.sys_obs_vaults     = self.cfg.get('sys_obs_vaults', {})
        self.vlt_cfg            = self.cfg.get('vlt_cfg', {})
        self.vault_name         = self.cfg.get('vault_name', '')
        self.vault_id           = self.cfg.get('vault_id', '')
        self.dir_vault          = self.cfg.get('dir_vault', '')
        self.sys_pn_batch       = self.cfg.get('sys_pn_batch', '')
        self.sys_pn_wbs         = self.cfg.get('sys_pn_wbs', '')
        self.dir_templates      = self.cfg.get('dir_templates', self.get_templates_dir())
        self.dirs_skip_rel_str  = self.cfg.get('dirs_skip_rel_str', '')
        self.dirs_skip_abs_lst  = self.cfg.get('dirs_skip_abs_lst', [])
        self.dirs_dot           = self.cfg.get('dirs_dot', [])
        self.ctot               = self.cfg.get('ctot', [0] * 13)
        self.bool_shw_notes     = self.cfg.get('bool_shw_notes', True)
        self.bool_rel_paths     = self.cfg.get('bool_rel_paths', True)
        self.bool_summ_rows     = self.cfg.get('bool_summ_rows', True)
        self.bool_unused_1      = self.cfg.get('bool_unused_1', False)
        self.bool_unused_2      = self.cfg.get('bool_unused_2', False)
        self.bool_unused_3      = self.cfg.get('bool_unused_3', False)
        self.link_lim_vals      = self.cfg.get('link_lim_vals', 0)
        self.link_lim_tags      = self.cfg.get('link_lim_tags', 0)
        self.v_chk_date             = self.cfg.get('v_chk_date', datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

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
    def validate_sys_pn_wb_exec(sys_pn_wb_exec):
        if not sys_pn_wb_exec or not sys_pn_wb_exec.strip():
            return False, "Executable path cannot be empty"
        path = Path(sys_pn_wb_exec.strip())
        if not path.exists():
            return False, "Executable file does not exist"
        if not path.is_file():
            return False, "Executable path must be a file"
        if not os.access(path, os.X_OK):
            return False, "File is not executable"
        return True, ""

def main() -> None:
    sys_cfg = SysConfig()
    setup_screen = SetupScreen(sys_cfg).show()

if __name__ == '__main__':
    main()
