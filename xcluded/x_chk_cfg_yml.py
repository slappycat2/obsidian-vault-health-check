
import yaml
import re
from pathlib import Path
from datetime import datetime
import os
import glob
from v_chk_class_lib import Colors

class Config:
    def __init__(self):
        self.c_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.bat_num = 0
        self.chk_yaml = []
        self.dbug = False
        self.vault_path = "E:\\o2"
        self.dirs_dot = [f.name for f in os.scandir(self.vault_path) if
                         f.is_dir() and f.path.startswith(f"{self.vault_path}\\.")]
        self.dirs_skip_rel_str = ["z_meta", "z_resources"]
        self.dirs_skip_abs_lst = self.dirs_dot + self.dirs_skip_rel_str
        self.dup_files = {}
        self.files = {}  # {"filename": {"links":, [list of values]...}
                     # pros {"key":    {"val":  [list of files], "val", [list of
        self.pros = {}  # {"status": {"📝/🟥": [list of files], "📝/🌲️", [list of files],...}, "links": {"[[2025-12-31]]":[list of files],...}...}
        self.Colors = Colors()
        self.tab_clrs = {  #  tab color,         tab hdr colors
              :  [self.Colors.clr_blud4, self.Colors.clr_blud4]
            , 'tags':   [self.Colors.clr_ora64, self.Colors.clr_ora64]
            , 'dups':   [self.Colors.clr_pur45, self.Colors.clr_pur45]
            , 'summ':   [self.Colors.clr_grn30, self.Colors.clr_grn30]
            , 'area51': [self.Colors.clr_red20, self.Colors.clr_red20]
        }

        self.rgx_boundary = re.compile('^---\\s*$', re.MULTILINE)
        self.rgx_body = re.compile('(^|(\\[))([)([A-Za-z0-9_]+)[:]{2}(.*?)(\\]?\\]?)($|\\])')
        self.rgx_noTZdatePattern = r"([0-9]{4})[-\/]([0-1]?[0-9]{1})[-\/]([0-3])?([0-9]{1})(\s+)([0-9]{2}:[0-9]{2}:[0-9]{2})(.*)"
        self.rgx_noTZdateReplace = r"\1-\2-\3\4 \6"
        self.rgx_sub_strip_code_blocks = r'```[\s\S]*?```'
        self.rgx_sub_strip_inline_code = r'`[^`]*`'

        self.v_chk_cfg_pname = ""
        self.v_chk_xls_pname = ""
        self.wb_exec_path = "C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE"

    def get_last_cfg(self):
        """Returns the name of the latest (most recent) file
        of the joined path(s)"""
        path = "/data/batch_files\\"
        pattern = "v_chk_*.yaml"
        fullpath = os.path.join(path, pattern)
        list_of_files = glob.iglob(fullpath)  # You may use iglob in Python3
        if not list_of_files:  # I prefer using the negation
            return None  # because it behaves like a shortcut
        latest_file = max(list_of_files, key=os.path.getctime)
        # dir, filename = os.path.split(latest_file)

        self.v_chk_cfg_pname = latest_file
        print(f"v_chk_cfg: Read Last Config file: {self.v_chk_cfg_pname}")

        return latest_file

    def get_next_cfg(self):
        c_num = 0
        c_file = f"G:\\dev\\v_chk\\batch_files\\v_chk_{c_num:04d}.yaml"

        while Path(c_file).exists():
            c_num += 1
            c_file = f"G:\\dev\\v_chk\\data\batch_files\\v_chk_{c_num:04d}.yaml"

        self.bat_num = c_num
        self.v_chk_cfg_pname = c_file
        self.v_chk_xls_pname = \
            f"G:\\dev\\v_chk\\workbooks\\v_chk_{self.bat_num:04d}.xlsx"

        print(f"v_chk_cfg: Write Next Config file: {self.v_chk_cfg_pname}")
        return 0

    def write_config(self, cfg_obj):
        self.get_next_cfg()
        print(f"pros={self.pros}")

        try:
            with open(self.v_chk_cfg_pname, 'w') as yaml_file:
                # yaml.dump(range(50), width=50, indent=4)
                yaml.dump({
                    'c_date':       self.c_date,
                    'bat_num':      self.bat_num,
                    'dbug':         self.dbug,
                    'v_chk_cfg_pname': self.v_chk_cfg_pname,
                    'v_chk_xls_pname': self.v_chk_xls_pname,
                    'vault_path':   self.vault_path,
                    'dirs_dot':     self.dirs_dot,
                    'dirs_skip_rel_str': self.dirs_skip_rel_str,
                    'dirs_skip_abs_lst':    self.dirs_skip_abs_lst,
                    'wb_exec_path': self.wb_exec_path,
                    :        self.pros,
                    'files':        self.files,
                    'chk_yaml':     self.chk_yaml,
                    'xkey_dup_files':    self.dup_files,
                }
                    , stream=yaml_file
                )
            return 0
        except Exception as e:
            print(f"v_chk_cfg: Error in Save Config: {e}")
            return 1

    def read_config(self, cfg):
        self.v_chk_cfg_pname = self.get_last_cfg()
        if self.v_chk_cfg_pname is None:
            print("v_chk_cfg: Error in Read Config- No Config File Found")
            return False
        else:
            print(f"v_chk_cfg-read_config: Reading Config file: {self.v_chk_cfg_pname}")
        try:
            with open(self.v_chk_cfg_pname, 'r') as file_y:
                cfg_data = file_y.read()

            config_data = yaml.safe_load(cfg_data)
            self.c_date         = config_data.get('c_date', "")
            self.bat_num        = config_data.get('bat_num', {})
            self.dbug           = config_data.get('dbug', False)
            self.v_chk_cfg_pname = config_data.get('v_chk_cfg_pname', {})
            self.v_chk_xls_pname = config_data.get('v_chk_xls_pname', {})
            self.vault_path     = config_data.get('vault_path', {})
            self.dirs_dot = config_data.get('dirs_dot', {})
            self.dirs_skip_rel_str = config_data.get('dirs_skip_rel_str', {})
            self.dirs_skip_abs_lst = config_data.get('dirs_skip_abs_lst', {})
            self.wb_exec_path = config_data.get('wb_exec_path', {})
            self.pros          = config_data.get(, {})
            self.files          = config_data.get('files', {})
            self.chk_yaml       = config_data.get('chk_yaml', [])
            self.dup_files      = config_data.get('xkey_dup_files', [])

        except Exception as e:
            print(f"v_chk_cfg: Error in Read Config: {e}")
            return False

        return cfg

if __name__ == "__main__":
    cfg = Config()
    cfg = cfg.read_config(cfg)

    print(f"cfg.vault_path = {cfg.vault_path}")

    if cfg:
        print(f"cfg.v_chk_cfg_pname: {cfg.v_chk_cfg_pname}")
        print(f"cfg.v_chk_xls_pname: {cfg.v_chk_xls_pname}")
        print(f"cfg.dirs_dot: {cfg.dirs_dot}")
        print(f"cfg.dirs_skip_rel_str: {cfg.dirs_skip_rel_str}")
        print(f"cfg.dirs_skip_abs_lst: {cfg.dirs_skip_abs_lst}")

        print(f"cfg.rgx_boundary: {cfg.rgx_boundary}")
        print(f"cfg.rgx_body: {cfg.rgx_body}")
        print(f"cfg.rgx_noTZdatePattern: {cfg.rgx_noTZdatePattern}")
        print(f"cfg.rgx_noTZdateReplace: {cfg.rgx_noTZdateReplace}")
        print(f"cfg.v_chk_cfg_pname: {cfg.v_chk_cfg_pname}")
        print(f"cfg.v_chk_xls_pname: {cfg.v_chk_xls_pname}")
        print(f"cfg.wb_exec_path: {cfg.wb_exec_path}")
        print(f"cfg.bat_num: {cfg.bat_num}")
        print(f"cfg.dbug: {cfg.dbug}")
        print(f"cfg.vault_path: {cfg.vault_path}")

        print(f"cfg.dup_files: {cfg.dup_files}")
        print(f"cfg.files: {cfg.files}")
        print(f"cfg.pros: {cfg.pros}")

    print(f'Standalone run of "v_chk_cfg_yml.py" complete.')