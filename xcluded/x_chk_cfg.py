
import pickle
import re
from pathlib import Path
from datetime import datetime
import os
import glob

class Config:
    def __init__(self):
        self.c_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.bat_num = 0
        self.chk_yaml = []
        self.dbug = False
        self.dir_vault = "E:\\o2"
        self.dirs_dot = [f.name for f in os.scandir(self.dir_vault) if
                         f.is_dir() and f.path.startswith(f"{self.dir_vault}\\.")]
        self.dirs_skip_rel_str = ["z_meta", "z_resources"]
        self.dirs_skip_abs_lst = self.dirs_dot + self.dirs_skip_rel_str
        self.dup_files = {}
        self.files = {}  # {"filename": {"links":, [list of values]...}
                     # pros {"key":    {"val":  [list of files], "val", [list of
        self.pros = {}  # {"status": {"📝/🟥": [list of files], "📝/🌲️", [list of files],...}, "links": {"[[2025-12-31]]":[list of files],...}...}
        self.rgx_boundary = re.compile('^---\\s*$', re.MULTILINE)
        self.rgx_body = re.compile('(^|(\\[))([)([A-Za-z0-9_]+)[:]{2}(.*?)(\\]?\\]?)($|\\])')
        self.rgx_noTZdatePattern = r"([0-9]{4})[-\/]([0-1]?[0-9]{1})[-\/]([0-3])?([0-9]{1})(\s+)([0-9]{2}:[0-9]{2}:[0-9]{2})(.*)"
        self.rgx_noTZdateReplace = r"\1-\2-\3\4 \6"
        self.v_chk_pn_cfg = ""
        self.v_chk_pn_wbs = ""
        self.pn_wb_exec = "C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE"

    def get_last_cfg(self):

        """Returns the name of the latest (most recent) file
        of the joined path(s)"""
        path = "G:\\dev\\v_chk\\batch_files\\"
        pattern = "v_chk_*.pickle"
        fullpath = os.path.join(path, pattern)
        list_of_files = glob.iglob(fullpath)  # You may use iglob in Python3
        if not list_of_files:  # I prefer using the negation
            return None  # because it behaves like a shortcut
        latest_file = max(list_of_files, key=os.path.getctime)
        _, filename = os.path.split(latest_file)

        print(f"v_chk_cfg: Read Last Config file: {self.v_chk_pn_cfg}")

        return latest_file

    def get_next_cfg(self):
        c_num = 0
        c_file = f"G:\\dev\\v_chk\\batch_files\\v_chk_{c_num:04d}.pickle"

        while Path(c_file).exists():
            c_num += 1
            c_file = f"G:\\dev\\v_chk\\batch_files\\v_chk_{c_num:04d}.pickle"

        self.bat_num = c_num
        self.v_chk_pn_cfg = c_file
        self.v_chk_pn_wbs = \
            f"G:\\dev\\v_chk\\workbooks\\v_chk_{self.bat_num:04d}.xlsx"

        print(f"v_chk_cfg: Write Next Config file: {self.v_chk_pn_cfg}")
        return 0

    def write_config(self, cfg_obj):
        self.get_next_cfg()
        cfg_file_name = self.v_chk_pn_cfg
        try:
            with open(cfg_file_name, "wb") as pkl_file:
                pickle.dump(cfg_obj, pkl_file)
                # once include , protocol=pickle.HIGHEST_PROTOCOL after pkl_file, but it failed
        except Exception as ex:
            print("Error during pickling object (Possibly unsupported):", ex)
            return 1
            
        return 0

    def read_config(self):
        self.v_chk_pn_cfg = self.get_last_cfg()
        if self.v_chk_pn_cfg is None:
            print("v_chk_cfg: Error in Read Config- No Config File Found")
            return False
        else:
            print(f"v_chk_cfg-read_config: Reading Config file: {self.v_chk_pn_cfg}")
        try:
            with open(self.v_chk_pn_cfg, "rb") as f:
                cfg = pickle.load(f)

        except Exception as ex:
            print("Error during unpickling object (Possibly unsupported):", ex)
            return False

        return cfg

if __name__ == "__main__":
    cfg = Config()
    cfg = cfg.read_config()
    print(f"cfg.v_chk_pn_cfg: {cfg.v_chk_pn_cfg}")
    print(f"cfg.v_chk_pn_wbs: {cfg.v_chk_pn_wbs}")
    print(f"cfg.dirs_dot: {cfg.dirs_dot}")
    print(f"cfg.dirs_skip_rel_str: {cfg.dirs_skip_rel_str}")
    print(f"cfg.dirs_skip_abs_lst: {cfg.dirs_skip_abs_lst}")

    print(f"cfg.rgx_boundary: {cfg.rgx_boundary}")
    print(f"cfg.rgx_body: {cfg.rgx_body}")
    print(f"cfg.rgx_noTZdatePattern: {cfg.rgx_noTZdatePattern}")
    print(f"cfg.rgx_noTZdateReplace: {cfg.rgx_noTZdateReplace}")
    print(f"cfg.v_chk_pn_cfg: {cfg.v_chk_pn_cfg}")
    print(f"cfg.v_chk_pn_wbs: {cfg.v_chk_pn_wbs}")
    print(f"cfg.pn_wb_exec: {cfg.pn_wb_exec}")
    print(f"cfg.bat_num: {cfg.bat_num}")
    print(f"cfg.dbug: {cfg.dbug}")
    print(f"cfg.dir_vault: {cfg.dir_vault}")

    print(f"cfg.dup_files: {cfg.dup_files}")
    print(f"cfg.files: {cfg.files}")
    print(f"cfg.pros: {cfg.pros}")