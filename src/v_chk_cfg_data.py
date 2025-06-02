import glob
import json
from v_chk_setup import *

class DataDefinition:
    """ This is a class that holds the Data Definition (the cell values, themselves)
    for a workbook as opposed to the Workbook Definition, which was supposed to hold the
    Workbook (Headings, filename, etc.) Definition, but it's all pretty much here, now.
    This is a child of WbConfig, and the parent to WbDataDef.
    """
    def __init__(self, sys_id, dbug_lvl=0):
        self.DBUG_LVL = dbug_lvl
        self.cfg = SysConfig(self.DBUG_LVL).cfg
        self.sys_id = sys_id
        self.OPEN_ON_CREATE = True
        self.xls_pname = ""     # assigned below
        self.bat_pname = ""
        self.bat_num = 0

        self.wb_data = {}     # NEED TO CHECK IF OK WITH v_chk
        self.wb_tabs = {}
        self.wb_def = {
              'cfg': self.cfg
            , 'wb_tabs': self.wb_tabs
            , 'wb_data': self.wb_data
        }
        self.tab_def = {}
        self.obs_props = {}
        self.obs_atags = {}
        self.obs_xyaml = {}
        self.obs_nests = {}
        self.obs_dupfn = {}
        self.obs_files = {}
        self.obs_tmplt = {}
        self.obs_datav = {}
        self.obs_plugs = {}
        self.pros = {}
        self.vals = {}
        self.tags = {}
        self.xyml = {}
        self.dups = {}
        self.file = {}
        self.tmpl = {}
        self.code = {}
        self.nest = {}
        self.plug = {}
        self.ar51 = {}
        self.plugin_id_def = {
                  'mapWithTag': 'Metadata Menu'
                , 'kindle-sync': 'Kindle Highlights'
                , 'NestedDictionary': 'Unknown Plugin'
        }
        self.xyml_descs = {
                     # 123456789 1 2345678 2 2345678 3 2345678 4 2345678 5 2345678 6 2345678 7 2345678 8
              'BadY': ["Invalid Properties"            , 'Cannot Load Frontmatter-Check YAML Markdown syntax.']
            , 'NoFm': ['No Properties'                 , 'Not a problem, if intentional.']
            , 'MtFm': ['YAML loaded, but empty' , 'Not a problem, if intentional.']
            , 'ErrY': ["YAML error"             , 'An Unknown Error occurred trying to process Frontmatter.']
            , 'NonD': ['YAML formatting error'  , 'Invalid Frontmatter--Not in dictionary format']
        }

        self.get_last_cfg()     # just so we have the name of the cfg file

    def wb_def_pack(self):
        self.wb_def = {
              'cfg': self.cfg
            , 'wb_tabs': self.wb_tabs
            , 'wb_data': self.wb_data
        }

    def get_last_cfg(self):
        """Returns the name of the latest (most recent) file
        of the path_pattern_lst
        requires: path_pattern_lst - a list formatted as:
        ["G:\\dev\\v_chk\\data\\batch_files\\", "v_chk_*.yaml"]
        """

        latest_file = f"{self.cfg_dir}\\{self.cfg_sys_id}_0000.yaml"
        full_path = f"{self.cfg_dir}\\{self.cfg_sys_id}_????.yaml"
        try:
            list_of_files = glob.iglob(full_path)
            if not list_of_files:
                return
            latest_file = max(list_of_files, key=os.path.getctime)
        except ValueError:
            # no files in dir (latest file is empty)
            pass
        except Exception as e:
            raise Exception(f"ConfigData: Error reading config file ({self.bat_pname}) Error : {e}")

        self.bat_pname = latest_file
        self.xls_pname = f"{self.xls_dir}\\{Path(latest_file).stem}.xlsx"
        if self.DBUG_LVL > 1:
            print(f"ConfigData: Read Last Config file: {self.bat_pname}")
        # Add these to the wb_def (not the sys file!)
        self.cfg['cfg_pname'] = self.bat_pname
        self.cfg['xls_pname'] = self.xls_pname
        self.cfg['cfg_pname'] = self.cfg_pname

        return

    def get_next_cfg(self):
        """Returns the name of the next available yaml config file
        using the path filename stub_provided.
        requires: path_stub formatted as:
        ["G:\\dev\\v_chk\\data\\batch_files\\"]
        """
        self.bat_num = 0
        c_file = f"{self.cfg_dir}{self.cfg_sys_id}_{self.bat_num:04d}.yaml"

        while Path(c_file).exists():
            self.bat_num += 1
            c_file = f"{self.cfg_dir}{self.cfg_sys_id}_{self.bat_num:04d}.yaml"
            if self.DBUG_LVL > 1:
                print(f"ConfigData: Next Config file: {c_file}")

        self.bat_pname = c_file
        self.xls_pname = f"{self.xls_dir}{Path(c_file).stem}.xlsx"
        self.cfg['cfg_pname'] = self.bat_pname
        self.cfg['xls_pname'] = self.xls_pname
        self.cfg['bat_num'] = self.bat_num

        if self.DBUG_LVL > 1:
            print(f"ConfigData: Init Next Config file: {self.bat_pname}")

        # Init everything except cfg, as this is a new file...
        self.tab_def = {}
        self.pros = {'tab_def': self.tab_def}
        self.vals = {'tab_def': self.tab_def}
        self.tags = {'tab_def': self.tab_def}
        self.xyml = {'tab_def': self.tab_def}
        self.dups = {'tab_def': self.tab_def}
        self.file = {'tab_def': self.tab_def}
        self.tmpl = {'tab_def': self.tab_def}
        self.code = {'tab_def': self.tab_def}
        self.nest = {'tab_def': self.tab_def}
        self.plug = {'tab_def': self.tab_def}
        self.summ = {'tab_def': self.tab_def}
        self.ar51 = {'tab_def': self.tab_def}
        self.wb_tabs = {
              'pros': self.pros
            , 'vals': self.tags
            , 'tags': self.tags
            , 'xyml': self.xyml
            , 'dups': self.dups
            , 'file': self.file
            , 'tmpl': self.tmpl
            , 'code': self.code
            , 'nest': self.nest
            , 'plug': self.plug
            , 'summ': self.summ
            , 'ar51': self.ar51
            , 'init': {}
        }
        self.obs_props = {}
        self.obs_atags = {}
        self.obs_files = {}
        self.wb_data = {
              'obs_props': self.obs_props
            , 'obs_atags': self.obs_atags
            , 'obs_files': self.obs_files
        }
        self.wb_def = {
              'cfg': self.cfg
            , 'wb_tabs': self.wb_tabs
            , 'wb_data': self.wb_data
        }

        return

    def write_cfg_data(self):
        if not self.bat_pname:
            self.get_next_cfg()

        # self.wb_def_pack()   # This clobbers wb_data during tags! Upd explicitly!

        try:
            with open(self.bat_pname, 'w') as yaml_file:
                # yaml.dump(range(50), width=50, indent=4)
                yaml.dump({
                    'wb_def':     self.wb_def
                }
                    , stream=yaml_file, sort_keys=False
                )
            return

        except Exception as e:
            print(f"ConfigData-write-wb_def ({self.bat_pname}): Error in Save Config: {e}")
            sys.exit(1)


    def read_cfg_data(self):
        # These comments may be used for the next ver version...
        # For now, we just return a copy of wb_def, and that's it.
        """Accepts a string that may contain C, T and/or D, for Cfg, Tabs & Data.
        If the flag is not set, then the whole wb_def is returned in a temp value.
        This avoids clobbering any existing wb_def.
        If the flag is set, then the wb_def is unpacked into the appropriate"""
        # if op_flg is not None:
        #     op_flg = op_flg.upper()

        if self.bat_pname == '' or self.cfg_pname is None:
            self.get_last_cfg()
            if self.DBUG_LVL > 1:
                print(f"ConfigData-read_config: Loaded last config file: {self.bat_pname}")
        else:
            if self.DBUG_LVL > 1:
                print(f"ConfigData-read_config: Reading Config file: {self.bat_pname}")
        try:
            with open(self.bat_pname, 'r') as file_y:
                cfg_data = file_y.read()

            wb_def_temp = yaml.safe_load(cfg_data)
            wb_def_temp = wb_def_temp.get('wb_def', {})

        except Exception as e:
            raise Exception(f"ConfigData: Error reading config file ({self.bat_pname}) Error : {e}")

        return wb_def_temp

class WbDataDef(DataDefinition):
    # This script, only initializes the environment. It does
    # not return the contents of any yaml files, at this time.
    # Read/write must be done explicitly.
    def __init__(self, dbug_lvl):
        self.DBUG_LVL = dbug_lvl
        super().__init__('v_chk', self.DBUG_LVL)

        pass

    def __getitem__(self, item):
        return self.cfg[item]

if __name__ == "__main__":
    DBUG_LVL = 1
    print(f"\nRunning standalone version of {Path(__file__).name}")

    # Build Tabs
    cfg = WbDataDef(DBUG_LVL)

    # shelve_file = shelve.open("v_def.db")
    # shelve_file['v_def'] = v_def
    # shelve_file.close()

    # self.tab_def['tab_cd_table_hdr']['Row']
    if DBUG_LVL:
        lin = "=" * 30
        dict_list = {
              'cfg': cfg.wb_def['cfg']
            , 'wb_tabs': cfg.wb_def['wb_tabs']
            , 'wb_data': cfg.wb_def['wb_data']
            # , 'tab_def': cfg.wb_def['wb_tabs']['pros']
        }

        for p_dict_name, p_dict in dict_list.items():
            print(f"\n{p_dict_name}: {lin}")
            for k,v in p_dict.items():
                k_name = f"{p_dict_name}['{k}']"
                print(f"{k_name: <20}: {v}")

        print(f'\nStandalone run of "v_chk_cfg_data.py" complete.')