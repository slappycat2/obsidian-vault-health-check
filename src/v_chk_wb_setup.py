import glob
import sys

from src.v_chk_setup import *
from src.v_chk_class_lib import Colors

class WbDataDef:
    def __init__(self, sys_id, dbug_lvl=0):
        self.DBUG_LVL = dbug_lvl
        self.syscfg = SysConfig(self.DBUG_LVL)
        self.cfg = self.syscfg.cfg
        self.cfg_sys_id = self.cfg['cfg_sys_id']
        self.dir_batch    = self.cfg['dir_batch']
        self.dir_wbs    = self.cfg['dir_wbs']
        self.pn_cfg  = self.cfg['pn_cfg']
        self.bat_num    = self.cfg['bat_num']
        self.wb_tabs = {}
        self.wb_data = {}

        self.OPEN_ON_CREATE = True
        self.pn_batch = None
        self.pn_wbs = None
        self.pn_batch = None
        self.pn_wbs = None
        self.tab_def = None
        self.summ = None

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
        self.wb_def = {
              'cfg': self.cfg
            , 'wb_tabs': self.wb_tabs
            , 'wb_data': self.wb_data
        }

        self.get_last_bat()     # just so we have the name of the bat file

    def wb_def_pack(self):
        self.wb_def = {
              'cfg': self.cfg
            , 'wb_tabs': self.wb_tabs
            , 'wb_data': self.wb_data
        }

    def get_last_bat(self):
        """Returns the name of the latest (most recent) file
        of the path_pattern_lst
        requires: path_pattern_lst - a list formatted as:
        ["G:/dev/v_chk/data/batch_files/", "v_chk_*.yaml"]
        """

        latest_file = f"{self.dir_batch}/{self.cfg_sys_id}_0000.yaml"
        full_path = f"{self.dir_batch}/{self.cfg_sys_id}_????.yaml"
        try:
            list_of_files = glob.iglob(full_path)
            if not list_of_files:
                return
            latest_file = max(list_of_files, key=os.path.getctime)
        except ValueError:
            # no files in dir (latest file is empty)
            pass
        except Exception as e:
            raise Exception(f"ConfigData: Error reading config file ({self.pn_batch}) Error : {e}")

        self.pn_batch = latest_file
        self.pn_wbs = f"{self.dir_wbs}/{Path(latest_file).stem}.xlsx"
        if self.DBUG_LVL > 1:
            print(f"ConfigData: Read Last Config file: {self.pn_batch}")
        # Add these to the wb_def (not the sys file!)
        self.cfg['pn_batch'] = self.pn_batch
        self.cfg['pn_wbs'] = self.pn_wbs
        self.cfg['pn_cfg'] = self.pn_cfg

        return

    def get_next_bat(self):
        """Returns the name of the next available yaml config file
        using the path filename stub_provided.
        requires: path_stub formatted as:
        ["G:/dev/v_chk/data/batch_files/"]
        """
        self.bat_num = 0
        c_file = f"{self.dir_batch}{self.cfg_sys_id}_{self.bat_num:04d}.yaml"

        while Path(c_file).exists():
            self.bat_num += 1
            c_file = f"{self.dir_batch}{self.cfg_sys_id}_{self.bat_num:04d}.yaml"
            if self.DBUG_LVL > 1:
                print(f"ConfigData: Next Config file: {c_file}")

        self.pn_batch = c_file
        self.pn_wbs = f"{self.dir_wbs}{Path(c_file).stem}.xlsx"
        self.cfg['pn_batch'] = self.pn_batch
        self.cfg['pn_wbs'] = self.pn_wbs
        self.cfg['bat_num'] = self.bat_num

        if self.DBUG_LVL > 1:
            print(f"ConfigData: Init Next Config file: {self.pn_batch}")

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

    def write_bat_data(self):
        if not self.pn_batch:
            self.get_next_bat()

        # self.wb_def_pack()   # This clobbers wb_data during tags! Upd explicitly!

        try:
            with open(self.pn_batch, 'w') as yaml_file:
                # yaml.dump(range(50), width=50, indent=4)
                yaml.dump({
                    'wb_def':     self.wb_def
                }
                    , stream=yaml_file, sort_keys=False
                )
            return

        except Exception as e:
            print(f"ConfigData-write-wb_def ({self.pn_batch}): Error in Save Config: {e}")
            sys.exit(1)


    def read_bat_data(self):
        # These comments may be used for the next ver version...
        # For now, we just return a copy of wb_def, and that's it.
        """Accepts a string that may contain C, T and/or D, for Cfg, Tabs & Data.
        If the flag is not set, then the whole wb_def is returned in a temp value.
        This avoids clobbering any existing wb_def.
        If the flag is set, then the wb_def is unpacked into the appropriate"""
        # if op_flg is not None:
        #     op_flg = op_flg.upper()

        if self.pn_batch == '' or self.pn_batch is None:
            self.get_last_bat()
            if self.DBUG_LVL > 1:
                print(f"ConfigData-read_config: Loaded last config file: {self.pn_batch}")
        else:
            if self.DBUG_LVL > 1:
                print(f"ConfigData-read_config: Reading Config file: {self.pn_batch}")
        try:
            with open(self.pn_batch, 'r') as file_y:
                bat_data = file_y.read()

            wb_def_temp = yaml.safe_load(bat_data)
            wb_def_temp = wb_def_temp.get('wb_def', {})

        except Exception as e:
            raise Exception(f"ConfigData: Error reading config file ({self.pn_batch}) Error : {e}")

        return wb_def_temp


if __name__ == "__main__":
    DBUG_LVL = 1
    print(f"\nRunning standalone version of {Path(__file__).name}")

    # Build Tabs
    bat = WbDataDef(DBUG_LVL)

    # shelve_file = shelve.open("v_def.db")
    # shelve_file['v_def'] = v_def
    # shelve_file.close()

    # self.tab_def['tab_cd_table_hdr']['Row']
    if DBUG_LVL:
        lin = "=" * 30
        dict_list = {
              'cfg': bat.wb_def['cfg']
            , 'wb_tabs': bat.wb_def['wb_tabs']
            , 'wb_data': bat.wb_def['wb_data']
            # , 'tab_def': bat.wb_def['wb_tabs']['pros']
        }

        for p_dict_name, p_dict in dict_list.items():
            print(f"\n{p_dict_name}: {lin}")
            for k,v in p_dict.items():
                k_name = f"{p_dict_name}['{k}']"
                print(f"{k_name: <20}: {v}")

        print(f'\nStandalone run of "v_chk_wb_setup.py" complete.')