import glob
import sys
import os

from pathlib import Path
import yaml

from src.v_chk_setup import SysConfig
# from src.v_chk_class_lib import Colors

from src.v_chk import logger
import v_chk_setup

class WbDataDef:
    def __init__(self):
        self.syscfg_obj = SysConfig()
        self.sys_cfg = self.syscfg_obj.sys_cfg

        self.sys_pn_cfg     = self.sys_cfg['sys_pn_cfg']

        self.sys_id         = self.sys_cfg.get('sys_id','v_chk')
        self.sys_dir_batch  = self.sys_cfg['sys_dir_batch']
        self.sys_dir_wbs    = self.sys_cfg['sys_dir_wbs']
        self.wb_tabs        = {}
        self.wb_data        = {}

        self.OPEN_ON_CREATE = True
        self.sys_pn_batch = None
        self.sys_pn_wbs = None
        self.sys_pn_batch = None
        self.sys_pn_wbs = None
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
              'sys_cfg': self.sys_cfg
            , 'wb_tabs': self.wb_tabs
            , 'wb_data': self.wb_data
        }

        self.get_last_bat()     # just so we have the name of the bat file

    def wb_def_pack(self):
        self.wb_def = {
              'sys_cfg': self.sys_cfg
            , 'wb_tabs': self.wb_tabs
            , 'wb_data': self.wb_data
        }

    def get_last_bat(self):
        """Returns the name of the latest (most recent) file
        of the path_pattern_lst
        requires: path_pattern_lst - a list formatted as:
        ["G:/dev/v_chk/data/batch_files/", "v_chk_*.yaml"]
        """

        latest_file = f"{self.sys_dir_batch}/{self.sys_id}_0000.yaml"
        full_path = f"{self.sys_dir_batch}/{self.sys_id}_????.yaml"
        try:
            list_of_files = glob.iglob(full_path)
            if not list_of_files:
                return
            latest_file = max(list_of_files, key=os.path.getctime)
        except ValueError:
            # no files in dir (latest file is empty)
            pass
        except Exception as e:
            raise Exception(f"ConfigData: Error reading config file ({self.sys_pn_batch}) Error : {e}")

        self.sys_pn_batch = latest_file
        self.sys_pn_wbs = f"{self.sys_dir_wbs}/{Path(latest_file).stem}.xlsx"
        logger.debug(f"ConfigData: Read Last Config file: {self.sys_pn_batch}")
        self.sys_cfg['sys_pn_batch'] = self.sys_pn_batch
        self.sys_cfg['sys_pn_wbs'] = self.sys_pn_wbs

        return

    def get_next_bat(self):
        """Returns the name of the next available yaml config file
        using the path filename stub_provided.
        requires: path_stub formatted as:
        ["G:/dev/v_chk/data/batch_files/"]
        """
        batch_num = 0
        c_file = f"{self.sys_dir_batch}{self.sys_id}_{batch_num:04d}.yaml"

        while Path(c_file).exists():
            batch_num += 1
            c_file = f"{self.sys_dir_batch}{self.sys_id}_{batch_num:04d}.yaml"
            logger.debug(f"ConfigData: Next Config file: {c_file}")

        self.sys_pn_batch = c_file
        self.sys_pn_wbs = f"{self.sys_dir_wbs}{Path(c_file).stem}.xlsx"
        self.sys_cfg['sys_pn_batch'] = self.sys_pn_batch
        self.sys_cfg['sys_pn_wbs'] = self.sys_pn_wbs

        logger.debug(f"ConfigData: Init Next Config file: {self.sys_pn_batch}")

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
              'sys_cfg': self.sys_cfg
            , 'wb_tabs': self.wb_tabs
            , 'wb_data': self.wb_data
        }

        return

    def write_bat_data(self):
        if not self.sys_pn_batch:
            self.get_next_bat()

        # self.wb_def_pack()   # This clobbers wb_data during tags! Upd explicitly!

        try:
            with open(self.sys_pn_batch, 'w') as yaml_file:
                # yaml.dump(range(50), width=50, indent=4)
                yaml.dump({
                    'wb_def':     self.wb_def
                }
                    , stream=yaml_file, sort_keys=False
                )
            return

        except Exception as e:
            print(f"ConfigData-write-wb_def ({self.sys_pn_batch}): Error in Save Config: {e}")
            sys.exit(1)


    def read_wb_data(self):
        if self.sys_pn_batch == '' or self.sys_pn_batch is None:
            self.get_last_bat()
            logger.debug(f"ConfigData-read_config: Loaded last config file: {self.sys_pn_batch}")
        else:
            logger.debug(f"ConfigData-read_config: Reading Config file: {self.sys_pn_batch}")
        try:
            with open(self.sys_pn_batch, 'r') as file_y:
                bat_data = file_y.read()

            wb_def_temp = yaml.safe_load(bat_data)
            wb_def_temp = wb_def_temp.get('wb_def', {})

        except Exception as e:
            raise Exception(f"ConfigData: Error reading config file ({self.sys_pn_batch}) Error : {e}")

        return wb_def_temp


if __name__ == "__main__":
    print(f"\nRunning standalone version of {Path(__file__).name}")

    # Build Tabs
    bat = WbDataDef()
