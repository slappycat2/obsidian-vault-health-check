import sys
from tempfile import template

import yaml
import json
import re
from pathlib import Path
from datetime import datetime
import os
from v_chk_class_lib import Colors

class ConfigSys:
    def __init__(self, sys_id, dbug_lvl=0):
        self.DBUG_LVL = dbug_lvl
        self.cfg_sys_id = sys_id
        self.cfg_sys_ver = "1.0.1"
        # Instantiate default values, presumably overridden in read_cfg_sys
        self.c_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.vault_id = 'o2'
        self.vault_path = 'E:\\o2\\'
        self.xl_exec_path  = "C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE"
        self.sys_dir = f"{Path(__file__).parent.parent.absolute()}\\"
        self.cfg_dir = f"{self.sys_dir}data\\cfgs\\"
        self.xls_dir = f"{self.sys_dir}data\\xlWork\\"
        self.tmpldir = self.get_templates_dir()
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
        self.cfg            = {}
        self.dirs_dot       = {}
        self.dirs_special   = {}
        self.dirs_skip      = {}
        self.ctot           = [0] * 10

        # NOTE: this is the SYSTEM config, not the wb runtime config
        pname = Path(self.sys_dir).joinpath(f"CONFIG.yaml")
        self.cfg_sys_pname = f"{pname}"
        if not os.path.isfile(self.cfg_sys_pname):
            self.init_cfg_sys()
        else:
            self.read_cfg_sys()

        self.Colors = Colors()

        # make sure directories exist
        data_dir = f"{self.sys_dir}data"
        self.mkdirs([data_dir, self.sys_dir, self.cfg_dir])

    def __getitem__(self, item):
        return self.cfg[item]

    def get_templates_dir(self):
        template_cfg_file = f"{self.vault_path}\\.obsidian\\plugins\\templater-obsidian\\data.json"
        try:
            if os.path.isfile(template_cfg_file):
                with open(template_cfg_file, 'r') as f:
                    template_cfg_json = f.read()
                template_cfg = json.loads(template_cfg_json)
                return template_cfg['templates_folder']
        except Exception as e:
                raise Exception(f"ConfigSys: Error in get_templates_dir: {e}")
        else:
            return ""

    def mkdirs(self, path):
        if isinstance(path, list):
            for p in path:
                self.mkdirs(p)
            return

        if os.path.isdir(path):
            return

        os.makedirs(path)

    def cfg_pack(self):
        self.cfg = {
              'cfg_sys_id':   self.cfg_sys_id
            , 'cfg_sys_ver':  self.cfg_sys_ver
            , 'c_date':       self.c_date
            , 'vault_id':     self.vault_id
            , 'vault_path':   self.vault_path
            , 'cfg_dir':      self.cfg_dir
            , 'xls_dir':      self.xls_dir
            , 'tmpldir':      self.tmpldir
            , 'xl_exec_path': self.xl_exec_path
            , 'tab_seq':      self.tab_seq
            , 'ctot':         self.ctot
        }

    def cfg_unpack(self):
        self.cfg_sys_id     = self.cfg['cfg_sys_id']
        self.cfg_sys_ver    = self.cfg['cfg_sys_ver']
        self.c_date         = self.cfg['c_date']
        self.vault_id       = self.cfg['vault_id']
        self.vault_path     = self.cfg['vault_path']
        self.cfg_dir        = self.cfg['cfg_dir']
        self.xls_dir        = self.cfg['xls_dir']
        self.tmpldir        = self.cfg['tmpldir']
        self.xl_exec_path   = self.cfg['xl_exec_path']
        self.tab_seq        = self.cfg['tab_seq']
        self.ctot           = self.cfg['ctot']

    def write_cfg_sys(self):
        try:
            ConfigSys.cfg_pack(self)
            # self.cfg_pack()
            with open(self.cfg_sys_pname, 'w') as yaml_file:
                # yaml.dump(range(50), width=50, indent=4)
                yaml.dump({'cfg': self.cfg}
                    , stream=yaml_file, sort_keys=False
                )
            return 0

        except Exception as e:
            raise Exception(f"ConfigSys: Uncaught Error raised in write_cfg_sys: {e}")
            # raise ConfigWriteError(f"ConfigSys ({self.cfg_sys_id}) write_cfg_sys: Error in Save Config: {e}")


    def read_cfg_sys(self):
        if self.DBUG_LVL > 1:
            print(f"ConfigSys ({self.cfg_sys_id}) read_cfg_config: Reading Config file: {self.cfg_sys_pname}")
        try:
            with open(self.cfg_sys_pname, 'r') as file_y:
                cfg_yaml = file_y.read()

            # ConfigSys.cfg_pack(self)
            self.cfg_pack()    # the values here, don't matter, but the keys must exist!
            cfg_loaded_yaml = yaml.safe_load(cfg_yaml)
            self.cfg = cfg_loaded_yaml['cfg']

            ConfigSys.cfg_unpack(self)

        except FileNotFoundError:
            res = input(f"ConfigSys: {self.cfg_sys_pname} Not Found. Create? (y/n):")
            if res == 'y' or res == 'Y':
                self.init_cfg_sys()
            else:
                print("ConfigSys: Unable to continue! Init Config Option Not Taken.")
                raise FileNotFoundError(f"ConfigSys: Error in read_cfg_sys: {self.cfg_sys_pname} Not Found.")

        except Exception as e:
            raise Exception(f"ConfigSys: Uncaught Error raised in read_cfg_sys: {e}")

        # test for validity, in case cfg file was edited...
        try:
            if self.cfg['cfg_sys_id'] != self.cfg_sys_id:
                raise Exception(f"ConfigSys: read_cfg_sys: cfg_sys_id mismatch: {self.cfg['cfg_sys_id']} != {self.cfg_sys_id}")


            test = Path(self.cfg['vault_path'])
            if not test.is_dir():
                raise Exception(f"ConfigSys: read_cfg_sys: Vault Not Found!: {self.vault_path}")

            test = Path(self.cfg['xl_exec_path'])
            if not test.is_file():
                raise Exception(f"ConfigSys: read_cfg_sys: xl_exec_path Not Found!: {self.xl_exec_path}")
        except Exception as e:
            raise Exception(f"ConfigSys: Uncaught Error validating {self.cfg_sys_pname}: {e}")

        return

    def init_cfg_sys(self):
        sys_id = self.cfg_sys_id
        if not os.path.isfile(self.cfg_sys_pname):
            self.cfg_sys_id = sys_id
            if sys_id == 'v_chk':
                self.cfg_pack()
                self.write_cfg_sys()
                if self.DBUG_LVL > 0:
                    print(f"ConfigSys: init_cfg_sys: Created new Config file: {self.cfg_sys_pname}")
            else:
               raise Exception(f"ConfigSys: init_cfg_sys: Error in init_cfg_sys: Unknown sys_id: {sys_id}")

class WbConfig(ConfigSys):
    def __init__(self, dbug_lvl):
        self.DBUG_LVL = dbug_lvl
        self.sys_id = 'v_chk'
        super().__init__(self.sys_id, self.DBUG_LVL)

        # Todo: Bug-007 - Need a way to find fileclass directory and Vault Path (from Vault Id)
        self.dirs_dot = [f.name for f in os.scandir(self.vault_path) if
                         f.is_dir() and f.path.startswith(f"{self.vault_path}\\.")]
        self.dirs_special = ["z_meta","z_resources"]
        self.dirs_skip = self.dirs_dot + self.dirs_special

        # regular expressions used in v_chk
        # REGEX is now hardcoded, here. Too much overhead maintaining in YAML.
        self.rgx_noTZdatePattern = r"([0-9]{4})[-\/]([0-1]?[0-9]{1})[-\/]([0-3])?([0-9]{1})(\s+)([0-9]{2}:[0-9]{2}:[0-9]{2})(.*)"
        self.rgx_noTZdateReplace = r"\1-\2-\3\4 \6"
        self.rgx_sub_strip_code_blocks = r'```[\s\S]*?```'
        self.rgx_sub_strip_inline_code = r'`[^`]*`'

        # regular expressions used in v_chk (compiled)
        self.rgx_boundary = re.compile('^---\\s*$', re.MULTILINE)
        self.rgx_body = re.compile('(^|(\\[))([)([A-Za-z0-9_]+)[:]{2}(.*?)(\\]?\\]?)($|\\])')
        self.rgx_tag_pattern = re.compile(r'#(\w+)')

        # WbConfig.cfg_pack(self) # include all of this in cfg{}
        self.cfg_pack() # include all of this in cfg{}

    def cfg_pack(self):        # pack cfg_wb
        # # ConfigSys.cfg_pack(self)
        # self.cfg_pack()

        wb_cfg = {
              'dirs_dot':     self.dirs_dot
            , 'dirs_special': self.dirs_special
            , 'dirs_skip':    self.dirs_skip
        # regex strings are now hardcoded in __init__ only. Too much overhead to maintain
        }

        self.cfg.update(wb_cfg)

    def cfg_unpack(self):
        # ConfigSys.cfg_unpack(self)
        self.cfg_unpack()

        self.dirs_dot = self.cfg['dirs_dot']
        self.dirs_special = self.cfg['dirs_special']
        self.dirs_skip = self.cfg['dirs_skip']
        # regex strings are now hardcoded in __init__ only. Too much overhead to maintain

class Error(Exception):
    """Base class for exceptions in this module."""
    pass

class WorkbookDefinitionError(Error):
    """Exception raised for errors in the Workbook Definition.

    Attributes:
        expression -- input expression in which the error occurred
        message -- explanation of the error
    """

    def __init__(self, expression, message):
        self.expression = expression
        self.message = message

class InputError(Error):
    """Exception raised for errors in the input.

    Attributes:
        expression -- input expression in which the error occurred
        message -- explanation of the error
    """

    def __init__(self, expression, message):
        self.expression = expression
        self.message = message

class ConfigWriteError(Exception):
    def __init__(self, expression, message):
        self.expression = expression
        self.message = message

if __name__ == "__main__":
    # Test if this works
    DBUG_LVL = 1

    cfg = WbConfig(DBUG_LVL)

    # cfg.read_cfg_sys()

    print(f"cfg.vault_path = {cfg.vault_path}")

    if cfg:
        print(f'cfg_sys_id                      :  {cfg.cfg_sys_id}')
        print(f'c_date                          :  {cfg.c_date}')
        print(f'vault_path                      :  {cfg.vault_path}')
        print(f'sys_dir                         :  {cfg.sys_dir}')
        print(f'cfg_dir                         :  {cfg.cfg_dir}')
        print(f'xls_dir                         :  {cfg.xls_dir}')
        print(f'xl_exec_path                    :  {cfg.xl_exec_path}')
        print(f'cfg_sys_id                      :  {cfg.cfg_sys_id}')
        print(f'dirs_dot                        :  {cfg.dirs_dot}')
        print(f'dirs_special                    :  {cfg.dirs_special}')
        print(f'dirs_skip                       :  {cfg.dirs_skip}')
        print(f'tab_seq                         :  {cfg.tab_seq}')
        print(f'ctot                            :  {cfg.ctot}')
        print(f'rgx_boundary                    :  {cfg.rgx_boundary}')
        print(f'rgx_body                        :  {cfg.rgx_body}')
        print(f'rgx_noTZdatePattern             :  {cfg.rgx_noTZdatePattern}')
        print(f'rgx_noTZdateReplace             :  {cfg.rgx_noTZdateReplace}')
        print(f'rgx_sub_strip_code_blocks       :  {cfg.rgx_sub_strip_code_blocks}')
        print(f'rgx_sub_strip_inline_code       :  {cfg.rgx_sub_strip_inline_code}')

    print(f'\nStandalone run of "v_chk_cfg_sys.py" complete.')