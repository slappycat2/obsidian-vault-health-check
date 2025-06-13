import platform
import os

from pathlib import Path
import json

class JsonFile:
    """
    JsonFile class is used to load a JSON file and store its contents.
    Attributes:
        json_path (str): Path to the JSON file.
        json_data (dict): Parsed JSON data.
        err_msg (str): Error message if any exception occurs during loading.
    """
    def __init__(self, json_path: str):
        self.json_path = json_path
        self.json_data = {}
        self.err_msg = ''

        try:
            with open(self.json_path, 'r', encoding='utf8') as f:
                self.json_data = json.load(f)

        except FileNotFoundError:
            self.err_msg = f"Exception: {json_path} not found. "

        except json.JSONDecodeError:
            self.err_msg = f"Exception: {json_path} is not a valid JSON file. "

        except Exception as e:
            raise Exception(f"JsonFile: load_json_file attempting to read: ({json_path}) - Error: {e}")


class Colors:
    # There are 16 million colors, these are ones I picked to play with,
    def __init__(self):
        """
        Initializes color schemes for a workbook.
        Class objects:
            name: name of base color
            clr1: base color, to be used for tab color, table headings, and title.
            txt1: text color for clr1.
            clr2: secondary color, to be used for Totals headings, and other highlights
            txt2: text color for clr2.
            table_style: Table style to be used for the tab based on a given
                table_row_style (see below).

        tbl_row_style: Defines how alternating rows will be displayed. Table Headers colors are
            overridden by the tabs primary color. This means theme changes will only change table
            row colors.
            Possible values are:
                0=wht/wht, 1=clr/wht, 2=clr/clrA, 3=gry/wht, 4=dark

        Class methods:
        get_tab_clrs(self, tab_id, shade=None, row_style=None)
        get_clr(self, name, shade)
        get_table_style(self, name, row_style)
        get_txt_clr(self, hex_color)
        get_base_clr(self, hex_color)
        init_tbl_clrs(self)
        init_tbl_txts(self)
        init_tab_clrs(self)
        init_table_style(self)
        complement(clr_in)

        get_tab_clrs(tab_id)

        """
        # , name='wht', shade=0, tbl_row_style=2
        #tabs need: clrtb, clrhi, clrfl, clrtx, clrflh, clrtxh, tab_table_style
        # renameTo: clrtb, clrf2, clrf1, clrt1, clrt2
        self.name = "blu"
        self.dflt_shade = 2         # 60% of self.name
        self.dflt_row_style = 5     # v_chk custom style
        self.tab_id = None
        self.tbl_row_style = 4      # 4 is the dark one, but this is init, only
        self.shade = 0              # must be 0-6 (0=solid, 6 is dark)
        self.table_style = None     # e.g., TableStyleMedium21
        self.base_clr = None        # if needed, when doing shade lookups
        self.table_styles = []
        self.row_clr_idx = {}
        self.err_txt = '00FF0000'   # bright red for error text
        self.clr1 = None
        self.clr2 = None
        self.txt1 = None
        self.txt2 = None

        # Todo: Add dflt_shade and dflt table_row_style to CFG in Colors

        self.tbl_clrs = self.init_tbl_clrs()
        self.tbl_txts = self.init_tbl_txts()
        self.tab_clrs = self.init_tab_clrs()
        self.table_styles, self.row_clr_idx = self.init_table_style()

        self.clr_wht = self.get_clr("wht", 0)
        self.clr_blk = self.get_clr("blk", 0)
        self.clr_aqu = self.get_clr("aqu", 0)
        self.clr_red = self.get_clr("red", 0)
        self.clr_ora = self.get_clr("ora", 0)
        self.clr_yel = self.get_clr("yel", 0)
        self.clr_sea = self.get_clr("sea", 0)
        self.clr_tea = self.get_clr("tea", 0)
        self.clr_pur = self.get_clr("pur", 0)
        self.clr_blu = self.get_clr("blu", 0)
        self.clr_grn = self.get_clr("grn", 0)

    def get_tab_clrs(self, tab_id, shade=None, row_style=None):
        # returns:
        #       self.clr1
        #       self.txt1
        #       self.clr2
        #       self.txt2
        #       self.tbl_style
        tab_clrs_def = self.tab_clrs[tab_id]
        self.name, self.base_clr = tab_clrs_def

        if shade is None:
            self.shade = tab_clrs_def[1]
        else:
            self.shade = shade

        self.clr1 = self.tbl_clrs[self.name][0]
        self.clr2 = self.tbl_clrs[self.name][self.shade]

        self.txt1 = self.tbl_txts[self.clr1][0]
        self.txt2 = self.tbl_txts[self.clr2][0]

        if row_style is None:
            self.tbl_row_style = self.dflt_row_style
        else:
            self.tbl_row_style = row_style

        self.table_style = self.get_table_style(self.name, self.tbl_row_style)

        return self.clr1, self.txt1, self.clr2, self.txt2, self.table_style

    def get_clr(self, name, shade):
        return self.tbl_clrs[name][shade]

    def get_table_style(self, name, row_style):
        # Table Styles based on table_styles[row_index]
        # These are based on the Excel Accent Colors (1-6)
        # Plus my two additional accent colors
        idx =  self.row_clr_idx[name]
        style_prefix = self.table_styles[row_style][0]
        return f"{style_prefix}{self.table_styles[row_style][idx]}"

    def get_txt_clr(self, hex_clr):
        return self.tbl_txts[hex_clr][0]

    def get_base_clr(self, hex_clr):
        return self.tbl_txts[hex_clr][1]

    @staticmethod
    def init_tbl_clrs():
        clrs = {
            # Based on Excel Ion Theme (%' refers to whiteness-so light to dark)
            # shades         0         1         2         3         4         5
            # Accents	   Solid	  80%   	60%	      40%   	25%       dark
                  "wht": ["FFFFFF", "D3D3D3", "B0B0B0", "757575", "A5A5A5", "808080"]
                , "blk": ["000000", "7F7F7F", "595959", "3F3F3F", "262626", "0D0D0D"]
                , "aqu": ["1E5155", "C4E7EA", "8AD0D5", "4FB8C1", "163C3F", "0F282B"]
                , "red": ["B01513", "F8C6C6", "F28E8D", "EC5654", "840F0E", "580C0A"]
                , "ora": ["EA6312", "FBDFCF", "F7C09F", "F3A16F", "AF4A0D", "753109"]
                , "yel": ["E6B729", "F9F0D4", "F4E2A9", "F0D37E", "B58E15", "7A600E"]
                , "sea": ["6AAC90", "E1EEE8", "C3DDD2", "A5CDBC", "4A856C", "325848"]
                , "tea": ["54849A", "DBE6EB", "B8CED8", "95B6C5", "3F6373", "29424E"]
                , "pur": ["9E5E9B", "EBDEEB", "D8BED7", "C59DC3", "764674", "4F2F4E"]
                , "blu": ["4169E1", "DADDF6", "AFB5EB", "7C85DE", "2E3BB4", "212A83"]
                , "grn": ["57BF70", "B1E1BC", "B6E4C1", "8AD29B", "2C763E", "1D4B28"]
        }

        return clrs

    def init_tbl_txts(self):
        # tbl_txts is a dict with keys for all clrs: [clrtx, baseclr (used for tbl_style)]
        # it will look like this:
        #    a_clr      clrtx     baseclr (so you can look up table_styles)
        #               # wht[['FFFFFF', 'D3D3D3', 'B0B0B0', '757575', 'A5A5A5', '808080']]
        # , 'FFFFFF': ['000000', 'FFFFFF']     # wht[0]  <== [0] is the baseclr which will be the
        # , 'D3D3D3': ['FFFFFF', 'FFFFFF']     # wht[1]          same for all shades [1-5]
        # , 'B0B0B0': ['FFFFFF', 'FFFFFF']     # wht[2]
        # , '757575': ['FFFFFF', 'FFFFFF']     # wht[3]
        # , 'A5A5A5': ['000000', 'FFFFFF']     # wht[4]
        # , '808080': ['000000', 'FFFFFF']     # wht[5]
        #
        #               # blk[['000000', '7F7F7F', '595959', '3F3F3F', '262626', '0D0D0D']]
        # , '000000': ['000000', '000000']     # blk[0]
        # , '7F7F7F': ['FFFFFF', '000000']     # blk[1]
        # , '595959': ['FFFFFF', '000000']     # blk[2]
        # , '3F3F3F': ['FFFFFF', '000000']     # blk[3]
        # ...

        # Non-Default Text Colors (see c_chk_colorsClass.xlsx)
                        # wht4      blk1      blk2      blk3      red3      yel0      sea0      grn0
        special_txts = ["A5A5A5", "7F7F7F", "595959", "3F3F3F", "EC5654", "E6B729", "6AAC90", "57BF70"]
        tbl_txts = {}
        for clr, val in self.tbl_clrs.items():
            baseclr = val[0]
            # print(f"\n              # {clr}[{val}]")
            for i in range(len(val)):
                if i > 0 or i > 3:
                    textclr = "000000" # blk
                else:
                    textclr = "FFFFFF"

                if val[i] in special_txts:
                    textclr = self.complement(textclr)

                tbl_txts[val[i]] = [textclr, baseclr]

        return tbl_txts

    def init_tab_clrs(self):
        tab_clrs = {
            #  tab_id: [tab_color, tab_clr highlights]
              'pros': ['blu', self.dflt_shade]    # , self.clr_redio]
            , 'vals': ['blu', self.dflt_shade]    # , self.clr_redio]
            , 'tags': ['grn', self.dflt_shade]    # , self.clr_yelio]
            , 'xyml': ['red', self.dflt_shade]    # , self.clr_yelio]
            , 'dups': ['red', self.dflt_shade]    # , self.clr_yelio]
            , 'file': ['pur', self.dflt_shade]    # , self.clr_yelio]
            , 'tmpl': ['ora', self.dflt_shade]    # , self.clr_grnio]
            , 'code': ['ora', self.dflt_shade]    # , self.clr_grnio]
            , 'nest': ['yel', self.dflt_shade]    # , self.clr_grnio]
            , 'plug': ['yel', self.dflt_shade]    # , self.clr_grnio]
            , 'summ': ['red', self.dflt_shade]    # , self.clr_yelio]
            , 'ar51': ['red', self.dflt_shade]    # , self.clr_yelio]
        }

        return tab_clrs

    @staticmethod
    def init_table_style():
        # Table Styles
        # These are based on the Excel Accent Colors (1-6)
        # Plus my two additional accent colors
        table_styles = [
                #          0             1    2     3      4      5     6      7
                #     style_prefix      wht   red  ora    yel    grn   blu    pur]   row_idx
                  ["TableStyleLight",   "8",  "9", "10",  "11",  "12", "13", "14"]    # 0
                , ["TableStyleMedium",  "1",  "2",  "3",   "4",   "5",  "6",  "7"]    # 1
                , ["TableStyleMedium",  "8",  "9", "10",  "11",  "12", "13", "14"]    # 2
                , ["TableStyleMedium", "15", "16", "17",  "18"   "19", "20", "21"]    # 3
                , ["TableStyleDark",    "1",  "2",  "3",   "4",   "5",  "6",  "7"]    # 4
                , ["TableStyleMedium", "15", "20", "19",  "19",  "17", "16", "18"]    # 5=v_chk style
                ]
        row_clr_idx = {
                  "wht": 1   # (accent colors from Ion Theme)
                , "red": 2   # (accent color)
                , "ora": 3   # (accent color)
                , "yel": 4   # (accent color)
                , "sea": 5   # (accent color)
                , "tea": 6   # (accent color)
                , "pur": 7   # (accent color)
                , "blu": 6   # uses teal
                , "grn": 5   # uses sea
                , "blk": 1   # uses wht
                , "aqu": 2   # uses red
                }

        return table_styles, row_clr_idx

    @staticmethod
    def complement(clr_in):
        a_red = 255 - int(clr_in[0:2], base=16)
        a_grn = 255 - int(clr_in[2:4], base=16)
        a_blu = 255 - int(clr_in[4:6], base=16)
        # f"{i:02x}"
        complement = f"{a_red:02x}{a_grn:02x}{a_blu:02x}"

        return complement

# End of Colors Class ==================================================================

class PluginMan:
    def __init__(self, v_path=None):
        self.v_path = v_path
        self.DBUG_LVL = 0
        self.id                 = ''
        self.name               = ''
        self.version            = ''
        self.minAppVersion      = ''
        self.description        = ''
        self.author             = ''
        self.authorUrl          = ''
        self.helpUrl            = ''
        self.plugin_dir         = ''
        self.plugs_lib = {}
        self.obs_plugs = {}
        self.known_plug_sigs = {
              'dataview': 'dataview'
            , 'dataviewjs': 'dataview'
            , 'button': 'buttons'
            , 'gevent': 'google-calendar'
            , 'ccard': 'folder-note-plugin'
            , 'leaflet': 'leaflet'
            , 'mermaid': 'mermaid-tools'
            , 'todoist': 'todoist-sync-plugin'
            , 'tracker': 'obsidian-tracker'
        }
        if self.v_path is not None:
            self.get_plugs_lib()
            self.get_obs_plugs()

    def __getitem__(self, key):
        return self.plugs_lib[key]
    def __setitem__(self, key, value):
        self.plugs_lib[key] = value
    def __delitem__(self, key):
        del self.plugs_lib[key]
    def __contains__(self, key):
        return key in self.plugs_lib
    def __len__(self):
        return len(self.plugs_lib)
    def __iter__(self):
        return iter(self.plugs_lib)
    def __str__(self):
        return str(self.plugs_lib)
    def __repr__(self):
        return repr(self.plugs_lib)
    def __call__(self):
        return self.plugs_lib

    def get_name(self, cb_sig):
        cb_sig = cb_sig.lower()
        cb_name = ''
        if cb_sig in self.known_plug_sigs:
            try:
                cb_sig = self.known_plug_sigs[cb_sig]
                cb_name  = self.plugs_lib[cb_sig]['name']
            except KeyError:
                pass
            except Exception as e:
                print(f"Error: {e}")
                raise Exception(f"PluginMan: {e}")
        return cb_name

    def get_plugin_sig_list(self, plug_id):
        plugin_sig_list = []
        for sig, sig_plug_id in self.known_plug_sigs.items():
            if plug_id == sig_plug_id:
                plugin_sig_list.append(sig)
        return plugin_sig_list

    def get_obs_plugs(self):
        self.obs_plugs = {}
        for plugin_id, a_plugin_dict in self.plugs_lib.items():
            a_plugin_list = [''] * 10
            a_plugin_list[0] = plugin_id
            a_plugin_list[1] = a_plugin_dict.setdefault('name', '')
            a_plugin_list[2] = a_plugin_dict.setdefault('enabled', '')
            a_plugin_list[3] = a_plugin_dict.setdefault('version', '')
            a_plugin_list[4] = a_plugin_dict.setdefault('minAppVersion', '')
            a_plugin_list[5] = a_plugin_dict.setdefault('author', '')
            a_plugin_list[6] = a_plugin_dict.setdefault('authorUrl', '')
            a_plugin_list[7] = a_plugin_dict.setdefault('isDesktopOnly', '')
            a_plugin_list[8] = a_plugin_dict.setdefault('description', '')
            a_plugin_list[9] = a_plugin_dict.setdefault('plugin_sig_list', '')
            self.obs_plugs[plugin_id] = {plugin_id: a_plugin_list}

        return self.obs_plugs

    def get_plugs_lib(self):
        o_path = f"{self.v_path}/.obsidian/"
        v_path = f"{o_path}plugins/"
        v_path_obj = Path(v_path)
        cp_json = f"{o_path}community-plugins.json"
        enabled_plugins = []
        # First, load the the list of enabled plugins from community_plugins.json
        enabled_plugins_obj = JsonFile(cp_json)
        if enabled_plugins_obj.err_msg:
            print(f"PluginMan: get_plugs_lib-Error: {enabled_plugins_obj.err_msg}")

        enabled_plugins = enabled_plugins_obj.json_data

        # Now, load the plugin descriptions from each manifest.json files
        #   These are the INSTALLED plugins.
        self.plugs_lib = {}
        plugin_dir = ''
        for mj_file in v_path_obj.rglob("manifest.json"):
            try:
                plugin_dir = mj_file.parent.name
                with open(mj_file, 'r', encoding='utf8') as f:
                    mj_data = json.load(f)
                    mj_data['plugin_dir'] = plugin_dir
                    mj_data['plugin_sig_list'] = self.get_plugin_sig_list(mj_data['id'])

                    if mj_data['id'] in enabled_plugins:
                        mj_data['enabled'] = True
                    else:
                        mj_data['enabled'] = False

                    if self.DBUG_LVL > 98:
                        print(f"Created obj: {mj_data['id']: <30} {mj_data['name']}")
                        # print(f"{t}dir: {plugin_dir}  id:{mj_data['id']}  name:{mj_data['name']}")
                    self.plugs_lib[mj_data['id']] = mj_data

            except Exception as e:
                print(f"  plugin_dir: {plugin_dir}")
                print(f"  mj_file: {mj_file}")
                raise Exception(f"PluginMan: get_plugs_lib-Error: {e}")

class ObsidianApp:
    """
    ObsidianApp class is a placeholder for Obsidian application related methods.
    Currently, it does not contain any methods or attributes.
    possible platforms: Linux, Darwin, Windows
    """
    def __init__(self):
        self.obs_platform = platform.system()
        self.dir_obs_json = None
        self.dir_obs_vault = None
        self.pn_obs_json = None
        self.obs_vaults_open = []
        self.obs_vaults = {}

        home_dir = Path.home()
        home_dir_str = str(home_dir)

        if self.obs_platform == 'Windows':
            home_dir_str = os.getenv('APPDATA', '')           # Not really Home, but it works for now...

        obs_json_locs = {
              'Linux': f'{home_dir_str}/.config/obsidian/'
            , 'Darwin': f'{home_dir_str}/Library/Application Support/obsidian/'
            , 'Windows': f'{home_dir_str}/obsidian/'
        }

        self.dir_obs_json = obs_json_locs.get(self.obs_platform, '')
        self.pn_obs_json = f"{self.dir_obs_json}obsidian.json"
        obs_json_obj = JsonFile(self.pn_obs_json)
        if obs_json_obj.err_msg:
            raise Exception(f"ObsidianApp: {obs_json_obj.err_msg}")

        obs_json_dict = obs_json_obj.json_data
        if 'vaults' not in obs_json_dict:
            raise Exception(f"ObsidianApp: No vaults found in obsidian.json at {self.dir_obs_json}")

        vaults_dict = obs_json_dict['vaults']

        for vault_id, vault_dict in vaults_dict.items():
            if 'path' in vault_dict:
                self.obs_vaults[vault_dict['path']] = vault_id
                if 'open' in vault_dict and vault_dict['open']:
                    self.obs_vaults_open.append(vault_id)

        if not self.obs_vaults:
            raise Exception(f"ObsidianApp: No Vaults with a 'path' key in obsidian.json")

def main() -> None:
    pass

    print(f'\nStandalone run of "{Path(__file__).name}" complete.')


if __name__ == '__main__':
    main()



