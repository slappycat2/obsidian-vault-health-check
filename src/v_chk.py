from operator import truediv

from v_chk_wb_tabs import *
from v_chk_xl import *

# This s/b WbDataDef and v_chk should just instantiate the system.
class VaultHealthCheck:   # WbConfig
    def __init__(self, dbug_lvl):
        # super().__init__()
        self.DBUG_LVL = dbug_lvl
        self.dbug = False
        # self.dbug = 'Terminal Color Escape Sequences.md'
        if self.DBUG_LVL >= 0:
            print(f"Obsidian Vault Health Check v1.0\n\nLoading configuration data...")
        self.DBUG_AREA51 = True
        self.DBUG_TAB = 'summ'  # DBUG_LVL must be greater than 2
        # self.DBUG_LVL = 0  # Do Not print anything
        # self.DBUG_LVL = 1  # print report level actions only (export, load, save, etc.) + all lower levels
        # self.DBUG_LVL = 2  # print object level actions + all lower levels
        # self.DBUG_LVL = 3  # print export_tab + all lower levels
        # self.DBUG_LVL = 4  # print hdr records + all lower levels
        # self.DBUG_LVL = 5  # print detail records + all lower levels
        # self.DBUG_LVL = 9  # print everything (includes export_cell!)
        self.key_stack = []
        # self.cfg_setup = WbDataDef(DBUG_LVL)
        self.wb_data_obj = WbDataDef(DBUG_LVL)

        self.wb_data_obj.get_next_cfg()     # this initials wb_data
        self.plugin_id_def = self.wb_data_obj.plugin_id_def
        self.wb_def = self.wb_data_obj.wb_def

        self.cfg = self.wb_def.get('cfg', {})
        self.tmpldir = Path(self.cfg.get('tmpldir', ''))
        self.isTemplate = False
        self.wb_data = self.wb_def.get('wb_data', {})
        self.obs_props = self.wb_data.get('obs_props', {})
        self.obs_atags = self.wb_data.get('obs_atags', {})
        self.obs_xyaml = self.wb_data.get('obs_xyaml', {})
        self.obs_dupfn = self.wb_data.get('obs_dupfn', {})
        self.obs_files = self.wb_data.get('obs_files', {})
        self.obs_tmplt = self.wb_data.get('obs_tmplt', {})
        self.obs_codes = self.wb_data.get('obs_codes', {})
        self.obs_nests = self.wb_data.get('obs_nests', {})
        self.obs_plugs = self.wb_data.get('obs_plugs', {})
        self.rgx_boundary = re.compile('^---\\s*$', re.MULTILINE)
        self.rgx_body_pros = re.compile('(^|(\\[))([)([A-Za-z0-9_]+)[:]{2}(.*?)(\\]?\\]?)($|\\])')
        self.rgx_tag_pattern = re.compile('[^|\w]#(\w+)', re.MULTILINE)
        self.rgx_noTZdatePattern = re.compile(r"([0-9]{4})[-\/]([0-1]?[0-9]{1})[-\/]([0-3])?([0-9]{1})(\s+)([0-9]{2}:[0-9]{2}:[0-9]{2})(.*)", re.MULTILINE)
        self.rgx_code_blocks = re.compile(r'^`{3}[\s\S]*?^`{3}', re.MULTILINE)
        self.rgx_code_inline = re.compile(r'`[^`]*`', re.MULTILINE)
        self.rgx_templater_strs = r"<%[\*]?\s*.*?\s*%>"
        self.rgx_wikilinks = re.compile(r"\[\[.*?\]\]", re.MULTILINE)

        self.filepath = ""
        self.inline = "F"
        self.actual_prop_key = ""
        self.plugin_id = ""
        self.ctot = [0] * 10

        plugin_lib = PluginMan(self.cfg['vault_path'])
        self.obs_plugs = plugin_lib.get_obs_plugs()

        self.process_vault()

    def process_vault(self):
        if self.DBUG_LVL >= 0:
            print(f"Gathering statistics on vault Id: {self.cfg['vault_id']}  Path: {self.cfg['vault_path']}...")

        v_path_obj = Path(self.cfg['vault_path'])
        for md_file in v_path_obj.rglob("*.md"):
            self.ctot[0] += 1

            # md_file and BASE_DIR are WindowsPath objects, not a strings!
            if self.dbug and str(md_file.name) != self.dbug:
                # we're debugging a single file and this isn't it!
                # print(f"dbug is on: Skipping_file: {md_file} is not equal to {self.dbug}")
                continue

            x_dir_test = False
            for x_dir in md_file.parts:
                if x_dir in self.cfg['dirs_skip']:
                    x_dir_test = True
                    continue  # this only exits this for loop
            if x_dir_test:
                self.ctot[1] += 1
                print(f"Skipping file: {md_file} is in dirs_skip")
                continue # this gets the next file...

            self.isTemplate = False
            if self.is_subdirectory(md_file, self.tmpldir):
                self.isTemplate = True
                self.ctot[2] += 1
                # right now, support for templates is not implemented.
                # this will require a special decoding of the markdown
                # without using PyYaml, since they would not load properly
                # otherwise it will like all be invalid properties...
                # continue

            md_pname = str(md_file)

            if self.DBUG_LVL > 2:
                print(f"Processing file: {md_file}")
            self.process_md_file(md_pname)

        # Vault Processing Complete! Write to wb_def

        self.wb_def['wb_data']['obs_props'] = self.obs_props
        self.wb_def['wb_data']['obs_xyaml'] = self.obs_xyaml
        self.wb_def['wb_data']['obs_dupfn'] = self.obs_dupfn
        self.wb_def['wb_data']['obs_files'] = self.obs_files
        self.wb_def['wb_data']['obs_tmplt'] = self.obs_tmplt
        self.wb_def['wb_data']['obs_codes'] = self.obs_codes
        self.wb_def['wb_data']['obs_nests'] = self.obs_nests
        self.wb_def['wb_data']['obs_plugs'] = self.obs_plugs
        self.cfg['ctot'] = self.ctot
        # Vault processing complete! Get wb tab defs,
        self.wb_data_obj.write_cfg_data()

        if self.DBUG_LVL > 0:
            print(f"Vault ({self.cfg['vault_path']}) processing complete.")

        return

    def process_md_file(self, filepath):
        self.filepath = filepath

        self.upd_obs_props(self.obs_dupfn, 'dupfn', Path(filepath).name, filepath)
        self.ctot[3] += 1
        if self.DBUG_LVL > 2:
            print(f"Processing file: {self.filepath}")
        self.parse_file()

        # self.obs_files[self.filepath] = self.file_pros

    def parse_file(self):
        self.plugin_id = ""
        with open(self.filepath, 'r', encoding='utf-8') as file:
            full_content = file.read()

        content = self.rgx_code_blocks.sub('', full_content)
        content = self.strip_inline_code(content)

        y_text, x_text = self.split_file(content)
        if len(y_text) != 0:
            self.plugin_id = ''.join([pid for pid in self.plugin_id_def if pid in y_text])
            if self.plugin_id == "NestedDictionary":  # YAML with the word NestedDictionary is okay!
                self.ctot[4] += 1
                self.plugin_id = ""
            self.prop_loc_F_I = "F"
            self.ctot[5] += 1
            self.process_yaml(y_text)

        if len(x_text) != 0:
            self.prop_loc_F_I = "I"
            self.ctot[6] += 1
            self.process_body(x_text)

        self.process_code_blocks(full_content)

    def process_code_blocks(self, content):
        cb_list = self.extract_codeblocks(content)
        for cb in cb_list:
            cb_sig = self.extract_codeblock_info(cb)
            # file_cb_sig = f"{self.filepath}|{cb_sig}"
            self.upd_obs_props(self.obs_codes, self.filepath, cb_sig, cb)

    def strip_templater_strs(self, markdown_text):
        clean_text = re.sub(self.rgx_templater_strs, "", markdown_text, flags=re.DOTALL)
        return clean_text

    def strip_wikilinks(self, markdown_text):
        clean_text = re.sub(self.rgx_wikilinks, "", markdown_text)
        return clean_text

    def strip_codeblocks(self, markdown_text):
        return re.sub(self.rgx_code_blocks, "", markdown_text)

    def strip_inline_code(self, markdown_text):
        return re.sub(r"`[^`]*`", "", markdown_text)

    def get_tags_list(self, markdown_text):
        temp_content = copy.deepcopy(markdown_text)
        temp_content = self.strip_codeblocks(temp_content)
        temp_content = self.strip_inline_code(temp_content)
        temp_content = self.strip_templater_strs(temp_content)
        temp_content = self.strip_wikilinks(temp_content)
        tag_list = set(self.rgx_tag_pattern.findall(temp_content))
        return tag_list

    def process_body(self, body_text):
        body_text = "".join(body_text)
        # strip code from text
        # body_text = re.sub(self.rgx_code_blocks, '', body_text)
        # body_text = re.sub(self.rgx_code_inline, '', body_text)

        match_pros = list(self.rgx_body_pros.finditer(body_text))

        # body pros
        for idx, match in enumerate(match_pros):
            m = match_pros[idx].group()

            if self.DBUG_LVL > 8:
                print(f"m={m}  len: {len(m.split('::'))}")

            k, v = m.split("::")
            k = k.strip()
            v = v.strip()
            if k.startswith("["):
                k = k[2:]
            if v.endswith("]"):
                v = v[:-1]
            # if k.endswith(":"):
            #     k = k[:-1]

            self.upd_val(k, v)

        # body tags
        body_tags = self.get_tags_list(body_text)

        for tag in body_tags:
            if tag.isnumeric():
                continue
            else:
                self.upd_val("tags", tag)

        return

    def process_yaml(self, front_text):
        data = {}
        try:
            data = yaml.safe_load(front_text) or {}
            if not isinstance(data, dict):
                self.upd_obs_props(self.obs_xyaml, 'NonD', f"{self.filepath}", self.filepath)
                return

        except yaml.YAMLError as e:
            if self.DBUG_LVL > 8:
                print(f"Error in YAML: {e}")
            self.upd_obs_props(self.obs_xyaml, 'BadY', self.filepath, self.filepath)
            return

        except Exception as e:
            if self.DBUG_LVL >= 0:
                print(f"Unhandled Exception:\t {self.filepath}\n{e}\n\n")
                # e = unhashable type: 'dict'
            self.upd_obs_props(self.obs_xyaml, 'ErrY', self.filepath, self.filepath)
            return

        if data:
            self.unpack_yaml(None, data)
        else:
            self.upd_obs_props(self.obs_xyaml, 'MtFm', self.filepath, self.filepath)
        return

    # new version here
    def unpack_yaml(self, key_passed, a_yaml_dict):
        # For nested dictionaries, this will run reciprocally
        obs_prop_key = ""
        slash = ""

        if key_passed is None:
            obs_prop_key = ""  # otherwise, we get a "/" appended to all keys
        else:
            # if len(self.key_stack) == 1:
            slash = "/"
            for ks in self.key_stack:
                obs_prop_key = f"{obs_prop_key}{ks}{slash}"
                slash = "/"

        # process yaml dictionary
        for key, value in a_yaml_dict.items():
            self.actual_prop_key = key
            key = key.lower()
            if self.actual_prop_key == key:
                self.actual_prop_key = ""

            # if key == "fields" or key_passed == "fields":
            #     print(key)
            # else:
            #     print(f"key_stack: {self.key_stack}   key:{key}: ({value.__class__}) {value}")

            if isinstance(value, list) and len(value) == 0:
                value = None

            if (isinstance(value, list)
                    and isinstance(value[0], list)
                    and len(value) == 1):
                value = f"[[{value[0][0]}]]"

            if isinstance(value, dict):
                # nested dicts not allowed in Obsidian, so this must be a plugin
                self.key_stack.append(key)
                if self.plugin_id is None or self.plugin_id == "":
                    # but no id tag was found using plugin_id_defs in process_yaml
                    self.plugin_id = 'NestedDictionary'
                self.unpack_yaml(key, value)
                self.key_stack.pop()
                continue
            if isinstance(value, list):
                for item in value:
                    if isinstance(item, dict):
                        # nested dicts not allowed in Obsidian--this must be a plugin
                        self.key_stack.append(key)
                        if self.plugin_id is None or self.plugin_id == "":
                            # but no id tag was found using plugin_id_defs in process_yaml
                            self.plugin_id = 'NestedDictionary'

                        self.unpack_yaml(key, item)
                        self.key_stack.pop()
                        continue
                    else:
                        self.upd_val(f"{obs_prop_key}{key}", item)
            else:
                # otherwise, the value is a single string, bool, etc.
                self.upd_val(f"{obs_prop_key}{key}", value)
        return

    def upd_val(self, k, v):
        if k == "tags":
            v = v.lower()

        if self.plugin_id != "":
            # Here, we're passing in a subset of obs_nests, the set for this plugin, only
            if v is None or v == "":
                v = "(-None-)"
            self.upd_obs_nests(self.obs_nests, k, v, self.filepath)
        else:
            if k == "tags":
                self.upd_obs_props(self.obs_atags, v, v, self.filepath) # Note k is not used here...
            else:
                self.upd_obs_props(self.obs_props, k, v, self.filepath)

        self.upd_obs_files(self.obs_files, k, v, self.filepath)
        # upd_props_dict(cfg.props, k, v, self.filepath)
        # upd_props_dict(self.file_props, k, v, self.filepath)
        return

    def upd_obs_files(self, o_files, ukey, uval, fkey):
        self.ctot[7] += 1
        ukey = ukey.lower()
        fkey = f"{fkey}|{self.prop_loc_F_I}"    # (F)frontmatter or (I)inline indicator

        o_files.setdefault(fkey, {ukey: [self.actual_prop_key]})    # First, make sure fkey exists...
        o_files[fkey].setdefault(ukey,  [self.actual_prop_key])     # Then, make sure ukey exists...
        o_files[fkey][ukey].append(uval)                            # Then update it

        return

    def upd_obs_nests(self, o_files, ukey, uval, fkey):
        self.ctot[8] += 1
        ukey = ukey.lower()
        fkey = f"{self.plugin_id}|{fkey}"

        o_files.setdefault(fkey, {ukey: []})    # First, make sure fkey exists...
        o_files[fkey].setdefault(ukey,  [])     # Then, make sure ukey exists...
        o_files[fkey][ukey].append(uval)        # Then update it

        return

    def upd_obs_props(self, obs_dict, ukey, uval, filepath):
        self.ctot[9] += 1
        f_list = []

        # f_list = obs_dict.setdefault(ukey, {uval: []})
        if ukey in obs_dict:
            k_dict = obs_dict[ukey]
        else:
            k_dict = {uval: []}
            obs_dict[ukey] = k_dict
        # we have to get the existing files list, before we can append to it...
        if uval in k_dict:
            f_list = k_dict[uval]

        f_list.append(filepath)
        obs_dict[ukey][uval] = f_list

        return

    def split_file(self, content):
        file_text = "".join(content)
        yaml_match = list(self.rgx_boundary.finditer(file_text))

        if len(yaml_match) < 2:
            self.upd_obs_props(self.obs_xyaml, 'NoFm', self.filepath, self.filepath)
            return "", file_text

        start, end = yaml_match[0].end(), yaml_match[1].start()
        yaml_text = file_text[start:end].strip()  # Extract YAML content only
        body_text = file_text[end:]

        return yaml_text, body_text

    @staticmethod
    def extract_codeblock_info(a_codeblock) -> tuple:
        """
        Extracts detailed information from a code block input.

        This function processes the input `a_codeblock`, which could either be a string or a list of strings,
        and extracts the type and action details of the code block. The type of the code block (e.g.,
        BUTTON, DATAVIEW) is derived from the first line of the input, while the action associated with
        the code block is obtained from subsequent details.

        :param a_codeblock: A code block representation that could either be a string or an iterable list
                     of strings. If a list is provided, it is concatenated into a single string.
        :type a_codeblock: list[str] | str
        :return: A tuple containing the code block type (`str`) and the associated action (`str`) as
                 parsed from the input.
        :rtype: tuple[str, str]
        """
        if isinstance(a_codeblock, list):
            a_codeblock = '\n'.join(map(str, a_codeblock)) # might need to raise a Typeerror here

        # cb_sig = cb_action = None
        cb_sig = a_codeblock.split('\n')[0].strip('`').upper()  # 1st line (w/o ```)
        cb_sig.upper()

        # cb_action = ''
        # if len(a_codeblock.split('\n')) < 2:
        #     return cb_sig, ''
        # for ln in a_codeblock[1:].split('\n'):
        #     if ln.strip() != '':
        #         cb_action = ln.strip()
        #         break
        #
        # if cb_sig == 'BUTTON':
        #     cb_action = ' '.join(cb_action.split(' ')[1:])  # get button text
        # elif cb_sig == 'DATAVIEW':
        #     cb_action = ''.join(cb_action.split(' ')[0])  # get DV Output Format
        #     cb_action = cb_action.upper()

        if cb_sig is None or cb_sig == '':
            cb_sig = 'CodeBlock'

        # if cb_action is None or cb_action == '':
        #     cb_action = 'Undefined'

        return cb_sig

    def extract_codeblocks(self, some_md_content) -> list:
        """
        Extracts all code blocks from the given content within triple backticks (```).

        This function searches for all occurrences of fenced code blocks (enclosed
        by triple backticks) in the provided content.

        :param some_md_content: The string content from which code blocks should be extracted.
        :return: A list of code blocks identified in the provided content.
        :rtype: list
        """
        return re.findall(self.rgx_code_blocks, some_md_content)

    @staticmethod
    def is_subdirectory(child, parent):
        return parent in child.parents

if __name__ == "__main__":
        # v_wb = WbDataDef()
        # t_wb = NewWb()
    DBUG_LVL = 9
        # self.DBUG_LVL = 0  # Do Not print anything
        # self.DBUG_LVL = 1  # print report level actions only (export, load, save, etc.) + all lower levels
        # self.DBUG_LVL = 2  # print object level actions + all lower levels
        # self.DBUG_LVL = 3  # print export_tab + all lower levels
        # self.DBUG_LVL = 4  # print hdr records + all lower levels
        # self.DBUG_LVL = 5  # print detail records + all lower levels
        # self.DBUG_LVL = 9  # print everything (includes export_cell!)
    vc_def = VaultHealthCheck(DBUG_LVL)
    tabs = NewWb(DBUG_LVL)
    exporter = ExcelExporter(DBUG_LVL)
    exporter.export(DBUG_LVL)

        # print(f"v_chk_xl:Loading Spreadsheet: {xl_exec_path} - {xls_pname}")
        # time.sleep(5)
        # pid = Popen([xl_exec_path, xls_pname]).pid

        # shelve_file = shelve.open("v_def.db")
        # shelve_file['v_def'] = v_def
        # shelve_file.close()

    if DBUG_LVL > 90:
        wb_cfg = WbDataDef(DBUG_LVL)
        wb_def = wb_cfg.read_cfg_data()
        cfg     = wb_def.get('cfg', {})
        wb_tabs = wb_def.get('wb_tabs', {})
        wb_data = wb_def.get('wb_data', {})

        lin = "=" * 80
        print(f"\n{lin}")
        # self.tab_def['tab_cd_table_hdr']['Row']
        print(f"Vault Health Check Complete.")
        print(f"\nConfig Sys File: {cfg['cfg_sys_pname']}")
        print(f"     Next Data File: {cfg['cfg_pname']}")
        print(f"            Wb File: {cfg['xls_pname']}")

        # dump wb_def
        dict_list = {
              'cfg': cfg
            , 'wb_tabs': wb_tabs
            , 'wb_data': wb_data
           # , 'tab_def': vc_def.wb_tabs['pros']
        }

        for p_dict_name, p_dict in dict_list.items():
            print(f"\n{p_dict_name}: {lin}")
            for k,v in p_dict.items():
                k_name = f"{p_dict_name}['{k}']"
                print(f"{k_name: <20}: {v}")

        print(f'\nStandalone run of "{Path(__file__).name}" complete.')
