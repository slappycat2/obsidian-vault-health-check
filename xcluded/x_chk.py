from datetime import datetime
import time
import os

import yaml
import re
from pathlib import Path
from subprocess import Popen

from v_chk_xl import ExcelExporter
from v_chk_cfg_data import Config


def upd_pros_dict(pros_dict, ukey, uval, filepath):
    # pros {'tags'  : {tag1: [file1, file2, file3...],
    #                   tag2: [file1, file2, file3...]},
    #       {'links' : {'[[xxx]]': [file1, file2, file3...],
    #                   '[[zzz]]': [file1, file2, file3...],...}
    #       {'dup_vals': {file1: [path1, path2, path3...],},
    #                     file21: [path1, path2, path3...],},...},

    # v2
    # tab_def_data:    (data gathered in v_chk, dumped in spreadsheet)
    # 	pros:
    # 		prop_id:
    # 			prop_val:
    # 			    [files_list]
    # 	tags:
    # 		tag_id:
    # 				[files_list]
    # 	dups:
    # 		file_name:
    # 				[pathnames_list]
    # 	xyml:
    # 		[file_list]
    # 	files:
    # 		filename:
    # 			pros: (including the "tags" property)
    # 	        vals: [vals_list]
    f_list = []

    # Format Data

    if ukey in pros_dict:
        k_dict = pros_dict[ukey]
    else:
        k_dict = {uval: []}
        pros_dict[ukey] = k_dict

    if uval in k_dict:
        f_list = k_dict[uval]

    f_list.append(filepath)
    pros_dict[ukey][uval] = f_list

    return pros_dict

class ObsidianVaultProcessor:
    def __init__(self, cfg):
        print(f"Processing vault: {cfg.dir_vault}")
        ret_val = self.process_vault()
        if ret_val == 0:
            print(f"Metadata extracted successfully from {cfg.dir_vault}")
        else:
            print(f"Error in extracting metadata from {cfg.dir_vault}")

    def process_vault(self):
        v_path_obj = Path(cfg.dir_vault)
        for md_file in v_path_obj.rglob("*.md"):

            # md_file and BASE_DIR are WindowsPath objects, not a strings!
            # To process templates, set BASE_DIR to the templates folder
            test_md_file = str(md_file)
            print(f"md_file: {md_file}  test_md_file: {test_md_file}")
            if cfg.dbug and test_md_file != cfg.dbug:
                # we're debugging a single file and this isn't it!
                # print(f"dbug is on: Skipping_file: {md_file} is not equal to {cfg.dbug}")
                continue

            x_dir_test = False
            for x_dir in md_file.parts:
                if x_dir in cfg.dirs_skip_abs_lst:
                    x_dir_test = True
                    continue  # this only exits this for loop
            if x_dir_test:
                continue

            md_pname = str(md_file)
            print(f"Processing file: {md_file}")
            upd_pros_dict(cfg.pros, 'xkey_dup_files', md_file.name, md_pname, )
            md_file = MarkdownFile(md_pname)

        ret = cfg.write_config(cfg)

        return ret

class MarkdownFile:
    def __init__(self, filepath):
        self.filepath = filepath
        self.file_pros = {}

        print(f"Processing file: {self.filepath}")
        self.parse_file()
        cfg.files[self.filepath] = self.file_pros

    def parse_file(self):
        with open(self.filepath, 'r', encoding='utf-8') as file:
            content = file.read()

        y_text, x_text = self.split_file(content)
        if len(y_text) != 0:
            self.process_yaml(y_text)
        if len(x_text) != 0:
            self.process_body(x_text)

    def upd_val(self, k, v):
        upd_pros_dict(cfg.pros, k, v, self.filepath)
        upd_pros_dict(self.file_pros, k, v, self.filepath)

    def process_body(self, body_text):
        body_text = "".join(body_text)
        # strip code from text
        body_text = re.sub(cfg.rgx_sub_strip_code_blocks, '', body_text)
        body_text = re.sub(cfg.rgx_sub_strip_inline_code, '', body_text)

        match_pros = list(cfg.rgx_body.finditer(body_text))

        # body pros
        for idx, match in enumerate(match_pros):
            m = match_pros[idx].group()
            # m = match_pros[idx].group().replace('\n', "")
            # while len(m.split("::")) > 2:
            #     m =
            #     m = m.replace("::", "::::", 1)
            print(f"m={m}  len: {len(m.split('::'))}")

            k, v = m.split("::")
            k = k.strip()
            v = v.strip()
            if k.startswith("["):
                k = k[2:]
            if v.endswith("]"):
                v = v[:-1]
            # k = k.replace(":", "")

            self.upd_val(k, v)

        # body tags
        body_tags = self.extract_tags(body_text)

        for tag in body_tags:
            if tag.isnumeric():
                continue
            else:
                self.upd_val("tags", tag)

    def process_yaml(self, front_text):
        try:
            data = yaml.safe_load(front_text) or {}
            if not isinstance(data, dict):
                cfg.chk_yaml.append(self.filepath)
                return

            for key, value in data.items():
                key = key.lower()

                if (isinstance(value, list)
                        and isinstance(value[0], list)
                        and len(value) == 1):
                    value = f"[[{value[0][0]}]]"

                if isinstance(value, list):
                    for item in value:
                        self.upd_val(key, item)
                else:
                    # otherwise, the value is a single string, bool, etc.
                    self.upd_val(key, value)

        except yaml.YAMLError as e:
            print(f"Error in YAML: {e}")
            cfg.chk_yaml.append(self.filepath)

        except Exception as e:
            print(f"Unhandled Exception:\t {self.filepath}\n{e}\n\n")
            cfg.chk_yaml.append(self.filepath)


    def split_file(self, content):
        file_text = "".join(content)
        yaml_match = list(cfg.rgx_boundary.finditer(file_text))

        if len(yaml_match) < 2:
            cfg.chk_yaml.append(self.filepath)
            return "", file_text

        start, end = yaml_match[0].end(), yaml_match[1].start()
        yaml_text = file_text[start:end].strip()  # Extract YAML content only
        body_text = file_text[end:]

        return yaml_text, body_text

    def extract_tags(self, content):
        rgx_tag_pattern = re.compile(r'#(\w+)')
        tag_list = set(rgx_tag_pattern.findall(content))
        return tag_list




if __name__ == "__main__":
    cfg = Config()
    cfg.dbug = False
    # cfg.dbug = "10th Step Homework.md"
    # cfg.dbug = "E:\\o2\\⚡ Build 118th Congress Spreadsheet.md"

    processor = ObsidianVaultProcessor(cfg)

    exporter = ExcelExporter(cfg)
    exporter.export(cfg)


    
    print(f"Metadata exported successfully to {cfg.v_chk_pn_wbs}")
    time.sleep(5)

    pid = Popen([cfg.pn_wb_exec, cfg.v_chk_pn_wbs]).pid

