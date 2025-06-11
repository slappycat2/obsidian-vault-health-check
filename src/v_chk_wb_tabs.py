from openpyxl.styles import Side
import copy
from src.v_chk_wb_setup import WbDataDef
from src.v_chk_class_lib import Colors

class NewWb(WbDataDef):
    def __init__(self, dbug_lvl):
        self.DBUG_LVL = dbug_lvl
        self.tab_id = 'init'
        super().__init__(self.DBUG_LVL)
        self.wb_def = self.read_bat_data()
        self.wb_tabs = self.wb_def['wb_tabs']
        self.tab_def = {}
        self.cfg = self.wb_def['cfg']
        self.ctot = self.cfg['ctot']
        self.Colors = Colors()

        if self.DBUG_LVL > 0:
            print("Loading Workbook Tab Definitions...")

        # Notes and hlp_txt s/b a max of 80 chars; Comments, no more than 3 lines.
        self.tab_common = {  # tab color,         tab hdr colors
              'summ': {
                  'tab_name': "Summary"
                , 'shw_grid': False
                , 'tab_titl': 'Obsidian Vault Healthcheck v1.0'
                , 'hdr_clrs': True          #  True=Force Tab Colors; False=Use TableStyle Colors
                , 'col_key1': "Property"
                , 'col_key2': ""
                , 'col_val1': "Values"
                , 'col_val2': ""
                , 'col_lnks': "Links"
                , 'help_txt': {
                #       'subtitle1': [
                #         'Lists all properties and tags found in an Obsidian vault, with '
                #       , 'links back to their respective markdown files. Duplicate Notes,'
                #       , 'existing in different folders, are also detected among various'
                #       , 'other issues and statistics.'
                #       ]
                #       , 'subtitle2': [
                #         'Links back to your vault make it easy to quickly review issues. '
                #       , 'Files with invalid frontmatter/YAML are also listed with links.'
                #       , 'Lastly, if this helps you in any way, PLEASE consider buying'
                #       , "me a coffee via the link, below. It's greatly appreciated! 😉"
                #       ]
                      'notes': [
                           # 123456789 1 2345678 2 2345678 3 2345678 4 2345678 5 2345678 6 2345678 7 2345678 8
                             'The maximum number of links displayed is currently 15 for Values'
                           , 'and 10 for tags. I kept the numbers low, because it can slow'
                           , 'performance. I may make this configurable in a future release,'
                           , 'depending on interest. For now, these numbers work for me. ;-)'
                           , 'Some Totals will appear high because items will be counted twice.'
                           , 'Use them for control purposes or feel free to add your own totals,'
                           , 'if you need something different.'
                           , 'For example, the Files Total on the Properties tab will count one'
                           , 'for every property in every file. So if "My Note" has an "aliases"'
                           , 'property and an "author" property, that file is counted twice'
                    ]
                  }
                , 'data_src': ['obs_props']
                }
            , 'pros': {
                     'tab_name': "Properties"
                    , 'shw_grid': False
                    , 'tab_titl': 'Properties Analysis'
                    , 'hdr_clrs': True  #  True=Force Tab Colors; False=Use TableStyle Colors
                    , 'col_key1': "Properties"
                    , 'col_key2': ""
                    , 'col_val1': "Values"
                    , 'col_val2': ""
                    , 'col_lnks': "Links"
                    , 'help_txt': {
                          'subtitle': [
                            'All Vault Properties with counts of values and note files.'
                        ]
                        , 'notes': [
                              ' - Anything in red bold/italics are likely errors or at least worthy of review.'
                            , ' - All Properties and Tags are listed in lowercase, as that is how Obsidian sees them.'
                            , '     The FileDetails Tab shows them as entered, if lowercase was not used.'
                        ]
                    }
                    , 'data_src': ['obs_props']
                    }
            , 'vals': {
                'tab_name': "Values"
                , 'shw_grid': False
                , 'tab_titl': 'Properties and Values Analysis'
                , 'hdr_clrs': True  # True=Force Tab Colors; False=Use TableStyle Colors
                , 'col_key1': "Properties"
                , 'col_key2': ""
                , 'col_val1': "Values"
                , 'col_val2': ""
                , 'col_lnks': "Links"   # NB: This column title may be referenced in the tab_notes, below!
                , 'help_txt': {
                      'subtitle': [
                        'All Vault Properties and Values with links to their Notes'
                      ]
                    , 'notes': [
                        ' - Anything in red bold/italics are likely errors or at least worthy of review.'
                      , ' - You can use Table Heading Filters to look at specific items.'
                      , '      When filters are applied, tab totals will reflect the filtered data and'
                      , '      a warning will display next to the tab Totals Section'
                      , ' - The "Links" column reflects the total number of links found in all markdown files.'
                      , '   If the columns-maximum is reached, this number will be more than what is displayed.'
                      , ' - All Properties are listed in lowercase, as that is how Obsidian sees them.'
                      , '     The "Files" Tab shows them as entered, if lowercase was not used.'
                      , ' - The Values Total does not include empty values, and therefore may differ from the'
                    , '       the Values Total in the "Properties" Tab.'
                      ]
                    }
                , 'data_src': ['obs_props']
            }
            , 'tags': {
                      'tab_name': "Tags"
                    , 'shw_grid': False
                    , 'tab_titl': 'Tags Analysis'
                    , 'hdr_clrs': True  #  True=Force Tab Colors; False=Use TableStyle Colors
                    , 'col_key1': "Tags"
                    , 'col_key2': ""
                    , 'col_val1': ""
                    , 'col_val2': ""
                    , 'col_lnks': "Links"
                    , 'help_txt': {
                          'subtitle': [
                            'All tags found in all markdown files.'
                        ]
                        , 'notes': [
                            'All tags found in frontmatter and in-line, in all vault markdown files.'
                        ]
                    }
                    , 'data_src': ['obs_atags']
                    }
            , 'file': {
                      'tab_name': "Files"
                    , 'shw_grid': False
                    , 'tab_titl': 'All Files Analysis'
                    , 'hdr_clrs': True  #  True=Force Tab Colors; False=Use TableStyle Colors
                    , 'col_key1': "Notes"
                    , 'col_key2': "Inline?"
                    , 'col_val1': "Properties"
                    , 'col_val2': "CaseDiffs"
                    , 'col_lnks': "ValCount"
                    , 'help_txt': {
                          'subtitle': [
                            'All tags found in all markdown files.'
                        ]
                        , 'notes': [
                            'All tags found in frontmatter and in-line, in all vault markdown files.'
                          , 'Note that "tags" are considered as "Properties" in this report, and the'
                          , 'actual tags themselves will appear in the "Values" column.'
                        ]
                    }
                    , 'data_src': ['obs_files']
                    }
            , 'xyml': {
                      'tab_name': "Xyml"
                    , 'shw_grid': False
                    , 'tab_titl': 'Possible Issues and Corrupt* Markdown Files'
                    , 'hdr_clrs': True  #  True=Force Tab Colors; False=Use TableStyle Colors
                    , 'col_key1': "Notes"
                    , 'col_key2': ""
                    , 'col_val1': "Fm Okay"
                    , 'col_val2': "Likely Issue"
                    , 'col_lnks': ""
                        # Possible errors needing more detail...
                        # These are defined above...
                        #    'NonD': 'YAML not in dictionary format.'
                        #  , 'BadY': "YAML loaded, but can't be unpacked."
                        # ,  'ErrY': "Invalid Properties--Can't Load Frontmatter."
                        #  , 'NoFm': 'No YAML defined--Ok, if intentional.'
                    , 'help_txt': {
                          'subtitle': [
                            'These are markdown files with frontmatter that may need to be reviewed.'
                            ]
                        , 'notes': [
                               '* Corrupt, meaning the Python package PyYAML (6.0.2) is unable to perform a'
                             , '  "safe_load" of the notes frontmatter, or the YAML is otherwise problematic.'
                               ' - Note that the terms "frontmatter" and "YAML" are interchangeable here.'
                             , ' - Items on this report may or may not be indicative of an error, depending '
                             , '    on how you setup and use your vault. At the very least, they should be worthy'
                             , '    of review.'
                             , 'Fm Okay - Frontmatter loaded OK and was parsed w/o error, however properties'
                             , '    and/or tags were found inline that caused a failure.'
                             , ''
                             , ''
                            ]
                        }
                    , 'data_src': ['obs_xyaml']
                    }
            , 'dups': {
                      'tab_name': "Duplicates"
                    , 'shw_grid': False
                    , 'tab_titl': 'Duplicate Files in Vault'
                    , 'hdr_clrs': True  #  True=Force Tab Colors; False=Use TableStyle Colors
                    , 'col_key1': "Duplicate Notes"
                    , 'col_key2': ""
                    , 'col_val1': ""
                    , 'col_val2': ""
                    , 'col_lnks': "Links"
                    , 'help_txt': {
                          'subtitle': [
                             'Duplicates here are defined as vault files that have identical filenames,'
                            , 'but exist in different folders.'
                            ]
                        , 'notes': [
                              'Duplicate markdown filenames are allowed in Obsidian, but should be'
                            , 'avoided as they can be a source of confusion. They also require fully'
                            , 'qualified pathnames, instead of just the filename, when attempting to'
                            , 'create links.'
                            , 'For this reason, the links provided show the full pathnames.'
                        ]
                    }
                    , 'data_src': ['obs_dupfn']
                    }
            , 'nest': {
                      'tab_name': "Nests"
                    , 'shw_grid': False
                    , 'tab_titl': 'Nested Dictionary Properties Analysis'
                    , 'hdr_clrs': True  #  True=Force Tab Colors; False=Use TableStyle Colors
                    , 'col_key1': "Plug-ins"
                    , 'col_key2': "Filenames"
                    , 'col_val1': "Values"
                    , 'col_val2': ""
                    , 'col_lnks': "All Values ('|' separator)"
                    , 'help_txt': {
                          'subtitle': [
                              'This is only a Proof-Of-Concept at this time. It may be'
                            , "useful, for some debugging, but I wouldn't spend a lot"
                            , "of time in these weeds. Unless perhaps there's no game on,"
                            , "or the boss is out playing golf, then freak away!"
                            ]
                        , 'notes': [
                              '                -- Proof-Of-Concept -- '
                            , 'The data here is from markdown files in your vault that'
                            , 'contain Nested Properties in their YAML. This is something'
                            , 'Obsidian does not support, natively. So, when I hit one, '
                            , 'I thought they s/b included in this report for documentation'
                            , "purposes. These are files created by Community Plugins and"
                            , "they almost certainly do not require any action on your part."
                            , "I wish I could be one of those people who can just enjoy"
                            , "their soup, without having to stop and figure out how the"
                            , "spoon works."
                            , 'DO NOT ATTEMPT TO EDIT THESE FILES UNLESS YOU'
                            , "REALLY KNOW WHAT YOU'RE DOING!'"
                            , "I certainly don't know what I'm doing, so don't be looking'"
                            , "down the table at me for help on this one!"
                            , "Just use the Plugin's interface and keep your head about"
                            , "you! CARRY ON!!! ;-)"
                        ]
                    }
                    , 'data_src': ['obs_nests']
                    }
            , 'plug': {
                      'tab_name': "Plugins"
                    , 'shw_grid': False
                    , 'tab_titl': 'Installed Plugins'
                    , 'hdr_clrs': True  #  True=Force Tab Colors; False=Use TableStyle Colors
                    , 'col_key1': "Plugin Id"
                    , 'col_key2': "Enabled"
                    , 'col_val1': "isDesktopOnly"
                    , 'col_val2': ""
                    , 'col_lnks': ""
                    , 'help_txt': {
                          'subtitle': [
                            'Installed Plugins manifests.'
                          ]
                        , 'notes': [
                              'These are the plugins that exist in the Vault Plugins directory'
                            , "Disabled Plugins are listed in red italics."
                            ]
                    }
                    , 'data_src': ['obs_plugs']
                    }
            , 'code': {
                'tab_name': "Code"
                , 'shw_grid': False
                , 'tab_titl': 'Code Blocks Analysis'
                , 'hdr_clrs': True  # True=Force Tab Colors; False=Use TableStyle Colors
                , 'col_key1': "Notes"
                , 'col_key2': "Plugins"
                , 'col_val1': "Signature"
                , 'col_val2': ""
                , 'col_lnks': "Count"  # JS or ''
                , 'help_txt': {
                    'subtitle': [
                        'All codeblocks and with known plugins signatures.'
                    ]
                    , 'notes': [
                          '-- This is only a Proof-Of-Concept at this time --'
                        , "- To best view Codeblocks, expand the height of the"
                        , "  Formula Bar to something like 15 or 25 lines,"
                        , "  then run the cursor over listed Codeblocks."
                    ]
                }
                , 'data_src': ['obs_codes']
            }
            , 'tmpl': {
                      'tab_name': "Templates"
                    , 'shw_grid': False
                    , 'tab_titl': 'Templates Analysis'
                    , 'hdr_clrs': True  #  True=Force Tab Colors; False=Use TableStyle Colors
                    # , 'col_key1': "Template Filename"
                    # , 'col_lnks': "Values"
                    # , 'col_val1': "Property"
                    , 'col_key1': "Property"
                    , 'col_key2': ""
                    , 'col_val1': "Values"
                    , 'col_val2': ""
                    , 'col_lnks': "Links"
                    , 'help_txt': {
                          'subtitle': []
                        , 'notes': []
                    }
                    , 'data_src': ['obs_tmplt']
                    }
            , 'ar51': {
                'tab_name': "Area51"
                , 'shw_grid': False
                , 'tab_titl': ''
                , 'hdr_clrs': True  # True=Force Tab Colors; False=Use TableStyle Colors
                , 'col_key1': "Properties"
                , 'col_key2': ""
                , 'col_val1': "Values"
                , 'col_val2': ""
                , 'col_lnks': "Links"
                , 'help_txt': {
                      'subtitle': []
                    , 'notes': []
                }
                , 'data_src': ['obs_props']
            }

            , 'init': {
                      'tab_name': ""
                    , 'col_key1': "Init only"
                    , 'col_val1': "dummy"
                    , 'col_val2': ""
                    , 'col_lnks': "dummy"
                    , 'help_txt': {}
                    , 'data_src': ['dummy']
                    }
        }

        for tab_id_key in self.wb_tabs.keys():
            if self.DBUG_LVL > 3:
                print(f"Building '{tab_id_key}' tab definition")
            self.tab_def = self.wb_tabs[tab_id_key]
            if tab_id_key == 'pros':
                self.tab_def_obj = DefPros(self)
            elif tab_id_key == 'vals':
                self.tab_def_obj = DefVals(self)
            elif tab_id_key == 'tags':
                self.tab_def_obj = DefTags(self)
            elif tab_id_key == 'dups':
                self.tab_def_obj = DefDups(self)
            elif tab_id_key == 'xyml':
                self.tab_def_obj = DefXyml(self)
            elif tab_id_key == 'file':
                self.tab_def_obj = DefFile(self)
            elif tab_id_key == 'tmpl':
                self.tab_def_obj = DefTmpl(self)
            elif tab_id_key == 'code':
                self.tab_def_obj = DefCode(self)
            elif tab_id_key == 'nest':
                self.tab_def_obj = DefNest(self)
            elif tab_id_key == 'plug':
                self.tab_def_obj = DefPlug(self)
            elif tab_id_key == 'summ':
                self.tab_def_obj = DefSumm(self)
            elif tab_id_key == 'ar51':
                self.tab_def_obj = DefAr51(self)
            elif tab_id_key == 'init':
                pass
            else:
                raise Exception(f"Unexpected key: {tab_id_key} in {self.wb_tabs.keys()}")

            if tab_id_key != 'init':
                self.wb_def['wb_tabs'][tab_id_key] = self.tab_def_obj.tab_def

        self.write_bat_data()

class NewTab:
    def __init__(self, tab_id, wb_obj):
        self.tab_id = tab_id
        self.cfg = wb_obj.cfg
        # self.tab_common = tab_common

        self.tab_name       = wb_obj.tab_common[tab_id]['tab_name']
        self.hdr_clrs       = wb_obj.tab_common[tab_id]['hdr_clrs']
        self.col_key1       = wb_obj.tab_common[tab_id]['col_key1']
        self.col_key2       = wb_obj.tab_common[tab_id]['col_key2']
        self.col_lnks       = wb_obj.tab_common[tab_id]['col_lnks']
        self.col_val1       = wb_obj.tab_common[tab_id]['col_val1']
        self.col_val2       = wb_obj.tab_common[tab_id]['col_val2']
        self.help_txt       = wb_obj.tab_common[tab_id]['help_txt']
        self.data_src       = wb_obj.tab_common[tab_id]['data_src']
        self.tab_title      = wb_obj.tab_common[tab_id]['tab_titl']
        self.showGridLines  = wb_obj.tab_common[tab_id]['shw_grid']
        self.xyml_descs     = wb_obj.xyml_descs

        ctot = self.cfg.get('ctot', [0] * 13)

        self.link_max_vals   = ctot[11]
        self.link_max_tags   = ctot[12]
        self.link_lim_vals = self.cfg.get('link_lim_vals', 0)
        self.link_lim_tags = self.cfg.get('link_lim_tags', 0)

        # Fill and text colors for grid tab headings
        self.colors = wb_obj.Colors
         # self.tab_clrs = TabColorsDef()

        # Border Styles must be one of {'medium', 'mediumDashDotDot', 'hair', 'thin',
        # 'dashed', 'double', 'slantDashDot', 'dashDotDot', 'dashDot', 'dotted',
        # 'mediumDashDot', 'thick', 'mediumDashed'}

        self.bdr_thick = Side(border_style="thick", color='000000')
        self.bdr_thin = Side(border_style="thin", color='000000')
        self.bdr_double = Side(border_style="double", color='000000')

        self.hdr_PVI = "P-V Index"
        self.hdr_IsVis = "IsVisible"
        self.hdr_links_pfx = "File"

        self.tab_clr_txt = self.colors.get_clr("blk", 0)
        self.tab_txt_sz = 11
        self.tab_link_clr = self.colors.get_clr("blu", 0)
        self.tab_fill_clr = ''
        self.font_title_lst = ['Berlin Sans FB Demi', 24, '']
        self.font_subs_lst = ['Berlin Sans', 14, '']
        self.font_body_lst = ['Calibri'    , 11, '']
        self.cell_width = 8
        self.tbl_name = f"tbl_{self.tab_id}"
        self.tbl_hdr_row = 10
        self.tbl_beg_col = 4
        self.tbl_end_col = 0
        self.tbl_fix_cols = 0
        self.hdr_RowId = "RowId"
        self.f_warn_null = "Properties w/Empty Values Detected!"
        self.f_warn_filter = "Column filters applied--Totals now reflect column filters!"
        # Aggregate 1st Digit: 9=Num, 3=Text       (SUM, COUNTA)
        # Aggregate 2nd Digit: 3=Vis, 4=All        (3=ignore hidden rows)

        self.f_txt_rows   = f'=_xlfn.AGGREGATE(3,3,{self.tbl_name}[{self.hdr_RowId}])'
        self.f_uniq_key1   = f'=COUNTA(_xlfn.UNIQUE(_xlfn.FILTER({self.tbl_name}[{self.col_key1}],{self.tbl_name}[{self.hdr_IsVis}])))'
        self.f_uniq_key2   = f'=COUNTA(_xlfn.UNIQUE(_xlfn.FILTER({self.tbl_name}[{self.col_key2}],{self.tbl_name}[{self.hdr_IsVis}])))'
        self.f_uniq_val1   = f'=COUNTA(_xlfn.UNIQUE(_xlfn.FILTER({self.tbl_name}[{self.col_val1}],{self.tbl_name}[{self.hdr_IsVis}])))'
        self.f_uniq_val2   = f'=COUNTA(_xlfn.UNIQUE(_xlfn.FILTER({self.tbl_name}[{self.col_val2}],{self.tbl_name}[{self.hdr_IsVis}])))'
        self.f_uniq_lnks   = f'=COUNTA(_xlfn.UNIQUE(_xlfn.FILTER({self.tbl_name}[{self.col_lnks}],{self.tbl_name}[{self.hdr_IsVis}])))'
        self.f_num_key1   = f'=_xlfn.AGGREGATE(9,3,{self.tbl_name}[{self.col_key1}])'
        self.f_num_key2   = f'=_xlfn.AGGREGATE(9,3,{self.tbl_name}[{self.col_key2}])'
        self.f_num_val1   = f'=_xlfn.AGGREGATE(9,3,{self.tbl_name}[{self.col_val1}])'
        self.f_num_val2   = f'=_xlfn.AGGREGATE(9,3,{self.tbl_name}[{self.col_val2}])'
        self.f_num_lnks   = f'=_xlfn.AGGREGATE(9,3,{self.tbl_name}[{self.col_lnks}])'
        self.f_txt_key1   = f'=_xlfn.AGGREGATE(3,3,{self.tbl_name}[{self.col_key1}])'
        self.f_txt_key2   = f'=_xlfn.AGGREGATE(3,3,{self.tbl_name}[{self.col_key2}])'
        self.f_txt_val1   = f'=_xlfn.AGGREGATE(3,3,{self.tbl_name}[{self.col_val1}])'
        self.f_txt_val2   = f'=_xlfn.AGGREGATE(3,3,{self.tbl_name}[{self.col_val2}])'
        self.f_txt_lnks   = f'=_xlfn.AGGREGATE(3,3,{self.tbl_name}[{self.col_lnks}])'
        self.f_cif_val1_true   = f'=COUNTIF({self.tbl_name}[{self.col_val1}],True)'


        self.f_null_values = f'=IFERROR(IF(_xlfn.AGGREGATE(3,3,{self.tbl_name}[{self.col_val1}])<>SUM({self.tbl_name}[{self.hdr_IsVis}]),"{self.f_warn_null}",""),"")'
        self.f_filters_on = f'=IFERROR(IF(COUNTA({self.tbl_name}[{self.hdr_RowId}])<>SUM({self.tbl_name}[{self.hdr_IsVis}]),"{self.f_warn_filter}",""),"")'
        self.f_isVisible = f'=SUBTOTAL(3,@[{self.hdr_RowId}])'
        sz = 11

        self.tab_table_files = {}
        self.tab_def = {
                      'tab_id': self.tab_id
                    , 'tab_clr_txt':        self.tab_clr_txt
                    , 'hdr_clrs':           self.hdr_clrs
                    , 'tab_txt_sz':         self.tab_txt_sz
                    , 'tab_link_clr':       self.tab_link_clr
                    , 'tab_fill_clr':       self.tab_fill_clr
                    , 'font_title_lst':     self.font_title_lst # not used
                    , 'font_subs_lst':      self.font_subs_lst  # not used
                    , 'font_body_lst':      self.font_body_lst   # not used
                    , 'cell_width':         self.cell_width
                    , 'tbl_name':           self.tbl_name
                    , 'tbl_hdr_row':        self.tbl_hdr_row
                    , 'tbl_beg_col':        self.tbl_beg_col
                    , 'tbl_end_col':        self.tbl_end_col
                    , 'tbl_fix_cols':       self.tbl_fix_cols
                    , 'hdr_RowId':          self.hdr_RowId
                    , 'hdr_PVI':            self.hdr_PVI
                    , 'hdr_IsVis':          self.hdr_IsVis
                    , 'hdr_links_pfx':      self.hdr_links_pfx
                    , 'f_warn_null':        self.f_warn_null
                    , 'f_warn_filter':      self.f_warn_filter
                    , 'f_uniq_key1':        self.f_uniq_key1
                    , 'f_uniq_val1':        self.f_uniq_val1
                    , 'f_txt_key1':         self.f_txt_key1
                    , 'f_txt_key2':         self.f_txt_key2
                    , 'f_txt_val1':         self.f_txt_val1
                    , 'f_txt_val2':         self.f_txt_val2
                    , 'f_num_lnks':         self.f_num_lnks
                    , 'f_null_values':      self.f_null_values
                    , 'f_filters_on':       self.f_filters_on
                    , 'f_isVisible':        self.f_isVisible
                    , 'tab_name':           self.tab_name
                    , 'data_src':           self.data_src
                    , 'tab_help_txt':       self.help_txt
                                     # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
                    , 'tab_cd_title_def':    [3,  2, 'Berlin Sans FB Demi', 24, 0, '', '', True, False, 'left', self.tab_title]
                    , 'tab_cd_subtitle_def': [3,  3, '', sz, 0, '', '', True, False, 'left', '']
                    , 'tab_cd_notes_def':    [3, 22, '', sz, 0, '', '', True, False, 'left', '']
                    , 'tab_color': ''
                    # cd is short for cell_def
                    , 'tab_cd_table_hdr': {}  # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
                    , 'tab_cd_table_dtl': {}
                    , "tab_cd_table_links":  [9, 0, '', sz, 18, '', '', False, False, 'left'  , '']
                    , 'tab_table_links_cols': 0
                    , 'tab_table_link_spcrs': True

                    , 'tab_has_isVisible_col': False
                    , 'tab_tots_isVisible_col': 0
                    , 'tab_cd_fixed_grid': {}  # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
        } # end tab_def

    def __getitem__(self, key):
        return self.tab_def[key]

    def tab_def_post(self):
        self.calc_col_pointers()
        self.files_cols_def()
        
    @staticmethod
    def get_links_cols(lnks_limit, lnks_max):
        if lnks_limit == 0:
            return lnks_max
        return min(lnks_limit, lnks_max)
        
    def files_cols_def(self):
        # If there are add`l columns for Files "Where used", add them to the tab_def
        if self.tab_id and self.tab_def['tab_table_links_cols'] > 0:
            tab_cd_table_hdr = self.tab_def['tab_cd_table_hdr']
            tab_cd_table_dtl = self.tab_def['tab_cd_table_dtl']
            tab_cd_table_hdr, tab_cd_table_dtl = self.set_table_links(tab_cd_table_hdr, tab_cd_table_dtl)
            self.tab_def['tab_cd_table_hdr'] = tab_cd_table_hdr
            self.tab_def['tab_cd_table_dtl'] = tab_cd_table_dtl

    def calc_col_pointers(self):
        """
        Calculates the pointer values to define the start, end, and fixed column positions
        for table definitions in the `tab_def` dictionary attribute.

        This method determines various parameters related to table column placements,
        including:
          tab_def['tbl_beg_col'] - the first absolute column in the table
          tab_def['tbl_end_col'] - the last absolute column in the table
          tab_def['tbl_fix_cols'] - the number of fixed columns in the table
            this is the size of tab_def['tab_cd_table_hdr'], it's the tbl sz w/o
            the links columns
          tab_def['tab_has_isVisible_col'] - this will be set if to the last column
            if tab_def['tab_has_isVisible_col'] is True

        :param self: Instance of the containing class that holds the `tab_def` attribute.
        :type self: Any
        :return: None; tab_def is updated with the calculated values.
        """
        first_col_list = []
        link_cols = self.tab_def['tab_table_links_cols']
        spcr_cols = 0

        if self.tab_def['tab_table_link_spcrs']:
            spcr_cols = link_cols

        col_num = 0
        for k, v in self.tab_def['tab_cd_table_hdr'].items():
            col_num = int(v[0])
            if col_num != 0:
                first_col_list.append(col_num)
        if len(first_col_list) == 0:
            raise ValueError("No columns defined/all columns zero in tab_cd_table_hdr")

        self.tab_def['tbl_beg_col'] = min(first_col_list)
        self.tab_def['tbl_fix_cols'] = len(self.tab_def['tab_cd_table_hdr'])
        self.tab_def['tbl_end_col'] = ((link_cols + spcr_cols)
            + (self.tab_def['tbl_fix_cols'] - 1)
            + (self.tab_def['tbl_beg_col']))
        if self.tab_def['tbl_end_col'] < self.tab_def['tab_tots_isVisible_col']:
            self.tab_def['tbl_end_col'] = self.tab_def['tab_tots_isVisible_col']

        if self.tab_def['tab_has_isVisible_col']:
            self.tab_def['tab_tots_isVisible_col'] = int(self.tab_def['tbl_end_col'])
            self.tab_def['tab_cd_fixed_grid']['isVisible'][0] = self.tab_def['tbl_end_col']

    def set_table_links(self, tab_cd_table_hdr, tab_cd_table_dtl):
        tab_table_links_hdr     = self.tab_def['tab_cd_table_links']
        tab_table_links_dtl     = copy.deepcopy(self.tab_def['tab_cd_table_links'])
        tab_cd_table_spacer_hdr = self.tab_def['tab_cd_table_spacer']
        tab_cd_table_spacer_dtl = copy.deepcopy(self.tab_def['tab_cd_table_spacer'])
        tab_table_links_dtl[5]     = tab_table_links_dtl[6]     = ''  # No Color fills in details!
        tab_cd_table_spacer_dtl[5] = tab_cd_table_spacer_dtl[6] = ''  # No Color fills in details!
        tab_table_link_spcrs = self.tab_def['tab_table_link_spcrs']
        tab_table_links_cols = self.tab_def['tab_table_links_cols']
        hdr_links_pfx       = self.tab_def['hdr_links_pfx']

        col_idx = tab_table_links_hdr[0]
        for i in range(1, tab_table_links_cols + 1):                           # [col,    w, txt, fill, bold, align
            # tab_cd_table_links[0] = col_idx
            # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
            link_key_and_val = f"{hdr_links_pfx}{i:02d}"
            spcr_key_and_val = f"S{i}"
            tab_cd_table_hdr[link_key_and_val] = [col_idx] + tab_table_links_hdr[1:10] + [link_key_and_val]
            tab_cd_table_dtl[link_key_and_val] = [col_idx] + tab_table_links_dtl[1:]
            col_idx += 1

            if tab_table_link_spcrs:
                tab_cd_table_hdr[spcr_key_and_val] = [col_idx] + tab_cd_table_spacer_hdr[1:10] + [f"{spcr_key_and_val}  "]
                tab_cd_table_dtl[spcr_key_and_val] = [col_idx] + tab_cd_table_spacer_dtl[1:]
                col_idx += 1

        return tab_cd_table_hdr, tab_cd_table_dtl

class DefPros(NewTab):
    def __init__(self,wb_obj):
        self.tab_id = 'pros'
        self.tab_common = wb_obj.tab_common
        super().__init__(self.tab_id, wb_obj)

        # hdr_clrs = True will turn on Table hdrs
        # clr1, txt1, clr2, txt2, table_style = self.color.get_tab_clrs(self.tab_id)
        # clr1, clr2,  clr1,  txt1,  clr2, txt2,    tab_table_style = \
        #    self.colors.get_tab_clrs(self, tab_id, shade=None, row_style=None)
        clr1, txt1, clr2, txt2, table_style = self.colors.get_tab_clrs(self.tab_id)
        self.tab_def['tab_table_style'] = table_style

        # clr1 = tab color, and main tab headings
        # clr2 = secondary "highlights" color, headings
        # clr1 = fill color on cells that use color fills
        # txt1 = text color        ''         ''
        # clr2 = fill color on Table Headings (if turned on w/hdr_clrs
        # txt2 = text color        ''         ''

        self.tab_def['tab_table_style'] = table_style

        sz = self.tab_txt_sz

        self.font_title_lst = ['Berlin Sans FB Demi', 24, clr1]
        self.font_subs_lst = ['Berlin Sans', 14, txt1]
        self.font_body_lst = ['Calibri', sz, txt1]
        self.tab_def['tab_color'] = clr1

        self.tab_def['tab_table_link_spcrs'] = True  # Always, TRUE for now
        self.tab_def['tab_txt_sz'] = sz
        self.tab_def['showGridLines'] = self.showGridLines

        self.tab_def['hdr_links_pfx'] = ""
        self.tab_def['tab_table_links_cols'] = 0
        self.tab_def['tab_has_isVisible_col'] = True
        self.tab_def['tab_tots_isVisible_col'] = 14

        # if self.tab_def['tab_name'] != 'Properties':
        #     raise WorkbookDefinitionError(f"Tab_Def: pros-tab_name tab name not defined.")
        # self.tab_def['tab_name'] = 'Properties'
        self.tab_def['tab_cd_title_def'] = [3, 2, 'Berlin Sans FB Demi', 24, 0, clr1, '', True, False, 'left',self.tab_title]
        self.tab_def['tab_cd_subtitle_def'] = [ 3,  3, '', sz, 0, '', '', False,  False, 'left', '']
        self.tab_def['tab_cd_notes_def']    = [ 3, 12, '', sz, 0, '', '', False,  False, 'left', '']
        self.tab_def['tab_help_txt'] = self.help_txt

        self.tab_def['tab_cd_table_hdr'] = {
                              # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
              "RowId":          [10, 10, '', sz, 12, txt1, clr1, True, False, 'center', self.hdr_RowId]
            , "Prop":           [11, 10, '', sz, 27, txt1, clr1, True, False, 'left'  , self.col_key1]
            , "Values Count":   [12, 10, '', sz, 12, txt1, clr1, True, False, 'center', self.col_val1]
            , "File Count":     [13, 10, '', sz, 12, txt1, clr1, True, False, 'center', self.col_lnks]
        }
        self.tab_def['tab_cd_table_dtl'] = {  # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
              "RowId":          [10,  0, '', sz,  0, "", "", False, False, 'center', self.hdr_RowId]
            , "Prop":           [11,  0, '', sz,  0, "", "", True,  False, 'left', self.col_key1]
            , "Values Count":   [12,  0, '', sz,  0, "", "", False, False, 'center', self.col_val1]
            , "File Count":     [13,  0, '', sz,  0, "", "", False, False, 'center', self.col_lnks]
        }
        self.tab_def['tab_cd_table_links']     = [0, 0, '', 0, 0, txt1, clr1, False, False, '', '']
        self.tab_def['tab_cd_table_spacer'] = [0, 0, '', 0, 0, txt1, clr1, False, False, 'right', '']
        self.tab_def['tab_cd_fixed_summ']   = {  # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
            # totals headers (across and down)
              'summ-title':  [3, 5, '', 14, 21, txt1, clr1, True, False, 'left', 'Analysis']
            , 'Totals':      [0, 0, '', sz, 15, txt1, clr1, True, False, 'center', 'Totals']
            , 'Rows':        [3, 6, '', sz,  0, txt2, clr2, True, False, 'right', 'Rows']
            , 'x-uniq-rows': [0, 0, '', sz,  0, "", "", False, False, 'center', self.f_txt_rows]
            , 'Prop':        [3, 7, '', sz,  0, txt2, clr2, True, False, 'right', self.col_key1]
            , 'x-ctot-key1': [0, 0, '', sz,  0, "", "", False, False, 'center', self.f_txt_key1]
            , 'Val':         [3, 8, '', sz,  0, txt2, clr2, True, False, 'right', self.col_val1]
            , 'x-uniq-val1': [0, 0, '', sz,  0, "", "", False, False, 'center', self.f_num_val1]
            , 'Files':       [3, 9, '', sz,  0, txt2, clr2, True, False, 'right', self.col_lnks]
            , 'x-ctot-val1': [0, 0, '', sz,  0, "", "", False, False, 'center', self.f_num_lnks]
        }
        self.tab_def['tab_cd_fixed_grid']   = {  # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
            # totals headers (across and down)
              'Notes-hdr':   [3, 11, '', sz, 0, txt1, clr1, True, False, 'left', 'Notes: ']
            , 'x-filters-on': [ 5,  6, '', 12, 0, self.colors.clr_red, "", True, True, 'left', self.f_filters_on]
            , 'x-null-values': [11, 8, '', 12, 0, self.colors.clr_red, "", True, True, 'left', self.f_null_values]
            # This one is for the IVisible Column, not the totals
            , 'isVisible':    [14, 0, '', sz, 1, clr2, clr2, False, False, 'right', self.f_isVisible]
        }

        self.tab_def_post()

class DefVals(NewTab):
    def __init__(self,wb_obj):
        self.tab_id = 'vals'
        self.tab_common = wb_obj.tab_common
        super().__init__(self.tab_id, wb_obj)

        # hdr_clrs = True will turn on Table hdrs
        # clr1, txt1, clr2, txt2, table_style = self.color.get_tab_clrs(self.tab_id)
        # clr1, clr2,  clr1,  txt1,  clr2, txt2,    tab_table_style = \
        #    self.colors.get_tab_clrs(self, tab_id, shade=None, row_style=None)
        clr1, txt1, clr2, txt2, table_style = self.colors.get_tab_clrs(self.tab_id)
        self.tab_def['tab_table_style'] = table_style

        # clr1 = tab color, and main tab headings
        # clr2 = secondary "highlights" color, headings
        # clr1 = fill color on cells that use color fills
        # txt1 = text color        ''         ''
        # clr2 = fill color on Table Headings (if turned on w/hdr_clrs
        # txt2 = text color        ''         ''

        self.tab_def['tab_table_style'] = table_style

        sz = self.tab_txt_sz

        self.font_title_lst = ['Berlin Sans FB Demi', 24, clr1]
        self.font_subs_lst = ['Berlin Sans', 14, txt1]
        self.font_body_lst = ['Calibri', sz, txt1]
        self.tab_def['tab_color'] = clr1

        self.tab_def['tab_table_link_spcrs'] = True  # Always, TRUE for now
        self.tab_def['tab_txt_sz'] = sz
        self.tab_def['showGridLines'] = self.showGridLines

        max_vals = self.link_max_vals
        lim_vals = self.link_lim_vals
        links_cols = self.get_links_cols(lim_vals, max_vals)
        self.tab_def['tab_table_links_cols'] = links_cols

        self.tab_def['hdr_links_pfx'] = "File"
        self.tab_def['tab_has_isVisible_col'] = True
        # self.tab_def['tab_tots_isVisible_col'] = 44

        # if self.tab_def['tab_name'] != 'Properties':
        #     raise WorkbookDefinitionError(f"Tab_Def: pros-tab_name tab name not defined.")
        # self.tab_def['tab_name'] = 'Properties'
        self.tab_def['tab_cd_title_def'] = [ 3,  2, 'Berlin Sans FB Demi', 24, 0, clr1, '', True, False, 'left',self.tab_title]
        self.tab_def['tab_cd_subtitle_def'] = [ 3,  3, '', sz, 0, '', '', False,  False, 'left', '']
        self.tab_def['tab_cd_notes_def']    = [ 3, 12, '', sz, 0, '', '', False,  False, 'left', '']
        self.tab_def['tab_help_txt'] = self.help_txt

        self.tab_def['tab_cd_table_hdr'] = {
            # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
             "RowId":  [10, 10, '', sz,  8, txt1, clr1, True, False, 'center', self.hdr_RowId]
           , "Prop":   [ 0,  0, '', sz, 27, txt1, clr1, True, False, 'left'  , self.col_key1]
           , "Val":    [ 0,  0, '', sz, 50, txt1, clr1, True, False, 'left'  , self.col_val1]
           , "FCnt":   [ 0,  0, '', sz,  8, txt1, clr1, True, False, 'center', self.col_lnks]
           , "PVI":    [ 0,  0, '', sz, 13, txt1, clr1, True, False, 'center', self.hdr_PVI]
           }
        self.tab_def['tab_cd_table_dtl'] =  {
            # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
             "RowId":  [10, 0, '', sz, 0, "", "", False, False, 'center', '']
           , "Prop":   [0,  0, '', sz, 0, "", "", True, False,  'left', '']
           , "Vals":   [0,  0, '', sz, 0, "", "", False, False, 'left', '']
           , "FCnt":   [0,  0, '', sz, 0, "", "", False, False, 'center', '']
           , "PVI":    [0,  0, '', sz, 0, "", "", False, False, 'center', '']
           }
        self.tab_def['tab_cd_table_links']  = [15, 0, '', sz, 18, txt1, clr1, False, False, 'left'  , '']
        self.tab_def['tab_cd_table_spacer'] = [16, 0, '', sz,  1, txt1, clr1, False, False, 'right'  , '']
        self.tab_def['tab_cd_fixed_summ']   = {
            # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
            # totals headers (across and down)
             'summ-title':    [ 3, 5, '', 14, 21, txt1, clr1, True,  False, 'left',    'Analysis']
          ,  'Totals':        [ 0, 0, '', sz, 15, txt1, clr1, True,  False, 'center',  'Totals']
          ,  'Rows':          [ 3, 6, '', sz, 0, txt2, clr2, True,  False, 'right',  "Rows"]
          ,  'x-uniq-rows':   [ 0, 0, '', sz, 0, "",     "", False, False, 'center', self.f_txt_rows]
          ,  'Prop':          [ 3, 7, '', sz, 0, txt2, clr2,  True,  False, 'right', self.col_key1]
          ,  'x-ctot-key1':   [ 0, 0, '', sz, 0, "",     "",  False, False, 'center', self.f_uniq_key1]
          ,  'Values':        [ 3, 8, '', sz, 0, txt2, clr2,  True,  False, 'right', self.col_val1]
          ,  'x-uniq-val1':   [ 0, 0, '', sz, 0, "",     "",  False, False, 'center', self.f_txt_val1]
          ,  'Files':         [ 3, 9, '', sz, 0, txt2, clr2,  True,  False, 'right', self.col_lnks]
          ,  'x-ctot-val1':   [ 0, 0, '', sz, 0, "",     "",  False, False, 'center', self.f_num_lnks]
        }
        self.tab_def['tab_cd_fixed_grid']   = {
            # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
            # totals headers (across and down)
            'Notes-hdr':      [ 3, 11, '', sz, 0, txt1, clr1, True, False, 'left', 'Notes: ']
          , 'x-null-values':  [ 5,  8, '', 12, 0, self.colors.clr_red,  "", True, True,   'left',   self.f_null_values]
          , 'x-filters-on':   [ 5,  6, '', 12, 0, self.colors.clr_red,  "", True, True,   'left',   self.f_filters_on]
          # This one is for the IVisible Column, not the totals
          # The 'isVisible' column, below, is set in calc_col_pointers()
          , 'isVisible':      [0, 0, '', sz, 0, clr2, clr2, False, False, 'right',  self.f_isVisible]
          }

        self.tab_def_post()

class DefTags(NewTab):
    def __init__(self,wb_obj):
        self.tab_id = 'tags'
        self.tab_common = wb_obj.tab_common
        super().__init__(self.tab_id, wb_obj)

        clr1, txt1, clr2, txt2, table_style = self.colors.get_tab_clrs(self.tab_id)
        self.tab_def['tab_table_style'] = table_style
        # clr1 = tab color,
        # clr2 = secondary "highlights" color, headings
        # clr1 = fill color on cells that use color fills
        # txt1 = text color on cells that use color fills

        sz = self.tab_txt_sz

        self.font_title_lst = ['Berlin Sans FB Demi', 24, clr1]
        self.font_subs_lst = ['Berlin Sans', 14, txt1]
        self.font_body_lst = ['Calibri', sz, txt1]
        self.tab_def['tab_color'] = clr1

        self.tab_def['tab_table_link_spcrs'] = True  # Always, TRUE for now
        self.tab_def['tab_txt_sz'] = sz
        self.tab_def['showGridLines'] = self.showGridLines

        max_tags = self.link_max_tags
        lim_tags = self.link_lim_tags
        links_cols = self.get_links_cols(lim_tags, max_tags)
        self.tab_def['tab_table_links_cols'] = links_cols

        self.tab_def['hdr_links_pfx'] = "File"
        self.tab_def['tab_has_isVisible_col'] = True
        # self.tab_def['tab_tots_isVisible_col'] = 32
        # self.tab_def['tab_table_links_cols'] = 10

        # self.tab_def['tab_name'] = 'Tags'
        self.tab_def['tab_cd_title_def']    = [3,  2, 'Berlin Sans FB Demi', 24, 0, clr1, '', True, False, 'left', self.tab_title]
        self.tab_def['tab_cd_subtitle_def'] = [3,  3, '', sz, 0, '', '', False,  False, 'left', '']
        self.tab_def['tab_cd_notes_def']    = [3, 11, '', sz, 0, '', '', False,  False, 'left', '']
        self.tab_def['tab_help_txt'] = self.help_txt

        self.tab_def['tab_cd_table_hdr'] = {  # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
              "RowId":    [10, 10, '', sz,  8, txt1, clr1, True, False, 'center', self.hdr_RowId]
            , "Tag Name": [11, 10, '', sz, 27, txt1, clr1, True, False, 'left', self.col_key1]
            , "FCnt":     [12, 10, '', sz,  8, txt1, clr1, True, False, 'center', self.col_lnks]
        }
        self.tab_def['tab_cd_table_dtl'] = {  # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
            "RowId":  [10, 0, '', sz,  0, "", "", False, False, 'center', '']
            , "Tags": [11, 0, '', sz,  0, "", "", True, False, 'left', '']
            , "FCnt": [12, 0, '', sz,  0, "", "", False, False, 'center', '']
        }
        self.tab_def['tab_cd_table_links']  = [13, 0, '', sz, 18, txt1, clr1, False, False, 'left', '']
        self.tab_def['tab_cd_table_spacer'] = [14, 0, '', sz, 1, txt1, clr1, False, False, 'right', '']
        self.tab_def['tab_cd_fixed_summ']   = {  # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
            # totals headers (across and down)
            'summ-title':       [3, 5, '', 14, 21, txt1, clr1, True, False, 'left', 'Analysis']
            , 'Totals':       [0, 0, '', sz, 15, txt1, clr1, True, False, 'center', 'Totals']
            , 'Rows':         [3, 6, '', sz, 0, txt2, clr2,  True, False, 'right',   "Rows"]
            , 'x-uniq-rows':  [0, 0, '', sz, 0, "", "",      False, False, 'center', self.f_txt_rows]
            , 'Tags':         [3, 7, '', sz, 0, txt2, clr2,  True, False, 'right', self.col_key1]
            , 'x-ctot-key1':  [0, 0, '', sz, 0, "", "",      False, False, 'center', self.f_txt_key1]
            , 'Files':        [3, 8, '', sz, 0, txt2, clr2,  True, False, 'right', self.col_lnks]
            , 'x-ctot-val1':  [0, 0, '', sz, 0, "", "",      False, False, 'center', self.f_num_lnks]
        }
        self.tab_def['tab_cd_fixed_grid']   = {  # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
            # totals headers (across and down)
             'Notes-hdr':    [3, 10, '', sz,  0, txt1, clr1, True, False, 'left', 'Notes: ']
            , 'x-filters-on': [ 5,  6, '', 12, 0, self.colors.clr_red, "", True, True, 'left', self.f_filters_on]
            # This one is for the IVisible Column, not the totals
            # The 'isVisible' column, below, is set in calc_col_pointers()
            , 'isVisible':    [ 0, 13, '', sz, 0, clr2, clr2, False, False, 'right', self.f_isVisible]
        }

        self.tab_def_post()

class DefFile(NewTab):
    def __init__(self,wb_obj):
        self.tab_id = 'file'
        self.tab_common = wb_obj.tab_common
        super().__init__(self.tab_id, wb_obj)

        clr1, txt1, clr2, txt2, table_style = self.colors.get_tab_clrs(self.tab_id)
        self.tab_def['tab_table_style'] = table_style
        # clr1 = tab color,
        # clr2 = secondary "highlights" color, headings
        # clr1 = fill color on cells that use color fills
        # txt1 = text color on cells that use color fills
        red0 = self.colors.get_clr('red', 0)

        sz = self.tab_txt_sz

        self.font_title_lst = ['Berlin Sans FB Demi', 24, clr1]
        self.font_subs_lst  = ['Berlin Sans', 14, txt1]
        self.font_body_lst  = ['Calibri', sz, txt1]
        self.tab_def['tab_color'] = clr1

        self.tab_def['tab_table_link_spcrs'] = True  # Always, TRUE for now
        self.tab_def['tab_txt_sz']    = sz
        self.tab_def['showGridLines'] = self.showGridLines

        self.tab_def['hdr_links_pfx'] = "File"
        self.tab_def['tab_table_links_cols'] = 0
        self.tab_def['tab_has_isVisible_col'] = True
        self.tab_def['tab_tots_isVisible_col'] = 18

        # self.tab_def['tab_name'] = 'All Files'
        self.tab_def['tab_cd_title_def'] = [3, 2, 'Berlin Sans FB Demi', 24, 0, clr1, '', True, False, 'left', self.tab_title]
        self.tab_def['tab_cd_subtitle_def'] = [ 3, 3, '', sz, 0, '', '', False,  False, 'left', '']
        self.tab_def['tab_cd_notes_def'] =    [ 3, 13, '', sz, 0, '', '', False,  False, 'left', '']
        self.tab_def['tab_help_txt'] = self.help_txt

        self.tab_def['tab_cd_table_hdr'] = {
                      # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
             "RowId":      [10, 10, '', sz,  8, txt1, clr1, True, False, 'center', self.hdr_RowId]
           , "FileNm":     [11, 10, '', sz, 30, txt1, clr1, True, False, 'left'  , self.col_key1]
           , "Spacer":     [12, 10, '', sz, 1, txt1, clr1, True, False, 'right'  , '  spcr']
           , "Inline?":    [13, 10, '', sz, 10, txt1, clr1, True, False, 'left', self.col_key2]
           , "Property":   [14, 10, '', sz, 18, txt1, clr1, True, False, 'left'  , self.col_val1]
           , "As Entered": [15, 10, '', sz, 18, txt1, clr1, True, False, 'left'  , self.col_val2]
           , "ValCount":   [16, 10, '', sz, 18, txt1, clr1, True, False, 'left'  , self.col_lnks]
           , "Values":     [17, 10, '', sz, 50, txt1, clr1, True, False, 'center'  , "Values (All)"]
          , 'isVisible':   [18,  0, '', sz,  1, txt1, clr1, False, False, 'right', self.hdr_IsVis]
           }
        self.tab_def['tab_cd_table_dtl'] =  {
            # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
             "RowId":     [10,  0, '', sz, 0, "", "",   False, False, 'center', '']
           , "filenm":    [11,  0, '', sz, 0, "", "",   True, False,  'left', '']
           , "spacer":    [12,  0, '', sz, 0, "", "",   True, False,  'left', '']
           , "inline":    [13,  0, '', sz, 0, "", "",   False, False, 'center', '']
           , "property":  [14,  0, '', sz, 0, "", "",   False, False, 'left', '']
           , "casediff":  [15,  0, '', sz, 0, red0, "", False, True,  'left', '']
           , "val-cnt":   [16,  0, '', sz, 0, "", "",   False, False, 'center', '']
           , "val":       [17,  0, '', sz, 0, "", "",   False, False, 'left', '']
           , 'isVisible': [18,  0, '',  8, 0, "", "",   False, False, 'right',  self.f_isVisible]
           }
        self.tab_def['tab_cd_table_links']     = [ 9, 0, '', sz, 18, txt1, clr1, False, False, 'left'  , '']
        self.tab_def['tab_cd_table_spacer'] = [10, 0, '', sz,  1, txt1, clr1, False, False, 'right'  , '']
        self.tab_def['tab_cd_fixed_summ']   = {  # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
            # totals headers (across and down)
            'summ-title':    [3,  5, '', 14, 21, txt1, clr1, True,  False, 'left',  'Analysis']
          , 'Totals':      [0,  0, '', sz, 15, txt1, clr1, True,  False, 'center', 'Totals']
          , 'Rows':        [3,  6, '', sz,  0, txt2, clr2, True,  False, 'right',  "Rows"]
          , 'x-uniq-rows': [0,  0, '', sz,  0, "", "",     False, False, 'center', self.f_txt_rows]
          , 'Files_Notes': [3,  7, '', sz,  0, txt2, clr2, True,  False, 'right',  self.col_key1]
          , 'x-ctot-key1': [0,  0, '', sz,  0, "", "",     False, False, 'center', self.f_uniq_key1]
          , 'Inline-ind':  [3,  8, '', sz,  0, txt2, clr2, True,  False, 'right',  self.col_key2]
          , 'x-ctot-key2': [0,  0, '', sz,  0, "", "",     False, False, 'center', self.f_txt_key2]
          , 'Properties':  [3,  9, '', sz,  0, txt2, clr2, True,  False, 'right',  self.col_val1]
          , 'x-ctot-val1': [0,  0, '', sz,  0, "", "",     False, False, 'center', self.f_uniq_val1]
          , 'CaseDiff':    [3, 10, '', sz,  0, txt2, clr2, True,  False, 'right',  self.col_val2]
          , 'x-ctot-val2': [0,  0, '', sz,  0, "", "",     False, False, 'center', self.f_txt_val2]
        }

        self.tab_def['tab_cd_fixed_grid'] = {
           # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
            'Notes-hdr':   [3, 12, '', sz, 0, txt1, clr1,  True,  False, 'left', 'Notes: ']

          , 'x-null-values': [11,  8, '', 12, 0, self.colors.clr_red,  "", True, True,   'left',   self.f_null_values]
          , 'x-filters-on':  [ 5,  6, '', 12, 0, self.colors.clr_red,  "", True, True,   'left',   self.f_filters_on]
          # This one is for the IVisible Column, not the totals
          , 'isVisible': [18, 0, '', sz, 0, clr2, clr2, False, False, 'right', self.f_isVisible]

        }

        self.tab_def_post()

class DefCode(NewTab):
    def __init__(self,wb_obj):
        self.tab_id = 'code'
        self.tab_common = wb_obj.tab_common
        super().__init__(self.tab_id, wb_obj)

        clr1, txt1, clr2, txt2, table_style = self.colors.get_tab_clrs(self.tab_id)
        self.tab_def['tab_table_style'] = table_style
        # clr1 = tab color,
        # clr2 = secondary "highlights" color, headings
        # clr1 = fill color on cells that use color fills
        # txt1 = text color on cells that use color fills

        sz = self.tab_txt_sz
        self.font_title_lst = ['Berlin Sans FB Demi', 24, clr1]
        self.font_subs_lst  = ['Berlin Sans', 14, txt1]
        self.font_body_lst  = ['Calibri', sz, txt1]
        self.tab_def['tab_color'] = clr1

        self.tab_def['tab_table_link_spcrs'] = True  # Always, TRUE for now
        self.tab_def['tab_txt_sz']    = sz
        self.tab_def['showGridLines'] = self.showGridLines

        self.tab_def['hdr_links_pfx'] = "CodeBlock-"
        self.tab_def['tab_table_links_cols'] = 20
        self.tab_def['tab_has_isVisible_col'] = True
        self.tab_def['tab_tots_isVisible_col'] = 44

        # self.tab_def['tab_name'] = 'Code'
        self.tab_def['tab_cd_title_def']    = [3,  2, 'Berlin Sans FB Demi', 24, 0, clr1, '', True, False, 'left', self.tab_title]
        self.tab_def['tab_cd_subtitle_def'] = [3,  3, '', sz, 0, '', '', False,  False, 'left', '']
        self.tab_def['tab_cd_notes_def']    = [3, 13, '', sz, 0, '', '', False,  False, 'left', '']
        self.tab_def['tab_help_txt'] = self.help_txt

        self.tab_def['tab_cd_table_hdr'] = {
                      # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
             "RowId":       [10, 10, '', sz,  8, txt1, clr1, True, False, 'center', self.hdr_RowId]
           , "FileNm":      [11, 10, '', sz, 30, txt1, clr1, True, False, 'left'  , self.col_key1]
           , "PluginId":    [12, 30, '', sz, 25, txt1, clr1, True, False, 'left' , self.col_key2]
           , "Signature":   [13, 10, '', sz, 17, txt1, clr1, True, False, 'left' , self.col_val1]
           , "CbCnt":       [14, 10, '', sz,  8, txt1, clr1, True, False, 'left' , self.col_lnks]
           }
        self.tab_def['tab_cd_table_dtl'] =  {
            # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
             "RowId":     [10,  0, '', sz, 0, "", "", False, False, 'center', '']
           , "filenm":    [11,  0, '', sz, 0, "", "", True, False,  'left', '']
           , "pluginId":  [12,  0, '', sz, 0, "", "", False, False, 'left', '']
           , "sig":       [13,  0, '', sz, 0, "", "", False, False, 'left', '']
           , "cbCnt":     [14,  0, '', sz, 0, "", "", False, False, 'center', '']
           }
        self.tab_def['tab_cd_table_links']     = [15, 0, '', sz, 25, txt1, clr1, False, False, 'left'  , '']
        self.tab_def['tab_cd_table_spacer'] = [16, 0, '', sz,  1, txt1, clr1, False, False, 'right'  , '']
        self.tab_def['tab_cd_fixed_summ']   = {  # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
            # totals headers (across and down)
            'summ-title':    [3,  5, '', 14, 21, txt1, clr1, True,  False, 'left',   'Analysis']
          , 'Totals':      [0,  0, '', sz, 15, txt1, clr1, True,  False, 'center', 'Totals']
          , 'Rows':        [3,  6, '', sz,  0, txt2, clr2, True,  False, 'right',  "Rows"]
          , 'x-uniq-rows': [0,  0, '', sz,  0, "", "",     False, False, 'center', self.f_txt_rows]
          , 'Files_Notes': [3,  7, '', sz,  0, txt2, clr2, True,  False, 'right',  self.col_key1]
          , 'x-ctot-key1': [0,  0, '', sz,  0, "", "",     False, False, 'center', self.f_uniq_key1]
          , 'Plug-ins':    [3,  8, '', sz,  0, txt2, clr2, True,  False, 'right',  self.col_key2]
          , 'x-ctot-key2': [0,  0, '', sz,  0, "", "",     False, False, 'center', self.f_uniq_key2]
          , 'Signatures':  [3,  9, '', sz,  0, txt2, clr2, True,  False, 'right',  self.col_val1]
          , 'x-ctot-val1': [0,  0, '', sz,  0, "", "",     False, False, 'center', self.f_uniq_val1]
          , 'Count':       [3, 10, '', sz,  0, txt2, clr2, True,  False, 'right',  self.col_lnks]
          , 'x-ctot-lnks': [0,  0, '', sz,  0, "", "",     False, False, 'center', self.f_num_lnks]
        }
        self.tab_def['tab_cd_fixed_grid']   = {
          # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
          # totals headers (across and down)
            'Notes-hdr':   [3, 12, '', sz,  0, txt1, clr1, True,  False, 'left',   'Notes: ']

          , 'x-null-values':  [11,  8, '', 12, 0, self.colors.clr_red,  "", True, True,   'left',   self.f_null_values]
          , 'x-filters-on':   [ 5,  6, '', 12, 0, self.colors.clr_red,  "", True, True,   'left',   self.f_filters_on]
          # This one is for the IVisible Column, not the totals
          , 'isVisible':      [44, 0, '', sz, 0, clr2, clr2, False, False, 'right',  self.f_isVisible]
          }

        self.tab_def_post()

class DefXyml(NewTab):
    def __init__(self,wb_obj):
        self.tab_id = 'xyml'
        self.tab_common = wb_obj.tab_common
        super().__init__(self.tab_id, wb_obj)

        clr1, txt1, clr2, txt2, table_style = self.colors.get_tab_clrs(self.tab_id)
        self.tab_def['tab_table_style'] = table_style
        # clr1 = tab color,
        # clr2 = secondary "highlights" color, headings
        # clr1 = fill color on cells that use color fills
        # txt1 = text color on cells that use color fills

        sz = self.tab_txt_sz
        self.font_title_lst = ['Berlin Sans FB Demi', 24, clr1]
        self.font_subs_lst =  ['Berlin Sans', 14, txt1]
        self.font_body_lst =  ['Calibri', sz, txt1]
        self.tab_def['tab_color'] = clr1

        # self.tab_def['tab_table_link_spcrs'] = True  # Always, TRUE for now

        self.tab_def['tab_has_isVisible_col'] = False
        self.tab_def['tab_tots_isVisible_col']  = 0
        self.tab_def['showGridLines'] = self.showGridLines

        self.tab_def['hdr_links_pfx'] = ""
        self.tab_def['tab_table_links_cols'] = 0

        # self.tab_def['tab_name'] = 'Xyml'
        self.tab_def['tab_cd_title_def']    = [3,  2, 'Berlin Sans FB Demi', 24, 0, clr1, '', True, False, 'left', self.tab_title]
        self.tab_def['tab_cd_subtitle_def'] = [3,  3, '', sz, 0, '', '', False,  False, 'left', '']
        self.tab_def['tab_cd_notes_def']    = [3, 16, '', sz, 0, '', '', False,  False, 'left', '']
        self.tab_def['tab_help_txt'] = self.help_txt

        self.tab_def['tab_cd_table_hdr'] = {  # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
             "RowId":         [10, 15, '', sz,  8, txt1, clr1, True, False, 'center', self.hdr_RowId]
            , "Filename":     [11, 15, '', sz, 30, txt1, clr1, True, False, 'left', self.col_key1]
            , "InFiles":      [12, 15, '', sz, 30, txt1, clr1, True, False, 'left', self.col_val1]
            , "Xyml-type":    [13, 15, '', sz, 80, txt1, clr1, True, False, 'left', self.col_val2]
        }

        self.tab_def['tab_cd_table_dtl'] = {  # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
             "RowId":         [10, 0, '', sz, 0, "", "", False, False, 'center', '']
            , "xyml-type":    [11, 0, '', sz, 0, "", "", True, False, 'left', '']
            , "in-files":     [12, 0, '', sz, 0, "", "", True, False, 'left', '']
            , "Filename":     [13, 0, '', sz, 0, "", "", True, False, 'left', '']
        }
        self.tab_def['tab_cd_table_links']  = [13, 10, '', sz, 40, txt1, clr1, False, False, 'left', '']
        self.tab_def['tab_cd_table_spacer'] = [14,  0, '', sz,  1, txt1, clr1, False, False, 'right', '']

        self.tab_def['tab_cd_fixed_summ'] = {  # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
            # totals headers (across and down)
              'summ-title':  [3, 5, '', 14, 26, txt1, clr1, True, False, 'left', 'Analysis']
            , 'Totals':      [0, 0, '', sz, 15, txt1, clr1, True, False, 'center', 'Totals']

            , 'Rows':        [3, 6, '', sz, 0, txt2, clr2, True, False, 'right', "Rows"]
            , 'x-uniq-rows': [0, 0, '', sz, 0, "", "", False, False, 'center', self.f_txt_rows]

            , 'Files':       [3, 7, '', sz, 0, txt2, clr2, True, False, 'right', self.col_key1]
            , 'x-ctot-key1': [0, 0, '', sz, 0, "", "", False, False, 'center', self.f_txt_key1]

            , 'in-Files':    [3, 8, '', sz, 0, txt2, clr2, True, False, 'right', self.col_val1]
            , 'x-ctot-key2': [0, 0, '', sz, 0, "", "", False, False, 'center', self.f_cif_val1_true]
        }
        self.tab_def['tab_cd_fixed_grid']   = {
                          # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
              'Notes-hdr':   [3, 15, '', 12, 32, txt1, clr1, True, False, 'left', 'Notes: ']
        }
        # Special Totals and Notes for GRID
        tab_note = self.tab_common['xyml']['help_txt']['notes']
        tab_note.append(f"{self.tab_common['xyml']['col_val2']} Details:")
        row_num = 9
        for xyml_key, xyml_desc_list in self.xyml_descs.items():
            key_id = f'x-tot-key{row_num}'
            val_id = f'f-tot-key{row_num}'
            cntif = f'=COUNTIF({self.tbl_name}[{self.col_val2}],"{xyml_desc_list[0]}")'
            self.tab_def['tab_cd_fixed_summ'][key_id] = [3, row_num, '', sz, 0, txt2, clr2, True, False, 'right', xyml_desc_list[0]]
            self.tab_def['tab_cd_fixed_summ'][val_id] = [0, 0, '', sz, 0, "", "", False, False, 'center', cntif]
            tab_note.append(f' - {xyml_desc_list[0]:<30}: {xyml_desc_list[1]}')
            row_num += 1

        self.tab_common['xyml']['help_txt']['notes'] = tab_note
        #   'BadY': ["Invalid Properties",            'Cannot Load Frontmatter-Check YAML Markdown syntax.']
        # , 'NoFm': ['No Properties',                 'Not a problem, if intentional.']
        # , 'MtFm': ['Frontmatter loaded, but empty', 'Not a problem, if intentional.']
        #             123456789012345678901234567890
        # , 'ErrY': ["Frontmatter error",             'An Unknown Error occurred trying to process Frontmatter.']
        # , 'NonD': ['Frontmatter formatting error',  'Invalid Frontmatter--Not in dictionary format']
        row_idx = 7
        for dkey, desc in self.xyml_descs.items():
            dkey_2 = f"{dkey}{row_idx}"
            f_tot = f"=COUNTIF(tbl_xyml[{self.col_val1}],C{row_idx})"
            # self.tab_def['tab_cd_fixed_grid'][dkey]   = [ 3, row_idx, '', sz, 0, txt2, clr2, True,  False, 'right', desc[0]]
            # self.tab_def['tab_cd_fixed_grid'][dkey_2] = [ 4, row_idx, '', sz, 0, "", "",       False, False, 'center', f_tot]
            row_idx += 1

        self.tab_def_post()

class DefDups(NewTab):
    def __init__(self,wb_obj):
        self.tab_id = 'dups'
        self.tab_common = wb_obj.tab_common
        super().__init__(self.tab_id, wb_obj)

        clr1, txt1, clr2, txt2, table_style = self.colors.get_tab_clrs(self.tab_id)
        self.tab_def['tab_table_style'] = table_style
        # clr1 = tab color,
        # clr2 = secondary "highlights" color, headings
        # clr1 = fill color on cells that use color fills
        # txt1 = text color on cells that use color fills

        sz = self.tab_txt_sz

        self.font_title_lst = ['Berlin Sans FB Demi', 24, clr1]
        self.font_subs_lst = ['Berlin Sans', 14, txt1]
        self.font_body_lst = ['Calibri', sz, txt1]
        self.tab_def['tab_color'] = clr1

        self.tab_def['tab_table_link_spcrs'] = True  # Always, TRUE for now

        self.tab_def['tab_has_isVisible_col'] = False
        self.tab_def['tab_tots_isVisible_col']  = 0
        self.tab_def['showGridLines'] = self.showGridLines

        self.tab_def['hdr_links_pfx'] = "Full Pathnames"
        self.tab_def['tab_table_links_cols']    = 4

        # self.tab_def['tab_name'] = 'Dups'
        self.tab_def['tab_cd_title_def']    = [ 3,  2, 'Berlin Sans FB Demi', 24, 0, clr1, '', True, False, 'left', self.tab_title]
        self.tab_def['tab_cd_subtitle_def'] = [ 3,  3, '', sz, 0, '', '', False,  False, 'left', '']
        self.tab_def['tab_cd_notes_def']    = [ 3, 11, '', sz, 0, '', '', False,  False, 'left', '']
        self.tab_def['tab_help_txt'] = self.help_txt

        self.tab_def['tab_cd_table_hdr'] = {
                        # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
             "RowId":       [10, 10, '', sz, 8, txt1, clr1, True, False, 'center', self.hdr_RowId]
            , "Filename":   [11, 10, '', sz, 35, txt1, clr1, True, False, 'left', self.col_key1]
            , "Links": [12, 10, '', sz, 8, txt1, clr1, True, False, 'center', self.col_lnks]
        }
        self.tab_def['tab_cd_table_dtl'] = {  # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
              "RowId":      [10, 0, '', sz, 0, "", "", False, False, 'center', '']
            , "Filename":   [11, 0, '', sz, 0, "", "", True, False, 'left', '']
            , "Links": [12, 0, '', sz, 0, "", "", False, False, 'center', '']
        }
        self.tab_def['tab_cd_table_links']     = [13, 0, '', sz, 25, txt1, clr1, False, False, 'left', '']
        self.tab_def['tab_cd_table_spacer'] = [14, 0, '', sz,  1, txt1, clr1, False, False, 'right', '']
        self.tab_def['tab_cd_fixed_summ'] = {  # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
            # totals headers (across and down)
            'summ-title':  [3, 5, '', 14, 21, txt1, clr1, True, False, 'left', 'Analysis']
            , 'Totals':      [0, 0, '', sz, 15, txt1, clr1, True, False, 'center', 'Totals']
            , 'Rows':        [3, 6, '', sz, 0, txt2, clr2, True, False, 'right', "Rows"]
            , 'x-uniq-rows': [0, 0, '', sz, 0, "", "", False, False, 'center', self.f_txt_rows]
            , 'Tags':        [3, 7, '', sz, 0, txt2, clr2, True, False, 'right', self.col_key1]
            , 'x-ctot-key1': [0, 0, '', sz, 0, "", "", False, False, 'center', self.f_txt_key1]
            , 'Files':       [3, 8, '', sz, 0, txt2, clr2, True, False, 'right', self.col_lnks]
            , 'x-ctot-val1': [0, 0, '', sz, 0, "", "", False, False, 'center', self.f_num_lnks]
        }
        self.tab_def['tab_cd_fixed_grid'] = {
            # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
            # totals headers (across, then down)
              'Notes-hdr':   [3, 10, '', sz, 0, txt1, clr1, True, False, 'left', 'Notes: ']
        }

        self.tab_def_post()

class DefNest(NewTab):
    def __init__(self,wb_obj):
        self.tab_id = 'nest'
        self.tab_common = wb_obj.tab_common
        super().__init__(self.tab_id, wb_obj)

        clr1, txt1, clr2, txt2, table_style = self.colors.get_tab_clrs(self.tab_id)
        self.tab_def['tab_table_style'] = table_style
        # clr1 = tab color,
        # clr2 = secondary "highlights" color, headings
        # clr1 = fill color on cells that use color fills
        # txt1 = text color on cells that use color fills

        sz = self.tab_txt_sz
        self.font_title_lst = ['Berlin Sans FB Demi', 24, clr1]
        self.font_subs_lst = ['Berlin Sans', 14, txt1]
        self.font_body_lst = ['Calibri', sz, txt1]
        self.tab_def['tab_color'] = clr1

        self.tab_def['tab_table_link_spcrs'] = True  # Always, TRUE for now
        self.tab_def['tab_txt_sz'] = sz
        self.tab_def['showGridLines'] = self.showGridLines

        self.tab_def['hdr_links_pfx'] = "File"
        self.tab_def['tab_table_links_cols'] = 0
        self.tab_def['tab_has_isVisible_col'] = True
        self.tab_def['tab_tots_isVisible_col'] = 16

        # if self.tab_def['tab_name'] != 'Plug-Ins':
        #     raise WorkbookDefinitionError(f"Tab_Def: pros-tab_name tab name not defined.")
        # self.tab_def['tab_name'] = 'Properties'
        self.tab_def['tab_cd_title_def'] = [3, 2, 'Berlin Sans FB Demi', 24, 0, clr1, '', True, False, 'left',
                                         self.tab_title]
        self.tab_def['tab_cd_subtitle_def'] = [3,  3, '', sz, 0, '', '', False, False, 'left', '']
        self.tab_def['tab_cd_notes_def']    = [3, 12, '', sz, 0, '', '', False, False, 'left', '']
        self.tab_def['tab_help_txt'] = self.help_txt

        self.tab_def['tab_cd_table_hdr'] = {
            # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
            # RowId
            # File
            # Plugin
            # Props
            # Values
            # Value Cnt
            # Value01
            # Value02

              "RowId":      [10, 10, '', sz,  8, txt1, clr1, True, False,  'center', self.hdr_RowId]
            , "Plugin":     [11, 10, '', sz, 30, txt1, clr1, True, False,  'left',   self.col_key1]
            , "FileName":   [12, 10, '', sz, 30, txt1, clr1, True, False,  'left',   self.col_key2]
            , "Prop":       [13, 10, '', sz, 30, txt1, clr1, True, False,  'left',   'Properties']
            , "Val-Cnt":    [14, 10, '', sz, 8,  txt1, clr1, True, False,  'left', self.col_val1]
            , "value_list": [15, 10, '', sz, 40,  txt1, clr1, True, False,  'left',   self.col_lnks]
            , 'isVisible':  [16,  0, '', sz, 1,  txt1, clr1, False, False, 'right', self.hdr_IsVis]
        }
        self.tab_def['tab_cd_table_dtl'] = {
            # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
              "rowId":      [10, 0, '', sz, 0, "", "", False, False, 'center', '']
            , "plugin":     [11, 0, '', sz, 0, "", "", True, False, 'left', '']
            , "filename":   [12, 0, '', sz, 0, "", "", True, False, 'left', '']
            , "prop":       [13, 0, '', sz, 0, "", "", False, False, 'left', '']
            , "val-cnt":    [14, 0, '', sz, 0, "", "", False, False, 'center', '']
            , "value_list": [15, 0, '', sz, 0, "", "", False, False, 'left', '']
            , 'isVisible':  [16, 0, '', sz, 0, "", "", False, False, 'right', self.f_isVisible]
        }
        self.tab_def['tab_cd_table_links']  = [17, 0, '', sz, 25, txt1, clr1, False, False, 'left', '']
        self.tab_def['tab_cd_table_spacer'] = [18, 0, '', sz,  1, txt1, clr1, False, False, 'right', '']
        self.tab_def['tab_cd_fixed_summ']   = {  # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
            # totals headers (across and down)

              'summ-title':    [3,  5, '', 14, 21, txt1, clr1, True,  False, 'left',   'Analysis']
            , 'Totals':      [0,  0, '', sz, 15, txt1, clr1, True,  False, 'center', 'Totals']
            , 'Rows':        [3,  6, '', sz,  0, txt2, clr2, True,  False, 'right',  "Rows"]
            , 'x-uniq-rows': [0,  0, '', sz,  0, "", "",     False, False, 'center', self.f_txt_rows]
            , 'Plug-In':     [3,  7, '', sz,  0, txt2, clr2, True,  False, 'right',  self.col_key1]
            , 'x-ctot-key1': [0,  0, '', sz,  0, "", "",     False, False, 'center', self.f_uniq_key1]
            , 'Filenames':   [3,  8, '', sz,  0, txt2, clr2, True,  False, 'right',  self.col_key2]
            , 'x-ctot-key2': [0,  0, '', sz,  0, "", "",     False, False, 'center', self.f_uniq_key2]
            , 'ValCount':    [3,  9, '', sz,  0, txt2, clr2, True,  False, 'right',  self.col_val1]
            , 'x-ctot-val1': [0,  0, '', sz,  0, "", "",     False, False, 'center', self.f_uniq_val1]
        }
        self.tab_def['tab_cd_fixed_grid'] = {
            # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
            # totals headers (across and down)
              'Notes-hdr':   [3, 11, '', sz,  0, txt1, clr1, True,  False, 'left', 'Notes: ']

            , 'x-null-values':  [11,  8, '', 12, 0, self.colors.clr_red, "", True, True, 'left', self.f_null_values]
            , 'x-filters-on':   [ 5,  6, '', 12, 0, self.colors.clr_red, "", True, True, 'left', self.f_filters_on]
            # This one is for the IVisible Column, not the totals
            , 'isVisible':      [16, 0, '', sz, 0, clr2, clr2, False, False, 'right', self.f_isVisible]
        }

        self.tab_def_post()

class DefPlug(NewTab):
    def __init__(self,wb_obj):
        self.tab_id = 'plug'
        self.tab_common = wb_obj.tab_common
        super().__init__(self.tab_id, wb_obj)

        clr1, txt1, clr2, txt2, table_style = self.colors.get_tab_clrs(self.tab_id)
        self.tab_def['tab_table_style'] = table_style
        # clr1 = tab color,
        # clr2 = secondary "highlights" color, headings
        # clr1 = fill color on cells that use color fills
        # txt1 = text color on cells that use color fills

        sz = self.tab_txt_sz
        self.font_title_lst = ['Berlin Sans FB Demi', 24, clr1]
        self.font_subs_lst = ['Berlin Sans', 14, txt1]
        self.font_body_lst = ['Calibri', sz, txt1]
        self.tab_def['tab_color'] = clr1

        self.tab_def['tab_table_link_spcrs'] = True  # Always, TRUE for now
        self.tab_def['tab_txt_sz'] = sz
        self.tab_def['showGridLines'] = self.showGridLines

        self.tab_def['hdr_links_pfx'] = "File"
        self.tab_def['tab_table_links_cols'] = 0
        self.tab_def['tab_has_isVisible_col'] = False
        self.tab_def['tab_tots_isVisible_col'] = 0

        # if self.tab_def['tab_name'] != 'Plug-Ins':
        #     raise WorkbookDefinitionError(f"Tab_Def: pros-tab_name tab name not defined.")
        # self.tab_def['tab_name'] = 'Properties'
        self.tab_def['tab_cd_title_def'] = [3, 2, 'Berlin Sans FB Demi', 24, 0, clr1, '', True, False, 'left', self.tab_title]
        self.tab_def['tab_cd_subtitle_def'] = [3, 3, '', sz, 0, '', '', False, False, 'left', '']
        self.tab_def['tab_cd_notes_def']    = [3, 11, '', sz, 0, '', '', False, False, 'left', '']
        self.tab_def['tab_help_txt'] = self.help_txt

        self.tab_def['tab_cd_table_hdr'] = {
            # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
              "RowId":          [10, 10, '', sz,  8, txt1, clr1, False, False, 'left', 'RowId']
            , 'name':           [ 0,  0, '', sz, 30, txt1, clr1, True,  False, 'left', self.col_key1]
            , 'enabled':        [ 0,  0, '', sz, 15, txt1, clr1, True,  False, 'left', self.col_key2]
            , 'version':        [ 0,  0, '', sz, 10, txt1, clr1, False, False, 'left', 'Version']
            , 'minAppVersion':  [ 0,  0, '', sz,  8, txt1, clr1, False, False, 'left', 'Min App Version']
            , 'author':         [ 0,  0, '', sz, 20, txt1, clr1, False, False, 'left', 'Author']
            , 'authorUrl':      [ 0,  0, '', sz, 20, txt1, clr1, False, False, 'left', 'Authors Url']
            , 'isDesktopOnly':  [ 0,  0, '', sz,  7, txt1, clr1, False, False, 'left', self.col_val1]
            , 'description':    [ 0,  0, '', sz, 50, txt1, clr1, False, False, 'left', 'Description']
            , 'cb-sig-list':    [ 0,  0, '', sz, 30, txt1, clr1, False, False, 'left', 'Plugin Codeblock Signatures']
        }
        self.tab_def['tab_cd_table_dtl'] = {
            # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
              "rowId":          [10, 0, '', sz, 0, "", "", False, False, 'center', '']
            , 'name':           [ 0, 0, '', sz, 0, "", "", False, False, 'left',   '']
            , 'enabled':        [ 0, 0, '', sz, 0, "", "", False, False, 'center', '']
            , 'version':        [ 0, 0, '', sz, 0, "", "", True,  False, 'left',   '']
            , 'minAppVersion':  [ 0, 0, '', sz, 0, "", "", False, False, 'left',   '']
            , 'author':         [ 0, 0, '', sz, 0, "", "", False, False, 'left',   '']
            , 'authorUrl':      [ 0, 0, '', sz, 0, "", "", False, False, 'left',   '']
            , 'isDesktopOnly':  [ 0, 0, '', sz, 0, "", "", False, False, 'left',   '']
            , 'Description':    [ 0, 0, '', sz, 0, "", "", False, False, 'left',   '']
            , 'cb-sig-list':    [ 0, 0, '', sz, 0, "", "", False, False, 'left',   '']
        }
        self.tab_def['tab_cd_table_links']  = [17, 0, '', sz, 25, txt1, clr1, False, False, 'left', '']
        self.tab_def['tab_cd_table_spacer'] = [18, 0, '', sz,  1, txt1, clr1, False, False, 'right', '']
        self.tab_def['tab_cd_fixed_summ']   = {  # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
            # totals headers (across and down)

            # col=0 means next col, row = 0 means same row
             'summ-title':         [ 3,  5, '', 14, 32, txt1, clr1, True, False,  'left', 'Analysis']
            , 'Total':          [ 0, 0, '',  sz, 15, txt1, clr1, True, False,  'center', 'Totals']

            , 'Plugins':        [ 3, 6, '',  sz, 0, txt2, clr2, True, False,  'right',  self.col_key1]
            , 'x-ctot-key1':    [ 0, 0, '',  sz, 0, "", "",     False, False, 'center', self.f_txt_key1]

            , 'Enabled':        [ 3, 7, '',  sz, 0, txt2, clr2, True, False,  'right',  self.col_key2]
            , 'x-ctot-val2':    [ 0, 0, '',  sz, 0, "", "",     False, False, 'center', self.f_txt_key2]

            , 'Desktop':        [ 3, 8, '',  sz, 0, txt2, clr2, True, False,  'right',  self.col_val1]
            , 'x=desktop':      [ 0, 0, '',  sz, 0, "", "",     False, False, 'center', self.f_txt_val1]
                                   # , f'=COUNTIF({self.tbl_name}[{self.col_val1}],"TRUE")']
        }
        self.tab_def['tab_cd_fixed_grid'] = {
            # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
            # totals headers (across and down)
              'Notes-hdr':      [ 3, 10, '', sz, 0, txt1, clr1, True, False, 'left', 'Notes: ']
            , 'x-null-values':  [11,  8, '', 12, 0, self.colors.clr_red, "", True, True, 'left', self.f_null_values]
            , 'x-filters-on':   [ 5,  6, '', 12, 0, self.colors.clr_red, "", True, True, 'left', self.f_filters_on]
            
              # This one is for the IVisible Column, not the totals
            , 'isVisible': [16, 0, '', sz, 0, clr2, clr2, False, False, 'right', self.f_isVisible]
        }

        self.tab_def_post()

class DefTmpl(NewTab):
    def __init__(self,wb_obj):
        self.tab_id = 'tmpl'
        self.tab_common = wb_obj.tab_common
        super().__init__(self.tab_id, wb_obj)

        clr1, txt1, clr2, txt2, table_style = self.colors.get_tab_clrs(self.tab_id)
        # clr1 = tab color,
        # clr2 = secondary "highlights" color, headings
        # clr1 = fill color on cells that use color fills
        # txt1 = text color on cells that use color fills
        self.tab_def['tab_table_style'] = table_style

        sz = self.tab_txt_sz

        self.font_title_lst = ['Berlin Sans FB Demi', 24, clr1]
        self.font_subs_lst = ['Berlin Sans', 14, txt1]
        self.font_body_lst = ['Calibri', sz, txt1]
        self.tab_def['tab_color'] = clr1

        self.tab_def['tab_table_link_spcrs'] = True  # Always, TRUE for now
        self.tab_def['tab_txt_sz'] = sz
        self.tab_def['showGridLines'] = self.showGridLines

        self.tab_def['hdr_links_pfx'] = "File"
        self.tab_def['tab_table_links_cols'] = 15
        self.tab_def['tab_has_isVisible_col'] = True
        self.tab_def['tab_tots_isVisible_col'] = 38

        # if self.tab_def['tab_name'] != 'Templates':
        #     raise WorkbookDefinitionError(f"Tab_Def: pros-tab_name tab name not defined.")
        # self.tab_def['tab_name'] = 'Properties'
        self.tab_def['tab_cd_title_def'] = [3, 2, 'Berlin Sans FB Demi', 24, 0, clr1, '', True, False, 'left',                                            self.tab_title]
        self.tab_def['tab_cd_subtitle_def'] = [3, 3, '', sz, 0, '', '', False, False, 'left', '']
        self.tab_def['tab_cd_notes_def'] = [3, 22, '', sz, 0, '', '', False, False, 'left', '']
        self.tab_def['tab_help_txt'] = self.help_txt

        self.tab_def['tab_cd_table_hdr'] = {
            # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
            "RowId":  [4, 10, '', sz,  8, txt1, clr1, True, False, 'center', self.hdr_RowId]
            , "Prop": [5, 10, '', sz, 18, txt1, clr1, True, False, 'left', self.col_key1]
            , "Val":  [6, 10, '', sz, 50, txt1, clr1, True, False, 'left', self.col_val1]
            , "FCnt": [7, 10, '', sz,  8, txt1, clr1, True, False, 'center', self.col_lnks]
            , "PVI":  [8, 10, '', sz, 13, txt1, clr1, True, False, 'center', self.hdr_PVI]
        }
        self.tab_def['tab_cd_table_dtl'] = {
            # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
            "RowId": [4, 0, '', sz, 8, "", "", False, False, 'center', '']
            , "Prop": [5, 0, '', sz, 18, "", "", True, False, 'left', '']
            , "Vals": [6, 0, '', sz, 50, "", "", False, False, 'left', '']
            , "FCnt": [7, 0, '', sz, 8, "", "", False, False, 'center', '']
            , "PVI": [8, 0, '', sz, 13, "", "", False, False, 'center', '']
        }
        self.tab_def['tab_cd_table_links'] = [9, 0, '', sz, 18, txt1, clr1, False, False, 'left', '']
        self.tab_def['tab_cd_table_spacer'] = [10, 0, '', sz, 1, txt1, clr1, False, False, 'right', '']
        self.tab_def['tab_cd_fixed_summ']   = {  # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
            # totals headers (across and down)

            'summ-title': [9, 2, '', 14, 0, txt1, clr1, True, False, 'center', 'Analysis']
            , 'S1  ': [10, 2, '', 12, 0, clr1, clr1, True, False, 'right', 'Spc1  ']
            , 'Prop': [11, 2, '', 12, 0, txt1, clr1, True, False, 'center', self.col_key1]
            , 'S2  ': [12, 2, '', 12, 0, clr1, clr1, True, False, 'right', 'Spc2  ']
            , 'Val': [13, 2, '', 12, 0, txt1, clr1, True, False, 'center', self.col_val1]
            , 'S3  ': [14, 2, '', 12, 0, clr1, clr1, True, False, 'right', 'Spc3  ']
            , 'Files Used': [15, 2, '', 12, 0, txt1, clr1, True, False, 'center', self.col_lnks]
            , 'Unique Values': [9, 3, '', 12, 0, txt1, clr1, True, False, 'left', 'Unique Values']
            , 'Column Totals': [9, 4, '', 12, 0, txt1, clr1, True, False, 'left', 'Totals']
        }
        self.tab_def['tab_cd_fixed_grid'] = {
            # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
            # totals headers (across and down)
              'Notes-hdr': [3, 21, '', sz, 0, txt1, clr1, True, False, 'left', 'Notes: ']
            , 'x-uniq-rows': [11, 3, '', sz, 0, "", "", False, False, 'center', self.f_uniq_key1]
            , 'x-uniq-val1': [13, 3, '', sz, 0, "", "", False, False, 'center', self.f_uniq_val1]
            , 'x-ctot-key1': [11, 4, '', sz, 0, "", "", False, False, 'center', self.f_txt_key1]
            , 'x-ctot-files': [15, 4, '', sz, 0, "", "", False, False, 'center', self.f_num_lnks]
            , 'x-null-values': [6, 9, '', 12, 0, self.colors.clr_red, "", True, True, 'left', self.f_null_values]
            , 'x-filters-on': [ 5,  6, '', 12, 0, self.colors.clr_red, "", True, True, 'left', self.f_filters_on]
            # This one is for the IVisible Column, not the totals
            , 'isVisible': [38, 0, '', sz, 0, clr2, clr2, False, False, 'right', self.f_isVisible]
        }

        self.tab_def_post()

class DefSumm(NewTab):
    def __init__(self, wb_obj):
        self.tab_id = 'summ'
        self.tab_common = wb_obj.tab_common
        wb_def = wb_obj.wb_def
        super().__init__(self.tab_id, wb_obj)
        ctot = self.cfg['ctot']
        dir_vault_txt = f"Vault Path: {self.cfg['dir_vault']}"
        c_date = self.cfg['c_date']
        comm = wb_obj.tab_common

        clr1, txt1, clr2, txt2, table_style = self.colors.get_tab_clrs(self.tab_id)
        self.tab_def['tab_table_style'] = table_style
        # clr1 = tab color,
        # clr2 = secondary "highlights" color, headings
        # clr1 = fill color on cells that use color fills
        # txt1 = text color on cells that use color fills

        sz = self.tab_txt_sz

        self.font_title_lst = ['Berlin Sans FB Demi', 24, clr1]
        self.font_subs_lst = ['Berlin Sans', 14, txt1]
        self.font_body_lst = ['Calibri', sz, txt1]
        self.tab_def['tab_color'] = clr1

        self.tab_def['tab_table_link_spcrs']    = True  # Always, TRUE for now

        self.tab_def['tab_has_isVisible_col']   = False
        self.tab_def['tab_tots_isVisible_col']  = 0
        self.tab_def['showGridLines']           = self.showGridLines

        self.tab_def['hdr_links_pfx']           = ""
        self.tab_def['tab_table_links_cols']    = 0
        self.tab_def['tbl_beg_col']             = 0     # disable call to xl.format_as_table for tbl_summ

        clr1, txt1, clr2, txt2,         table_style = self.colors.get_tab_clrs(self.tab_id)
        wht0 = 'FFFFFF'
        val_version = f'=HYPERLINK("https://github.com/slappycat2/obsidian-vault-health-check","v.0.9 (beta)")'
        val_donate  = f'=HYPERLINK("https://ko-fi.com/swenlarsen","support this project!")'
        pros_key1 = f"{comm['pros']['col_key1']}"
        pros_lnks = f"{comm['pros']['col_lnks']}"
        tags_lnks = f"{comm['tags']['col_lnks']}"
        dups_key1 = f"{comm['dups']['col_key1']}"
        xyml_key1 = f"{comm['xyml']['col_key1']}"
        xyml_val2 = f"{comm['xyml']['col_val2']}"
        file_val2 = f"{comm['file']['col_val2']}"
        # Note that the following formulas do not care about hidden rows (op=2). They are true totals.
        hdr_summ = {
              'Total Markdown Files':            [3, 3, ctot[0]]
            , 'Templates (skipped)':             [3, 4, ctot[1]]
            , 'Ignored Folder Notes':            [3, 5, ctot[2]]
            , 'Notes Analyzed':                  [3, 6, ctot[3]]
            , 'Duplicate Notes':                 [3, 7, f"=_xlfn.AGGREGATE(3,2,tbl_dups[{dups_key1}])"]
            , 'Disabled Plugins':                [6, 3, '=J28-J29']
            , 'YAML Issues':                     [6, 4, f"=_xlfn.AGGREGATE(3,2,tbl_xyml[{xyml_key1}])"]
            , 'Notes w/No Properties':           [6, 5, f'=COUNTIF(tbl_xyml[{xyml_val2}],"No Properties")']
            , 'Dataview Query Notes':            [6, 6, f'=COUNTIF(tbl_code[Plugins],"Dataview")']
            , 'Upper Case in YAML':              [6, 7, f"=_xlfn.AGGREGATE(3,2,tbl_file[{file_val2}])"]
            , 'Unique Properties':               [9, 3, f"=_xlfn.AGGREGATE(3,2,tbl_pros[{pros_key1}])"]
            , 'Deprecated Properties':           [9, 4, f'=COUNTIF(tbl_pros[{pros_key1}],"*Deprecated in*")']
            , 'Properties w/colons':             [9, 5, f'=COUNTIF(tbl_pros[{pros_key1}],"*:*")']
            , 'Single Use Properties':           [9, 6, f'=COUNTIF(tbl_pros[{pros_lnks}],1)']
            , 'Single Use Tags':                 [9, 7, f'=COUNTIF(tbl_tags[{tags_lnks}],1)']
            }

        # self.tab_def['tab_name'] = 'Summary'
        self.tab_def['tab_cd_title_def']    = [3,  2, 'Berlin Sans FB Demi', sz, 0, clr1, '', True, False, 'left', '']
        self.tab_def['tab_cd_notes_def']    = [14, 3, '', sz, 0, '', '', False, False, 'left', '']
        # self.tab_def['tab_help_txt']        = self.tab_common['summ']['help_txt']
        self.tab_def['tab_cd_subtitle1_def'] = [3, 2, '', sz, 0, '', '', False, False, 'left', '']
        self.tab_def['tab_cd_subtitle2_def'] = [7, 2, '', sz, 0, '', '', False, False, 'left', '']
        self.tab_def['tab_cd_table_hdr']    = {}
        self.tab_def['tab_cd_table_dtl']    = {}
        self.tab_def['tab_cd_table_links']  = [0, 0, '', 0, 0, txt1, clr1, False, False, '', '']
        self.tab_def['tab_cd_table_spacer'] = [0, 0, '', 0, 0, txt1, clr1, False, False, 'right', '']
        self.tab_def['tab_cd_fixed_grid']   = {  # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
            # totals headers (across and down)

            # This paints the red bar with Workbook Summary, and sets col widths, among other things
              'wbSumm': [ 3, 9, '', 14, 29, txt1, clr1, True,  False, 'left', 'Workbook Summary']
            , 'col-02': [ 0, 0, '', 11, 10, clr1, clr1, False, False, 'left', "."]
            , 'col-03': [ 0, 0, '', 11,  3, clr1, clr1, False, False, 'left', "."]
            , 'col-04': [ 0, 0, '', 11, 25, clr1, clr1, False, False, 'left', "."]
            , 'col-05': [ 0, 0, '', 11, 14, clr1, clr1, False, False, 'left', "."]
            , 'col-06': [ 0, 0, '', 11,  3, clr1, clr1, False, False, 'left', "."]
            , 'col-07': [ 0, 0, '', 11, 25, clr1, clr1, False, False, 'left', "."]
            , 'col-08': [ 0, 0, '', 11, 10, clr1, clr1, False, False, 'left', "."]

            , 'img-01':    [  3,  1, '', sz,  0, '', '',     False, False, 'left', "../img/v_chkBanner2.png"]
            , 'row-01':    [  3,  1, '', 72,  0, wht0, wht0, False, False, 'left', "."]
            , 'dir-vault': [  3,  2, '', 16,  0, '','', True, False, 'left', dir_vault_txt]
            , 'c-date':    [ 10,  2, '', 10,  0, '','', True, False, 'right', c_date]
            , 'version':   [  3, 31, '', sz,  0, '', '',  False, False, 'left', val_version]
            , 'donate':    [ 10, 31, '', sz,  0, '', '',  False, False, 'right', val_donate]
            , 'Notes-hdr': [ 14,  2, '', 16, 32, txt1, clr1, True, False, 'left', 'Notes: ']

        }
        # Continue Building Summary Grid...
        cnt = 0
        for key, val in hdr_summ.items():
            cnt += 1
            uniq_key: str = f'hdr-summ-k-{cnt}'
            uniq_val: str = f'hdr-summ-v-{cnt}'
            def_key = [ val[0], val[1], '', 11, 0, '', '', True, False, 'left', key]
            def_val = [ 0, 0, '', 11, 0, '', '', False, False, 'right', val[2]]
            self.tab_def['tab_cd_fixed_grid'][uniq_key] = def_key
            self.tab_def['tab_cd_fixed_grid'][uniq_val] = def_val

        self.tab_def['borders'] = {
              'footer':["C31:J31", "thin",  self.colors.clr_blk, "top"]
            , 'summ_1':["C11:D11", "thick", self.colors.clr_blk, "bottom"]
            , 'summ_2':["C17:D17", "thick", self.colors.clr_blk, "bottom"]
            , 'summ_3':["C26:D26", "thick", self.colors.clr_blk, "bottom"]
            , 'summ_4':["F11:G11", "thick", self.colors.clr_blk, "bottom"]
            , 'summ_5':["F17:G17", "thick", self.colors.clr_blk, "bottom"]
            , 'summ_6':["F26:G26", "thick", self.colors.clr_blk, "bottom"]
            , 'summ_7':["I11:J11", "thick", self.colors.clr_blk, "bottom"]
            , 'summ_8':["I17:J17", "thick", self.colors.clr_blk, "bottom"]
            , 'summ_9':["I26:J26", "thick", self.colors.clr_blk, "bottom"]
        }
        tab_summ_map = {
              'pros': [ 3, 11]
            , 'vals': [ 3, 17]
            , 'tags': [ 3, 26]
            , 'file': [ 6, 11]
            , 'xyml': [ 6, 17]
            , 'dups': [ 6, 26]
            , 'code': [ 9, 11]
            , 'nest': [ 9, 17]
            , 'plug': [ 9, 26]
        }
        # Builds the colored totals boxes on the summary pages by reading in all the tab_defs
        # from each tab, and changing the col/row to match the tab_summ_map, above
        # It skips the Rows Totals and sets the widths to zero, since they are assigned above.
        self.tab_def['tab_cd_fixed_summ'] = {}
        new_tab_fix_summ = {}

        for tab_id, new_cell_addr in tab_summ_map.items():
            # clr1_fl, clr1_tx, clr2_fl, clr2_tx, _ = self.colors.get_tab_clrs(tab_id)
            col_idx = new_cell_addr[0]
            row_idx = new_cell_addr[1] - 1
            tab_name = wb_def['wb_tabs'][tab_id]['tab_name']
            tab_cd_fix_summ = copy.deepcopy(wb_def['wb_tabs'][tab_id]['tab_cd_fixed_summ'])

            # if col or row is 0, use last row or next col
            for tab_def_key, tab_def_summ in tab_cd_fix_summ.items():
                if tab_def_key == 'Rows' or tab_def_key == 'x-uniq-rows':
                    continue

                if tab_def_summ[0] == 0:
                    col_idx += 1
                else:
                    row_idx += 1
                    col_idx = new_cell_addr[0]
                tab_def_summ[0] = col_idx
                tab_def_summ[1] = row_idx
                tab_def_summ[4] = 0   # all widths are set in the "Workbook Summary" line in the grid
                if tab_def_key == 'summ-title':
                    tab_def_summ[10] = tab_name
                summ_key = f"{tab_id}_{tab_def_key}"
                new_tab_fix_summ[summ_key] = tab_def_summ

        self.tab_def['tab_cd_fixed_summ'] = new_tab_fix_summ

class DefAr51(NewTab):
    def __init__(self, wb_obj):
        self.tab_id = 'ar51'
        self.colors = wb_obj.Colors
      # self.tab_common = wb_obj.tab_common
        super().__init__(self.tab_id, wb_obj)
        cfg  = wb_obj.cfg
        ctot = wb_obj.ctot
        tab_common = wb_obj.tab_common
        sea2 = self.colors.tbl_clrs['sea'][2]
        pur2 = self.colors.tbl_clrs['pur'][2]
        red2 = self.colors.tbl_clrs['red'][2]
        pur0 = self.colors.tbl_clrs['pur'][0]
        red0 = self.colors.tbl_clrs['red'][0]
        
        purt = self.colors.tbl_txts[pur0][0]
        redt = self.colors.tbl_txts[red0][0]

        file_key1 = f"{tab_common['file']['col_key1']}"
        xyml_key1 = f"{tab_common['xyml']['col_key1']}"
        xyml_val1 = f"{tab_common['xyml']['col_val1']}"
        dups_key1 = f"{tab_common['dups']['col_key1']}"

        file_tot1 = f'=_xlfn.AGGREGATE(3,3,tbl_file[{file_key1}])'
        file_tot1 = f'=COUNTA(_xlfn.UNIQUE(_xlfn.FILTER(tbl_file[{file_key1}],tbl_file[IsVisible])))*-1'
        xyml_tot1 = f'=_xlfn.AGGREGATE(3,3,tbl_xyml[{xyml_key1}])*(-1)'
        xyml_tot2 = f'=COUNTIF(tbl_xyml[{xyml_val1}],TRUE)'
        dups_tot1 = f'=_xlfn.AGGREGATE(3,3,tbl_dups[{dups_key1}])*(-1)'

        file_desc1 = f"    less Total {file_key1} from {tab_common['file']['tab_name']} tab"
        xyml_desc1 = f"    less Total {xyml_key1} from {tab_common['xyml']['tab_name']} tab"
        xyml_desc2 = f"    plus Total {xyml_val1} from {tab_common['xyml']['tab_name']} tab"
        dups_desc1 = f"    less Total {dups_key1} from {tab_common['dups']['tab_name']} tab"

        ctot_descs = [
                  '00-Total MD Files in Vault'
                , '01-Templates Processed'
                , '02-MD Files in dirs_skip_rel_str'
                , '03-Files Processed/Dups Teste'
                , '04-Known Nested Tags Files Found'
                , '05-Frontmatter YAML Files'
                , '06-Inline YAML Files'
                , '07-upd_obs_files'
                , '08-upd_obs_nests'
                , '09-upd_obs_props'
                , '10-Empty Fm/Body in Markdown'
                , '11-Max Props'
                , '12-Max Tags'
                ]

        # f'=COUNTA(_xlfn.UNIQUE(_xlfn.FILTER({self.tbl_name}[{self.col_key1}],{self.tbl_name}[{self.hdr_IsVis}])))'
        # self.f_txt_key1 = f'=_xlfn.AGGREGATE(3,3,{self.tbl_name}[{self.col_key1}])'
        # self.f_num_key2 = f'=_xlfn.AGGREGATE(9,3,{self.tbl_name}[{self.col_key2}])'

        clr1, txt1, clr2, txt2, table_style = self.colors.get_tab_clrs(self.tab_id)
        self.tab_def['borders'] = {
              'proc_3':["C24:C24", "double", self.colors.clr_blk,"top"]
        }
        self.tab_def['tab_table_style'] = table_style
        # clr1 = tab color,
        # clr2 = secondary "highlights" color, headings
        # clr1 = fill color on cells that use color fills
        # txt1 = text color on cells that use color fills

        sz = self.tab_txt_sz

        self.font_title_lst = ['Berlin Sans FB Demi', 24, clr1]
        self.font_subs_lst = ['Berlin Sans', 14, txt1]
        self.font_body_lst = ['Calibri', sz, txt1]
        self.tab_def['tab_color'] = clr1

        self.tab_def['tab_table_link_spcrs'] = True  # Always, TRUE for now

        self.tab_def['tab_has_isVisible_col'] = False
        self.tab_def['tab_tots_isVisible_col']  = 0
        self.tab_def['showGridLines'] = self.showGridLines

        self.tab_def['hdr_links_pfx'] = ""
        self.tab_def['tab_table_links_cols']    = 0
        self.tab_def['tbl_beg_col'] = 0

        sz = 11
        self.tab_def['tab_options'] = {
            "FullPath", True
        }
        # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
        self.tab_def['cfg-dump'] = [7, 2, '', sz, 20, "", "", False, False, 'left', '']

        self.tab_def['tab_cd_fixed_grid']   = {  # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
            # totals headers (across and down)
              'Ctrl-Key':   [2, 20, '', 14, 48,   '', sea2, True,  False, 'left', 'Control totals from v_chk']
            , 'Ctrl-Val':   [0,  0, '', 10, 11,   '', sea2, True,  False, 'right', 'Totals']
            , 'Run-Bal':    [0,  0, '', 10, 10,   '', sea2, True,  False, 'right', 'Run Tot']
            , 'Zero-Bal':   [0,  0, '', 10, 13,   '', sea2, True,  False, 'center', 'Zero Bal Chk']

            , 'f-tot-00':   [2, 21, '', 12,  0,   '',   '', False, False, 'left', ctot_descs[0]]
            , 'x-tot-00':   [0,  0, '', 12,  0,   '',   '', False, False, 'right', ctot[0]]
            , 'r-tot-00':   [0,  0, '', 12,  0,   '',   '', False, False, 'right', '=C21']

            , 'f-tot-01':   [2, 22, '', 12,  0,   '',   '', False, False, 'left', ctot_descs[1]]
            , 'x-tot-01':   [0,  0, '', 12,  0,   '',   '', False, False, 'right', -ctot[1]]
            , 'r-tot-01':   [0,  0, '', 12,  0,   '',   '', False, False, 'right', '=D21+C22']

            , 'f-tot-02':   [2, 23, '', 12,  0,   '',   '', False, False, 'left', ctot_descs[2]]
            , 'x-tot-02':   [0,  0, '', 12,  0,   '',   '', False, False, 'right', -ctot[2]]
            , 'r-tot-02':   [0,  0, '', 12,  0,   '',   '', False, False, 'right', '=D22+C23']

            , 'f-tot-03':   [2, 24, '', 12,  0,   '',   '', False, False, 'left', ctot_descs[3]]
            , 'x-tot-03':   [0,  0, '', 12,  0,   '',   '', False, False, 'right', ctot[3]]
            , 'x-total1':   [5,  0, '', 12,  0,   '',   '', False, False, 'center', '=sum(C20:C23)-C24']

            , 'file-notes': [2, 25, '', sz,  0, purt, pur0, False, True,  'left',  file_desc1]
            , 'file-tot-1': [0,  0, '', 12,  0,   '', pur2, False, False, 'right', file_tot1]
            , 'file-tot-r': [0,  0, '', 12,  0,   '',   '', False, False, 'right', '=D23+C25']

            , 'xyml-notes': [2, 26, '', sz,  0, redt, red0, False, True,  'left',  xyml_desc1]
            , 'xyml-tot-1': [0,  0, '', 12,  0,   '', red2, False, False, 'right', xyml_tot1]
            , 'xyml-run-1': [0,  0, '', 12,  0,   '',   '', False, False, 'right', '=D25+C26']

            , 'xyml-inFls': [2, 27, '', sz,  0, redt, red0, False, True,  'left',  xyml_desc2]
            , 'xyml-tot-2': [0,  0, '', 12,  0,   '', red2, False, False, 'right', xyml_tot2]
            , 'xyml-run-2': [0,  0, '', 12,  0,   '',   '', False, False, 'right', '=D26+C27']

            , 'dups-notes': [2, 28, '', sz,  0, redt, red0, False, True,  'left',  dups_desc1]
            , 'dups-tot-1': [0,  0, '', 12,  0,   '', red2, False, False, 'right', dups_tot1]
            , 'dups-run-1': [0,  0, '', 12,  0,   '',   '', False, False, 'right', '=D27+C28']
            , 'dups-bal-1': [0,  0, '', 12,  0,   '',   '', False, False, 'center', '=SUM(C24:C28)']

            , 'f-tot-04':   [2, 30, '', 12,  0,   '',   '', False, False, 'left', ctot_descs[4]]
            , 'x-tot-04':   [0, 30, '', 12,  0,   '',   '', False, False, 'right', ctot[4]]
            , 'f-tot-05':   [2, 31, '', 12,  0,   '',   '', False, False, 'left', ctot_descs[5]]
            , 'x-tot-05':   [0, 31, '', 12,  0,   '',   '', False, False, 'right', ctot[5]]
            , 'f-tot-06':   [2, 32, '', 12,  0,   '',   '', False, False, 'left', ctot_descs[6]]
            , 'x-tot-06':   [0, 32, '', 12,  0,   '',   '', False, False, 'right', ctot[6]]
            , 'f-tot-07':   [2, 33, '', 12,  0,   '',   '', False, False, 'left', ctot_descs[7]]
            , 'x-tot-07':   [0, 33, '', 12,  0,   '',   '', False, False, 'right', ctot[7]]
            , 'f-tot-08':   [2, 34, '', 12,  0,   '',   '', False, False, 'left', ctot_descs[8]]
            , 'x-tot-08':   [0, 34, '', 12,  0,   '',   '', False, False, 'right', ctot[8]]
            , 'f-tot-09':   [2, 35, '', 12,  0,   '',   '', False, False, 'left', ctot_descs[9]]
            , 'x-tot-09':   [0, 35, '', 12,  0,   '',   '', False, False, 'right', ctot[9]]
            , 'f-tot-10':   [2, 36, '', 12,  0,   '',   '', False, False, 'left', ctot_descs[10]]
            , 'x-tot-10':   [0, 36, '', 12,  0,   '',   '', False, False, 'right', ctot[10]]
            , 'f-tot-11':   [2, 37, '', 12,  0,   '',   '', False, False, 'left', ctot_descs[11]]
            , 'x-tot-11':   [0, 37, '', 12,  0,   '',   '', False, False, 'right', ctot[11]]
            , 'f-tot-12':   [2, 38, '', 12,  0,   '',   '', False, False, 'left', ctot_descs[12]]
            , 'x-tot-12':   [0, 38, '', 12,  0,   '',   '', False, False, 'right', ctot[12]]

            , 'cfg-keys':  [7, 1, '', 14, 0,   '', sea2, True, False, 'left', 'CFG Keys']
            , 'cfg-vals':  [8, 1, '', 14, 0,   '', sea2, True, False, 'left', 'Values']


        }


if __name__ == '__main__':
    # Build Tabs
    # cfg = WbDataDef()
    DBUG_LVL = 1

    tabs = NewWb(DBUG_LVL)

    # shelve_file = shelve.open("v_def.db")
    # shelve_file['v_def'] = v_def
    # shelve_file.close()

    # self.tab_def['tab_cd_table_hdr']["RowId"]
    if DBUG_LVL:
        lin = "=" * 30
        dict_list = {
              'cfg': tabs.wb_def['cfg']
            , 'wb_tabs': tabs.wb_def['wb_tabs']
            , 'wb_data': tabs.wb_def['wb_data']
            # , 'tab_def': tabs.wb_def['wb_tabs']['pros']
        }

        for p_dict_name, p_dict in dict_list.items():
            print(f"\n{p_dict_name}: {lin}")
            for k,v in p_dict.items():
                k_name = f"{p_dict_name}['{k}']"
                print(f"{k_name: <20}: {v}")

        print("Standalone run of v_chk_xl_tabs.py completed.")





