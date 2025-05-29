from openpyxl.styles import Side
from v_chk_cfg_data import *

class NewWb(WbDataDef):
    def __init__(self, dbug_lvl):
        self.DBUG_LVL = dbug_lvl
        self.tab_id = 'init'
        super().__init__(self.DBUG_LVL)
        self.wb_def = self.read_cfg_data()
        self.wb_tabs = self.wb_def['wb_tabs']
        self.tab_def = {}
        if self.DBUG_LVL >= 0:
            print("Loading Workbook Tab Definitions...")

        for tab_id_key in self.wb_tabs.keys():
            if self.DBUG_LVL > 3:
                print(f"Building '{tab_id_key}' tab definition")
            self.tab_def = self.wb_tabs[tab_id_key]
            if tab_id_key == 'pros':
                self.tab_def_obj = DefPros()
            elif tab_id_key == 'tags':
                self.tab_def_obj = DefTags()
            elif tab_id_key == 'dups':
                self.tab_def_obj = DefDups()
            elif tab_id_key == 'xyml':
                self.tab_def_obj = DefXyml()
            elif tab_id_key == 'ar51':
                self.tab_def_obj = DefAr51()
            elif tab_id_key == 'ar52':
                self.tab_def_obj = DefAr52()
            elif tab_id_key == 'summ':
                self.tab_def_obj = DefSumm()
            elif tab_id_key == 'init':
                pass
            else:
                raise Exception(f"Unexpected key: {tab_id_key} in {self.wb_tabs.keys()}")

            if tab_id_key != 'init':
                self.wb_def['wb_tabs'][tab_id_key] = self.tab_def_obj.tab_def

        self.write_cfg_data()

class NewTab:
    def __init__(self, tab_id):
        self.tab_id = tab_id

        # Fill and text colors for grid tab headings
        self.colors = Colors()
        self.bdr_thick = Side(border_style="thick", color='000000')
        self.bdr_thin = Side(border_style="thin", color='000000')
        self.tab_clrs = {  # tab color,         tab hdr colors
              'pros': [self.colors.clr_blud4, self.colors.clr_blud4]
            , 'tags': [self.colors.clr_ora64, self.colors.clr_ora64]
            , 'dups': [self.colors.clr_pur45, self.colors.clr_pur45]
            , 'summ': [self.colors.clr_grn30, self.colors.clr_grn30]
            , 'xyml': [self.colors.clr_aqu56, self.colors.clr_aqu56]
            , 'ar51': [self.colors.clr_red20, self.colors.clr_red20]
            , 'ar52': [self.colors.clr_red20, self.colors.clr_red20]
        }
        self.hdr_key_col = {
              'pros': "Property"
            , 'tags': "Tag Name"
            , 'dups': "MD Filename"
            , 'summ': "Property"
            , 'xyml': "MD Filename"
            , 'ar51': "MD Filename"
            , 'ar52': "MD Filename"
            , 'init': "Init only"
        }
        self.hdr_files_col = {
              'pros': "Files"
            , 'tags': "Files"
            , 'dups': "Duplicate Notes"
            , 'summ': "Files"
            , 'xyml': "MD Filename"
            , 'ar51': "MD Filename"
            , 'ar52': "MD Filename"
            , 'init': "Init only"
        }
        self.hdr_val_col = {
              'pros': "Values"
            , 'tags': ""
            , 'dups': ""
            , 'summ': "Values"
            , 'xyml': ""
            , 'ar51': "MD Filename"
            , 'ar52': "MD Filename"
            , 'init': "Init only"
        }
        # for prop_name, values_dict in sorted(self.wb_def['wb_data']['obs_props'].items()):
        #
        #     # filter props for those belonging to this tab only
        #     if tab_id == 'dups' and prop_name != 'xkey_dups':
        #             continue
        #     if tab_id == 'xyml' and 'xkey_xyml' not in prop_name:
        #             continue
        #     if tab_id == 'pros' and prop_name in ["tags", "tags:", 'xkey_dups', 'xkey_xyml']:
        #         continue
        #     if tab_id == 'tags' and prop_name != "tags" and prop_name != "tags:":
        #         continue
        #     if tab_id == 'summ' and prop_name in ["tags", "tags:", 'xkey_dups', 'xkey_xyml']:
        #         continue
        self.xkey_sub_keys = {
            # obs_props key, tab_id, 'KeyDesc', incl_filter, excl_filter]
              'xkey_dups': ['dups', 'Duplicates', ['xkey_dups'], [], ]
            , 'xkey_xyml_Non-Dictionary-YAML': ['xyml', 'Frontmatter Not in Dictionary Format', [], []]
        }
        self.xyaml_descs = {
              'NonD': 'YAML Not a valid Dictionary'
            , 'BadY': 'Invalid Properties in Frontmatter'
            , 'ErrY': 'Unknown YAML Exception'
            , 'NoFm': 'No Frontmatter'
            , 'MtFm': 'Empty Frontmatter'       # This should never happen?!
        }

        self.hdr_PVI = "P-V Index"
        self.hdr_IsVis = "IsVisible"
        self.hdr_links_pfx = "File"
        self.hdr_dups_found = "Dups Found"

        self.tab_clr_txt = self.colors.clr_blk00
        self.tab_txt_sz = 11
        self.tab_link_clr = self.colors.clr_blud4
        self.tab_fill_clr = self.colors.clr_wht00
        self.font_title_lst = ['Berlin Sans FB Demi', 24, self.colors.clr_blud4]
        self.font_subs_lst = ['Berlin Sans', 14, self.colors.clr_blk00]
        self.font_body_lst = ['Calibri'    , 11, self.colors.clr_blk00]
        self.cell_width = 8
        self.tbl_name = f"tbl_{self.tab_id}"
        self.tbl_hdr_row = 10
        self.tbl_beg_col = 4
        self.tbl_end_col = 0
        self.tbl_fix_cols = 0
        self.hdr_RowId = "RowId"
        self.f_warn_null = "Properties w/Empty Values Detected!"
        self.f_warn_filter = "Column filters applied--Totals now reflect column filters!"
        self.f_uniq_key    = f'=COUNTA(_xlfn.UNIQUE(_xlfn.FILTER({self.tbl_name}[{self.hdr_key_col[self.tab_id]}],{self.tbl_name}[{self.hdr_IsVis}])))'
        self.f_ctot_key    = f'=_xlfn.AGGREGATE(3,3,{self.tbl_name}[{self.hdr_key_col[self.tab_id]}])'
        self.f_uniq_val    = f'=COUNTA(_xlfn.UNIQUE(_xlfn.FILTER({self.tbl_name}[{self.hdr_val_col[self.tab_id]}],{self.tbl_name}[{self.hdr_IsVis}])))'
        self.f_ctot_val    = f'=COUNTA({self.tbl_name}[{self.hdr_val_col[self.tab_id]}])'
        self.f_uniq_files  = f'=_xlfn.AGGREGATE(9,3,{self.tbl_name}[{self.hdr_files_col[self.tab_id]}])' # needs work!
        self.f_ctot_files  = f'=_xlfn.AGGREGATE(9,3,{self.tbl_name}[{self.hdr_files_col[self.tab_id]}])'
        self.f_ctot_fnames = f'=COUNTA({self.tbl_name}[{self.hdr_files_col[self.tab_id]}])'
        self.f_ctot_dup_founds = f'=SUM({self.tbl_name}[{self.hdr_dups_found}])'
        self.f_null_values = f'=IFERROR(IF(_xlfn.AGGREGATE(3,3,{self.tbl_name}[{self.hdr_val_col[self.tab_id]}])<>SUM({self.tbl_name}[{self.hdr_IsVis}]),"{self.f_warn_null}",""),"")'
        self.f_filters_on  = f'=IFERROR(IF(COUNTA({self.tbl_name}[{self.hdr_RowId}])<>SUM({self.tbl_name}[{self.hdr_IsVis}]),"{self.f_warn_filter}",""),"")'
        self.f_isVisible   = f'=SUBTOTAL(3,@[{self.hdr_RowId}])'

        self.tab_table_files = {}
        self.tab_def = {
                      'tab_id': self.tab_id
                    , 'tab_clr_txt':        self.tab_clr_txt
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
                    , 'f_uniq_key':         self.f_uniq_key
                    , 'f_uniq_val':         self.f_uniq_val
                    , 'f_ctot_key':         self.f_ctot_key
                    , 'f_ctot_val':         self.f_ctot_val
                    , 'f_ctot_files':       self.f_ctot_files
                    , 'f_null_values':      self.f_null_values
                    , 'f_filters_on':       self.f_filters_on
                    , 'f_isVisible':        self.f_isVisible
                    , 'tab_name': ''
                    , 'tab_comments': []  # up to 3 lines
                    , 'tab_notes': []
                                     # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
                    , 'tab_title_def':    [3,  2, 'Berlin Sans FB Demi', 24, 0, '', '', True, False, 'left', '']
                    , 'tab_comments_def': [3,  3, '', 10, 0, '', '', True, False, 'left', '']
                    , 'tab_notes_def':    [3, 22, '', 10, 0, '', '', True, False, 'left', '']
                    , 'tab_table_style': 'TableStyleMedium20'
                    , 'tab_color': self.colors.clr_blud4
                    # cd is short for cell_def
                    , 'tab_cd_table_hdr': {}  # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
                    , 'tab_cd_table_dtl': {}
                    , "tab_table_links":  [9, 0, '', 10, 18, "", "", False, False, 'left'  , '']
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

    def files_cols_def(self):
        # If there are add`l columns for Files "Where used", add them to the tab_def
        if self.tab_id and self.tab_def['tab_table_links_cols'] > 0:
            tab_cd_table_hdr = self.tab_def['tab_cd_table_hdr']
            tab_cd_table_dtl = self.tab_def['tab_cd_table_dtl']
            tab_cd_table_hdr, tab_cd_table_dtl = self.set_table_links(tab_cd_table_hdr, tab_cd_table_dtl)
            self.tab_def['tab_cd_table_hdr'] = tab_cd_table_hdr
            self.tab_def['tab_cd_table_dtl'] = tab_cd_table_dtl

    def calc_col_pointers(self):
        first_col_list = []
        spcr_cols = 0
        if self.tab_def['tab_table_link_spcrs']:
            spcr_cols = self.tab_def['tab_table_links_cols']

        for k, v in self.tab_def['tab_cd_table_hdr'].items():
            col_num = int(v[0])
            first_col_list.append(col_num)
        self.tab_def['tbl_beg_col'] = min(first_col_list)
        self.tab_def['tbl_fix_cols'] = len(self.tab_def['tab_cd_table_hdr'])
        self.tab_def['tbl_end_col'] = ((self.tab_def['tab_table_links_cols'] + spcr_cols)
            + (self.tab_def['tbl_fix_cols'] - 1)
            + (self.tab_def['tbl_beg_col']))
        if self.tab_def['tbl_end_col'] < self.tab_def['tab_tots_isVisible_col']:
            self.tab_def['tbl_end_col'] = self.tab_def['tab_tots_isVisible_col']

        if self.tab_def['tab_has_isVisible_col']:
            self.tab_def['tab_tots_isVisible_col'] = int(self.tab_def['tbl_end_col'])    #  38

    def set_table_links(self, tab_cd_table_hdr, tab_cd_table_dtl):
        tab_table_links = self.tab_def['tab_table_links']      # [9, 18, "", "", False, 'left']
        tab_cd_table_spacer = self.tab_def['tab_cd_table_spacer']
        tab_table_link_spcrs = self.tab_def['tab_table_link_spcrs']
        tab_table_links_cols = self.tab_def['tab_table_links_cols']
        hdr_links_pfx = self.tab_def['hdr_links_pfx']

        col_idx = tab_table_links[0]
        for i in range(1, tab_table_links_cols + 1):                           # [col,    w, txt, fill, bold, align
            # tab_table_links[0] = col_idx
            # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
            link_key_and_val = f"{hdr_links_pfx}{i:02d}"
            spcr_key_and_val = f"S{i}"
            tab_cd_table_hdr[link_key_and_val] = [col_idx] + tab_table_links[1:10] + [link_key_and_val]
            tab_cd_table_dtl[link_key_and_val] = [col_idx] + tab_table_links[1:]
            col_idx += 1

            if tab_table_link_spcrs:
                tab_cd_table_hdr[spcr_key_and_val] = [col_idx] + tab_cd_table_spacer[1:10] + [f"{spcr_key_and_val}  "]
                tab_cd_table_dtl[spcr_key_and_val] = [col_idx] + tab_cd_table_spacer[1:]
                col_idx += 1

        return tab_cd_table_hdr, tab_cd_table_dtl

class DefPros(NewTab):
    def __init__(self):
        self.tab_id = 'pros'
        super().__init__(self.tab_id)
        clrtb, clrhi = self.tab_clrs[self.tab_id]

        clrfl = self.colors.Code_LOV[clrhi][0]
        clrtx = self.colors.Code_LOV[clrhi][1]
        sz = self.tab_txt_sz
        self.tab_def['tab_table_style'] = 'TableStyleMedium6'

        self.font_title_lst = ['Berlin Sans FB Demi', 24, clrtb]
        self.font_subs_lst = ['Berlin Sans', 14, clrtx]
        self.font_body_lst = ['Calibri', sz, clrtx]
        self.tab_def['tab_color'] = clrtb

        self.tab_def['tab_table_link_spcrs'] = True  # Always, TRUE for now
        self.tab_def['tab_txt_sz'] = sz
        self.tab_def['showGridLines'] = True

        self.tab_def['hdr_links_pfx'] = "File"
        self.tab_def['tab_table_links_cols'] = 15
        self.tab_def['tab_has_isVisible_col'] = True
        self.tab_def['tab_tots_isVisible_col'] = 38

        self.tab_def['tab_name'] = 'Properties'
        self.tab_def['tab_title_def'] = [3, 2, 'Berlin Sans FB Demi', 24, 0, clrhi, '', True, False, 'left', "Properties and Values Analysis"]
        self.tab_def['tab_comments_def'] = [ 3, 3, '', 10, 0, '', '', False,  False, 'left', '']
        self.tab_def['tab_notes_def'] =    [ 3, 22, '', 10, 0, '', '', False,  False, 'left', '']
        self.tab_def['tab_notes'] = []
        self.tab_def['tab_comments'] = [
             'Totals at right will reflect column filters. This can be useful'
            ,'to perform analysis on specific properties, and values.'
            ]  # up to 3 lines

        self.tab_def['tab_cd_table_hdr'] = {
            # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
             "RowId":  [ 4, 10, '', sz,  8, "", "", True, False, 'center', self.hdr_RowId]
           , "Prop": [ 5, 10, '', sz, 18, "", "", True, False, 'left'  , self.hdr_key_col[self.tab_id]]
           , "Val":  [ 6, 10, '', sz, 50, "", "", True, False, 'left'  , self.hdr_val_col[self.tab_id]]
           , "FCnt": [ 7, 10, '', sz,  8, "", "", True, False, 'center', self.hdr_files_col[self.tab_id]]
           , "PVI":  [ 8, 10, '', sz, 13, "", "", True, False, 'center', self.hdr_PVI]
           }
        self.tab_def['tab_cd_table_dtl'] =  {
            # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
             "RowId":  [ 4,  0, '', sz,  8, "", "", False, False, 'center', '']
           , "Prop": [ 5,  0, '', sz, 18, "", "", True, False,  'left', '']
           , "Vals": [ 6,  0, '', sz, 50, "", "", False, False, 'left', '']
           , "FCnt": [ 7,  0, '', sz,  8, "", "", False, False, 'center', '']
           , "PVI":  [ 8,  0, '', sz, 13, "", "", False, False, 'center', '']
           }
        self.tab_def['tab_table_links']  = [ 9, 0, '', sz, 18, "", "", False, False, 'left'  , '']
        self.tab_def['tab_cd_table_spacer'] = [10, 0, '', sz,  1, "", "", False, False, 'right'  , '']
        self.tab_def['tab_cd_fixed_grid'] = {
            # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
            # totals headers (across and down)
             'Analysis':   [ 9, 2, '', 14, 0, clrtx, clrfl, True,  False, 'center', 'Analysis']
          ,  'S1  ':   [10, 2, '', 12, 0, clrfl, clrfl, True,  False, 'right', 'Spc1  ']
          ,  'Prop':   [11, 2, '', 12, 0, clrtx, clrfl, True,  False, 'center', self.hdr_key_col[self.tab_id]]
          ,  'S2  ':   [12, 2, '', 12, 0, clrfl, clrfl, True,  False, 'right', 'Spc2  ']
          ,  'Val':    [13, 2, '', 12, 0, clrtx, clrfl, True,  False, 'center', self.hdr_val_col[self.tab_id]]
          ,  'S3  ':   [14, 2, '', 12, 0, clrfl, clrfl, True,  False, 'right', 'Spc3  ']
          ,  'Files Used':   [15, 2, '', 12, 0, clrtx, clrfl, True,  False, 'center', self.hdr_files_col[self.tab_id]]
          ,  'Unique Values': [ 9, 3, '', 12, 0, clrtx, clrfl, True,  False, 'left', 'Unique Values']
          ,  'Column Totals': [ 9, 4, '', 12, 0, clrtx, clrfl, True,  False, 'left', 'Column Totals']
          ,  'x-uniq-key':   [11, 3, '', sz, 0, "",              "", False, False, 'center', self.f_uniq_key]
          ,  'x-uniq-val':    [13, 3, '', sz, 0, "",              "", False, False, 'center', self.f_uniq_val]
          ,  'x-ctot-key':   [11, 4, '', sz, 0, "",              "", False, False, 'center', self.f_ctot_key]
          ,  'x-ctot-val':    [13, 4, '', sz, 0, "",              "", False, False, 'center', self.f_ctot_val]
          ,  'x-ctot-files':  [15, 4, '', sz, 0, "",              "", False, False, 'center', self.f_ctot_files]
          , 'x-null-values':  [ 6, 9, '', 12, 0, self.colors.clr_red20,  "", True, True,   'left',   self.f_null_values]
          , 'x-filters-on':   [ 9, 6, '', 12, 0, self.colors.clr_red20,  "", True, True,   'left',   self.f_filters_on]
          # This one is for the IVisible Column, not the totals
          , 'isVisible':      [38, 0, '', 10, 0, "",              "", False, False, 'right',  self.f_isVisible]
          }

        self.tab_def_post()

class DefTags(NewTab):
    def __init__(self):
        self.tab_id = 'tags'
        super().__init__(self.tab_id)
        clrtb, clrhi = self.tab_clrs[self.tab_id]

        clrfl = self.colors.Code_LOV[clrhi][0]
        clrtx = self.colors.Code_LOV[clrhi][1]
        sz = self.tab_txt_sz
        self.tab_def['tab_table_style'] = 'TableStyleMedium7'

        self.font_title_lst = ['Berlin Sans FB Demi', 24, clrtb]
        self.font_subs_lst = ['Berlin Sans', 14, clrtx]
        self.font_body_lst = ['Calibri', sz, clrtx]
        self.tab_def['tab_color'] = clrtb

        self.tab_def['tab_table_link_spcrs'] = True  # Always, TRUE for now

        self.tab_def['tab_has_isVisible_col'] = True
        self.tab_def['tab_tots_isVisible_col'] = 0

        self.tab_def['showGridLines'] = True

        self.tab_def['hdr_links_pfx'] = "File"
        self.tab_def['tab_table_links_cols']    = 10

        self.tab_def['tab_name'] = 'Tags'
        self.tab_def['tab_title_def'] = [3, 2, 'Berlin Sans FB Demi', 24, 0, clrhi, '', True, False, 'left', "Tag Analysis"]
        self.tab_def['tab_comments_def'] = [ 3, 3, '', 10, 0, '', '', False,  False, 'left', '']
        self.tab_def['tab_notes_def'] = [ 3, 22, '', 10, 0, '', '', False,  False, 'left', '']
        self.tab_def['tab_comments'] = [
              'All tags found in all markdown files.'
            , 'Totals at right will reflect column filters. This can be useful'
            , 'to perform analysis on specific properties, and values.']  # up to 3 lines
        self.tab_def['tab_notes'] = []

        self.tab_def['tab_cd_table_hdr'] = {  # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
              "RowId": [4, 10, '', sz, 8, "", "", True, False, 'center', self.hdr_RowId]
            , "Tag Name": [5, 10, '', sz, 18, "", "", True, False, 'left', self.hdr_key_col[self.tab_id]]
            , "FCnt": [6, 10, '', sz, 8, "", "", True, False, 'center', self.hdr_files_col[self.tab_id]]
        }
        self.tab_def['tab_cd_table_dtl'] = {  # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
            "RowId": [4, 0, '', sz, 8, "", "", False, False, 'center', '']
            , "Tags": [5, 0, '', sz, 18, "", "", True, False, 'left', '']
            , "FCnt": [6, 0, '', sz, 8, "", "", False, False, 'center', '']
        }
        self.tab_def['tab_table_links'] = [7, 0, '', sz, 18, "", "", False, False, 'left', '']
        self.tab_def['tab_cd_table_spacer'] = [8, 0, '', sz, 1, "", "", False, False, 'right', '']
        self.tab_def['tab_cd_fixed_grid'] = {  # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
            # totals headers (across and down)
            'Analysis': [9, 2, '', 14, 0, clrtx, clrfl, True, False, 'center', 'Analysis']
            , 'Spc1  ': [10, 2, '', 12, 0, clrfl, clrfl, True, False, 'right', 'Spc1  ']
            , 'Property': [11, 2, '', 12, 0, clrtx, clrfl, True, False, 'center', self.hdr_key_col[self.tab_id]]
            , 'Spc2  ': [12, 2, '', 12, 0, clrfl, clrfl, True, False, 'right', 'Spc2  ']
            , 'FCnt': [13, 2, '', 12, 0, clrtx, clrfl, True, False, 'center', self.hdr_files_col[self.tab_id]]
            , 'UVals': [9, 3, '', 12, 0, clrtx, clrfl, True, False, 'left', 'Unique Values']
            , 'CTot': [9, 4, '', 12, 0, clrtx, clrfl, True, False, 'left', 'Column Totals']

            , 'x-uniq-key': [11, 3, '', sz, 0, "", "", False, False, 'center', self.f_uniq_key]
            , 'x-ctot-key': [11, 4, '', sz, 0, "", "", False, False, 'center', self.f_ctot_key]
            , 'x-ctot-files': [13, 4, '', sz, 0, "", "", False, False, 'center', self.f_ctot_files]
            , 'x-filters-on': [9, 6, '', 12, 0, self.colors.clr_red20, "", True, True, 'left', self.f_filters_on]
            # This one is for the IVisible Column, not the totals
            , 'isVisible': [26, 0, '', 10, 0, "", "", False, False, 'right', self.f_isVisible]
        }

        self.tab_def_post()

class DefSumm(NewTab):
    def __init__(self):
        self.tab_id = 'summ'
        super().__init__(self.tab_id)
        clrtb, clrhi = self.tab_clrs[self.tab_id]

        clrfl = self.colors.Code_LOV[clrhi][0]
        clrtx = self.colors.Code_LOV[clrhi][1]
        sz = 12
        self.tab_def['tab_table_style'] = 'TableStyleMedium4'

        self.font_title_lst = ['Berlin Sans FB Demi', 24, clrtb]
        self.font_subs_lst = ['Berlin Sans', 14, clrtx]
        self.font_body_lst = ['Calibri', sz, clrtx]
        self.tab_def['tab_color'] = clrtb

        self.tab_def['tab_table_link_spcrs'] = True  # Always, TRUE for now

        self.tab_def['tab_has_isVisible_col'] = True
        self.tab_def['tab_tots_isVisible_col']  = 0
        self.tab_def['showGridLines'] = False

        self.tab_def['hdr_links_pfx'] = ""
        self.tab_def['tab_table_links_cols']    = 0

        _, clr_p_hi = self.tab_clrs['pros']
        _, clr_t_hi = self.tab_clrs['tags' ]
        _, clr_d_hi = self.tab_clrs['dups' ]
        _, clr_x_hi = self.tab_clrs['xyml' ]
        _, clr_s_hi = self.tab_clrs['summ' ]

        clr_p_fl, clr_p_tx, _ = self.colors.Code_LOV[clr_p_hi]
        clr_t_fl, clr_t_tx, _ = self.colors.Code_LOV[clr_t_hi]
        clr_d_fl, clr_d_tx, _ = self.colors.Code_LOV[clr_d_hi]
        clr_x_fl, clr_x_tx, _ = self.colors.Code_LOV[clr_x_hi]

        self.tab_def['tab_name'] = 'Summary'
        self.tab_def['tab_title_def'] = [3, 2, 'Berlin Sans FB Demi', 24, 60, clrhi, '', True, False, 'left', "Obsidian Vault Healthcheck v1.0"]
        self.tab_def['tab_comments'] = [
                  'Lists All properties and Tags found in an Obsidian Vault, with  a sample number '
                , 'of links back to the markdown files where they are used. Duplicate Markdown files'
                , 'existing in different folders are also detected.'
                , '']  # up to 3 lines
        self.tab_def['tab_notes'] = [
              ' - Inline Properties and Tags will end with a ":" and appear in bold/italics.'
            , '      This was actually a bug that I decided to turn into a feature, for now.;-)'
            , '      Version 2 will handle inline P+T, properly, grouping them and having'
            , '      separate totals.'
            , ' - You can use Table Heading Filters to look at specific tabs.'
            , '      When filters are applied, tab totals will reflect the filtered data and'
            , '      a warning will display next to the tab Totals Section'
            , ' - The "Total" column reflects the total number of links found in all markdown files.'
            , ' - All Properties and Tags are listed in lowercase, as that is how Obsidian sees them.'
            , '     The FileDetails Tab shows them as entered, if lowercase was not used.'
            ]

        self.tab_def['tab_cd_table_hdr'] = {
                              # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
              "RowId":          [10, 10, '', sz, 12, "", "", True, False, 'center', self.hdr_RowId]
            , "Prop":           [11, 10, '', sz, 18, "", "", True, False, 'left', self.hdr_key_col[self.tab_id]]
            , "Values Count":   [12, 10, '', sz, 12, "", "", True, False, 'center', self.hdr_val_col[self.tab_id]]
            , "File Count":     [13, 10, '', sz, 12, "", "", True, False, 'center', self.hdr_files_col[self.tab_id]]
        }
        self.tab_def['tab_cd_table_dtl'] = {  # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
              "RowId":          [10, 10, '', sz,  0, "", "", False, False, 'center', self.hdr_RowId]
            , "Prop":           [11, 10, '', sz,  0, "", "", False, False, 'left', self.hdr_key_col[self.tab_id]]
            , "Values Count":   [12, 10, '', sz,  0, "", "", False, False, 'center', self.hdr_val_col[self.tab_id]]
            , "File Count":     [13, 10, '', sz,  0, "", "", False, False, 'center', self.hdr_files_col[self.tab_id]]
        }
        self.tab_def['tab_table_links']     = [0, 0, '', 0, 0, "", "", False, False, '', '']
        self.tab_def['tab_cd_table_spacer'] = [0, 0, '', 0, 0, "", "", False, False, 'right', '']
        self.tab_def['tab_comments_def']    = [3, 3, '', 10, 0, '', '', False, False, 'left', '']
        self.tab_def['tab_notes_def']       = [3, 24, '', 10, 0, '', '', False, False, 'left', '']
        self.tab_def['tab_cd_fixed_grid']   = {  # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
            # totals headers (across and down)
              'Table Totals': [10, 2, '', 14, 0, clrtx, clrfl, True, False, 'center', 'Analysis']
            , 'Property':     [11, 2, '', 12, 0, clrtx, clrfl, True, False, 'center', self.hdr_key_col[self.tab_id]]
            , 'Values':       [12, 2, '', 12, 0, clrtx, clrfl, True, False, 'center', self.hdr_val_col[self.tab_id]]
            , 'FCnt':         [13, 2, '', 12, 0, clrtx, clrfl, True, False, 'center', self.hdr_files_col[self.tab_id]]
            , 'UVals':        [10, 3, '', 12, 0, clrtx, clrfl, True, False, 'left', 'Unique']
            , 'CTot':         [10, 4, '', 12, 0, clrtx, clrfl, True, False, 'left', 'Total']
            , 'Notes':        [3, 23, '', 12, 0, clrtx, clrfl, True, False, 'left', 'Notes']

            , 'x-uniq-key':   [11, 3, '', sz, 0, "", "", False, False, 'center', self.f_uniq_key]
            , 'x-ctot-key':   [11, 4, '', sz, 0, "", "", False, False, 'center', self.f_ctot_key]
            , 'x-ctot-val':   [12, 4, '', sz, 0, "", "", False, False, 'center', self.f_ctot_val]
            , 'x-ctot-files': [13, 4, '', sz, 0, "", "", False, False, 'center', self.f_ctot_files]
            , 'x-filters-on': [10, 6, '', 12, 0, self.colors.clr_red20, "", True, True, 'left', self.f_filters_on]
            # This one is for the IVisible Column, not the totals
            , 'isVisible':    [14, 0, '', 10, 1, "", "", False, False, 'right', self.f_isVisible]

            # Left side - WORKBOOK Totals
            #                   [col,row,font,sz, w, t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
            , 'Workbook Totals': [ 3, 10, '', 14, 35, clrtx, clrfl,      True, False, 'left',    'Workbook Summary']
            , 'wb_Total':        [ 4, 10, '', 11, 0, clrtx, clrfl,       True, False, 'right',  "Count"]
            , 'wb_Uniq':         [ 5, 10, '', 11, 0, clrtx, clrfl,       False, False, 'right',  "Unique"]
            , 'wb_Property':     [ 3, 11, '', 11, 0, clr_p_tx, clr_p_fl, False, False, 'right',   f"{self.hdr_key_col['pros']} : "]
            , 'wb_Prop_Vals':    [ 3, 12, '', 11, 0, clr_p_tx, clr_p_fl, False, False, 'right',   f"{self.hdr_val_col['pros']} : "]
            , 'wb_Prop_Files':   [ 3, 13, '', 11, 0, clr_p_tx, clr_p_fl, False, False, 'right',   f"{self.hdr_files_col['pros']} : "]
            , 'wb_Tags':         [ 3, 15, '', 11, 0, clr_t_tx, clr_t_fl, False, False, 'right',   f"{self.hdr_key_col['tags']} : "]
            , 'wb_Tag_Files':    [ 3, 16, '', 11, 0, clr_t_tx, clr_t_fl, False, False, 'right',   f"{self.hdr_files_col['tags']} : "]
            , 'wb_Duplicates':   [ 3, 18, '', 11, 0, clr_d_tx, clr_d_fl, False, False, 'right',   f"{self.hdr_key_col['dups']} : "]
            , 'wb_Xyaml':        [ 3, 19, '', 11, 0, clr_x_tx, clr_x_fl, False, False, 'right',   f"{self.hdr_key_col['xyml']} : "]

            , 'x_p_keys_ctot':   [ 4, 11, '', 11,  0, "", "",             False, False, 'right', "=Properties!K4"]
            , 'x_p_keys_uniq':   [ 5, 11, '', 11,  0, "", "",             False, False, 'right', "=Properties!K3"]
            , 'x_p_vals_ctot':   [ 4, 12, '', 11,  0, "", "",             False, False, 'right', "=Properties!M4"]
            , 'x_p_vals_uniq':   [ 5, 12, '', 11,  0, "", "",             False, False, 'right', "=Properties!M3"]
            , 'x_p_file_ctot':   [ 4, 13, '', 11,  0, "", "",             False, False, 'right', "=Properties!O4"]
            , 'x_t_vals_ctot':   [ 4, 15, '', 11,  0, "", "",             False, False, 'right', "=Tags!K4"]
            , 'x_t_vals_uniq':   [ 5, 15, '', 11,  0, "", "",             False, False, 'right', "=Tags!K3"]
            , 'x_t_file_ctot':   [ 4, 16, '', 11,  0, "", "",             False, False, 'right', "=Tags!M4"]
            , 'x_d_file_ctot':   [ 4, 18, '', 11,  0, "", "",             False, False, 'right', f"=COUNTA(tbl_dups[{self.hdr_key_col['dups']}])"]
            , 'x_x_file_ctot':   [ 4, 19, '', 11,  0, "", "",             False, False, 'right', f"=COUNTA(tbl_xyml[{self.hdr_key_col['xyml']}])"]

        }

        self.tab_def_post()

class DefDups(NewTab):
    def __init__(self):
        self.tab_id = 'dups'
        super().__init__(self.tab_id)
        clrtb, clrhi = self.tab_clrs[self.tab_id]

        clrfl = self.colors.Code_LOV[clrhi][0]
        clrtx = self.colors.Code_LOV[clrhi][1]
        sz = self.tab_txt_sz
        self.tab_def['tab_table_style'] = 'TableStyleMedium19'

        self.font_title_lst = ['Berlin Sans FB Demi', 24, clrtb]
        self.font_subs_lst = ['Berlin Sans', 14, clrtx]
        self.font_body_lst = ['Calibri', sz, clrtx]
        self.tab_def['tab_color'] = clrtb

        self.tab_def['tab_table_link_spcrs'] = True  # Always, TRUE for now

        self.tab_def['tab_has_isVisible_col'] = False
        self.tab_def['tab_tots_isVisible_col']  = 0
        self.tab_def['showGridLines'] = True

        self.tab_def['hdr_links_pfx'] = "Full Pathnames"
        self.tab_def['tab_table_links_cols']    = 4

        self.tab_def['tab_name'] = 'Dups'
        self.tab_def['tab_title_def'] = [3, 2, 'Berlin Sans FB Demi', 24, 80, clrhi, '', True, False, 'left', "Vault Duplicate Markdown Filenames"]
        self.tab_def['tab_comments_def'] = [ 3, 3, '', 10, 0, '', '', False,  False, 'left', '']
        self.tab_def['tab_notes_def'] = [ 3, 22, '', 10, 0, '', '', False,  False, 'left', '']
        self.tab_def['tab_comments'] = [
              'Duplicates are vault files that have identical filenames,'
            , 'but exist in different folders.' ]  # up to 3 lines
             # 123456789 1 2345678 2 2345678 3 2345678 4 2345678 5 2345678 6 2345678 7 2345678 8

        self.tab_def['tab_notes'] = [
            'Duplicate markdown filenames are allowed in Obsidian, but should be avoided as they can'
            , 'be a source of confusion. They also require fully qualified pathnames, instead of just'
            , 'the filename, when attempting to create links.' ]

        self.tab_def['tab_cd_table_hdr'] = {
                        # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
             "RowId":       [6, 10, '', sz, 8, "", "", True, False, 'center', self.hdr_RowId]
            , "Filename":   [7, 10, '', sz, 35, "", "", True, False, 'left', self.hdr_files_col[self.tab_id]]
            , "Dups Found": [8, 10, '', sz, 8, "", "", True, False, 'left', self.hdr_dups_found]
        }
        self.tab_def['tab_cd_table_dtl'] = {  # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
             "RowId":       [6, 0, '', sz, 0, "", "", False, False, 'center', '']
            , "Filename":   [7, 0, '', sz, 0, "", "", True, False, 'left', '']
            , "Dups Found": [8, 0, '', sz, 0, "", "", False, False, 'center', '']
        }
        self.tab_def['tab_table_links']     = [9, 10, '', sz, 25, "", "", False, False, 'left', '']
        self.tab_def['tab_cd_table_spacer'] = [10,  0, '', sz,  1, "", "", False, False, 'right', '']
        self.tab_def['tab_cd_fixed_grid'] = {
            # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
            # totals headers (across, then down)
              'Analysis':       [7, 2, '', 14, 0, clrtx, clrfl, True, False, 'center', 'Analysis']
            , 'Dup-Notes-Top':  [8, 2, '', 12, 0, clrtx, clrfl, True, False, 'left', 'Paths']
            , 'Dup-Notes':      [7, 3, '', 12, 0, clrtx, clrfl, True, False, 'left', self.hdr_files_col[self.tab_id]]
            , 'Full-Pathnames': [7, 4, '', 12, 0, clrtx, clrfl, True, False, 'left', self.tab_def['hdr_links_pfx']]
            , 'x-ctot-files':   [8, 3, '', sz, 0, "", "", False, False, 'center', self.f_ctot_fnames ] #f_ctot_fnames
            , 'x-ctot-dups':    [8, 4, '', sz, 0, "", "", False, False, 'center', self.f_ctot_dup_founds]
            , 'Notes':          [3, 21, '', 12, 0, clrtx, clrfl, True, False, 'left', 'Notes:']
        }

        self.tab_def_post()

class DefXyml(NewTab):
    def __init__(self):
        self.tab_id = 'xyml'
        super().__init__(self.tab_id)
        clrtb, clrhi = self.tab_clrs[self.tab_id]

        clrfl = self.colors.Code_LOV[clrhi][0]
        clrtx = self.colors.Code_LOV[clrhi][1]
        sz = self.tab_txt_sz
        self.tab_def['tab_table_style'] = 'TableStyleMedium20'

        self.font_title_lst = ['Berlin Sans FB Demi', 24, clrtb]
        self.font_subs_lst = ['Berlin Sans', 14, clrtx]
        self.font_body_lst = ['Calibri', sz, clrtx]
        self.tab_def['tab_color'] = clrtb

        self.tab_def['tab_table_link_spcrs'] = True  # Always, TRUE for now

        self.tab_def['tab_has_isVisible_col'] = False
        self.tab_def['tab_tots_isVisible_col']  = 0
        self.tab_def['showGridLines'] = True

        self.tab_def['hdr_links_pfx'] = ""
        self.tab_def['tab_table_links_cols'] = 0

        self.tab_def['tab_name'] = 'Xyml'
        self.tab_def['tab_title_def'] = [3, 2, 'Berlin Sans FB Demi', 24, 80, clrhi, '', True, False, 'left', "Possibly Corrupt Markdown Files"]
        self.tab_def['tab_comments_def'] = [ 3, 3, '', 10, 0, '', '', False,  False, 'left', '']
        self.tab_def['tab_notes_def'] = [ 3, 22, '', 10, 0, '', '', False,  False, 'left', '']
        self.tab_def['tab_comments'] = [
              'Corrupt, meaning the Python package PyYAML 6.0.2 is unable to perform a "safe_load" of'
            , 'the notes frontmatter. This may or may not indicate an error, depending on how you setup'
            , 'and use your vault.'
            ]
             # 123456789 1 2345678 2 2345678 3 2345678 4 2345678 5 2345678 6 2345678 7 2345678 8
        self.tab_def['tab_notes'] = [
              'This system was intentionally designed to be "generous" when it comes to identifying '
            , 'what might be an error. Our purpose is only to bring issues to ones attention that'
            , 'may, (or may not) require further investigation.'
            ]
        self.tab_def['tab_cd_table_hdr'] = {  # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
             "RowId":         [6, 10, '', sz, 8, "", "", True, False, 'center', self.hdr_RowId]
            , "Filename":   [7, 10, '', sz, 80, "", "", True, False, 'left', self.hdr_files_col[self.tab_id]]
        }
        self.tab_def['tab_cd_table_dtl'] = {  # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
             "RowId":         [6, 0, '', sz, 0, "", "", False, False, 'center', '']
            , "Filename":   [7, 0, '', sz, 0, "", "", True, False, 'left', '']
        }
        self.tab_def['tab_table_links']  = [7, 10, '', sz, 40, "", "", False, False, 'left', '']
        self.tab_def['tab_cd_table_spacer'] = [8,  0, '', sz,  1, "", "", False, False, 'right', '']
        self.tab_def['tab_cd_fixed_grid'] = {
                          # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
            'Analysis':       [5, 2, '', 14, 16, clrtx, clrfl, True, False, 'left', 'Analysis']
            , 'Total':        [6, 2, '', 12,  0, clrtx, clrfl, True, False, 'center', '']
            , 'Files':        [5, 3, '', 12,  0, clrtx, clrfl, True, False, 'left', 'Total Files']
            , 'x-ctot-files': [6, 3, '', sz,  0, "", "", False, False, 'center', self.f_ctot_fnames ]
            , 'Notes': [3, 21, '', 12, 0, clrtx, clrfl, True, False, 'left', 'Notes:']
        }

        self.tab_def_post()

class DefFiles(NewTab):
    def __init__(self):
        self.tab_id = 'file'
        super().__init__(self.tab_id)
        clrtb, clrhi = self.tab_clrs[self.tab_id]

        clrfl = self.colors.Code_LOV[clrhi][0]
        clrtx = self.colors.Code_LOV[clrhi][1]
        sz = self.tab_txt_sz
        self.tab_def['tab_table_style'] = 'TableStyleMedium20'

        self.font_title_lst = ['Berlin Sans FB Demi', 24, clrtb]
        self.font_subs_lst = ['Berlin Sans', 14, clrtx]
        self.font_body_lst = ['Calibri', sz, clrtx]
        self.tab_def['tab_color'] = clrtb

        self.tab_def['tab_table_link_spcrs'] = True  # Always, TRUE for now
        self.tab_def['tab_txt_sz'] = sz
        self.tab_def['showGridLines'] = True

        self.tab_def['hdr_links_pfx'] = "File"
        self.tab_def['tab_table_links_cols'] = 15
        self.tab_def['tab_has_isVisible_col'] = True
        self.tab_def['tab_tots_isVisible_col'] = 38

        self.tab_def['tab_name'] = 'Properties'
        self.tab_def['tab_title_def'] = [3, 2, 'Berlin Sans FB Demi', 24, 0, clrhi, '', True, False, 'left', "Properties and Values Analysis"]
        self.tab_def['tab_comments_def'] = [ 3, 3, '', 10, 0, '', '', False,  False, 'left', '']
        self.tab_def['tab_notes_def'] =    [ 3, 22, '', 10, 0, '', '', False,  False, 'left', '']
        self.tab_def['tab_notes'] = []
        self.tab_def['tab_comments'] = [
             'Totals at right will reflect column filters. This can be useful'
            ,'to perform analysis on specific properties, and values.'
            ]  # up to 3 lines

        self.tab_def['tab_cd_table_hdr'] = {
            # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
             "RowId":  [ 4, 10, '', sz,  8, "", "", True, False, 'center', self.hdr_RowId]
           , "Prop": [ 5, 10, '', sz, 18, "", "", True, False, 'left'  , self.hdr_key_col[self.tab_id]]
           , "Val":  [ 6, 10, '', sz, 50, "", "", True, False, 'left'  , self.hdr_val_col[self.tab_id]]
           , "FCnt": [ 7, 10, '', sz,  8, "", "", True, False, 'center', self.hdr_files_col[self.tab_id]]
           , "PVI":  [ 8, 10, '', sz, 13, "", "", True, False, 'center', self.hdr_PVI]
           }
        self.tab_def['tab_cd_table_dtl'] =  {
            # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
             "RowId":  [ 4,  0, '', sz,  8, "", "", False, False, 'center', '']
           , "Prop": [ 5,  0, '', sz, 18, "", "", True, False,  'left', '']
           , "Vals": [ 6,  0, '', sz, 50, "", "", False, False, 'left', '']
           , "FCnt": [ 7,  0, '', sz,  8, "", "", False, False, 'center', '']
           , "PVI":  [ 8,  0, '', sz, 13, "", "", False, False, 'center', '']
           }
        self.tab_def['tab_table_links']  = [ 9, 0, '', sz, 18, "", "", False, False, 'left'  , '']
        self.tab_def['tab_cd_table_spacer'] = [10, 0, '', sz,  1, "", "", False, False, 'right'  , '']
        self.tab_def['tab_cd_fixed_grid'] = {
            # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
            # totals headers (across and down)
             'Analysis':   [ 9, 2, '', 14, 0, clrtx, clrfl, True,  False, 'center', 'Analysis']
          ,  'S1  ':   [10, 2, '', 12, 0, clrfl, clrfl, True,  False, 'right', 'Spc1  ']
          ,  'Prop':   [11, 2, '', 12, 0, clrtx, clrfl, True,  False, 'center', self.hdr_key_col[self.tab_id]]
          ,  'S2  ':   [12, 2, '', 12, 0, clrfl, clrfl, True,  False, 'right', 'Spc2  ']
          ,  'Val':    [13, 2, '', 12, 0, clrtx, clrfl, True,  False, 'center', self.hdr_val_col[self.tab_id]]
          ,  'S3  ':   [14, 2, '', 12, 0, clrfl, clrfl, True,  False, 'right', 'Spc3  ']
          ,  'Files Used':   [15, 2, '', 12, 0, clrtx, clrfl, True,  False, 'center', self.hdr_files_col[self.tab_id]]
          ,  'Unique Values': [ 9, 3, '', 12, 0, clrtx, clrfl, True,  False, 'left', 'Unique Values']
          ,  'Column Totals': [ 9, 4, '', 12, 0, clrtx, clrfl, True,  False, 'left', 'Column Totals']
          ,  'x-uniq-key':   [11, 3, '', sz, 0, "",              "", False, False, 'center', self.f_uniq_key]
          ,  'x-uniq-val':    [13, 3, '', sz, 0, "",              "", False, False, 'center', self.f_uniq_val]
          ,  'x-ctot-key':   [11, 4, '', sz, 0, "",              "", False, False, 'center', self.f_ctot_key]
          ,  'x-ctot-val':    [13, 4, '', sz, 0, "",              "", False, False, 'center', self.f_ctot_val]
          ,  'x-ctot-files':  [15, 4, '', sz, 0, "",              "", False, False, 'center', self.f_ctot_files]
          , 'x-null-values':  [ 6, 9, '', 12, 0, self.colors.clr_red20,  "", True, True,   'left',   self.f_null_values]
          , 'x-filters-on':   [ 9, 6, '', 12, 0, self.colors.clr_red20,  "", True, True,   'left',   self.f_filters_on]
          # This one is for the IVisible Column, not the totals
          , 'isVisible':      [38, 0, '', 10, 0, "",              "", False, False, 'right',  self.f_isVisible]
          }

        self.tab_def_post()

class DefAr51(NewTab):
    def __init__(self):
        self.tab_id = 'ar51'
        super().__init__(self.tab_id)
        clrtb, clrhi = self.tab_clrs[self.tab_id]

        clrfl = self.colors.Code_LOV[clrhi][0]
        clrtx = self.colors.Code_LOV[clrhi][1]
        sz = self.tab_txt_sz
        self.tab_def['tab_table_style'] = 'TableStyleMedium19'

        self.font_title_lst = ['Berlin Sans FB Demi', 24, clrtb]
        self.font_subs_lst = ['Berlin Sans', 14, clrtx]
        self.font_body_lst = ['Calibri', sz, clrtx]
        self.tab_def['tab_color'] = clrtb

        self.tab_def['tab_table_link_spcrs'] = True  # Always, TRUE for now

        self.tab_def['tab_has_isVisible_col'] = False
        self.tab_def['tab_tots_isVisible_col']  = 0
        self.tab_def['showGridLines'] = True

        self.tab_def['hdr_links_pfx'] = ""
        self.tab_def['tab_table_links_cols'] = 0

        self.tab_def['tab_name'] = 'Xyml'
        self.tab_def['tab_title_def'] = [3, 2, 'Berlin Sans FB Demi', 24, 80, clrhi, '', True, False, 'left', "Possibly Corrupt Markdown Files"]
        self.tab_def['tab_comments_def'] = [ 3, 3, '', 10, 0, '', '', False,  False, 'left', '']
        self.tab_def['tab_notes_def'] = [ 3, 22, '', 10, 0, '', '', False,  False, 'left', '']
        self.tab_def['tab_comments'] = [
              'Corrupt, meaning the Python package PyYAML 6.0.2 is unable to perform a "safe_load".'
            , 'This may or may not indicate an error, depending on how you setup and use your vault.'
            , 'They are logged here for further manual analysis, just in case.' ]  # up to 3 lines
             # 123456789 1 2345678 2 2345678 3 2345678 4 2345678 5 2345678 6 2345678 7 2345678 8
        self.tab_def['tab_notes'] = [
            'This system was intentionally designed to be "generous" when it comes to identifying errors,'
            'or deciding what constitutes an issue. It may even be true that files with no YAML at all are consider "Corrupt".'
            'Duplicate markdown filenames are allowed in Obsidian, but should be'
            'avoided. When linking to a duplicate file, you will be required to '
            'use full pathnames instead of just the filename.' ]

        self.tab_def['tab_cd_table_hdr'] = {  # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
             "RowId":         [4, 10, '', sz, 8, "", clrhi, True, False, 'center', self.hdr_RowId]
            , "Filename":   [5, 10, '', sz, 35, "", clrhi, True, False, 'left', self.hdr_files_col[self.tab_id]]
        }
        self.tab_def['tab_cd_table_dtl'] = {  # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
             "RowId":         [4, 0, '', sz, 0, "", "", False, False, 'center', '']
            , "Filename":   [5, 0, '', sz, 0, "", "", True, False, 'left', '']
        }
        self.tab_def['tab_table_links']  = [7, 10, '', sz, 25, "", "", False, False, 'left', '']
        self.tab_def['tab_cd_table_spacer'] = [8,  0, '', sz,  1, "", "", False, False, 'right', '']
        self.tab_def['tab_cd_fixed_grid'] = {  # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
            # totals headers (across and down)
            'Analysis': [ 9, 2, '', 14, 0, clrtx, clrfl, True, False, 'center', 'Analysis']
            , 'Spc1  ': [10, 2, '', 12, 0, clrfl, clrfl, True, False, 'right', 'Spc1  ']
            , 'Total':  [11, 2, '', 12, 0, clrtx, clrfl, True, False, 'center', self.hdr_files_col[self.tab_id]]
            , 'Files':  [9, 3, '', 12, 0, clrtx, clrfl, True, False, 'left', 'Files']
            , 'CTot':   [9, 4, '', 12, 0, clrtx, clrfl, True, False, 'left', 'Dups Found']
            , 'Notes':  [3, 21, '', 12, 0, clrtx, clrfl, True, False, 'left', 'Notes:']
            , 'x-ctot-files': [11, 3, '', sz, 0, "", "", False, False, 'center', self.f_ctot_files ]
            , 'x-ctot-dups': [11, 4, '', sz, 0, "", "", False, False, 'center', self.f_ctot_dup_founds]
        }

        self.tab_def_post()

class DefAr52(NewTab):
    def __init__(self):
        self.tab_id = 'xyml'
        super().__init__(self.tab_id)
        clrtb, clrhi = self.tab_clrs[self.tab_id]

        clrfl = self.colors.Code_LOV[clrhi][0]
        clrtx = self.colors.Code_LOV[clrhi][1]
        sz = self.tab_txt_sz
        self.tab_def['tab_table_style'] = 'TableStyleMedium19'

        self.font_title_lst = ['Berlin Sans FB Demi', 24, clrtb]
        self.font_subs_lst = ['Berlin Sans', 14, clrtx]
        self.font_body_lst = ['Calibri', sz, clrtx]
        self.tab_def['tab_color'] = clrtb

        self.tab_def['tab_table_link_spcrs'] = True  # Always, TRUE for now

        self.tab_def['tab_has_isVisible_col'] = False
        self.tab_def['tab_tots_isVisible_col']  = 0
        self.tab_def['showGridLines'] = True

        self.tab_def['hdr_links_pfx'] = ""
        self.tab_def['tab_table_links_cols'] = 0

        self.tab_def['tab_name'] = 'Xyml'
        self.tab_def['tab_title_def'] = [3, 2, 'Berlin Sans FB Demi', 24, 80, clrhi, '', True, False, 'left', "Possibly Corrupt Markdown Files"]
        self.tab_def['tab_comments_def'] = [ 3, 3, '', 10, 0, '', '', False,  False, 'left', '']
        self.tab_def['tab_notes_def'] = [ 3, 22, '', 10, 0, '', '', False,  False, 'left', '']
        self.tab_def['tab_comments'] = [
              'Corrupt, meaning the Python package PyYAML 6.0.2 is unable to perform a "safe_load".'
            , 'This may or may not indicate an error, depending on how you setup and use your vault.'
            , 'They are logged here for further manual analysis, just in case.' ]  # up to 3 lines
             # 123456789 1 2345678 2 2345678 3 2345678 4 2345678 5 2345678 6 2345678 7 2345678 8
        self.tab_def['tab_notes'] = [
            'This system was intentionally designed to be "generous" when it comes to identifying errors,'
            'or deciding what constitutes an issue. It may even be true that files with no YAML at all are consider "Corrupt".'
            'Duplicate markdown filenames are allowed in Obsidian, but should be'
            'avoided. When linking to a duplicate file, you will be required to '
            'use full pathnames instead of just the filename.' ]

        self.tab_def['tab_cd_table_hdr'] = {  # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
             "RowId":         [4, 10, '', sz, 8, "", clrhi, True, False, 'center', self.hdr_RowId]
            , "Filename":   [5, 10, '', sz, 35, "", clrhi, True, False, 'left', self.hdr_files_col[self.tab_id]]
        }
        self.tab_def['tab_cd_table_dtl'] = {  # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
             "RowId":         [4, 0, '', sz, 0, "", "", False, False, 'center', '']
            , "Filename":   [5, 0, '', sz, 0, "", "", True, False, 'left', '']
        }
        self.tab_def['tab_table_links']  = [7, 10, '', sz, 25, "", "", False, False, 'left', '']
        self.tab_def['tab_cd_table_spacer'] = [8,  0, '', sz,  1, "", "", False, False, 'right', '']
        self.tab_def['tab_cd_fixed_grid'] = {  # [col,row,font,sz, w,t_clr,fill_clr,Bold,Ital,  Align,  val ] = 11
            # totals headers (across and down)
            'Analysis': [ 9, 2, '', 14, 0, clrtx, clrfl, True, False, 'center', 'Analysis']
            , 'Spc1  ': [10, 2, '', 12, 0, clrfl, clrfl, True, False, 'right', 'Spc1  ']
            , 'Total':  [11, 2, '', 12, 0, clrtx, clrfl, True, False, 'center', self.hdr_files_col[self.tab_id]]
            , 'Files':  [9, 3, '', 12, 0, clrtx, clrfl, True, False, 'left', 'Files']
            , 'CTot':   [9, 4, '', 12, 0, clrtx, clrfl, True, False, 'left', 'Dups Found']
            , 'Notes':  [3, 21, '', 12, 0, clrtx, clrfl, True, False, 'left', 'Notes:']
            , 'x-ctot-files': [11, 3, '', sz, 0, "", "", False, False, 'center', self.f_ctot_files ]
            , 'x-ctot-dups': [11, 4, '', sz, 0, "", "", False, False, 'center', self.f_ctot_dup_founds]
        }

        self.tab_def_post()


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





