def complement(clr_in):
    a_red = 255 - int(clr_in[0:2], base=16)
    a_grn = 255 - int(clr_in[2:4], base=16)
    a_blu = 255 - int(clr_in[4:6], base=16)
    # f"{i:02x}"
    complement = f"{a_red:02x}{a_grn:02x}{a_blu:02x}"

    return complement

class Colors:
    def __init__(self):
        self.clr_pur44 = '00CC99FF'
        self.clr_pur45 = '00660066'
        self.clr_ora64 = '00FFCC99'
        self.clr_blud4 = '004F81BD'
        self.clr_red20 = '00C0504D'
        self.clr_grn30 = '009BBB59'
        self.clr_aqu56 = '0033CCCC'
        self.clr_gry35 = '00C0C0C0'
        self.clr_wht15 = '00F2F2F2'
        self.clr_wht00 = '00FFFFFF'
        self.clr_blk00 = '00000000'
        self.Name_LOV = {
            'pur44': [self.clr_pur44, self.clr_wht00],
            'pur45': [self.clr_pur45, self.clr_wht00],
            'ora64': [self.clr_ora64, self.clr_blk00],
            'blud4': [self.clr_blud4, self.clr_wht00],
            'red20': [self.clr_red20, self.clr_wht00],
            'grn30': [self.clr_grn30, self.clr_blk00],
            'aqu56': [self.clr_aqu56, self.clr_blk00],
            'gry35': [self.clr_gry35, self.clr_blk00],
            'wht15': [self.clr_wht15, self.clr_blk00],
            'wht00': [self.clr_wht00, self.clr_blk00],
            'blk00': [self.clr_blk00, self.clr_wht00],
        }
        self.Code_LOV = {
            '00CC99FF': [self.clr_pur44, self.clr_wht00, 'clr_pur44'],
            '00660066': [self.clr_pur45, self.clr_wht00, 'clr_pur45'],
            '00FFCC99': [self.clr_ora64, self.clr_blk00, 'clr_ora64'],
            '004F81BD': [self.clr_blud4, self.clr_wht00, 'clr_blud4'],
            '00C0504D': [self.clr_red20, self.clr_wht00, 'clr_red20'],
            '009BBB59': [self.clr_grn30, self.clr_blk00, 'clr_grn30'],
            '0033CCCC': [self.clr_aqu56, self.clr_blk00, 'clr_aqu56'],
            '00C0C0C0': [self.clr_gry35, self.clr_blk00, 'clr_gry35'],
            '00F2F2F2': [self.clr_wht15, self.clr_blk00, 'clr_wht15'],
            '00FFFFFF': [self.clr_wht00, self.clr_blk00, 'clr_wht00'],
            '00000000': [self.clr_blk00, self.clr_wht00, 'clr_blk00'],
        }
