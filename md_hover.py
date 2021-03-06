import asyncio
import threading
import time
import webbrowser
from ctypes import windll
from json import loads, dumps
from tkinter import *
from tkinter.ttk import *

import dhash
import pymem
import pyperclip
import win32gui
import win32ui
from PIL import Image

n_flags = 3
try:
    _win_v = sys.getwindowsversion()
    if _win_v.major == 6 and _win_v.minor == 1:
        n_flags = 1
except Exception:
    pass
'''
WIN_10 = (10, 0, 0)
WIN_8 = (6, 2, 0)
WIN_7 = (6, 1, 0)
WIN_SERVER_2008 = (6, 0, 1)
WIN_VISTA_SP1 = (6, 0, 1)
WIN_VISTA = (6, 0, 0)
WIN_SERVER_2003_SP2 = (5, 2, 2)
WIN_SERVER_2003_SP1 = (5, 2, 1)
WIN_SERVER_2003 = (5, 2, 0)
WIN_XP_SP3 = (5, 1, 3)
WIN_XP_SP2 = (5, 1, 2)
WIN_XP_SP1 = (5, 1, 1)
WIN_XP = (5, 1, 0)
'''
windll.user32.SetProcessDPIAware()
BOXES = (
    ((55, 159, 126, 231), (41, 169, 124, 253), (125, 192, 349, 414)),  # 1280x720
    ((58, 170, 135, 246), (44, 181, 132, 269), (134, 205, 373, 442)),  # 1366x768
    ((61, 179, 142, 259), (46, 190, 140, 284), (141, 216, 392, 466)),  # 1440x810
    ((68, 199, 158, 288), (51, 212, 156, 316), (156, 240, 437, 518)),  # 1600x900
    ((82, 239, 189, 346), (61, 254, 186, 379), (187, 288, 524, 622)),  # 1920x1080
    ((87, 255, 202, 369), (66, 271, 199, 404), (200, 307, 559, 663)),  # 2048x1152
    ((109, 318, 252, 461), (82, 339, 249, 505), (250, 384, 698, 829)),  # 2560x1440
    ((136, 398, 316, 576), (102, 424, 312, 632), (312, 480, 874, 1036)),  # 3200x1800
    ((164, 478, 378, 692), (122, 508, 372, 758), (374, 576, 1048, 1244)),  # 3840x2160
)
F2A = {
    37: {6: 0x348, 211: 0x360, 159: 0x360, 45: 0x378},  # 34/35/36 x 24 = 816/840/864
    17: {6: 0xC0, 211: 0xD8, 159: 0xD8, 45: 0xF0},  # 11/12/13 = 264/288/312
    89: {6: 0xF0, 211: 0x108, 159: 0x108, 45: 0x120},  # 9/10/11 = 216/240/264
    # 53/52/54 = 1272/1248/1296
}


def text_event(event):
    if event.state == 12 and event.keysym == 'c':
        return
    else:
        return "break"


def screenshot():
    global n_flags
    hwnd = win32gui.FindWindow(None, 'masterduel')
    if hwnd:
        box = win32gui.GetWindowRect(hwnd)
        box_w = box[2] - box[0]
        box_h = box[3] - box[1]
        hwnd_dc = win32gui.GetWindowDC(hwnd)
        mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
        save_dc = mfc_dc.CreateCompatibleDC()
        save_bitmap = win32ui.CreateBitmap()
        save_bitmap.CreateCompatibleBitmap(mfc_dc, box_w, box_h)  # 10%
        save_dc.SelectObject(save_bitmap)
        result = windll.user32.PrintWindow(hwnd, save_dc.GetSafeHdc(), n_flags)  # 58%
        bmpinfo = save_bitmap.GetInfo()
        bmpstr = save_bitmap.GetBitmapBits(True)  # 19%
        im = Image.frombuffer('RGB', (bmpinfo['bmWidth'], bmpinfo['bmHeight']), bmpstr, 'raw', 'BGRX', 0, 1)  # 12%
        win32gui.DeleteObject(save_bitmap.GetHandle())
        save_dc.DeleteDC()
        mfc_dc.DeleteDC()
        win32gui.ReleaseDC(hwnd, hwnd_dc)
        if result != 0:
            return im
    return None


def get_pm():
    try:
        pm = pymem.Pymem('masterduel.exe')
        addr = pymem.process.module_from_name(pm.process_handle, 'GameAssembly.dll').lpBaseOfDll
        return pm, addr
    except Exception:
        return None, None


def copy_widget_text(event):
    pyperclip.copy(event.widget['text'])


class MainWindow(Tk):
    def __init__(self):
        super().__init__()
        self.title('MD_HOVER')
        self.loop = None
        self.cfg = {
            'w': 440,
            'h': 320,
            'x': 5,
            'y': 466,
            'f': 12,
            'm': 0,
            'r': 0,
            'j': 1,
            'e': 1,
            'a': 1,
            'p': 0,
            't': 0,
            's': 1,
        }
        try:
            with open('md_hover.cfg', 'r') as f:
                ini = loads(f.read())
                self.cfg.update(ini)
        except Exception:
            pass
        self.geometry(f"{self.cfg['w']}x{self.cfg['h']}+{self.cfg['x']}+{self.cfg['y']}")
        self.wm_attributes('-topmost', True)
        self.protocol('WM_DELETE_WINDOW', self.save_exit)
        # self.wm_attributes("-transparentcolor", "white")
        self.override = False
        self.toolwindow = False
        self.ygoid = 89631139
        self.scan_cid = 0
        self.last_hash = [0] * 3
        self.pm = None
        self.base_addr = None
        self.deck_addr = None
        self.revi_addr = None
        self.duel_addr = None
        self.play_addr = None
        self.shop_flag_addr = None
        self.shop_count = None
        self.shop_cids = None
        self.offsets = [816, 840, 864]
        self.deck_cid = -1
        self.revi_cid = -1
        self.duel_cid = -1
        self.play_cid = -1
        self.shop_cid = -1
        self.cards_info = {}
        with open('cards.json', 'r', encoding='utf-8') as f:
            data = loads(f.read())
            for cid in data.keys():
                cid_int = int(cid)
                card_info = data[cid]
                cn_name = card_info['cn_name'] if 'cn_name' in card_info else ''
                jp_name = card_info['jp_name'] if 'jp_name' in card_info else ''
                en_name = card_info['en_name'] if 'en_name' in card_info else ''
                text = card_info['text']
                pdesc = text['pdesc']
                tdesc = text['types']  # replace('\n', ' ')
                desc = text['desc'].replace('\r\n', '\n')
                fdesc = '\n????????????????????????\n'.join((pdesc, desc)) if pdesc else desc
                self.cards_info[cid_int] = (card_info['id'], cn_name, jp_name, en_name, fdesc, tdesc)
        with open('tokens.json', 'r', encoding='utf-8') as f:
            data = loads(f.read())
            for cid, info in iter(data.items()):
                self.cards_info[int(cid)] = (89631139, info[0], info[1], info[2], info[3], '')
        with open('hash.json', 'r', encoding='utf-8') as f:
            self.art_hash = loads(f.read())

        self.fontsize = self.cfg['f']
        self.tf = Label(self)
        self.tf.pack(side='top', fill='x', pady=5)

        self.cname_label = Label(self.tf, font=('????????????', self.fontsize + 1, 'bold'), text='????????????', padding=-3)
        self.cname_label.pack(side='top', fill='x', padx=(5, 0))
        self.cname_label.bind('<Button-1>', copy_widget_text)
        self.jname_label = Label(self.tf, font=('????????????', self.fontsize + 1, 'bold'), text='???????????????', padding=-3)
        self.jname_label.pack(side='top', fill='x', padx=(5, 0))
        self.jname_label.bind('<Button-1>', copy_widget_text)
        self.ename_label = Label(self.tf, font=('????????????', self.fontsize + 1, 'bold'),
                                 text='Blue-Eyes White Dragon', padding=-3)
        self.ename_label.pack(side='top', fill='x', padx=(5, 0))
        self.ename_label.bind('<Button-1>', copy_widget_text)
        self.attr_label = Label(self.tf, font=('????????????', self.fontsize),
                                text='[??????|??????] ???/???\n[???8] 3000/2500', padding=-3)
        self.attr_label.pack(side='top', fill='x', padx=(5, 0))

        self.desc_label = Text(self, font=('????????????', self.fontsize))
        self.desc_label.insert(END, '?????????????????????????????????????????????????????????????????????????????????????????????')
        self.desc_scroll = Scrollbar(self, orient='vertical', command=self.desc_label.yview)
        self.desc_label.configure(yscrollcommand=self.desc_scroll.set)
        self.desc_scroll.pack(side='right', fill=Y)
        self.desc_label.pack(side='left', expand=1, fill=BOTH)
        self.desc_label.bind("<Key>", lambda xx: text_event(xx))
        self.widgets = [self.tf, self.cname_label, self.jname_label, self.ename_label, self.attr_label,
                        self.desc_label, self]
        self.widgets_bg = [widget['background'] for widget in self.widgets]
        self.widgets_fg = [widget['foreground'] for widget in self.widgets[:-1]]

        self.mode_t0 = ['??????', '??????']
        self.mode_t = f'???????????????????????????\n???????????????????????????\n???????????????????????????\n????????????????????????????????????\n\n' \
                      f'????????????{self.mode_t0[self.cfg["m"]]}?????????'
        self.Button1 = Button(self.desc_label, text=self.mode_t, width=30, command=self.main_start)
        self.Button1.pack()
        self.Button1.place(x=50, y=50)

        self.Menu = Menu(self, tearoff=False)

        self.mode_menu = Menu(self.Menu, tearoff=0)
        self.mode_var = IntVar(value=self.cfg['m'])
        for i, lb in enumerate(['??????', '??????']):
            self.mode_menu.add_radiobutton(label=lb, command=self.set_mode, variable=self.mode_var, value=i)
        self.Menu.add_cascade(label='??????', menu=self.mode_menu)

        self.play_menu = Menu(self.Menu, tearoff=0)
        self.play_var = IntVar(value=self.cfg['p'])
        for i, lb in enumerate(['??????', '??????']):
            self.play_menu.add_radiobutton(label=lb, variable=self.play_var, value=i)
        self.Menu.add_cascade(label='??????(??????)', menu=self.play_menu)

        self.rate_menu = Menu(self.Menu, tearoff=0)
        self.rate_var = IntVar(value=self.cfg['r'])
        for i, lb in enumerate(['4?????????', '16?????????']):
            self.rate_menu.add_radiobutton(label=lb, variable=self.rate_var, value=i)
        self.Menu.add_cascade(label='????????????', menu=self.rate_menu)

        self.Menu.add_separator()

        self.Menu.add_command(label='????????????/??????', command=self.over)
        self.Menu.add_command(label='????????????/??????', command=self.icon)

        self.Menu.add_separator()

        self.fontsize_menu = Menu(self.Menu, tearoff=0)
        self.fontsize_menu.add_command(label='-', command=self.font_size_down)
        self.fontsize_menu.add_command(label='+', command=self.font_size_up)
        self.Menu.add_cascade(label='????????????', menu=self.fontsize_menu)

        self.Menu.add_separator()

        self.jname_menu = Menu(self.Menu, tearoff=0)
        self.jname_var = IntVar(value=self.cfg['j'])
        for i, lb in enumerate(['??????', '??????']):
            self.jname_menu.add_radiobutton(label=lb, command=self.set_jname, variable=self.jname_var, value=i)
        self.Menu.add_cascade(label='????????????', menu=self.jname_menu)

        self.ename_menu = Menu(self.Menu, tearoff=0)
        self.ename_var = IntVar(value=self.cfg['e'])
        for i, lb in enumerate(['??????', '??????']):
            self.ename_menu.add_radiobutton(label=lb, command=self.set_ename, variable=self.ename_var, value=i)
        self.Menu.add_cascade(label='????????????', menu=self.ename_menu)

        self.attr_menu = Menu(self.Menu, tearoff=0)
        self.attr_var = IntVar(value=self.cfg['a'])
        for i, lb in enumerate(['??????', '??????']):
            self.attr_menu.add_radiobutton(label=lb, command=self.set_attr, variable=self.attr_var, value=i)
        self.Menu.add_cascade(label='????????????', menu=self.attr_menu)

        self.Menu.add_separator()

        self.Menu.add_command(label='OUROCG', command=self.ourocg)
        self.Menu.add_command(label='????????????', command=self.ygocdb)

        self.Menu.add_separator()

        self.theme_menu = Menu(self.Menu, tearoff=0)
        self.theme_var = IntVar(value=self.cfg['t'])
        for i, lb in enumerate(['??????', '??????']):
            self.theme_menu.add_radiobutton(label=lb, command=self.set_theme, variable=self.theme_var, value=i)
        self.Menu.add_cascade(label='??????', menu=self.theme_menu)

        self.scroll_menu = Menu(self.Menu, tearoff=0)
        self.scroll_var = IntVar(value=self.cfg['s'])
        for i, lb in enumerate(['??????', '??????']):
            self.scroll_menu.add_radiobutton(label=lb, command=self.set_scroll, variable=self.scroll_var, value=i)
        self.Menu.add_cascade(label='?????????', menu=self.scroll_menu)

        self.Menu.add_separator()

        self.Menu.add_command(label='12??? 220228')

        self.bind("<Button-3>", self.popupmenu)
        self.set_jname()
        self.set_ename()
        self.set_attr()
        self.set_theme()
        self.set_scroll()
        self.mainloop()

    def popupmenu(self, event):
        try:
            self.Menu.post(event.x_root, event.y_root)
        finally:
            self.Menu.grab_release()

    def save_exit(self):
        self.cfg = {
            'w': self.winfo_width(),
            'h': self.winfo_height(),
            'x': self.winfo_x(),
            'y': self.winfo_y(),
            'f': self.fontsize,
            'm': self.mode_var.get(),
            'p': self.play_var.get(),
            'r': self.rate_var.get(),
            'j': self.jname_var.get(),
            'e': self.ename_var.get(),
            'a': self.attr_var.get(),
            't': self.theme_var.get(),
            's': self.scroll_var.get(),
        }
        try:
            with open('md_hover.cfg', 'w') as f:
                print(dumps(self.cfg, separators=(',', ':')), file=f)
        finally:
            self.destroy()

    def set_jname(self):
        if self.jname_var.get():
            self.jname_label.pack(side='top', fill='x', padx=(5, 0))
        else:
            self.jname_label.pack_forget()

    def set_ename(self):
        if self.ename_var.get():
            self.ename_label.pack(side='top', fill='x', padx=(5, 0))
        else:
            self.ename_label.pack_forget()

    def set_attr(self):
        if self.attr_var.get():
            self.attr_label.pack(side='top', fill='x', padx=(5, 0))
        else:
            self.attr_label.pack_forget()

    def set_mode(self):
        try:
            self.Button1['text'] = f'???????????????????????????\n???????????????????????????\n???????????????????????????\n????????????????????????????????????\n\n' \
                                   f'????????????{self.mode_t0[self.mode_var.get()]}?????????'
        except Exception:
            pass

    def set_scroll(self):
        if self.scroll_var.get():
            self.desc_label.pack_forget()
            self.desc_scroll.pack(side='right', fill=Y)
            self.desc_label.pack(side='left', expand=1, fill=BOTH)
        else:
            self.desc_scroll.pack_forget()

    def set_theme(self):
        if self.theme_var.get():
            for i in range(6):
                self.widgets[i]['background'] = '#1F1F1F'
                self.widgets[i]['foreground'] = '#DFDFDF'
            self.widgets[6]['background'] = '#1C1C1C'
        else:
            for i in range(6):
                self.widgets[i]['background'] = self.widgets_bg[i]
                self.widgets[i]['foreground'] = self.widgets_fg[i]
            self.widgets[6]['background'] = self.widgets_bg[6]

    def apply_font(self):
        f0 = ('????????????', self.fontsize)
        f1 = ('????????????', self.fontsize + 1, 'bold')
        self.cname_label['font'] = f1
        self.jname_label['font'] = f1
        self.ename_label['font'] = f1
        self.attr_label['font'] = f0
        self.desc_label['font'] = f0

    def font_size_up(self):
        self.fontsize += 1
        self.apply_font()

    def font_size_down(self):
        if self.fontsize > 6:
            self.fontsize -= 1
            self.apply_font()

    def get_loop(self, loop):
        self.loop = loop
        asyncio.set_event_loop(self.loop)
        self.loop.run_forever()

    def main_start(self):
        self.Button1.destroy()
        coroutine1 = self.handler()
        new_loop = asyncio.new_event_loop()
        t = threading.Thread(target=self.get_loop, args=(new_loop,))
        t.daemon = True
        t.start()
        asyncio.run_coroutine_threadsafe(coroutine1, new_loop)

    def ygocdb(self):
        webbrowser.open_new('https://ygocdb.com/card/' + str(self.ygoid))

    def ourocg(self):
        webbrowser.open_new('https://www.ourocg.cn/search/' + str(self.ygoid))

    def over(self):
        self.override = not self.override
        self.overrideredirect(self.override)

    def icon(self):
        self.toolwindow = not self.toolwindow
        self.wm_attributes('-toolwindow', self.toolwindow)

    def get_scan(self):
        full_img = screenshot()
        if not full_img:
            return -1
        imgx, imgy = full_img.size
        if imgx < 1280:
            return -1
        elif imgx < 1366:
            resolution = 0
        elif imgx < 1440:
            resolution = 1
        elif imgx < 1600:
            resolution = 2
        elif imgx < 1920:
            resolution = 3
        elif imgx < 2048:
            resolution = 4
        elif imgx < 2560:
            resolution = 5
        elif imgx < 3200:
            resolution = 6
        elif imgx < 3840:
            resolution = 7
        else:
            resolution = 8

        results = []
        for i in range(3):
            _sample = full_img.crop(BOXES[resolution][i])
            _hash = dhash.dhash_int(_sample, 16)
            if dhash.get_num_bits_different(_hash, self.last_hash[i]) > 10:
                hash_compare = [(dhash.get_num_bits_different(_hash, item[0]), item[1]) for item in self.art_hash]
                results.append(min(hash_compare, key=lambda xx: xx[0]))
                self.last_hash[i] = _hash
        if results:
            result = min(results, key=lambda xx: xx[0])
            if result[0] < 60:
                return result[1]
        return 0

    def get_memory(self):
        pm = self.pm
        base_addr = self.base_addr
        deck_cid, revi_cid, duel_cid, play_cid, shop_cid = -1, -1, -1, -1, -1

        try:
            value = pm.read_longlong(base_addr + 0x01CCE3C0)
            for offset in [0xB8, 0x0, 0xF8]:
                value = pm.read_longlong(value + offset)
            deck_cid = pm.read_int(pm.read_longlong(value + 0x1D8) + 0x20)
            revi_cid = pm.read_int(pm.read_longlong(value + 0x140) + 0x20)
        except Exception:
            pass

        try:
            value = pm.read_longlong(base_addr + 0x01CB3CE0)
            for offset in [0xB8, 0x0]:
                value = pm.read_longlong(value + offset)
            duel_cid = pm.read_int(value + 0x44)
        except Exception:
            pass

        try:
            value = pm.read_longlong(base_addr + 0x01CF5A18)
            for offset in [0xB8, 0x0, 0x10]:
                value = pm.read_longlong(value + offset)
            play_cid = pm.read_int(value + 0x1DC)
        except Exception:
            pass

        if self.shop_flag_addr:
            try:
                shop_flag = pm.read_uchar(self.shop_flag_addr)
                if shop_flag == 6 or shop_flag == 108:
                    pos = 0
                elif shop_flag == 221 or shop_flag == 159 or shop_flag == 211:
                    pos = 1
                elif shop_flag == 45:
                    pos = 2
                else:
                    pos = -1
                if pos >= 0:
                    try:
                        _addr = pm.read_longlong(base_addr + 0x01D9ED08)
                        for _offset in [0x40, 0x20, 0x110, 0x18]:
                            _addr = pm.read_longlong(_addr + _offset)
                        _addr += 0x18
                        shop_count = pm.read_longlong(_addr)
                        shop_cid = pm.read_int(_addr + self.offsets[pos])
                        shop_cids = [pm.read_longlong(_addr + 24 * i) for i in range(1, shop_count + 1)]
                        if self.shop_count == shop_count:
                            diff = [True if self.shop_cids[i] - shop_cids[i] else False for i in range(shop_count)]
                            if diff.count(True) == 2:
                                start = diff.index(True)
                                if diff[start + 2]:
                                    off0 = start * 24 + 24
                                    off1 = off0 + 24
                                    off2 = off1 + 24
                                    self.offsets = [off0, off1, off2]
                        self.shop_cids = shop_cids
                        self.shop_count = shop_count
                    except Exception:
                        pass
            except Exception:
                self.shop_flag_addr = None
        else:
            try:
                _addr = pm.read_longlong(base_addr + 0x01BD4508)
                for _offset in [0xB8, 0x0, 0x20, 0x18, 0x2E0, 0x420]:
                    _addr = pm.read_longlong(_addr + _offset)
                self.shop_flag_addr = _addr + 0x21
            except Exception:
                pass
        return deck_cid, revi_cid, duel_cid, play_cid, shop_cid

    async def handler(self):
        while True:
            _t0 = time.time()
            _scan_mode = not self.mode_var.get()
            rate_sleep = 0.0625 if self.rate_var.get() else 0.25
            final_cid = 0

            if _scan_mode:
                new_cid = self.get_scan()
                if new_cid != self.scan_cid and new_cid != 0:
                    final_cid = new_cid
                    self.scan_cid = new_cid
            else:
                if not self.pm:
                    self.pm, self.base_addr = get_pm()
                if self.pm:
                    try:
                        self.pm.read_longlong(self.base_addr)
                        deck_cid, revi_cid, duelcid, play_cid, shop_cid = self.get_memory()
                        if 3899 < duelcid < 18000 and duelcid != self.duel_cid:
                            final_cid = duelcid
                            self.duel_cid = duelcid
                        elif 3899 < revi_cid < 18000 and revi_cid != self.revi_cid:
                            final_cid = revi_cid
                            self.revi_cid = revi_cid
                        elif 3899 < deck_cid < 18000 and deck_cid != self.deck_cid:
                            final_cid = deck_cid
                            self.deck_cid = deck_cid
                        elif self.play_var.get() and 3899 < play_cid < 18000 and play_cid != self.play_cid:
                            final_cid = play_cid
                            self.play_cid = play_cid
                        elif 3899 < shop_cid < 18000 and shop_cid != self.shop_cid:
                            final_cid = shop_cid
                            self.shop_cid = shop_cid
                    except Exception:
                        self.pm, self.base_addr = None, 0
                        self.deck_addr = None
                        self.revi_addr = None
                        self.duel_addr = None
                        self.play_addr = None
                        self.shop_flag_addr = None
                        self.shop_count = None
                        self.shop_cids = None
                        self.offsets = [816, 840, 864]
            if final_cid in self.cards_info:
                hover_info = self.cards_info[final_cid]
                self.ygoid = hover_info[0]
                self.cname_label['text'] = hover_info[1]
                self.jname_label['text'] = hover_info[2]
                self.ename_label['text'] = hover_info[3]
                self.desc_label.delete('1.0', END)
                self.desc_label.insert(END, hover_info[4])
                self.attr_label['text'] = hover_info[5]
            dt = _t0 - time.time() + rate_sleep
            if dt > 0:
                await asyncio.sleep(dt)


if __name__ == '__main__':
    MainWindow()
