import asyncio
import threading
import time
import webbrowser
from ctypes import windll
# from datetime import datetime
from json import loads, dumps
from tkinter import *
from tkinter.ttk import *

import dhash
import pymem
import pyperclip
import win32gui
import win32ui
from PIL import Image

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


def text_event(event):
    if event.state == 12 and event.keysym == 'c':
        return
    else:
        return "break"


def screenshot():
    hwnd = win32gui.FindWindow(None, 'masterduel')
    if hwnd:
        box = win32gui.GetWindowRect(hwnd)
        box_w = box[2] - box[0]
        box_h = box[3] - box[1]
        hwndDC = win32gui.GetWindowDC(hwnd)
        mfcDC = win32ui.CreateDCFromHandle(hwndDC)
        saveDC = mfcDC.CreateCompatibleDC()
        saveBitMap = win32ui.CreateBitmap()
        saveBitMap.CreateCompatibleBitmap(mfcDC, box_w, box_h)  # 10%
        saveDC.SelectObject(saveBitMap)
        result = windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 3)  # 58%
        bmpinfo = saveBitMap.GetInfo()
        bmpstr = saveBitMap.GetBitmapBits(True)  # 19%
        im = Image.frombuffer('RGB', (bmpinfo['bmWidth'], bmpinfo['bmHeight']), bmpstr, 'raw', 'BGRX', 0, 1)  # 12%
        win32gui.DeleteObject(saveBitMap.GetHandle())
        saveDC.DeleteDC()
        mfcDC.DeleteDC()
        win32gui.ReleaseDC(hwnd, hwndDC)
        if result == 1:
            return True, im
    return False, None


def get_pm():
    try:
        pm = pymem.Pymem('masterduel.exe')
        addr = pymem.process.module_from_name(pm.process_handle, 'GameAssembly.dll').lpBaseOfDll
        return pm, addr
    except Exception:
        return None, 0


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
        self.base_addr = 0
        self.deck_cid = -1
        self.repl_cid = -1
        self.duel_cid = -1
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
                fdesc = '\n　　○○○○　　\n'.join((pdesc, desc)) if pdesc else desc
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

        self.cname_label = Label(self.tf, font=('微软雅黑', self.fontsize + 1, 'bold'), text='青眼白龙', padding=-3)
        self.cname_label.pack(side='top', fill='x', padx=(5, 0))
        self.cname_label.bind('<Button-1>', copy_widget_text)
        self.jname_label = Label(self.tf, font=('微软雅黑', self.fontsize + 1, 'bold'), text='青眼の白龍', padding=-3)
        self.jname_label.pack(side='top', fill='x', padx=(5, 0))
        self.jname_label.bind('<Button-1>', copy_widget_text)
        self.ename_label = Label(self.tf, font=('微软雅黑', self.fontsize + 1, 'bold'),
                                 text='Blue-Eyes White Dragon', padding=-3)
        self.ename_label.pack(side='top', fill='x', padx=(5, 0))
        self.ename_label.bind('<Button-1>', copy_widget_text)
        self.attr_label = Label(self.tf, font=('微软雅黑', self.fontsize),
                                text='[怪兽|通常] 龙/光\n[★8] 3000/2500', padding=-3)
        self.attr_label.pack(side='top', fill='x', padx=(5, 0))

        self.desc_label = Text(self, font=('微软雅黑', self.fontsize))
        # self.desc_label.configure(state='disabled')
        self.desc_label.insert(END, '以高攻击力著称的传说之龙。任何对手都能粉碎，其破坏力不可估量。')
        self.desc_scroll = Scrollbar(self, orient='vertical', command=self.desc_label.yview)
        self.desc_label.configure(yscrollcommand=self.desc_scroll.set)
        self.desc_scroll.pack(side='right', fill=Y)
        self.desc_label.pack(side='left', expand=1, fill=BOTH)
        self.desc_label.bind("<Key>", lambda xx: text_event(xx))

        self.mode_t0 = ['识图', '内存']
        self.mode_t = f'以{self.mode_t0[self.cfg["m"]]}模式\n运行插件\n\n右击菜单\n调整模式'
        self.Button1 = Button(self.desc_label, text=self.mode_t, width=30, command=self.main_start)
        self.Button1.pack()
        self.Button1.place(x=50, y=50)

        self.Menu = Menu(self, tearoff=False)

        self.mode_menu = Menu(self.Menu, tearoff=0)
        self.mode_var = IntVar(value=self.cfg['m'])
        for i, lb in enumerate(['识图', '内存']):
            self.mode_menu.add_radiobutton(label=lb, command=self.set_mode, variable=self.mode_var, value=i)
        self.Menu.add_cascade(label='模式', menu=self.mode_menu)
        self.rate_menu = Menu(self.Menu, tearoff=0)
        self.rate_var = IntVar(value=self.cfg['r'])
        for i, lb in enumerate(['4帧节能', '16帧够用']):
            self.rate_menu.add_radiobutton(label=lb, variable=self.rate_var, value=i)
        self.Menu.add_cascade(label='刷新频率', menu=self.rate_menu)

        self.Menu.add_separator()

        self.Menu.add_command(label='锁定窗口/解锁', command=self.over)
        self.Menu.add_command(label='简化窗口/恢复', command=self.icon)

        self.Menu.add_separator()

        self.fontsize_menu = Menu(self.Menu, tearoff=0)
        self.fontsize_menu.add_command(label='-', command=self.font_size_down)
        self.fontsize_menu.add_command(label='+', command=self.font_size_up)
        self.Menu.add_cascade(label='调整字号', menu=self.fontsize_menu)

        self.Menu.add_separator()

        self.jname_menu = Menu(self.Menu, tearoff=0)
        self.jname_var = IntVar(value=self.cfg['j'])
        for i, lb in enumerate(['隐藏', '显示']):
            self.jname_menu.add_radiobutton(label=lb, command=self.set_jname, variable=self.jname_var, value=i)
        self.Menu.add_cascade(label='日语牌名', menu=self.jname_menu)

        self.ename_menu = Menu(self.Menu, tearoff=0)
        self.ename_var = IntVar(value=self.cfg['e'])
        for i, lb in enumerate(['隐藏', '显示']):
            self.ename_menu.add_radiobutton(label=lb, command=self.set_ename, variable=self.ename_var, value=i)
        self.Menu.add_cascade(label='英语牌名', menu=self.ename_menu)

        self.attr_menu = Menu(self.Menu, tearoff=0)
        self.attr_var = IntVar(value=self.cfg['a'])
        for i, lb in enumerate(['隐藏', '显示']):
            self.attr_menu.add_radiobutton(label=lb, command=self.set_attr, variable=self.attr_var, value=i)
        self.Menu.add_cascade(label='类别属性', menu=self.attr_menu)

        self.Menu.add_separator()

        self.Menu.add_command(label='OUROCG', command=self.ourocg)
        self.Menu.add_command(label='百鸽查卡', command=self.ygocdb)

        self.bind("<Button-3>", self.popupmenu)
        self.set_jname()
        self.set_ename()
        self.set_attr()
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
            'r': self.rate_var.get(),
            'j': self.jname_var.get(),
            'e': self.ename_var.get(),
            'a': self.attr_var.get(),
        }
        try:
            with open('md_hover.cfg', 'w') as f:
                print(dumps(self.cfg), file=f)
        finally:
            self.destroy()

    def set_dark(self):
        if self.dark_var.get():
            self.dark_label.pack(side='top', fill='x', padx=(5, 0))
        else:
            self.dark_label.pack_forget()

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

    def set_mode(self):
        mode = self.mode_var.get()
        mode_t = self.mode_t0[mode]
        try:
            self.Button1['text'] = f'以{mode_t}模式\n运行插件\n\n右击菜单\n调整模式'
        except Exception:
            pass

    def font_size_up(self):
        self.fontsize += 1
        f0 = ('微软雅黑', self.fontsize, 'bold')
        f1 = ('微软雅黑', self.fontsize + 1, 'bold')
        self.cname_label['font'] = f1
        self.jname_label['font'] = f1
        self.ename_label['font'] = f1
        self.attr_label['font'] = f0
        self.desc_label['font'] = f0

    def font_size_down(self):
        if self.fontsize > 6:
            self.fontsize -= 1
            f0 = ('微软雅黑', self.fontsize, 'bold')
            f1 = ('微软雅黑', self.fontsize + 1, 'bold')
            self.cname_label['font'] = f1
            self.jname_label['font'] = f1
            self.ename_label['font'] = f1
            self.attr_label['font'] = f0
            self.desc_label['font'] = f0

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
        # now_str = str(datetime.now())[2:19].replace(":", "_")
        flag, full_img = screenshot()
        if not flag:
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
        base_addr = self.base_addr
        pm = self.pm
        deck_cid, repl_cid, duel_cid = -1, -1, -1
        try:
            value = pm.read_longlong(base_addr + 0x01CCD278)
            for offset in [0xB8, 0x0, 0xF8]:
                value = pm.read_longlong(value + offset)
            deck_cid = pm.read_int(pm.read_longlong(value + 0x1D8) + 0x20)
            repl_cid = pm.read_int(pm.read_longlong(value + 0x140) + 0x20)
        except Exception:
            pass
        try:
            value = pm.read_longlong(base_addr + 0x01cb2b90)
            for offset in [0xB8, 0x0]:
                value = pm.read_longlong(value + offset)
            duel_cid = pm.read_int(value + 0x44)
        except Exception:
            pass
        return deck_cid, repl_cid, duel_cid

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
                        deck_cid, repl_cid, duelcid = self.get_memory()
                        if duelcid != self.duel_cid and 3899 < duelcid < 18000:
                            final_cid = duelcid
                            self.duel_cid = duelcid
                        elif repl_cid != self.repl_cid and 3899 < deck_cid < 18000:
                            final_cid = repl_cid
                            self.repl_cid = repl_cid
                        elif deck_cid != self.deck_cid and 3899 < deck_cid < 18000:
                            final_cid = deck_cid
                            self.deck_cid = deck_cid
                    except Exception:
                        self.pm, self.base_addr = None, 0
            if final_cid in self.cards_info:
                hover_info = self.cards_info[final_cid]
                # print(final_cid, hover_info)
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
