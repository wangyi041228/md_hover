import asyncio
import threading
import time
import webbrowser
from ctypes import windll
# from datetime import datetime
from json import loads
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
        self.geometry(f'440x320+5+466')
        self.wm_attributes('-topmost', True)
        self.override = False
        self.toolwindow = False
        self.scan_mode = True
        self.low_rate = True
        self.ygoid = 89631139
        self.scan_cid = 0
        self.last_hash = [0] * 3
        self.pm = None
        self.base_addr = 0
        self.deck_cid = -1
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
                pdesc = card_info['text']['pdesc']
                desc = card_info['text']['desc'].replace('\r\n', '\n')
                fdesc = '\n　　○○○○　　\n'.join((pdesc, desc)) if pdesc else desc
                self.cards_info[cid_int] = (card_info['id'], cn_name, jp_name, en_name, fdesc)
        with open('tokens.json', 'r', encoding='utf-8') as f:
            data = loads(f.read())
            for cid, info in iter(data.items()):
                self.cards_info[int(cid)] = (89631139, info[0], info[1], info[2], info[3])
        with open('hash.json', 'r', encoding='utf-8') as f:
            self.art_hash = loads(f.read())
        self.fontsize = 12
        self.Label1 = Label(self, width=3)
        self.Label1.pack(side='left', expand=0, fill='y')
        self.Label2 = Label(self, font=('微软雅黑', self.fontsize + 1, 'bold'), text='青眼白龙', padding=-3)
        self.Label2.pack(side='top', fill='x')
        self.Label2.bind('<Button-1>', copy_widget_text)
        self.Label4 = Label(self, font=('微软雅黑', self.fontsize + 1, 'bold'), text='青眼の白龍', padding=-3)
        self.Label4.pack(side='top', fill='x')
        self.Label4.bind('<Button-1>', copy_widget_text)
        self.Label5 = Label(self, font=('微软雅黑', self.fontsize + 1, 'bold'), text='Blue-Eyes White Dragon', padding=-3)
        self.Label5.pack(side='top', fill='x')
        self.Label5.bind('<Button-1>', copy_widget_text)
        self.Label3 = Label(self, font=('微软雅黑', self.fontsize), wraplength=400, anchor='nw',
                            text='以高攻击力著称的传说之龙。任何对手都能粉碎，其破坏力不可估量。')
        self.Label3.pack(side='right', expand=1, fill=BOTH)
        self.Button1 = Button(self.Label3, text='内存模式', width=30, command=self.main_start)
        self.Button1.pack()
        self.Button1.place(x=50, y=50)
        self.Button1b = Button(self.Label3, text='识图模式', width=30, command=self.main_start_b)
        self.Button1b.pack()
        self.Button1b.place(x=50, y=100)

        self.Button2 = Button(self.Label1, text='+', width=3, command=self.font_size_up)
        self.Button2.pack()
        self.Button2.place(x=-3, y=55)
        self.Button3 = Button(self.Label1, text='-', width=3, command=self.font_size_down)
        self.Button3.pack()
        self.Button3.place(x=-3, y=80)
        self.Button4 = Button(self.Label1, text='网', width=3, command=self.web)
        self.Button4.pack()
        self.Button4.place(x=-3, y=160)
        self.Button5 = Button(self.Label1, text='图', width=3, command=self.switch)
        self.Button5.pack()
        self.Button5.place(x=-3, y=210)
        self.Button6 = Button(self.Label1, text='4', width=3, command=self.rate)
        self.Button6.pack()
        self.Button6.place(x=-3, y=110)
        self.Button7 = Button(self.Label1, text='顶', width=3, command=self.over)
        self.Button7.pack()
        self.Button7.place(x=-3, y=0)
        self.Button8 = Button(self.Label1, text='标', width=3, command=self.icon)
        self.Button8.pack()
        self.Button8.place(x=-3, y=25)
        self.mainloop()

    def get_loop(self, loop):
        self.loop = loop
        asyncio.set_event_loop(self.loop)
        self.loop.run_forever()

    def main_start(self):
        self.scan_mode = False
        self.Button1.destroy()
        self.Button1b.destroy()
        coroutine1 = self.handler()
        new_loop = asyncio.new_event_loop()
        t = threading.Thread(target=self.get_loop, args=(new_loop,))
        t.daemon = True
        t.start()
        asyncio.run_coroutine_threadsafe(coroutine1, new_loop)

    def main_start_b(self):
        self.Button1.destroy()
        self.Button1b.destroy()
        coroutine1 = self.handler()
        new_loop = asyncio.new_event_loop()
        t = threading.Thread(target=self.get_loop, args=(new_loop,))
        t.daemon = True
        t.start()
        asyncio.run_coroutine_threadsafe(coroutine1, new_loop)

    def font_size_up(self):
        self.fontsize += 1
        self.Label2['font'] = ('微软雅黑', self.fontsize + 1, 'bold')
        self.Label3['font'] = ('微软雅黑', self.fontsize)
        self.Label3['wraplength'] = self.winfo_width() - 40
        self.Label4['font'] = ('微软雅黑', self.fontsize + 1, 'bold')
        self.Label5['font'] = ('微软雅黑', self.fontsize + 1, 'bold')

    def font_size_down(self):
        if self.fontsize > 6:
            self.fontsize -= 1
            self.Label2['font'] = ('微软雅黑', self.fontsize + 1, 'bold')
            self.Label3['font'] = ('微软雅黑', self.fontsize)
            self.Label3['wraplength'] = self.winfo_width() - 40
            self.Label4['font'] = ('微软雅黑', self.fontsize + 1, 'bold')
            self.Label5['font'] = ('微软雅黑', self.fontsize + 1, 'bold')

    def web(self):
        webbrowser.open_new('https://ygocdb.com/card/' + str(self.ygoid))

    def switch(self):
        self.scan_mode = not self.scan_mode
        if self.scan_mode:
            self.Button5['text'] = '图'
        else:
            self.Button5['text'] = '内'

    def rate(self):
        self.low_rate = not self.low_rate
        if self.low_rate:
            self.Button6['text'] = '4'
        else:
            self.Button6['text'] = '16'

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
        deck_cid, duel_cid = -1, -1
        try:
            value = pm.read_longlong(base_addr + 0x01CCD278)
            for offset in [0xB8, 0x0, 0xF8, 0x1D8]:
                value = pm.read_longlong(value + offset)
            deck_cid = pm.read_int(value + 0x20)
        except Exception:
            pass
        try:
            value = pm.read_longlong(base_addr + 0x01cb2b90)
            for offset in [0xB8, 0x0]:
                value = pm.read_longlong(value + offset)
            duel_cid = pm.read_int(value + 0x44)
        except Exception:
            pass
        return deck_cid, duel_cid

    async def handler(self):
        while True:
            _t0 = time.time()
            _scan_mode = self.scan_mode
            _low_rate = self.low_rate
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
                        deck_cid, duelcid = self.get_memory()
                        if duelcid != self.duel_cid and 0 < duelcid < 18000:
                            final_cid = duelcid
                            self.duel_cid = duelcid
                        elif deck_cid != self.deck_cid and 0 < deck_cid < 18000:
                            final_cid = deck_cid
                            self.deck_cid = deck_cid
                    except Exception:
                        self.pm, self.base_addr = None, 0
            if final_cid in self.cards_info:
                hover_info = self.cards_info[final_cid]
                # print(final_cid, hover_info)
                self.ygoid = hover_info[0]
                self.Label2['text'] = hover_info[1]
                self.Label4['text'] = hover_info[2]
                self.Label5['text'] = hover_info[3]
                self.Label3['text'] = hover_info[4]
                self.Label3['wraplength'] = self.winfo_width() - 40
            dt = _t0 - time.time() + (0.25 if _low_rate else 0.0625)
            if dt > 0:
                await asyncio.sleep(dt)


if __name__ == '__main__':
    main = MainWindow()
