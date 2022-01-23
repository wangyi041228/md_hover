import asyncio
import os
import platform
import threading
import winreg
from ctypes import windll, Structure, c_long, byref
from datetime import datetime
from json import loads, dumps
from tkinter import *
from tkinter.ttk import *
import time
import win32gui
from PIL import Image, ImageTk, ImageGrab, ImageChops
from imagehash import average_hash, dhash, hex_to_hash  # phash, whash


def screenshot_a():
    toplist, winlist = [], []

    def enum_cb(hwnd, results):
        winlist.append((hwnd, win32gui.GetWindowText(hwnd)))

    win32gui.EnumWindows(enum_cb, toplist)
    mtga_window = [(hwnd, title) for hwnd, title in winlist if 'masterduel' == title]
    if mtga_window:
        hwnd = mtga_window[0][0]
        bbox = win32gui.GetWindowRect(hwnd)
        # win32gui.SetForegroundWindow(hwnd)
        img = ImageGrab.grab(bbox, all_screens=True)
        return img


def img_dhash(path):
    img = Image.open(path)
    return dhash(img, 16)


class MainWindow(Tk):

    def __init__(self):
        super().__init__()
        self.title('MD_HOVER')
        self.loop = None
        self.geometry(f'400x320+50+466')
        self['bg'] = 'white'
        self.wm_attributes('-topmost', True)
        self.dhash1 = []
        files = os.listdir('a')
        for file in files:
            self.dhash1.append((img_dhash('a\\' + file), file[:-4]))
        self.hover_id = -1
        self.cards_info = [
            [hex_to_hash('07b01df16ce4f668b774e354659d74d1f23238329dac4e21ce31b6f167991963'),
             [0, '宝贝龙', 'Baby Dragon', 'ベビードラゴン',
              '怪兽 通常', '不要因为这条龙还是个孩子就轻视它。体内隐藏的力量无法估计。']],
            [hex_to_hash('a698cea08faaa794b7b9b210b8a476b6e1b2f1b259b2690267ab6a96c584c269'),
             [1, '守城翼龙', 'Winged Dragon, Guardian of the Fortress #1', '砦を守る翼竜',
              '怪兽 通常', '守护山寨的龙，从空中急速下降攻击敌人。']],
            [hex_to_hash('660866f0f6b0d698da58d8c8e72ccc9cc0d4e634be30d674f2ba78e3b9f1cfa4'),
             [2, '黑魔导士', 'Dark Magician', 'ブラック・マジシャン',
              '怪兽 通常', '在魔法师中，攻击力、守备力都是最高级别。']],
            [hex_to_hash('86668fa69f063e023f017c4179b13bb1666c65dce61a371d274f207338e33c43'),
             [3, '栗子球', 'Kuriboh', 'クリボー',
              '怪兽 效果', '①：对手怪兽的攻击要使自己受到战斗伤害的伤害计算时可以将这张卡从手牌舍弃发动。那次战斗发生的对自己的战斗伤害变为0。']]
        ]
        self.cards_db = [(x[0], x[1][0]) for x in self.cards_info]
        self.label = Label(self)
        self.label.pack(expand=1, fill=BOTH)
        self.Button = Button(self, text='开始监测', width=9, command=self.main_start)
        self.Button.pack()
        self.Button.place(x=0, y=0)
        self.bind('<Enter>', self.alpha_max)
        self.mainloop()

    def get_loop(self, loop):
        self.loop = loop
        asyncio.set_event_loop(self.loop)
        self.loop.run_forever()

    def main_start(self):
        self.Button.destroy()
        coroutine1 = self.handler()
        new_loop = asyncio.new_event_loop()
        t = threading.Thread(target=self.get_loop, args=(new_loop,))
        t.daemon = True
        t.start()
        asyncio.run_coroutine_threadsafe(coroutine1, new_loop)

    def alpha_max(self, event):
        self.attributes('-alpha', 1.0)

    def alpha_min(self, event):
        self.attributes('-alpha', 0.3)

    async def handler(self):
        while True:
            await asyncio.sleep(0.01)
            now_str = str(datetime.now())[2:19].replace(":", "_")

            now_img = screenshot_a()
            if not now_img:
                continue
            imgx, imgy = now_img.size

            if imgx == 1920 and imgy == 1080:
                art_img = None
                art_mode = ''
                art_hash = None
                hash_diff = 0
                card_type = 0

                sample1 = now_img.crop((64, 237, 74, 347))
                _hash = dhash(sample1, 16)
                hash_compare = [(_hash - item[0], item[1]) for item in self.dhash1]
                hash_min = min(hash_compare, key=lambda xx: xx[0])
                # print(hash_min[1], hash_min[0])
                if hash_min[0] < 30:
                    if hash_min[1] in ['10', '11', '12', '13', '14']:
                        art_img = now_img.crop((74, 238, 198, 331))
                    else:
                        art_img = now_img.crop((82, 239, 189, 346))
                    art_hash = dhash(art_img, 16)
                    art_mode = '卡查'
                    hash_diff = hash_min[0]
                    card_type = hash_min[1]

                if not art_mode:
                    pass

                if not art_mode:
                    pass

                if art_mode:
                    # print(now_str, art_mode, card_type, hash_diff, str(art_hash))
                    hash_compare = [(art_hash - item[0], item[1]) for item in self.cards_db]
                    hash_min = min(hash_compare, key=lambda xx: xx[0])
                    # if hash_min[0] < 30:
                    new_id = hash_min[1]
                    # art_img.save(now_str + '.png')
                else:
                    new_id = -1
                if new_id != self.hover_id:
                    if new_id == -1:
                        self.attributes('-alpha', 0.3)
                    else:
                        new_text = ''
                        for line in self.cards_info[new_id][1][1:]:
                            print(line)
                            new_text += line + '\n'
                        self.label['text'] = new_text
                        self.label['wraplength'] = self.winfo_width() - 20
                        self.attributes('-alpha', 1.0)
                    self.hover_id = new_id
                else:
                    pass
                    # print(time, self.hover_id)
            # print(now_str, imgx, imgy)


if __name__ == '__main__':
    main = MainWindow()
    # print(img_dhash(r'D:\PycharmProjects\md_hover\img\4064.png'))
"""
卡查
左64上237 10宽110高 - 左82上239 107x107 - 左74上238 124宽93高
对战

放大

只做了4张牌的数据。感谢文本来源 www.ourocg.cn！
只做了组牌界面，还没做放大界面和对战界面的识别。
只适配1080P全屏，没适配其他分辨率和窗口。
随时烂尾，仅供学习交流。
"""