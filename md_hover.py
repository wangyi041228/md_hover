import asyncio
import threading
from ctypes import windll
from datetime import datetime
from json import loads  # dumps
from tkinter import *
from tkinter.ttk import *
import time
import win32gui
import win32ui
from PIL import Image  # ImageChops ImageTk ImageGrab
import dhash
import webbrowser
import pymem
import pyperclip
windll.user32.SetProcessDPIAware()


def screenshot():
    hwnd = win32gui.FindWindow(0, 'masterduel')
    if hwnd:
        box = win32gui.GetWindowRect(hwnd)
        box_w = box[2] - box[0]
        box_h = box[3] - box[1]
        hwndDC = win32gui.GetWindowDC(hwnd)
        mfcDC = win32ui.CreateDCFromHandle(hwndDC)
        saveDC = mfcDC.CreateCompatibleDC()
        saveBitMap = win32ui.CreateBitmap()
        saveBitMap.CreateCompatibleBitmap(mfcDC, box_w, box_h)
        saveDC.SelectObject(saveBitMap)
        result = windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 3)
        bmpinfo = saveBitMap.GetInfo()
        bmpstr = saveBitMap.GetBitmapBits(True)
        im = Image.frombuffer('RGB', (bmpinfo['bmWidth'], bmpinfo['bmHeight']), bmpstr, 'raw', 'BGRX', 0, 1)
        win32gui.DeleteObject(saveBitMap.GetHandle())
        saveDC.DeleteDC()
        mfcDC.DeleteDC()
        win32gui.ReleaseDC(hwnd, hwndDC)
        if result == 1:
            return True, im
    return False, None


def get_mem():
    try:
        pm = pymem.Pymem('masterduel.exe')
        base_addr = pymem.process.module_from_name(pm.process_handle, 'GameAssembly.dll').lpBaseOfDll
        deck_addr = base_addr + int("0x01CCD278", base=16)
        duel_addr = base_addr + int("0x01cb2b90", base=16)
        try:
            value = pm.read_longlong(deck_addr)
            for offset in [0xB8, 0x0, 0xF8, 0x1D8]:
                value = pm.read_longlong(value + offset)
            deck_pointer_value = (value + 0x20)
            deck_cid = pm.read_int(deck_pointer_value)
        except Exception as e:
            deck_cid = 0
        try:
            value = pm.read_longlong(duel_addr)
            for offset in [0xB8, 0x0]:
                value = pm.read_longlong(value + offset)
            duel_pointer_value = (value + 0x44)
            duel_cid = pm.read_int(duel_pointer_value)
        except Exception as e:  # pymem.exception.ProcessNotFound pymem.exception.MemoryReadError
            duel_cid = 0
    except Exception as e:
        return 0, 0
    return deck_cid, duel_cid


def get_widget_text(event):
    pyperclip.copy(event.widget['text'])


class MainWindow(Tk):
    def __init__(self):
        super().__init__()
        self.title('MD_HOVER')
        self.loop = None
        self.geometry(f'440x320+5+466')
        self.wm_attributes('-topmost', True)
        self.dhash1 = [
(948671198198901504627105345415049419891654810783989213831072240419345639644946855625040574212598844597633988647234510868093502590461869797514101830713248, 1),
(948262018129417386261369233939751510490099131423826503827002607480306931279812272207079792218269570120200954555554926180559924938619222545511495026868192, 2),
(948262018129417386261369233939751510411295119031037545402444592463181653427568407687789425317959770737592958198838055362072154039159372729064970752360352, 3),
(947852844303430368527617585227648060517268038775282162303759181042169636703017823368429016556419032717949097363319728376627396851686161625547578522206144, 4),
(948671191955404403995120882651854960305322199286792928501130003886182963097758138575771336744014874978341969158033748064166239516641787686844862255071200, 5),
(942942689712070082522248437688321314581288528288600622706749974246106895490248856758558731973776073865958600275067584766887569394520426957693809354276835, 6),
(948875775746649362546004476641694006007281405798214888500445835805357318891666301756482444091952413525762857057234990327689292695550363293518401693089775, 7),
(948262024372914486893353696702945970076431742921022789156944909198620850530555749846611032458881776247184896875209128705740883947185928282560699520974816, 8),
(948262024372914486893353696702945969997627730528233830732386828998333622920396271305680139122985189438693453332214367200544638219745342622193334682058721, 9),
(947648254268735944174223690678787619061363835960563918602472659654527406208416521936070804574554638746535372093396586614295172506415151171862573140148192, 10),
(947648257390484494490227012908391660163674704120884761758414695413917929736841239490163903013694240166998900197881889115683153136084352496694339897917376, 11),
(947648257390484494490227012908391660163674704120884621772232322265233859814303682256226941968175041339749641699109116353262093353744192748072654615871392, 12),
(947648254268688309344749529489999105777088192898806328091581278914857318561835012861729424544839841817470721571646038155252243076499383770679738704592800, 13),
(943351863538104735085330068934354487936353793042566535785382428076795582748869238087820434819338246128376785928952522891884900082650463245821224011431904, 14)
        ]
        self.dhash2 = [
(10081378544931573597285150352582354296017530317712813943145007883319630980422536141189191253059645902124046345440957524426550475108703748309339389768531967, 1),
(10081378551174641984451778637862400227822103729195607870168332001042192998620172357834802245059669882011473483523197085022659784725902652018434777565659135, 2),
(10079128063914227883756222255960628627462988690198182707703679558710621851902234293627913579573966518178106871946078881043769225701733181626922882499117055, 3),
(10079128063914227884483060980256235518012311896868460186404021219253559377880291167827366246277514265414468325956932416682979597169940143666168778759110655, 4),
(10080764784192545436581887442621020080691416550300092652822032387865687732482145602170525781213182771540477924094998038267880721238391625995442776111611903, 5),
(10062965351268459399603840721037249499168166929021275777044102514327285944205708703796112492064320822066751728933116695734680832653055477156718422402301951, 6),
(10080355594757434587905453896241746160023839173407702752076274161787126594982864963698082735318000724185707972660057511493078231801400219130289965278199807, 7),
(10080764759218175955418078665952656128858162774097635665643265664429806873944027830708812981507897221949988122555897221209005474680377104716260190768365567, 8),
(10078923483244731475521330894367776526602063334157779423221316821031167065829745270365638845220271237712595167527528975565145433416605892079287758747631615, 9),
(10080560194157422298036298647478032010791043608506628973786031316693252435536649632585269104479048161273691494636717520632431152760866097724250650268696575, 10),
(10080560169183433895508360795071390890049038477209237559880716868796912916452474999555146580541732432544818024477073159134790252781561831340463781052448767, 11),
(10080560194157803376672180662418531325141730567059256984927066749905527631207688817108407196088278182433465928138010829082363014590202794818957761484718079, 13),
        ]
        self.dhash3 = [
(157125146718485008028911076206643652098164776420723718369629555676148269974502504295297656977517658769023821602461395419655787045374589830787323853340671, 1),
(158761842022433110218149170532476978985292682508606875081249722522127268583239491333556941363142995412094743924459460335651116225378751134854361714327551, 2),
(157329733631478525617840496015805262494776454708185019622107083541779532773280332384808981478518017991277045242327828657962150278528754363875502843822079, 3),
(157125146718485008028899980112455357299000673558698901226869675693140750079391176471156697791556369423665099188636474271065110053359338654512078506754047, 4),
(157329733631478516895786895147319716215705621273254637515108778097306804625568527455945132602115942912083090560108506739182211591091567156552802232238079, 5),
(158966482005914085937562247775405260998587873548727670566787558937635415565248451883668591387452572619612103493774091619686650268638120945530155053416447, 6),
(157125146718485047278224369526917067183821622813623227689685722438565430393262631325928486904162082721022195073547008357619691766289733364775565326876287, 7),
(157125146718485040010214204234027864113510605725167728852336303536708057283605310677552000945260214892615383748484080558121776288499352850995410874007551, 8),
(157125146718485031287783520286997281571992326893950109028422108395084038430870165069503104474117951471819662295401833510666133630850903440226842621050879, 9),
(157125146718485001487340370773206528204839443811931915985224615682891810103056469099680364214652263849208964445309060320264527014607706657663740880683007, 10),
(157125146718485001487340370773206528204839443811931915985224615682891810103056469099680364227507768107499915345614157309123894836079499232523204239097855, 11),
(157125146718485001487340370773206528204839443811931915985224615682891810103056016786831780948263890525048774258256462381913683518570434859262128869670911, 12),
(157125146718485001487340370773206528204839443811931915985224615682891810103056016786831780948263890429267802954040745041906504216088196751638068195295231, 13),
(157125149840995697445521779218674501234309751947888467377676419517597034438163721637964354320790985702164827137787406637026798438085812829480909180829695, 14)
        ]
        self.scan_cid = 0
        self.mem_cids = (0, 0)
        self.scan_mode = True
        self.low_rate = True
        self.ygoid = 89631139
        self.last_hash = 0
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
        self.Label2.bind('<Button-1>', get_widget_text)
        self.Label4 = Label(self, font=('微软雅黑', self.fontsize + 1, 'bold'), text='青眼の白龍', padding=-3)
        self.Label4.pack(side='top', fill='x')
        self.Label4.bind('<Button-1>', get_widget_text)
        self.Label5 = Label(self, font=('微软雅黑', self.fontsize + 1, 'bold'), text='Blue-Eyes White Dragon', padding=-3)
        self.Label5.pack(side='top', fill='x')
        self.Label5.bind('<Button-1>', get_widget_text)
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
        self.Button2.place(x=-3, y=0)
        self.Button3 = Button(self.Label1, text='-', width=3, command=self.font_size_down)
        self.Button3.pack()
        self.Button3.place(x=-3, y=25)
        self.Button4 = Button(self.Label1, text='网', width=3, command=self.web)
        self.Button4.pack()
        self.Button4.place(x=-3, y=85)
        # self.Button5 = Button(self.Label1, text='图', width=3, command=self.switch)
        # self.Button5.pack()
        # self.Button5.place(x=-3, y=165)
        self.Button6 = Button(self.Label1, text='4', width=3, command=self.rate)
        self.Button6.pack()
        self.Button6.place(x=-3, y=55)
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
        self.Label3['wraplength'] = self.winfo_width() - 45
        self.Label4['font'] = ('微软雅黑', self.fontsize + 1, 'bold')
        self.Label5['font'] = ('微软雅黑', self.fontsize + 1, 'bold')

    def font_size_down(self):
        if self.fontsize > 6:
            self.fontsize -= 1
            self.Label2['font'] = ('微软雅黑', self.fontsize + 1, 'bold')
            self.Label3['font'] = ('微软雅黑', self.fontsize)
            self.Label3['wraplength'] = self.winfo_width() - 45
            self.Label4['font'] = ('微软雅黑', self.fontsize + 1, 'bold')
            self.Label5['font'] = ('微软雅黑', self.fontsize + 1, 'bold')

    def web(self):
        webbrowser.open_new('https://ygocdb.com/card/' + str(self.ygoid))

    # def switch(self):
    #     self.scan_mode = not self.scan_mode
    #     if self.scan_mode:
    #         self.Button5['text'] = '图'
    #     else:
    #         self.Button5['text'] = '内'

    def rate(self):
        self.low_rate = not self.low_rate
        if self.low_rate:
            self.Button6['text'] = '4'
        else:
            self.Button6['text'] = '16'

    def get_scan(self):
        now_str = str(datetime.now())[2:19].replace(":", "_")
        flag, now_img = screenshot()
        if not flag:
            return -1
        imgx, imgy = now_img.size
        if imgx < 1280:
            return -1
        elif imgx == 1280:
            pass
        elif imgx < 1366:
            now_img = now_img.crop((0, 0, 1280, 720))
        elif imgx == 1366:
            pass
        elif imgx < 1440:
            now_img = now_img.crop((0, 0, 1366, 768))
        elif imgx == 1440:
            pass
        elif imgx < 1600:
            now_img = now_img.crop((0, 0, 1440, 810))
        elif imgx == 1600:
            pass
        elif imgx < 1920:
            now_img = now_img.crop((0, 0, 1600, 900))
        elif imgx == 1920:
            pass
        elif imgx < 2048:
            now_img = now_img.crop((0, 0, 1920, 1080))
        elif imgx == 2048:
            pass
        elif imgx < 2560:
            now_img = now_img.crop((0, 0, 2048, 1152))
        elif imgx == 2560:
            pass
        elif imgx < 3200:
            now_img = now_img.crop((0, 0, 2560, 1440))
        elif imgx == 3200:
            pass
        elif imgx < 3840:
            now_img = now_img.crop((0, 0, 3200, 1800))
        elif imgx == 3840:
            pass
        else:
            now_img = now_img.crop((0, 0, 3840, 2160))
        if now_img.size[0] != 1920:
            now_img = now_img.resize((1920, 1080), Image.LANCZOS)
        # now_img.save(now_str + '.png')

        art_img = None
        art_mode = ''
        art_hash = None
        hash_diff = 0
        card_type = 0

        sample1 = now_img.crop((57, 235, 74, 354))  # 17x119
        # sample1.save(now_str + '_1.png')
        _hash = dhash.dhash_int(sample1, 16)
        hash_compare = [(dhash.get_num_bits_different(_hash, item[0]), item[1]) for item in self.dhash1]
        hash_min = min(hash_compare, key=lambda xx: xx[0])
        # print('卡查', hash_min[1], hash_min[0])
        if hash_min[0] < 65:
            if hash_min[1] > 9:
                art_img = now_img.crop((74, 238, 198, 331))  # 124x93 133.33%
            else:
                art_img = now_img.crop((82, 239, 189, 346))
            art_hash = dhash.dhash_int(art_img, 16)
            art_mode = '卡查'
            hash_diff = hash_min[0]
            card_type = hash_min[1]

        if not art_mode:
            sample1 = now_img.crop((40, 248, 51, 367))  # 17x119
            # sample1.save(now_str + '_2.png')
            _hash = dhash.dhash_int(sample1, 16)
            hash_compare = [(dhash.get_num_bits_different(_hash, item[0]), item[1]) for item in self.dhash2]
            hash_min = min(hash_compare, key=lambda xx: xx[0])
            # print('对战', hash_min[1], hash_min[0])
            if hash_min[0] < 65:
                if hash_min[1] > 9:
                    art_img = now_img.crop((53, 254, 196, 361))  # 143x107 133.64%
                else:
                    art_img = now_img.crop((62, 254, 187, 379))
                art_hash = dhash.dhash_int(art_img, 16)
                art_mode = '对战'
                hash_diff = hash_min[0]
                card_type = hash_min[1]

        if not art_mode:
            sample1 = now_img.crop((141, 331, 158, 586))  # 17x255
            # sample1.save(now_str + '_3.png')
            _hash = dhash.dhash_int(sample1, 16)
            hash_compare = [(dhash.get_num_bits_different(_hash, item[0]), item[1]) for item in self.dhash3]
            hash_min = min(hash_compare, key=lambda xx: xx[0])
            # print('放大', hash_min[1], hash_min[0])
            if hash_min[0] < 65:
                if hash_min[1] > 9:
                    art_img = now_img.crop((164, 285, 548, 570))  # 384x285 134.74%
                else:
                    art_img = now_img.crop((189, 288, 523, 622))
                art_hash = dhash.dhash_int(art_img, 16)
                art_mode = '放大'
                hash_diff = hash_min[0]
                card_type = hash_min[1]

        if art_mode and dhash.get_num_bits_different(art_hash, self.last_hash) > 10:
            self.last_hash = art_hash
            hash_compare = [(dhash.get_num_bits_different(art_hash, item[0]), item[1]) for item in self.art_hash]
            hash_min = min(hash_compare, key=lambda xx: xx[0])
            new_cid = hash_min[1]
            # print(now_str, art_mode, '牌类别', card_type, '误差', hash_diff, 'cid', new_cid)
            # art_img.save(now_str + '_4.png')
        else:
            new_cid = 0
        return new_cid

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
                    self.scan_cid = new_cid  # 识别空不存
            else:
                new_cids = get_mem()
                if new_cids[0] != self.mem_cids[0] and new_cids[0] != 0:
                    final_cid = new_cids[0]
                elif new_cids[1] != self.mem_cids[1] and new_cids[1] != 0:
                    final_cid = new_cids[1]
                self.mem_cids = new_cids
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
